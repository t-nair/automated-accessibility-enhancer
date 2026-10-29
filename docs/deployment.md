# Deploying the web page

The website is served by [Waitress](https://docs.pylonsproject.org/projects/waitress/),
which runs on Windows and Linux. Everything, including the image descriptions, runs on
the server itself: no course file or picture is sent to an outside service.

The [open decisions](#open-decisions) at the end still need an owner before this runs
anywhere real.

---

## How the app runs (read this before changing any server setting)

- **One process, one worker thread.** When `app.py` is imported it starts a queue in
  memory and one background thread that takes files off it one at a time. That thread
  loads the captioning model (`Qwen/Qwen2-VL-2B-Instruct`, about 4.5 GB in fp16 on a GPU,
  about 9 GB in fp32 on a CPU) the first time it needs it and keeps it loaded.
- **Progress is kept in memory** in that same process. The progress bar asks the process
  that handled the upload.
- **Submissions are in SQLite** (`data/submissions.db`). Each request opens its own
  connection.

So the app must run as **exactly one process**. More processes (Gunicorn `-w 4`, several
Waitress instances behind a load balancer, several containers) would each start their own
queue and their own copy of the model: GPU memory runs out, and the progress bar shows 0
whenever a request lands on a process that is not the one doing the work. Gunicorn's
`--preload` is worse: the worker thread starts before the fork and does not survive it,
so nothing is ever processed.

Threads are fine. Web requests are short, because uploads are only saved and queued, and
Waitress reads the whole request body before handing it to the app, so a slow upload does
not hold a thread. Waitress's default of 4 threads is plenty. More threads would never
process more files at the same time.

Running files in parallel, or on a second machine, needs a shared queue and shared
progress first (see open decisions).

---

## What the server needs

| Need | Why |
|------|-----|
| Python 3.11 or 3.12 | OCR for scanned PDFs (`rapidocr-onnxruntime`) does not support newer versions |
| An NVIDIA GPU with at least 8 GB, or a lot of patience | On a CPU, describing a deck with many pictures takes minutes per file |
| About 10 GB of disk for the model, plus room for uploads | See "Where data is kept" |
| LibreOffice (optional) | Only for old `.ppt` files |
| Chrome, Chromium or Edge | Only for PDF uploads; set `PDF_PIPELINE_CHROME` if it is not on `PATH` |
| A reverse proxy that handles HTTPS | Waitress does not do TLS |

---

## Install

```bash
git clone https://github.com/t-nair/automated-accessibility-enhancer.git /srv/accessibility-enhancer
cd /srv/accessibility-enhancer
python3.12 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m pipeline.setup_models
```

For a GPU, install the CUDA build of PyTorch that matches the server's driver (see
pytorch.org) before `requirements.txt`, or the plain `torch` from PyPI may be CPU-only.

Download the captioning model once, while the server can still reach Hugging Face:

```bash
HF_HOME=/srv/accessibility-enhancer/hf-cache .venv/bin/python -c "from huggingface_hub import snapshot_download; snapshot_download('Qwen/Qwen2-VL-2B-Instruct')"
```

After that, set `HF_HUB_OFFLINE=1` (below) so the running app never contacts Hugging Face.
Without it the model is downloaded on the first upload instead, which makes that first
person wait several minutes.

---

## Environment variables

| Variable | Needed | What it does |
|----------|--------|--------------|
| `SECRET_KEY` | Yes | Signs the cookie that ties each visitor to their files. Keep it the same across restarts and secret. If unset, one is made and saved in `data/secret_key.txt`. Make one with `python -c "import secrets; print(secrets.token_hex(32))"` |
| `FLASK_SESSION_COOKIE_SECURE` | Yes, behind HTTPS | Set to `true` so the cookie is only sent over HTTPS |
| `HF_HOME` | Recommended | Where the model is stored. Default is `~/.cache/huggingface` of the user running the app |
| `HF_HUB_OFFLINE` | Recommended | Set to `1` after downloading the model, so nothing is fetched at run time |
| `MIN_FONT_PT` | No | Font size below which text is reported (default `9`) |
| `PDF_PIPELINE_CHROME` | Only if Chrome is not on `PATH` | Path to Chrome, Chromium or Edge for PDF export |
| `HOST`, `PORT` | No | Address to listen on (default `127.0.0.1:5000`). Keep `127.0.0.1` behind a proxy on the same machine |

Any other `FLASK_`-prefixed variable sets the Flask setting of that name. Do not set
`FLASK_DEBUG` on a server.

---

## Where data is kept

All paths are relative to the **working directory**, so always start the app from the
project folder (`WorkingDirectory=` in systemd). Back up `data/`, `uploads/` and
`processed/` together; they only make sense as a set.

| Path | What | Keep across restarts? |
|------|------|------------------------|
| `data/submissions.db` | Every submission and its status | Yes |
| `data/secret_key.txt` | The cookie key, if `SECRET_KEY` is not set | Yes, or everyone loses their file list |
| `uploads/` | The original files people uploaded | Yes, retry needs them |
| `processed/` | Fixed files, reports and PDF output folders | Yes, downloads come from here |
| `pipeline.log` | Log of every file processed | Yes, rotate it (it is never trimmed) |
| `$HF_HOME` | The captioning model | Yes, or it is downloaded again |
| `pipeline/ocr_models/` | OCR model for scanned PDFs | Yes |

These hold course material, so they must be on a disk only the service account can read,
and they must not be on storage that is wiped on redeploy (a container's own filesystem,
for example). Nothing is deleted automatically; people delete their own files from the
status page.

---

## Start command

```bash
.venv/bin/python app.py
```

This starts Waitress on `HOST`:`PORT` (default `127.0.0.1:5000`), so only the reverse
proxy on the same machine can reach it.

The app does not use `X-Forwarded-*` headers (every redirect is relative), so Waitress
needs no trusted-proxy settings.

For the Flask debugger on a laptop, use `flask --app app run --debug --no-reload`, never on
a server. (The reloader restarts the app in the middle of a file, because loading PyTorch
touches files it watches.)

### systemd (Linux)

`/etc/systemd/system/accessibility-enhancer.service`:

```ini
[Unit]
Description=Automated accessibility enhancer
After=network.target

[Service]
User=accessibility
WorkingDirectory=/srv/accessibility-enhancer
EnvironmentFile=/etc/accessibility-enhancer.env
ExecStart=/srv/accessibility-enhancer/.venv/bin/python app.py
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

`/etc/accessibility-enhancer.env` (readable only by root and the service user):

```ini
SECRET_KEY=<the generated key>
FLASK_SESSION_COOKIE_SECURE=true
HF_HOME=/srv/accessibility-enhancer/hf-cache
HF_HUB_OFFLINE=1
```

```bash
sudo systemctl enable --now accessibility-enhancer
```

On Windows, run `.venv\Scripts\python.exe app.py` from the project folder with a service
wrapper such as NSSM, and put the variables in the service's environment.

---

## Reverse proxy and HTTPS (nginx)

```nginx
server {
    listen 80;
    server_name accessibility.example.edu;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name accessibility.example.edu;

    ssl_certificate     /etc/ssl/certs/accessibility.pem;
    ssl_certificate_key /etc/ssl/private/accessibility.key;

    # each file can be up to 50 MB and several can be sent at once (see open decisions)
    client_max_body_size 205m;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
    }
}
```

No long proxy timeouts are needed: processing happens after the upload request has
returned, and the page polls for progress.

---

## Restarting and upgrading

```bash
sudo systemctl restart accessibility-enhancer
```

The queue is only in memory, so on start the app puts every file still shown as "queued" or
"processing" back on the queue. A file in progress starts again from its original upload,
which is never moved or changed. If that original is gone, the file is marked as an error
that asks the person to upload it again.

---

## Open decisions

These are not settled in the code or the repository and need someone to decide:

1. **The host.** No server has been chosen. The `aws-deployment` and
   `infra/initial-terraform-setup` branches target AWS and Render, but they replace the
   local model with Claude on Amazon Bedrock, which sends course images off the server.
   They are not merged and this guide does not use them. Whoever picks a host needs to
   confirm it may hold course material and has a GPU (or accept CPU speed). (#102)
2. **Domain and certificate.** `accessibility.example.edu` and the certificate paths above
   are placeholders. (#102)
3. **Who can use it.** There is no login; anyone who can reach the site can upload. Until
   there is university sign-in, keep it on the campus network or VPN, or put sign-in in
   front of it at the proxy. Real sign-in is #77; the interim choice is #102.
4. **Upload limits.** Each file may be 50 MB (`MAX_FILE_SIZE_MB`), but nothing limits how
   many files one upload sends; nginx's `client_max_body_size 205m` above is the only cap.
   Pick a total and enforce it in the app too (#74).
5. **Keeping files.** Nothing is deleted automatically. Decide how long uploads and results
   may be kept and whether something should clear them. (#78)
6. **Model version.** The captioning model is not pinned to a revision, so a fresh
   download can get a newer one. Pin it, or keep the `HF_HOME` folder backed up. (#103)
7. **More than one file at a time.** Needs a shared queue and progress outside the process
   and enough GPU memory for several models. Until then, one process. (#104)
