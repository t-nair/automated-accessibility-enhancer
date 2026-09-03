# Infrastructure

Terraform for running the pipeline on AWS: a deck is uploaded to S3, a Lambda
picks it up, fixes the reading order and writes descriptions for the pictures
using Claude on Bedrock, and puts the result in a second bucket.

This is the deployment side of the project. The Flask app in the repository root
still runs on a laptop and is unaffected by any of it.

---

## What gets created

| Piece | What it is for |
|---|---|
| `bootstrap/` | The S3 bucket that holds Terraform's own state. Run once per AWS account. |
| `modules/s3_buckets` | The uploads and processed buckets, both with lifecycle expiry. |
| `modules/ecr` | The registry the pipeline container image is pushed to. |
| `modules/lambda` | The function, its execution role, and the S3 trigger. |
| `modules/observability` | Log group with a retention period, an SNS topic, and four alarms. |
| `modules/github_oidc` | A role GitHub Actions assumes to run `terraform plan` without an access key. |
| `envs/dev` | One of everything, wired together. Copy the directory to add an environment. |

### Why the buckets expire

Uploads are deleted after 7 days and processed files after 30. A deck is a
working file, not a record. Once the accessible copy has been downloaded there
is no reason to keep paying to store someone's lecture, and holding real course
material longer than it is needed is a liability rather than a feature. Both
periods are variables if they turn out to be wrong.

---

## Prerequisites

- Terraform 1.11 or newer. The S3 backend uses `use_lockfile`, which is what
  replaced the separate DynamoDB lock table, and needs 1.11.
- AWS credentials with permission to create the resources above.
- Docker, to build the Lambda image.
- Bedrock model access enabled for the Claude model you plan to use. This is a
  per-account, per-model switch in the AWS console under **Bedrock → Model
  access**, and nothing works until it is on.

---

## First time setup

### 1. Create the state bucket

```bash
cd infra/bootstrap
terraform init
terraform apply -var 'state_bucket_name=aae-tfstate-<something-unique>'
```

S3 bucket names are global across all of AWS, so pick something nobody else will
have. This step keeps its state in a local file, because the bucket that would
hold it does not exist yet. It only ever needs to run once.

### 2. Point the environment at it

```bash
cd ../envs/dev
```

Edit `backend.tf` and replace `REPLACE_WITH_STATE_BUCKET_NAME` with the bucket
name from step 1. Then:

```bash
cp terraform.tfvars.example terraform.tfvars
```

Fill in `terraform.tfvars`. It is gitignored, so real values stay out of the
repository.

### 3. Create the registry first

Lambda cannot be created pointing at an image that does not exist yet, so the
registry and the image come before everything else:

```bash
terraform init
terraform apply -target=module.ecr
```

### 4. Build and push the image

From the repository root, with `<account>`, `<region>` and `<repo>` taken from
the `ecr_repository_url` output:

```bash
aws ecr get-login-password --region <region> | docker login --username AWS --password-stdin <account>.dkr.ecr.<region>.amazonaws.com
```

```bash
docker build --platform linux/amd64 -t <account>.dkr.ecr.<region>.amazonaws.com/<repo>:latest .
```

```bash
docker push <account>.dkr.ecr.<region>.amazonaws.com/<repo>:latest
```

`--platform linux/amd64` matters. Building on an Apple Silicon machine without
it produces an arm64 image, which Lambda will refuse unless
`lambda_architecture` is set to `arm64` to match.

### 5. Create everything else

```bash
terraform apply
```

### 6. Try it

```bash
aws s3 cp "Error Test PPTX/issue_missing_alt_text.pptx" s3://<uploads_bucket>/test/
```

Then look in the processed bucket for `test/issue_missing_alt_text_updated.pptx`,
the `_alt_text` report next to it, and a `.status.json` saying how it went. If
nothing appears, the log group named in the outputs will say why.

---

## Deploying a change

```bash
docker build --platform linux/amd64 -t <repo-url>:latest .
docker push <repo-url>:latest
aws lambda update-function-code --function-name <function_name> --image-uri <repo-url>:latest
```

Terraform does not need to run for a code-only change, because the image tag it
points at has not changed.

This is also the weakness of using `latest`: two different images can carry the
tag over time and nothing records which one is deployed. Moving `image_tag` to a
build number or an image digest is the fix, and is worth doing before this is
used for real.

---

## CI

`modules/github_oidc` creates a role that GitHub Actions assumes through OIDC,
so no AWS access key is ever stored in the repository. GitHub mints a
short-lived token for the workflow run and AWS trades it for temporary
credentials.

The role can plan but not apply. It gets `ReadOnlyAccess` plus the S3 writes the
backend needs to take and release its lock — a plan is not read-only against the
backend, it writes a lock file while it runs. It is also explicitly denied
`s3:GetObject` on the uploads and processed buckets, because `ReadOnlyAccess`
would otherwise let a workflow download real course material, and a plan has no
reason to read a deck.

Applying from CI would need a second role with permissions to match, added
deliberately rather than by widening this one.

The workflow side is a job with `permissions: id-token: write` using
`aws-actions/configure-aws-credentials` with the `github_actions_role_arn`
output as `role-to-assume`. That workflow is not in the repository yet.

---

## Alarms

Four, all reporting to one SNS topic:

- **errors** — the pipeline raised an unhandled error on a deck.
- **throttles** — uploads are arriving faster than the reserved concurrency
  allows, so someone is waiting on a deck being turned away.
- **approaching-timeout** — a run took 80% of the timeout. The warning before
  decks start being cut off partway through instead of failing cleanly.
- **caption-failures** — a log metric filter counting the pipeline's own
  "Could not generate a caption" line.

The last one exists because the pipeline deliberately catches a captioning
failure and writes placeholder alt text rather than failing the whole deck. That
is right for the person waiting, and exactly why it needs its own alarm:
otherwise a broken Bedrock setup looks like a series of successful runs that
quietly describe nothing.

Email subscriptions stay pending until the recipient clicks the confirmation
link AWS sends. Terraform cannot do that for them.

---

## Known gaps

These are deliberate, not oversights:

- **`bedrock_model_arns` defaults to `*`.** AWS does not currently publish the
  resource ARN format for `bedrock-mantle:CreateInference`. Run the pipeline
  once, read the resource off the CloudTrail event, then narrow the variable.
- **Dependencies are unpinned.** `requirements.txt` has no versions, so two
  builds of the same commit can differ. Worth pinning before this is relied on.
- **`image_tag` defaults to `latest`.** See the note under *Deploying a change*.
- **No `terraform plan` has been run against these files.** They were written
  without Terraform installed locally, so they are unvalidated. Expect to fix
  something on the first `terraform init && terraform validate`.
- **The Flask app does not use any of this.** It still processes files on
  whichever machine it runs on. Connecting it to the buckets is separate work.
