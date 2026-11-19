# Infrastructure

Terraform for running the pipeline on AWS: a deck is uploaded to S3, a Lambda
picks it up, fixes the reading order and writes descriptions for the pictures
using Claude on Bedrock, and puts the result in a second bucket.

This is the deployment side of the project. The same Flask app runs either way:
on a laptop it keeps submissions in a sqlite file and does the work in a thread,
and on AWS it keeps them in DynamoDB and leaves the work to the Lambda.

---

## What gets created

| Piece | What it is for |
|---|---|
| `bootstrap/` | The S3 bucket that holds Terraform's own state. Run once per AWS account. |
| `modules/s3_buckets` | The uploads and processed buckets, both with lifecycle expiry. |
| `modules/ecr` | The registries the two container images are pushed to, one each for the pipeline and the website. |
| `modules/dynamodb` | The table of submissions, which replaces the sqlite file the local version uses. |
| `modules/ecs_web` | The website itself on ECS Express Mode, which brings its own load balancer and HTTPS certificate, and the permissions it runs with. |
| `modules/cognito` | Sign-in: a Cognito user pool, its sign-in pages, the website's app client, Google when configured, and SES for email. See *Sign-in* below. |
| `modules/lambda` | The function, its execution role, and the S3 trigger. |
| `modules/observability` | Log group with a retention period, an SNS topic, and four alarms. |
| `modules/github_oidc` | Two roles GitHub Actions assumes without an access key: one to run `terraform plan` on pull requests, one to deploy from `main`. |
| `envs/dev` | One of everything, wired together. Copy the directory to add an environment. |

### Why the buckets expire

Uploads are deleted after 7 days and processed files after 30. A deck is a
working file, not a record. Once the accessible copy has been downloaded there
is no reason to keep paying to store someone's lecture, and holding real course
material longer than it is needed is a liability rather than a feature. Both
periods are variables if they turn out to be wrong.

### Backups

- **Files:** both buckets are versioned. A file that is deleted or overwritten,
  by the app, by the expiry or by mistake, keeps its previous version for 3 days
  (`noncurrent_version_expiration_days`) and is then removed for good. So the
  longest a file can exist is its expiry plus 3 days: 10 days for an upload, 33
  for a processed file, and a file someone deletes from the status page is gone
  3 days later.
- **Submissions:** the table has point-in-time recovery, so it can be restored
  to any second in the last 35 days.
- **Tearing down:** `buckets_force_destroy` is false, so `terraform destroy`
  stops at a bucket that still holds files instead of emptying it.

Nothing is copied to another region or account. These are working files with a
short life, and a copy that outlived them would keep course material longer than
people are told. See *Restoring from a backup* below.

---

## Prerequisites

- Terraform 1.11 or newer. The S3 backend uses `use_lockfile`, which is what
  replaced the separate DynamoDB lock table, and needs 1.11.
- AWS credentials with permission to create the resources above.
- Docker, to build the Lambda image.
- A default VPC in the region, which the website runs in. An account without
  one (some university accounts) sets `web_subnet_ids` to two or more subnets
  instead.
- Bedrock working for the account. The Model access page has been retired, and
  models now enable themselves when first used, but Anthropic ones need a use
  case form submitting once, and a new account can be left with every Bedrock
  quota applied at 0, which throttles the first request. Both are worth
  confirming before deploying, because neither shows up until a caption is
  attempted. `aws bedrock-runtime converse` is the quickest way to check.

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

### 3. Create the registries first

Lambda and ECS cannot be created pointing at images that do not exist
yet, so the registries and the images come before everything else:

```bash
terraform init
terraform apply -target=module.ecr -target=module.ecr_web
```

### 4. Build and push the images

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

Then the website image, from `web_repository_url`, which is a second repository
and a second Dockerfile:

```bash
docker build --platform linux/amd64 -t <account>.dkr.ecr.<region>.amazonaws.com/<web-repo>:latest -f Dockerfile.web .
```

```bash
docker push <account>.dkr.ecr.<region>.amazonaws.com/<web-repo>:latest
```

The two images carry opposite things on purpose: the pipeline one has
LibreOffice, for converting an older `.ppt`, and the website one has Flask and
the pages.

`--platform linux/amd64` matters. Building on an Apple Silicon machine without
it produces an arm64 image, which Lambda will refuse unless
`lambda_architecture` is set to `arm64` to match.

### 5. Create everything else

```bash
terraform apply
```

### 6. Open the site

`terraform output website_url` is the address ECS serves it on, with its own
HTTPS certificate. Uploading a deck there exercises the whole thing: the
site puts it in the uploads bucket, the bucket starts the Lambda, the Lambda
writes its progress to the table, and the page follows along.

### 7. Try the pipeline on its own

```bash
aws s3 cp "Error Test PPTX/issue_missing_alt_text.pptx" s3://<uploads_bucket>/test/
```

Then look in the processed bucket for `test/issue_missing_alt_text_updated.pptx`,
the `_alt_text` report next to it, and a `.status.json` saying how it went. If
nothing appears, the log group named in the outputs will say why.

---

## Deploying a change

Merging to `main` deploys it (see *CI* below). The workflow builds both images,
tags them with the commit, and runs `terraform apply -var image_tag=<commit>`,
which moves the Lambda and the website onto them together. Every deploy is then
a commit you can name, and rolling back is re-running the deploy of an older
one.

By hand, the same thing is to build and push both images under one tag, the
pipeline from `Dockerfile` to `ecr_repository_url` and the website from
`Dockerfile.web` to `web_repository_url`:

```bash
docker build --provenance=false --platform linux/amd64 -t <repo-url>:<tag> .
docker push <repo-url>:<tag>
```

and then:

```bash
terraform apply -var image_tag=<tag>
```

---

## Sign-in

Accounts are in an Amazon Cognito user pool (`modules/cognito`, #77). The website
sends people to Cognito's sign-in pages and gets them back at
`<public_url>/auth/callback`. The website's half of this is the second part of
#77; until it ships the pool exists but nothing uses it.

Out of the box the pool has **email and password** only, with verification and
password reset by email. Google and SES are each turned on after the steps below.

Microsoft is left off for now. Cognito checks a token's issuer against one exact
URL, and Microsoft's multi-tenant endpoints put each tenant's own id in the
issuer, so one Microsoft provider can only sign in one tenant (UW's, or personal
accounts). That choice is in #77.

The pool has deletion protection on, because deleting it deletes every account.
To tear an environment down, turn it off in the console first.

### Google

1. In Google Cloud, create a project and open **APIs & Services → OAuth consent
   screen**. Choose **External**, add the `openid`, `email` and `profile` scopes,
   and publish it.
2. Under **Credentials**, create an **OAuth client ID** of type **Web
   application**. Add `terraform output cognito_identity_provider_redirect_uri` as
   an authorized redirect URI, and the host of `cognito_sign_in_domain` as an
   authorized JavaScript origin.
3. Put the client secret in Parameter Store yourself, so it never goes through
   GitHub:

   ```bash
   aws ssm put-parameter --name /aae/dev/google-client-secret \
     --type SecureString --value '<client secret>'
   ```

4. Set `google_client_id` in the `TFVARS` secret. The id is not secret; it is in
   every sign-in URL.

To rotate the secret, run `put-parameter` again with `--overwrite` and re-run the
deploy.

### Where secrets live

Nothing secret goes in the `TFVARS` secret or in the repository. The session key
and the Cognito client secret are generated by Terraform, and the Google secret is
put in by hand, all as SecureString parameters in Parameter Store.

Terraform's state holds a copy of each, because the resources that use them store
them there. So the state bucket is the real boundary: it is private, encrypted and
versioned, and only the CI roles and account administrators can read it. Treat
access to it like access to the secrets themselves.

### Email through SES

Cognito's own sender stops at about 50 emails a day. SES has no such limit, but
takes two applies and a request to AWS:

1. Set `email_domain = "accessibilityenhancer.com"` and apply. Add the three
   records from `terraform output email_dkim_records` in Cloudflare as **CNAME,
   DNS only** (not proxied).
2. When the SES console shows the domain as verified, set
   `email_domain_verified = true` and apply again. Verification and reset emails
   now come from `no-reply@accessibilityenhancer.com`.
3. A new SES account is in the **sandbox** and can only send to addresses it has
   verified. Request production access in the SES console, under **Account
   dashboard**, before real people sign up.

### Trying sign-in locally

Set `cognito_extra_callback_urls = ["http://localhost:5000/auth/callback"]` and
`cognito_extra_logout_urls = ["http://localhost:5000/"]`. Cognito allows http only
for localhost.

---

## Restoring from a backup

Both need someone with more than the CI roles have: they are denied reading
decks, and restoring one means reading it.

### A file

A deleted file is hidden behind a delete marker. Removing the marker brings the
previous version back:

```bash
# find the delete marker for the key
aws s3api list-object-versions --bucket <bucket> --prefix <key>

# delete the marker itself (its VersionId, with IsLatest true)
aws s3api delete-object --bucket <bucket> --key <key> --version-id <marker-version-id>
```

An overwritten file comes back by copying the older version over the current
one:

```bash
aws s3api copy-object --bucket <bucket> --key <key>   --copy-source "<bucket>/<key>?versionId=<old-version-id>"
```

This only works within 3 days of the delete or overwrite.

### The submissions table

A point-in-time restore always makes a **new** table; it never rewinds the
existing one:

```bash
aws dynamodb restore-table-to-point-in-time   --source-table-name aae-dev-submissions   --target-table-name aae-dev-submissions-restored   --restore-date-time <ISO 8601 time, before the damage>
```

The new table does not get the original's time to live, point-in-time recovery
or tags; those come from Terraform. Then either copy the items you need back into
`aae-dev-submissions`, or, to swap the whole table, delete the original, restore
again under the original name, and run `terraform apply` so Terraform puts its
settings back on it.

---

## CI

`modules/github_oidc` creates two roles that GitHub Actions assumes through
OIDC, so no AWS access key is ever stored in the repository or handed to a
developer. GitHub mints a short-lived token for the workflow run and AWS trades
it for temporary credentials.

- **The plan role** (`github_actions_role_arn`) can plan but not apply. It gets
  `ReadOnlyAccess` plus the S3 writes the backend needs to take and release its
  lock. Any workflow in the repository may assume it.
- **The apply role** (`github_actions_apply_role_arn`) can push images and run
  `terraform apply`. Only a job using the repository's `production` environment
  may assume it, so pull requests and branches never can. It gets
  `PowerUserAccess` and IAM read access, and may only create, change and pass
  roles whose names start with `aae-dev-`.

Both are explicitly denied `s3:GetObject` and `s3:GetObjectVersion` on the
uploads and processed buckets,
because their broad policies would otherwise let a workflow download real
course material, and neither has a reason to read a deck.

A role that can make roles can, in effect, give itself more, so the apply role
is only as safe as the `production` environment. Required reviewers on it are
what stop a merge from deploying without someone saying yes.

[`.github/workflows/ci.yml`](../.github/workflows/ci.yml) runs:

| Job | When | What |
|---|---|---|
| `terraform-check` | every push and pull request | `terraform fmt -check` and `terraform validate`, with no AWS access |
| `plan` | pull requests from this repository | `terraform plan`, written to the job summary |
| `deploy` | pushes to `main`, after approval | builds and pushes both images, then `terraform apply` |

`plan` and `deploy` are skipped until the repository is set up for them.

### Setting CI up

After the first `terraform apply` by hand (above), in the repository's
**Settings**:

1. **Environments → New environment** named `production`. Add the people who
   may approve a deploy as required reviewers, and limit it to the `main`
   branch.
2. **Secrets and variables → Actions → Variables**:
   - `AWS_PLAN_ROLE_ARN`: `terraform output -raw github_actions_role_arn`
   - `AWS_APPLY_ROLE_ARN`: `terraform output -raw github_actions_apply_role_arn`
   - `TF_STATE_BUCKET`: the bucket from step 1
   - `AWS_REGION`: only if it is not `us-east-1`
3. **Secrets and variables → Actions → Secrets**: `TFVARS`, holding the whole
   of your `terraform.tfvars`. It is a secret because it carries the alarm
   email addresses.
4. **Branches**: protect `main` so a change needs a pull request, a review and
   passing checks.

Developers then need nothing from AWS to ship a change. They open a pull
request, it is reviewed and merged, and the deploy waits for an approver.

---

## Giving people access to the AWS console

Use IAM Identity Center rather than IAM users, so nobody holds a long-lived
access key. If the account belongs to the university, ask UW-IT first: it may
already sign people in to AWS with their NetID, and then you only ask for them
to be assigned.

Otherwise, in IAM Identity Center, create a group for each kind of person and
give each a permission set on this account:

- **Viewers**: `ReadOnlyAccess`, plus an inline policy denying `s3:GetObject` and
  `s3:GetObjectVersion` on the uploads and processed buckets, for the same reason as the CI roles.
- **Developers**: viewer access plus what investigating takes, such as
  CloudWatch Logs Insights queries. Deploying goes through CI, not the console.
- **Administrators**: one or two people.

Everyone signs in at the Identity Center portal and gets short-lived
credentials for the group they are in.

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
- **`image_tag` defaults to `latest`.** CI always passes the commit, so this
  only matters for an apply by hand.
- **No `terraform plan` has been run against these files.** They were written
  without an AWS account to plan against. They pass `terraform validate` with
  the AWS provider 6.73, and the `terraform-check` CI job keeps checking that,
  but only a real plan proves the IAM is right.
- **The Flask app does not use any of this.** It still processes files on
  whichever machine it runs on. Connecting it to the buckets is separate work.
