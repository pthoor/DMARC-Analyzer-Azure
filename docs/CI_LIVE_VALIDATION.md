# CI: Live Validation of Detection Rules Against Azure

`.github/workflows/detection-rules-live-validate.yml` validates `detections/*.yaml`
against a real Microsoft Sentinel workspace's own rule-schema validator - catching
things static YAML/schema checks can't, like whether Azure actually accepts the KQL
syntax and the full analytics-rule payload. This is in addition to, not instead of,
the static schema checks in `tests/DetectionRules.Tests.ps1` (which run on every PR,
including from forks, and need no Azure access at all).

This doc is maintainer setup - it is not needed to build, test, or contribute to the
project day to day.

## Safety model

This workflow authenticates to a real Azure tenant and creates/deletes real (but
disabled) Sentinel analytics rules. Given that, several things are deliberately
layered on top of each other rather than relying on any single control:

- **No stored credential.** Authentication is via OpenID Connect (federated
  credential) - Azure AD trusts short-lived tokens minted per CI run for this
  specific repo. There is no client secret sitting in GitHub to leak or rotate.
- **Never triggered by a pull request.** The workflow's `on:` block has no
  `pull_request` trigger at all - only `push` to `main` (and only when
  `detections/**` actually changed) and manual `workflow_dispatch`. A contributor's
  PR - including from a fork - can never cause this workflow to run.
- **Gated behind a GitHub Environment.** The job declares `environment:
  azure-live-validation`. Configure that environment's protection rules (below) to
  require a manual approval before the job runs, so even an automatic push to `main`
  pauses for a human click before it touches Azure.
- **Federated credential is scoped to the environment**, not just the repo - so even
  if someone edited the workflow's trigger, a run outside the `azure-live-validation`
  environment context still couldn't mint an Azure token.
- **Least-privilege role**, scoped to the one Log Analytics workspace resource, not
  the resource group or subscription.
- **Test rules never actually evaluate.** Every rule created is `enabled: false` -
  this validates schema/KQL-syntax acceptance only, never touches production data.
- **Unmistakable, fixed test IDs.** Test rules use fixed GUIDs in the pattern
  `00000000-cccc-cccc-cccc-00000000000N`, deliberately distinct from the real
  production rule IDs, with a `[CI-VALIDATION]` name prefix. If cleanup ever fails,
  an orphaned rule is instantly recognizable and safe to delete by hand.
- **Self-healing cleanup.** The script sweeps (deletes) those same fixed IDs at the
  *start* of every run before creating anything, so a previous run's failed cleanup
  can never accumulate. Deletion of what a run itself created is wrapped in
  try/finally, so cleanup still runs even if validation of a later file throws.

## One-time Azure setup

Run these yourself - none of this can be done from the repo or by an AI assistant,
since it requires your own Azure AD and subscription admin access.

```bash
# 1. Create the app registration CI will authenticate as.
APP_ID=$(az ad app create --display-name "dmarc-defender-xdr-ci-live-validate" --query appId -o tsv)
az ad sp create --id "$APP_ID"

# 2. Trust GitHub's OIDC tokens for this repo's specific environment only -
#    replace <owner>/<repo> with your fork/repo path if different.
az ad app federated-credential create \
  --id "$APP_ID" \
  --parameters '{
    "name": "github-actions-live-validate-environment",
    "issuer": "https://token.actions.githubusercontent.com",
    "subject": "repo:pthoor/DMARC-for-Defender-XDR:environment:azure-live-validation",
    "audiences": ["api://AzureADTokenExchange"]
  }'

# 3. Grant the minimum role needed - scoped to the single workspace, not the
#    resource group or subscription.
WORKSPACE_ID=$(az monitor log-analytics workspace show \
  --resource-group <YOUR_RESOURCE_GROUP> \
  --workspace-name <YOUR_WORKSPACE_NAME> \
  --query id -o tsv)

az role assignment create \
  --assignee "$APP_ID" \
  --role "Microsoft Sentinel Contributor" \
  --scope "$WORKSPACE_ID"

# 4. Collect the values you'll need for GitHub secrets (step below).
echo "AZURE_LIVE_VALIDATE_CLIENT_ID=$APP_ID"
echo "AZURE_LIVE_VALIDATE_TENANT_ID=$(az account show --query tenantId -o tsv)"
echo "AZURE_LIVE_VALIDATE_SUBSCRIPTION_ID=$(az account show --query id -o tsv)"
echo "AZURE_LIVE_VALIDATE_RESOURCE_GROUP=<YOUR_RESOURCE_GROUP>"
echo "AZURE_LIVE_VALIDATE_WORKSPACE_NAME=<YOUR_WORKSPACE_NAME>"
```

If you'd rather scope even tighter than the built-in "Microsoft Sentinel
Contributor" role (which also grants dashboard/workbook/incident management), define
a custom role limited to `Microsoft.SecurityInsights/alertRules/*` actions on that
one workspace and assign that instead.

## One-time GitHub setup

1. **Create the environment.** Repo → **Settings → Environments → New environment**,
   name it exactly `azure-live-validation` (must match the workflow's `environment:`
   and the federated credential's `subject` above).
2. **Add required reviewers (recommended).** In that environment's protection rules,
   enable **Required reviewers** and add yourself (or another trusted maintainer).
   This is what makes even an automatic push-to-`main` run pause for a manual
   approval click before it touches your tenant. Available on GitHub Free for public
   repos.
3. **Add environment secrets** (not repo-level secrets - these must be scoped to the
   `azure-live-validation` environment specifically, so no other workflow in the repo
   can read them):

   | Secret | Value |
   |---|---|
   | `AZURE_LIVE_VALIDATE_CLIENT_ID` | App registration's app ID |
   | `AZURE_LIVE_VALIDATE_TENANT_ID` | Your Azure AD tenant ID |
   | `AZURE_LIVE_VALIDATE_SUBSCRIPTION_ID` | Subscription containing the workspace |
   | `AZURE_LIVE_VALIDATE_RESOURCE_GROUP` | Resource group containing the workspace |
   | `AZURE_LIVE_VALIDATE_WORKSPACE_NAME` | Log Analytics/Sentinel workspace name |

   Via `gh` CLI (run once you're on the environment - note `--env`, not a repo
   secret):
   ```bash
   gh secret set AZURE_LIVE_VALIDATE_CLIENT_ID --env azure-live-validation --body "$APP_ID"
   gh secret set AZURE_LIVE_VALIDATE_TENANT_ID --env azure-live-validation --body "<tenant-id>"
   gh secret set AZURE_LIVE_VALIDATE_SUBSCRIPTION_ID --env azure-live-validation --body "<subscription-id>"
   gh secret set AZURE_LIVE_VALIDATE_RESOURCE_GROUP --env azure-live-validation --body "<resource-group>"
   gh secret set AZURE_LIVE_VALIDATE_WORKSPACE_NAME --env azure-live-validation --body "<workspace-name>"
   ```

## Testing it

Trigger a run manually without waiting for a `detections/**` change: repo → **Actions
→ Detection Rules Live Validation → Run workflow**. If you added a required
reviewer, approve the pending deployment when prompted.

## If something goes wrong

- **A `[CI-VALIDATION]` rule is visible in your Sentinel workspace.** This means a
  cleanup step failed (e.g. the runner was killed mid-job). It's disabled and
  harmless, but delete it - the fixed IDs are listed in
  `scripts/Test-DetectionRulesLive.ps1`'s `$testRuleIds` map, or just search the rule
  list for the `[CI-VALIDATION]` name prefix. The next scheduled run will also sweep
  it automatically.
- **The job fails with a role/permission error.** Re-check the role assignment scope
  in step 3 above - it must be on the workspace resource, not a parent scope that
  might have been revoked or never applied.
- **The job fails with an OIDC/federated-credential error.** Confirm the federated
  credential's `subject` exactly matches `repo:<owner>/<repo>:environment:azure-live-validation`
  and that the job's `environment:` value in the workflow matches too - a mismatch
  in either is the most common cause.
