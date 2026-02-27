# DMARC Analyzer Azure

Azure-native DMARC aggregate report analyzer using Azure Functions, Event Grid, and Log Analytics.

## Overview

This solution automatically processes **DMARC Aggregate Reports (RUA)** from your email infrastructure. It provides real-time analysis and visualization of email authentication results, helping you monitor SPF, DKIM, and DMARC compliance, detect spoofing attempts, and progress toward stricter DMARC policies.

### What are DMARC Reports?

DMARC (Domain-based Message Authentication, Reporting & Conformance) defines two types of reports:

#### RUA - Aggregate Reports (Supported ✓)

**This tool processes RUA (Aggregate Reports).**

RUA reports provide statistical summaries of email authentication results:
- **Content**: Aggregate authentication data including pass/fail counts, source IPs, SPF/DKIM/DMARC results, and policy disposition
- **Format**: XML files, typically compressed (ZIP/GZIP), sent as email attachments
- **Frequency**: Usually sent once per day by receiving mail servers (Google, Microsoft, Yahoo, etc.)
- **Volume**: One report per sending organization per day per domain
- **Use Case**: Monitoring overall email authentication health, identifying legitimate vs. unauthorized senders, policy tuning

To receive aggregate reports, configure the `rua=` tag in your DMARC DNS record:
```dns
_dmarc.example.com. IN TXT "v=DMARC1; p=none; rua=mailto:dmarc@example.com"
```

#### RUF - Forensic/Failure Reports (Not Supported)

**This tool does NOT process RUF (Forensic/Failure Reports).**

RUF reports provide per-message failure details:
- **Content**: Full headers and sometimes message bodies of individual failed messages
- **Format**: Various formats (ARF, AFRF), sent as individual emails per failure
- **Frequency**: Real-time, one report per authentication failure
- **Volume**: Can be very high for domains with many failures
- **Privacy Concerns**: Contains PII and message content, raising GDPR/privacy issues
- **Industry Trend**: Increasingly deprecated by major providers (Google stopped sending in 2023, Microsoft never implemented) due to privacy and volume concerns

**Recommendation**: Focus on `rua=` configuration only. RUF is largely obsolete and not necessary for DMARC monitoring and policy enforcement.

## Features

- **Real-time Processing**: Automatic ingestion via Microsoft Graph change notifications and Event Grid
- **Rich Analytics**: Azure Monitor Workbook with comprehensive visualizations
- **GeoIP Mapping**: Geographic distribution of email sources with pass/fail rates
- **ASN Enrichment**: Identify mail providers by Autonomous System Number
- **Subdomain Discovery**: Track email from all subdomains to prevent shadow IT
- **Threat Detection**: Identify spoofing attempts and suspicious source IPs
- **Compliance Tracking**: Monitor SPF, DKIM, and DMARC pass rates per domain
- **Policy Guidance**: Built-in recommendations for DMARC policy progression
- **Zero Secrets**: Uses Azure Managed Identity for authentication (no client secrets)
- **Scalable & Cost-Effective**: Serverless architecture, typically under $1/month

## Architecture

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for detailed architecture, data flow, and security model.

## Prerequisites

- Azure subscription with permissions to create resources
- Microsoft 365 tenant with Exchange Online
- A shared mailbox to receive DMARC reports (e.g., `dmarc@example.com`)
- Azure CLI
- PowerShell 7.4+ (for setup scripts)
- [Azure Functions Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local) (`func` CLI for publishing)
- PowerShell modules for Step 3 (install once):
  ```powershell
  Install-Module Az.Accounts -Scope CurrentUser -Force
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser -Force
  Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
  ```

## Deployment

### 1. Deploy Azure Resources

First, generate the Graph client state secret and set it as an environment variable. This secret validates Graph change notifications. Using `readEnvironmentVariable()` in the `.bicepparam` file ensures the secret is never stored in source control.

**PowerShell:**
```powershell
# Generates a random GUID in-memory — safe even with PowerShell Script Block Logging enabled
$env:GRAPH_CLIENT_STATE = [guid]::NewGuid().ToString()
```

**Bash:**
```bash
export GRAPH_CLIENT_STATE=$(cat /proc/sys/kernel/random/uuid)
```

Then create a new resource group:

**PowerShell:**
```powershell
az group create --name "rg-dmarc-prod" --location "eastus"
```

**Bash:**
```bash
az group create --name "rg-dmarc-prod" --location "eastus"
```

Then deploy:

**PowerShell:**
```powershell
az deployment group create `
  --resource-group "rg-dmarc-prod" `
  --template-file infra/main.bicep `
  --parameters infra/main.bicepparam
```

**Bash:**
```bash
az deployment group create \
  --resource-group "rg-dmarc-prod" \
  --template-file infra/main.bicep \
  --parameters infra/main.bicepparam
```

Or use the Azure Portal to deploy `infra/main.bicep` with custom parameters.

**Key Parameters:**
- `baseName`: Base name for all resources (e.g., "dmarc")
- `location`: Azure region
- `mailboxUserId`: Object ID of the shared mailbox user (required)
- `existingWorkspaceId`: (Optional) Use an existing Log Analytics Workspace
- `deployPartnerConfig`: (Optional, default `true`) Deploy Event Grid partner configuration for Microsoft Graph API. Set to `false` if you already have a partner configuration in the resource group.

The Bicep deployment automatically authorizes Microsoft Graph API as an Event Grid partner — no manual Portal steps required.

### 2. Publish Function Code

Publish the function app code (including the `SetupHelper` function used in the next step):

```bash
cd src/function
func azure functionapp publish <function-app-name> --powershell
```

The function app name is shown in the Bicep deployment output as `functionAppName`.

### 3. Grant Exchange RBAC Permissions

Grant the Function's Managed Identity permissions to access the shared mailbox. **This must be done before creating the Graph subscription.**

Requires PowerShell modules: `Az.Accounts`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `ExchangeOnlineManagement` (see Prerequisites).

**PowerShell:**
```powershell
./scripts/Grant-MIExchangeRBAC.ps1 `
  -FunctionAppName "dmarc-func-xyz123" `
  -ResourceGroupName "rg-dmarc-prod" `
  -MailboxAddress "dmarc@example.com"
```

This assigns Graph `Mail.Read`/`Mail.ReadWrite` app roles and creates an Exchange Application RBAC policy scoped to the DMARC mailbox only.

### 4. Create Graph Subscription

Run the setup script to create the Graph change notification subscription and wire up the Event Grid pipeline.

**PowerShell:**
```powershell
./scripts/New-GraphSubscription.ps1 `
  -FunctionAppName "dmarc-func-xyz123" `
  -ResourceGroupName "rg-dmarc-prod" `
  -SubscriptionId "11111111-1111-1111-1111-111111111111"
```

**Bash:**
```bash
./scripts/New-GraphSubscription.sh \
  --function-app "dmarc-func-xyz123" \
  --resource-group "rg-dmarc-prod" \
  --subscription "11111111-1111-1111-1111-111111111111"
```

The script automates:
1. Invokes the `SetupHelper` function (uses the MI to create the Graph subscription)
2. Waits for the Partner Topic, activates it, and creates an Event Subscription
3. Saves the subscription ID to the Function App settings

### 5. Import Existing DMARC Reports (Optional)

If the DMARC mailbox already contains reports (e.g., you had DNS configured before deploying this solution), you can import them using the `BackfillProcessor` function. This is an admin-secured HTTP endpoint — no public access.

**Retrieve the master host key securely via ARM and trigger the backfill:**

**PowerShell:**
```powershell
# Get the master key from ARM (no secrets in shell history)
$keys = az rest --method POST `
  --uri "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Web/sites/<function-app-name>/host/default/listkeys?api-version=2024-04-01" `
  --query masterKey -o tsv

# Import the last 7 days of unread messages
az rest --method POST `
  --url "https://<function-app-name>.azurewebsites.net/api/BackfillProcessor?days=7" `
  --headers "x-functions-key=$keys" `
  --skip-authorization-header
```

**Bash:**
```bash
# Get the master key from ARM
KEY=$(az rest --method POST \
  --uri "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Web/sites/<function-app-name>/host/default/listkeys?api-version=2024-04-01" \
  --query masterKey -o tsv)

# Import the last 7 days of unread messages
curl -s -X POST "https://<function-app-name>.azurewebsites.net/api/BackfillProcessor?days=7" \
  -H "x-functions-key: $KEY"
```

**Parameters:**

| Parameter | Default | Description |
|-----------|---------|-------------|
| `days` | `7` | How far back to look (1–365) |
| `includeRead` | `false` | Set to `true` to re-process already-read messages |

The function returns a JSON summary with `processed`, `failed`, and `skipped` counts. Messages are marked as read after processing, so running it twice is safe — already-processed messages are skipped by default.

> **Tip:** If you have more than 7 days of historical reports, increase the window: `?days=30&includeRead=true`

### 6. Configure DMARC DNS Records

Add or update the `_dmarc` TXT record for your domain to send reports to your shared mailbox:

```dns
_dmarc.example.com. IN TXT "v=DMARC1; p=none; rua=mailto:dmarc@example.com; pct=100; sp=none"
```

**Important Notes:**
- Use `rua=mailto:...` to specify the aggregate report recipient (this tool)
- Do NOT configure `ruf=mailto:...` — forensic reports are not supported and not recommended
- Start with `p=none` (monitor mode) and progressively move to `p=quarantine` or `p=reject` as compliance improves
- Set `pct=100` to ensure all traffic is reported

### 7. Import the Workbook

1. Navigate to your Log Analytics Workspace in the Azure Portal
2. Select **Workbooks** > **+ New**
3. Click the **Advanced Editor** button (</> icon)
4. Paste the contents of `workbook/dmarc-workbook.json`
5. Click **Apply**
6. Save the workbook with a name like "DMARC Analytics"

## Alerting & Notifications

Azure Monitor Alert Rules provide automated notifications for critical DMARC events — the Azure-native equivalent of email alerts from commercial DMARC tools.

### Setting Up Alert Rules

1. **Navigate to your Log Analytics Workspace** → **Alerts** → **Create alert rule**
2. **Select a signal type**: "Custom log search"
3. **Enter a KQL query** (see examples below)
4. **Configure alert logic**: Threshold, frequency, evaluation period
5. **Add action groups**: Email, SMS, webhook, etc.

### Example Alert Queries

#### Alert on New Source IPs

Detect when a new source IP is seen for the first time in the last 24 hours:

```kql
let known = DMARCReports_CL
| where TimeGenerated between (ago(30d) .. ago(1d))
| distinct SourceIP;
DMARCReports_CL
| where TimeGenerated > ago(1d)
| where SourceIP !in (known)
| summarize
    Messages = sum(MessageCount),
    Domains = make_set(Domain, 10),
    FirstSeen = min(TimeGenerated)
    by SourceIP
| where Messages > 10  // Threshold: only alert if > 10 messages
| project SourceIP, Messages, Domains, FirstSeen
```

**Recommendation**: Run every 1 hour, alert when result count > 0

#### Alert on Authentication Failure Spike

Detect when the failure rate exceeds 10%:

```kql
DMARCReports_CL
| where TimeGenerated > ago(1h)
| summarize
    Total = sum(MessageCount),
    Failed = sumif(MessageCount, PolicyEvaluated_dkim == 'fail' and PolicyEvaluated_spf == 'fail')
| extend FailRate = 100.0 * Failed / Total
| where FailRate > 10.0  // Alert if failure rate > 10%
| project Total, Failed, FailRate
```

**Recommendation**: Run every 1 hour, alert when result count > 0

#### Alert on DMARC Pass Rate Drop

Detect when overall pass rate drops below 95%:

```kql
DMARCReports_CL
| where TimeGenerated > ago(6h)
| summarize
    Total = sum(MessageCount),
    Passed = sumif(MessageCount, PolicyEvaluated_dkim == 'pass' or PolicyEvaluated_spf == 'pass')
| extend PassRate = 100.0 * Passed / Total
| where PassRate < 95.0  // Alert if pass rate drops below 95%
| project Total, Passed, PassRate
```

**Recommendation**: Run every 6 hours, alert when result count > 0

### Action Groups

Configure **Action Groups** to define how you're notified:
- **Email/SMS**: Notify security or operations team
- **Webhook**: Integrate with Microsoft Teams, Slack, PagerDuty, etc.
- **Azure Function/Logic App**: Custom remediation workflows

For more on Azure Monitor Alerts, see: [Microsoft Docs - Create log alert rules](https://learn.microsoft.com/en-us/azure/azure-monitor/alerts/alerts-create-log-alert-rule)

## Workbook Capabilities

The Azure Monitor Workbook (`workbook/dmarc-workbook.json`) is organized into 5 tabs with 43 KQL visualizations:

### Executive Summary
- Total message volume, unique source IPs, reporters, pass rate (KPI tiles)
- Week-over-week comparison with deltas
- Per-domain risk scoring (weighted compliance score with color-coded risk levels)
- Daily pass/fail trend and compliance rate trend
- Priority action items with categorized fix recommendations

### Authentication
- Combined SPF/DKIM/DMARC authentication matrix
- SPF and DKIM pass rate trends over time
- Failure reason analysis (forwarding, mailing lists, local policy overrides)
- Policy override reasons and forwarding detection
- Envelope vs header from mismatch analysis
- DKIM selector-level detail and alignment mode breakdown

### Sources & Senders
- **Top 50 Source IPs** with pass rates, ASN/organization info, and IP reputation links
- Authorized vs unknown sender classification
- Volume anomaly detection (Z-score based)
- **GeoIP Map** showing geographic distribution of email sources
- **Suspicious IPs** (both SPF and DKIM failing)
- Top failing senders with remediation hints
- Volume by sending identity (HeaderFrom)

### Domain Compliance
- Per-domain compliance scoring and pass rates
- Domain readiness for policy enforcement progression
- Published DMARC policies
- Disposition over time (none/quarantine/reject)
- **Subdomain discovery** (identify shadow IT or unauthorized subdomains)
- DKIM selector rotation tracking
- BIMI readiness assessment

### Reporting & Ops
- Report freshness monitor (color-coded per reporter)
- Reporter coverage and missing expected reporters
- Reports by provider and report volume over time
- Azure Monitor alert rule templates (5 ready-to-use KQL queries)
- Policy guidance for DMARC policy progression

## Troubleshooting

### No Data in Workbook

1. **Check Function App logs**: Navigate to Function App → **Functions** → **DmarcReportProcessor** → **Monitor**
2. **Verify Graph subscription**: Check Function App → **Configuration** → Confirm `GRAPH_SUBSCRIPTION_ID` is set
3. **Check shared mailbox**: Ensure DMARC reports are arriving (may take 24-48 hours after DNS change)
4. **Check Event Grid delivery**: Navigate to Partner Topic → **Metrics** → Check for failed deliveries

### Reports Not Processing

1. **Verify Managed Identity permissions**:
   - Graph API: `Mail.Read`, `Mail.ReadWrite` app roles
   - Exchange RBAC: Application permissions scoped to mailbox
   - Log Analytics: `Monitoring Metrics Publisher` on DCR
2. **Check Graph subscription status**: Should be "enabled" and not expired
3. **Review Function logs** for errors

### Permission Errors

If you see errors like "Access Denied" or "Insufficient privileges":
- Ensure the Managed Identity has been granted permissions (may take 5-10 minutes to propagate)
- Verify Exchange RBAC scope is configured correctly
- Check that the MI has `Monitoring Metrics Publisher` on the DCR

## Security Considerations

- **No Client Secrets**: Uses Azure Managed Identity exclusively
- **Least Privilege**: Exchange RBAC limits MI to the shared mailbox only
- **Key Vault Integration**: Secrets (Graph client state) stored in Key Vault
- **DTD Attack Protection**: XML parsing explicitly disables DTD processing
- **Input Validation**: All report data sanitized before ingestion
- **RBAC**: All Azure resources secured via Azure RBAC

## Testing

The repository includes comprehensive Pester tests to verify PowerShell scripts work as expected:

```powershell
# Run all tests
Invoke-Pester -Path ./tests

# Run with detailed output
Invoke-Pester -Path ./tests -Output Detailed
```

### Test Coverage
- ✅ **147 tests** covering all PowerShell scripts
- ✅ Module functions (token acquisition, Graph API, DMARC XML parsing, attachment extraction)
- ✅ Setup scripts (New-GraphSubscription.ps1, Grant-MIExchangeRBAC.ps1)
- ✅ Azure Functions (DmarcReportProcessor, RenewGraphSubscription, CatchupProcessor, BackfillProcessor)
- ✅ Security validations (DTD protection, size limits, client state validation)
- ✅ Error handling and logging patterns

See [tests/README.md](tests/README.md) for detailed test documentation.

## Contributing

Contributions are welcome! Please submit issues or pull requests via GitHub.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Microsoft Graph API and Event Grid teams
- Azure Monitor and Log Analytics teams
- The DMARC community

## Support

For issues, questions, or feature requests, please open a GitHub issue.
