# Architecture

## Data Flow

```
DMARC Aggregate Reports (ZIP/GZ containing XML)
  │ SMTP delivery
  ▼
Exchange Online Shared Mailbox
  │ Microsoft Graph Change Notification
  ▼
Azure Event Grid (Partner Topic)
  │ Event Subscription (CloudEvents v1.0)
  ▼
Azure Function: DmarcReportProcessor (Event Grid trigger)
  ├─ Graph API → Fetch message + attachments (Managed Identity token)
  ├─ Decompress ZIP/GZ → Extract XML
  ├─ Parse DMARC XML → Flat records (one per <record> element)
  ├─ Logs Ingestion API → POST to DCR endpoint (Managed Identity token)
  └─ Graph API → Mark message as read
  │
  ▼
Log Analytics Workspace (DMARCReports_CL custom table)
  │
  ▼
Azure Monitor Workbook / Microsoft Sentinel Workbook
```

## Functions

| Function | Trigger | Purpose |
|---|---|---|
| `DmarcReportProcessor` | Event Grid | Real-time processing of new DMARC reports |
| `RenewGraphSubscription` | Timer (every 2h) | Keeps Graph subscription alive (max 4230 min) |
| `CatchupProcessor` | Timer (daily 06:00 UTC) | Processes any missed unread reports |

## Authentication & Security

**No app registration or client secrets.** The Function App uses a **system-assigned Managed Identity** for all API calls:

- **Microsoft Graph**: MI is granted `Mail.Read` + `Mail.ReadWrite` app roles, scoped to the DMARC shared mailbox only via Exchange Online Application RBAC (Management Scope).
- **Logs Ingestion API**: MI is granted `Monitoring Metrics Publisher` role on the DCR via Azure RBAC.

### Exchange Application RBAC (replaces Application Access Policies)

Instead of the legacy Application Access Policy (which is being deprecated), this solution uses the modern Exchange Online Application RBAC:

1. A **Management Scope** restricts access to the DMARC shared mailbox only.
2. **Application roles** (`Application Mail.Read`, `Application Mail.ReadWrite`) are assigned to the MI's service principal, scoped to the management scope.
3. The MI **cannot access any other mailbox** in the tenant.

## Graph Change Notifications via Event Grid

Uses Event Grid as the delivery mechanism (not webhooks):

- No public webhook endpoint to validate or secure
- Built-in retry logic (up to 24 hours)
- Dead-letter storage for failed deliveries
- Partner topic auto-created by the Graph subscription

The Graph subscription watches: `users/{mailboxUserId}/mailFolders('Inbox')/messages`

## Log Analytics Schema

The `DMARCReports_CL` table uses a flat schema — one row per `<record>` element in the DMARC XML. Report metadata and policy are repeated across records from the same report.

### Columns (32)

| Column | Type | Source |
|---|---|---|
| TimeGenerated | datetime | Ingestion time |
| ReportOrgName | string | `report_metadata/org_name` |
| ReportEmail | string | `report_metadata/email` |
| ReportExtraContactInfo | string | `report_metadata/extra_contact_info` |
| ReportId | string | `report_metadata/report_id` |
| ReportDateRangeBegin | datetime | `report_metadata/date_range/begin` (epoch→ISO) |
| ReportDateRangeEnd | datetime | `report_metadata/date_range/end` (epoch→ISO) |
| Domain | string | `policy_published/domain` |
| PolicyPublished_p | string | `policy_published/p` |
| PolicyPublished_sp | string | `policy_published/sp` |
| PolicyPublished_pct | int | `policy_published/pct` |
| PolicyPublished_adkim | string | `policy_published/adkim` |
| PolicyPublished_aspf | string | `policy_published/aspf` |
| PolicyPublished_fo | string | `policy_published/fo` |
| SourceIP | string | `record/row/source_ip` |
| MessageCount | int | `record/row/count` |
| PolicyEvaluated_disposition | string | `record/row/policy_evaluated/disposition` |
| PolicyEvaluated_dkim | string | `record/row/policy_evaluated/dkim` |
| PolicyEvaluated_spf | string | `record/row/policy_evaluated/spf` |
| PolicyEvaluated_reason_type | string | `record/row/policy_evaluated/reason/type` |
| PolicyEvaluated_reason_comment | string | `record/row/policy_evaluated/reason/comment` |
| HeaderFrom | string | `record/identifiers/header_from` |
| EnvelopeFrom | string | `record/identifiers/envelope_from` |
| EnvelopeTo | string | `record/identifiers/envelope_to` |
| DkimResult | string | First `auth_results/dkim/result` |
| DkimDomain | string | First `auth_results/dkim/domain` |
| DkimSelector | string | First `auth_results/dkim/selector` |
| SpfResult | string | First `auth_results/spf/result` |
| SpfDomain | string | First `auth_results/spf/domain` |
| SpfScope | string | First `auth_results/spf/scope` |
| DkimAuthResults | string | Full DKIM results as JSON array |
| SpfAuthResults | string | Full SPF results as JSON array |

## Cost Estimate (typical mid-size org)

| Resource | Monthly Cost |
|---|---|
| Azure Function (Consumption plan) | ~$0 (1M executions/month free) |
| Event Grid | ~$0 (100K operations/month free) |
| Log Analytics (~100-300 MB/month) | ~$0 (5 GB/month free) |
| Storage Account | ~$0.01 |
| **Total** | **~$0** |
