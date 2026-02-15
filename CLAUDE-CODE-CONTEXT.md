# DMARC-to-Sentinel Pipeline — Claude Code Context

You're working on **DMARC-Analyzer-Azure**, an open-source DMARC aggregate report ingestion pipeline for Azure. Read `docs/ARCHITECTURE.md` for the full design.

## What it does

DMARC XML reports arrive at an Exchange Online shared mailbox → Microsoft Graph change notifications trigger an Azure Function via Event Grid → the Function fetches the email, extracts/decompresses attachments (ZIP/GZ/XML), parses the DMARC XML into flat records, and sends them to a Log Analytics custom table via the Logs Ingestion API → an Azure Monitor Workbook provides visualization.

## Key architecture decisions

- **PowerShell Azure Functions** (7.4 runtime) — no Python, no C#.
- **No external PowerShell modules** — all API calls use raw `Invoke-RestMethod` for reliability.
- **System-assigned Managed Identity** — no app registration, no client secrets.
- **Exchange Online Application RBAC** (not legacy Application Access Policies) — scopes Graph permissions to the DMARC shared mailbox only via Management Scope.
- **DCR with `kind: Direct`** — built-in `logsIngestion` endpoint, no separate DCE needed.
- **Event Grid delivery** for Graph change notifications (not webhooks) — built-in retry and dead-letter.
- **Bicep** for all infrastructure-as-code.
- Multiple domains supported by default (domain comes from the XML).

## Repo structure

```
src/function/                    # PowerShell Azure Function App
  modules/DmarcHelpers.psm1      # Shared: MI tokens, Graph calls, XML parsing, log ingestion
  DmarcReportProcessor/           # Event Grid trigger — processes new reports
  RenewGraphSubscription/         # Timer (2h) — renews Graph subscription
  CatchupProcessor/               # Timer (daily) — processes missed unread messages
  profile.ps1, host.json, requirements.psd1

infra/                           # Bicep templates
  main.bicep                      # LAW, custom table, DCR, storage, Function App, RBAC
  main.bicepparam                 # Parameter file

workbook/
  dmarc-workbook.json             # 12-panel Azure Monitor Workbook

scripts/
  Grant-MIExchangeRBAC.ps1        # One-time: Exchange RBAC + Graph app roles for MI
  New-GraphSubscription.ps1        # One-time: create Graph→Event Grid subscription
```

## Log Analytics schema

`DMARCReports_CL` — 32 columns, flat (one row per `<record>` in the DMARC XML). Report metadata and policy are denormalized across records. Primary DKIM/SPF results as individual columns; full auth results as JSON string columns for drill-down.

## Current state

All core files are in place. The project still needs:
- README.md (will be written last)
- Testing with real DMARC reports
- Potential refinements based on testing (Event Grid payload structure, error edge cases)
- CONTRIBUTING.md
- Possibly a PREREQUISITES.md with step-by-step setup instructions

## Style preferences

- PowerShell for all code and scripts
- Bicep for infrastructure (not ARM JSON)
- No unnecessary external dependencies
- Security-first: minimal permissions, scoped access
- Clear inline comments explaining "why", not just "what"
