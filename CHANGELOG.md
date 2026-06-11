# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Security
- `DmarcReportProcessor` now fails closed when `GRAPH_CLIENT_STATE` is missing or its Key Vault reference is unresolved (previously it logged a warning and processed the notification unvalidated), and validates the client state with a constant-time comparison.
- `SetupHelper` only accepts `EventGrid:` notification URLs by default. Direct HTTPS webhook URLs (which would receive the client state secret) require the `ALLOW_WEBHOOK_NOTIFICATION_URL=true` app setting.
- Attachment extraction enforces a cumulative 100 MB decompression cap per message (in addition to existing per-entry limits) to prevent memory exhaustion from crafted ZIP attachments.
- `BackfillProcessor` no longer returns internal exception details in HTTP 500 responses.
- Mail subjects are sanitized (control characters stripped, length capped) before logging.
- Mailbox user IDs are URL-encoded consistently in Microsoft Graph request URIs.

### Fixed
- A malformed `<record>` (e.g., non-numeric `<count>` or `<pct>`) no longer aborts processing of the whole report; the bad record is skipped (or its value defaulted) with a warning, so messages cannot get stuck unread in a retry loop.
- A ZIP entry that fails extraction no longer discards entries already extracted from the same archive, and `ZipArchive` handles are always disposed.

### Changed
- Transient HTTP 502/504 responses are retried with backoff (in addition to 429/503).
- Removed the unused `Get-UnreadMessages` helper (superseded by `Get-MailboxMessages`) and the always-zero `skipped` counter from the BackfillProcessor response.

## [1.1.0] - 2026-05-08

### Added
- Added a root `VERSION` file for tenant-deployment-safe artifact version tracking.
- Added `tests/Versioning.Tests.ps1` to validate version metadata consistency and KQL safety guards.

### Changed
- Hardened workbook metric queries against null and divide-by-zero edge cases.
- Updated workbook header to include release metadata (`v1.1.0`).
- Bumped Sentinel detection rule versions to `1.1.0`.

### Fixed
- Corrected policy override detection entity mapping to use a scalar IP field.
- Corrected pass-rate logic in Azure Monitor pass-rate alert to use DMARC semantics (SPF **or** DKIM pass).
- Corrected alert and workbook query calculations for null handling, zero denominators, and safer join-key usage.
