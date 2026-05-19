# agent365-active-users-report

PowerShell reporting utility that produces an HTML report of Agent 365 (Microsoft 365 Copilot) active users, split into **licensed** and **unlicensed** active users for a configurable reporting period.

## What it does

- Pulls top-level Agent 365 usage summary metrics from Microsoft Graph.
- Builds active-user lists from the Unified Audit Log.
- Classifies each active user as licensed or unlicensed based on the tenant's subscribed SKUs.
- Generates a self-contained HTML report with tabbed user tables.
- Writes console progress and a persistent execution log.

## Requirements

- PowerShell 7 or later (`pwsh`).
- Network access to `graph.microsoft.com` and Exchange Online.
- A sign-in account (interactive) with permission to:
  - Read Microsoft 365 usage reports (Microsoft Graph scope `Reports.Read.All`).
  - Read user and organization data (`User.Read.All`, `Organization.Read.All`).
  - Search the Unified Audit Log in Exchange Online (e.g. **View-Only Audit Logs** or **Audit Logs** role).

The script installs the required Microsoft Graph and Exchange Online modules on first run (per-user scope).

## Quick start

```powershell
pwsh ./Get-Agent365ActiveUsers.ps1
```

You will be prompted to sign in to Microsoft Graph and (on first run) to Exchange Online. When the run completes, the HTML report path and log path are printed to the console.

## Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-Period` | `D30` | Reporting window: `D7`, `D30`, `D90`, `D180`, or `ALL`. |
| `-CopilotSkuPartNumbers` | `MICROSOFT_365_COPILOT`, `E7` | Exact SKU part numbers treated as Agent 365 licenses. |
| `-CopilotSkuPartNumberPatterns` | `*MICROSOFT_365_COPILOT*`, `*E7*`, `*AGENT*`, `*COPILOT*` | Wildcard patterns also matched against subscribed SKU part numbers. |
| `-AuditOperations` | `CopilotInteraction` | Unified Audit Log operations used to identify active users. |
| `-AuditResultSize` | `5000` | Max audit records returned per query (1–5000). |
| `-ReportPath` | `./Agent365-ActiveUsers-Report.html` | Output path for the HTML report. |
| `-LogPath` | `./Agent365-ActiveUsers.log` | Output path for the execution log. |
| `-NoProgress` | _off_ | Suppress progress bars. |
| `-VerboseLog` | _off_ | Log per-user classification details to console and log. |
| `-ReturnRaw` | _off_ | Emit a JSON object with summary + user lists instead of writing the HTML report and tables. |

## Common usage

```powershell
# 7-day report
pwsh ./Get-Agent365ActiveUsers.ps1 -Period D7

# custom report and log paths
pwsh ./Get-Agent365ActiveUsers.ps1 -ReportPath ./reports/agent365.html -LogPath ./logs/agent365.log

# no progress bars (good for CI / unattended)
pwsh ./Get-Agent365ActiveUsers.ps1 -NoProgress

# per-user verbose logging
pwsh ./Get-Agent365ActiveUsers.ps1 -VerboseLog

# override Agent 365 license matching
pwsh ./Get-Agent365ActiveUsers.ps1 `
  -CopilotSkuPartNumbers 'MICROSOFT_365_COPILOT','E7' `
  -CopilotSkuPartNumberPatterns '*MICROSOFT_365_COPILOT*','*E7*','*AGENT*','*COPILOT*'

# raw JSON output (no HTML, no console tables)
pwsh ./Get-Agent365ActiveUsers.ps1 -ReturnRaw
```

## Outputs

- `Agent365-ActiveUsers-Report.html` — tabbed HTML report (summary cards + licensed/unlicensed tables + assumptions).
- `Agent365-ActiveUsers.log` — execution log with timestamps.

Both files are ignored by `.gitignore` and never committed.

## Notes and caveats

- Licensed/unlicensed classification depends entirely on the configured SKU part numbers and patterns matching a SKU in your tenant. If no SKUs match, the run stops with an error.
- The active-user identity list comes from the Unified Audit Log; tenants that log different operation names should override `-AuditOperations`.
- Summary metrics come from Microsoft Graph Copilot reports; user detail lists come from audit records plus Graph user lookups. Timing/latency between sources can produce minor discrepancies.
- Results depend on Unified Audit Log retention and on the permissions of the signed-in account.

## Development

A GitHub Actions workflow at `.github/workflows/lint-powershell.yml` runs [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) on every push and pull request that touches a `*.ps1` file.
