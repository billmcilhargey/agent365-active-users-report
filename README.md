# agent365-active-users-report

PowerShell reporting utility for Agent 365 usage in Microsoft 365.

## What This Script Does

- Pulls top-level Agent 365 usage summary metrics.
- Builds detailed active-user lists split by:
  - Licensed Agent 365 users
  - Unlicensed active users
- Generates an HTML report with tabbed user tables.
- Writes real-time console progress and a persistent execution log.

## Script

- `Get-Agent365ActiveUsers.ps1`

## Quick Start

```powershell
pwsh ./Get-Agent365ActiveUsers.ps1
```

## Prerequisites

- PowerShell 7+
- Permissions/roles to read Microsoft 365 usage reports (`Reports.Read.All`) and query Unified Audit Log
- Access to connect Microsoft Graph and Exchange Online PowerShell modules

## Common Options

```powershell
# 7-day report
pwsh ./Get-Agent365ActiveUsers.ps1 -Period D7

# custom report and log paths
pwsh ./Get-Agent365ActiveUsers.ps1 -ReportPath ./reports/agent365.html -LogPath ./logs/agent365.log

# disable progress bars
pwsh ./Get-Agent365ActiveUsers.ps1 -NoProgress

# enable per-user verbose logging to console and log file
pwsh ./Get-Agent365ActiveUsers.ps1 -VerboseLog

# control audit query size (1-5000)
pwsh ./Get-Agent365ActiveUsers.ps1 -AuditResultSize 5000

# override Agent 365 license matching (exact SKU part numbers + wildcard patterns)
pwsh ./Get-Agent365ActiveUsers.ps1 \
  -CopilotSkuPartNumbers "MICROSOFT_365_COPILOT","E7" \
  -CopilotSkuPartNumberPatterns "*MICROSOFT_365_COPILOT*","*E7*","*AGENT*","*COPILOT*"
```

## Notes

- Licensed/unlicensed classification is based on configured Agent 365 SKU part numbers and each user's assigned licenses.
- Active-user identity list is built from Unified Audit Log operations configured in the script parameters.
