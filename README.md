# agent365-active-users-report

PowerShell reporting utility that produces an HTML report of Agent 365 (Microsoft 365 Copilot) active users, split into **licensed** and **unlicensed** active users for a configurable reporting period.

## What it does

- Pulls top-level usage summary metrics from Microsoft Graph (`getMicrosoft365CopilotUserCountSummary`).
- Pulls the per-user activity list from Microsoft Graph (`getMicrosoft365CopilotUsageUserDetail`) — cross-platform, works on Linux, macOS, and Windows.
- *(Optional, Windows-only)* also queries the Unified Audit Log for `CopilotInteraction` activity to surface unlicensed Copilot Chat usage that the Graph user-detail report does not include.
- Classifies each active user as **licensed** or **unlicensed** based on the tenant's subscribed Agent 365 / Copilot SKUs.
- Generates a self-contained HTML report with tabbed user tables, including a per-user **Last Activity Date** and **Source** column (`Graph`, `UAL`, or `Graph+UAL`).
- Writes console progress and a fresh execution log on every run.

## Requirements

- PowerShell 7 or later (`pwsh`).
- Network access to `graph.microsoft.com` (and Exchange Online if you opt in to `-IncludeUnifiedAuditLog`).
- A sign-in account (interactive) with permission to:
  - Read Microsoft 365 usage reports (Microsoft Graph scope `Reports.Read.All`).
  - Read user and organization data (`User.Read.All`, `Organization.Read.All`).
  - **A Microsoft Entra directory role that qualifies for the per-user Copilot usage report** — see [Required permissions](#required-permissions) below.
  - *(Only when `-IncludeUnifiedAuditLog` is used)* Search the Unified Audit Log in Exchange Online (e.g. **View-Only Audit Logs** or **Audit Logs** role).

The script installs the required modules on first run (per-user scope). `ExchangeOnlineManagement` is only installed when you opt in to `-IncludeUnifiedAuditLog`.

### Required permissions

The `getMicrosoft365CopilotUsageUserDetail` Graph endpoint has stricter access requirements than the summary endpoint. **`Reports.Read.All` scope alone is not sufficient** — the signed-in account must also hold one of these Microsoft Entra directory roles:

**Full-access roles (return per-user detail — what this script needs):**

- **Reports Reader** (least privilege — recommended)
- **AI Administrator**
- **Global Administrator** (Company Administrator)
- **Exchange Administrator**
- **SharePoint Administrator**
- **Teams Administrator** (Teams Service Administrator)
- **Teams Communications Administrator**
- **Skype for Business Administrator** (Lync Administrator)

**Tenant-only roles (insufficient — return aggregate data only, no per-user detail):**

- **Global Reader**
- **Usage Summary Reports Reader**

Per [Microsoft's authorization documentation](https://learn.microsoft.com/graph/reportroot-authorization), "Global Reader and Usage Summary Reports Reader roles will only have access to tenant-level data, without visibility into detailed metrics." The newer `/copilot/reports/` endpoint variants don't accept these roles at all. **If you hold only one of these, the per-user call returns `403 Forbidden`** — add **Reports Reader** in addition.

On every run, the script performs a directory-role pre-flight check using `/me/transitiveMemberOf` (which catches PIM activations and group-based assignments) and prints a clear `WARN` if the signed-in account holds only a tenant-only role or no qualifying role at all. The script still continues — a custom role may grant equivalent access — but the per-user call will likely return `403 Forbidden`. Pass `-SkipRoleCheck` to suppress the warning, or assign **Reports Reader** in the Microsoft Entra admin center and re-run.

## Quick start

```powershell
pwsh ./Get-Agent365ActiveUsers.ps1
```

You will be prompted to sign in to Microsoft Graph. The script tries an interactive browser sign-in by default and **automatically falls back to device code authentication** in environments where a browser cannot be launched (SSH sessions, GitHub Codespaces, dev containers, headless Linux). Use `-UseDeviceCode` to force device code mode. When the run completes, the HTML report path and log path are printed to the console.

### Device code sign-in tips

Microsoft enforces a hard **120-second** timeout from the moment a device code is generated. To make that comfortable:

1. **Open https://login.microsoft.com/device in a browser tab first.**
2. Start the script. It will prompt: *"When the sign-in page is open and ready, press Enter to receive a fresh code (attempt 1 of 3)..."*
3. Press Enter, copy the code, paste it on the device-login page, and sign in.

If a code does time out, the script automatically requests up to **3 fresh codes** before giving up — no need to re-run the whole script.

## Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-Period` | `D30` | Reporting window: `D7`, `D30`, `D90`, `D180`, or `ALL`. |
| `-CopilotSkuPartNumbers` | `MICROSOFT_365_COPILOT`, `E7` | Exact SKU part numbers treated as Agent 365 licenses. |
| `-CopilotSkuPartNumberPatterns` | `*MICROSOFT_365_COPILOT*`, `*E7*`, `*AGENT*`, `*COPILOT*` | Wildcard patterns also matched against subscribed SKU part numbers. |
| `-IncludeUnifiedAuditLog` | _off_ | **Windows-only.** Also pull `CopilotInteraction` (or `-AuditOperations`) events from the Unified Audit Log and merge with Graph results. Auto-skipped with a warning on Linux/macOS. |
| `-AuditOperations` | `CopilotInteraction` | Unified Audit Log operations used to identify active users (only applied when `-IncludeUnifiedAuditLog` is set). |
| `-AuditResultSize` | `5000` | Max audit records returned per query (1–5000). |
| `-ReportPath` | `./Agent365-ActiveUsers-Report.html` | Output path for the HTML report. |
| `-LogPath` | `./Agent365-ActiveUsers.log` | Output path for the execution log (recreated on every run). |
| `-NoProgress` | _off_ | Suppress progress bars. |
| `-VerboseLog` | _off_ | Log per-user classification details to console and log. |
| `-UseDeviceCode` | _off_ | Force device code sign-in. Auto-detected when no browser is available. |
| `-SkipRoleCheck` | _off_ | Skip the Microsoft Entra directory-role pre-flight check (see [Required permissions](#required-permissions)). Use when access is granted via a custom role the check doesn't recognise. |
| `-ReturnRaw` | _off_ | Emit a JSON object with summary + user lists + data-source metadata instead of writing the HTML report and tables. |

## Common usage

```powershell
# 7-day report (Graph-only, works on any OS)
pwsh ./Get-Agent365ActiveUsers.ps1 -Period D7

# Windows: include Unified Audit Log to catch unlicensed Copilot Chat activity
pwsh ./Get-Agent365ActiveUsers.ps1 -IncludeUnifiedAuditLog

# Custom report and log paths
pwsh ./Get-Agent365ActiveUsers.ps1 -ReportPath ./reports/agent365.html -LogPath ./logs/agent365.log

# No progress bars (good for CI / unattended)
pwsh ./Get-Agent365ActiveUsers.ps1 -NoProgress

# Per-user verbose logging
pwsh ./Get-Agent365ActiveUsers.ps1 -VerboseLog

# Force device code sign-in (SSH, Codespaces, headless servers)
pwsh ./Get-Agent365ActiveUsers.ps1 -UseDeviceCode

# Override Agent 365 license matching
pwsh ./Get-Agent365ActiveUsers.ps1 `
  -CopilotSkuPartNumbers 'MICROSOFT_365_COPILOT','E7' `
  -CopilotSkuPartNumberPatterns '*MICROSOFT_365_COPILOT*','*E7*','*AGENT*','*COPILOT*'

# Raw JSON output (no HTML, no console tables)
pwsh ./Get-Agent365ActiveUsers.ps1 -ReturnRaw
```

## Data sources

| Source | Endpoint / cmdlet | Used for | Platform |
| --- | --- | --- | --- |
| Microsoft Graph summary | `GET /v1.0/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='X')` | Top tiles (Active / Enabled / Copilot Chat active) and per-app counts | Any |
| Microsoft Graph user detail | `GET /beta/reports/getMicrosoft365CopilotUsageUserDetail(period='X')` (falls back to `/beta/copilot/reports/...`) | Per-user **Last Activity Date** list (rows tagged `Graph`) | Any |
| Microsoft 365 subscribed SKUs | `GET /v1.0/subscribedSkus` | Resolve Agent 365 / Copilot SKU GUIDs for the licensed-vs-unlicensed split | Any |
| User license assignments | `GET /v1.0/users/{id}` (`assignedLicenses`) | Per-user license check + canonical DisplayName / ObjectId | Any |
| Unified Audit Log *(optional)* | `Search-UnifiedAuditLog -Operations CopilotInteraction` | Catches unlicensed Copilot Chat activity not in the Graph user-detail report (rows tagged `UAL` or `Graph+UAL`) | **Windows only** |

The **Source** column in the HTML report indicates which path produced each row.

## Outputs

- `Agent365-ActiveUsers-Report.html` — tabbed HTML report (summary cards + licensed/unlicensed tables + Last Activity Date + Source + assumptions).
- `Agent365-ActiveUsers.log` — execution log with timestamps. **Recreated on every run** so the log only ever reflects the latest invocation.

Both files are ignored by `.gitignore` and never committed. The log and HTML report contain tenant-specific identifiers (user principal names, SKU part numbers, tenant display name) — review and redact before sharing externally.

## Notes and caveats

- The Agent 365 admin center's **active users over time** and **trending agents by active users** charts are not yet exposed via a public Graph API. This report uses the closest API-available signal: per-user Microsoft 365 Copilot activity, optionally augmented with `CopilotInteraction` audit events.
- Licensed/unlicensed classification depends on the configured SKU part numbers and patterns matching a SKU in your tenant. If no SKUs match, the run stops with an error.
- Microsoft Graph usage reports typically have a 24–48 hour reporting latency.
- The `getMicrosoft365CopilotUsageUserDetail` endpoint reports users with assigned Microsoft 365 Copilot / Agent 365 licenses. Unlicensed Copilot Chat activity is only visible when you also enable `-IncludeUnifiedAuditLog` on Windows.
- The per-user detail report has stricter access requirements than the summary tiles. The signed-in account needs **Reports Reader** (least privilege) or a higher Microsoft Entra admin role such as Global Administrator, AI Administrator, or one of the Exchange / SharePoint / Teams / Lync admin roles — `Reports.Read.All` Graph scope alone is not sufficient. The summary tiles use less restrictive permissions, so they may succeed even when user-detail returns `403 Forbidden`.
- The Unified Audit Log step requires `Search-UnifiedAuditLog` from the Exchange Online PowerShell module, which is **Windows-only** in PowerShell 7. The script auto-skips this step with a warning on Linux/macOS instead of failing.

## Development

A GitHub Actions workflow at `.github/workflows/lint-powershell.yml` runs [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) on every push and pull request that touches a `*.ps1` file.

## License

Released under the [MIT License](LICENSE).
