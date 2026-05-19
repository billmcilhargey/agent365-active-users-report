param(
    [ValidateSet('D7', 'D30', 'D90', 'D180', 'ALL')]
    [string]$Period = 'D30',

    [switch]$ReturnRaw,

    [string[]]$CopilotSkuPartNumbers = @(
        'MICROSOFT_365_COPILOT',
        'E7'
    ),

    [string[]]$CopilotSkuPartNumberPatterns = @(
        '*MICROSOFT_365_COPILOT*',
        '*E7*',
        '*AGENT*',
        '*COPILOT*'
    ),

    [string[]]$AuditOperations = @('CopilotInteraction'),

    [ValidateRange(1, 5000)]
    [int]$AuditResultSize = 5000,

    [string]$ReportPath = './Agent365-ActiveUsers-Report.html',

    [string]$LogPath = './Agent365-ActiveUsers.log',

    [switch]$NoProgress,

    [switch]$VerboseLog
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:StepId = 1
$script:TotalSteps = 9
$script:ScriptVersion = '1.0.0'
$script:LogFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $LogPath))

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $script:LogFullPath -Value $line
    Write-Host $line
}

function Write-VerboseLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($VerboseLog) {
        Write-Log -Message $Message
    }
}

function Update-StepProgress {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,
        [Parameter(Mandatory = $true)]
        [string]$Status
    )

    if ($NoProgress) {
        return
    }

    $percent = [math]::Round((($script:StepId - 1) / $script:TotalSteps) * 100, 0)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percent
}

"" | Set-Content -Path $script:LogFullPath -Encoding UTF8
Write-Log -Message "Starting Agent 365 active users report run."
if ($VerboseLog) {
    Write-Log -Message 'Verbose logging is enabled.'
}

function Initialize-RequiredModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log -Message "Installing module: $Name"
        Write-Host "Installing module: $Name"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log -Message "Module loaded: $Name"
}

Initialize-RequiredModule -Name 'Microsoft.Graph.Authentication'

function Get-PeriodDays {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PeriodValue
    )

    switch ($PeriodValue) {
        'D7' { return 7 }
        'D30' { return 30 }
        'D90' { return 90 }
        'D180' { return 180 }
        'ALL' { return 180 }
        default { return 30 }
    }
}

function Get-TenantContextInfo {
    [CmdletBinding()]
    param()

    $info = [PSCustomObject]@{
        TenantName    = ''
        Domain        = ''
        TenantId      = ''
        Account       = ''
        ScriptVersion = $script:ScriptVersion
    }

    try {
        $ctx = Get-MgContext
        if ($ctx) {
            if ($ctx.Account)  { $info.Account  = [string]$ctx.Account }
            if ($ctx.TenantId) { $info.TenantId = [string]$ctx.TenantId }
        }
    } catch {
        Write-VerboseLog -Message "Get-MgContext failed: $($_.Exception.Message)"
    }

    try {
        $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,verifiedDomains"
        if ($org.value -and $org.value.Count -gt 0) {
            $first = $org.value[0]
            if ($first.displayName) { $info.TenantName = [string]$first.displayName }
            if (-not $info.TenantId -and $first.id) { $info.TenantId = [string]$first.id }
            if ($first.verifiedDomains) {
                $primary = $first.verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1
                if (-not $primary) { $primary = $first.verifiedDomains | Select-Object -First 1 }
                if ($primary -and $primary.name) { $info.Domain = [string]$primary.name }
            }
        }
    } catch {
        Write-VerboseLog -Message "Organization lookup failed: $($_.Exception.Message)"
    }

    return $info
}

function Get-ActiveUsersFromUnifiedAudit {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Days,
        [Parameter(Mandatory = $true)]
        [string[]]$Operations,
        [Parameter(Mandatory = $true)]
        [int]$ResultSize
    )

    Initialize-RequiredModule -Name 'ExchangeOnlineManagement'

    try {
        $connectionInfo = Get-ConnectionInformation -ErrorAction Stop
    } catch {
        $connectionInfo = $null
    }

    if (-not $connectionInfo) {
        Write-Log -Message 'Connecting to Exchange Online for Unified Audit Log access.'
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    }

    $startDate = (Get-Date).ToUniversalTime().AddDays(-$Days)
    $endDate = (Get-Date).ToUniversalTime()

    $records = @()
    try {
        Write-Log -Message "Querying Unified Audit Log for the last $Days days. Operations: $($Operations -join ', '). ResultSize: $ResultSize"
        $records = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -Operations $Operations -ResultSize $ResultSize
    } catch {
        Write-Log -Level 'ERROR' -Message "Unified Audit Log query failed: $($_.Exception.Message)"
        throw "Failed to query Unified Audit Log with operations [$($Operations -join ', ')]. Error: $($_.Exception.Message)"
    }

    if (-not $records) {
        return @()
    }

    $activeUsers = $records |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserIds) } |
        Select-Object -ExpandProperty UserIds -Unique

    Write-VerboseLog -Message "Unified Audit Log records retrieved: $($records.Count). Unique active users: $($activeUsers.Count)."

    return $activeUsers
}

function Get-CopilotSkuIds {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$SkuPartNumbers,
        [Parameter(Mandatory = $true)]
        [string[]]$SkuPartNumberPatterns
    )

    $uri = "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuId,skuPartNumber"
    $skus = Invoke-MgGraphRequest -Method GET -Uri $uri

    if (-not $skus.value) {
        return @()
    }

    $wantedExact = $SkuPartNumbers | ForEach-Object { $_.ToUpperInvariant() }
    $wantedPatterns = $SkuPartNumberPatterns

    $matchedSkuRecords = $skus.value | Where-Object {
        $skuPart = [string]$_.skuPartNumber
        if ($wantedExact -contains $skuPart.ToUpperInvariant()) {
            return $true
        }

        foreach ($pattern in $wantedPatterns) {
            if ($skuPart -like $pattern) {
                return $true
            }
        }

        return $false
    }

    $matchedSkuIds = $matchedSkuRecords | ForEach-Object { [Guid]$_.skuId } | Select-Object -Unique
    $matchedPartNumbers = $matchedSkuRecords | ForEach-Object { [string]$_.skuPartNumber } | Sort-Object -Unique

    Write-Log -Message "Matched Agent 365 SKU IDs count: $($matchedSkuIds.Count)"
    Write-Log -Message "Matched SKU part numbers: $($matchedPartNumbers -join ', ')"

    return [PSCustomObject]@{
        SkuIds                = $matchedSkuIds
        MatchedSkuPartNumbers = $matchedPartNumbers
    }
}

function Get-ActiveUserLicenseClassification {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ActiveUsers,
        [Parameter(Mandatory = $true)]
        [Guid[]]$CopilotSkuIds
    )

    if ($CopilotSkuIds.Count -eq 0) {
        throw 'No matching Agent 365 SKUs found in tenant. Check -CopilotSkuPartNumbers and -CopilotSkuPartNumberPatterns values.'
    }

    $licensed = New-Object System.Collections.Generic.List[object]
    $unlicensed = New-Object System.Collections.Generic.List[object]
    $total = $ActiveUsers.Count
    $index = 0

    foreach ($userId in $ActiveUsers) {
        $index++
        if (-not $NoProgress) {
            $pct = if ($total -eq 0) { 100 } else { [math]::Round(($index / $total) * 100, 0) }
            Write-Progress -Id 2 -Activity 'Classifying active users by Agent 365 license' -Status "Processing $index of $total" -PercentComplete $pct
        }

        try {
            $encodedUserId = [uri]::EscapeDataString($userId)
            $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserId?`$select=id,userPrincipalName,displayName,assignedLicenses"
        } catch {
            Write-Log -Level 'WARN' -Message "Skipping user '$userId' due to lookup failure."
            continue
        }

        $assignedSkuIds = @()
        if ($user.assignedLicenses) {
            $assignedSkuIds = $user.assignedLicenses | ForEach-Object { [Guid]$_.skuId }
        }

        $hasCopilot = $false
        foreach ($assigned in $assignedSkuIds) {
            if ($CopilotSkuIds -contains $assigned) {
                $hasCopilot = $true
                break
            }
        }

                $record = [PSCustomObject]@{
                        UserPrincipalName = $user.userPrincipalName
                        DisplayName       = $user.displayName
                        ObjectId          = $user.id
                        IsAgent365Licensed = $hasCopilot
                }

                if ($hasCopilot) {
                        $licensed.Add($record)
                        Write-VerboseLog -Message "Licensed user: $($record.UserPrincipalName) ($($record.ObjectId))"
                } else {
                        $unlicensed.Add($record)
                        Write-VerboseLog -Message "Unlicensed user: $($record.UserPrincipalName) ($($record.ObjectId))"
                }
        }

        if (-not $NoProgress) {
            Write-Progress -Id 2 -Activity 'Classifying active users by Agent 365 license' -Completed
        }

        Write-Log -Message "Classified users. Licensed: $($licensed.Count), Unlicensed: $($unlicensed.Count)"

        return [PSCustomObject]@{
                Licensed   = $licensed
                Unlicensed = $unlicensed
        }
}

function Convert-UsersToHtmlRows {
        param(
                [Parameter(Mandatory = $true)]
                [object[]]$Users
        )

        if (-not $Users -or $Users.Count -eq 0) {
                return '<tr><td colspan="3">No users found.</td></tr>'
        }

        $rows = foreach ($u in ($Users | Sort-Object UserPrincipalName)) {
                $upn = [System.Net.WebUtility]::HtmlEncode([string]$u.UserPrincipalName)
                $name = [System.Net.WebUtility]::HtmlEncode([string]$u.DisplayName)
                $oid = [System.Net.WebUtility]::HtmlEncode([string]$u.ObjectId)
                "<tr><td>$upn</td><td>$name</td><td>$oid</td></tr>"
        }

        return ($rows -join [Environment]::NewLine)
}

function New-Agent365HtmlReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Summary,
        [Parameter(Mandatory = $true)]
        [object[]]$LicensedUsers,
        [Parameter(Mandatory = $true)]
        [object[]]$UnlicensedUsers,
        [Parameter(Mandatory = $true)]
        [string[]]$MatchedSkuPartNumbers,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$TenantInfo
    )

    $licensedRows = Convert-UsersToHtmlRows -Users $LicensedUsers
    $unlicensedRows = Convert-UsersToHtmlRows -Users $UnlicensedUsers

    $nowUtc = (Get-Date).ToUniversalTime()
    $generatedUtc = $nowUtc.ToString('yyyy-MM-dd HH:mm:ss')
    $generatedUtcDisplay = $nowUtc.ToString('MMMM d, yyyy HH:mm:ss') + ' UTC'
    $year = (Get-Date).Year

    $repoUrl = 'https://github.com/billmcilhargey/agent365-active-users-report'
    $repoIssuesUrl = "$repoUrl/issues"
    $repoReadmeUrl = "$repoUrl/blob/main/README.md"

    $encode = {
        param($Value, $Default = 'Unknown')
        if ($null -eq $Value) { return $Default }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
        return [System.Net.WebUtility]::HtmlEncode($text)
    }

    $tenantNameHtml    = & $encode $TenantInfo.TenantName    'Unknown tenant'
    $tenantDomainHtml  = & $encode $TenantInfo.Domain        'Unknown'
    $tenantIdHtml      = & $encode $TenantInfo.TenantId      'Unknown'
    $accountHtml       = & $encode $TenantInfo.Account       'Unknown'
    $scriptVersionHtml = & $encode $TenantInfo.ScriptVersion '0.0.0'

    $html = @"
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Agent 365 Active Users Report</title>
    <style>
        :root {
            --bg: #f4f7fb;
            --card: #ffffff;
            --text: #1b2430;
            --muted: #5a6777;
            --muted-2: #6b7280;
            --accent: #0b5cab;
            --line: #e2e8f0;
        }
        * { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: var(--bg); color: var(--text); line-height: 1.5; }
        a { color: var(--accent); text-decoration: none; }
        a:hover { text-decoration: underline; }
        .container { max-width: 1200px; margin: 0 auto; padding: 0 24px; }

        /* Header */
        .site-header { position: sticky; top: 0; z-index: 50; width: 100%; border-bottom: 1px solid var(--line); background: rgba(255, 255, 255, 0.9); backdrop-filter: saturate(180%) blur(8px); -webkit-backdrop-filter: saturate(180%) blur(8px); }
        .site-header .inner { display: flex; align-items: center; min-height: 56px; gap: 16px; }
        .brand { display: inline-flex; align-items: center; gap: 10px; font-weight: 600; color: var(--text); }
        .brand:hover { text-decoration: none; }
        .brand-mark { width: 28px; height: 28px; border-radius: 6px; background: linear-gradient(135deg, #0b5cab, #36b9ff); color: #fff; display: inline-flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 700; }
        .header-right { margin-left: auto; display: flex; align-items: center; gap: 4px; }
        .header-btn { display: inline-flex; align-items: center; gap: 6px; padding: 6px 10px; border-radius: 6px; border: 1px solid transparent; background: transparent; color: var(--text); cursor: pointer; font: inherit; font-size: 13px; }
        .header-btn:hover { background: #eef2f7; text-decoration: none; }
        .gh-icon { width: 16px; height: 16px; fill: currentColor; }
        .tenant-menu { position: relative; }
        .tenant-panel { display: none; position: absolute; right: 0; top: calc(100% + 6px); min-width: 320px; background: #fff; border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 12px 28px rgba(14, 28, 45, 0.12); padding: 6px 0; }
        .tenant-menu.open .tenant-panel { display: block; }
        .tenant-row { padding: 8px 14px; }
        .tenant-row + .tenant-row { border-top: 1px solid var(--line); }
        .tenant-row .label { font-size: 13px; font-weight: 600; color: var(--text); }
        .tenant-row .value { font-size: 12px; color: var(--muted-2); margin-top: 2px; word-break: break-all; }
        .caret { font-size: 10px; opacity: 0.7; }

        /* Main */
        main { padding: 24px 0; }
        .panel { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 20px; box-shadow: 0 8px 24px rgba(14, 28, 45, 0.06); }
        h1 { margin: 0 0 12px 0; font-size: 28px; }
        .meta { color: var(--muted); margin-bottom: 16px; font-size: 13px; }
        .notes { border: 1px solid var(--line); background: #f9fbfe; border-radius: 10px; padding: 12px 14px; margin-top: 16px; }
        .notes h2 { margin: 0 0 8px 0; font-size: 16px; }
        .notes ul { margin: 0; padding-left: 20px; color: var(--muted); }
        .notes li { margin-bottom: 6px; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 18px; }
        .stat { border: 1px solid var(--line); border-radius: 10px; padding: 12px; background: #fbfdff; }
        .stat .k { color: var(--muted); font-size: 13px; }
        .stat .v { font-weight: 700; font-size: 24px; }
        .tabs { display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }
        .tab-btn { border: 1px solid var(--line); background: #fff; border-radius: 8px; padding: 8px 12px; cursor: pointer; font: inherit; font-size: 13px; }
        .tab-btn.active { background: var(--accent); color: #fff; border-color: var(--accent); }
        .tab { display: none; }
        .tab.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th, td { border: 1px solid var(--line); text-align: left; padding: 8px; }
        th { background: #f0f5fb; }

        /* Footer */
        .site-footer { border-top: 1px solid var(--line); background: #fff; margin-top: 32px; }
        .footer-grid { display: grid; grid-template-columns: 1fr; gap: 32px; padding: 32px 0; }
        @media (min-width: 768px) { .footer-grid { grid-template-columns: 1fr 1fr 1fr; } }
        .footer-col h4 { margin: 0 0 12px 0; font-size: 14px; font-weight: 600; color: var(--text); }
        .footer-col p { margin: 0; font-size: 13px; color: var(--muted-2); }
        .footer-col a { display: block; font-size: 13px; color: var(--muted-2); padding: 4px 0; }
        .footer-col a:hover { color: var(--text); }
        .footer-bottom { border-top: 1px solid var(--line); padding: 16px 0; display: flex; flex-direction: column; gap: 8px; align-items: center; text-align: center; font-size: 12px; color: var(--muted-2); }
        @media (min-width: 768px) { .footer-bottom { flex-direction: row; justify-content: space-between; text-align: left; } }
        .footer-bottom .left { display: flex; flex-direction: column; gap: 2px; }
        .footer-bottom .right { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
        .footer-bottom a { color: var(--muted-2); }
        .footer-bottom a:hover { color: var(--text); }
        .dot { color: var(--line); }
    </style>
</head>
<body>
    <header class="site-header">
        <div class="container inner">
            <a class="brand" href="#top">
                <span class="brand-mark">A</span>
                <span>Agent 365 Active Users Report</span>
            </a>
            <div class="header-right">
                <a class="header-btn" href="$repoUrl" target="_blank" rel="noreferrer noopener" title="View on GitHub" aria-label="View on GitHub">
                    <svg class="gh-icon" viewBox="0 0 16 16" aria-hidden="true"><path fill-rule="evenodd" d="M8 0C3.58 0 0 3.58 0 8a8 8 0 0 0 5.47 7.59c.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8z"/></svg>
                    <span>GitHub</span>
                </a>
                <div class="tenant-menu" id="tenant-menu">
                    <button type="button" class="header-btn" id="tenant-btn" aria-haspopup="true" aria-expanded="false">
                        <span>$tenantNameHtml</span>
                        <span class="caret" aria-hidden="true">&#9662;</span>
                    </button>
                    <div class="tenant-panel" role="menu">
                        <div class="tenant-row"><div class="label">Tenant</div><div class="value">$tenantNameHtml</div></div>
                        <div class="tenant-row"><div class="label">Domain</div><div class="value">$tenantDomainHtml</div></div>
                        <div class="tenant-row"><div class="label">Tenant ID</div><div class="value">$tenantIdHtml</div></div>
                        <div class="tenant-row"><div class="label">Generated by</div><div class="value">$accountHtml</div></div>
                        <div class="tenant-row"><div class="label">Run on</div><div class="value">$generatedUtcDisplay</div></div>
                        <div class="tenant-row"><div class="label">Version</div><div class="value">$scriptVersionHtml</div></div>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <main class="container" id="top">
        <div class="panel">
            <h1>Agent 365 Active Users Report</h1>
            <div class="meta">Generated UTC: $generatedUtc &middot; Period: $($Summary.Period) &middot; Report Refresh Date: $($Summary.ReportRefreshDate)</div>
            <div class="stats">
                <div class="stat"><div class="k">Agent 365 Active Users</div><div class="v">$($Summary.ActiveUsers)</div></div>
                <div class="stat"><div class="k">Enabled Users</div><div class="v">$($Summary.EnabledUsers)</div></div>
                <div class="stat"><div class="k">Copilot Chat Active Users</div><div class="v">$($Summary.CopilotChatActiveUsers)</div></div>
                <div class="stat"><div class="k">Active Licensed Users Total</div><div class="v">$($LicensedUsers.Count)</div></div>
                <div class="stat"><div class="k">Active Unlicensed Users Total</div><div class="v">$($UnlicensedUsers.Count)</div></div>
            </div>

            <div class="tabs">
                <button class="tab-btn active" data-tab="licensed">Licensed Active Users</button>
                <button class="tab-btn" data-tab="unlicensed">Unlicensed Active Users</button>
            </div>

            <div id="licensed" class="tab active">
                <table>
                    <thead>
                        <tr><th>UserPrincipalName</th><th>DisplayName</th><th>ObjectId</th></tr>
                    </thead>
                    <tbody>
$licensedRows
                    </tbody>
                </table>
            </div>

            <div id="unlicensed" class="tab">
                <table>
                    <thead>
                        <tr><th>UserPrincipalName</th><th>DisplayName</th><th>ObjectId</th></tr>
                    </thead>
                    <tbody>
$unlicensedRows
                    </tbody>
                </table>
            </div>

            <div class="notes">
                <h2>Notes, Assumptions, and Caveats</h2>
                <ul>
                    <li>How Agent 365 license is determined: the script matches tenant subscribed SKUs using exact SKU part numbers ($($CopilotSkuPartNumbers -join ', ')) plus wildcard patterns ($($CopilotSkuPartNumberPatterns -join ', ')), then maps these to skuId GUIDs. Each active user is marked licensed only if any users/{id}.assignedLicenses.skuId matches one of those GUIDs.</li>
                    <li>SKUs matched in this tenant for this run: $($MatchedSkuPartNumbers -join ', ').</li>
                    <li>Active user identity lists come from Unified Audit Log operations: $($AuditOperations -join ', '). If your tenant logs different operation names, totals may differ.</li>
                    <li>Summary metrics (top cards) come from Microsoft Graph Copilot reports; user detail lists come from audit records and Graph user lookup. These sources can have timing/latency differences.</li>
                    <li>Report accuracy depends on retention and availability of Unified Audit Log data and permissions for Graph and Exchange Online access.</li>
                    <li>This script queries up to $AuditResultSize audit records per run (configurable via -AuditResultSize, max 5000).</li>
                </ul>
            </div>
        </div>
    </main>

    <footer class="site-footer">
        <div class="container">
            <div class="footer-grid">
                <div class="footer-col">
                    <h4>Agent 365 Active Users Report</h4>
                    <p>PowerShell utility that builds an HTML report of Agent 365 (Microsoft 365 Copilot) active users, split into licensed and unlicensed.</p>
                </div>
                <div class="footer-col">
                    <h4>Resources</h4>
                    <a href="https://learn.microsoft.com/en-us/graph/api/resources/microsoft365copilotusage-api-overview" target="_blank" rel="noreferrer noopener">Microsoft Graph Copilot reporting API</a>
                    <a href="https://learn.microsoft.com/en-us/microsoft-365/admin/activity-reports/activity-reports" target="_blank" rel="noreferrer noopener">Microsoft 365 usage reports</a>
                    <a href="https://learn.microsoft.com/en-us/purview/audit-log-search" target="_blank" rel="noreferrer noopener">Unified Audit Log</a>
                </div>
                <div class="footer-col">
                    <h4>Support</h4>
                    <a href="$repoIssuesUrl" target="_blank" rel="noreferrer noopener">Report an issue</a>
                    <a href="$repoUrl" target="_blank" rel="noreferrer noopener">View on GitHub</a>
                    <a href="$repoReadmeUrl" target="_blank" rel="noreferrer noopener">Documentation (README)</a>
                </div>
            </div>
            <div class="footer-bottom">
                <div class="left">
                    <div>&copy; $year Agent 365 Active Users Report</div>
                    <div>This is a community project and not an official Microsoft product.</div>
                </div>
                <div class="right">
                    <a href="$repoReadmeUrl" target="_blank" rel="noreferrer noopener">README</a>
                    <span class="dot">&bull;</span>
                    <a href="$repoUrl" target="_blank" rel="noreferrer noopener">GitHub</a>
                    <span class="dot">&bull;</span>
                    <span>$generatedUtcDisplay</span>
                </div>
            </div>
        </div>
    </footer>

    <script>
        const buttons = document.querySelectorAll('.tab-btn');
        const tabs = document.querySelectorAll('.tab');
        buttons.forEach(btn => {
            btn.addEventListener('click', () => {
                buttons.forEach(b => b.classList.remove('active'));
                tabs.forEach(t => t.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(btn.dataset.tab).classList.add('active');
            });
        });

        const tenantMenu = document.getElementById('tenant-menu');
        const tenantBtn = document.getElementById('tenant-btn');
        if (tenantBtn && tenantMenu) {
            tenantBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const open = tenantMenu.classList.toggle('open');
                tenantBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
            });
            document.addEventListener('click', (e) => {
                if (!tenantMenu.contains(e.target)) {
                    tenantMenu.classList.remove('open');
                    tenantBtn.setAttribute('aria-expanded', 'false');
                }
            });
        }
    </script>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

function Get-UserTableProjection {
        param(
                [Parameter(Mandatory = $true)]
                [object[]]$Users
        )

        return $Users | Select-Object UserPrincipalName, DisplayName, ObjectId
}

$scopes = @('Reports.Read.All', 'User.Read.All', 'Organization.Read.All')

Update-StepProgress -Activity 'Agent 365 report' -Status 'Checking Graph context'
Write-Log -Message 'Checking Microsoft Graph authentication context.'

try {
    $ctx = Get-MgContext
} catch {
    $ctx = $null
}

if (-not $ctx) {
    Write-Log -Message 'Connecting to Microsoft Graph.'
    Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null
} elseif (-not ($ctx.Scopes -contains 'Reports.Read.All')) {
    Write-Log -Message 'Reconnecting to Microsoft Graph to include required scopes.'
    Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null
}
$script:StepId++

Write-Log -Message 'Resolving tenant context for report header.'
$tenantInfo = Get-TenantContextInfo

# Microsoft recommends the /copilot/reports endpoint for Copilot usage reporting APIs.
Update-StepProgress -Activity 'Agent 365 report' -Status 'Pulling summary metrics from Graph'
Write-Log -Message 'Requesting Copilot summary usage metrics from Graph.'
$uri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')?`$format=application/json"
$response = Invoke-MgGraphRequest -Method GET -Uri $uri

if (-not $response.value -or $response.value.Count -eq 0) {
    throw 'No data returned from getMicrosoft365CopilotUserCountSummary.'
}

$latest = $response.value | Sort-Object reportRefreshDate -Descending | Select-Object -First 1

if (-not $latest.adoptionByProduct -or $latest.adoptionByProduct.Count -eq 0) {
    throw 'No adoptionByProduct data found in response.'
}

$periodRow = $latest.adoptionByProduct | Where-Object { $_.reportPeriod -eq [int]($Period.TrimStart('D')) } | Select-Object -First 1
if (-not $periodRow) {
    $periodRow = $latest.adoptionByProduct | Select-Object -First 1
}

$result = [PSCustomObject]@{
    ReportRefreshDate      = $latest.reportRefreshDate
    Period                 = $Period
    ActiveUsers            = $periodRow.anyAppActiveUsers
    EnabledUsers           = $periodRow.anyAppEnabledUsers
    CopilotChatActiveUsers = $periodRow.copilotChatActiveUsers
}
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Collecting active users from audit log'
Write-Log -Message 'Collecting active users from Unified Audit Log.'
$days = Get-PeriodDays -PeriodValue $Period
$activeUsersFromAudit = Get-ActiveUsersFromUnifiedAudit -Days $days -Operations $AuditOperations -ResultSize $AuditResultSize
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Resolving Agent 365 SKUs'
Write-Log -Message 'Resolving Agent 365 SKU IDs from subscribed SKUs.'
$skuMatchResults = Get-CopilotSkuIds -SkuPartNumbers $CopilotSkuPartNumbers -SkuPartNumberPatterns $CopilotSkuPartNumberPatterns
$copilotSkuIds = @($skuMatchResults.SkuIds)
$matchedSkuPartNumbers = @($skuMatchResults.MatchedSkuPartNumbers)
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Classifying licensed vs unlicensed users'
Write-Log -Message "Classifying $($activeUsersFromAudit.Count) active users by license assignment."
$classification = Get-ActiveUserLicenseClassification -ActiveUsers $activeUsersFromAudit -CopilotSkuIds $copilotSkuIds
$script:StepId++

$licensedActiveUsers = @($classification.Licensed)
$unlicensedActiveUsers = @($classification.Unlicensed)

$licensedTableUsers = Get-UserTableProjection -Users $licensedActiveUsers
$unlicensedTableUsers = Get-UserTableProjection -Users $unlicensedActiveUsers

if ($ReturnRaw) {
    $reportFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))

    [PSCustomObject]@{
        Summary                    = $result
        ActiveLicensedUsersTotal   = $licensedTableUsers.Count
        ActiveUnlicensedUsersTotal = $unlicensedTableUsers.Count
        InputAgent365SkuPartNumbers = $CopilotSkuPartNumbers
        InputAgent365SkuPatterns    = $CopilotSkuPartNumberPatterns
        MatchedAgent365SkuPartNumbers = $matchedSkuPartNumbers
        LicensedActiveUsers        = $licensedTableUsers
        UnlicensedActiveUsers      = $unlicensedTableUsers
        ReportPath                 = $reportFullPath
    } | ConvertTo-Json -Depth 8
    Write-Log -Message 'Completed run in ReturnRaw mode.'
    if (-not $NoProgress) {
        Write-Progress -Activity 'Agent 365 report' -Completed
    }
    return
}

Write-Host "Agent 365 Active Users (Period: $Period): $($result.ActiveUsers)"
Write-Host "Report Refresh Date: $($result.ReportRefreshDate)"
Write-Host "Enabled Users: $($result.EnabledUsers)"
Write-Host "Copilot Chat Active Users: $($result.CopilotChatActiveUsers)"
Write-Host "Agent 365 SKU exact matches configured: $($CopilotSkuPartNumbers -join ', ')"
Write-Host "Agent 365 SKU pattern matches configured: $($CopilotSkuPartNumberPatterns -join ', ')"
Write-Host "Agent 365 SKUs matched in tenant: $($matchedSkuPartNumbers -join ', ')"
Write-Host "Active Licensed Users Total: $($licensedTableUsers.Count)"
Write-Host "Active Unlicensed Users Total: $($unlicensedTableUsers.Count)"

New-Agent365HtmlReport -OutputPath $ReportPath -Summary $result -LicensedUsers $licensedTableUsers -UnlicensedUsers $unlicensedTableUsers -MatchedSkuPartNumbers $matchedSkuPartNumbers -TenantInfo $tenantInfo
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Finalizing report output'
Write-Log -Message "HTML report generated at: $([System.IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath)))"

Write-Host "Report generated: $(Resolve-Path $ReportPath)"
Write-Host "Execution log: $script:LogFullPath"
if ($VerboseLog) {
    Write-Host 'Verbose logging: enabled'
}

Write-Host ''
Write-Host 'Licensed Active Users (Agent 365):'
if ($licensedTableUsers.Count -gt 0) {
    $licensedTableUsers | Sort-Object UserPrincipalName | Format-Table -AutoSize UserPrincipalName, DisplayName, ObjectId
} else {
    Write-Host 'No licensed active users found.'
}

Write-Host ''
Write-Host 'Unlicensed Active Users (Agent 365):'
if ($unlicensedTableUsers.Count -gt 0) {
    $unlicensedTableUsers | Sort-Object UserPrincipalName | Format-Table -AutoSize UserPrincipalName, DisplayName, ObjectId
} else {
    Write-Host 'No unlicensed active users found.'
}

$script:StepId++
Update-StepProgress -Activity 'Agent 365 report' -Status 'Complete'
if (-not $NoProgress) {
    Write-Progress -Activity 'Agent 365 report' -Completed
}
Write-Log -Message 'Run completed successfully.'
