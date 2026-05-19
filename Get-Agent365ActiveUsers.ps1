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

function Ensure-Module {
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

Ensure-Module -Name 'Microsoft.Graph.Authentication'

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

function Get-ActiveUsersFromUnifiedAudit {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Days,
        [Parameter(Mandatory = $true)]
        [string[]]$Operations,
        [Parameter(Mandatory = $true)]
        [int]$ResultSize
    )

    Ensure-Module -Name 'ExchangeOnlineManagement'

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

    $matches = $matchedSkuRecords | ForEach-Object { [Guid]$_.skuId } | Select-Object -Unique
    $matchedPartNumbers = $matchedSkuRecords | ForEach-Object { [string]$_.skuPartNumber } | Sort-Object -Unique

    Write-Log -Message "Matched Agent 365 SKU IDs count: $($matches.Count)"
    Write-Log -Message "Matched SKU part numbers: $($matchedPartNumbers -join ', ')"

    return [PSCustomObject]@{
        SkuIds                = $matches
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
                [string[]]$MatchedSkuPartNumbers
        )

        $licensedRows = Convert-UsersToHtmlRows -Users $LicensedUsers
        $unlicensedRows = Convert-UsersToHtmlRows -Users $UnlicensedUsers
        $generatedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

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
            --accent: #0b5cab;
            --line: #dce3ec;
        }
        body { font-family: Segoe UI, Tahoma, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 24px; }
        .wrap { max-width: 1200px; margin: 0 auto; }
        .panel { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 20px; box-shadow: 0 8px 24px rgba(14, 28, 45, 0.06); }
        h1 { margin: 0 0 12px 0; font-size: 28px; }
        .meta { color: var(--muted); margin-bottom: 16px; }
        .notes { border: 1px solid var(--line); background: #f9fbfe; border-radius: 10px; padding: 12px 14px; margin-top: 16px; }
        .notes h2 { margin: 0 0 8px 0; font-size: 16px; }
        .notes ul { margin: 0; padding-left: 20px; color: var(--muted); }
        .notes li { margin-bottom: 6px; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 18px; }
        .stat { border: 1px solid var(--line); border-radius: 10px; padding: 12px; background: #fbfdff; }
        .stat .k { color: var(--muted); font-size: 13px; }
        .stat .v { font-weight: 700; font-size: 24px; }
        .tabs { display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }
        .tab-btn { border: 1px solid var(--line); background: #fff; border-radius: 8px; padding: 8px 12px; cursor: pointer; }
        .tab-btn.active { background: var(--accent); color: #fff; border-color: var(--accent); }
        .tab { display: none; }
        .tab.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th, td { border: 1px solid var(--line); text-align: left; padding: 8px; }
        th { background: #f0f5fb; }
    </style>
</head>
<body>
    <div class="wrap">
        <div class="panel">
            <h1>Agent 365 Active Users Report</h1>
            <div class="meta">Generated UTC: $generatedUtc | Period: $($Summary.Period) | Report Refresh Date: $($Summary.ReportRefreshDate)</div>
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
    </div>

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

New-Agent365HtmlReport -OutputPath $ReportPath -Summary $result -LicensedUsers $licensedTableUsers -UnlicensedUsers $unlicensedTableUsers -MatchedSkuPartNumbers $matchedSkuPartNumbers
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
