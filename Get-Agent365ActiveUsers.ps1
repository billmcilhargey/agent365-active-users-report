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

    [switch]$VerboseLog,

    [switch]$UseDeviceCode,

    # Optional login hint used across all auth paths. When supplied, this
    # UPN is shown as the account to pick in browser / device-code flows,
    # and passed through to Unified Audit Log child-session connections.
    # Omit this parameter (or pass an empty string) to use the normal
    # account picker on machines with interactive sign-in support.
    [AllowEmptyString()]
    [string]$SignInAccount,

    # Explicitly request account-picker behavior without providing a UPN.
    # Useful in automation/scripts where using "-SignInAccount" with no
    # value would be a PowerShell parse error.
    [switch]$PickSignInAccount,

    # Unified Audit Log step (Exchange Online PowerShell, Windows-only).
    # Disabled by default; pass -IncludeUnifiedAuditLog to opt in.
    [bool]$IncludeUnifiedAuditLog = $false,

    # Convenience force-disable switch; equivalent to
    # -IncludeUnifiedAuditLog:$false.
    [switch]$SkipUnifiedAuditLog,

    # Skip the Microsoft Entra directory-role pre-check (use for custom roles
    # the check does not recognise; see README "Required permissions").
    [switch]$SkipRoleCheck
)

# Honour the convenience opt-out switch.
if ($SkipUnifiedAuditLog) { $IncludeUnifiedAuditLog = $false }

if ($SignInAccount) {
    $SignInAccount = $SignInAccount.Trim()
}
if ([string]::IsNullOrWhiteSpace($SignInAccount)) {
    $SignInAccount = $null
}
if ($PickSignInAccount) {
    $SignInAccount = $null
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 does not define the automatic variables
# $IsWindows / $IsLinux / $IsMacOS that the rest of the script (and
# Set-StrictMode) expect. Define safe shims so platform checks below work
# regardless of host. PowerShell 7+ already has these as read-only
# automatic variables and will skip this block.
if (-not (Get-Variable -Name 'IsWindows' -Scope Global -ErrorAction SilentlyContinue)) {
    Set-Variable -Name 'IsWindows' -Value $true  -Scope Global -Option ReadOnly -Force
    Set-Variable -Name 'IsLinux'   -Value $false -Scope Global -Option ReadOnly -Force
    Set-Variable -Name 'IsMacOS'   -Value $false -Scope Global -Option ReadOnly -Force
}

$script:StepId = 1
$script:TotalSteps = 9
$script:ScriptVersion = '1.0.0'
$script:LogFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $LogPath))
$script:ReportFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))

# Set by the directory-role pre-check (see Test-Agent365ReportsAccess) when
# the signed-in account is known to lack a role that qualifies for the
# per-user Copilot usage report. The main flow uses this to skip the
# user-detail call and degrade gracefully to a summary-only report.
$script:userDetailLikelyForbidden = $false
$script:userDetailSkipReason      = $null

# Collector of report steps that were intentionally skipped, so the console
# summary and HTML report can show [SKIPPED] in place of misleading 0
# values and explain WHY in a dedicated notes section. Populated via
# Add-SkippedItem from the pre-check, main flow, and UAL paths.
$script:SkippedItems = New-Object System.Collections.Generic.List[object]

function Add-SkippedItem {
    <#
    .SYNOPSIS
        Record a report step that was intentionally skipped, with a
        human-readable reason for later display in the console summary and
        the HTML notes section.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Item,

        [Parameter(Mandatory = $true)]
        [string]$Reason,

        [string]$Remediation
    )

    # De-duplicate by Item so the same step is not recorded twice when
    # multiple code paths flag the same skip (e.g. pre-check then 403 catch).
    foreach ($existing in $script:SkippedItems) {
        if ($existing.Item -eq $Item) { return }
    }

    $script:SkippedItems.Add([pscustomobject]@{
        Item        = $Item
        Reason      = $Reason
        Remediation = $Remediation
    })
}

function Format-CountOrSkipped {
    <#
    .SYNOPSIS
        Return the numeric Count as a string, OR the literal token
        "[SKIPPED]" when the upstream step that would have produced data
        was skipped. Used by console summary and HTML stat tiles to avoid
        showing a misleading 0 when the real answer is "we could not check".
    #>
    param(
        [int]$Count,
        [bool]$WasSkipped
    )

    if ($WasSkipped) { return '[SKIPPED]' }
    return [string]$Count
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $script:LogFullPath -Value $line

    # Console output: keep the timestamp neutral and color the [LEVEL] tag so
    # INFO / WARN / ERROR stand out at a glance. Falls back to a plain
    # Write-Host if the host doesn't support color (e.g. redirected stdout).
    $levelColor = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        default   { $null }
    }

    try {
        Write-Host "[$ts] " -NoNewline
        if ($levelColor) {
            Write-Host "[$Level]" -NoNewline -ForegroundColor $levelColor
        } else {
            Write-Host "[$Level]" -NoNewline
        }
        Write-Host " $Message"
    } catch {
        Write-Host $line
    }
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

# Always start with a fresh log file so each run is self-contained.
if (Test-Path -LiteralPath $script:LogFullPath) {
    try {
        Remove-Item -LiteralPath $script:LogFullPath -Force -ErrorAction Stop
    } catch {
        Write-Warning "Could not remove existing log file '$script:LogFullPath': $($_.Exception.Message). Falling back to truncate."
        '' | Set-Content -LiteralPath $script:LogFullPath -Encoding UTF8 -Force
    }
}

# Always start with a fresh HTML report so each run reflects only the latest data.
if (Test-Path -LiteralPath $script:ReportFullPath) {
    try {
        Remove-Item -LiteralPath $script:ReportFullPath -Force -ErrorAction Stop
    } catch {
        Write-Warning "Could not remove existing report file '$script:ReportFullPath': $($_.Exception.Message). It will be overwritten when the new report is written."
    }
}

# Clear the console on interactive runs so the report output starts on a
# fresh screen. Wrapped because Clear-Host can throw when the script is
# executed in a non-interactive host (CI runners, redirected stdout, etc.).
try { Clear-Host } catch { Write-Verbose "Clear-Host skipped: $($_.Exception.Message)" }

Write-Log -Message ("=" * 72)
Write-Log -Message "Agent 365 Active Users Report v$script:ScriptVersion"
Write-Log -Message "Run started: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))"
Write-Log -Message "Log file:    $script:LogFullPath"
Write-Log -Message ("=" * 72)
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
        # Suppress harmless PowerShellGet bootstrap warnings for cleaner output.
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -WarningAction SilentlyContinue
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log -Message "Module loaded: $Name"
}

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

function Get-PeriodDescription {
    <#
    .SYNOPSIS
        Render a human-friendly description for a Microsoft Graph reporting
        period code (e.g. D30 -> "trailing 30 days (2026-04-18 -> 2026-05-17)").

    .DESCRIPTION
        Microsoft Graph usage reports accept period codes D7, D30, D90, D180,
        and ALL. The codes alone are opaque to anyone reading the report, so
        this helper turns them into a short phrase that also includes the
        actual date window when a report refresh date is supplied.

        Window math follows the same convention as the Graph reports: the
        period is INCLUSIVE of the refresh date. For example, D30 with a
        refresh date of 2026-05-17 covers 2026-04-18 through 2026-05-17 (30
        calendar days inclusive on both ends).

    .PARAMETER PeriodValue
        Period code: D7, D30, D90, D180, or ALL.

    .PARAMETER ReportRefreshDate
        Optional. The Report Refresh Date string from the Graph summary
        response (yyyy-MM-dd). When supplied, the description includes the
        concrete date window. When omitted, only the trailing-day phrase is
        returned.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$PeriodValue,

        [string]$ReportRefreshDate
    )

    $base = switch ($PeriodValue) {
        'D7'   { 'trailing 7 days' }
        'D30'  { 'trailing 30 days' }
        'D90'  { 'trailing 90 days' }
        'D180' { 'trailing 180 days' }
        'ALL'  { 'all available history (up to last 180 days)' }
        default { "trailing $PeriodValue" }
    }

    if ([string]::IsNullOrWhiteSpace($ReportRefreshDate)) {
        return $base
    }

    $end = $null
    try {
        $end = [datetime]::ParseExact($ReportRefreshDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        try { $end = [datetime]$ReportRefreshDate } catch { $end = $null }
    }
    if (-not $end) { return $base }

    $days = Get-PeriodDays -PeriodValue $PeriodValue
    if ($PeriodValue -eq 'ALL') {
        return "$base (through $($end.ToString('yyyy-MM-dd')))"
    }

    # Window is inclusive on both ends: refresh date counts as day N, so the
    # start is (days - 1) days earlier (e.g. D30 ending 2026-05-17 starts
    # 2026-04-18 -- 30 calendar days inclusive).
    $start = $end.AddDays(-($days - 1))
    return "$base ($($start.ToString('yyyy-MM-dd')) -> $($end.ToString('yyyy-MM-dd')))"
}

function Invoke-Agent365Preflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Period,
        [Parameter(Mandatory = $true)]
        [string[]]$RequiredModules,
        [Parameter(Mandatory = $true)]
        [string[]]$AuditOperations,
        [switch]$IncludeUnifiedAuditLog
    )

    Write-Log -Message ('-' * 72)
    Write-Log -Message 'Pre-flight checks'
    Write-Log -Message ('-' * 72)

    $psv = $PSVersionTable.PSVersion
    Write-Log -Message "PowerShell version: $psv ($($PSVersionTable.PSEdition))"

    $platformName = if ($IsWindows) { 'Windows' } elseif ($IsLinux) { 'Linux' } elseif ($IsMacOS) { 'macOS' } else { 'Unknown' }
    Write-Log -Message "Platform: $platformName"

    Write-Log -Message "Report period: $Period ($(Get-PeriodDescription -PeriodValue $Period))"
    Write-Log -Message 'Primary active-user source: Microsoft Graph (getMicrosoft365CopilotUsageUserDetail).'

    if ($IncludeUnifiedAuditLog) {
        if ($AuditOperations -and $AuditOperations.Count -gt 0) {
            Write-Log -Message "Unified Audit Log operations: $($AuditOperations -join ', ')"
        } else {
            Write-Log -Level 'WARN' -Message 'IncludeUnifiedAuditLog is set but -AuditOperations is empty; UAL step will be skipped.'
        }
        if (-not $IsWindows) {
            Write-Log -Level 'WARN' -Message 'Search-UnifiedAuditLog requires Exchange Online PowerShell on Windows. On this platform the UAL step will be skipped and the report will use Graph data only.'
        }
    } else {
        Write-Log -Message 'Unified Audit Log step: disabled by default (pass -IncludeUnifiedAuditLog to enable, or -SkipUnifiedAuditLog to force-disable).'
    }

    Write-Log -Message "Installing/loading required modules: $($RequiredModules -join ', ')"
    foreach ($m in $RequiredModules) {
        Initialize-RequiredModule -Name $m
    }

    Write-Log -Message 'Pre-flight checks complete. Proceeding to authentication.'
    Write-Log -Message ('-' * 72)
}

$preflightModules = [System.Collections.Generic.List[string]]::new()
$preflightModules.Add('Microsoft.Graph.Authentication')

# The Unified Audit Log step (ExchangeOnlineManagement) cannot share a PS 5.1
# session with Microsoft.Graph due to a Microsoft.Identity.Client assembly
# conflict (BrokerExtension.WithBroker overload mismatch). If the user wants
# UAL on Desktop PowerShell, prompt to upgrade to PowerShell 7+ via winget;
# on decline, auto-disable UAL up-front so the run doesn't bother loading
# EXO or attempting a doomed Exchange Online sign-in.
if ($IncludeUnifiedAuditLog -and $IsWindows -and $PSVersionTable.PSEdition -eq 'Desktop') {
    Write-Log -Level 'WARN' -Message 'Unified Audit Log step is enabled, but PowerShell 7+ is REQUIRED for it: Microsoft.Graph and ExchangeOnlineManagement cannot coexist in a Windows PowerShell 5.1 session (Microsoft.Identity.Client BrokerExtension.WithBroker assembly mismatch).'
    $upgradeAnswer = $null
    try {
        $upgradeAnswer = Read-Host -Prompt 'Install PowerShell 7 now via winget (recommended), or skip the Unified Audit Log step and continue? [Y] Upgrade  [N] Skip UAL'
    } catch {
        Write-Log -Level 'WARN' -Message "Could not prompt for upgrade decision: $($_.Exception.Message). Skipping UAL step."
    }

    if ($upgradeAnswer -and $upgradeAnswer.Trim().ToUpperInvariant() -in @('Y', 'YES')) {
        Write-Log -Message 'User chose to upgrade. Installing PowerShell 7 via winget per https://learn.microsoft.com/powershell/scripting/install/install-powershell-on-windows#winget.'

        $winget = Get-Command -Name 'winget' -ErrorAction SilentlyContinue
        if (-not $winget) {
            Write-Log -Level 'ERROR' -Message 'winget was not found on PATH. Install App Installer from the Microsoft Store, or download PowerShell 7 manually.'
            try { Start-Process 'https://learn.microsoft.com/powershell/scripting/install/install-powershell-on-windows#winget' | Out-Null } catch { }
            exit 1
        }

        Write-Log -Message 'Running: winget search Microsoft.PowerShell'
        & winget search Microsoft.PowerShell

        Write-Log -Message 'Running: winget install --id Microsoft.PowerShell --source winget'
        & winget install --id Microsoft.PowerShell --source winget --accept-source-agreements --accept-package-agreements
        $wingetExit = $LASTEXITCODE

        if ($wingetExit -eq 0) {
            Write-Log -Level 'SUCCESS' -Message 'PowerShell 7 installation completed.'
        } else {
            Write-Log -Level 'WARN' -Message "winget exited with code $wingetExit. Review the output above for details."
        }

        Write-Host ''
        Write-Host 'Open a NEW terminal (so PATH refreshes) and re-run this script with:' -ForegroundColor Yellow
        Write-Host '    pwsh -File .\Get-Agent365ActiveUsers.ps1' -ForegroundColor Yellow
        exit $wingetExit
    }

    Write-Log -Level 'WARN' -Message 'User declined the PowerShell 7 upgrade. Unified Audit Log step will be SKIPPED (PowerShell 7+ is required for it).'
    $IncludeUnifiedAuditLog = $false
    Add-SkippedItem -Item 'Unified Audit Log (Search-UnifiedAuditLog)' -Reason 'PowerShell 7+ is required for the Unified Audit Log step (Microsoft.Graph + ExchangeOnlineManagement cannot coexist on Windows PowerShell 5.1), and the user declined the upgrade.' -Remediation 'Re-run this script with pwsh 7+, or accept the upgrade prompt next time.'
}

if ($IncludeUnifiedAuditLog -and $IsWindows) {
    $preflightModules.Add('ExchangeOnlineManagement')
}

Invoke-Agent365Preflight `
    -Period $Period `
    -RequiredModules $preflightModules.ToArray() `
    -AuditOperations $AuditOperations `
    -IncludeUnifiedAuditLog:$IncludeUnifiedAuditLog

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

function Test-Agent365ReportsAccess {
    <#
    .SYNOPSIS
        Checks whether the signed-in account holds a Microsoft Entra admin
        role that qualifies for the per-user Microsoft 365 Copilot usage
        report (getMicrosoft365CopilotUsageUserDetail).

    .DESCRIPTION
        Reports.Read.All Graph scope alone is NOT sufficient for the
        per-user Copilot usage report. The signed-in account must also
        hold one of the qualifying directory roles documented at:
          https://learn.microsoft.com/graph/reportroot-authorization
          https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail

        Roles split into two tiers:

        * FullAccess: returns per-user detail (this script needs this).
          Global Admin, AI Admin, Reports Reader, Exchange / SharePoint /
          Teams Service / Teams Communications / Skype (Lync) Admin.

        * TenantOnly: documented as authorized for the report but
          Microsoft restricts these roles to "tenant-level data, without
          visibility into detailed metrics". The per-user call may
          return 403 or empty per-user records depending on which Graph
          endpoint variant is hit. Global Reader and Usage Summary
          Reports Reader are in this tier.

        This function enumerates the signed-in user's transitive directory
        roles (so PIM and group-based assignments are caught) and matches
        them against both tiers by roleTemplateId (stable) and displayName
        (readable).

        Returns a [pscustomobject] with HasQualifyingRole (= FullAccess),
        HasTenantOnlyRole, AssignedRoles, FullAccessRoles, TenantOnlyRoles,
        and a Checked flag (false if the lookup itself failed). The caller
        decides how to warn.
    #>
    [CmdletBinding()]
    param()

    # Microsoft Entra admin roles that grant FULL access to the per-user
    # Microsoft 365 Copilot usage report (including detailed per-user
    # records). roleTemplateId is the stable identifier; displayName is
    # the human-readable name (may differ slightly across tenants).
    $fullAccessRoles = @(
        [pscustomobject]@{ DisplayName = 'Global Administrator';              TemplateId = '62e90394-69f5-4237-9190-012177145e10'; AliasNames = @('Company Administrator') }
        [pscustomobject]@{ DisplayName = 'AI Administrator';                  TemplateId = 'd2562ede-74db-457e-a7b6-544e236ebb61'; AliasNames = @() }
        [pscustomobject]@{ DisplayName = 'Reports Reader';                    TemplateId = '4a5d8f65-41da-4de4-8968-e035b65339cf'; AliasNames = @() }
        [pscustomobject]@{ DisplayName = 'Exchange Administrator';            TemplateId = '29232cdf-9323-42fd-ade2-1d097af3e4de'; AliasNames = @() }
        [pscustomobject]@{ DisplayName = 'SharePoint Administrator';          TemplateId = 'f28a1f50-f6e7-4571-818b-6a12f2af6b6c'; AliasNames = @() }
        [pscustomobject]@{ DisplayName = 'Teams Administrator';               TemplateId = '69091246-20e8-4a56-aa4d-066075b2a7a8'; AliasNames = @('Teams Service Administrator') }
        [pscustomobject]@{ DisplayName = 'Teams Communications Administrator';TemplateId = 'baf37b3a-610e-45da-9e62-d9d1e5e8914b'; AliasNames = @() }
        [pscustomobject]@{ DisplayName = 'Skype for Business Administrator';  TemplateId = '75941009-915a-4869-abe7-691bff18279e'; AliasNames = @('Lync Administrator') }
    )

    # Roles documented as authorized but restricted to TENANT-LEVEL data
    # only. The per-user detail report will either return 403 or omit
    # per-user records when only one of these is held.
    $tenantOnlyRoles = @(
        [pscustomobject]@{ DisplayName = 'Global Reader';                  TemplateId = 'f2ef992c-3afb-46b9-b7cf-a126ee74c451' }
        [pscustomobject]@{ DisplayName = 'Usage Summary Reports Reader';   TemplateId = '75934031-6c7e-415a-99d7-48dbd49e875e' }
    )

    $result = [pscustomobject]@{
        Checked            = $false
        HasQualifyingRole  = $false
        HasTenantOnlyRole  = $false
        UserPrincipalName  = $null
        AssignedRoles      = @()
        FullAccessRoles    = @()
        TenantOnlyRoles    = @()
        QualifyingRoleList = @($fullAccessRoles | Select-Object -ExpandProperty DisplayName -Unique)
        TenantOnlyRoleList = @($tenantOnlyRoles | Select-Object -ExpandProperty DisplayName -Unique)
    }

    # Pull /me first so we can log who we checked even if the role pull fails.
    try {
        $me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,displayName,id'
        if ($me -and $me.userPrincipalName) {
            $result.UserPrincipalName = [string]$me.userPrincipalName
        }
    } catch {
        Write-VerboseLog -Message "/me lookup failed during role check: $($_.Exception.Message)"
    }

    # transitiveMemberOf catches PIM activations and nested group-based
    # role assignments; memberOf would miss those. Filter server-side to
    # directoryRole entities to minimise payload.
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=id,displayName,roleTemplateId&$top=200'
    } catch {
        Write-VerboseLog -Message "transitiveMemberOf lookup failed during role check: $($_.Exception.Message)"
        return $result
    }

    $result.Checked = $true

    $roles = @()
    if ($resp -and $resp.value) { $roles = @($resp.value) }

    $assignedTemplateIds = @($roles | ForEach-Object { $_.roleTemplateId } | Where-Object { $_ })
    $assignedDisplay     = @($roles | ForEach-Object { $_.displayName }    | Where-Object { $_ })
    # Wrap Sort-Object -Unique in @(...) -- on empty/single input it returns
    # $null or a scalar, which would break .Count access under StrictMode.
    $result.AssignedRoles = @($assignedDisplay | Sort-Object -Unique)

    # Helper: project assigned roles back to canonical display names for
    # a given tier definition, matching by either roleTemplateId or
    # displayName.
    $projectMatches = {
        param($tierDefs, $assignedIds, $assignedNames)

        $tierTemplateIds  = @($tierDefs | Select-Object -ExpandProperty TemplateId  -Unique)
        $tierDisplayNames = @($tierDefs | Select-Object -ExpandProperty DisplayName -Unique)
        $tierAliasNames   = @(
            $tierDefs |
                ForEach-Object {
                    if ($_.PSObject.Properties.Match('AliasNames').Count -gt 0) {
                        @($_.AliasNames)
                    }
                } |
                Where-Object { $_ }
        )
        $tierAllNames     = @($tierDisplayNames + $tierAliasNames | Sort-Object -Unique)

        $matchedTemplateIds = @($assignedIds   | Where-Object { $tierTemplateIds  -contains $_ })
        $matchedDisplayHits = @($assignedNames | Where-Object { $tierAllNames -contains $_ })

        $matched = @()
        foreach ($tid in ($matchedTemplateIds | Sort-Object -Unique)) {
            $name = ($tierDefs | Where-Object { $_.TemplateId -eq $tid } | Select-Object -First 1).DisplayName
            if ($name) { $matched += $name }
        }
        foreach ($name in $matchedDisplayHits) {
            if ($matched -notcontains $name) {
                $canonical = ($tierDefs | Where-Object { $_.DisplayName -eq $name -or ($_.AliasNames -contains $name) } | Select-Object -First 1).DisplayName
                if ($canonical -and ($matched -notcontains $canonical)) {
                    $matched += $canonical
                }
            }
        }
        return @($matched | Sort-Object -Unique)
    }

    # Wrap scriptblock invocations in @(...) -- PowerShell unwraps an empty
    # array returned via the call operator (&) back to $null, which would
    # then fail .Count access under StrictMode.
    $result.FullAccessRoles   = @(& $projectMatches $fullAccessRoles $assignedTemplateIds $assignedDisplay)
    $result.TenantOnlyRoles   = @(& $projectMatches $tenantOnlyRoles $assignedTemplateIds $assignedDisplay)
    $result.HasQualifyingRole = ($result.FullAccessRoles.Count -gt 0)
    $result.HasTenantOnlyRole = ($result.TenantOnlyRoles.Count -gt 0)

    return $result
}

function Test-CanLaunchBrowser {
    [CmdletBinding()]
    param()

    # Explicit environment overrides
    if ($env:AGENT365_FORCE_DEVICE_CODE -eq 'true') { return $false }
    if ($env:AGENT365_FORCE_BROWSER     -eq 'true') { return $true }

    # Non-interactive PowerShell host (e.g. CI / scheduled task)
    try {
        if (-not [Environment]::UserInteractive) { return $false }
        $hostName = $Host.Name
        if ($hostName -like '*Default Host*' -or $hostName -like '*ServerRemoteHost*') {
            return $false
        }
    } catch {
        return $false
    }

    # SSH session (no local browser)
    if ($env:SSH_CLIENT -or $env:SSH_TTY -or $env:SSH_CONNECTION) { return $false }

    # GitHub Codespaces / VS Code dev containers
    if ($env:CODESPACES         -eq 'true') { return $false }
    if ($env:CODESPACE_NAME)                 { return $false }
    if ($env:REMOTE_CONTAINERS  -eq 'true') { return $false }
    if ($env:DEVCONTAINER       -eq 'true') { return $false }

    # Generic Linux container marker
    if ($IsLinux -and (Test-Path '/.dockerenv')) { return $false }

    # Linux without an X / Wayland display (no GUI)
    if ($IsLinux -and -not $env:DISPLAY -and -not $env:WAYLAND_DISPLAY) { return $false }

    return $true
}

function Test-CanPromptOnStdin {
    try {
        return (-not [System.Console]::IsInputRedirected)
    } catch {
        return $false
    }
}

function Open-ReportInDefaultBrowser {
    <#
    .SYNOPSIS
        Best-effort launch of the generated HTML report in the OS default
        browser. No-op (returns $false) on headless / non-interactive hosts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    if (-not (Test-CanLaunchBrowser))        { return $false }

    $full = [System.IO.Path]::GetFullPath($Path)
    try {
        if ($IsWindows) {
            Start-Process -FilePath $full -ErrorAction Stop | Out-Null
        } elseif ($IsMacOS) {
            Start-Process -FilePath 'open' -ArgumentList $full -ErrorAction Stop | Out-Null
        } else {
            # Linux desktop
            Start-Process -FilePath 'xdg-open' -ArgumentList $full -ErrorAction Stop | Out-Null
        }
        return $true
    } catch {
        Write-Verbose "Could not auto-launch browser: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-DeviceCodeMgGraphSignIn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,
        [int]$MaxAttempts = 3,
        [switch]$Immediate
    )

    $canPrompt = Test-CanPromptOnStdin

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        # On a retry, clear any stale Graph session so the next device code
        # doesn't collide with a half-cached identity from the previous attempt.
        if ($attempt -gt 1) {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose "Disconnect-MgGraph cleanup ignored: $($_.Exception.Message)" }
        }

        Write-Host ''
        Write-Host '=== Microsoft Graph sign-in (device code) ==='
        Write-Host ''
        Write-Host 'Microsoft enforces a 120-second timeout once a code is generated.'
        Write-Host 'For best results, OPEN this URL FIRST in any browser:'
        Write-Host ''
        Write-Host '    https://login.microsoft.com/device'
        Write-Host ''

        if ($Immediate) {
            Write-Host 'Generating a device code now.'
        } elseif ($canPrompt) {
            Write-Host ("When the sign-in page is open and ready, press Enter to receive a fresh code (attempt {0} of {1})..." -f $attempt, $MaxAttempts)
            [void](Read-Host)
        } else {
            Write-Host ("Non-interactive shell detected; requesting device code now (attempt {0} of {1})." -f $attempt, $MaxAttempts)
        }

        try {
            Connect-MgGraph -Scopes $Scopes -NoWelcome -UseDeviceCode
            return
        } catch {
            $message = $_.Exception.Message

            # MSAL sometimes reports "timed out due to inactivity" even after the
            # user's browser-side sign-in succeeded. Re-check Get-MgContext: if a
            # valid context with all required scopes is now present, treat the
            # call as successful instead of forcing another device code.
            $ctxAfter = $null
            try { $ctxAfter = Get-MgContext } catch { $ctxAfter = $null }
            if ($ctxAfter -and $ctxAfter.Account) {
                $missingScopes = @($Scopes | Where-Object { $ctxAfter.Scopes -notcontains $_ })
                if ($missingScopes.Count -eq 0) {
                    Write-Log -Message ("Microsoft Graph sign-in raised '{0}' but an authenticated context with all required scopes is present (account: {1}). Continuing." -f $message, $ctxAfter.Account)
                    return
                }
            }

            $isTransient = ($message -match 'timed out|authorization_pending|expired_token|inactivity')
            if ($attempt -lt $MaxAttempts -and $isTransient) {
                Write-Log -Level 'WARN' -Message ("Microsoft Graph sign-in attempt {0} failed: {1}. Retrying with a fresh device code." -f $attempt, $message)
                continue
            }
            throw
        }
    }
}

function Connect-Agent365Graph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,
        [switch]$UseDeviceCode
    )

    # Inspect any cached Graph session and ask the user whether to reuse
    # that account or sign in as a different one. We never silently assume
    # the cached identity is the one the operator wants — wrong-account
    # sign-ins are the most common cause of 403 errors on this report.
    $existingCtx = $null
    try { $existingCtx = Get-MgContext } catch { $existingCtx = $null }

    $loginHint = $null
    $forceDisconnect = $false

    if ($existingCtx -and $existingCtx.Account) {
        if ($PickSignInAccount) {
            $forceDisconnect = $true
            Write-Log -Message "-PickSignInAccount specified; clearing cached Microsoft Graph session ($($existingCtx.Account)) so account picker can be used."
        }
        if ($SignInAccount -and $existingCtx.Account -ne $SignInAccount) {
            $forceDisconnect = $true
            Write-Log -Message "Cached Microsoft Graph account ($($existingCtx.Account)) does not match requested sign-in account ($SignInAccount). Clearing cached session."
        }
        Write-Host ''
        Write-Host ("A Microsoft Graph session is already cached for: {0}" -f $existingCtx.Account) -ForegroundColor Cyan
        if (-not $forceDisconnect) {
            $reuse = $null
            try {
                $reuse = Read-Host -Prompt 'Use this account? [Y] Yes  [N] Sign in as a different account'
            } catch {
                Write-Log -Level 'WARN' -Message "Could not prompt for account selection: $($_.Exception.Message). Reusing cached account $($existingCtx.Account)."
                $reuse = 'Y'
            }
            if ($reuse -and $reuse.Trim().ToUpperInvariant() -in @('N', 'NO')) {
                $forceDisconnect = $true
                Write-Log -Message "User chose to sign in with a different account; clearing cached session ($($existingCtx.Account))."
            } else {
                Write-Log -Message "Reusing cached Microsoft Graph account: $($existingCtx.Account)."
            }
        } else {
            Write-Log -Message "Will sign in again to satisfy requested account hint."
        }
    } else {
        Write-Host ''
        Write-Host 'No cached Microsoft Graph session found.' -ForegroundColor Cyan
    }

    if (-not $existingCtx -or $forceDisconnect) {
        if ($forceDisconnect) {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose "Disconnect-MgGraph cleanup ignored: $($_.Exception.Message)" }
        }
    }

    # Use the optional -SignInAccount parameter as a display-only login
    # hint. We do NOT prompt the user for an account; if they want to
    # target a specific one they pass it via -SignInAccount.
    if ($SignInAccount) {
        $loginHint = $SignInAccount.Trim()
        if ($loginHint) {
            Write-Log -Message "Target sign-in account (from -SignInAccount): $loginHint"
        }
    }

    $useDevice = $UseDeviceCode.IsPresent
    if ($useDevice) {
        Write-Log -Message 'Device code authentication requested for Microsoft Graph.'
    } elseif (-not (Test-CanLaunchBrowser)) {
        Write-Log -Message 'No interactive browser detected in this environment. Using device code flow for Microsoft Graph.'
        $useDevice = $true
    }

    if ($useDevice) {
        if (-not $NoProgress) { Write-Progress -Activity 'Agent 365 report' -Completed }
        if ($loginHint) {
            Write-Host ("When prompted on https://login.microsoft.com/device, sign in as: {0}" -f $loginHint) -ForegroundColor Yellow
        }
        Invoke-DeviceCodeMgGraphSignIn -Scopes $Scopes -Immediate:$UseDeviceCode.IsPresent
        return
    }

    try {
        Write-Log -Message 'Attempting interactive (browser) Microsoft Graph sign-in.'
        if ($loginHint) {
            Write-Host ("In the browser account picker, select: {0}" -f $loginHint) -ForegroundColor Yellow
        }
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
    } catch {
        Write-Log -Level 'WARN' -Message "Interactive Microsoft Graph sign-in failed: $($_.Exception.Message). Falling back to device code."
        if (-not $NoProgress) { Write-Progress -Activity 'Agent 365 report' -Completed }
        Invoke-DeviceCodeMgGraphSignIn -Scopes $Scopes
    }

    # Confirm the account that actually signed in matches what the user
    # asked for via -SignInAccount; warn loudly if it doesn't so they
    # don't waste time debugging a 403 caused by the wrong identity.
    if ($loginHint) {
        try {
            $signedIn = (Get-MgContext).Account
            if ($signedIn -and $signedIn -ne $loginHint) {
                Write-Log -Level 'WARN' -Message ("Signed-in account ({0}) does not match -SignInAccount ({1}). If you hit 403 errors, run Disconnect-MgGraph and re-run." -f $signedIn, $loginHint)
            }
        } catch { }
    }
}

function Connect-Agent365ExchangeOnline {
    [CmdletBinding()]
    param(
        [switch]$UseDeviceCode
    )

    $useDevice = $UseDeviceCode.IsPresent
    if ($useDevice) {
        Write-Log -Message 'Device code authentication requested for Exchange Online.'
    } elseif (-not (Test-CanLaunchBrowser)) {
        Write-Log -Message 'No interactive browser detected in this environment. Using device code flow for Exchange Online.'
        $useDevice = $true
    } elseif ($IsWindows) {
        # ExchangeOnlineManagement uses the Windows broker/WAM path for its
        # interactive flow. In this script that path is brittle on Windows,
        # especially after Microsoft Graph has already loaded MSAL into the
        # session. Use device code instead so the UAL step stays reliable.
        Write-Log -Message 'Windows detected: using device code flow for Exchange Online to avoid broker/WAM sign-in failures.'
        $useDevice = $true
    }

    if ($useDevice) {
        if (-not $NoProgress) { Write-Progress -Activity 'Agent 365 report' -Completed }
        Invoke-DeviceCodeExoSignIn -Immediate:$UseDeviceCode.IsPresent
        return
    }

    try {
        Write-Log -Message 'Attempting interactive (browser) Exchange Online sign-in.'
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    } catch {
        Write-Log -Level 'WARN' -Message "Interactive Exchange Online sign-in failed: $($_.Exception.Message). Falling back to device code."
        if (-not $NoProgress) { Write-Progress -Activity 'Agent 365 report' -Completed }
        Invoke-DeviceCodeExoSignIn
    }
}

function Invoke-Agent365UnifiedAuditLogQuery {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [int]$Days,

        [Parameter(Mandatory = $true)]
        [string[]]$Operations,

        [Parameter(Mandatory = $true)]
        [int]$ResultSize,

        [Parameter(Mandatory = $true)]
        [string]$OutputJsonPath,

        [Parameter(Mandatory = $true)]
        [string]$ErrorLogPath,

        [switch]$UseDeviceCode,
        [switch]$PickSignInAccount
    )

    $pwshCommand = Get-Command -Name 'pwsh' -ErrorAction Stop
    $tempScriptPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.ps1" -f ([guid]::NewGuid().ToString('N')))

    $childScript = @'
param(
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [int]$Days,

    [Parameter(Mandatory = $true)]
    [string[]]$Operations,

    [Parameter(Mandatory = $true)]
    [int]$ResultSize

    , [Parameter(Mandatory = $true)]
    [string]$OutputJsonPath

    , [Parameter(Mandatory = $true)]
    [string]$ErrorLogPath

    , [switch]$UseDeviceCode

    , [switch]$PickSignInAccount
)

try {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    Set-StrictMode -Version Latest

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $ipssConnectCommand = Get-Command -Name 'Connect-IPPSSession' -ErrorAction Stop
    $ipssSupportsDisableWam = $ipssConnectCommand.Parameters.ContainsKey('DisableWAM')
    $ipssSupportsUseRps = $ipssConnectCommand.Parameters.ContainsKey('UseRPSSession')
    $ipssSupportsDevice = $ipssConnectCommand.Parameters.ContainsKey('Device')
    $ipssSupportsUseDeviceAuth = $ipssConnectCommand.Parameters.ContainsKey('UseDeviceAuthentication')

    function Connect-Agent365IppsSessionWithFallback {
        param(
            [hashtable]$BaseParams = @{},
            [switch]$UseDeviceCode
        )

        $attempts = [System.Collections.Generic.List[hashtable]]::new()

        # Default connection is generally the most stable. Restricting import
        # with -CommandName can hang in some tenants/sessions, so try it last.
        $attempts.Add(@{ Label = 'default'; Params = @{ ShowBanner = $false } })

        if ($ipssSupportsUseRps) {
            $attempts.Add(@{ Label = 'UseRPSSession'; Params = @{ ShowBanner = $false; UseRPSSession = $true } })
        }

        $attempts.Add(@{ Label = 'CommandName=Search-UnifiedAuditLog'; Params = @{ ShowBanner = $false; CommandName = 'Search-UnifiedAuditLog' } })

        $lastError = $null
        foreach ($attempt in $attempts) {
            $params = @{} + $attempt.Params
            foreach ($kv in $BaseParams.GetEnumerator()) {
                $params[$kv.Key] = $kv.Value
            }
            if ($ipssSupportsDisableWam) {
                $params.DisableWAM = $true
            }
            if ($UseDeviceCode) {
                if ($ipssSupportsDevice) {
                    $params.Device = $true
                } elseif ($ipssSupportsUseDeviceAuth) {
                    $params.UseDeviceAuthentication = $true
                }
            }

            Write-Host ("Connecting Security & Compliance session using {0}..." -f $attempt.Label)
            try {
                Connect-IPPSSession @params -ErrorAction Stop | Out-Null
                Write-Host ("Security & Compliance connection established using {0}." -f $attempt.Label)
                return
            } catch {
                $lastError = $_
                Write-Host ("Security & Compliance connection attempt failed ({0}): {1}" -f $attempt.Label, $_.Exception.Message)
            }
        }

        if ($lastError -and $lastError.Exception.Message -match 'Error Acquiring Token') {
            throw [System.InvalidOperationException]::new(
                "Security & Compliance sign-in failed with token acquisition errors after all connection strategies. Try re-running with -UseDeviceCode again and complete sign-in in the same browser profile where your tenant account is active. If this persists, re-run with -SkipUnifiedAuditLog to finish the report while EXO auth is investigated. Last error: $($lastError.Exception.Message)"
            )
        }

        if ($lastError) {
            throw [System.InvalidOperationException]::new("All Security & Compliance connection strategies failed. Last error: $($lastError.Exception.Message)")
        }

        throw [System.InvalidOperationException]::new('All Security & Compliance connection strategies failed with no detailed error.')
    }

    function Test-EnsureSearchUnifiedAuditLogAvailable {
        param(
            [hashtable]$BaseParams = @{},
            [switch]$UseDeviceCode
        )

        if (Get-Command -Name 'Search-UnifiedAuditLog' -ErrorAction SilentlyContinue) {
            return $true
        }

        Write-Host 'Search-UnifiedAuditLog not yet available. Attempting manual import from compliance remote session...'
        $complianceSessions = @(
            Get-PSSession -ErrorAction SilentlyContinue |
                Where-Object {
                    ($_.ComputerName -like '*ps.compliance.protection.outlook.com*') -or
                    ($_.ComputerName -like '*outlook.office365.com*')
                }
        )

        foreach ($session in $complianceSessions) {
            try {
                Import-PSSession -Session $session -CommandName 'Search-UnifiedAuditLog' -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
                if (Get-Command -Name 'Search-UnifiedAuditLog' -ErrorAction SilentlyContinue) {
                    Write-Host 'Search-UnifiedAuditLog imported successfully from compliance remote session.'
                    return $true
                }
            } catch {
                Write-Host ("Manual import attempt failed from session '{0}': {1}" -f $session.Name, $_.Exception.Message)
            }
        }

        if (-not $ipssSupportsUseRps) {
            return $false
        }

        Write-Host 'Search-UnifiedAuditLog still unavailable. Reconnecting Security & Compliance session with UseRPSSession for cmdlet import...'
        $rpsParams = @{ ShowBanner = $false; UseRPSSession = $true }
        foreach ($kv in $BaseParams.GetEnumerator()) {
            $rpsParams[$kv.Key] = $kv.Value
        }
        if ($ipssSupportsDisableWam) {
            $rpsParams.DisableWAM = $true
        }
        if ($UseDeviceCode) {
            if ($ipssSupportsDevice) {
                $rpsParams.Device = $true
            } elseif ($ipssSupportsUseDeviceAuth) {
                $rpsParams.UseDeviceAuthentication = $true
            }
        }

        try {
            Connect-IPPSSession @rpsParams -ErrorAction Stop | Out-Null
        } catch {
            Write-Host ("UseRPSSession reconnect failed: {0}" -f $_.Exception.Message)
            return $false
        }

        return [bool](Get-Command -Name 'Search-UnifiedAuditLog' -ErrorAction SilentlyContinue)
    }

    function Test-HasUnifiedAuditLogPermission {
        param(
            [Parameter(Mandatory = $true)]
            [string]$PrincipalName
        )

    $allowedRoleNames = @(
        'Audit Logs',
        'View-Only Audit Logs',
        'Audit Manager',
        'Audit Reader',
        'Organization Management',
        'Global Administrator',
        'Company Administrator',
        'Compliance Administrator',
        'Compliance Data Administrator',
        'Security Administrator',
        'Security Operator',
        'Global Reader'
    )

    try {
        $roleAssignments = @(Get-ManagementRoleAssignment -RoleAssignee $PrincipalName -Enabled $true -ErrorAction Stop)
        foreach ($assignment in $roleAssignments) {
            $roleName = $null
            try { $roleName = [string]$assignment.Role } catch { $roleName = $null }
            if ($roleName -and ($allowedRoleNames -contains $roleName.Trim())) {
                return $true
            }
        }
    } catch {
        Write-Verbose "Get-ManagementRoleAssignment lookup failed: $($_.Exception.Message)"
    }

    foreach ($roleGroup in @('Audit Logs', 'View-Only Audit Logs')) {
        try {
            $members = @(Get-RoleGroupMember $roleGroup -ErrorAction Stop)
        } catch {
            continue
        }

        foreach ($member in $members) {
            foreach ($propertyName in @('UserPrincipalName', 'PrimarySmtpAddress', 'WindowsLiveID', 'DisplayName', 'Name', 'EmailAddress')) {
                $memberValue = $null
                try { $memberValue = [string]$member.$propertyName } catch { $memberValue = $null }
                if ($memberValue -and $memberValue.Trim().ToUpperInvariant() -eq $PrincipalName.Trim().ToUpperInvariant()) {
                    return $true
                }
            }
        }
    }

    return $false
}

    $useDeviceCode = $UseDeviceCode.IsPresent
    if ($PickSignInAccount.IsPresent -and -not $useDeviceCode) {
        Write-Host 'PickSignInAccount requested: using interactive browser sign-in for Unified Audit Log session.'
        $exoInteractiveParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($UserPrincipalName) {
            $exoInteractiveParams.UserPrincipalName = $UserPrincipalName
            Write-Host ("Unified Audit Log account hint for Exchange Online: {0}" -f $UserPrincipalName)
        }

        try {
            Connect-ExchangeOnline @exoInteractiveParams | Out-Null
        } catch {
            Write-Host ("Interactive Exchange Online sign-in failed in picker mode: {0}. Falling back to device code for UAL session." -f $_.Exception.Message)
            $useDeviceCode = $true
        }
    }

    if ($useDeviceCode) {
        Write-Host 'Using device code authentication for Unified Audit Log access.'
        $exoConnectParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($UserPrincipalName) {
            $exoConnectParams.UserPrincipalName = $UserPrincipalName
            Write-Host ("Unified Audit Log account hint for Exchange Online: {0}" -f $UserPrincipalName)
        }
        try {
            Connect-ExchangeOnline @exoConnectParams -Device | Out-Null
        } catch [System.Management.Automation.ParameterBindingException] {
            Connect-ExchangeOnline @exoConnectParams -UseDeviceAuthentication | Out-Null
        }
        Write-Host 'Exchange Online device-code sign-in completed. Connecting Security & Compliance session (this can take up to a few minutes)...'
        if (-not $ipssSupportsDevice -and -not $ipssSupportsUseDeviceAuth) {
            Write-Host 'Connect-IPPSSession does not expose a device-code switch in this module version; using default auth for IPPS session.'
        }
        Connect-Agent365IppsSessionWithFallback -UseDeviceCode
        Write-Host 'Security & Compliance session connected. Verifying Search-UnifiedAuditLog availability...'
        if (-not (Test-EnsureSearchUnifiedAuditLogAvailable -UseDeviceCode)) {
            throw [System.InvalidOperationException]::new('Device code authentication completed, but Search-UnifiedAuditLog could not be imported into the session (manual import and UseRPSSession recovery both failed).')
        }
    } else {
        $connectParams = @{}
        if ($UserPrincipalName) {
            $connectParams.UserPrincipalName = $UserPrincipalName
            Write-Host ("Unified Audit Log account hint for Security & Compliance: {0}" -f $UserPrincipalName)
        }

        Connect-Agent365IppsSessionWithFallback -BaseParams $connectParams
        if (-not (Test-EnsureSearchUnifiedAuditLogAvailable -BaseParams $connectParams)) {
            throw [System.InvalidOperationException]::new('Security & Compliance authentication completed, but Search-UnifiedAuditLog could not be imported into the session (manual import and UseRPSSession recovery both failed).')
        }
    }

    Write-Host 'Unified Audit Log authentication complete. Running Search-UnifiedAuditLog query...'

$effectivePrincipal = $null
if ($UserPrincipalName) {
    $effectivePrincipal = $UserPrincipalName
} else {
    try {
        $activeConnection = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        if ($activeConnection -and $activeConnection.UserPrincipalName) {
            $effectivePrincipal = [string]$activeConnection.UserPrincipalName
        }
    } catch {
        $effectivePrincipal = $null
    }
}

if ($effectivePrincipal) {
    Write-Host ("Unified Audit Log effective signed-in account: {0}" -f $effectivePrincipal)
    if (-not (Test-HasUnifiedAuditLogPermission -PrincipalName $effectivePrincipal)) {
        throw [System.InvalidOperationException]::new(
            "Signed-in account '$effectivePrincipal' does not appear to have an audit-log role assignment that grants Search-UnifiedAuditLog access. Assign Audit Logs / View-Only Audit Logs, or one of the equivalent admin roles documented by Microsoft (for example Global Administrator via Organization Management, Global Reader via View-Only Audit Logs, or Audit Manager / Audit Reader), then re-run."
        )
    }
} else {
    Write-Host 'Unified Audit Log effective signed-in account could not be determined; continuing without explicit audit-role pre-check.'
}

$startDate = (Get-Date).ToUniversalTime().AddDays(-$Days)
$endDate = (Get-Date).ToUniversalTime()
    $records = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -Operations $Operations -ResultSize $ResultSize

    Write-Host ("Unified Audit Log query complete. Records returned: {0}" -f @($records).Count)

    $records | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8
} catch {
    $detail = $_ | Out-String
    Set-Content -LiteralPath $ErrorLogPath -Value $detail -Encoding UTF8
    throw
}
'@

    Set-Content -LiteralPath $tempScriptPath -Value $childScript -Encoding UTF8

    try {
        Write-Log -Message 'Connecting to Security & Compliance PowerShell in a separate pwsh process for Unified Audit Log access.'

        $outputJsonPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.json" -f ([guid]::NewGuid().ToString('N')))
        $errorLogPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.err.txt" -f ([guid]::NewGuid().ToString('N')))

        $childArgs = @(
            '-NoLogo'
            '-NoProfile'
            '-File'
            $tempScriptPath
            '-Days'
            $Days
            '-Operations'
            $Operations
            '-ResultSize'
            $ResultSize
            '-OutputJsonPath'
            $outputJsonPath
            '-ErrorLogPath'
            $errorLogPath
        )
        if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            $childArgs += @('-UserPrincipalName', $UserPrincipalName)
        }
        if ($UseDeviceCode) {
            $childArgs += @('-UseDeviceCode')
        }
        if ($PickSignInAccount) {
            $childArgs += @('-PickSignInAccount')
        }

        $childLogPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.log" -f ([guid]::NewGuid().ToString('N')))
        & $pwshCommand.Path @childArgs 2>&1 | Tee-Object -FilePath $childLogPath | Out-Host

        if ($LASTEXITCODE -ne 0) {
            $message = "pwsh exited with code $LASTEXITCODE while querying Unified Audit Log."
            if (Test-Path -LiteralPath $childLogPath) {
                $childError = (Get-Content -LiteralPath $childLogPath -Raw).Trim()
                if ($childError) {
                    $message = "$message Original child error:`n$childError"
                }
            }
            if (Test-Path -LiteralPath $errorLogPath) {
                $errorDetail = (Get-Content -LiteralPath $errorLogPath -Raw).Trim()
                if ($errorDetail) {
                    $message = "$message Detailed child exception:`n$errorDetail"
                }
            }
            throw [System.InvalidOperationException]::new($message)
        }

        $json = ''
        if (Test-Path -LiteralPath $outputJsonPath) {
            $json = (Get-Content -LiteralPath $outputJsonPath -Raw).Trim()
        }
        if (-not $json) {
            return @()
        }

        $records = $json | ConvertFrom-Json
        if ($null -eq $records) {
            return @()
        }

        if ($records -isnot [System.Array]) {
            return @($records)
        }

        return @($records)
    } catch {
        if (Test-Path -LiteralPath $childLogPath) {
            Write-Log -Level 'ERROR' -Message "Unified Audit child console log preserved at: $childLogPath"
        }
        if (Test-Path -LiteralPath $errorLogPath) {
            Write-Log -Level 'ERROR' -Message "Unified Audit child exception log preserved at: $errorLogPath"
        }
        if (Test-Path -LiteralPath $errorLogPath) {
            try {
                $errorDetail = (Get-Content -LiteralPath $errorLogPath -Raw).Trim()
                if ($errorDetail) {
                    Write-Log -Level 'ERROR' -Message "Unified Audit child exception detail:`n$errorDetail"
                }
            } catch { }
        }
        throw [System.InvalidOperationException]::new(
            "Failed to query Unified Audit Log in a separate PowerShell session. Original error: $($_.Exception.Message)"
        )
    } finally {
        if (Test-Path -LiteralPath $outputJsonPath) {
            Remove-Item -LiteralPath $outputJsonPath -ErrorAction SilentlyContinue
        }
        if ((Test-Path -LiteralPath $childLogPath) -and $LASTEXITCODE -eq 0) {
            Remove-Item -LiteralPath $childLogPath -ErrorAction SilentlyContinue
        }
        if ((Test-Path -LiteralPath $errorLogPath) -and $LASTEXITCODE -eq 0) {
            Remove-Item -LiteralPath $errorLogPath -ErrorAction SilentlyContinue
        }
        Remove-Item -LiteralPath $tempScriptPath -ErrorAction SilentlyContinue
    }
}

function Invoke-DeviceCodeExoSignIn {
    [CmdletBinding()]
    param(
        [int]$MaxAttempts = 3,
        [switch]$Immediate
    )

    $canPrompt = Test-CanPromptOnStdin

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        # On a retry, clear any half-established EXO session to avoid the next
        # device code colliding with a stale connection.
        if ($attempt -gt 1) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose "Disconnect-ExchangeOnline cleanup ignored: $($_.Exception.Message)" }
        }

        Write-Host ''
        Write-Host '=== Exchange Online sign-in (device code) ==='
        Write-Host ''
        Write-Host 'Microsoft enforces a short timeout once a code is generated.'
        Write-Host 'For best results, OPEN this URL FIRST in any browser:'
        Write-Host ''
        Write-Host '    https://login.microsoft.com/device'
        Write-Host ''

        if ($Immediate) {
            Write-Host 'Generating a device code now.'
        } elseif ($canPrompt) {
            Write-Host ("When the sign-in page is open and ready, press Enter to receive a fresh code (attempt {0} of {1})..." -f $attempt, $MaxAttempts)
            [void](Read-Host)
        } else {
            Write-Host ("Non-interactive shell detected; requesting device code now (attempt {0} of {1})." -f $attempt, $MaxAttempts)
        }

        try {
            # ExchangeOnlineManagement renamed its device-code switch across
            # versions: v3.x uses -Device, earlier 2.0.x builds use
            # -UseDeviceAuthentication. Get-Command.Parameters does not
            # always surface these on PS 5.1 (proxy functions hide dynamic
            # parameters), so probe by trying the calls directly.
            try {
                Connect-ExchangeOnline -ShowBanner:$false -Device -ErrorAction Stop
            } catch [System.Management.Automation.ParameterBindingException] {
                Write-VerboseLog -Message '-Device not accepted by this Connect-ExchangeOnline; trying -UseDeviceAuthentication.'
                Connect-ExchangeOnline -ShowBanner:$false -UseDeviceAuthentication -ErrorAction Stop
            }
            return
        } catch {
            $message = $_.Exception.Message

            # Hard-fail fast on the well-known PS 5.1 + Microsoft.Graph +
            # ExchangeOnlineManagement MSAL assembly conflict. Microsoft.Graph
            # loads an older Microsoft.Identity.Client into the AppDomain
            # first; EXO then tries to call a newer BrokerExtension.WithBroker
            # overload that doesn't exist on that older assembly. Retrying
            # device codes will not fix this -- the only workarounds are
            # running on PowerShell 7+, or running EXO in a separate process.
            if ($message -match 'BrokerExtension\.WithBroker' -or $_.Exception -is [System.MissingMethodException]) {
                throw [System.InvalidOperationException]::new(
                    "Exchange Online sign-in cannot proceed in this session due to a Microsoft.Identity.Client assembly conflict between Microsoft.Graph and ExchangeOnlineManagement on Windows PowerShell 5.1. Remediation: run this script on PowerShell 7+ (pwsh), or re-run with -SkipUnifiedAuditLog to bypass the UAL step. Original error: $message"
                )
            }

            # Exchange Online's auth flow can throw a timeout even after the
            # browser-side sign-in succeeded. Check Get-ConnectionInformation:
            # if an active connection exists, treat this as success.
            $existingConnection = $null
            try { $existingConnection = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1 } catch { $existingConnection = $null }
            if ($existingConnection) {
                Write-Log -Message ("Exchange Online sign-in raised '{0}' but an active connection is present (account: {1}). Continuing." -f $message, $existingConnection.UserPrincipalName)
                return
            }

            $isTransient = ($message -match 'timed out|authorization_pending|expired_token|inactivity')
            if ($attempt -lt $MaxAttempts -and $isTransient) {
                Write-Log -Level 'WARN' -Message ("Exchange Online sign-in attempt {0} failed: {1}. Retrying with a fresh device code." -f $attempt, $message)
                continue
            }
            throw
        }
    }
}

function Get-ActiveUsersFromUnifiedAudit {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Days,
        [Parameter(Mandatory = $true)]
        [string[]]$Operations,
        [Parameter(Mandatory = $true)]
        [int]$ResultSize,
        [string]$UserPrincipalName,
        [switch]$PickSignInAccount
    )

    if (-not $IsWindows) {
        Write-Log -Level 'WARN' -Message 'Skipping Unified Audit Log step: Search-UnifiedAuditLog / Connect-IPPSSession require Exchange Online PowerShell on Windows. The Graph user-detail report is being used instead.'
        return @()
    }

    # Pre-emptive skip on Windows PowerShell 5.1: once Microsoft.Graph has
    # loaded its (older) Microsoft.Identity.Client into the AppDomain, any
    # ExchangeOnlineManagement sign-in attempt in the same session fails
    # with a 'Method not found: BrokerExtension.WithBroker(BrokerOptions)'
    # MissingMethodException. The only workarounds are PowerShell 7+ or a
    # separate process for EXO. Skip cleanly with remediation instead of
    # asking the user to wait through a doomed device-code prompt.
    if ($PSVersionTable.PSEdition -eq 'Desktop' -and (Get-Module -Name Microsoft.Graph.Authentication)) {
        $skipReason = 'Windows PowerShell 5.1 cannot load Microsoft.Graph and ExchangeOnlineManagement in the same session (Microsoft.Identity.Client assembly mismatch: BrokerExtension.WithBroker overload not found).'
        Write-Log -Level 'WARN' -Message $skipReason
        Add-SkippedItem -Item 'Unified Audit Log (Search-UnifiedAuditLog)' -Reason $skipReason -Remediation 'Re-run this script on PowerShell 7+ (pwsh), or run with -SkipUnifiedAuditLog to suppress this step and rely on Graph data only.'
        return @()
    }

    if (-not $Operations -or $Operations.Count -eq 0) {
        Write-Log -Level 'WARN' -Message 'Skipping Unified Audit Log step: -AuditOperations is empty.'
        return @()
    }

    $signInAccount = $null
    if ($UserPrincipalName) {
        $signInAccount = $UserPrincipalName.Trim()
        if ($signInAccount) {
            Write-Log -Message "Unified Audit Log sign-in account hint: $signInAccount"
        }
    }
    if (-not $signInAccount -and -not $PickSignInAccount) {
        try {
            $signInAccount = (Get-MgContext).Account
        } catch {
            $signInAccount = $null
        }
    }
    if ($PickSignInAccount -and -not $signInAccount) {
        Write-Log -Message '-PickSignInAccount requested: Unified Audit Log child session will prompt for explicit account sign-in (no Graph account hint fallback).'
    }

    $startDate = (Get-Date).ToUniversalTime().AddDays(-$Days)
    $endDate = (Get-Date).ToUniversalTime()

    try {
        Write-Log -Message "Querying Unified Audit Log for the last $Days days. Operations: $($Operations -join ', '). ResultSize: $ResultSize"
        $outputJsonPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.json" -f ([guid]::NewGuid().ToString('N')))
        $errorLogPath = Join-Path ([System.IO.Path]::GetTempPath()) ("agent365-ual-{0}.err.txt" -f ([guid]::NewGuid().ToString('N')))
        $ualUseDeviceCode = $UseDeviceCode.IsPresent
        $records = Invoke-Agent365UnifiedAuditLogQuery `
            -UserPrincipalName $signInAccount `
            -Days $Days `
            -Operations $Operations `
            -ResultSize $ResultSize `
            -OutputJsonPath $outputJsonPath `
            -ErrorLogPath $errorLogPath `
            -UseDeviceCode:$ualUseDeviceCode `
            -PickSignInAccount:$PickSignInAccount
    } catch {
        $message = $_.Exception.Message
        if ($message -match 'Audit Logs or View-Only Audit Logs role|required for Search-UnifiedAuditLog|does not appear to have an audit-log role assignment') {
            Write-Log -Level 'WARN' -Message $message
            Add-SkippedItem -Item 'Unified Audit Log (Search-UnifiedAuditLog)' -Reason $message -Remediation 'Assign the Audit Logs or View-Only Audit Logs role in Microsoft Purview, then re-run the script.'
            return @()
        }
        $ualImportFailure = (($message -match 'Search-UnifiedAuditLog') -and ($message -match 'import')) -or ($message -match 'UseRPSSession recovery both failed')
        if ($ualImportFailure -or $message -match 'Error Acquiring Token') {
            $skipReason = "Unified Audit Log step could not load Search-UnifiedAuditLog in the isolated Security & Compliance session. $message"
            Write-Log -Level 'WARN' -Message $skipReason
            Add-SkippedItem -Item 'Unified Audit Log (Search-UnifiedAuditLog)' -Reason $skipReason -Remediation 'Update ExchangeOnlineManagement to the latest version, then re-run with -UseDeviceCode. If this tenant/session still cannot import Search-UnifiedAuditLog, run with -SkipUnifiedAuditLog to complete from Graph-only data.'
            return @()
        }

        Write-Log -Level 'ERROR' -Message "Unified Audit Log query failed: $message"
        throw "Failed to query Unified Audit Log with operations [$($Operations -join ', ')]. Error: $message"
    }

    if (-not $records) {
        return @()
    }

    $byUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $records) {
        $upn = $null
        try { $upn = [string]$r.UserIds } catch { $upn = $null }
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        if (-not $byUpn.ContainsKey($upn)) {
            $byUpn[$upn] = [PSCustomObject]@{
                UserPrincipalName = $upn
                DisplayName       = $null
                LastActivityDate  = $null
                Source            = 'UAL'
            }
        }
    }

    Write-VerboseLog -Message "Unified Audit Log records retrieved: $($records.Count). Unique active users: $($byUpn.Count)."

    return @($byUpn.Values)
}

function ConvertFrom-CopilotSummaryCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvText,
        [Parameter(Mandatory = $true)]
        [string]$Period
    )

    if ([string]::IsNullOrWhiteSpace($CsvText)) { return $null }

    $rows = $CsvText | ConvertFrom-Csv
    if (-not $rows) { return $null }
    $row = @($rows) | Select-Object -First 1
    if (-not $row) { return $null }

    $toInt = {
        param($v)
        if ($null -eq $v) { return 0 }
        $t = [string]$v
        if ([string]::IsNullOrWhiteSpace($t)) { return 0 }
        $n = 0
        if ([int]::TryParse($t, [ref]$n)) { return $n }
        return 0
    }

    $adoption = [PSCustomObject]@{
        reportPeriod               = [int]($Period.TrimStart('D'))
        anyAppEnabledUsers         = & $toInt $row.'Any App Enabled Users'
        anyAppActiveUsers          = & $toInt $row.'Any App Active Users'
        microsoftTeamsEnabledUsers = & $toInt $row.'Microsoft Teams Enabled Users'
        microsoftTeamsActiveUsers  = & $toInt $row.'Microsoft Teams Active Users'
        wordEnabledUsers           = & $toInt $row.'Word Enabled Users'
        wordActiveUsers            = & $toInt $row.'Word Active Users'
        powerPointEnabledUsers     = & $toInt $row.'PowerPoint Enabled Users'
        powerPointActiveUsers      = & $toInt $row.'PowerPoint Active Users'
        outlookEnabledUsers        = & $toInt $row.'Outlook Enabled Users'
        outlookActiveUsers         = & $toInt $row.'Outlook Active Users'
        excelEnabledUsers          = & $toInt $row.'Excel Enabled Users'
        excelActiveUsers           = & $toInt $row.'Excel Active Users'
        oneNoteEnabledUsers        = & $toInt $row.'OneNote Enabled Users'
        oneNoteActiveUsers         = & $toInt $row.'OneNote Active Users'
        loopEnabledUsers           = & $toInt $row.'Loop Enabled Users'
        loopActiveUsers            = & $toInt $row.'Loop Active Users'
        copilotChatEnabledUsers    = & $toInt $row.'Copilot Chat Enabled Users'
        copilotChatActiveUsers     = & $toInt $row.'Copilot Chat Active Users'
    }

    return [PSCustomObject]@{
        value = @(
            [PSCustomObject]@{
                reportRefreshDate = [string]$row.'Report Refresh Date'
                adoptionByProduct = @($adoption)
            }
        )
    }
}

function Get-CopilotSummaryMetrics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Period
    )

    $base = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')"

    # Try JSON-formatted variants first.
    $jsonUris = @(
        "$base`?`$format=application/json",
        "$base`?`$format=json",
        $base
    )

    foreach ($u in $jsonUris) {
        try {
            Write-VerboseLog -Message "Trying Copilot summary URI: $u"
            $resp = Invoke-MgGraphRequest -Method GET -Uri $u
            if ($resp -is [string]) {
                # Likely CSV returned as raw text - try to parse it.
                $parsed = ConvertFrom-CopilotSummaryCsv -CsvText $resp -Period $Period
                if ($parsed) { return $parsed }
            } elseif ($resp -and $resp.value) {
                return $resp
            }
        } catch {
            Write-VerboseLog -Message "Copilot summary attempt failed for '$u': $($_.Exception.Message)"
        }
    }

    # Final fallback: ask explicitly for CSV. The Graph CSV response comes back
    # as an octet-stream attachment that Invoke-MgGraphRequest will not return
    # inline, so write it to a temp file and read it back.
    $tmpCsv = $null
    try {
        $csvUri = "$base`?`$format=text/csv"
        $tmpCsv = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            "copilot-summary-$([guid]::NewGuid()).csv"
        )
        Write-VerboseLog -Message "Falling back to CSV via temp file: $csvUri -> $tmpCsv"
        Invoke-MgGraphRequest -Method GET -Uri $csvUri -OutputFilePath $tmpCsv | Out-Null
        if (-not (Test-Path -LiteralPath $tmpCsv)) {
            throw "Graph did not write any CSV output to '$tmpCsv'."
        }
        $csvText = Get-Content -LiteralPath $tmpCsv -Raw
        $parsed = ConvertFrom-CopilotSummaryCsv -CsvText $csvText -Period $Period
        if ($parsed) { return $parsed }
    } catch {
        throw "Failed to retrieve Copilot summary metrics from Graph. Last error: $($_.Exception.Message)"
    } finally {
        if ($tmpCsv -and (Test-Path -LiteralPath $tmpCsv)) {
            Remove-Item -LiteralPath $tmpCsv -Force -ErrorAction SilentlyContinue
        }
    }

    throw 'No data returned from getMicrosoft365CopilotUserCountSummary (no JSON or CSV variant succeeded).'
}

function ConvertFrom-CopilotUserDetailCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvText
    )

    if ([string]::IsNullOrWhiteSpace($CsvText)) { return @() }

    $rows = $CsvText | ConvertFrom-Csv
    if (-not $rows) { return @() }

    $records = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($rows)) {
        $rec = [PSCustomObject]@{
            reportRefreshDate                    = [string]$row.'Report Refresh Date'
            userPrincipalName                    = [string]$row.'User Principal Name'
            displayName                          = [string]$row.'Display Name'
            lastActivityDate                     = [string]$row.'Last Activity Date'
            microsoftTeamsCopilotLastActivityDate = [string]$row.'Microsoft Teams Copilot Last Activity Date'
            wordCopilotLastActivityDate           = [string]$row.'Word Copilot Last Activity Date'
            excelCopilotLastActivityDate          = [string]$row.'Excel Copilot Last Activity Date'
            powerPointCopilotLastActivityDate     = [string]$row.'PowerPoint Copilot Last Activity Date'
            outlookCopilotLastActivityDate        = [string]$row.'Outlook Copilot Last Activity Date'
            oneNoteCopilotLastActivityDate        = [string]$row.'OneNote Copilot Last Activity Date'
            loopCopilotLastActivityDate           = [string]$row.'Loop Copilot Last Activity Date'
            copilotChatLastActivityDate           = [string]$row.'Copilot Chat Last Activity Date'
        }
        $records.Add($rec)
    }

    return $records.ToArray()
}

function Get-CopilotUserDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Period
    )

    # The Microsoft 365 Copilot user-detail report is published on the beta
    # surface in two parallel namespaces:
    #   1. /beta/reports/...                (reportRoot, original publication)
    #   2. /beta/copilot/reports/...        (copilotReportRoot, newer)
    # Tenants vary in which one is reachable, so we try both before giving up.
    # Both expose the same JSON/CSV format switching used by the summary report.
    $bases = @(
        "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')",
        "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')"
    )

    $attemptErrors = New-Object System.Collections.Generic.List[string]

    foreach ($base in $bases) {
        $jsonUris = @(
            "$base`?`$format=application/json",
            "$base`?`$format=json",
            $base
        )

        foreach ($u in $jsonUris) {
            try {
                Write-VerboseLog -Message "Trying Copilot user-detail URI: $u"
                $records = New-Object System.Collections.Generic.List[object]
                $next = $u
                $page = 0
                while ($next) {
                    $page++
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $next
                    if ($resp -is [string]) {
                        # JSON path returned raw text - probably CSV. Parse it and stop paging.
                        $parsed = ConvertFrom-CopilotUserDetailCsv -CsvText $resp
                        foreach ($p in @($parsed)) { $records.Add($p) }
                        $next = $null
                        break
                    }
                    if ($resp -and $resp.value) {
                        foreach ($v in @($resp.value)) { $records.Add($v) }
                    }
                    $next = $null
                    if ($resp -and $resp.'@odata.nextLink') {
                        $next = [string]$resp.'@odata.nextLink'
                        Write-VerboseLog -Message "Following nextLink (page $page)"
                    }
                }
                if ($records.Count -gt 0) {
                    Write-VerboseLog -Message "Copilot user-detail rows retrieved (JSON path): $($records.Count)"
                    return $records.ToArray()
                }
            } catch {
                $detail = $_.Exception.Message
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $detail = "$detail | body: $($_.ErrorDetails.Message)"
                }
                Write-VerboseLog -Message "Copilot user-detail attempt failed for '$u': $detail"
                $attemptErrors.Add("GET $u -> $detail")
            }
        }

        # CSV fallback via temp file (mirrors Get-CopilotSummaryMetrics).
        $tmpCsv = $null
        $csvUri = "$base`?`$format=text/csv"
        try {
            $tmpCsv = [System.IO.Path]::Combine(
                [System.IO.Path]::GetTempPath(),
                "copilot-user-detail-$([guid]::NewGuid()).csv"
            )
            Write-VerboseLog -Message "Trying user-detail CSV via temp file: $csvUri -> $tmpCsv"
            Invoke-MgGraphRequest -Method GET -Uri $csvUri -OutputFilePath $tmpCsv | Out-Null
            if (-not (Test-Path -LiteralPath $tmpCsv)) {
                throw "Graph did not write any CSV output to '$tmpCsv'."
            }
            $csvText = Get-Content -LiteralPath $tmpCsv -Raw
            $records = ConvertFrom-CopilotUserDetailCsv -CsvText $csvText
            Write-VerboseLog -Message "Copilot user-detail rows retrieved (CSV path): $($records.Count)"
            return @($records)
        } catch {
            $detail = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $detail = "$detail | body: $($_.ErrorDetails.Message)"
            }
            Write-VerboseLog -Message "Copilot user-detail CSV attempt failed for '$csvUri': $detail"
            $attemptErrors.Add("GET $csvUri -> $detail")
        } finally {
            if ($tmpCsv -and (Test-Path -LiteralPath $tmpCsv)) {
                Remove-Item -LiteralPath $tmpCsv -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # Every attempt failed. Build a comprehensive error message with hints.
    $allErrors = ($attemptErrors -join "`n  - ")
    $hint = ''
    if ($allErrors -match 'Forbidden|\b403\b') {
        $hint = @'

This is almost certainly a permissions issue. The Microsoft 365 Copilot
user-detail report requires the signed-in account to hold one of these
Microsoft Entra admin roles in addition to the Reports.Read.All Graph scope:
  - Global Administrator (Company Administrator)
  - AI Administrator
  - Reports Reader  (least-privilege option)
  - Exchange / SharePoint / Teams Service / Teams Communications Admin
  - Lync Administrator

Holding only Global Reader, Usage Summary Reports Reader, or a non-admin
account is NOT sufficient for the per-user detail report (the summary report
has lower requirements, which is why it succeeded earlier in this run).

Assign Reports Reader (or a higher admin role) to your sign-in account,
then run Disconnect-MgGraph and re-run this script to refresh the token.
'@
    } elseif ($allErrors -match 'BadRequest|\b400\b') {
        $hint = @'

400 BadRequest from every attempt usually means the report endpoint is not
exposed on this tenant's API surface. Verify Microsoft 365 Copilot is
enabled for the tenant and that you signed in to the correct tenant.
'@
    } elseif ($allErrors -match 'NotFound|\b404\b') {
        $hint = @'

404 NotFound from every attempt means the report function does not exist on
this tenant's API surface. The endpoint may have been renamed or the
Microsoft 365 Copilot service is not provisioned in this tenant.
'@
    }

    throw "Failed to retrieve Copilot user detail from Graph after $($attemptErrors.Count) attempt(s).`nAttempts:`n  - $allErrors$hint"
}

function ConvertTo-Agent365ActiveUserCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$UserDetailRecords
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    if (-not $UserDetailRecords) { return $candidates.ToArray() }

    foreach ($r in $UserDetailRecords) {
        if (-not $r) { continue }
        $upn = $null
        try { $upn = [string]$r.userPrincipalName } catch { $upn = $null }
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        $last = $null
        try { $last = [string]$r.lastActivityDate } catch { $last = $null }

        # Only count users who have an actual lastActivityDate inside the
        # requested window. A blank lastActivityDate means "enabled but never
        # active during this period".
        if ([string]::IsNullOrWhiteSpace($last)) { continue }

        $displayName = $null
        try { $displayName = [string]$r.displayName } catch { $displayName = $null }

        $candidates.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            DisplayName       = $displayName
            LastActivityDate  = $last
            Source            = 'Graph'
        })
    }

    return $candidates.ToArray()
}

function Merge-Agent365ActiveUserCandidates {
    [CmdletBinding()]
    param(
        [object[]]$GraphCandidates,
        [object[]]$AuditCandidates
    )

    $byKey = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($g in @($GraphCandidates)) {
        if (-not $g) { continue }
        $upn = [string]$g.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $byKey[$upn] = [PSCustomObject]@{
            UserPrincipalName = $upn
            DisplayName       = $g.DisplayName
            LastActivityDate  = $g.LastActivityDate
            Source            = 'Graph'
        }
    }

    foreach ($a in @($AuditCandidates)) {
        if (-not $a) { continue }
        $upn = if ($a -is [string]) { $a } else { [string]$a.UserPrincipalName }
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        if ($byKey.ContainsKey($upn)) {
            $existing = $byKey[$upn]
            $existing.Source = 'Graph+UAL'
            if ([string]::IsNullOrWhiteSpace([string]$existing.DisplayName) -and -not ($a -is [string]) -and -not [string]::IsNullOrWhiteSpace([string]$a.DisplayName)) {
                $existing.DisplayName = [string]$a.DisplayName
            }
        } else {
            $displayName = if ($a -is [string]) { $null } else { [string]$a.DisplayName }
            $byKey[$upn] = [PSCustomObject]@{
                UserPrincipalName = $upn
                DisplayName       = $displayName
                LastActivityDate  = $null
                Source            = 'UAL'
            }
        }
    }

    return @($byKey.Values)
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
        [AllowEmptyCollection()]
        [object[]]$ActiveUsers,
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

    foreach ($candidate in $ActiveUsers) {
        $index++
        if (-not $NoProgress) {
            $pct = if ($total -eq 0) { 100 } else { [math]::Round(($index / $total) * 100, 0) }
            Write-Progress -Id 2 -Activity 'Classifying active users by Agent 365 license' -Status "Processing $index of $total" -PercentComplete $pct
        }

        if ($candidate -is [string]) {
            $upn = $candidate
            $candidateDisplayName = $null
            $candidateLastActivity = $null
            $candidateSource = 'UAL'
        } else {
            $upn = [string]$candidate.UserPrincipalName
            $candidateDisplayName = try { [string]$candidate.DisplayName } catch { $null }
            $candidateLastActivity = try { [string]$candidate.LastActivityDate } catch { $null }
            $candidateSource = try { [string]$candidate.Source } catch { 'Graph' }
            if ([string]::IsNullOrWhiteSpace($candidateSource)) { $candidateSource = 'Graph' }
        }

        if ([string]::IsNullOrWhiteSpace($upn)) {
            Write-VerboseLog -Message 'Skipping candidate with empty UserPrincipalName.'
            continue
        }

        try {
            $encodedUserId = [uri]::EscapeDataString($upn)
            $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserId?`$select=id,userPrincipalName,displayName,assignedLicenses"
        } catch {
            Write-Log -Level 'WARN' -Message "Skipping user '$upn' due to lookup failure."
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

        $resolvedDisplayName = if ([string]::IsNullOrWhiteSpace([string]$user.displayName)) { $candidateDisplayName } else { [string]$user.displayName }

        $record = [PSCustomObject]@{
            UserPrincipalName  = [string]$user.userPrincipalName
            DisplayName        = $resolvedDisplayName
            ObjectId           = [string]$user.id
            IsAgent365Licensed = $hasCopilot
            LastActivityDate   = $candidateLastActivity
            Source             = $candidateSource
        }

        if ($hasCopilot) {
            $licensed.Add($record)
            Write-VerboseLog -Message "Licensed user: $($record.UserPrincipalName) ($($record.ObjectId)) [$($record.Source)]"
        } else {
            $unlicensed.Add($record)
            Write-VerboseLog -Message "Unlicensed user: $($record.UserPrincipalName) ($($record.ObjectId)) [$($record.Source)]"
        }
    }

    if (-not $NoProgress) {
        Write-Progress -Id 2 -Activity 'Classifying active users by Agent 365 license' -Completed
    }

    Write-Log -Message "Classified users. Licensed: $($licensed.Count), Unlicensed: $($unlicensed.Count)"

    # Convert List[object] to plain arrays before returning. PowerShell 7.6.x
    # has a bug where @($psObj.SomeProperty) throws
    # "OperationStopped: Argument types do not match" when the property holds
    # an empty System.Collections.Generic.List[object]. Returning arrays
    # sidesteps the bug entirely for every caller.
    return [PSCustomObject]@{
        Licensed   = $licensed.ToArray()
        Unlicensed = $unlicensed.ToArray()
    }
}

function Convert-UsersToHtmlRows {
        param(
                [Parameter(Mandatory = $true)]
                [AllowEmptyCollection()]
                [object[]]$Users
        )

        if (-not $Users -or $Users.Count -eq 0) {
                return '<tr><td colspan="5">No users found.</td></tr>'
        }

        $rows = foreach ($u in ($Users | Sort-Object UserPrincipalName)) {
                $upn = [System.Net.WebUtility]::HtmlEncode([string]$u.UserPrincipalName)
                $name = [System.Net.WebUtility]::HtmlEncode([string]$u.DisplayName)
                $oid = [System.Net.WebUtility]::HtmlEncode([string]$u.ObjectId)
                $lastRaw = $null
                try { $lastRaw = [string]$u.LastActivityDate } catch { $lastRaw = $null }
                if ([string]::IsNullOrWhiteSpace($lastRaw)) { $lastRaw = '' }
                $last = [System.Net.WebUtility]::HtmlEncode($lastRaw)
                $srcRaw = $null
                try { $srcRaw = [string]$u.Source } catch { $srcRaw = $null }
                if ([string]::IsNullOrWhiteSpace($srcRaw)) { $srcRaw = 'Graph' }
                $src = [System.Net.WebUtility]::HtmlEncode($srcRaw)
                "<tr><td>$upn</td><td>$name</td><td>$oid</td><td>$last</td><td>$src</td></tr>"
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
        [AllowEmptyCollection()]
        [object[]]$LicensedUsers,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$UnlicensedUsers,
        [Parameter(Mandatory = $true)]
        [string[]]$MatchedSkuPartNumbers,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$TenantInfo,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$DataSources
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

    # Human-readable period description (e.g. "trailing 30 days (2026-04-18
    # -> 2026-05-17)") so readers don't have to know what "D30" means.
    $periodDescription = Get-PeriodDescription -PeriodValue $Summary.Period -ReportRefreshDate $Summary.ReportRefreshDate
    $periodHtml = & $encode "$($Summary.Period) - $periodDescription" $Summary.Period

    # Build an in-report warning banner when the per-user Graph call was
    # skipped or returned 403 Forbidden. Without this the user might not
    # realise the empty Licensed/Unlicensed tables are due to permissions
    # rather than zero activity.
    $userDetailBannerHtml = ''
    if ($DataSources.PSObject.Properties.Name -contains 'GraphUserDetailUnavailable' -and $DataSources.GraphUserDetailUnavailable) {
        $reasonHtml = & $encode $DataSources.GraphUserDetailReason 'The per-user Copilot detail call returned 403 Forbidden.'
        $userDetailBannerHtml = @"
            <div class="alert warning" role="alert">
                <h3>Per-user Copilot activity unavailable</h3>
                <p>$reasonHtml</p>
                <p>This report shows tenant-level summary metrics only. The <strong>Licensed</strong> and <strong>Unlicensed</strong> active-user tables are empty because Microsoft Graph did not return per-user records for this account.</p>
                <p><strong>To populate the per-user tables:</strong></p>
                <ul>
                    <li>Assign <code>Reports Reader</code> (least privilege, read-only) or a higher admin role (Global Admin, AI Admin, Exchange / SharePoint / Teams admin) to the signed-in account in the Microsoft Entra admin center.</li>
                    <li>Run <code>Disconnect-MgGraph</code> and re-run this script to refresh the token.</li>
                </ul>
            </div>
"@
    }

    # Build the [SKIPPED] tokens for the top stat tiles so empty Licensed /
    # Unlicensed values don't masquerade as "zero activity" when the real
    # reason is that we never received per-user data.
    $userDetailSkipped     = ($DataSources.PSObject.Properties.Name -contains 'GraphUserDetailUnavailable') -and [bool]$DataSources.GraphUserDetailUnavailable
    $licensedTotalDisplay   = & $encode (Format-CountOrSkipped -Count $LicensedUsers.Count   -WasSkipped $userDetailSkipped) '0'
    $unlicensedTotalDisplay = & $encode (Format-CountOrSkipped -Count $UnlicensedUsers.Count -WasSkipped $userDetailSkipped) '0'

    # Render a per-step "Skipped steps" section inside the Notes block so the
    # HTML report explains WHY any value rendered as [SKIPPED] above. Kept
    # empty when nothing was skipped.
    $skippedNotesHtml = ''
    if (($DataSources.PSObject.Properties.Name -contains 'SkippedItems') -and @($DataSources.SkippedItems).Count -gt 0) {
        $skippedListItems = foreach ($s in $DataSources.SkippedItems) {
            $itemHtml        = & $encode $s.Item        'Unknown step'
            $reasonHtml      = & $encode $s.Reason      'Not recorded'
            $remediationHtml = ''
            if ($s.PSObject.Properties.Name -contains 'Remediation' -and -not [string]::IsNullOrWhiteSpace([string]$s.Remediation)) {
                $remediationHtml = '<br/><em>Remediation:</em> ' + (& $encode $s.Remediation '')
            }
            "                        <li><strong>$itemHtml</strong> &mdash; $reasonHtml$remediationHtml</li>"
        }
        $skippedNotesHtml = @"
            <div class="notes" id="skipped-steps">
                <h2>Skipped steps</h2>
                <p>The following report steps were skipped. Values above shown as <code>[SKIPPED]</code> are a consequence of these skips, <em>not</em> a true zero.</p>
                <ul>
$($skippedListItems -join [Environment]::NewLine)
                </ul>
            </div>
"@
    }

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
        .alert { border-radius: 10px; padding: 12px 16px; margin: 0 0 16px 0; border: 1px solid; }
        .alert.warning { background: #fff8e1; border-color: #f0c97a; color: #6b4f00; }
        .alert h3 { margin: 0 0 6px 0; font-size: 15px; }
        .alert p { margin: 0 0 6px 0; }
        .alert ul { margin: 6px 0 0 18px; padding: 0; }
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
            <div class="meta">Generated UTC: $generatedUtc &middot; Period: $periodHtml &middot; Report Refresh Date: $($Summary.ReportRefreshDate)</div>
$userDetailBannerHtml
            <div class="stats">
                <div class="stat"><div class="k">Agent 365 Active Users</div><div class="v">$($Summary.ActiveUsers)</div></div>
                <div class="stat"><div class="k">Enabled Users</div><div class="v">$($Summary.EnabledUsers)</div></div>
                <div class="stat"><div class="k">Copilot Chat Active Users</div><div class="v">$($Summary.CopilotChatActiveUsers)</div></div>
                <div class="stat"><div class="k">Active Licensed Users Total</div><div class="v">$licensedTotalDisplay</div></div>
                <div class="stat"><div class="k">Active Unlicensed Users Total</div><div class="v">$unlicensedTotalDisplay</div></div>
            </div>

            <div class="tabs">
                <button class="tab-btn active" data-tab="licensed">Licensed Active Users</button>
                <button class="tab-btn" data-tab="unlicensed">Unlicensed Active Users</button>
            </div>

            <div id="licensed" class="tab active">
                <table>
                    <thead>
                        <tr><th>UserPrincipalName</th><th>DisplayName</th><th>ObjectId</th><th>Last Activity Date</th><th>Source</th></tr>
                    </thead>
                    <tbody>
$licensedRows
                    </tbody>
                </table>
            </div>

            <div id="unlicensed" class="tab">
                <table>
                    <thead>
                        <tr><th>UserPrincipalName</th><th>DisplayName</th><th>ObjectId</th><th>Last Activity Date</th><th>Source</th></tr>
                    </thead>
                    <tbody>
$unlicensedRows
                    </tbody>
                </table>
            </div>

            <div class="notes">
                <h2>Notes, Assumptions, and Caveats</h2>
                <ul>
                    <li><strong>Summary metrics</strong> (top tiles) come from Microsoft Graph <code>getMicrosoft365CopilotUserCountSummary</code> for the period <code>$($Summary.Period)</code>.</li>
                    <li><strong>Per-user list</strong> comes from Microsoft Graph <code>getMicrosoft365CopilotUsageUserDetail</code>. This endpoint reports users with assigned Microsoft 365 Copilot / Agent 365 licenses; the <em>Source</em> column shows <code>Graph</code> for these rows.$($DataSources.UalNoteHtml)</li>
                    <li>How Agent 365 license is determined: the script matches tenant subscribed SKUs using exact SKU part numbers ($($CopilotSkuPartNumbers -join ', ')) plus wildcard patterns ($($CopilotSkuPartNumberPatterns -join ', ')), then maps these to skuId GUIDs. Each active user is marked licensed only if any users/{id}.assignedLicenses.skuId matches one of those GUIDs.</li>
                    <li>SKUs matched in this tenant for this run: $($MatchedSkuPartNumbers -join ', ').</li>
                    <li><strong>Active users (per the Agent 365 admin center)</strong> — the per-agent breakdown and “trending agents” view shown in <em>Microsoft 365 admin center &rarr; Agent 365</em> are not yet exposed via a public Graph API. This report reflects the closest API-available signal: per-user Microsoft 365 Copilot activity (plus, optionally, audit log activity).</li>
                    <li>Report accuracy depends on Graph reporting latency (typically 24–48 hours), license assignment freshness, and your tenant’s Reports.Read.All / User.Read.All permissions.</li>
                </ul>
            </div>
$skippedNotesHtml
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
                [AllowEmptyCollection()]
                [object[]]$Users
        )

        # Use the unary comma operator on the local variable so the function
        # always returns an actual array (even empty) instead of letting
        # PowerShell unwrap an empty pipeline to $null. Without this,
        # $licensedTableUsers.Count under Set-StrictMode throws
        # "The property 'Count' cannot be found on this object".
        $projected = @($Users | Select-Object UserPrincipalName, DisplayName, ObjectId, LastActivityDate, Source)
        return ,$projected
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
    Connect-Agent365Graph -Scopes $scopes -UseDeviceCode:$UseDeviceCode
} elseif (-not ($ctx.Scopes -contains 'Reports.Read.All')) {
    Write-Log -Message 'Reconnecting to Microsoft Graph to include required scopes.'
    Connect-Agent365Graph -Scopes $scopes -UseDeviceCode:$UseDeviceCode
}
$script:StepId++

Write-Log -Message 'Resolving tenant context for report header.'
$tenantInfo = Get-TenantContextInfo

# Pre-flight: confirm the signed-in account holds a directory role that
# qualifies for the per-user Copilot usage report. Reports.Read.All Graph
# scope alone is not enough -- the user-detail endpoint also requires one
# of: Global Administrator, AI Administrator, Reports Reader (least
# privilege), Exchange / SharePoint / Teams Service / Teams Communications
# / Skype for Business (Lync) Administrator. This is a WARN-only check
# (custom roles may grant access via different mechanisms), but it gives
# the user immediate, actionable guidance instead of waiting for a 403.
if ($SkipRoleCheck) {
    Write-Log -Message 'Skipping Microsoft Entra directory-role pre-check (-SkipRoleCheck specified).'
} else {
    Write-Log -Message 'Checking signed-in account for a directory role that qualifies for the per-user Copilot usage report.'
    $roleCheck = Test-Agent365ReportsAccess
    if (-not $roleCheck.Checked) {
        $rcUpn = if ($roleCheck.UserPrincipalName) { $roleCheck.UserPrincipalName } else { '<unknown>' }
        Write-Log -Level 'WARN' -Message ('Could not enumerate directory roles for the signed-in account ({0}). Skipping role pre-check; the per-user Copilot report call may fail with 403 Forbidden.' -f $rcUpn)
    } elseif ($roleCheck.HasQualifyingRole) {
        $rcUpn = if ($roleCheck.UserPrincipalName) { $roleCheck.UserPrincipalName } else { '<unknown>' }
        Write-Log -Message ('Directory-role pre-check passed for {0}. Full-access roles assigned: {1}.' -f $rcUpn, ($roleCheck.FullAccessRoles -join ', '))
        if ($roleCheck.HasTenantOnlyRole) {
            Write-Log -Message ('  Additionally holds tenant-only role(s): {0} (not needed; full-access role takes precedence).' -f ($roleCheck.TenantOnlyRoles -join ', '))
        }
    } elseif ($roleCheck.HasTenantOnlyRole) {
        # Role IS documented as authorized, but per
        # https://learn.microsoft.com/graph/reportroot-authorization
        # "Global Reader and Usage Summary Reports Reader roles will only
        # have access to tenant-level data, without visibility into
        # detailed metrics." Newer /copilot/reports/ variants do not
        # include these roles at all.
        $upn = $roleCheck.UserPrincipalName
        if (-not $upn) { $upn = '<unknown>' }
        Write-Log -Level 'WARN' -Message ('Directory-role pre-check: account {0} holds {1}, which is documented as TENANT-LEVEL ONLY for the Copilot usage report.' -f $upn, ($roleCheck.TenantOnlyRoles -join ', '))
        Write-Log -Level 'WARN' -Message '  Microsoft restricts these roles to aggregate data only; per-user records are not returned and the call typically yields 403 Forbidden on /beta/reports and /copilot/reports endpoints.'
        Write-Log -Level 'WARN' -Message ("  Required for per-user detail (any one of): $($roleCheck.QualifyingRoleList -join ', ')")
        Write-Log -Level 'WARN' -Message '  Remediation: in the Microsoft Entra admin center, assign Reports Reader (least privilege, read-only) IN ADDITION to your current role, then run: Disconnect-MgGraph; and re-run this script.'
        Write-Log -Level 'WARN' -Message '  Skipping per-user detail call. The report will still be generated from Graph summary metrics + license data; the Licensed/Unlicensed active-user tables will be empty.'
        $script:userDetailLikelyForbidden = $true
        $script:userDetailSkipReason = ('Signed-in account ({0}) holds only tenant-level role(s): {1}. Microsoft restricts these roles to aggregate data only.' -f $upn, ($roleCheck.TenantOnlyRoles -join ', '))
        Add-SkippedItem -Item 'Per-user Copilot activity (Graph getMicrosoft365CopilotUsageUserDetail)' -Reason $script:userDetailSkipReason -Remediation 'In the Microsoft Entra admin center, assign Reports Reader (least privilege, read-only) IN ADDITION to the current role, then run Disconnect-MgGraph and re-run this script.'
    } else {
        $upn = $roleCheck.UserPrincipalName
        if (-not $upn) { $upn = '<unknown>' }
        $assignedText = if (@($roleCheck.AssignedRoles).Count -gt 0) { $roleCheck.AssignedRoles -join ', ' } else { '<none>' }
        Write-Log -Level 'WARN' -Message ('Directory-role pre-check FAILED. Signed-in account {0} does not hold any role that qualifies for getMicrosoft365CopilotUsageUserDetail.' -f $upn)
        Write-Log -Level 'WARN' -Message ("  Assigned directory roles: $assignedText")
        Write-Log -Level 'WARN' -Message ("  Required (any one of): $($roleCheck.QualifyingRoleList -join ', ')")
        Write-Log -Level 'WARN' -Message ("  Documented but tenant-only (insufficient for per-user detail): $($roleCheck.TenantOnlyRoleList -join ', ')")
        Write-Log -Level 'WARN' -Message '  Reports.Read.All Graph scope alone is NOT sufficient. The per-user Copilot report will likely return 403 Forbidden.'
        Write-Log -Level 'WARN' -Message '  Remediation: in the Microsoft Entra admin center, assign Reports Reader (least privilege) to this account, then run: Disconnect-MgGraph; and re-run this script.'
        Write-Log -Level 'WARN' -Message '  Skipping per-user detail call. The report will still be generated from Graph summary metrics + license data; the Licensed/Unlicensed active-user tables will be empty.'
        Write-Log -Level 'WARN' -Message '  If access is granted via a custom directory role the check does not recognise, re-run with -SkipRoleCheck to suppress this warning and attempt the call anyway.'
        $script:userDetailLikelyForbidden = $true
        $script:userDetailSkipReason = ('Signed-in account ({0}) holds no directory role that qualifies for the per-user Copilot usage report. Assigned roles: {1}.' -f $upn, $assignedText)
        Add-SkippedItem -Item 'Per-user Copilot activity (Graph getMicrosoft365CopilotUsageUserDetail)' -Reason $script:userDetailSkipReason -Remediation 'In the Microsoft Entra admin center, assign Reports Reader (least privilege) to this account, then run Disconnect-MgGraph and re-run this script.'
    }
}

# Microsoft's /copilot/reports endpoints currently return a CSV-backed Stream.
# Get-CopilotSummaryMetrics tries JSON-format variants first and falls back to CSV.
Update-StepProgress -Activity 'Agent 365 report' -Status 'Pulling summary metrics from Graph'
Write-Log -Message 'Requesting Copilot summary usage metrics from Graph.'
$response = Get-CopilotSummaryMetrics -Period $Period

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

Update-StepProgress -Activity 'Agent 365 report' -Status 'Collecting active users from Microsoft Graph'
$days = Get-PeriodDays -PeriodValue $Period

# User-detail availability tracking. The role pre-check above may have
# already set $script:userDetailLikelyForbidden when it knows the call will
# fail. We still wrap the actual call in try/catch as a safety net for
# custom-role tenants where the pre-check can't predict success.
$userDetailRecords        = @()
$userDetailUnavailable    = $false
$userDetailFailureReason  = $null

if ($script:userDetailLikelyForbidden) {
    Write-Log -Level 'WARN' -Message 'Skipping per-user Copilot detail call (role pre-check predicted 403 Forbidden). Report will be built from Graph summary metrics and license data only.'
    $userDetailUnavailable   = $true
    $userDetailFailureReason = $script:userDetailSkipReason
} else {
    Write-Log -Message 'Pulling per-user activity from Microsoft Graph (getMicrosoft365CopilotUsageUserDetail).'
    try {
        $userDetailRecords = Get-CopilotUserDetail -Period $Period
    } catch {
        $msg = [string]$_.Exception.Message
        if ($msg -match 'Forbidden|\b403\b|S2SUnauthorized') {
            Write-Log -Level 'WARN' -Message 'Per-user Copilot detail returned 403 Forbidden despite role pre-check passing (or pre-check skipped). Continuing with Graph summary + license data only.'
            Write-Log -Level 'WARN' -Message '  Assign Reports Reader (or higher admin role) to this account, then run Disconnect-MgGraph and re-run.'
            $userDetailRecords       = @()
            $userDetailUnavailable   = $true
            $userDetailFailureReason = '403 Forbidden from Microsoft Graph getMicrosoft365CopilotUsageUserDetail. The signed-in account does not have sufficient permissions to read per-user Copilot activity.'
            Add-SkippedItem -Item 'Per-user Copilot activity (Graph getMicrosoft365CopilotUsageUserDetail)' -Reason $userDetailFailureReason -Remediation 'Assign Reports Reader (or higher admin role such as Global Administrator / AI Administrator) to this account, then run Disconnect-MgGraph and re-run this script.'
        } else {
            throw
        }
    }
}

# Wrap in @(...) so an empty result from the function is preserved as an
# empty array rather than being unwrapped to $null (which would break
# .Count under Set-StrictMode).
$graphCandidates = @(ConvertTo-Agent365ActiveUserCandidate -UserDetailRecords $userDetailRecords)
if ($userDetailUnavailable) {
    Write-Log -Level 'WARN' -Message 'No per-user Graph activity available; Licensed/Unlicensed active-user tables will be empty unless UAL is enabled.'
} else {
    Write-Log -Message "Graph reported $($graphCandidates.Count) users with activity in the last $days day(s)."
}
$script:StepId++

$auditCandidates = @()
$ualAttempted = $false
$ualSkippedReason = $null
if ($IncludeUnifiedAuditLog) {
    $ualAttempted = $true
    if (-not $IsWindows) {
        $ualSkippedReason = 'Search-UnifiedAuditLog requires Exchange Online PowerShell on Windows. Skipped on this platform.'
        Write-Log -Level 'WARN' -Message $ualSkippedReason
        Add-SkippedItem -Item 'Unified Audit Log (CopilotInteraction)' -Reason $ualSkippedReason -Remediation 'Re-run this script on Windows PowerShell 7 with the ExchangeOnlineManagement module installed.'
    } else {
        Update-StepProgress -Activity 'Agent 365 report' -Status 'Collecting active users from Unified Audit Log'
        Write-Log -Message 'Collecting active users from Unified Audit Log.'
        # Wrap in @(...) so an empty/$null return is preserved as an empty
        # array under Set-StrictMode (otherwise .Count below throws
        # PropertyNotFoundStrict).
        $auditCandidates = @(Get-ActiveUsersFromUnifiedAudit -Days $days -Operations $AuditOperations -ResultSize $AuditResultSize -UserPrincipalName $SignInAccount -PickSignInAccount:$PickSignInAccount)
        Write-Log -Message "UAL returned $($auditCandidates.Count) unique users."
    }
} else {
    Write-Log -Message 'Unified Audit Log step skipped (disabled via -IncludeUnifiedAuditLog:$false or -SkipUnifiedAuditLog, or running on a non-Windows host).'
}
$script:StepId++

# Wrap in @(...) so an empty merge result is preserved as an empty array
# rather than being unwrapped to $null under Set-StrictMode.
$mergedCandidates = @(Merge-Agent365ActiveUserCandidates -GraphCandidates $graphCandidates -AuditCandidates $auditCandidates)
$mergedCountMsg = if ($userDetailUnavailable) {
    "Merged candidate count after dedup: {0} (per-user step skipped; only audit/license sources contributed)." -f $mergedCandidates.Count
} else {
    "Merged candidate count after dedup: {0}." -f $mergedCandidates.Count
}
Write-Log -Message $mergedCountMsg

Update-StepProgress -Activity 'Agent 365 report' -Status 'Resolving Agent 365 SKUs'
Write-Log -Message 'Resolving Agent 365 SKU IDs from subscribed SKUs.'
$skuMatchResults = Get-CopilotSkuIds -SkuPartNumbers $CopilotSkuPartNumbers -SkuPartNumberPatterns $CopilotSkuPartNumberPatterns
$copilotSkuIds = @($skuMatchResults.SkuIds)
$matchedSkuPartNumbers = @($skuMatchResults.MatchedSkuPartNumbers)
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Classifying licensed vs unlicensed users'
$classifyMsg = if ($userDetailUnavailable) {
    "Classifying active users by license assignment (count: {0}; per-user step skipped)." -f $mergedCandidates.Count
} else {
    "Classifying active users by license assignment (count: {0})." -f $mergedCandidates.Count
}
Write-Log -Message $classifyMsg
$classification = Get-ActiveUserLicenseClassification -ActiveUsers $mergedCandidates -CopilotSkuIds $copilotSkuIds
$script:StepId++

$licensedActiveUsers = @($classification.Licensed)
$unlicensedActiveUsers = @($classification.Unlicensed)

$licensedTableUsers = Get-UserTableProjection -Users $licensedActiveUsers
$unlicensedTableUsers = Get-UserTableProjection -Users $unlicensedActiveUsers

# Describe which data sources actually contributed to this report so the HTML
# notes block can render an accurate disclaimer.
$ualNoteHtml = if ($IncludeUnifiedAuditLog -and $IsWindows) {
    " Unified Audit Log was also queried for operations [$($AuditOperations -join ', ')]; rows tagged <code>UAL</code> or <code>Graph+UAL</code> were observed there."
} elseif ($IncludeUnifiedAuditLog) {
    ' Unified Audit Log was requested but skipped because Search-UnifiedAuditLog requires Exchange Online PowerShell on Windows.'
} else {
    ' The Unified Audit Log path is disabled by default. Re-run on Windows with <code>-IncludeUnifiedAuditLog</code> to also include unlicensed Copilot Chat activity from audit records, or use <code>-SkipUnifiedAuditLog</code> to force-disable it.'
}

$dataSources = [PSCustomObject]@{
    GraphUserDetail            = -not $userDetailUnavailable
    GraphUserDetailUnavailable = [bool]$userDetailUnavailable
    GraphUserDetailReason      = $userDetailFailureReason
    GraphCandidates            = $graphCandidates.Count
    UalRequested               = [bool]$IncludeUnifiedAuditLog
    UalAttempted               = [bool]$ualAttempted
    UalSkippedReason           = $ualSkippedReason
    UalCandidates              = $auditCandidates.Count
    MergedCandidates           = $mergedCandidates.Count
    AuditOperations            = $AuditOperations
    UalNoteHtml                = $ualNoteHtml
    # PowerShell 7.6.x can throw "Argument types do not match" for @() over
    # Generic.List[object] values in hashtable/object initializers; use ToArray().
    SkippedItems               = $script:SkippedItems.ToArray()
}

if ($ReturnRaw) {
    $reportFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))

    [PSCustomObject]@{
        Summary                       = $result
        ActiveLicensedUsersTotal      = $licensedTableUsers.Count
        ActiveUnlicensedUsersTotal    = $unlicensedTableUsers.Count
        InputAgent365SkuPartNumbers   = $CopilotSkuPartNumbers
        InputAgent365SkuPatterns      = $CopilotSkuPartNumberPatterns
        MatchedAgent365SkuPartNumbers = $matchedSkuPartNumbers
        LicensedActiveUsers           = $licensedTableUsers
        UnlicensedActiveUsers         = $unlicensedTableUsers
        DataSources                   = $dataSources
        ReportPath                    = $reportFullPath
    } | ConvertTo-Json -Depth 8
    Write-Log -Message 'Completed run in ReturnRaw mode.'
    if (-not $NoProgress) {
        Write-Progress -Activity 'Agent 365 report' -Completed
    }
    return
}

Write-Host "Agent 365 Active Users (Period: $Period - $(Get-PeriodDescription -PeriodValue $Period -ReportRefreshDate $result.ReportRefreshDate)): $($result.ActiveUsers)"
Write-Host "Report Refresh Date: $($result.ReportRefreshDate)"
Write-Host "Enabled Users: $($result.EnabledUsers)"
Write-Host "Copilot Chat Active Users: $($result.CopilotChatActiveUsers)"
Write-Host "Agent 365 SKU exact matches configured: $($CopilotSkuPartNumbers -join ', ')"
Write-Host "Agent 365 SKU pattern matches configured: $($CopilotSkuPartNumberPatterns -join ', ')"
Write-Host "Agent 365 SKUs matched in tenant: $($matchedSkuPartNumbers -join ', ')"
Write-Host ("Active Licensed Users Total: {0}" -f (Format-CountOrSkipped -Count $licensedTableUsers.Count -WasSkipped $dataSources.GraphUserDetailUnavailable))
Write-Host ("Active Unlicensed Users Total: {0}" -f (Format-CountOrSkipped -Count $unlicensedTableUsers.Count -WasSkipped $dataSources.GraphUserDetailUnavailable))
if ($dataSources.GraphUserDetailUnavailable) {
    Write-Host '[SKIPPED] Graph user-detail rows: per-user call skipped or returned 403 Forbidden.' -ForegroundColor Yellow
    Write-Host "  Reason: $($dataSources.GraphUserDetailReason)"
    Write-Host '  Report still generated using Graph summary metrics + license data only.'
} else {
    Write-Host "Graph user-detail rows: $($dataSources.GraphCandidates)"
}
if ($dataSources.UalRequested) {
    if ($dataSources.UalSkippedReason) {
        Write-Host '[SKIPPED] Unified Audit Log was requested but could not be queried.' -ForegroundColor Yellow
        Write-Host ("  Reason: {0}" -f $dataSources.UalSkippedReason)
    } else {
        Write-Host "Unified Audit Log unique users: $($dataSources.UalCandidates)"
    }
} elseif (-not $IsWindows) {
    Write-Host '[SKIPPED] Unified Audit Log is unavailable on this platform.' -ForegroundColor Yellow
    Write-Host ("  Reason: Search-UnifiedAuditLog requires Exchange Online PowerShell, which is Windows-only in PowerShell 7. Detected platform: {0}." -f $(if ($IsLinux) { 'Linux' } elseif ($IsMacOS) { 'macOS' } else { 'non-Windows' }))
}
# Note: when UAL was disabled (declined upgrade, -SkipUnifiedAuditLog, etc.)
# the reason is already reported in the Skipped-steps section printed below.

New-Agent365HtmlReport -OutputPath $ReportPath -Summary $result -LicensedUsers $licensedTableUsers -UnlicensedUsers $unlicensedTableUsers -MatchedSkuPartNumbers $matchedSkuPartNumbers -TenantInfo $tenantInfo -DataSources $dataSources
$script:StepId++

Update-StepProgress -Activity 'Agent 365 report' -Status 'Finalizing report output'
# Report path is printed at the very end of the run (browser-launch block),
# so we don't echo it here.

if ($VerboseLog) {
    Write-Host 'Verbose logging: enabled'
}

Write-Host ''
Write-Host 'Licensed Active Users (Agent 365):'
if ($licensedTableUsers.Count -gt 0) {
    $licensedTableUsers | Sort-Object UserPrincipalName | Format-Table -AutoSize UserPrincipalName, DisplayName, ObjectId, LastActivityDate, Source
} elseif ($dataSources.GraphUserDetailUnavailable) {
    Write-Host '[SKIPPED] Per-user data could not be retrieved.' -ForegroundColor Yellow
    if ($dataSources.GraphUserDetailReason) {
        Write-Host ("  Reason: {0}" -f $dataSources.GraphUserDetailReason)
    }
} else {
    Write-Host 'No licensed active users found.'
}

Write-Host ''
Write-Host 'Unlicensed Active Users (Agent 365):'
if ($unlicensedTableUsers.Count -gt 0) {
    $unlicensedTableUsers | Sort-Object UserPrincipalName | Format-Table -AutoSize UserPrincipalName, DisplayName, ObjectId, LastActivityDate, Source
} elseif ($dataSources.GraphUserDetailUnavailable) {
    Write-Host '[SKIPPED] Per-user data could not be retrieved.' -ForegroundColor Yellow
    if ($dataSources.GraphUserDetailReason) {
        Write-Host ("  Reason: {0}" -f $dataSources.GraphUserDetailReason)
    }
} else {
    Write-Host 'No unlicensed active users found.'
}

# Skipped-steps summary: explain WHY any value above shows [SKIPPED] or
# UNAVAILABLE, and what the user can do about it.
if ($script:SkippedItems.Count -gt 0) {
    Write-Host ''
    Write-Host 'Skipped steps (values shown as [SKIPPED] above):' -ForegroundColor Yellow
    foreach ($s in $script:SkippedItems) {
        Write-Host ("  - {0}" -f $s.Item) -ForegroundColor Yellow
        Write-Host ("      Reason     : {0}" -f $s.Reason)
        if ($s.Remediation) {
            Write-Host ("      Remediation: {0}" -f $s.Remediation)
        }
    }
}

$script:StepId++
Update-StepProgress -Activity 'Agent 365 report' -Status 'Complete'
if (-not $NoProgress) {
    Write-Progress -Activity 'Agent 365 report' -Completed
}

# Final summary block: report path + execution log + best-effort browser launch.
# Printed AFTER the Skipped-steps section so the user always sees "where the
# output is" as the last thing on screen.
Write-Host ''
Write-Host ''
if (Open-ReportInDefaultBrowser -Path $script:ReportFullPath) {
    Write-Host "Opened report in default browser: $script:ReportFullPath" -ForegroundColor Green
    Write-Log  -Level SUCCESS -Message "Opened report in default browser: $script:ReportFullPath"
} else {
    Write-Host "Report saved (open manually): $script:ReportFullPath" -ForegroundColor Green
    Write-Log  -Message "Browser auto-launch skipped (headless/non-interactive host). Report path: $script:ReportFullPath"
}
Write-Host "Execution log: $script:LogFullPath" -ForegroundColor Green

Write-Log -Level SUCCESS -Message 'Run completed successfully.'
