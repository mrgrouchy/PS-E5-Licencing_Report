<#
.SYNOPSIS
    Reports on licence assignments recorded in Entra ID audit logs within a given lookback window.

.DESCRIPTION
    Queries auditLogs/directoryAudits for "Add user" / "Change user license" / group-based licence
    events, then enriches each result with UPN, DisplayName, UsageLocation, and ExtensionAttribute1
    via a batched Graph call. Outputs to console and exports a timestamped CSV.

.PARAMETER Hours
    Lookback window in hours. Mutually exclusive with -Days, -StartDate/-EndDate.

.PARAMETER Days
    Lookback window in days. Mutually exclusive with -Hours, -StartDate/-EndDate.

.PARAMETER StartDate
    Start of a custom date range (UTC). Must be paired with -EndDate.

.PARAMETER EndDate
    End of a custom date range (UTC). Must be paired with -StartDate.

.PARAMETER ExportPath
    Full path for the CSV output file.
    Defaults to .\LicenceAssignments_<timestamp>.csv in the working directory.

.PARAMETER IncludeRemoved
    Include licence removal events in addition to assignments.

.EXAMPLE
    # Last 24 hours, interactive session already connected
    .\Get-LicenceAssignmentReport.ps1 -Hours 24

.EXAMPLE
    # Last 7 days
    .\Get-LicenceAssignmentReport.ps1 -Days 7

.EXAMPLE
    # Custom range, include removals
    .\Get-LicenceAssignmentReport.ps1 -StartDate "2026-03-01" -EndDate "2026-03-10" -IncludeRemoved

.NOTES
    Required Graph scopes (delegated or application):
        AuditLog.Read.All
        User.Read.All
        Directory.Read.All

    Recommended RBAC roles (least privilege):
        Reports Reader  +  Directory Reader
        — or —
        Security Reader  (covers both)

    Install the SDK authentication module if not present:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>

[CmdletBinding(DefaultParameterSetName = 'Hours')]
param(
    [Parameter(ParameterSetName = 'Hours')]
    [ValidateRange(1, 720)]
    [int]$Hours = 24,

    [Parameter(ParameterSetName = 'Days')]
    [ValidateRange(1, 30)]
    [int]$Days,

    [Parameter(ParameterSetName = 'Range', Mandatory)]
    [datetime]$StartDate,

    [Parameter(ParameterSetName = 'Range', Mandatory)]
    [datetime]$EndDate,

    [string]$ExportPath,

    [switch]$IncludeRemoved,

    [ValidateSet('ENTERPRISEPREMIUM', 'SPE_E5', 'All')]
    [string]$SkuFilter = 'All'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Helpers ─────────────────────────────────────────────────────────────

function Write-Section ([string]$Text) {
    $bar = '─' * 60
    Write-Host "`n$bar" -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor Cyan
}

function Assert-GraphConnection {
    try {
        $ctx = Get-MgContext
        if (-not $ctx) { throw }
        Write-Host "[+] Connected as: $($ctx.Account)  |  Tenant: $($ctx.TenantId)" -ForegroundColor Green

        $required = @('AuditLog.Read.All', 'User.Read.All', 'Directory.Read.All')
        $missing  = $required | Where-Object { $_ -notin $ctx.Scopes }
        if ($missing) {
            Write-Warning "Possibly missing scope(s): $($missing -join ', '). Re-connect with -Scopes if calls fail."
        }
    }
    catch {
        Write-Error "Not connected to Microsoft Graph. Run Connect-MgGraph first."
        exit 1
    }
}

function Get-LookbackWindow {
    switch ($PSCmdlet.ParameterSetName) {
        'Hours' { return [datetime]::UtcNow.AddHours(-$Hours) }
        'Days'  { return [datetime]::UtcNow.AddDays(-$Days)   }
        'Range' { return $StartDate.ToUniversalTime()          }
    }
}

# Retrieve all pages of a Graph collection
function Invoke-GraphGetAll ([string]$Uri) {
    $results = [System.Collections.Generic.List[object]]::new()
    $next    = $Uri
    do {
        $page  = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($page.value) { $results.AddRange([object[]]$page.value) }
        $next  = $page.ContainsKey('@odata.nextLink') ? $page.'@odata.nextLink' : $null
    } while ($next)
    return $results
}

# Friendly SKU name lookup (extend as needed)
$SkuNames = @{
    'c7df2760-2c81-4ef7-b578-5b5392b571df' = 'Microsoft 365 E5'    # ENTERPRISEPREMIUM
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06' = 'Microsoft 365 E5'    # SPE_E5
}

function Resolve-SkuName ([string]$SkuId) {
    if ($SkuNames.ContainsKey($SkuId)) { return $SkuNames[$SkuId] }
    return $SkuId   # fall back to raw GUID
}

#endregion

#region ── Main ────────────────────────────────────────────────────────────────

Write-Section 'Licence Assignment Report'
Assert-GraphConnection

# ── Build time window ──────────────────────────────────────────────────────

$windowStart = Get-LookbackWindow
$windowEnd   = if ($PSCmdlet.ParameterSetName -eq 'Range') { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

$startStr = $windowStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
$endStr   = $windowEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')

Write-Host "[*] Query window: $startStr  →  $endStr" -ForegroundColor Yellow

# SKU part number → GUID (used for post-filtering)
$SkuGuids = @{
    'ENTERPRISEPREMIUM' = 'c7df2760-2c81-4ef7-b578-5b5392b571df'
    'SPE_E5'            = '06ebc4ee-1bb5-47dd-8120-11324bc54e06'
}
$filterSkuGuids = if ($SkuFilter -eq 'All') {
    $SkuGuids.Values
} else {
    @($SkuGuids[$SkuFilter])
}
Write-Host "[*] SKU filter: $SkuFilter" -ForegroundColor Yellow

# ── Fetch audit events ─────────────────────────────────────────────────────
#    activityDisplayName values that represent licence changes:
#      "Change user license"      – direct assignment / removal via portal or Graph
#      "Add member to group"      – group-based licensing (indirect)
#      "Add user"                 – new user creation (often paired with licence)
#    We post-filter on modifiedProperties to isolate actual licence changes.

$filterClauses = @(
    "activityDisplayName eq 'Change user license'",
    "activityDisplayName eq 'Add member to group'"
)
if ($IncludeRemoved) {
    $filterClauses += "activityDisplayName eq 'Remove member from group'"
}

$auditFilter = "($(($filterClauses -join ' or ')))"
$auditFilter += " and activityDateTime ge $startStr and activityDateTime le $endStr"

$auditUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?" +
            "`$filter=$([uri]::EscapeDataString($auditFilter))" +
            "&`$orderby=activityDateTime desc" +
            "&`$top=999"

Write-Host "[*] Querying audit logs..." -ForegroundColor Yellow
$auditEvents = @(Invoke-GraphGetAll -Uri $auditUri)
Write-Host "[+] Audit events retrieved: $($auditEvents.Count)" -ForegroundColor Green

# Post-filter to events that reference one of the target SKU GUIDs
# (auditLogs don't support $filter on modifiedProperties values directly)
if ($SkuFilter -ne 'All') {
    $auditEvents = @($auditEvents | Where-Object {
        $props = $_.targetResources.modifiedProperties
        $props | Where-Object {
            ($_.displayName -like '*license*' -or $_.displayName -eq 'AssignedLicense') -and
            ($filterSkuGuids | Where-Object { $_.newValue -like "*$_*" -or $_.oldValue -like "*$_*" })
        }
    })
    Write-Host "[*] Events after SKU filter ($SkuFilter): $($auditEvents.Count)" -ForegroundColor Yellow
}

if ($auditEvents.Count -eq 0) {
    Write-Host "`n[!] No licence events found in the specified window." -ForegroundColor Yellow
    exit 0
}

# ── Extract unique target user object IDs ─────────────────────────────────

$targetIds = @($auditEvents |
    ForEach-Object {
        $_.targetResources | Where-Object { $_.type -eq 'User' } | Select-Object -ExpandProperty id
    } |
    Where-Object { $_ } |
    Sort-Object -Unique)

Write-Host "[*] Unique users to enrich: $($targetIds.Count)" -ForegroundColor Yellow

# ── Batch-enrich users (20 per batch via $batch endpoint) ─────────────────

$userCache = @{}

$batches = [System.Collections.Generic.List[object[]]]::new()
for ($i = 0; $i -lt $targetIds.Count; $i += 20) {
    $batches.Add($targetIds[$i..([Math]::Min($i + 19, $targetIds.Count - 1))])
}

Write-Host "[*] Fetching user properties in $($batches.Count) batch(es)..." -ForegroundColor Yellow

foreach ($batch in $batches) {
    $batchRequests = @{
        requests = @(
            $batch | ForEach-Object -Begin { $idx = 1 } -Process {
                @{
                    id     = "$idx"
                    method = 'GET'
                    url    = "/users/$_`?`$select=id,userPrincipalName,displayName,usageLocation,onPremisesExtensionAttributes,assignedLicenses"
                }
                $idx++
            }
        )
    }

    $batchResponse = Invoke-MgGraphRequest -Method POST `
        -Uri 'https://graph.microsoft.com/v1.0/$batch' `
        -Body ($batchRequests | ConvertTo-Json -Depth 10) `
        -ContentType 'application/json'

    foreach ($resp in $batchResponse.responses) {
        if ($resp.status -eq 200) {
            $u = $resp.body
            $userCache[$u.id] = @{
                UPN              = $u.userPrincipalName
                DisplayName      = $u.displayName
                UsageLocation    = $u.usageLocation
                ExtensionAttr1   = $u.onPremisesExtensionAttributes?.extensionAttribute1
                AssignedLicenses = ($u.assignedLicenses | ForEach-Object { Resolve-SkuName $_.skuId }) -join '; '
            }
        }
        else {
            Write-Verbose "Batch sub-request $($resp.id) returned HTTP $($resp.status)"
        }
    }
}

Write-Host "[+] User enrichment complete." -ForegroundColor Green

# ── Build report rows ──────────────────────────────────────────────────────

Write-Section 'Processing Events'

$report = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($event in $auditEvents) {
    $activityTime = [datetime]::Parse($event.activityDateTime).ToUniversalTime()
    $activity     = $event.activityDisplayName
    $initiatedBy  = $event.initiatedBy.user.userPrincipalName ??
                    $event.initiatedBy.app.displayName        ??
                    'Unknown'

    # Determine which licences changed (present in modifiedProperties for direct assignments)
    $licencesBefore = ''
    $licencesAfter  = ''
    $assignedSkus   = ''

    $licenceProp = $event.targetResources.modifiedProperties |
        Where-Object { $_.displayName -in @('AssignedLicense', 'AccountEnabled') -or
                       $_.displayName -like '*license*' } |
        Select-Object -First 1

    if ($licenceProp) {
        $licencesBefore = $licenceProp.oldValue -replace '[\[\]"]', '' | Out-String -NoNewline
        $licencesAfter  = $licenceProp.newValue -replace '[\[\]"]', '' | Out-String -NoNewline
    }

    # Pull target user(s) from this event
    $targets = $event.targetResources | Where-Object { $_.type -eq 'User' }

    if (-not $targets) {
        # For group-based events the target may be a Group; record the group name at minimum
        $groupTarget = $event.targetResources | Where-Object { $_.type -eq 'Group' } | Select-Object -First 1
        $report.Add([PSCustomObject]@{
            EventDateTime    = $activityTime.ToString('yyyy-MM-dd HH:mm:ss')
            Activity         = $activity
            InitiatedBy      = $initiatedBy
            UserObjectId     = ''
            UPN              = ''
            DisplayName      = $groupTarget?.displayName ?? ''
            UsageLocation    = ''
            ExtensionAttr1   = ''
            CurrentLicences  = ''
            LicenceBefore    = $licencesBefore
            LicenceAfter     = $licencesAfter
            CorrelationId    = $event.correlationId
        })
        continue
    }

    foreach ($target in $targets) {
        $userId   = $target.id
        $cached   = $userCache[$userId]

        $report.Add([PSCustomObject]@{
            EventDateTime    = $activityTime.ToString('yyyy-MM-dd HH:mm:ss')
            Activity         = $activity
            InitiatedBy      = $initiatedBy
            UserObjectId     = $userId
            UPN              = $cached?.UPN              ?? $target.userPrincipalName ?? $userId
            DisplayName      = $cached?.DisplayName      ?? $target.displayName       ?? ''
            UsageLocation    = $cached?.UsageLocation    ?? ''
            ExtensionAttr1   = $cached?.ExtensionAttr1  ?? ''
            CurrentLicences  = $cached?.AssignedLicenses ?? ''
            LicenceBefore    = $licencesBefore
            LicenceAfter     = $licencesAfter
            CorrelationId    = $event.correlationId
        })
    }
}

# ── Output ─────────────────────────────────────────────────────────────────

Write-Section 'Summary'
Write-Host "  Total events processed : $($auditEvents.Count)"
Write-Host "  Report rows generated  : $($report.Count)"
Write-Host "  Unique users enriched  : $($userCache.Count)"

# Console preview (top 20)
$report | Select-Object -First 20 |
    Format-Table EventDateTime, Activity, UPN, DisplayName, UsageLocation, ExtensionAttr1 -AutoSize

if ($report.Count -gt 20) {
    Write-Host "  ... $($report.Count - 20) additional rows in CSV export." -ForegroundColor DarkGray
}

# CSV export
if (-not $ExportPath) {
    $ts         = Get-Date -Format 'yyyyMMdd-HHmmss'
    $ExportPath = Join-Path (Get-Location) "LicenceAssignments_$ts.csv"
}

$report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`n[+] Report exported to: $ExportPath" -ForegroundColor Green

#endregion
