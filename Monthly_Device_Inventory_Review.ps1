<#
.SYNOPSIS
    Monthly Entra ID device inventory review: disable stale devices (90+ days inactive)
    and delete long-stale devices (120+ days inactive), then export a CSV audit report.

.DESCRIPTION
    1.  Prompts for interactive tenant sign-in (Microsoft Graph).
    2.  Retrieves every Entra ID device object and calculates days since last sign-in
        (approximateLastSignInDateTime).
    3.  Prints all devices inactive for 90+ days, then asks whether to disable them.
    4.  Prints all devices inactive for 120+ days, then asks whether to delete them
        (requires typing DELETE to confirm, since deletion is not reversible).
    5.  Exports a CSV containing every device from both lists, showing which bucket it
        fell into and whether the disable/delete action succeeded.
    6.  Signs out of the tenant and reports whether the sign-out succeeded, so the script
        can be run safely against one tenant after another in the same session.

    Devices with no recorded sign-in activity at all are reported separately and are NOT
    acted upon unless -IncludeNeverSignedIn is specified.

.PARAMETER TenantId
    Optional tenant ID or domain to sign in to. Omit to be prompted / use the default tenant.

.PARAMETER DisableAfterDays
    Inactivity threshold (days) for the disable pass. Default 90.

.PARAMETER DeleteAfterDays
    Inactivity threshold (days) for the delete pass. Default 120.

.PARAMETER CsvPath
    Output path for the report. Defaults to a file beside this script named after the tenant
    and today's date, e.g. signing in as chadmin@contoso.com produces
    Contoso_Device_Review_2026-09-02.csv. A second run on the same day gets a _2, _3, ... suffix
    rather than overwriting the earlier report.

.PARAMETER IncludeNeverSignedIn
    Include devices that have never reported sign-in activity in the disable/delete passes.

.PARAMETER LineDelayMilliseconds
    Pause between printed lines so the inventory and device lists scroll by at a readable pace
    instead of appearing all at once. Default 60ms; use 0 for full speed. Total added delay is
    capped at 45 seconds, after which output prints at full speed automatically.

.EXAMPLE
    .\Monthly_Device_Inventory_Review.ps1

.EXAMPLE
    .\Monthly_Device_Inventory_Review.ps1 -TenantId contoso.onmicrosoft.com -IncludeNeverSignedIn

.NOTES
    Requires the Microsoft.Graph PowerShell SDK and an account with permission to consent to
    Device.ReadWrite.All + Directory.ReadWrite.All (Cloud Device Administrator / Intune
    Administrator / Global Administrator).
#>

[CmdletBinding()]
param(
    [string]$TenantId,

    [ValidateRange(1, 3650)]
    [int]$DisableAfterDays = 90,

    [ValidateRange(1, 3650)]
    [int]$DeleteAfterDays = 120,

    [string]$CsvPath,

    [switch]$IncludeNeverSignedIn,

    [ValidateRange(0, 2000)]
    [int]$LineDelayMilliseconds = 60
)

$ErrorActionPreference = 'Stop'

# Paced-output settings. Artificial delays make the run followable in real time, but they are
# capped by a total budget so a large tenant cannot turn a 30-second run into a 10-minute one.
$script:LineDelayMs      = $LineDelayMilliseconds
$script:DelayBudgetMs    = 45000
$script:DelaySpentMs     = 0
$script:DelayBudgetSpent = $false

#region Helpers ---------------------------------------------------------------

function Wait-Line {
    <#
        Sleeps for the configured per-line delay, drawing from a shared budget.
        Once the budget is exhausted, remaining output prints at full speed.
    #>
    param([int]$Milliseconds = $script:LineDelayMs)

    if ($Milliseconds -le 0 -or $script:DelayBudgetSpent) { return }

    if ($script:DelaySpentMs -ge $script:DelayBudgetMs) {
        $script:DelayBudgetSpent = $true
        Write-Host '  (output pacing limit reached - printing the remainder at full speed)' -ForegroundColor DarkGray
        return
    }

    Start-Sleep -Milliseconds $Milliseconds
    $script:DelaySpentMs += $Milliseconds
}

function Write-PacedHost {
    <# Write-Host with a pause afterwards so lines appear one at a time. #>
    param(
        [Parameter(Position = 0)][string]$Object = '',
        [System.ConsoleColor]$ForegroundColor,
        [int]$Milliseconds = $script:LineDelayMs
    )

    if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        Write-Host $Object -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Object
    }

    Wait-Line -Milliseconds $Milliseconds
}

function Write-Banner {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}

function Get-YesNoAnswer {
    param(
        [Parameter(Mandatory)][string]$Question
    )
    while ($true) {
        $answer = (Read-Host "$Question [Y/N]").Trim()
        switch -Regex ($answer) {
            '^(y|yes)$' { return $true }
            '^(n|no)$'  { return $false }
            default     { Write-Host 'Please answer Y or N.' -ForegroundColor Yellow }
        }
    }
}

function Get-TenantSlug {
    <#
        Derives a friendly, filename-safe tenant name for the report.
        chadmin@contoso.com            -> Contoso
        admin@contoso.onmicrosoft.com  -> Contoso
        admin@contoso.co.uk            -> Contoso
        Falls back to the tenant's initial verified domain (app-only sign-in has no UPN),
        then to the tenant ID.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string]$Account,
        [string]$TenantIdValue
    )

    function Format-Slug {
        param([string]$Text)
        # Keep letters, digits, dash and underscore; everything else would be illegal or noisy.
        $clean = ($Text -replace '[^A-Za-z0-9\-_]', '')
        if (-not $clean) { return $null }
        return [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($clean.ToLower())
    }

    $domain = $null
    if ($Account -and $Account.Contains('@')) {
        $domain = ($Account -split '@')[-1]
    }

    if (-not $domain) {
        try {
            $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
            $initialDomain = $org.VerifiedDomains | Where-Object { $_.IsInitial } | Select-Object -First 1
            if ($initialDomain) { $domain = $initialDomain.Name }
            elseif ($org.DisplayName) {
                $fromName = Format-Slug -Text $org.DisplayName
                if ($fromName) { return $fromName }
            }
        }
        catch {
            Write-Verbose "Could not read organization details for the report name: $($_.Exception.Message)"
        }
    }

    if ($domain) {
        # First label of the domain is the tenant name in every form above.
        $slug = Format-Slug -Text (($domain -split '\.')[0])
        if ($slug) { return $slug }
    }

    if ($TenantIdValue) { return "Tenant-$TenantIdValue" }
    return 'UnknownTenant'
}

function Disconnect-MgGraphSession {
    <#
        Signs out of the current Microsoft Graph session and reports the outcome.
        Verifies the sign-out by confirming no Graph context remains, so a run against the
        next tenant cannot silently inherit this tenant's session.
        Returns $true on a verified sign-out, otherwise $false.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $activeContext = Get-MgContext -ErrorAction SilentlyContinue

    if (-not $activeContext) {
        Write-Host '  [ -- ]   No active Microsoft Graph session to sign out of.' -ForegroundColor DarkGray
        return $true
    }

    $account = $activeContext.Account
    $tenant  = $activeContext.TenantId

    try {
        Disconnect-MgGraph -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "  [FAIL]   Sign-out failed for tenant $tenant ($account)" -ForegroundColor Red
        Write-Host "           Reason    : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host '           The session may still be active - run Disconnect-MgGraph manually' -ForegroundColor Red
        Write-Host '           before running this script against another tenant.' -ForegroundColor Red
        return $false
    }

    # Confirm the context is really gone rather than trusting the cmdlet's silence.
    $remainingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($remainingContext) {
        Write-Host "  [FAIL]   Sign-out reported no error but a Graph session is still present" -ForegroundColor Red
        Write-Host "           Tenant    : $($remainingContext.TenantId) ($($remainingContext.Account))" -ForegroundColor Red
        Write-Host '           Run Disconnect-MgGraph manually before switching tenants.' -ForegroundColor Red
        return $false
    }

    Write-Host "  [ OK ]   Signed out of tenant $tenant ($account)" -ForegroundColor Green
    return $true
}

function Show-DeviceTable {
    param([object[]]$Devices)

    if (-not $Devices -or $Devices.Count -eq 0) {
        Write-Host '  (none)' -ForegroundColor DarkGray
        return
    }

    # Let Format-Table handle column alignment, then emit the rendered lines one at a time so
    # the table builds up on screen instead of appearing all at once.
    $Devices |
        Sort-Object -Property @{ Expression = { $_.DaysInactive }; Descending = $true } |
        Format-Table -AutoSize -Property `
            @{ Label = 'Days Inactive'; Expression = { if ($null -eq $_.DaysInactive) { 'never' } else { $_.DaysInactive } } },
            @{ Label = 'Display Name';  Expression = { $_.DisplayName } },
            @{ Label = 'OS';            Expression = { (@($_.OperatingSystem, $_.OperatingSystemVersion) | Where-Object { $_ }) -join ' ' } },
            @{ Label = 'Trust';         Expression = { $_.TrustType } },
            @{ Label = 'Enabled';       Expression = { $_.AccountEnabled } },
            @{ Label = 'Last Sign-In';  Expression = { if ($_.ApproximateLastSignIn) { $_.ApproximateLastSignIn.ToString('yyyy-MM-dd') } else { 'never' } } } |
        Out-String -Stream -Width 240 |
        ForEach-Object {
            $line = $_.TrimEnd()
            Write-Host $line
            # Don't burn the delay budget on the blank/header lines Format-Table emits.
            if ($line) { Wait-Line }
        }
}

#endregion

# Guarantee sign-out even if the script terminates early, so a failed run never leaves a
# live session behind for the next tenant.
trap {
    Write-Host ''
    Write-Host "Script stopped with an unhandled error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    Write-Host 'Signing out of Microsoft Graph...' -ForegroundColor Gray
    Disconnect-MgGraphSession | Out-Null
    break
}

#region Module + sign-in ------------------------------------------------------

$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
$missingModules = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }

if ($missingModules) {
    Write-Host "Missing required module(s): $($missingModules -join ', ')" -ForegroundColor Yellow
    if (Get-YesNoAnswer -Question 'Install them now from the PowerShell Gallery for the current user?') {
        foreach ($module in $missingModules) {
            Write-Host "Installing $module ..." -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        }
    }
    else {
        throw "Cannot continue without: $($missingModules -join ', ')"
    }
}

foreach ($module in $requiredModules) {
    Import-Module -Name $module -ErrorAction Stop
}

Write-Banner 'Entra ID Device Inventory Review'

# A session left over from a previous tenant would silently be reused - clear it first.
$staleContext = Get-MgContext -ErrorAction SilentlyContinue
if ($staleContext) {
    Write-Host "An existing Microsoft Graph session was found for tenant $($staleContext.TenantId) ($($staleContext.Account))." -ForegroundColor Yellow
    Write-Host 'Signing out of it before connecting to the target tenant...' -ForegroundColor Yellow
    if (-not (Disconnect-MgGraphSession)) {
        throw 'Could not clear the existing Microsoft Graph session - aborting to avoid acting against the wrong tenant.'
    }
    Write-Host ''
}

Write-Host 'Signing in to Microsoft Graph (a browser/device-code prompt will appear)...' -ForegroundColor Gray

$connectParams = @{
    Scopes      = @('Device.ReadWrite.All', 'Directory.ReadWrite.All')
    NoWelcome   = $true
    ErrorAction = 'Stop'
}
if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Connect-MgGraph @connectParams

$context = Get-MgContext
if (-not $context) { throw 'Sign-in failed - no Microsoft Graph context available.' }

Write-Host ''
Write-Host "Signed in as : $($context.Account)"    -ForegroundColor Green
Write-Host "Tenant       : $($context.TenantId)"   -ForegroundColor Green
Write-Host "Scopes       : $($context.Scopes -join ', ')" -ForegroundColor DarkGray

# Resolved now so the report name is known before any changes are made.
$tenantSlug = Get-TenantSlug -Account $context.Account -TenantIdValue $context.TenantId
Write-Host "Report name  : ${tenantSlug}_Device_Review_$((Get-Date).ToString('yyyy-MM-dd')).csv" -ForegroundColor DarkGray

#endregion

#region Inventory -------------------------------------------------------------

$now         = (Get-Date).ToUniversalTime()
$disableFrom = $now.AddDays(-$DisableAfterDays)
$deleteFrom  = $now.AddDays(-$DeleteAfterDays)

Write-Host ''
Write-Host 'Retrieving device objects from Entra ID...' -ForegroundColor Gray
Write-Host "  Legend: " -NoNewline -ForegroundColor DarkGray
Write-Host 'active' -NoNewline -ForegroundColor DarkGray
Write-Host ' | ' -NoNewline -ForegroundColor DarkGray
Write-Host "$DisableAfterDays+ days idle" -NoNewline -ForegroundColor Yellow
Write-Host ' | ' -NoNewline -ForegroundColor DarkGray
Write-Host "$DeleteAfterDays+ days idle" -NoNewline -ForegroundColor Red
Write-Host ' | ' -NoNewline -ForegroundColor DarkGray
Write-Host 'no sign-in data' -ForegroundColor Magenta
Write-Host ''

# Piping Get-MgDevice (rather than assigning it) lets each page render as it arrives, so the
# inventory scrolls past in real time instead of appearing in one block at the end.
$retrievalDelayMs = [math]::Min(25, $script:LineDelayMs)
$counter = 0

$devices = Get-MgDevice -All -Property @(
    'id', 'deviceId', 'displayName', 'operatingSystem', 'operatingSystemVersion',
    'accountEnabled', 'approximateLastSignInDateTime', 'trustType',
    'isManaged', 'isCompliant', 'registrationDateTime', 'profileType'
) -ErrorAction Stop | ForEach-Object {
    $device = $_
    $counter++

    $lastSignIn = $null
    if ($device.ApproximateLastSignInDateTime) {
        $lastSignIn = ([datetime]$device.ApproximateLastSignInDateTime).ToUniversalTime()
    }

    $daysInactive = if ($lastSignIn) { [int][math]::Floor(($now - $lastSignIn).TotalDays) } else { $null }

    if ($null -eq $daysInactive) {
        $colour   = [System.ConsoleColor]::Magenta
        $ageText  = 'no sign-in data'
    }
    else {
        $colour = if ($daysInactive -ge $DeleteAfterDays)      { [System.ConsoleColor]::Red }
                  elseif ($daysInactive -ge $DisableAfterDays) { [System.ConsoleColor]::Yellow }
                  else                                         { [System.ConsoleColor]::DarkGray }
        $ageText = "{0,4} days idle  (last sign-in {1})" -f $daysInactive, $lastSignIn.ToString('yyyy-MM-dd')
    }

    $name = if ($device.DisplayName) { $device.DisplayName } else { '(no display name)' }
    if ($name.Length -gt 40) { $name = $name.Substring(0, 37) + '...' }

    Write-Host ("  [{0,5}]  {1,-40}  {2}" -f $counter, $name, $ageText) -ForegroundColor $colour
    Wait-Line -Milliseconds $retrievalDelayMs

    [pscustomobject]@{
        ObjectId               = $device.Id
        DeviceId               = $device.DeviceId
        DisplayName            = $device.DisplayName
        OperatingSystem        = $device.OperatingSystem
        OperatingSystemVersion = $device.OperatingSystemVersion
        TrustType              = $device.TrustType
        ProfileType            = $device.ProfileType
        IsManaged              = $device.IsManaged
        IsCompliant            = $device.IsCompliant
        RegistrationDateTime   = $device.RegistrationDateTime
        AccountEnabled         = $device.AccountEnabled
        ApproximateLastSignIn  = $lastSignIn
        DaysInactive           = $daysInactive
    }
}

$devices = @($devices)

Write-Host ''
Write-Host "Retrieved $($devices.Count) device object(s)." -ForegroundColor Gray

$neverSignedIn = @($devices | Where-Object { $null -eq $_.ApproximateLastSignIn })

$staleDisable = @($devices | Where-Object {
    $_.ApproximateLastSignIn -and $_.ApproximateLastSignIn -le $disableFrom
})
$staleDelete = @($devices | Where-Object {
    $_.ApproximateLastSignIn -and $_.ApproximateLastSignIn -le $deleteFrom
})

if ($IncludeNeverSignedIn) {
    $staleDisable += $neverSignedIn
    $staleDelete  += $neverSignedIn
}

# Report rows keyed by object ID so disable + delete results land on the same record.
$report = [ordered]@{}
function Add-ReportRow {
    param([Parameter(Mandatory)][object]$Device)

    if ($report.Contains($Device.ObjectId)) { return $report[$Device.ObjectId] }

    $row = [pscustomobject]@{
        DisplayName            = $Device.DisplayName
        DeviceId               = $Device.DeviceId
        ObjectId               = $Device.ObjectId
        OperatingSystem        = $Device.OperatingSystem
        OperatingSystemVersion = $Device.OperatingSystemVersion
        TrustType              = $Device.TrustType
        ProfileType            = $Device.ProfileType
        IsManaged              = $Device.IsManaged
        IsCompliant            = $Device.IsCompliant
        RegistrationDateTime   = if ($Device.RegistrationDateTime) { ([datetime]$Device.RegistrationDateTime).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        LastSignInUtc          = if ($Device.ApproximateLastSignIn) { $Device.ApproximateLastSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { 'never' }
        DaysInactive           = $Device.DaysInactive
        AccountEnabledBefore   = $Device.AccountEnabled
        "Inactive${DisableAfterDays}Plus" = 'No'
        "Inactive${DeleteAfterDays}Plus"  = 'No'
        DisableAttempted       = 'No'
        DisableResult          = ''
        DeleteAttempted        = 'No'
        DeleteResult           = ''
        ErrorMessage           = ''
        ActionTimestampUtc     = ''
    }

    $report[$Device.ObjectId] = $row
    return $row
}

foreach ($device in $staleDisable) { (Add-ReportRow -Device $device)."Inactive${DisableAfterDays}Plus" = 'Yes' }
foreach ($device in $staleDelete)  { (Add-ReportRow -Device $device)."Inactive${DeleteAfterDays}Plus"  = 'Yes' }

#endregion

#region Pass 1 - disable 90+ day devices --------------------------------------

Write-Banner "Devices with no activity in $DisableAfterDays+ days (on or before $($disableFrom.ToString('yyyy-MM-dd')) UTC)"
Write-Host "Count: $($staleDisable.Count)" -ForegroundColor White
Show-DeviceTable -Devices $staleDisable

if (-not $IncludeNeverSignedIn -and $neverSignedIn.Count -gt 0) {
    Write-Host "NOTE: $($neverSignedIn.Count) device(s) have no recorded sign-in activity and were EXCLUDED from all actions." -ForegroundColor Yellow
    Write-Host '      Re-run with -IncludeNeverSignedIn to include them.' -ForegroundColor Yellow
    Show-DeviceTable -Devices $neverSignedIn
}

$alreadyDisabled = @($staleDisable | Where-Object { $_.AccountEnabled -eq $false })
$toDisable       = @($staleDisable | Where-Object { $_.AccountEnabled -ne $false })

foreach ($device in $alreadyDisabled) {
    $row = Add-ReportRow -Device $device
    $row.DisableResult = 'Already Disabled'
}

if ($staleDisable.Count -eq 0) {
    Write-Host 'Nothing to disable.' -ForegroundColor Green
}
elseif ($toDisable.Count -eq 0) {
    Write-Host "All $($staleDisable.Count) device(s) are already disabled - nothing to do." -ForegroundColor Green
}
else {
    Write-Host ''
    if ($alreadyDisabled.Count -gt 0) {
        Write-Host "$($alreadyDisabled.Count) of these are already disabled and will be skipped." -ForegroundColor DarkGray
    }

    if (Get-YesNoAnswer -Question "Disable the $($toDisable.Count) enabled device(s) listed above?") {
        Write-Host ''
        $disableSucceeded = 0
        $disableFailed    = 0

        foreach ($device in $toDisable) {
            $row = Add-ReportRow -Device $device
            $row.DisableAttempted   = 'Yes'
            $row.ActionTimestampUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

            try {
                Update-MgDevice -DeviceId $device.ObjectId -BodyParameter @{ accountEnabled = $false } -ErrorAction Stop
                $row.DisableResult = 'Success'
                $disableSucceeded++
                Write-Host "  [ OK ]   Disabled  : $($device.DisplayName)  ($($device.ObjectId))" -ForegroundColor Green
            }
            catch {
                $row.DisableResult = 'Failed'
                $row.ErrorMessage  = $_.Exception.Message
                $disableFailed++
                Write-Host "  [FAIL]   Disable   : $($device.DisplayName)  ($($device.ObjectId))" -ForegroundColor Red
                Write-Host "           Reason    : $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        Write-Host ''
        Write-Host "Disable summary: $disableSucceeded succeeded, $disableFailed failed, $($alreadyDisabled.Count) skipped (already disabled)." -ForegroundColor $(if ($disableFailed) { 'Yellow' } else { 'Green' })
    }
    else {
        Write-Host 'Skipped - no devices were disabled.' -ForegroundColor Yellow
        foreach ($device in $toDisable) {
            (Add-ReportRow -Device $device).DisableResult = 'Declined by operator'
        }
    }
}

#endregion

#region Pass 2 - delete 120+ day devices --------------------------------------

Write-Banner "Devices with no activity in $DeleteAfterDays+ days (on or before $($deleteFrom.ToString('yyyy-MM-dd')) UTC)"
Write-Host "Count: $($staleDelete.Count)" -ForegroundColor White
Show-DeviceTable -Devices $staleDelete

if ($staleDelete.Count -eq 0) {
    Write-Host 'Nothing to delete.' -ForegroundColor Green
}
else {
    Write-Host ''
    Write-Host 'WARNING: Deleting an Entra ID device object is PERMANENT. Hybrid-joined devices may' -ForegroundColor Yellow
    Write-Host '         resync from AD Connect, and affected users will have to re-register/re-enroll.' -ForegroundColor Yellow
    Write-Host ''

    if (Get-YesNoAnswer -Question "Delete the $($staleDelete.Count) device(s) listed above?") {
        $typed = (Read-Host "Type DELETE (all caps) to permanently remove $($staleDelete.Count) device object(s)").Trim()

        if ($typed -cne 'DELETE') {
            Write-Host 'Confirmation text did not match - no devices were deleted.' -ForegroundColor Yellow
            foreach ($device in $staleDelete) {
                (Add-ReportRow -Device $device).DeleteResult = 'Declined by operator'
            }
        }
        else {
            Write-Host ''
            $deleteSucceeded = 0
            $deleteFailed    = 0

            foreach ($device in $staleDelete) {
                $row = Add-ReportRow -Device $device
                $row.DeleteAttempted    = 'Yes'
                $row.ActionTimestampUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

                try {
                    Remove-MgDevice -DeviceId $device.ObjectId -ErrorAction Stop
                    $row.DeleteResult = 'Success'
                    $deleteSucceeded++
                    Write-Host "  [ OK ]   Deleted   : $($device.DisplayName)  ($($device.ObjectId))" -ForegroundColor Green
                }
                catch {
                    $row.DeleteResult = 'Failed'
                    $row.ErrorMessage = (@($row.ErrorMessage, $_.Exception.Message) | Where-Object { $_ }) -join ' | '
                    $deleteFailed++
                    Write-Host "  [FAIL]   Delete    : $($device.DisplayName)  ($($device.ObjectId))" -ForegroundColor Red
                    Write-Host "           Reason    : $($_.Exception.Message)" -ForegroundColor Red
                }
            }

            Write-Host ''
            Write-Host "Delete summary: $deleteSucceeded succeeded, $deleteFailed failed." -ForegroundColor $(if ($deleteFailed) { 'Yellow' } else { 'Green' })
        }
    }
    else {
        Write-Host 'Skipped - no devices were deleted.' -ForegroundColor Yellow
        foreach ($device in $staleDelete) {
            (Add-ReportRow -Device $device).DeleteResult = 'Declined by operator'
        }
    }
}

#endregion

#region CSV report -----------------------------------------------------------

Write-Banner 'Report'

if (-not $CsvPath) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $stamp     = (Get-Date).ToString('yyyy-MM-dd')
    $CsvPath   = Join-Path $scriptDir "${tenantSlug}_Device_Review_$stamp.csv"

    # The name is date-only, so a second run on the same day would overwrite the first.
    $suffix = 2
    while (Test-Path -LiteralPath $CsvPath) {
        $CsvPath = Join-Path $scriptDir "${tenantSlug}_Device_Review_${stamp}_$suffix.csv"
        $suffix++
    }
}

$rows = @($report.Values | Sort-Object -Property @{ Expression = { $_.DaysInactive }; Descending = $true })

if ($rows.Count -eq 0) {
    Write-Host 'No devices met either inactivity threshold - no CSV written.' -ForegroundColor Green
}
else {
    try {
        $csvDir = Split-Path -Path $CsvPath -Parent
        if ($csvDir -and -not (Test-Path -LiteralPath $csvDir)) {
            New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
        }

        $rows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV report written: $CsvPath" -ForegroundColor Green
        Write-Host "Rows: $($rows.Count)" -ForegroundColor Gray
    }
    catch {
        Write-Host "Failed to write CSV to '$CsvPath': $($_.Exception.Message)" -ForegroundColor Red
    }

    $disabledOk = @($rows | Where-Object { $_.DisableResult -eq 'Success' }).Count
    $deletedOk  = @($rows | Where-Object { $_.DeleteResult  -eq 'Success' }).Count
    $failed     = @($rows | Where-Object { $_.DisableResult -eq 'Failed' -or $_.DeleteResult -eq 'Failed' }).Count

    Write-Host ''
    Write-Host "Devices inactive $DisableAfterDays+ days : $($staleDisable.Count)"
    Write-Host "Devices inactive $DeleteAfterDays+ days : $($staleDelete.Count)"
    Write-Host "Successfully disabled              : $disabledOk"
    Write-Host "Successfully deleted               : $deletedOk"
    Write-Host "Failed actions                     : $failed" -ForegroundColor $(if ($failed) { 'Red' } else { 'Gray' })
}

#endregion

#region Sign-out -------------------------------------------------------------

Write-Banner 'Sign-out'

$signOutOk = Disconnect-MgGraphSession

Write-Host ''
if ($signOutOk) {
    Write-Host 'Sign-out successful - safe to run this script against the next tenant.' -ForegroundColor Green
}
else {
    Write-Host 'Sign-out UNSUCCESSFUL - resolve this before running against another tenant.' -ForegroundColor Red
}

#endregion
