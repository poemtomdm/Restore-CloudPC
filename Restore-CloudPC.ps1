# ============================================================
#  SCRIPT   : Restore-CloudPC.ps1
#  VERSION  : 1.0
#  AUTHOR   : Tom Machado
#  CREATED  : 2026-07-10
# ============================================================
#
#  DESCRIPTION
#  -----------
#  This interactive script restores Windows 365 Cloud PCs to a
#  chosen snapshot (restore point) for a bulk set of users defined
#  in a CSV file.
#
#  The script will:
#    1. Connect to Microsoft Graph using app-only authentication
#       (client secret or certificate).
#    2. Prompt the user to provide the path to a CSV file containing
#       user UPNs (column header: "UPN").
#    3. Retrieve all available restore points and let the user pick
#       one interactively.
#    4. For each UPN, locate the matching Cloud PC via the Graph API,
#       find the snapshot that best matches the chosen restore date,
#       and trigger the restore action.
#    5. Print a summary of successes and failures.
#
#  RESTORE POINT AVAILABILITY
#  --------------------------
#  Windows 365 keeps automatic snapshots on the following schedule:
#    - Daily snapshots for the last 7 days (AM and PM)
#    - Weekly snapshots beyond 7 days (up to a few weeks)
#  The list of available restore points is fetched live from the first
#  Cloud PC found for the first user, so it reflects the actual state
#  of your tenant at runtime.
#
#  CSV FORMAT
#  ----------
#  The CSV file must contain at least one column named "UPN":
#
#    UPN
#    alice@contoso.com
#    bob@contoso.com
#    carol@contoso.com
#
#  PREREQUISITES
#  -------------
#  - PowerShell 7+ (recommended) or Windows PowerShell 5.1
#  - Microsoft.Graph PowerShell SDK installed:
#      Install-Module Microsoft.Graph -Scope CurrentUser
#  - An Entra ID (Azure AD) App Registration with:
#      * Application permissions (NOT delegated):
#          - CloudPC.ReadWrite.All
#      * Admin consent granted for the above permissions
#  - Authentication: either a Client Secret or a Certificate associated
#    with the App Registration (see usage examples below)
#
#  USAGE EXAMPLES
#  --------------
#  # Client Secret auth:
#  .\Restore-CloudPC.ps1 `
#      -TenantId     "contoso.onmicrosoft.com" `
#      -ClientId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
#      -ClientSecret "your-client-secret-here"
#
#  # Certificate auth (by thumbprint — cert must be in the current user's cert store):
#  .\Restore-CloudPC.ps1 `
#      -TenantId              "contoso.onmicrosoft.com" `
#      -ClientId              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
#      -CertificateThumbprint "AABBCCDDEEFF00112233445566778899AABBCCDD"
#
#  # Certificate auth (by .pfx file path):
#  .\Restore-CloudPC.ps1 `
#      -TenantId            "contoso.onmicrosoft.com" `
#      -ClientId            "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
#      -CertificatePath     "C:\certs\myapp.pfx" `
#      -CertificatePassword (Read-Host -AsSecureString "PFX password")
# ============================================================

[CmdletBinding(DefaultParameterSetName = 'ClientSecret')]
param (
    # ── Identity ──────────────────────────────────────────────────────────────
    [Parameter(Mandatory, HelpMessage = "Azure AD tenant ID or domain name (e.g. contoso.onmicrosoft.com).")]
    [string] $TenantId,

    [Parameter(Mandatory, HelpMessage = "Application (client) ID of the registered app.")]
    [string] $ClientId,

    # ── Auth: Client Secret ───────────────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'ClientSecret',
               HelpMessage = "Client secret for the registered app.")]
    [string] $ClientSecret,

    # ── Auth: Certificate (thumbprint) ────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'CertThumbprint',
               HelpMessage = "Thumbprint of the certificate already installed in the local certificate store.")]
    [string] $CertificateThumbprint,

    # ── Auth: Certificate (pfx file) ──────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'CertFile',
               HelpMessage = "Path to the .pfx certificate file.")]
    [string] $CertificatePath,

    [Parameter(ParameterSetName = 'CertFile',
               HelpMessage = "Password for the .pfx file (leave empty if the file has no password).")]
    [SecureString] $CertificatePassword
)

#region --- Connection ---

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow

try {
    switch ($PSCmdlet.ParameterSetName) {

        'ClientSecret' {
            $secureSecret           = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $ClientSecretCredential = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -ErrorAction Stop
        }

        'CertThumbprint' {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
        }

        'CertFile' {
            $cert = if ($CertificatePassword) {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    (Resolve-Path $CertificatePath).Path,
                    $CertificatePassword
                )
            } else {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    (Resolve-Path $CertificatePath).Path
                )
            }
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $cert -ErrorAction Stop
        }
    }

    Write-Host "Connected to Microsoft Graph successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

#endregion

#region --- CSV input ---

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Windows 365 Cloud PC Restore Tool" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script restores Cloud PCs for a list of users defined in a CSV file." -ForegroundColor Yellow
Write-Host "The CSV must contain a column named 'UPN' with one user principal name per row." -ForegroundColor Yellow
Write-Host ""

# Prompt for CSV path until a valid file is provided
while ($true) {
    $csvPath = Read-Host "Enter the full path to the UPN CSV file (e.g. C:\Users\you\users.csv)"
    $csvPath = $csvPath.Trim('"').Trim("'")   # strip accidental quotes from drag-and-drop

    if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
        Write-Host "File not found: $csvPath — please try again." -ForegroundColor Red
        continue
    }

    try {
        # Try with header first, fall back to headerless (plain list of UPNs)
        [string[]]$rawLines = @(Get-Content -LiteralPath $csvPath -ErrorAction Stop) |
                    Where-Object { $_.Trim() -ne '' }

        if ($rawLines.Count -eq 0) {
            Write-Host "The file is empty. Please check the file and try again." -ForegroundColor Red
            continue
        }

        $firstLine = $rawLines[0].Trim()

        if ($firstLine -match '@') {
            # No header — every line is a UPN
            $userList = $rawLines | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' }
        } else {
            # Has a header row — parse as CSV and look for a UPN column
            $csvData = Import-Csv -LiteralPath $csvPath -ErrorAction Stop
            $colName = ($csvData | Select-Object -First 1).PSObject.Properties.Name |
                       Where-Object { $_ -match '^upn$' } |
                       Select-Object -First 1

            if (-not $colName) {
                Write-Host "The CSV does not contain a 'UPN' column. Please check the file and try again." -ForegroundColor Red
                continue
            }

            $userList = $csvData | Select-Object -ExpandProperty $colName |
                        ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' }
        }
    }
    catch {
        Write-Host "Failed to read file: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    if ($userList.Count -eq 0) {
        Write-Host "No valid UPNs found in the file. Please check the file and try again." -ForegroundColor Red
        continue
    }

    Write-Host "$($userList.Count) UPN(s) loaded from CSV." -ForegroundColor Green
    break
}

#endregion

#region --- Discover restore points ---
# Restore points are per-Cloud PC, but they are aligned across the tenant.
# We fetch snapshots from the FIRST Cloud PC we can locate to build the
# interactive menu, then reuse the chosen date when processing all users.

Write-Host ""
Write-Host "Fetching available restore points..." -ForegroundColor Yellow

$referenceSnapshots = $null

foreach ($upn in $userList) {
    $uriSearch    = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=userPrincipalName eq '$($upn.ToLower())'"
    $searchResult = Invoke-MgGraphRequest -Uri $uriSearch -Method GET

    if ($searchResult.value) {
        $refCloudPcId   = $searchResult.value[0].id
        $uriSnap        = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$refCloudPcId/retrieveSnapshots"
        $snapResponse   = Invoke-MgGraphRequest -Uri $uriSnap -Method GET
        $referenceSnapshots = $snapResponse.value | Sort-Object { [datetime]$_.createdDateTime } -Descending
        break
    }
}

if (-not $referenceSnapshots -or $referenceSnapshots.Count -eq 0) {
    Write-Host "Could not retrieve any restore points. Verify that at least one user has a Cloud PC." -ForegroundColor Red
    exit 1
}

#endregion

#region --- Restore point selection ---

Write-Host ""
Write-Host "Available restore points:" -ForegroundColor Yellow
Write-Host ""

# Build a deduplicated list of dates (one entry per distinct calendar day).
# Within a day, multiple snapshots can exist (AM / PM). The user picks a DATE;
# the script then selects the LATEST snapshot on that day for each Cloud PC.
$distinctDates = $referenceSnapshots |
    ForEach-Object { ([datetime]$_.createdDateTime).Date } |
    Select-Object -Unique |
    Sort-Object -Descending

$menuIndex = 1
$dateMenu  = @{}

$today = (Get-Date).Date

foreach ($date in $distinctDates) {
    $dayLabel = $date.ToString("dddd, MMMM d, yyyy")
    $age      = ($today - $date).Days

    if ($age -eq 0) {
        $suffix = " (today)"
        $color  = "Green"
    } elseif ($age -le 7) {
        $suffix = " ($age day(s) ago)"
        $color  = "Cyan"
    } else {
        $suffix = " ($age day(s) ago — weekly)"
        $color  = "DarkGray"
    }

    Write-Host "  $menuIndex. $dayLabel$suffix" -ForegroundColor $color
    $dateMenu[$menuIndex.ToString()] = $date
    $menuIndex++
}

Write-Host ""

$selectedDate = $null

while ($null -eq $selectedDate) {
    $choice = Read-Host "Enter the number of the restore point to use"
    $choice = $choice.Trim()

    if ($dateMenu.ContainsKey($choice)) {
        $selectedDate = $dateMenu[$choice]
    } else {
        Write-Host "Invalid choice. Please enter a number between 1 and $($dateMenu.Count)." -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Selected restore date : $($selectedDate.ToString('dddd, MMMM d, yyyy'))" -ForegroundColor Green
Write-Host "The script will use the LATEST snapshot available on that date for each Cloud PC." -ForegroundColor DarkGray

#endregion

#region --- Processing loop ---

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Starting Cloud PC restores..." -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$successCount = 0
$failCount    = 0
$skippedCount = 0

foreach ($upn in $userList) {

    Write-Host "Processing: $upn" -ForegroundColor Cyan

    # ── 1. Find the Cloud PC for this user ──────────────────────────────────
    $uriSearch    = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=userPrincipalName eq '$($upn.ToLower())'"
    $searchResult = Invoke-MgGraphRequest -Uri $uriSearch -Method GET

    if (-not $searchResult.value) {
        Write-Host "  [SKIP] No Cloud PC found for $upn." -ForegroundColor Yellow
        $skippedCount++
        Write-Host "  ------------------------------------"
        continue
    }

    $cloudPCs = $searchResult.value
    Write-Host "  $($cloudPCs.Count) Cloud PC(s) found." -ForegroundColor DarkGray

    foreach ($cloudPC in $cloudPCs) {

        $cloudPcId   = $cloudPC.id
        $cloudPcName = if ($cloudPC.displayName) { $cloudPC.displayName } else { $cloudPcId }
        Write-Host "  ── $cloudPcName  ($cloudPcId)" -ForegroundColor DarkGray

        # ── 2. Retrieve snapshots for this Cloud PC ──────────────────────────
        $uriSnap      = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$cloudPcId/retrieveSnapshots"
        $snapResponse = Invoke-MgGraphRequest -Uri $uriSnap -Method GET

        # Pick the LATEST snapshot whose date matches the user-selected day
        $targetSnapshot = $snapResponse.value |
            Where-Object { ([datetime]$_.createdDateTime).Date -eq $selectedDate } |
            Sort-Object { [datetime]$_.createdDateTime } -Descending |
            Select-Object -First 1

        if (-not $targetSnapshot) {
            Write-Host "     [SKIP] No snapshot found on $($selectedDate.ToString('yyyy-MM-dd'))." -ForegroundColor Yellow
            $skippedCount++
            continue
        }

        $snapshotId = $targetSnapshot.id
        $snapshotDt = ([datetime]$targetSnapshot.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss")
        Write-Host "     Snapshot : $snapshotId  (created $snapshotDt)" -ForegroundColor DarkGray

        # ── 3. Trigger the restore ────────────────────────────────────────────
        $body       = @{ cloudPcSnapshotId = $snapshotId } | ConvertTo-Json
        $uriRestore = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$cloudPcId/restore"

        try {
            Invoke-MgGraphRequest -Method POST -Uri $uriRestore -Body $body -ContentType "application/json"
            Write-Host "     [OK] Restore initiated successfully." -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "     [FAIL] $($_.Exception.Message)" -ForegroundColor Red
            $failCount++
        }
    }

    Write-Host "  ------------------------------------"
}

#endregion

#region --- Summary ---

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Done!" -ForegroundColor Cyan
Write-Host "  Restore date : $($selectedDate.ToString('dddd, MMMM d, yyyy'))" -ForegroundColor Cyan
Write-Host "  Total users  : $($userList.Count)" -ForegroundColor Cyan
Write-Host "  Success      : $successCount" -ForegroundColor Green
Write-Host "  Skipped      : $skippedCount" -ForegroundColor Yellow
Write-Host "  Failed       : $failCount" -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

#endregion
