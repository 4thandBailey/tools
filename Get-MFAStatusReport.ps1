#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 MFA Status & Authentication Methods Report

.DESCRIPTION
    Generates a per-user MFA status and registered authentication methods
    report across an M365 tenant using the Microsoft Graph API. Identifies
    users without MFA, users relying on legacy SMS/voice only, and accounts
    with no authentication methods registered.

    Compatible with PowerShell 7.0+ on Windows, macOS, and Linux.

.PARAMETER TenantId
    The Entra ID Tenant ID. Required for app-only authentication.

.PARAMETER ClientId
    The App Registration Client ID. Required for app-only authentication.

.PARAMETER CertificateThumbprint
    Certificate thumbprint. Windows only.

.PARAMETER CertificatePath
    Path to a .pfx certificate file. macOS/Linux compatible.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate file.

.PARAMETER OutputPath
    Directory for output files. Defaults to current directory.

.PARAMETER HtmlReport
    Switch to generate an HTML report.

.PARAMETER LicensedUsersOnly
    Switch to include only licensed users in the report.

.PARAMETER ExcludeGuests
    Switch to exclude guest accounts.

.EXAMPLE
    .\Get-MFAStatusReport.ps1 -HtmlReport -LicensedUsersOnly -ExcludeGuests

.EXAMPLE
    .\Get-MFAStatusReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificateThumbprint "your-thumbprint" `
        -HtmlReport -LicensedUsersOnly

.NOTES
    Author:      4TH AND BAILEY | Information Technology Consulting
                 4thandbailey.com — Where IT Works
    Version:     1.0.0
    GitHub:      https://github.com/4thandbailey
    Website:     https://4thandbailey.com

    Required Modules:
        Microsoft.Graph.Authentication
        Microsoft.Graph.Users
        Microsoft.Graph.Identity.SignIns

    Required Permissions:
        User.Read.All
        UserAuthenticationMethod.Read.All

    Install modules:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Users -Scope CurrentUser
        Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser

    Note: UserAuthenticationMethod.Read.All requires Global Admin or
          Authentication Administrator role in the tenant for delegated
          access. For app-only access it requires the same permission
          granted via admin consent.
#>

[CmdletBinding(DefaultParameterSetName = 'Delegated')]
param (
    [Parameter(ParameterSetName = 'AppCert',     Mandatory)]
    [Parameter(ParameterSetName = 'AppCertFile', Mandatory)]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'AppCert',     Mandatory)]
    [Parameter(ParameterSetName = 'AppCertFile', Mandatory)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppCert', Mandatory)]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'AppCertFile', Mandatory)]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'AppCertFile')]
    [securestring]$CertificatePassword,

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [switch]$HtmlReport,

    [Parameter()]
    [switch]$LicensedUsersOnly,

    [Parameter()]
    [switch]$ExcludeGuests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Banner ────────────────────────────────────────────────────────────────────
$version = '1.0.0'
Write-Host ''
Write-Host '  4TH AND BAILEY | Information Technology Consulting — Where IT Works' -ForegroundColor Cyan
Write-Host '  M365 MFA Status & Authentication Methods Report' -ForegroundColor Cyan
Write-Host "  Version $version  |  4thandbailey.com" -ForegroundColor DarkCyan
Write-Host ''

# ── Method type labels ────────────────────────────────────────────────────────
$methodLabels = @{
    'microsoftAuthenticatorAuthenticationMethod' = 'Microsoft Authenticator'
    'phoneAuthenticationMethod'                  = 'Phone (SMS/Voice)'
    'emailAuthenticationMethod'                  = 'Email OTP'
    'fido2AuthenticationMethod'                  = 'FIDO2 / Passkey'
    'windowsHelloForBusinessAuthenticationMethod'= 'Windows Hello'
    'softwareOathAuthenticationMethod'           = 'TOTP App (OATH)'
    'temporaryAccessPassAuthenticationMethod'    = 'Temporary Access Pass'
    'passwordAuthenticationMethod'               = 'Password'
}

function Get-MethodLabel ([string]$OdataType) {
    $key = $OdataType -replace '#microsoft.graph.',''
    if ($methodLabels.ContainsKey($key)) { return $methodLabels[$key] }
    return $key
}

# ── Module check ──────────────────────────────────────────────────────────────
Write-Host '  Checking required modules...' -ForegroundColor Gray
$requiredMods = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.SignIns'
)
foreach ($mod in $requiredMods) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Error "Module '$mod' not installed. Run: Install-Module $mod -Scope CurrentUser"
    }
    Import-Module $mod -ErrorAction Stop
}
Write-Host '  Modules loaded.' -ForegroundColor Green

# ── Authentication ────────────────────────────────────────────────────────────
Write-Host '  Authenticating to Microsoft Graph...' -ForegroundColor Gray

switch ($PSCmdlet.ParameterSetName) {
    'AppCert' {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    'AppCertFile' {
        $certBytes = [System.IO.File]::ReadAllBytes($CertificatePath)
        $x509      = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                         $certBytes, $CertificatePassword,
                         [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                     )
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -Certificate $x509 -NoWelcome
    }
    default {
        Connect-MgGraph -Scopes 'User.Read.All', 'UserAuthenticationMethod.Read.All' -NoWelcome
    }
}

$ctx    = Get-MgContext
$tenant = $ctx.TenantId
Write-Host "  Connected  |  Tenant: $tenant" -ForegroundColor Green

# ── Output path ───────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath   = Join-Path $OutputPath "MFAStatus_$timestamp.csv"
$htmlPath  = Join-Path $OutputPath "MFAStatus_$timestamp.html"

# ── Fetch users ───────────────────────────────────────────────────────────────
Write-Host '  Fetching users...' -ForegroundColor Gray

$selectProps = 'id,displayName,userPrincipalName,userType,accountEnabled,' +
               'department,assignedLicenses,createdDateTime'

$allUsers = Get-MgUser -All -Property $selectProps -ConsistencyLevel eventual

Write-Host "  Retrieved $($allUsers.Count) users." -ForegroundColor Green

# ── Apply filters ─────────────────────────────────────────────────────────────
$filteredUsers = $allUsers

if ($ExcludeGuests) {
    $filteredUsers = @($filteredUsers | Where-Object { $_.UserType -ne 'Guest' })
}

if ($LicensedUsersOnly) {
    $filteredUsers = @($filteredUsers | Where-Object {
        $null -ne $_.AssignedLicenses -and @($_.AssignedLicenses).Count -gt 0
    })
}

Write-Host "  Processing $($filteredUsers.Count) users after filters." -ForegroundColor Green

# ── Fetch authentication methods per user ─────────────────────────────────────
Write-Host '  Fetching authentication methods (this takes time for large tenants)...' -ForegroundColor Gray

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$processed = 0

foreach ($user in $filteredUsers) {
    $processed++
    if ($processed % 25 -eq 0) {
        Write-Host "  Processing user $processed of $($filteredUsers.Count)..." -ForegroundColor Gray
    }

    $methods       = @()
    $methodNames   = @()
    $hasMFA        = $false
    $hasAuthApp    = $false
    $hasFido2      = $false
    $hasPhone      = $false
    $hasLegacyOnly = $false
    $methodCount   = 0

    try {
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop

        foreach ($method in $authMethods) {
            $odataType = $method.AdditionalProperties['@odata.type']
            if ($null -ne $odataType) {
                $label = Get-MethodLabel $odataType
                $methodNames += $label

                # Identify MFA-capable methods (exclude password-only)
                if ($odataType -notlike '*password*') {
                    $hasMFA = $true
                }
                if ($odataType -like '*microsoftAuthenticator*') { $hasAuthApp = $true }
                if ($odataType -like '*fido2*')                  { $hasFido2  = $true }
                if ($odataType -like '*phone*')                  { $hasPhone  = $true }
            }
        }

        $methodCount = $methodNames.Count
        # Legacy only = has phone/email but NOT authenticator app or FIDO2
        $hasLegacyOnly = $hasPhone -and -not $hasAuthApp -and -not $hasFido2

    } catch {
        $methodNames = @('Error retrieving methods')
    }

    $mfaStatus = if (-not $hasMFA) { 'No MFA' }
                 elseif ($hasAuthApp -or $hasFido2) { 'Strong MFA' }
                 elseif ($hasLegacyOnly) { 'Legacy MFA Only' }
                 else { 'MFA Registered' }

    $licenseCount = if ($null -ne $user.AssignedLicenses) {
        @($user.AssignedLicenses).Count
    } else { 0 }

    $results.Add([PSCustomObject]@{
        DisplayName         = if ($user.DisplayName)       { $user.DisplayName }       else { 'N/A' }
        UserPrincipalName   = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { 'N/A' }
        UserType            = if ($user.UserType)          { $user.UserType }          else { 'Member' }
        AccountEnabled      = $user.AccountEnabled
        Department          = if ($user.Department)        { $user.Department }        else { 'N/A' }
        LicenseCount        = $licenseCount
        MFAStatus           = $mfaStatus
        HasAuthenticatorApp = $hasAuthApp
        HasFIDO2            = $hasFido2
        HasPhoneSMS         = $hasPhone
        MethodCount         = $methodCount
        RegisteredMethods   = ($methodNames -join '; ')
        CreatedDate         = if ($null -ne $user.CreatedDateTime) {
                                  $user.CreatedDateTime.ToString('yyyy-MM-dd')
                              } else { 'N/A' }
        ReportDate          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

# Sort: No MFA first, then Legacy, then Strong
$sortOrder = @{ 'No MFA' = 0; 'Legacy MFA Only' = 1; 'MFA Registered' = 2; 'Strong MFA' = 3 }
$results   = [System.Collections.Generic.List[PSCustomObject]]($results |
    Sort-Object { $sortOrder[$_.MFAStatus] }, DisplayName)

# ── Summary ───────────────────────────────────────────────────────────────────
$totalUsers   = $results.Count
$noMFA        = @($results | Where-Object { $_.MFAStatus -eq 'No MFA' }).Count
$legacyOnly   = @($results | Where-Object { $_.MFAStatus -eq 'Legacy MFA Only' }).Count
$strongMFA    = @($results | Where-Object { $_.MFAStatus -eq 'Strong MFA' }).Count
$mfaReg       = @($results | Where-Object { $_.MFAStatus -eq 'MFA Registered' }).Count
$mfaCoverage  = if ($totalUsers -gt 0) {
    [math]::Round((($totalUsers - $noMFA) / $totalUsers) * 100, 1)
} else { 0 }

# ── Export CSV ────────────────────────────────────────────────────────────────
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "  CSV exported  →  $csvPath" -ForegroundColor Green

Write-Host ''
Write-Host '  ── Summary ──────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host "  Total users       : $totalUsers"
Write-Host "  No MFA            : $noMFA  ← immediate action required" -ForegroundColor $(if ($noMFA -gt 0) { 'Red' } else { 'Gray' })
Write-Host "  Legacy MFA only   : $legacyOnly  ← review and upgrade" -ForegroundColor $(if ($legacyOnly -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host "  MFA Registered    : $mfaReg"
Write-Host "  Strong MFA        : $strongMFA  (Authenticator App or FIDO2)"
Write-Host "  MFA Coverage      : $mfaCoverage%"
Write-Host ''

# ── HTML Report ───────────────────────────────────────────────────────────────
if ($HtmlReport) {
    Write-Host '  Generating HTML report...' -ForegroundColor Gray

    $tableRows = foreach ($r in $results) {
        $statusClass = switch ($r.MFAStatus) {
            'No MFA'          { 'status-none' }
            'Legacy MFA Only' { 'status-legacy' }
            'Strong MFA'      { 'status-strong' }
            default           { 'status-reg' }
        }
        $enabClass = if (-not $r.AccountEnabled) { 'warn' } else { '' }

        "<tr>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.DisplayName))</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.UserPrincipalName))</td>
          <td>$($r.UserType)</td>
          <td class='$enabClass'>$($r.AccountEnabled)</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.Department))</td>
          <td>$($r.LicenseCount)</td>
          <td><span class='status-badge $statusClass'>$($r.MFAStatus)</span></td>
          <td>$($r.MethodCount)</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.RegisteredMethods))</td>
        </tr>"
    }

    $coverageColor = if ($mfaCoverage -ge 95) { '#1D9E75' }
                     elseif ($mfaCoverage -ge 75) { '#D48806' }
                     else { '#E24B4A' }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 MFA Status Report</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#1a1a1a;background:#f4f4f0}
.page{max-width:1200px;margin:0 auto;padding:24px 20px 60px}
.header{background:#0C447C;color:#fff;padding:28px 32px;border-radius:10px;margin-bottom:24px}
.header h1{font-size:20px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.header .meta{display:flex;gap:28px;margin-top:14px;flex-wrap:wrap}
.header .meta-item strong{display:block;opacity:.65;font-size:10px;text-transform:uppercase;letter-spacing:.06em}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;margin-bottom:24px}
.metric{background:#fff;border-radius:8px;padding:14px 16px;border:1px solid #e0dfd8}
.metric-val{font-size:22px;font-weight:700;color:#0C447C}
.metric-lbl{font-size:11px;color:#888;margin-top:3px}
.metric-val.danger{color:#E24B4A}
.metric-val.warning{color:#D48806}
.metric-val.success{color:#1D9E75}
.section{background:#fff;border-radius:10px;padding:20px 24px;margin-bottom:20px;border:1px solid #e0dfd8}
.section h2{font-size:14px;font-weight:600;margin-bottom:14px;color:#0C447C;border-bottom:1px solid #eee;padding-bottom:8px}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f4f4f0;text-align:left;padding:8px 10px;font-weight:600;color:#444;border-bottom:2px solid #e0dfd8;white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid #f0efea;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#fafaf8}
td.warn{color:#E24B4A;font-weight:600}
.status-badge{font-size:11px;padding:2px 9px;border-radius:999px;font-weight:600;display:inline-block}
.status-none{background:#FEE2E2;color:#991B1B}
.status-legacy{background:#FAEEDA;color:#633806}
.status-reg{background:#E6F1FB;color:#0C447C}
.status-strong{background:#E1F5EE;color:#085041}
.coverage-bar-track{background:#e0dfd8;border-radius:999px;height:8px;margin-top:8px;width:100%;max-width:300px}
.coverage-bar{border-radius:999px;height:8px;background:$coverageColor}
.footer{text-align:center;font-size:11px;color:#aaa;margin-top:40px}
.footer a{color:#0C447C;text-decoration:none}
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <h1>Microsoft 365 MFA Status &amp; Authentication Methods Report</h1>
    <p>4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works</p>
    <p style="opacity:.5;margin-top:2px;font-size:11px">Microsoft Graph API v1.0</p>
    <div class="meta">
      <div class="meta-item"><strong>Tenant ID</strong>$tenant</div>
      <div class="meta-item"><strong>Generated</strong>$generatedAt</div>
    </div>
  </div>

  <div class="metrics">
    <div class="metric">
      <div class="metric-val" style="color:$coverageColor">$mfaCoverage%</div>
      <div class="metric-lbl">MFA coverage</div>
      <div class="coverage-bar-track"><div class="coverage-bar" style="width:$mfaCoverage%"></div></div>
    </div>
    <div class="metric"><div class="metric-val$(if ($noMFA -gt 0) {' danger'} else {''})">$noMFA</div><div class="metric-lbl">No MFA — act now</div></div>
    <div class="metric"><div class="metric-val$(if ($legacyOnly -gt 0) {' warning'} else {''})">$legacyOnly</div><div class="metric-lbl">Legacy MFA only</div></div>
    <div class="metric"><div class="metric-val success">$strongMFA</div><div class="metric-lbl">Strong MFA</div></div>
    <div class="metric"><div class="metric-val">$mfaReg</div><div class="metric-lbl">MFA registered</div></div>
    <div class="metric"><div class="metric-val">$totalUsers</div><div class="metric-lbl">Total users</div></div>
  </div>

  <div class="section">
    <h2>User MFA Status — No MFA and Legacy First</h2>
    <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Display Name</th><th>User Principal Name</th><th>Type</th>
          <th>Enabled</th><th>Department</th><th>Licenses</th>
          <th>MFA Status</th><th>Methods</th><th>Registered Methods</th>
        </tr>
      </thead>
      <tbody>$($tableRows -join "`n")</tbody>
    </table>
    </div>
  </div>

  <div class="footer">
    Generated by 4TH AND BAILEY | Information Technology Consulting &bull;
    <a href="https://4thandbailey.com">4thandbailey.com</a> — Where IT Works &bull;
    MFA Status Report v$version &bull; Microsoft Graph API v1.0 &bull; $generatedAt
  </div>
</div>
</body>
</html>
"@

    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "  HTML exported →  $htmlPath" -ForegroundColor Green
}

# ── Disconnect ────────────────────────────────────────────────────────────────
Disconnect-MgGraph | Out-Null
Write-Host '  Disconnected from Microsoft Graph.' -ForegroundColor Gray
Write-Host '  Done.' -ForegroundColor Cyan
Write-Host ''
