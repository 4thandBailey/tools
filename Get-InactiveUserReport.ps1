#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Inactive User Report

.DESCRIPTION
    Identifies users with no sign-in activity within a specified threshold
    (30, 60, or 90 days) using the Microsoft Graph API. Reports last sign-in
    date, account status, assigned licenses, and department. Useful for
    license reclamation, security hygiene, and offboarding audits.

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

.PARAMETER InactiveDays
    Number of days of inactivity to flag. Default is 90.

.PARAMETER OutputPath
    Directory for output files. Defaults to current directory.

.PARAMETER HtmlReport
    Switch to generate an HTML report.

.PARAMETER ExcludeGuests
    Switch to exclude guest accounts from results.

.PARAMETER ExcludeDisabled
    Switch to exclude already-disabled accounts from results.

.EXAMPLE
    .\Get-InactiveUserReport.ps1 -InactiveDays 60 -HtmlReport -ExcludeGuests

.EXAMPLE
    .\Get-InactiveUserReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificatePath "/certs/app.pfx" `
        -CertificatePassword (Read-Host -AsSecureString "Password") `
        -InactiveDays 90 -HtmlReport -ExcludeGuests -ExcludeDisabled

.NOTES
    Author:      4TH AND BAILEY | Information Technology Consulting
                 4thandbailey.com — Where IT Works
    Version:     1.0.0
    GitHub:      https://github.com/4thandbailey
    Website:     https://4thandbailey.com

    Required Modules:
        Microsoft.Graph.Authentication
        Microsoft.Graph.Users

    Required Permissions:
        User.Read.All
        AuditLog.Read.All   (required for signInActivity)

    Install modules:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Users -Scope CurrentUser

    Note: signInActivity requires Entra ID P1 or P2 license in the tenant.
          Without it, LastSignInDateTime will return null for all users.
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
    [ValidateSet(30, 60, 90)]
    [int]$InactiveDays = 90,

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [switch]$HtmlReport,

    [Parameter()]
    [switch]$ExcludeGuests,

    [Parameter()]
    [switch]$ExcludeDisabled
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Banner ────────────────────────────────────────────────────────────────────
$version = '1.0.0'
Write-Host ''
Write-Host '  4TH AND BAILEY | Information Technology Consulting — Where IT Works' -ForegroundColor Cyan
Write-Host '  M365 Inactive User Report' -ForegroundColor Cyan
Write-Host "  Version $version  |  4thandbailey.com" -ForegroundColor DarkCyan
Write-Host ''

# ── Module check ──────────────────────────────────────────────────────────────
Write-Host '  Checking required modules...' -ForegroundColor Gray
foreach ($mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')) {
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
        Connect-MgGraph -Scopes 'User.Read.All', 'AuditLog.Read.All' -NoWelcome
    }
}

$ctx       = Get-MgContext
$tenant    = $ctx.TenantId
$threshold = (Get-Date).AddDays(-$InactiveDays)

Write-Host "  Connected  |  Tenant: $tenant" -ForegroundColor Green
Write-Host "  Inactivity threshold: $InactiveDays days (before $($threshold.ToString('yyyy-MM-dd')))" -ForegroundColor Gray

# ── Output path ───────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath   = Join-Path $OutputPath "InactiveUsers_${InactiveDays}d_$timestamp.csv"
$htmlPath  = Join-Path $OutputPath "InactiveUsers_${InactiveDays}d_$timestamp.html"

# ── Fetch users ───────────────────────────────────────────────────────────────
Write-Host '  Fetching users (this may take a moment for large tenants)...' -ForegroundColor Gray

$selectProps = 'id,displayName,userPrincipalName,userType,accountEnabled,' +
               'department,jobTitle,assignedLicenses,createdDateTime,signInActivity'

$allUsers = Get-MgUser -All -Property $selectProps -ConsistencyLevel eventual

Write-Host "  Retrieved $($allUsers.Count) users." -ForegroundColor Green

# ── Apply filters ─────────────────────────────────────────────────────────────
$filteredUsers = $allUsers

if ($ExcludeGuests) {
    $filteredUsers = @($filteredUsers | Where-Object { $_.UserType -ne 'Guest' })
    Write-Host "  After guest exclusion: $($filteredUsers.Count) users." -ForegroundColor Gray
}

if ($ExcludeDisabled) {
    $filteredUsers = @($filteredUsers | Where-Object { $_.AccountEnabled -eq $true })
    Write-Host "  After disabled exclusion: $($filteredUsers.Count) users." -ForegroundColor Gray
}

# ── Classify inactivity ───────────────────────────────────────────────────────
Write-Host '  Classifying activity status...' -ForegroundColor Gray

$results = foreach ($user in $filteredUsers) {
    $lastSignIn     = $null
    $lastSignInStr  = 'Never / No Data'
    $daysSinceLogin = $null
    $isInactive     = $false

    if ($null -ne $user.SignInActivity -and
        $null -ne $user.SignInActivity.LastSignInDateTime) {
        $lastSignIn     = $user.SignInActivity.LastSignInDateTime
        $lastSignInStr  = $lastSignIn.ToString('yyyy-MM-dd')
        $daysSinceLogin = [math]::Round(((Get-Date) - $lastSignIn).TotalDays, 0)
        $isInactive     = $lastSignIn -lt $threshold
    } else {
        # No sign-in data — treat as inactive
        $isInactive = $true
    }

    # Only return inactive users
    if (-not $isInactive) { continue }

    $licenseCount = if ($null -ne $user.AssignedLicenses) {
        @($user.AssignedLicenses).Count
    } else { 0 }

    [PSCustomObject]@{
        DisplayName       = if ($user.DisplayName)       { $user.DisplayName }       else { 'N/A' }
        UserPrincipalName = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { 'N/A' }
        UserType          = if ($user.UserType)          { $user.UserType }          else { 'Member' }
        AccountEnabled    = $user.AccountEnabled
        Department        = if ($user.Department)        { $user.Department }        else { 'N/A' }
        JobTitle          = if ($user.JobTitle)          { $user.JobTitle }          else { 'N/A' }
        LicenseCount      = $licenseCount
        LastSignIn        = $lastSignInStr
        DaysSinceSignIn   = if ($null -ne $daysSinceLogin) { $daysSinceLogin } else { 'N/A' }
        CreatedDate       = if ($null -ne $user.CreatedDateTime) {
                                $user.CreatedDateTime.ToString('yyyy-MM-dd')
                            } else { 'N/A' }
        InactiveThreshold = "$InactiveDays days"
        ReportDate        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

# Sort by days since sign-in descending (longest inactive first)
$results = @($results | Sort-Object {
    if ($_.DaysSinceSignIn -eq 'N/A') { 99999 }
    else { [int]$_.DaysSinceSignIn }
} -Descending)

# ── Summary ───────────────────────────────────────────────────────────────────
$totalInactive     = $results.Count
$neverSignedIn     = @($results | Where-Object { $_.LastSignIn -eq 'Never / No Data' }).Count
$licensedInactive  = @($results | Where-Object { $_.LicenseCount -gt 0 }).Count
$disabledInactive  = @($results | Where-Object { -not $_.AccountEnabled }).Count

# ── Export CSV ────────────────────────────────────────────────────────────────
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "  CSV exported  →  $csvPath" -ForegroundColor Green

Write-Host ''
Write-Host '  ── Summary ──────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host "  Inactive users ($InactiveDays+ days) : $totalInactive"
Write-Host "  Never signed in               : $neverSignedIn"
Write-Host "  Inactive WITH licenses        : $licensedInactive  ← review for reclamation"
Write-Host "  Inactive AND disabled         : $disabledInactive"
Write-Host ''

# ── HTML Report ───────────────────────────────────────────────────────────────
if ($HtmlReport) {
    Write-Host '  Generating HTML report...' -ForegroundColor Gray

    $tableRows = foreach ($r in $results) {
        $licClass  = if ($r.LicenseCount -gt 0) { 'warn' } else { '' }
        $dayClass  = if ($r.DaysSinceSignIn -eq 'N/A' -or [int]$r.DaysSinceSignIn -gt 180) { 'warn' } else { '' }
        $enabClass = if (-not $r.AccountEnabled) { 'caution' } else { '' }

        "<tr>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.DisplayName))</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.UserPrincipalName))</td>
          <td>$($r.UserType)</td>
          <td class='$enabClass'>$($r.AccountEnabled)</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.Department))</td>
          <td class='$licClass'>$($r.LicenseCount)</td>
          <td>$($r.LastSignIn)</td>
          <td class='$dayClass'>$($r.DaysSinceSignIn)</td>
          <td>$($r.CreatedDate)</td>
        </tr>"
    }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 Inactive User Report</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#1a1a1a;background:#f4f4f0}
.page{max-width:1200px;margin:0 auto;padding:24px 20px 60px}
.header{background:#0C447C;color:#fff;padding:28px 32px;border-radius:10px;margin-bottom:24px}
.header h1{font-size:20px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.header .meta{display:flex;gap:28px;margin-top:14px;flex-wrap:wrap}
.header .meta-item strong{display:block;opacity:.65;font-size:10px;text-transform:uppercase;letter-spacing:.06em}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px}
.metric{background:#fff;border-radius:8px;padding:14px 16px;border:1px solid #e0dfd8}
.metric-val{font-size:22px;font-weight:700;color:#0C447C}
.metric-lbl{font-size:11px;color:#888;margin-top:3px}
.section{background:#fff;border-radius:10px;padding:20px 24px;margin-bottom:20px;border:1px solid #e0dfd8}
.section h2{font-size:14px;font-weight:600;margin-bottom:14px;color:#0C447C;border-bottom:1px solid #eee;padding-bottom:8px}
.section-note{font-size:12px;color:#888;margin-bottom:10px;font-style:italic}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f4f4f0;text-align:left;padding:8px 10px;font-weight:600;color:#444;border-bottom:2px solid #e0dfd8;white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid #f0efea;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#fafaf8}
td.warn{color:#E24B4A;font-weight:600}
td.caution{color:#D48806;font-weight:600}
.footer{text-align:center;font-size:11px;color:#aaa;margin-top:40px}
.footer a{color:#0C447C;text-decoration:none}
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <h1>Microsoft 365 Inactive User Report</h1>
    <p>4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works</p>
    <p style="opacity:.5;margin-top:2px;font-size:11px">Microsoft Graph API v1.0</p>
    <div class="meta">
      <div class="meta-item"><strong>Tenant ID</strong>$tenant</div>
      <div class="meta-item"><strong>Threshold</strong>$InactiveDays days</div>
      <div class="meta-item"><strong>Generated</strong>$generatedAt</div>
    </div>
  </div>

  <div class="metrics">
    <div class="metric"><div class="metric-val">$totalInactive</div><div class="metric-lbl">Inactive users</div></div>
    <div class="metric"><div class="metric-val">$neverSignedIn</div><div class="metric-lbl">Never signed in</div></div>
    <div class="metric"><div class="metric-val">$licensedInactive</div><div class="metric-lbl">Licensed + inactive</div></div>
    <div class="metric"><div class="metric-val">$disabledInactive</div><div class="metric-lbl">Disabled + inactive</div></div>
  </div>

  <div class="section">
    <h2>Inactive Users — Longest Inactive First</h2>
    <p class="section-note">
      Users with no sign-in activity in the last $InactiveDays days.
      Red license count = active license on an inactive account (review for reclamation).
      signInActivity requires Entra ID P1 or P2.
    </p>
    <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Display Name</th><th>User Principal Name</th><th>Type</th>
          <th>Enabled</th><th>Department</th><th>Licenses</th>
          <th>Last Sign-In</th><th>Days Inactive</th><th>Created</th>
        </tr>
      </thead>
      <tbody>$($tableRows -join "`n")</tbody>
    </table>
    </div>
  </div>

  <div class="footer">
    Generated by 4TH AND BAILEY | Information Technology Consulting &bull;
    <a href="https://4thandbailey.com">4thandbailey.com</a> — Where IT Works &bull;
    Inactive User Report v$version &bull; Microsoft Graph API v1.0 &bull; $generatedAt
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
