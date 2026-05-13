#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 License Assignment Report

.DESCRIPTION
    Generates a per-user license assignment report across an M365 tenant
    using the Microsoft Graph API. Shows every user, their assigned SKUs,
    and available service plans. Identifies unlicensed users and unused
    license seats. Exports to CSV and optional HTML.

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
    Switch to also generate an HTML report.

.PARAMETER IncludeServicePlans
    Switch to include per-service-plan detail in CSV output.

.EXAMPLE
    .\Get-LicenseAssignmentReport.ps1 -HtmlReport

.EXAMPLE
    .\Get-LicenseAssignmentReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificateThumbprint "your-thumbprint" `
        -HtmlReport -IncludeServicePlans

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
        Organization.Read.All

    Install modules:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Users -Scope CurrentUser
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
    [switch]$IncludeServicePlans
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Banner ────────────────────────────────────────────────────────────────────
$version = '1.0.0'
Write-Host ''
Write-Host '  4TH AND BAILEY | Information Technology Consulting — Where IT Works' -ForegroundColor Cyan
Write-Host '  M365 License Assignment Report' -ForegroundColor Cyan
Write-Host "  Version $version  |  4thandbailey.com" -ForegroundColor DarkCyan
Write-Host ''

# ── Friendly SKU name map (common SKUs) ───────────────────────────────────────
$skuNames = @{
    'SPB'                                           = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'                         = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'                      = 'Microsoft 365 Business Basic'
    'Microsoft_365_Business_Premium_(no_Teams)'     = 'M365 Business Premium (no Teams)'
    'ENTERPRISEPACK'                                = 'Microsoft 365 E3'
    'SPE_E5'                                        = 'Microsoft 365 E5'
    'EXCHANGESTANDARD'                              = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'                            = 'Exchange Online Plan 2'
    'AAD_PREMIUM'                                   = 'Entra ID P1'
    'AAD_PREMIUM_P2'                                = 'Entra ID P2'
    'MCOEV'                                         = 'Teams Phone'
    'MCOMEETADV'                                    = 'Audio Conferencing'
    'PROJECTPREMIUM'                                = 'Project Plan 5'
    'PBI_PREMIUM_PER_USER'                          = 'Power BI Premium Per User'
    'POWER_BI_STANDARD'                             = 'Power BI (free)'
    'FLOW_FREE'                                     = 'Power Automate (free)'
    'POWERAPPS_VIRAL'                               = 'Power Apps (free)'
    'Microsoft_Teams_Enterprise_New'                = 'Microsoft Teams Enterprise'
    'Microsoft365_Lighthouse'                       = 'Microsoft 365 Lighthouse'
    'POWERAUTOMATE_ATTENDED_RPA'                    = 'Power Automate RPA'
}

function Get-FriendlySkuName ([string]$SkuPartNumber) {
    if ($skuNames.ContainsKey($SkuPartNumber)) { return $skuNames[$SkuPartNumber] }
    return $SkuPartNumber
}

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
        Connect-MgGraph -Scopes 'User.Read.All', 'Organization.Read.All' -NoWelcome
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
$csvPath   = Join-Path $OutputPath "LicenseAssignment_$timestamp.csv"
$htmlPath  = Join-Path $OutputPath "LicenseAssignment_$timestamp.html"

# ── Fetch tenant subscribed SKUs ──────────────────────────────────────────────
Write-Host '  Fetching tenant subscribed SKUs...' -ForegroundColor Gray

$skuResponse    = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus'
$subscribedSkus = @($skuResponse.value)   # @() forces array — prevents Count failure under StrictMode

$skuSeatMap = @{}
foreach ($sku in $subscribedSkus) {
    $total     = if ($null -ne $sku.prepaidUnits -and $null -ne $sku.prepaidUnits.enabled) { [int]$sku.prepaidUnits.enabled } else { 0 }
    $consumed  = if ($null -ne $sku.consumedUnits) { [int]$sku.consumedUnits } else { 0 }
    $skuSeatMap[$sku.skuId] = @{
        SkuPartNumber = $sku.skuPartNumber
        FriendlyName  = Get-FriendlySkuName $sku.skuPartNumber
        Total         = $total
        Used          = $consumed
        Available     = $total - $consumed
    }
}

Write-Host "  Found $($subscribedSkus.Count) subscribed SKUs." -ForegroundColor Green

# ── Fetch all users with license detail ───────────────────────────────────────
Write-Host '  Fetching users and license assignments (this may take a moment)...' -ForegroundColor Gray

$selectProps = 'id,displayName,userPrincipalName,userType,accountEnabled,' +
               'assignedLicenses,licenseAssignmentStates,createdDateTime,signInActivity'

$users = Get-MgUser -All -Property $selectProps -ConsistencyLevel eventual

Write-Host "  Retrieved $($users.Count) users." -ForegroundColor Green

# ── Process users ─────────────────────────────────────────────────────────────
Write-Host '  Processing license assignments...' -ForegroundColor Gray

$results = foreach ($user in $users) {
    $licenses     = @(if ($null -ne $user.AssignedLicenses) { $user.AssignedLicenses } else { @() })
    $licenseCount = @($licenses).Count
    $isLicensed   = $licenseCount -gt 0

    $skuList = if ($isLicensed) {
        ($licenses | ForEach-Object {
            $skuId = if ($null -ne $_.SkuId) { $_.SkuId } else { '' }
            if ($skuId -ne '' -and $skuSeatMap.ContainsKey($skuId)) {
                $skuSeatMap[$skuId].FriendlyName
            } elseif ($skuId -ne '') { $skuId }
            else { 'Unknown' }
        }) -join '; '
    } else { 'Unlicensed' }

    $lastSignIn = 'N/A'
    try {
        if ($null -ne $user.SignInActivity -and
            $null -ne $user.SignInActivity.LastSignInDateTime) {
            $lastSignIn = $user.SignInActivity.LastSignInDateTime.ToString('yyyy-MM-dd')
        }
    } catch { $lastSignIn = 'N/A' }

    $row = [PSCustomObject]@{
        DisplayName       = if ($null -ne $user.DisplayName -and $user.DisplayName -ne '')       { $user.DisplayName }       else { 'N/A' }
        UserPrincipalName = if ($null -ne $user.UserPrincipalName -and $user.UserPrincipalName -ne '') { $user.UserPrincipalName } else { 'N/A' }
        UserType          = if ($null -ne $user.UserType -and $user.UserType -ne '')          { $user.UserType }          else { 'Member' }
        AccountEnabled    = if ($null -ne $user.AccountEnabled) { $user.AccountEnabled } else { $false }
        IsLicensed        = $isLicensed
        LicenseCount      = $licenseCount
        AssignedLicenses  = $skuList
        LastSignIn        = $lastSignIn
        CreatedDate       = if ($null -ne $user.CreatedDateTime) {
                                $user.CreatedDateTime.ToString('yyyy-MM-dd')
                            } else { 'N/A' }
        ReportDate        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }

    if ($IncludeServicePlans -and $isLicensed) {
        $plans = foreach ($lic in $licenses) {
            $skuId = if ($null -ne $lic.SkuId) { $lic.SkuId } else { '' }
            if ($skuId -ne '' -and $skuSeatMap.ContainsKey($skuId)) {
                $skuSeatMap[$skuId].SkuPartNumber
            }
        }
        $row | Add-Member -MemberType NoteProperty -Name 'SkuPartNumbers' -Value ($plans -join '; ')
    }

    $row
}

# ── Summary ───────────────────────────────────────────────────────────────────
$totalUsers      = @($results).Count
$licensedUsers   = @($results | Where-Object { $_.IsLicensed }).Count
$unlicensedUsers = @($results | Where-Object { -not $_.IsLicensed }).Count
$guestUsers      = @($results | Where-Object { $_.UserType -eq 'Guest' }).Count
$disabledUsers   = @($results | Where-Object { -not $_.AccountEnabled }).Count

# ── Export CSV ────────────────────────────────────────────────────────────────
$results | Sort-Object DisplayName |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "  CSV exported  →  $csvPath" -ForegroundColor Green

Write-Host ''
Write-Host '  ── Summary ──────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host "  Total users       : $totalUsers"
Write-Host "  Licensed          : $licensedUsers"
Write-Host "  Unlicensed        : $unlicensedUsers"
Write-Host "  Guest users       : $guestUsers"
Write-Host "  Disabled accounts : $disabledUsers"
Write-Host ''
Write-Host '  ── Tenant SKU Inventory ─────────────────────────' -ForegroundColor DarkCyan
foreach ($skuId in $skuSeatMap.Keys) {
    $s = $skuSeatMap[$skuId]
    if ($s.Total -gt 0 -and $s.Total -lt 10000) {
        $availColor = if ($s.Available -le 0) { 'Red' } else { 'Gray' }
        Write-Host ("  {0,-45} Total:{1,6}  Used:{2,6}  Avail:{3,6}" -f `
            $s.FriendlyName, $s.Total, $s.Used, $s.Available) -ForegroundColor $availColor
    }
}
Write-Host ''

# ── HTML Report ───────────────────────────────────────────────────────────────
if ($HtmlReport) {
    Write-Host '  Generating HTML report...' -ForegroundColor Gray

    $userRows = foreach ($r in ($results | Sort-Object DisplayName)) {
        $licClass  = if (-not $r.IsLicensed) { 'warn' } else { '' }
        $enabClass = if (-not $r.AccountEnabled) { 'warn' } else { '' }

        "<tr>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.DisplayName))</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.UserPrincipalName))</td>
          <td>$($r.UserType)</td>
          <td class='$enabClass'>$($r.AccountEnabled)</td>
          <td class='$licClass'>$($r.IsLicensed)</td>
          <td>$($r.LicenseCount)</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.AssignedLicenses))</td>
          <td>$($r.LastSignIn)</td>
        </tr>"
    }

    $skuRows = foreach ($skuId in ($skuSeatMap.Keys | Sort-Object)) {
        $s = $skuSeatMap[$skuId]
        if ($s.Total -lt 10000) {
            $availClass = if ($s.Available -le 0) { 'warn' } else { '' }
            "<tr>
              <td>$([System.Web.HttpUtility]::HtmlEncode($s.FriendlyName))</td>
              <td>$([System.Web.HttpUtility]::HtmlEncode($s.SkuPartNumber))</td>
              <td>$($s.Total)</td>
              <td>$($s.Used)</td>
              <td class='$availClass'>$($s.Available)</td>
            </tr>"
        }
    }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 License Assignment Report</title>
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
.section{background:#fff;border-radius:10px;padding:20px 24px;margin-bottom:20px;border:1px solid #e0dfd8}
.section h2{font-size:14px;font-weight:600;margin-bottom:14px;color:#0C447C;border-bottom:1px solid #eee;padding-bottom:8px}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f4f4f0;text-align:left;padding:8px 10px;font-weight:600;color:#444;border-bottom:2px solid #e0dfd8;white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid #f0efea;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#fafaf8}
td.warn{color:#E24B4A;font-weight:600}
.footer{text-align:center;font-size:11px;color:#aaa;margin-top:40px}
.footer a{color:#0C447C;text-decoration:none}
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <h1>Microsoft 365 License Assignment Report</h1>
    <p>4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works</p>
    <p style="opacity:.5;margin-top:2px;font-size:11px">Microsoft Graph API v1.0</p>
    <div class="meta">
      <div class="meta-item"><strong>Tenant ID</strong>$tenant</div>
      <div class="meta-item"><strong>Generated</strong>$generatedAt</div>
    </div>
  </div>

  <div class="metrics">
    <div class="metric"><div class="metric-val">$totalUsers</div><div class="metric-lbl">Total users</div></div>
    <div class="metric"><div class="metric-val">$licensedUsers</div><div class="metric-lbl">Licensed</div></div>
    <div class="metric"><div class="metric-val">$unlicensedUsers</div><div class="metric-lbl">Unlicensed</div></div>
    <div class="metric"><div class="metric-val">$guestUsers</div><div class="metric-lbl">Guest users</div></div>
    <div class="metric"><div class="metric-val">$disabledUsers</div><div class="metric-lbl">Disabled accounts</div></div>
  </div>

  <div class="section">
    <h2>Tenant License Inventory</h2>
    <div class="table-wrap">
    <table>
      <thead><tr><th>License</th><th>SKU Part Number</th><th>Total</th><th>Used</th><th>Available</th></tr></thead>
      <tbody>$($skuRows -join "`n")</tbody>
    </table>
    </div>
  </div>

  <div class="section">
    <h2>User License Assignments</h2>
    <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Display Name</th><th>User Principal Name</th><th>Type</th>
          <th>Enabled</th><th>Licensed</th><th>Count</th><th>Licenses</th><th>Last Sign-In</th>
        </tr>
      </thead>
      <tbody>$($userRows -join "`n")</tbody>
    </table>
    </div>
  </div>

  <div class="footer">
    Generated by 4TH AND BAILEY | Information Technology Consulting &bull;
    <a href="https://4thandbailey.com">4thandbailey.com</a> — Where IT Works &bull;
    License Assignment Report v$version &bull; Microsoft Graph API v1.0 &bull; $generatedAt
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
