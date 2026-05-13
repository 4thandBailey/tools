#Requires -Version 7.0
<#
.SYNOPSIS
    Exchange Online Mailbox Statistics Report

.DESCRIPTION
    Generates a detailed mailbox statistics report for all mailboxes in an
    Exchange Online tenant using the Microsoft Graph API. Exports results to
    CSV and optionally to a branded HTML report.

    Compatible with PowerShell 7.0+ on Windows, macOS, and Linux.

.PARAMETER TenantId
    The Entra ID Tenant ID (GUID). Required for app-only authentication.

.PARAMETER ClientId
    The App Registration Client ID. Required for app-only authentication.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication. Windows only.

.PARAMETER CertificatePath
    Path to a .pfx certificate file. macOS/Linux compatible alternative
    to CertificateThumbprint.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate file.

.PARAMETER OutputPath
    Directory where output files will be saved.
    Defaults to the current directory.

.PARAMETER HtmlReport
    Switch to generate an HTML report in addition to CSV.

.PARAMETER Period
    Reporting period in days. Valid values: 7, 30, 90, 180.
    Defaults to 30.

.EXAMPLE
    # Interactive (delegated) authentication
    .\Get-MailboxStatisticsReport.ps1 -HtmlReport

.EXAMPLE
    # App-only authentication (Windows - certificate thumbprint)
    .\Get-MailboxStatisticsReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificateThumbprint "your-thumbprint" `
        -HtmlReport

.EXAMPLE
    # App-only authentication (macOS/Linux - certificate file)
    .\Get-MailboxStatisticsReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificatePath "/path/to/cert.pfx" `
        -CertificatePassword (Read-Host -AsSecureString "Cert Password") `
        -OutputPath "/tmp/reports"

.NOTES
    Author:      4TH AND BAILEY | Information Technology Consulting
                 4thandbailey.com — Where IT Works
    Version:     1.0.0
    GitHub:      https://github.com/4thandbailey
    Website:     https://4thandbailey.com

    Required Modules:
        Microsoft.Graph.Authentication
        Microsoft.Graph.Reports

    Required Permissions (app-only):
        Reports.Read.All
        User.Read.All

    Required Permissions (delegated):
        Reports.Read.All
        User.Read.All

    Install modules:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Reports -Scope CurrentUser
#>

[CmdletBinding(DefaultParameterSetName = 'Delegated')]
param (
    [Parameter(ParameterSetName = 'AppCert',   Mandatory)]
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
    [ValidateSet(7, 30, 90, 180)]
    [int]$Period = 30,

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [switch]$HtmlReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Banner ──────────────────────────────────────────────────────────────────
$version = '1.0.0'
Write-Host ''
Write-Host '  4TH AND BAILEY | Information Technology Consulting — Where IT Works' -ForegroundColor Cyan
Write-Host '  Exchange Online Mailbox Statistics Report' -ForegroundColor Cyan
Write-Host "  Version $version  |  4thandbailey.com" -ForegroundColor DarkCyan
Write-Host ''

# ── Helper: Safe property access ─────────────────────────────────────────────
function Get-SafeValue {
    param($Object, [string]$Property, $Default = 'N/A')
    if ($null -eq $Object) { return $Default }
    $val = $Object.$Property
    if ($null -eq $val -or $val -eq '') { return $Default }
    return $val
}

# ── Helper: Format bytes to human-readable ───────────────────────────────────
function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# ── Module check ─────────────────────────────────────────────────────────────
Write-Host '  Checking required modules...' -ForegroundColor Gray
foreach ($mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Error "Module '$mod' is not installed. Run: Install-Module $mod -Scope CurrentUser"
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
        # Cross-platform: load cert from .pfx file
        $certBytes  = [System.IO.File]::ReadAllBytes($CertificatePath)
        $x509       = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                          $certBytes, $CertificatePassword,
                          [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                      )
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -Certificate $x509 -NoWelcome
    }
    default {
        # Delegated / interactive
        Connect-MgGraph -Scopes 'Reports.Read.All', 'User.Read.All' -NoWelcome
    }
}

$ctx    = Get-MgContext
$tenant = $ctx.TenantId
Write-Host "  Connected  |  Tenant: $tenant" -ForegroundColor Green

# ── Output path ───────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath    = Join-Path $OutputPath "MailboxStatistics_$timestamp.csv"
$htmlPath   = Join-Path $OutputPath "MailboxStatistics_$timestamp.html"
$tempCsv    = Join-Path ([System.IO.Path]::GetTempPath()) "4ab_mbox_$timestamp.csv"

# ── Fetch mailbox usage detail ────────────────────────────────────────────────
Write-Host "  Fetching mailbox usage detail (D$Period)..." -ForegroundColor Gray
$periodParam = "D$Period"

# Suppress Graph SDK progress bar — on macOS/Linux it throws a harmless
# but noisy PercentComplete overflow error (int32 max value bug in SDK)
$savedProgress      = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'

try {
    # Direct REST call — avoids the PercentComplete SDK bug and
    # requests non-obfuscated display names for tenants with privacy settings
    $reportUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$periodParam')"
    Invoke-MgGraphRequest -Method GET -Uri $reportUri `
        -Headers @{ 'Accept' = 'application/octet-stream' } `
        -OutputFilePath $tempCsv -ErrorAction Stop
} catch {
    # Fallback to cmdlet if direct REST call fails
    try {
        Get-MgReportMailboxUsageDetail -Period $periodParam -OutFile $tempCsv -ErrorAction Stop
    } catch {
        Write-Error "Failed to retrieve mailbox usage report: $_"
    }
} finally {
    $ProgressPreference = $savedProgress
}

if (-not (Test-Path $tempCsv) -or (Get-Item $tempCsv).Length -eq 0) {
    Write-Error 'Report file is empty. Verify Reports.Read.All permission is granted.'
}

$raw = Import-Csv -Path $tempCsv
Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue

# Emit actual column names at -Verbose for diagnostics
$detectedHeaders = if (@($raw).Count -gt 0) { $raw[0].PSObject.Properties.Name } else { @() }
Write-Verbose "  Columns detected: $($detectedHeaders -join ' | ')"
Write-Host "  Columns: $($detectedHeaders -join ' | ')" -ForegroundColor DarkGray

Write-Host "  Retrieved $(@($raw).Count) mailbox records." -ForegroundColor Green

# ── Detect actual CSV column names (Graph API column names vary by tenant/locale) ──
Write-Host '  Processing records...' -ForegroundColor Gray

$headers = $raw[0].PSObject.Properties.Name
Write-Verbose "  CSV columns detected: $($headers -join ', ')"

# Build a flexible column resolver — finds the best match for each logical field
function Resolve-Column {
    param([string[]]$Headers, [string[]]$Candidates)
    foreach ($c in $Candidates) {
        $match = $Headers | Where-Object { $_ -eq $c } | Select-Object -First 1
        if ($match) { return $match }
    }
    # Fuzzy fallback — partial match
    foreach ($c in $Candidates) {
        $match = $Headers | Where-Object { $_ -like "*$c*" } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

$colDisplayName     = Resolve-Column $headers @('Display Name', 'DisplayName')
$colUPN             = Resolve-Column $headers @('User Principal Name', 'UserPrincipalName')
$colRecipientType   = Resolve-Column $headers @('Recipient Type', 'RecipientType', 'Recipient type')
$colItemCount       = Resolve-Column $headers @('Item Count', 'ItemCount')
$colStorage         = Resolve-Column $headers @('Storage Used (Byte)', 'Storage Used (Bytes)')
$colDeletedCount    = Resolve-Column $headers @('Deleted Item Count', 'DeletedItemCount')
$colDeletedSize     = Resolve-Column $headers @('Deleted Item Size (Byte)', 'Deleted Item Size (Bytes)')
$colLastActivity    = Resolve-Column $headers @('Last Activity Date', 'LastActivityDate')
$colCreatedDate     = Resolve-Column $headers @('Created Date', 'CreatedDate')
$colIsDeleted       = Resolve-Column $headers @('Is Deleted', 'IsDeleted')

$results = foreach ($row in $raw) {
    $storageBytes = 0L
    if ($colStorage) {
        $rawStorage = $row.$colStorage
        if ($rawStorage -match '^\d+$') { $storageBytes = [long]$rawStorage }
    }

    $recipientType = if ($colRecipientType) { $row.$colRecipientType } else { 'N/A' }
    if ([string]::IsNullOrWhiteSpace($recipientType)) { $recipientType = 'N/A' }

    [PSCustomObject]@{
        DisplayName           = if ($colDisplayName)   { $row.$colDisplayName }   else { 'N/A' }
        UserPrincipalName     = if ($colUPN)            { $row.$colUPN }            else { 'N/A' }
        RecipientType         = $recipientType
        ItemCount             = if ($colItemCount)      { $row.$colItemCount }      else { 'N/A' }
        StorageUsedBytes      = $storageBytes
        StorageUsedFormatted  = Format-Bytes $storageBytes
        DeletedItemCount      = if ($colDeletedCount)  { $row.$colDeletedCount }   else { 'N/A' }
        DeletedItemSizeBytes  = if ($colDeletedSize)   { $row.$colDeletedSize }    else { 'N/A' }
        LastActivityDate      = if ($colLastActivity)  { $row.$colLastActivity }   else { 'N/A' }
        CreatedDate           = if ($colCreatedDate)   { $row.$colCreatedDate }    else { 'N/A' }
        IsDeleted             = if ($colIsDeleted)     { $row.$colIsDeleted }      else { 'N/A' }
        ReportPeriod          = "$Period days"
        ReportDate            = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

# Sort descending by storage
$results = $results | Sort-Object StorageUsedBytes -Descending

# ── Export CSV ────────────────────────────────────────────────────────────────
$results | Select-Object -ExcludeProperty StorageUsedBytes |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "  CSV exported  →  $csvPath" -ForegroundColor Green

# ── Summary stats ─────────────────────────────────────────────────────────────
$totalMailboxes   = @($results).Count
$userMailboxes    = @($results | Where-Object { $_.RecipientType -eq 'User' }).Count
$sharedMailboxes  = @($results | Where-Object { $_.RecipientType -eq 'Shared' }).Count
$totalStorageGB   = [math]::Round(($results | Measure-Object StorageUsedBytes -Sum).Sum / 1GB, 2)
$avgStorageMB     = if ($totalMailboxes -gt 0) {
    [math]::Round(($results | Measure-Object StorageUsedBytes -Average).Average / 1MB, 2)
} else { 0 }
$largestMailbox   = $results | Select-Object -First 1
$zeroActivity     = @($results | Where-Object { $_.LastActivityDate -eq 'N/A' -or $_.LastActivityDate -eq '' }).Count

Write-Host ''
Write-Host '  ── Summary ──────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host "  Total mailboxes   : $totalMailboxes"
Write-Host "  User mailboxes    : $userMailboxes"
Write-Host "  Shared mailboxes  : $sharedMailboxes"
Write-Host "  Total storage     : $totalStorageGB GB"
Write-Host "  Avg per mailbox   : $avgStorageMB MB"
Write-Host "  No activity       : $zeroActivity"
if ($largestMailbox) {
    Write-Host "  Largest mailbox   : $($largestMailbox.DisplayName) ($($largestMailbox.StorageUsedFormatted))"
}
Write-Host ''

# ── HTML Report ───────────────────────────────────────────────────────────────
if ($HtmlReport) {
    Write-Host '  Generating HTML report...' -ForegroundColor Gray

    $tableRows = foreach ($r in $results) {
        $sizeClass = if ($r.StorageUsedBytes -gt 45GB) { 'warn' }
                     elseif ($r.StorageUsedBytes -gt 40GB) { 'caution' }
                     else { '' }
        $actClass  = if ($r.LastActivityDate -eq 'N/A' -or $r.LastActivityDate -eq '') { 'warn' } else { '' }

        "<tr>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.DisplayName))</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($r.UserPrincipalName))</td>
          <td>$($r.RecipientType)</td>
          <td>$($r.ItemCount)</td>
          <td class='$sizeClass'>$($r.StorageUsedFormatted)</td>
          <td>$($r.DeletedItemCount)</td>
          <td class='$actClass'>$($r.LastActivityDate)</td>
          <td>$($r.IsDeleted)</td>
        </tr>"
    }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Exchange Online Mailbox Statistics Report</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#1a1a1a;background:#f4f4f0}
.page{max-width:1200px;margin:0 auto;padding:24px 20px 60px}
.header{background:#0C447C;color:#fff;padding:28px 32px;border-radius:10px;margin-bottom:24px}
.header h1{font-size:20px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.header .meta{display:flex;gap:28px;margin-top:14px;flex-wrap:wrap}
.header .meta-item{font-size:11px}
.header .meta-item strong{display:block;opacity:.65;font-size:10px;text-transform:uppercase;letter-spacing:.06em}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px}
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
td.caution{color:#D48806;font-weight:600}
.footer{text-align:center;font-size:11px;color:#aaa;margin-top:40px}
.footer a{color:#0C447C;text-decoration:none}
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <h1>Exchange Online Mailbox Statistics Report</h1>
    <p>4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works</p>
    <p style="opacity:.5;margin-top:2px;font-size:11px">Microsoft Graph API v1.0</p>
    <div class="meta">
      <div class="meta-item"><strong>Tenant ID</strong>$tenant</div>
      <div class="meta-item"><strong>Report Period</strong>Last $Period days</div>
      <div class="meta-item"><strong>Generated</strong>$generatedAt</div>
    </div>
  </div>

  <div class="metrics">
    <div class="metric"><div class="metric-val">$totalMailboxes</div><div class="metric-lbl">Total mailboxes</div></div>
    <div class="metric"><div class="metric-val">$userMailboxes</div><div class="metric-lbl">User mailboxes</div></div>
    <div class="metric"><div class="metric-val">$sharedMailboxes</div><div class="metric-lbl">Shared mailboxes</div></div>
    <div class="metric"><div class="metric-val">$totalStorageGB GB</div><div class="metric-lbl">Total storage used</div></div>
    <div class="metric"><div class="metric-val">$avgStorageMB MB</div><div class="metric-lbl">Avg per mailbox</div></div>
    <div class="metric"><div class="metric-val">$zeroActivity</div><div class="metric-lbl">No activity recorded</div></div>
  </div>

  <div class="section">
    <h2>Mailbox Detail — Sorted by Storage Used (Descending)</h2>
    <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Display Name</th>
          <th>User Principal Name</th>
          <th>Type</th>
          <th>Item Count</th>
          <th>Storage Used</th>
          <th>Deleted Items</th>
          <th>Last Activity</th>
          <th>Deleted</th>
        </tr>
      </thead>
      <tbody>
        $($tableRows -join "`n")
      </tbody>
    </table>
    </div>
  </div>

  <div class="footer">
    Generated by 4TH AND BAILEY | Information Technology Consulting &bull;
    <a href="https://4thandbailey.com">4thandbailey.com</a> — Where IT Works &bull;
    Mailbox Statistics Report v$version &bull; Microsoft Graph API v1.0 &bull; $generatedAt
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
Write-Host ''
Write-Host '  Done.' -ForegroundColor Cyan
Write-Host ''
