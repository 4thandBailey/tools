#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Group & Distribution List Membership Export

.DESCRIPTION
    Exports all Microsoft 365 groups, security groups, and distribution lists
    with their members to a flat CSV file using the Microsoft Graph API.
    Optionally generates a grouped HTML report showing each group and its
    members in a readable format.

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

.PARAMETER GroupType
    Filter by group type. Valid values: All, Security, M365, Distribution.
    Defaults to All.

.PARAMETER ExcludeEmptyGroups
    Switch to exclude groups with no members from the report.

.EXAMPLE
    .\Get-GroupMembershipReport.ps1 -HtmlReport

.EXAMPLE
    .\Get-GroupMembershipReport.ps1 -GroupType Security -HtmlReport -ExcludeEmptyGroups

.EXAMPLE
    .\Get-GroupMembershipReport.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -CertificatePath "/certs/app.pfx" `
        -CertificatePassword (Read-Host -AsSecureString "Password") `
        -HtmlReport -GroupType All

.NOTES
    Author:      4TH AND BAILEY | Information Technology Consulting
                 4thandbailey.com — Where IT Works
    Version:     1.0.0
    GitHub:      https://github.com/4thandbailey
    Website:     https://4thandbailey.com

    Required Modules:
        Microsoft.Graph.Authentication
        Microsoft.Graph.Groups

    Required Permissions:
        Group.Read.All
        GroupMember.Read.All
        User.Read.All

    Install modules:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Groups -Scope CurrentUser
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
    [ValidateSet('All', 'Security', 'M365', 'Distribution')]
    [string]$GroupType = 'All',

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [switch]$HtmlReport,

    [Parameter()]
    [switch]$ExcludeEmptyGroups
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Banner ────────────────────────────────────────────────────────────────────
$version = '1.0.0'
Write-Host ''
Write-Host '  4TH AND BAILEY | Information Technology Consulting — Where IT Works' -ForegroundColor Cyan
Write-Host '  M365 Group Membership Export' -ForegroundColor Cyan
Write-Host "  Version $version  |  4thandbailey.com" -ForegroundColor DarkCyan
Write-Host ''

# ── Module check ──────────────────────────────────────────────────────────────
Write-Host '  Checking required modules...' -ForegroundColor Gray
foreach ($mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')) {
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
        Connect-MgGraph -Scopes 'Group.Read.All', 'GroupMember.Read.All', 'User.Read.All' -NoWelcome
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
$csvPath   = Join-Path $OutputPath "GroupMembership_$timestamp.csv"
$htmlPath  = Join-Path $OutputPath "GroupMembership_$timestamp.html"

# ── Fetch groups ──────────────────────────────────────────────────────────────
Write-Host "  Fetching groups (Type: $GroupType)..." -ForegroundColor Gray

$allGroups = Get-MgGroup -All -Property 'id,displayName,mail,groupTypes,securityEnabled,mailEnabled,description,membershipRule,createdDateTime'

# Classify and filter by type
$groups = foreach ($g in $allGroups) {
    $isM365        = $g.GroupTypes -contains 'Unified'
    $isDynamic     = $g.GroupTypes -contains 'DynamicMembership'
    $isSecurity    = $g.SecurityEnabled -and -not $isM365
    $isDistrib     = $g.MailEnabled -and -not $g.SecurityEnabled -and -not $isM365

    $typeLabel = if ($isM365)     { 'Microsoft 365' }
                 elseif ($isDistrib) { 'Distribution' }
                 elseif ($isSecurity) { 'Security' }
                 else { 'Other' }

    $membershipLabel = if ($isDynamic) { 'Dynamic' } else { 'Assigned' }

    $include = switch ($GroupType) {
        'M365'         { $isM365 }
        'Security'     { $isSecurity }
        'Distribution' { $isDistrib }
        default        { $true }
    }

    if ($include) {
        [PSCustomObject]@{
            Id             = $g.Id
            DisplayName    = if ($g.DisplayName) { $g.DisplayName } else { 'N/A' }
            Mail           = if ($g.Mail)        { $g.Mail }        else { 'N/A' }
            TypeLabel      = $typeLabel
            MembershipType = $membershipLabel
            Description    = if ($g.Description) { $g.Description } else { '' }
            CreatedDate    = if ($null -ne $g.CreatedDateTime) {
                                 $g.CreatedDateTime.ToString('yyyy-MM-dd')
                             } else { 'N/A' }
        }
    }
}

$groups = @($groups)
Write-Host "  Found $($groups.Count) groups after filter." -ForegroundColor Green

# ── Fetch members for each group ──────────────────────────────────────────────
Write-Host '  Fetching group members (this may take several minutes for large tenants)...' -ForegroundColor Gray

$csvRows      = [System.Collections.Generic.List[PSCustomObject]]::new()
$groupSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
$processed    = 0

foreach ($group in $groups) {
    $processed++
    if ($processed % 10 -eq 0) {
        Write-Host "  Processing group $processed of $($groups.Count)..." -ForegroundColor Gray
    }

    $members = @()
    try {
        $rawMembers = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop |
            Select-Object -ExpandProperty AdditionalProperties -ErrorAction SilentlyContinue

        # Alternatively use ExpandProperty for richer data
        $memberObjects = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop

        $members = foreach ($m in $memberObjects) {
            $displayName = if ($m.AdditionalProperties -and $m.AdditionalProperties['displayName']) {
                $m.AdditionalProperties['displayName']
            } else { 'N/A' }

            $upn = if ($m.AdditionalProperties -and $m.AdditionalProperties['userPrincipalName']) {
                $m.AdditionalProperties['userPrincipalName']
            } else { 'N/A' }

            $odataType = if ($m.AdditionalProperties -and $m.AdditionalProperties['@odata.type']) {
                $m.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.',''
            } else { 'unknown' }

            [PSCustomObject]@{
                MemberId          = $m.Id
                MemberDisplayName = $displayName
                MemberUPN         = $upn
                MemberType        = $odataType
            }
        }
    } catch {
        Write-Warning "  Could not retrieve members for '$($group.DisplayName)': $_"
        $members = @()
    }

    $memberCount = @($members).Count

    if ($ExcludeEmptyGroups -and $memberCount -eq 0) { continue }

    $groupSummary.Add([PSCustomObject]@{
        GroupName      = $group.DisplayName
        GroupMail      = $group.Mail
        GroupType      = $group.TypeLabel
        MembershipType = $group.MembershipType
        MemberCount    = $memberCount
        CreatedDate    = $group.CreatedDate
    })

    if ($memberCount -eq 0) {
        $csvRows.Add([PSCustomObject]@{
            GroupName         = $group.DisplayName
            GroupMail         = $group.Mail
            GroupType         = $group.TypeLabel
            MembershipType    = $group.MembershipType
            MemberId          = 'N/A'
            MemberDisplayName = '(No members)'
            MemberUPN         = 'N/A'
            MemberType        = 'N/A'
            GroupCreatedDate  = $group.CreatedDate
            ReportDate        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        })
    } else {
        foreach ($m in $members) {
            $csvRows.Add([PSCustomObject]@{
                GroupName         = $group.DisplayName
                GroupMail         = $group.Mail
                GroupType         = $group.TypeLabel
                MembershipType    = $group.MembershipType
                MemberId          = $m.MemberId
                MemberDisplayName = $m.MemberDisplayName
                MemberUPN         = $m.MemberUPN
                MemberType        = $m.MemberType
                GroupCreatedDate  = $group.CreatedDate
                ReportDate        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            })
        }
    }
}

# ── Export CSV ────────────────────────────────────────────────────────────────
$csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "  CSV exported  →  $csvPath" -ForegroundColor Green

# ── Summary ───────────────────────────────────────────────────────────────────
$totalGroups   = $groupSummary.Count
$emptyGroups   = @($groupSummary | Where-Object { $_.MemberCount -eq 0 }).Count
$totalMembers  = ($groupSummary | Measure-Object MemberCount -Sum).Sum
$largestGroup  = $groupSummary | Sort-Object MemberCount -Descending | Select-Object -First 1

Write-Host ''
Write-Host '  ── Summary ──────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host "  Total groups        : $totalGroups"
Write-Host "  Empty groups        : $emptyGroups"
Write-Host "  Total memberships   : $totalMembers"
if ($largestGroup) {
    Write-Host "  Largest group       : $($largestGroup.GroupName) ($($largestGroup.MemberCount) members)"
}
Write-Host ''

# ── HTML Report ───────────────────────────────────────────────────────────────
if ($HtmlReport) {
    Write-Host '  Generating HTML report...' -ForegroundColor Gray

    # Group the CSV rows by group name for HTML rendering
    $groupedData  = $csvRows | Group-Object GroupName | Sort-Object Name

    $groupSections = foreach ($grp in $groupedData) {
        $firstRow     = $grp.Group[0]
        $memberCount  = if ($firstRow.MemberDisplayName -eq '(No members)') { 0 } else { $grp.Count }
        $emptyClass   = if ($memberCount -eq 0) { ' style="color:#888;font-style:italic"' } else { '' }

        $memberRows = foreach ($m in $grp.Group) {
            if ($m.MemberDisplayName -eq '(No members)') {
                "<tr><td colspan='4' style='color:#aaa;font-style:italic'>No members</td></tr>"
            } else {
                "<tr>
                  <td>$([System.Web.HttpUtility]::HtmlEncode($m.MemberDisplayName))</td>
                  <td>$([System.Web.HttpUtility]::HtmlEncode($m.MemberUPN))</td>
                  <td>$($m.MemberType)</td>
                  <td>$($m.MemberId)</td>
                </tr>"
            }
        }

        "<div class='group-card'>
          <div class='group-header'>
            <div>
              <span class='group-name'$emptyClass>$([System.Web.HttpUtility]::HtmlEncode($firstRow.GroupName))</span>
              <span class='group-mail'>$([System.Web.HttpUtility]::HtmlEncode($firstRow.GroupMail))</span>
            </div>
            <div class='group-meta'>
              <span class='badge'>$($firstRow.GroupType)</span>
              <span class='badge'>$($firstRow.MembershipType)</span>
              <span class='badge'>$memberCount members</span>
            </div>
          </div>
          <table>
            <thead><tr><th>Display Name</th><th>UPN</th><th>Type</th><th>Object ID</th></tr></thead>
            <tbody>$($memberRows -join "`n")</tbody>
          </table>
        </div>"
    }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 Group Membership Report</title>
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
.group-card{background:#fff;border-radius:10px;padding:16px 20px;margin-bottom:16px;border:1px solid #e0dfd8}
.group-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.group-name{font-size:14px;font-weight:600;color:#0C447C;display:block}
.group-mail{font-size:11px;color:#888;margin-top:2px;display:block}
.group-meta{display:flex;gap:6px;flex-wrap:wrap}
.badge{background:#E6F1FB;color:#0C447C;font-size:11px;padding:2px 8px;border-radius:999px;font-weight:500}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f4f4f0;text-align:left;padding:7px 10px;font-weight:600;color:#444;border-bottom:2px solid #e0dfd8}
td{padding:6px 10px;border-bottom:1px solid #f0efea;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#fafaf8}
.footer{text-align:center;font-size:11px;color:#aaa;margin-top:40px}
.footer a{color:#0C447C;text-decoration:none}
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <h1>Microsoft 365 Group Membership Report</h1>
    <p>4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works</p>
    <p style="opacity:.5;margin-top:2px;font-size:11px">Microsoft Graph API v1.0</p>
    <div class="meta">
      <div class="meta-item"><strong>Tenant ID</strong>$tenant</div>
      <div class="meta-item"><strong>Group Filter</strong>$GroupType</div>
      <div class="meta-item"><strong>Generated</strong>$generatedAt</div>
    </div>
  </div>

  <div class="metrics">
    <div class="metric"><div class="metric-val">$totalGroups</div><div class="metric-lbl">Total groups</div></div>
    <div class="metric"><div class="metric-val">$emptyGroups</div><div class="metric-lbl">Empty groups</div></div>
    <div class="metric"><div class="metric-val">$totalMembers</div><div class="metric-lbl">Total memberships</div></div>
  </div>

  $($groupSections -join "`n")

  <div class="footer">
    Generated by 4TH AND BAILEY | Information Technology Consulting &bull;
    <a href="https://4thandbailey.com">4thandbailey.com</a> — Where IT Works &bull;
    Group Membership Report v$version &bull; Microsoft Graph API v1.0 &bull; $generatedAt
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
