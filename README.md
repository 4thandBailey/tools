# 4th and Bailey — Microsoft 365 PowerShell Tools

**Free, open-source Microsoft 365 administration tools built on the Microsoft Graph API.**

[![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0-0078D4?logo=microsoft)](https://learn.microsoft.com/en-us/graph/overview)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

---

## About

**4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works**

4th and Bailey is a boutique Enterprise IT Consulting firm headquartered in Houston, TX, specializing in Microsoft Cloud infrastructure, cybersecurity, and infrastructure governance. We are a Microsoft Cloud Solution Provider (CSP) serving organizations nationwide.

These tools are the freeware subset of our Microsoft Cloud Practice Toolkit — practical, production-tested scripts we use daily that we believe the broader IT community will find useful.

Every script is:
- **Cross-platform** — PowerShell 7.0+ on Windows, macOS, and Linux
- **Graph API v1.0 only** — no beta endpoints in production scripts
- **Read-only** — no write, modify, or delete operations
- **HTML report ready** — branded, client-deliverable output via `-HtmlReport`
- **App-only and delegated auth** — supports both interactive and unattended execution

---

## Tools

| Script | Description | Key Permissions |
|--------|-------------|-----------------|
| [`Get-MailboxStatisticsReport.ps1`](#1-get-mailboxstatisticsreportps1) | Mailbox size, item count, last activity | `Reports.Read.All` |
| [`Get-LicenseAssignmentReport.ps1`](#2-get-licenseassignmentreportps1) | Per-user license assignments + tenant SKU inventory | `User.Read.All` |
| [`Get-InactiveUserReport.ps1`](#3-get-inactiveuserreportps1) | Users with no sign-in activity (30/60/90 days) | `User.Read.All`, `AuditLog.Read.All` |
| [`Get-GroupMembershipReport.ps1`](#4-get-groupmembershipreportps1) | All groups and members — flat CSV + grouped HTML | `Group.Read.All` |
| [`Get-MFAStatusReport.ps1`](#5-get-mfastatusreportps1) | Per-user MFA status and registered auth methods | `UserAuthenticationMethod.Read.All` |

---

## Prerequisites

### PowerShell 7.0+

```bash
# macOS
brew install powershell

# Linux (Ubuntu/Debian)
sudo snap install powershell --classic

# Windows
winget install Microsoft.PowerShell
```

### Microsoft Graph PowerShell SDK

```powershell
# Install all required modules
Install-Module Microsoft.Graph.Authentication      -Scope CurrentUser
Install-Module Microsoft.Graph.Reports             -Scope CurrentUser
Install-Module Microsoft.Graph.Users               -Scope CurrentUser
Install-Module Microsoft.Graph.Groups              -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns    -Scope CurrentUser
```

---

## Authentication

All scripts support two authentication modes:

### Delegated (Interactive) — Quickest to start

```powershell
# No parameters required — browser prompt will open
.\Get-MailboxStatisticsReport.ps1 -HtmlReport
```

### App-Only (Unattended) — Recommended for automation

Create an App Registration in Entra ID with the required permissions for each script, then:

```powershell
# Windows — certificate thumbprint
.\Get-MailboxStatisticsReport.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -CertificateThumbprint "your-cert-thumbprint" `
    -HtmlReport

# macOS / Linux — certificate file
.\Get-MailboxStatisticsReport.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -CertificatePath "/path/to/cert.pfx" `
    -CertificatePassword (Read-Host -AsSecureString "Certificate Password") `
    -HtmlReport
```

---

## Tool Reference

### 1. Get-MailboxStatisticsReport.ps1

Generates a mailbox statistics report using the Graph Reports API. Shows storage usage, item count, deleted item count, and last activity date for every mailbox. Sorted by storage descending — largest mailboxes first.

**Required permissions:** `Reports.Read.All`, `User.Read.All`

```powershell
# Interactive, HTML report, last 30 days (default)
.\Get-MailboxStatisticsReport.ps1 -HtmlReport

# Last 90 days, save to specific directory
.\Get-MailboxStatisticsReport.ps1 -Period 90 -OutputPath ~/Reports -HtmlReport
```

**Output:**
- `MailboxStatistics_YYYYMMDD_HHmmss.csv`
- `MailboxStatistics_YYYYMMDD_HHmmss.html` (with `-HtmlReport`)

---

### 2. Get-LicenseAssignmentReport.ps1

Per-user license assignment report with tenant SKU inventory. Shows every user's assigned licenses, account status, and last sign-in. The SKU inventory table shows total, used, and available seats across all subscribed plans — useful for identifying over-licensed or under-utilized SKUs.

**Required permissions:** `User.Read.All`, `Organization.Read.All`

```powershell
# Interactive, with HTML report
.\Get-LicenseAssignmentReport.ps1 -HtmlReport

# Include SKU part numbers in CSV
.\Get-LicenseAssignmentReport.ps1 -HtmlReport -IncludeServicePlans
```

**Output:**
- `LicenseAssignment_YYYYMMDD_HHmmss.csv`
- `LicenseAssignment_YYYYMMDD_HHmmss.html` (with `-HtmlReport`)

---

### 3. Get-InactiveUserReport.ps1

Identifies users with no sign-in activity beyond a configurable threshold (30, 60, or 90 days). Highlights licensed accounts that are inactive — the most common source of M365 license waste. Requires Entra ID P1 or P2 for `signInActivity` data.

**Required permissions:** `User.Read.All`, `AuditLog.Read.All`

```powershell
# 90-day threshold, exclude guests and disabled, HTML report
.\Get-InactiveUserReport.ps1 -InactiveDays 90 -ExcludeGuests -ExcludeDisabled -HtmlReport

# 60-day threshold, all user types
.\Get-InactiveUserReport.ps1 -InactiveDays 60 -HtmlReport
```

**Output:**
- `InactiveUsers_90d_YYYYMMDD_HHmmss.csv`
- `InactiveUsers_90d_YYYYMMDD_HHmmss.html` (with `-HtmlReport`)

---

### 4. Get-GroupMembershipReport.ps1

Exports all groups and their members to a flat CSV where each row is a group-member pair. The HTML report renders each group as a card with its members in a table — significantly more readable than a flat list for presentation to clients or leadership.

**Required permissions:** `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All`

```powershell
# All group types, HTML report
.\Get-GroupMembershipReport.ps1 -HtmlReport

# Security groups only, exclude empty groups
.\Get-GroupMembershipReport.ps1 -GroupType Security -ExcludeEmptyGroups -HtmlReport

# Microsoft 365 groups only
.\Get-GroupMembershipReport.ps1 -GroupType M365 -HtmlReport
```

**Output:**
- `GroupMembership_YYYYMMDD_HHmmss.csv`
- `GroupMembership_YYYYMMDD_HHmmss.html` (with `-HtmlReport`)

---

### 5. Get-MFAStatusReport.ps1

Per-user MFA status and registered authentication methods. Classifies users as:

| Status | Meaning |
|--------|---------|
| `No MFA` | No MFA methods registered — immediate risk |
| `Legacy MFA Only` | Phone/SMS only — no Authenticator App or FIDO2 |
| `MFA Registered` | MFA registered but type unclear |
| `Strong MFA` | Microsoft Authenticator App or FIDO2/Passkey |

Report is sorted No MFA → Legacy → Registered → Strong for immediate action prioritization.

**Required permissions:** `User.Read.All`, `UserAuthenticationMethod.Read.All`

> **Note:** `UserAuthenticationMethod.Read.All` requires Global Admin or Authentication Administrator consent.

```powershell
# Licensed users only, exclude guests, HTML report
.\Get-MFAStatusReport.ps1 -LicensedUsersOnly -ExcludeGuests -HtmlReport

# All users
.\Get-MFAStatusReport.ps1 -HtmlReport
```

**Output:**
- `MFAStatus_YYYYMMDD_HHmmss.csv`
- `MFAStatus_YYYYMMDD_HHmmss.html` (with `-HtmlReport`)

---

## Output Format

All scripts produce:

1. **CSV** — machine-readable, importable into Excel, Power BI, or any reporting tool
2. **HTML** (optional, `-HtmlReport`) — branded, client-ready report with summary metrics

HTML reports use the 4th and Bailey design system: Segoe UI typography, Communication Blue (`#0C447C`) brand color, and a consistent card-based layout.

---

## Common Issues

**`signInActivity` returns null for all users**
Entra ID P1 or P2 is required in the tenant for sign-in activity data. Without it, `LastSignInDateTime` will be null and all users will appear as inactive.

**`UserAuthenticationMethod.Read.All` permission denied**
This permission requires admin consent. In delegated mode, the signed-in user must hold Global Admin or Authentication Administrator. In app-only mode, grant admin consent in the Azure Portal under App Registrations → API Permissions.

**Rate limiting on large tenants**
Graph API enforces throttling on high-volume requests. Scripts include basic error handling. For tenants with 10,000+ users, consider running during off-peak hours.

**Module not found**
Run `Install-Module <module-name> -Scope CurrentUser` for each required module listed in the script header.

---

## Security & Privacy

- All scripts are **read-only**. No write, modify, or delete Graph API calls are made.
- No data is transmitted to any third party. All output is written locally.
- Credentials are handled by the Microsoft Graph PowerShell SDK and never stored by these scripts.
- For app-only authentication, use certificate-based auth (not client secrets) in production.

---

## Contributing

Issues and pull requests are welcome. Please open an issue before submitting a PR for significant changes.

---

## License

MIT License — free to use, modify, and distribute. See [LICENSE](LICENSE) for details.

---

## About 4th and Bailey

**4TH AND BAILEY | Information Technology Consulting**
**4thandbailey.com — Where IT Works**

Enterprise IT Consulting · Microsoft CSP · Houston, TX · Nationwide

- [Website](https://4thandbailey.com)
- [LinkedIn](https://www.linkedin.com/company/4thandbailey)
- [GitHub](https://github.com/4thandbailey)
- [(888) 305-5977](tel:+18883055977)

*Boutique firm. Enterprise standards. Principal-led on every engagement.*

> **4TH AND BAILEY | Information Technology Consulting (4thandbailey.com) — Where IT Works**
