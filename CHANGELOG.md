# Changelog

All notable changes to this project will be documented in this file.

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] — 2025-05-12

### Added

**Initial public release of the 4th and Bailey Microsoft 365 PowerShell Tools.**

These are the freeware subset of our Microsoft Cloud Practice Toolkit —
production-tested scripts used daily in our Microsoft 365 assessments,
infrastructure governance reviews, and client onboarding engagements.

#### `Get-MFAStatusReport.ps1`
- Per-user MFA status report using `UserAuthenticationMethod.Read.All`
- Classifies each user as `No MFA`, `Legacy MFA Only`, `MFA Registered`, or `Strong MFA`
- Strong MFA classification covers Microsoft Authenticator App and FIDO2/Passkey methods
- Report sorted by risk priority: No MFA → Legacy → Registered → Strong
- Supports `-LicensedUsersOnly` and `-ExcludeGuests` filters
- CSV and branded HTML output via `-HtmlReport`
- Delegated (interactive) and app-only (unattended) authentication

#### `Get-InactiveUserReport.ps1`
- Identifies users with no sign-in activity beyond a configurable threshold
- Supported thresholds: 30, 60, and 90 days via `-InactiveDays`
- Requires Entra ID P1 or P2 for `signInActivity` data
- Supports `-ExcludeGuests` and `-ExcludeDisabled` filters
- Highlights licensed inactive accounts — the most common source of M365 license waste
- CSV and branded HTML output via `-HtmlReport`
- Delegated and app-only authentication

#### `Get-LicenseAssignmentReport.ps1`
- Per-user license assignment report with tenant-wide SKU inventory table
- Shows assigned licenses, account status, and last sign-in per user
- SKU inventory table shows total, used, and available seats across all subscribed plans
- Supports `-IncludeServicePlans` to include SKU part numbers in CSV output
- CSV and branded HTML output via `-HtmlReport`
- Delegated and app-only authentication

#### `Get-MailboxStatisticsReport.ps1`
- Mailbox statistics report using the Graph Reports API
- Reports storage usage, item count, deleted item count, and last activity date per mailbox
- Results sorted by storage descending — largest mailboxes first
- Configurable reporting period: 30, 60, or 90 days via `-Period`
- Supports `-OutputPath` to specify report output directory
- CSV and branded HTML output via `-HtmlReport`
- Delegated and app-only authentication

#### `Get-GroupMembershipReport.ps1`
- Exports all groups and their members across the tenant
- Flat CSV output: one row per group-member pair, suitable for Excel or Power BI
- HTML report renders each group as a card with members in a table
- Supports filtering by group type: All, Security, or M365 via `-GroupType`
- Supports `-ExcludeEmptyGroups` to omit groups with no members
- CSV and branded HTML output via `-HtmlReport`
- Delegated and app-only authentication

#### Cross-cutting features across all scripts
- PowerShell 7.0+ required; cross-platform on Windows, macOS, and Linux
- Microsoft Graph API v1.0 only — no beta endpoints
- Read-only — no write, modify, or delete Graph API calls
- App-only authentication supports both certificate thumbprint (Windows) and
  certificate file path (macOS/Linux)
- HTML reports use the 4th and Bailey design system: Segoe UI typography,
  Communication Blue (`#0C447C`) brand color, card-based layout
- MIT License

---

## [Unreleased]

Changes staged for the next release will appear here.

---

*Maintained by [4TH AND BAILEY | Information Technology Consulting](https://4thandbailey.com)*
*Microsoft Cloud Solution Provider · Houston, TX · Nationwide*
