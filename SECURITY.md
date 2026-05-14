# Security Policy

## Overview

4th and Bailey takes the security of this project seriously. This repository
contains PowerShell tools that run against Microsoft 365 tenants using the
Microsoft Graph API. We hold these tools to the same security standards we
apply in client engagements.

If you believe you have found a security vulnerability in any script in this
repository, please report it to us as described below.

---

## Supported Versions

Security fixes are applied to the **latest release only**.

| Version | Supported          |
|---------|--------------------|
| 1.0.x   | ✅ Yes             |
| < 1.0   | ❌ No              |

---

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub Issues.**

Instead, report vulnerabilities by emailing us directly:

**📧 security@4thandbailey.com**

Include the following in your report:

- A description of the vulnerability and its potential impact
- The affected script(s) and version
- Steps to reproduce or a proof-of-concept (if applicable)
- Any suggested remediation, if you have one

We will acknowledge receipt of your report within **2 business days** and
provide a remediation timeline within **5 business days**.

We ask that you give us reasonable time to address the issue before any public
disclosure. We will credit researchers who responsibly disclose vulnerabilities,
unless they prefer to remain anonymous.

---

## Security Design of These Tools

Understanding how these scripts are built helps contextualize the threat model.

### Read-Only by Design

All scripts in this repository make **read-only** Microsoft Graph API calls
exclusively. No script issues any `POST`, `PATCH`, `PUT`, or `DELETE` request.
No script creates, modifies, or deletes any resource in any Microsoft 365 tenant.

### No Data Transmission to Third Parties

All output — CSV files and HTML reports — is written **locally** to the machine
running the script. No tenant data is transmitted to 4th and Bailey or any
third party. The scripts have no telemetry, no callbacks, and no network
destinations other than `graph.microsoft.com` and the Microsoft identity
platform (`login.microsoftonline.com`).

### Credential Handling

Credentials are handled exclusively by the
[Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview).
These scripts do not store, log, cache, or transmit authentication tokens or
credentials. Tokens are managed entirely in memory by the SDK for the duration
of the session.

### Authentication Recommendations

For unattended (app-only) execution, we recommend and document
**certificate-based authentication** rather than client secrets:

- Certificate thumbprint on Windows
- Certificate file path on macOS and Linux

Client secrets are **not recommended** for production use. They are not
referenced in any script and are not a supported authentication path in this
toolkit.

### Principle of Least Privilege

Each script documents the **minimum Graph API permissions** required for its
function. We recommend granting only those specific permissions when creating
App Registrations for use with these scripts. Permissions required per script:

| Script | Required Permissions |
|--------|----------------------|
| `Get-MailboxStatisticsReport.ps1` | `Reports.Read.All`, `User.Read.All` |
| `Get-LicenseAssignmentReport.ps1` | `User.Read.All`, `Organization.Read.All` |
| `Get-InactiveUserReport.ps1` | `User.Read.All`, `AuditLog.Read.All` |
| `Get-GroupMembershipReport.ps1` | `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All` |
| `Get-MFAStatusReport.ps1` | `User.Read.All`, `UserAuthenticationMethod.Read.All` |

`UserAuthenticationMethod.Read.All` is a sensitive permission requiring
Global Admin or Authentication Administrator consent. Use it only in
app-only registrations with scoped, audited access.

### Graph API Endpoint Policy

All scripts target **Microsoft Graph API v1.0 only**. No beta endpoints
(`/beta/`) are used in any production script in this repository. This is a
deliberate policy decision: beta endpoints may change without notice and are
not appropriate for production tooling running in client environments.

---

## Scope

The following are **in scope** for security reports:

- Logic errors that could cause a script to behave in a destructive or
  unintended way against a Microsoft 365 tenant
- Authentication or credential handling issues introduced by the scripts
  themselves (not by the Microsoft Graph SDK)
- Output files that inadvertently expose sensitive data beyond what is
  documented
- Dependency vulnerabilities in the Microsoft Graph PowerShell SDK modules
  the scripts depend on (report these upstream to Microsoft, but notify us
  as well)

The following are **out of scope**:

- Vulnerabilities in Microsoft Graph API itself — report those to
  [Microsoft Security Response Center (MSRC)](https://msrc.microsoft.com)
- Vulnerabilities in the Microsoft Graph PowerShell SDK — report those to
  [Microsoft](https://github.com/microsoftgraph/msgraph-sdk-powershell/security)
- Social engineering attacks
- Denial-of-service against the Graph API (covered by Microsoft's own rate
  limiting and throttling)

---

## Microsoft Security Resources

Since these scripts interact exclusively with Microsoft cloud services,
the following Microsoft security resources are relevant:

- [Microsoft Security Response Center (MSRC)](https://msrc.microsoft.com)
- [Report a vulnerability to Microsoft](https://msrc.microsoft.com/create-report)
- [Microsoft Graph security documentation](https://learn.microsoft.com/en-us/graph/security-concept-overview)
- [Microsoft Entra ID security best practices](https://learn.microsoft.com/en-us/entra/identity/standards/standards-overview)

---

*4TH AND BAILEY | Information Technology Consulting*
*Microsoft Cloud Solution Provider · Houston, TX · Nationwide*
*[4thandbailey.com](https://4thandbailey.com) · security@4thandbailey.com*
