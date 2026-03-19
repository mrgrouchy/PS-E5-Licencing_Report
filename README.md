# PS E5 Licensing Report

PowerShell scripts for Microsoft 365 E5 license analysis, with two reporting modes:

- `E5_Report.ps1`: AD-based last sign-in reporting for users with E5 licenses.
- `PS-E5_tenant_overview.ps1`: tenant-wide Entra + Exchange overview with E5 status, mailbox type, and inactivity signals.

## What Each Script Is For

### `E5_Report.ps1` (AD-focused)
- Uses Microsoft Graph to identify users with `ENTERPRISEPREMIUM` and/or `SPE_E5`.
- Uses on-premises Active Directory `lastLogonTimestamp` for activity.
- Exports only E5-licensed users.
- Best for hybrid identity environments where AD sign-in recency matters.

### `PS-E5_tenant_overview.ps1` (Tenant-wide governance)
- Pulls all Entra users, then marks E5 vs non-E5.
- Uses Entra `SignInActivity.LastSignInDateTime`.
- Connects to Exchange Online to classify mailbox type (`User`, `Shared`, `Room`, etc.).
- Produces a summary of disabled E5 users, shared mailbox E5 users, and E5 users inactive for 90+ days.

## Prerequisites

- PowerShell 7+ recommended (Windows PowerShell 5.1 may also work depending on module versions).
- Permissions to connect to:
  - Microsoft Graph
  - Exchange Online (for tenant overview script)
  - Active Directory (for AD report script)

Install modules:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ActiveDirectory
```

Notes:
- `ActiveDirectory` typically requires RSAT/domain connectivity.
- `SignInActivity` may require appropriate Entra ID licensing and Graph permissions.

## Required Graph Scopes

### `E5_Report.ps1`
- `User.Read.All`

### `PS-E5_tenant_overview.ps1`
- `User.Read.All`
- `Directory.Read.All`

## Usage

From the repository root:

```powershell
pwsh .\E5_Report.ps1
pwsh .\PS-E5_tenant_overview.ps1
```

If you are already in an authenticated session, the scripts will reuse the session where possible.

## Output Files

### AD report
- File pattern: `AD_E5_Users_Report_yyyyMMdd-HHmm.csv`
- Includes:
  - user identity columns
  - matched E5 SKU labels
  - all assigned SKUs
  - AD last sign-in and days since sign-in
  - employee type handling flags (indirectly via output and console warnings)

### Tenant overview report
- File pattern: `E5_License_Report_yyyyMMdd-HHmmss.csv`
- Includes:
  - account status (enabled/disabled)
  - E5 license status
  - target E5 SKUs
  - mailbox classification (`RecipientTypeDetails`, friendly mailbox type)
  - Entra last sign-in and days since sign-in
  - org profile fields (country, employee type, user type)

## Key Differences At A Glance

1. Data scope:
   - `E5_Report.ps1`: only E5-licensed users.
   - `PS-E5_tenant_overview.ps1`: all users, with E5 status indicated.
2. Activity source:
   - `E5_Report.ps1`: AD `lastLogonTimestamp`.
   - `PS-E5_tenant_overview.ps1`: Entra sign-in activity.
3. Mailbox awareness:
   - `E5_Report.ps1`: none.
   - `PS-E5_tenant_overview.ps1`: Exchange mailbox type classification.
4. Main use case:
   - `E5_Report.ps1`: hybrid/AD operational reporting.
   - `PS-E5_tenant_overview.ps1`: licensing governance and cleanup analysis.

## Troubleshooting

- `No matching E5 SKUs found`:
  - Verify the tenant has `ENTERPRISEPREMIUM` and/or `SPE_E5` subscriptions.
- Graph auth prompts repeatedly:
  - Clear token cache and reconnect with required scopes.
- `Get-ADUser` fails:
  - Confirm RSAT AD module availability and domain connectivity.
- Exchange cmdlets fail in tenant overview:
  - Ensure `ExchangeOnlineManagement` is installed and account has Exchange admin visibility.

## Security And Operational Notes

- Scripts export user metadata to local CSV files.
- Review export handling and storage retention in line with your internal policies.
- Run scripts using least-privilege accounts that still have required read permissions.

## Contributing

Please use the issue templates for bugs/feature requests and submit pull requests with clear change descriptions and test notes.

## License

This project is licensed under the terms in the [LICENSE](LICENSE) file.
