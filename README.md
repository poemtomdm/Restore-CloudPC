# Restore-CloudPC.ps1

A PowerShell script to bulk-restore Windows 365 Cloud PCs to a chosen date, for a list of users loaded from a file.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-Beta-orange?logo=microsoft&logoColor=white)
![Windows 365](https://img.shields.io/badge/Windows%20365-Cloud%20PC-0078D4?logo=microsoft&logoColor=white)

---

## What it does

When users report issues on their Cloud PC (ransomware, corruption, bad update…), restoring machines one by one through the Intune portal is slow and error-prone.

This script automates the full restore flow across any number of users and machines:

1. Load a list of users from a file (plain list or CSV)
2. Fetch the available restore points live from your tenant and display them as a menu
3. Pick a date — the script automatically selects the most recent snapshot of that day for each machine
4. Trigger the restore on every Cloud PC found for every user, including users with multiple machines
5. Print a summary with success / skipped / failed counts

---

## Step-by-step flow

### 1 — Load users

At startup the script asks for the path to the file containing the UPNs to process. Two formats are accepted automatically:

**Plain list** (no header):
```
alice@contoso.com
bob@contoso.com
carol@contoso.com
```

**CSV with header**:
```csv
UPN
alice@contoso.com
bob@contoso.com
```

The file is validated before continuing (must exist, be readable, and contain at least one valid UPN).

---

### 2 — Discover available restore points

The script queries the Graph API to retrieve all snapshots available for the first Cloud PC it finds. Windows 365 keeps:
- **Daily snapshots** (AM + PM) for the last 7 days
- **Weekly snapshots** beyond that (retained for a few weeks)

These are displayed as a numbered menu:

```
Available restore points:

  1. Thursday, July 10, 2026 (today)
  2. Wednesday, July 9, 2026 (1 day(s) ago)
  3. Tuesday, July 8, 2026 (2 day(s) ago)
  4. Monday, July 7, 2026 (3 day(s) ago)
  ...
  9. Saturday, June 27, 2026 (13 day(s) ago — weekly)
```

---

### 3 — Pick a restore date

Enter the number for the date you want to restore to. The script will automatically pick the **most recent snapshot of that day** for each machine — no need to choose between AM and PM snapshots.

---

### 4 — Restore all machines

For each UPN in the list, the script:
- Finds **all Cloud PCs** assigned to that user (handles users with multiple machines)
- Retrieves the snapshots available for each machine
- Selects the best matching snapshot for the chosen date
- Calls the Graph API restore action

Output during processing:
```
Processing: alice@contoso.com
  2 Cloud PC(s) found.
  ── CPC-Alice-Primary  (c86b874c-...)
     Snapshot : CPC_xxx...  (created 2026-07-09 13:07:42)
     [OK] Restore initiated successfully.
  ── CPC-Alice-Secondary  (a12b3c4d-...)
     Snapshot : CPC_yyy...  (created 2026-07-09 17:06:35)
     [OK] Restore initiated successfully.
  ------------------------------------
Processing: bob@contoso.com
  [SKIP] No Cloud PC found for bob@contoso.com.
  ------------------------------------
```

---

### 5 — Summary

```
============================================
  Done!
  Restore date : Wednesday, July 9, 2026
  Total users  : 42
  Success      : 40
  Skipped      : 1
  Failed       : 1
============================================
```

---

## Important notes

- **The restore is asynchronous.** The script triggers the action and moves on immediately. Monitor completion in the Intune portal or the Windows 365 provisioning status.
- **Users with no Cloud PC** are silently skipped and logged as `[SKIP]`.
- **Machines with no snapshot on the chosen date** are skipped individually, without interrupting the rest of the batch.
- **Restoring a Cloud PC will disconnect the active session** if the user is currently logged in.
- The script uses the Microsoft Graph **beta** endpoint — review and test after major Intune/Graph updates.

---

## Prerequisites

- PowerShell 7+ (recommended) or Windows PowerShell 5.1
- [Microsoft.Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation):
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- An Entra ID App Registration with the application permission **`CloudPC.ReadWrite.All`** (admin consent required, NOT delegated)

---

## Usage

```powershell
# Client Secret
.\Restore-CloudPC.ps1 `
    -TenantId     "contoso.onmicrosoft.com" `
    -ClientId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your-client-secret-here"

# Certificate (thumbprint — cert must be installed in the local cert store)
.\Restore-CloudPC.ps1 `
    -TenantId              "contoso.onmicrosoft.com" `
    -ClientId              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AABBCCDDEEFF00112233445566778899AABBCCDD"

# Certificate (.pfx file)
.\Restore-CloudPC.ps1 `
    -TenantId            "contoso.onmicrosoft.com" `
    -ClientId            "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificatePath     "C:\certs\myapp.pfx" `
    -CertificatePassword (Read-Host -AsSecureString "PFX password")
```

---

## License

This project is licensed under the [MIT License](LICENSE).

---

> **Disclaimer:** This script makes live changes to your tenant. Always validate in a non-production environment before running against production.
