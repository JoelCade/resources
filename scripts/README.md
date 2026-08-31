# Get-ManagedEnvLicenseInventory.ps1

A read-only pre-flight inventory for **Power Platform Managed Environments**.

Before you flag an environment as managed, you need to know who is in it and whether each of
those people holds a license that actually qualifies. Microsoft's own report
(*Users requiring licenses in Managed Environments*) only covers environments that are **already
managed**, and only lists users who launched an app in the selected month. It is a compliance
monitor, not a planning tool.

This script closes that gap. It joins three sources that Microsoft does not join for you and
tells you, per user, whether they are covered.

> **Companion guide:**
> [Managed Environments: Pre-Deployment Inventory and Checklist](https://joelcade.github.io/resources/power-platform-managed-environments-preflight.html)

---

## What it does

| Step | Source | What it answers |
|---|---|---|
| 1 | Dataverse `systemuser` | Who is *actually* in this environment? |
| 2 | Microsoft Graph | What licenses does each of those people hold, including licenses inherited from a group? |
| 3 | Rules table | Does each license qualify for Managed Environments? |

Every user lands in exactly one bucket:

| Bucket | Meaning | Action |
|---|---|---|
| `Covered` | Holds a qualifying standalone or Dynamics 365 Enterprise/Premium license | None |
| `NotQualifying` | Microsoft 365 seeded rights, Developer Plan, Dynamics 365 Pro, free/viral | Assign a license, or remove from the environment |
| `Unclassified` | SKU is not in the rules table | Review manually and extend the rules |
| `NoLicense` *(reported as `NoLicense`)* | In the environment, holds nothing | Likely a stale account |

The script **never writes anything** to Dataverse or Entra ID.

---

## Requirements

- **PowerShell 5.1 or PowerShell 7+**
- **Azure CLI** on the PATH, signed in: `az login`
- A Dataverse account in the target environment that can read the `systemuser` table
  (System Administrator, or equivalent read access)
- Directory read access in Entra ID. The Azure CLI first-party client normally has what is
  needed. If Graph returns 403, sign in with an account holding `Directory.Read.All`,
  `User.Read.All`, or `Organization.Read.All`.
- *Optional:* `Install-Module ImportExcel -Scope CurrentUser` for a single multi-tab workbook.
  Without it the script writes CSVs instead.

---

## Usage

```powershell
az login

.\Get-ManagedEnvLicenseInventory.ps1 -EnvironmentUrl "https://contoso.crm.dynamics.com"
```

With options:

```powershell
.\Get-ManagedEnvLicenseInventory.ps1 `
    -EnvironmentUrl "https://contoso.crm.dynamics.com" `
    -OutputPath "C:\Reports" `
    -IncludeDisabled
```

### Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-EnvironmentUrl` | *placeholder* | Dataverse environment URL. Edit the default in the param block or pass at runtime. |
| `-OutputPath` | script folder | Where to write the workbook or CSVs. |
| `-IncludeDisabled` | off | Include disabled Dataverse users. They do not consume licenses, but you may want them for an access review. |
| `-AllAccessModes` | off | Include all access modes rather than only Read-Write. Non-interactive, administrative, and application users do not consume user licenses and would inflate the count. |
| `-GraphPageAll` | `$true` | Page all tenant users once and join locally. Set `$false` to look users up individually, which pulls far less data on a large tenant with a small environment population. |

---

## Output

A five-tab workbook (or five CSVs):

| Tab | Contents |
|---|---|
| **Summary** | Count and share per bucket |
| **Action List** | Only the users who need attention, worst first |
| **All Users** | Full joined dataset, one row per user |
| **License Mix** | Every SKU seen, with its bucket and user count |
| **Run Notes** | Parameters used, timestamps, and the caveats below |

---

## Known limits - read these

**This is entitlement, not usage.** A user can be unlicensed and never launch an app, in which
case the answer may be to remove them rather than buy them a license. Cross-check against
**Power Platform admin center > Licensing > Power Apps > Download Reports > Active users**
before you purchase anything.

**Per-app plans and pay-as-you-go meters are invisible here.** Both are allocated at the
*environment* level, not to a user, so they cannot appear in per-user Entra data. If you use
either, treat `NotQualifying` as an **upper bound** and reconcile against
**Licensing > Power Apps > Environments**.

**Unknown SKUs are surfaced, never assumed.** Anything the rules table does not recognize is
reported as `Unclassified` and listed at the end of the run so you can extend the rules. The
script will not quietly guess.

---

## Licensing rules encoded

Built from the **Power Platform Licensing Guide, August 2026, p.24**, which lists the Managed
Environments entitlement, plus the fifteen Dynamics 365 SKUs named on Microsoft Learn.

### Qualifying

- Power Apps and Power Automate standalone user subscription licenses
- Power Automate Process and Hosted Process
- Power Pages user subscription capacity packs
- Copilot Studio tenant license, and Copilot Studio for Microsoft 365 Copilot
- Dynamics 365 Premium, Enterprise, and Team Members standalone licenses
- Dynamics 365 Customer Insights
- Power Apps per app, Power Pages, and Copilot Studio pay-as-you-go meters *(environment level)*

### Not qualifying

- **Dynamics 365 Pro / Professional.** Licensing Guide p.9 footnote 2 is explicit:
  *"Dynamics 365 Pro does not have use rights for Power Apps or Power Pages."*
- **Microsoft 365 and Office 365 seeded rights.** Guide p.24 footnote 1 confirms the
  entitlement covers standalone licenses and *"does not include the limited Power Apps,
  Power Automate and Power Pages use rights that come with select Dynamics 365 and
  Microsoft 365 licenses."*
- **Power Apps Developer Plan**, when running assets
- **Dynamics 365 Operations Activity and Device** licenses
- Free and viral Power Platform plans

### A trap worth knowing

Some Professional SKU part numbers embed the word `ENTERPRISE`, for example
`DYN365_ENTERPRISE_SALES_PROFESSIONAL`. A naive pattern match classifies these as Enterprise and
therefore covered, which is a false negative on exactly the population you most need to catch.
The rules table places the Professional patterns **above** the Enterprise patterns, and the
ordering is load-bearing. Preserve it if you extend the table.

---

## Extending the rules

The rules table is a plain ordered array near the middle of the script. First match wins.

```powershell
@{ Pattern = '^YOUR_SKU_PREFIX'; Bucket = 'Covered'; Family = 'Friendly name' }
```

Match on `SkuPartNumber`, the stable readable identifier, rather than the GUID. Run the script
once, read the *unrecognized SKUs* list it prints at the end, and add what you find.

---

## References

- [Managed environments licensing (Microsoft Learn)](https://learn.microsoft.com/power-platform/admin/managed-environment-licensing)
- [Managed environments overview (Microsoft Learn)](https://learn.microsoft.com/power-platform/admin/managed-environment-overview)
- [Microsoft Power Platform Licensing Guide (PDF)](https://go.microsoft.com/fwlink/?linkid=2085130)
- [Licensing overview for Microsoft Power Platform (Microsoft Learn)](https://learn.microsoft.com/power-platform/admin/pricing-billing-skus)

---

## Disclaimer

Guidance, not a contract. Licensing SKUs and rules change frequently. Always confirm current
terms against the Power Platform Licensing Guide and your organization's agreement before making
purchasing decisions.
