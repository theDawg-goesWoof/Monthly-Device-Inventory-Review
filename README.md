# Monthly Device Inventory Review

Interactive PowerShell script that reviews Entra ID device objects for inactivity, disables
devices idle 90+ days, deletes devices idle 120+ days, and writes a per-tenant CSV audit trail.

Built to be run by a human, one client tenant at a time. Every destructive step is gated behind
a prompt, and nothing is changed until the operator has seen the list on screen.

- **Script:** [Monthly_Device_Inventory_Review.ps1](Monthly_Device_Inventory_Review.ps1)
- **Output:** `<Tenant>_Device_Review_<yyyy-MM-dd>.csv` beside the script

---

## ⚠️ Read this before your first run in a new tenant

The first time anyone runs this in a client tenant, Entra ID shows a **"Permissions requested"**
consent dialog for the *Microsoft Graph Command Line Tools* application. That dialog contains a
checkbox:

> ☐ **Consent on behalf of your organization**

### Leave that checkbox UNCHECKED.

Ticking it does not help you finish the review any faster — the script works identically either
way — but it makes a **permanent, tenant-wide configuration change** to your client's tenant:

| | Box unchecked (correct) | Box checked (avoid) |
|---|---|---|
| Who the grant covers | Only the admin signing in | **Every user in the tenant** |
| Grant record created | `consentType: Principal` | `consentType: AllPrincipals` |
| Lifetime | Tied to your account | **Persists after your engagement ends** |
| Removal | Nothing to clean up | Must be manually revoked |

**Why it matters.** The delegated permissions the script asks for (`Device.ReadWrite.All`,
`Directory.ReadWrite.All`) stay bounded by each user's own directory roles, so ticking the box does
not hand standard users admin powers. What it *does* do is pre-consent a highly privileged
first-party app for the entire tenant, permanently. That removes the per-user consent gate for
every future sign-in to that app by anyone — including an attacker who phishes a privileged
account, who then finds the tooling already approved and raises no consent prompt. It also shows up
in a tenant's consent-grant audit as an unexplained org-wide grant.

The correct scope for a monthly maintenance task is **you, for this task, in this tenant**.

### What the dialog looks like

```
  Permissions requested
  Microsoft Graph Command Line Tools
  unverified

  This application is not published by Microsoft.

  This app would like to:
    ✔ Read and write devices
    ✔ Read and write directory data
    ✔ Sign in and read user profile
    ✔ Maintain access to data you have given it access to

  ☐  Consent on behalf of your organization      <-- LEAVE THIS UNCHECKED
                                                     (it may be pre-ticked; untick it)

           [ Cancel ]   [ Accept ]
```

Click **Accept** with the box clear. If the box appears pre-ticked, untick it before accepting.

> **Note:** The checkbox only appears for accounts that *can* grant org-wide consent (Global
> Administrator, Privileged Role Administrator, Cloud Application Administrator). If you don't see
> it, there is nothing to worry about.

### If someone ticks it by mistake

The grant is reversible. In the client tenant:

**Entra admin center →** Identity → Applications → Enterprise applications → search
*Microsoft Graph Command Line Tools* → **Permissions** → *Admin consent* tab → review the listed
permissions → **Revoke permission**.

Before revoking, confirm the grant wasn't already there deliberately — some tenants legitimately
admin-consent this app so their own staff can use Graph PowerShell. Revoking a pre-existing,
intentional grant will break that tooling for everyone in the tenant. If the grant predates your
run, leave it alone and note it in your ticket.

To check without changing anything (run in a separate session, requires its own consent):

```powershell
Connect-MgGraph -Scopes 'Directory.Read.All' -TenantId contoso.onmicrosoft.com
$sp = Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'"
Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" |
    Select-Object ConsentType, PrincipalId, Scope
Disconnect-MgGraph
```

`ConsentType` of `AllPrincipals` means org-wide consent exists. `Principal` means it is scoped to
individual users, which is what you want.

---

## Requirements

| Requirement | Detail |
|---|---|
| PowerShell | 7.x recommended (works on 5.1) |
| Modules | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement` |
| Graph scopes | `Device.ReadWrite.All`, `Directory.ReadWrite.All` |
| Entra role | Cloud Device Administrator (minimum), Intune Administrator, or Global Administrator |

The script checks for the modules on startup and offers to install them for the current user if
they're missing. Decline and it exits without connecting to anything.

---

## Usage

### Standard monthly run

```powershell
.\Monthly_Device_Inventory_Review.ps1
```

### Pin the tenant explicitly (recommended when working across clients)

```powershell
.\Monthly_Device_Inventory_Review.ps1 -TenantId contoso.onmicrosoft.com
```

Your browser may still hold an SSO cookie from the previous client and silently sign you in as the
wrong account. Passing `-TenantId` forces the correct tenant, and the script prints the resolved
account and tenant immediately after sign-in — **read that line before answering any prompt.**

### Dry run (see the lists, change nothing)

```powershell
.\Monthly_Device_Inventory_Review.ps1 -TenantId contoso.onmicrosoft.com
# answer N to the disable prompt
# answer N to the delete prompt
```

You still get a full CSV, with every device marked `Declined by operator`. Useful for sending a
client the proposed changes before acting. A second run the same day won't overwrite the first
report — it gets a `_2` suffix.

### Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-TenantId` | *(prompted)* | Tenant ID or domain to sign in to |
| `-DisableAfterDays` | `90` | Inactivity threshold for the disable pass |
| `-DeleteAfterDays` | `120` | Inactivity threshold for the delete pass |
| `-CsvPath` | *(auto)* | Override the report path |
| `-IncludeNeverSignedIn` | *(off)* | Also act on devices with no sign-in data at all |
| `-LineDelayMilliseconds` | `60` | Output pacing; `0` for full speed |

### Other examples

```powershell
# Stricter thresholds for a tightly managed tenant
.\Monthly_Device_Inventory_Review.ps1 -DisableAfterDays 60 -DeleteAfterDays 90

# Include never-signed-in registrations (read the caveat below first)
.\Monthly_Device_Inventory_Review.ps1 -IncludeNeverSignedIn

# Full speed, custom report location
.\Monthly_Device_Inventory_Review.ps1 -LineDelayMilliseconds 0 -CsvPath 'C:\Reports\contoso.csv'
```

---

## How it works

### 1. Module check and sign-in

Confirms both Graph modules are available (offering to install them if not), then clears any
**leftover Graph session** before connecting. This matters when reviewing several tenants in one
PowerShell window: without it, `Connect-MgGraph` can quietly reuse the previous client's session
and you'd act on the wrong tenant. If the leftover session can't be cleared, the script **aborts
before touching anything**.

After connecting it prints the account, tenant ID, granted scopes, and the report filename it will
use — all before any change is made.

### 2. Inventory

Retrieves every device object with `Get-MgDevice -All`, requesting only the fields it needs.
Rather than assigning the result, it **pipes** it into `ForEach-Object`, so each page renders as it
arrives from Graph instead of after the whole fetch completes. Each device prints on its own line,
colour-coded by staleness:

```
  Legend: active | 90+ days idle | 120+ days idle | no sign-in data

  [    1]  LAPTOP-FINANCE-01                            5 days idle  (last sign-in 2026-08-28)
  [    2]  DESKTOP-RECEPTION                           95 days idle  (last sign-in 2026-05-30)
  [    3]  SURFACE-CEO                                130 days idle  (last sign-in 2026-04-25)
  [    4]  (no display name)                        no sign-in data
```

| Colour | Meaning |
|---|---|
| Grey | Active within the disable threshold |
| Yellow | Idle past `-DisableAfterDays` |
| Red | Idle past `-DeleteAfterDays` |
| Magenta | No sign-in data recorded |

Inactivity is measured from **`approximateLastSignInDateTime`** in UTC. Microsoft only updates this
value approximately (it can lag by hours to a day or so), which is why a device sitting exactly on
the boundary may move between buckets between runs.

### 3. Pass 1 — disable (90+ days)

Prints the 90+ day list as a table that builds up row by row, then asks:

```
Disable the 12 enabled device(s) listed above? [Y/N]
```

On **Y**, each device is disabled with `Update-MgDevice` setting `accountEnabled = $false`, one at
a time, printing `[ OK ]` or `[FAIL]` per device plus a summary. A failure on one device does not
stop the rest of the run — the reason is captured per device in the CSV.

Devices already disabled are **skipped, not re-disabled**, and recorded as `Already Disabled`, so
month-over-month runs don't inflate the success count.

### 4. Pass 2 — delete (120+ days)

Prints the 120+ day list, warns that deletion is permanent, then asks twice:

```
Delete the 4 device(s) listed above? [Y/N]
Type DELETE (all caps) to permanently remove 4 device object(s)
```

The second confirmation is case-sensitive — anything other than `DELETE` cancels the pass and marks
every device `Declined by operator`. On confirmation, each device is removed with `Remove-MgDevice`,
again reporting per-device success or failure.

**The 120+ day set is a subset of the 90+ day set.** If you approve both passes, a device you delete
was disabled moments earlier in the same run. The CSV records both actions on one row, which is what
you want for an audit trail — just don't be surprised to see it.

### 5. CSV report

Every device that met **either** threshold gets one row, keyed internally by object ID so the
disable and delete outcomes land on the same record. Named from the domain of the signed-in
account:

| Signed in as | Report |
|---|---|
| `admin@contoso.com` | `Contoso_Device_Review_2026-09-02.csv` |
| `admin@contoso.onmicrosoft.com` | `Contoso_Device_Review_2026-09-02.csv` |
| `admin@contoso.co.uk` | `Contoso_Device_Review_2026-09-02.csv` |

Columns:

| Column | Notes |
|---|---|
| `DisplayName`, `DeviceId`, `ObjectId` | Device identity; `ObjectId` is what the Graph calls targeted |
| `OperatingSystem`, `OperatingSystemVersion` | As reported by the device |
| `TrustType` | `AzureAd` (cloud), `ServerAd` (hybrid), `Workplace` (registered) |
| `ProfileType` | e.g. `RegisteredDevice` |
| `IsManaged`, `IsCompliant` | Intune state at the time of the run |
| `RegistrationDateTime` | UTC |
| `LastSignInUtc` | UTC, or `never` |
| `DaysInactive` | Whole days since last sign-in; blank when never |
| `AccountEnabledBefore` | Enabled state *before* the script acted |
| `Inactive90Plus`, `Inactive120Plus` | Which bucket(s) the device fell into — column names follow your threshold parameters |
| `DisableAttempted`, `DisableResult` | `Success` / `Failed` / `Already Disabled` / `Declined by operator` |
| `DeleteAttempted`, `DeleteResult` | `Success` / `Failed` / `Declined by operator` |
| `ErrorMessage` | Graph's reason when an action failed |
| `ActionTimestampUtc` | When the action was attempted |

If no device met either threshold, no CSV is written and the script says so.

### 6. Sign-out

Signs out with `Disconnect-MgGraph`, then **re-checks `Get-MgContext` to confirm the session is
actually gone** rather than trusting the cmdlet's silence, and prints the verdict:

```
  [ OK ]   Signed out of tenant 1234abcd-... (admin@contoso.com)

Sign-out successful - safe to run this script against the next tenant.
```

If sign-out fails you get a red `Sign-out UNSUCCESSFUL` and instructions to run
`Disconnect-MgGraph` manually **before** starting the next tenant.

A script-scope `trap` guarantees this runs even if the script dies unexpectedly partway through, so
a failed run never leaves a live session behind for the next client.

---

## Operational caveats

**Never-signed-in devices are excluded by default.** Null sign-in data usually means a stale
registration *or* a brand-new one, not 90 days of inactivity. They're printed as a separate warning
block and skipped. `-IncludeNeverSignedIn` folds them into both passes — check what they actually
are first.

**Hybrid-joined devices come back.** Deleting an Entra object with `TrustType = ServerAd` won't
stick; Entra Connect resyncs it on the next cycle. Permanent removal has to happen in on-prem AD.
Filter the CSV by `TrustType` to see which devices this affects.

**Deleting a device object has user-visible consequences.** Affected users must re-register or
re-enrol, Windows Hello for Business and BitLocker recovery key escrow tied to that object are
lost, and Conditional Access policies keyed on device compliance will block the device until it
re-registers. Treat deletion as irreversible: device objects are not covered by the 30-day
recoverable-items restore that applies to users and groups, so re-registration — not restore — is
the recovery path. Confirm current behaviour against Microsoft's docs before relying on any undo.

**Disabling is the safer lever.** A disabled device blocks CA-gated access but keeps the object,
so it's trivially reversible. Consider running disable-only for a month, letting anything that
comes back re-enable itself organically, and only deleting on the following pass.

**Approximate timestamps.** `approximateLastSignInDateTime` is not a precise audit signal. Don't
use these thresholds as the only input for a device the client considers important.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Insufficient privileges to complete the operation` | Your account lacks a device-write role. Cloud Device Administrator or higher. |
| Consent dialog reappears every run | Normal when consent is per-user and you use a different account. **Do not** tick the org-wide box to stop it. |
| Signed in as the wrong tenant | Browser SSO reused a cookie. Use `-TenantId`, or `Connect-MgGraph -UseDeviceCode` to bypass browser SSO. |
| `Sign-out UNSUCCESSFUL` | Run `Disconnect-MgGraph` manually, confirm `Get-MgContext` returns nothing, then start the next tenant. |
| Run feels slow on a large tenant | `-LineDelayMilliseconds 0`. Pacing self-caps at 45 seconds total regardless. |
| Deleted devices reappear next month | Hybrid-joined — they resync from on-prem AD. Remove them in AD. |
| Fewer devices listed than expected | Only `approximateLastSignInDateTime` is considered; devices with no sign-in data are excluded unless `-IncludeNeverSignedIn`. |

---

## Suggested monthly routine

1. `.\Monthly_Device_Inventory_Review.ps1 -TenantId <client>` and answer **N** to both prompts.
2. Send the CSV to the client contact for sign-off.
3. Re-run and approve the passes the client agreed to.
4. Confirm `Sign-out successful` before moving to the next tenant.
5. File both CSVs against the ticket — the declined one is the proposal, the second is the record
   of what was actually changed.
