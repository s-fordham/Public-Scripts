# Active Directory Extension Attribute Migration

Safely exports and migrates on-premises Active Directory user `extensionAttribute` values, then provides a separately invoked, verification-gated cleanup operation.

## Files

| File | Purpose |
| --- | --- |
| `Move-ADUserExtensionAttributes.ps1` | Exports all 15 attributes, copies the mapped values from a CSV snapshot, verifies changes, and optionally clears verified source values. |

## Purpose

Use this script to free `extensionAttribute1` through `extensionAttribute7` by copying their current values into the seven attributes from `extensionAttribute9` through `extensionAttribute15` according to the supplied mapping.

The script separates the work into three explicit switches:

- `-Export` creates a timestamped snapshot of all 15 extension attributes for every user in scope.
- `-Copy` imports a script-generated snapshot and copies populated source values into the mapped targets. It never clears source values.
- `-ClearSource` is an isolated cleanup run with an all-or-nothing, per-user fail-safe.
- `-TestUser` limits any of those operations to one explicitly selected AD user.

## Attribute mapping

| Source | Target |
| --- | --- |
| `extensionAttribute1` | `extensionAttribute10` |
| `extensionAttribute2` | `extensionAttribute15` |
| `extensionAttribute3` | `extensionAttribute9` |
| `extensionAttribute4` | `extensionAttribute14` |
| `extensionAttribute5` | `extensionAttribute13` |
| `extensionAttribute6` | `extensionAttribute11` |
| `extensionAttribute7` | `extensionAttribute12` |

`extensionAttribute8` is not changed.

## When to use it

- The attributes are stored in on-premises Active Directory.
- On-premises AD is the authoritative source for attributes synchronised to Entra ID or Exchange Online.
- You need an auditable CSV snapshot before changes.
- You want source cleanup to occur only after an exact live comparison confirms the copy.

## When not to use it

- Do not use this script for cloud-only Exchange Online `CustomAttribute1` through `CustomAttribute15` values.
- Do not use it for Entra ID directory extensions that are not backed by these on-premises AD attributes.
- Do not run cleanup until a copy run has completed, directory synchronisation requirements have been considered, and the copy result report has been reviewed.
- Do not commit generated snapshots or reports. They can contain personal and customer data.

## Prerequisites

- A domain-joined Windows machine, or a Windows machine with connectivity to a domain controller.
- RSAT Active Directory tools.
- Windows PowerShell 5.1, or PowerShell 7 on Windows using Windows PowerShell compatibility for the Active Directory module.
- Secure storage for the generated CSV files.

## Required roles and permissions

No Microsoft 365 or Entra administrative role is required because the script works against on-premises AD.

The running account needs:

- Read access to the user objects and `extensionAttribute1` through `extensionAttribute15` for all operations.
- Write access to `extensionAttribute9` through `extensionAttribute15` for `-Copy`.
- Clear/write access to `extensionAttribute1` through `extensionAttribute7` for `-ClearSource`.

Delegate only the required attribute-level permissions where practical.

## Required PowerShell module

- `ActiveDirectory`

Install the appropriate RSAT Active Directory tools for the operating system. The script does not install modules automatically.

## Recommended workflow

### Test with one user first

Supply an identity accepted by `Get-ADUser -Identity`, such as a sAMAccountName, distinguished name, ObjectGUID, or SID.

Export only that user's 15 extension attributes:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Export `
    -TestUser "test.user" `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration\Test"
```

This creates a file named similarly to `ADUserExtensionAttributes-TestUser-Export-YYYYMMDD-HHMMSSfff.csv` containing exactly one row.

Select that test CSV and preview copying only the same user:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Copy `
    -TestUser "test.user" `
    -InputCsvPath "C:\Secure\ExtensionAttributeMigration\Test\ADUserExtensionAttributes-TestUser-Export-20260820-100000000.csv" `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration\Test" `
    -WhatIf `
    -Verbose
```

The script refuses to continue if the selected test CSV has more than one row or if its `ObjectGuid` does not match `-TestUser`. After reviewing the `-WhatIf` reports, remove `-WhatIf` to perform the one-user copy.

The cleanup fail-safe can also be tested with only that user:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -ClearSource `
    -TestUser "test.user" `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration\Test" `
    -WhatIf
```

### 1. Export a baseline

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Export `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration"
```

The CSV contains identity fields and `extensionAttribute1` through `extensionAttribute15` for every user in scope.

### 2. Preview export and copy together

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Export `
    -Copy `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration" `
    -WhatIf `
    -Verbose
```

This creates a fresh snapshot, imports that file as the desired source state, and reports the changes that would be made. `-WhatIf` still writes audit snapshots and reports; it does not write to AD.

### 3. Perform the copy

Interactive confirmation is enabled because the script has high confirmation impact:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Export `
    -Copy `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration"
```

For a reviewed large run, suppress per-user confirmation explicitly:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Export `
    -Copy `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration" `
    -Confirm:$false
```

You can also import an earlier script-generated snapshot:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -Copy `
    -InputCsvPath "C:\Secure\ExtensionAttributeMigration\ADUserExtensionAttributes-Export-20260819-190000000.csv" `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration" `
    -WhatIf
```

When an earlier snapshot is supplied, the script also creates a fresh `PreCopy` snapshot. If any source value changed after the input snapshot, the entire user is skipped.

### 4. Review copy results

Review the copy summary and detailed results. Resolve all `SkippedTargetConflict`, `SkippedSourceChanged`, `Failed`, and `PostChangeVerificationFailed` records before considering cleanup.

The copy operation does not alter `extensionAttribute1` through `extensionAttribute7`.

### 5. Preview the separate cleanup operation

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -ClearSource `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration" `
    -WhatIf `
    -Verbose
```

### 6. Clear only verified source values

Interactive confirmation:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -ClearSource `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration"
```

After reviewing the `-WhatIf` output, a large approved run can suppress per-user confirmation:

```powershell
.\Move-ADUserExtensionAttributes.ps1 `
    -ClearSource `
    -OutputFolder "C:\Secure\ExtensionAttributeMigration" `
    -Confirm:$false
```

## Cleanup fail-safe

Cleanup is deliberately stricter than a per-attribute best effort:

1. A fresh snapshot of every user in scope is written before any cleanup.
2. Each user is read again immediately before evaluation. If any mapped source or target changed after the pre-cleanup snapshot, that user is skipped so the backup is not stale.
3. For a user, each populated source attribute is compared with its mapped target using an exact ordinal string comparison.
4. Blank source attributes are ignored because there is nothing to clear.
5. If any populated pair differs, none of that user's source attributes are cleared.
6. If every populated pair matches, all populated source attributes for that user are cleared in one `Set-ADUser` call.
7. The script reads the user again and verifies that the sources are empty and the target values remain unchanged.

`-ClearSource` cannot be combined with `-Export` or `-Copy`, which prevents copy and cleanup from occurring in the same command.

## Copy safeguards

- The CSV must contain the script's supported snapshot schema and a valid, unique `ObjectGuid` for every row.
- With `-TestUser`, the input CSV must contain exactly one row and its `ObjectGuid` must match the selected live AD user.
- Test-user validation completes before any AD write is attempted.
- Users are resolved by immutable AD `ObjectGuid`, not by display name or UPN.
- When `-SearchBase` is specified, a user currently outside that scope is skipped.
- If a live source value differs from the imported snapshot, the entire user is skipped.
- A populated target containing a different value causes the entire user to be skipped by default.
- `-OverwriteTarget` is required to replace conflicting target values. Review conflicts in a `-WhatIf` result report before using it.
- Every write is read back from AD and verified.
- Source attributes are never cleared during `-Copy`.

## Parameters

| Parameter | Description |
| --- | --- |
| `-Export` | Export all 15 attributes for all users in scope. |
| `-Copy` | Import a snapshot and copy populated source values to mapped targets. |
| `-ClearSource` | Separately verify and clear source attributes. |
| `-InputCsvPath` | Script-generated input snapshot; required for `-Copy` unless combined with `-Export`. |
| `-SearchBase` | Optional OU/container distinguished name. |
| `-TestUser` | Limit export, copy, or cleanup to one AD identity. Test-user copy requires an exactly matching one-row CSV. |
| `-OutputFolder` | Secure destination for snapshots and reports. Defaults to `.\Output`. |
| `-Server` | Optional domain controller or AD LDS instance. |
| `-Credential` | Optional AD credential. |
| `-OverwriteTarget` | Explicitly allow replacement of conflicting mapped targets during copy. |
| `-DisableProgress` | Suppress progress bars. |
| `-WhatIf` | Preview AD changes while still writing audit reports. |
| `-Confirm` | Control per-user confirmation for AD changes. |

## Output files

Filenames include a millisecond timestamp and are not intentionally overwritten.

| File pattern | Purpose |
| --- | --- |
| `ADUserExtensionAttributes-Export-*.csv` | Standalone export of all 15 attributes. |
| `ADUserExtensionAttributes-ExportAndCopyInput-*.csv` | Fresh export used immediately as copy input. |
| `ADUserExtensionAttributes-TestUser-Export-*.csv` | One-user test export selected later with `-InputCsvPath`. |
| `ADUserExtensionAttributes-TestUser-ExportAndCopyInput-*.csv` | One-user export used immediately by `-Export -Copy -TestUser`. |
| `ADUserExtensionAttributes-TestUser-PreCopy-*.csv` | Fresh one-user snapshot before copying from an earlier test CSV. |
| `ADUserExtensionAttributes-TestUser-PostCopy-*.csv` | One-user snapshot after a verified test copy. |
| `ADUserExtensionAttributes-TestUser-PreClear-*.csv` | Mandatory one-user snapshot before test cleanup. |
| `ADUserExtensionAttributes-TestUser-PostClear-*.csv` | One-user snapshot after verified test cleanup. |
| `ADUserExtensionAttributes-PreCopy-*.csv` | Fresh live snapshot created when copying from an earlier input CSV. |
| `ADUserExtensionAttributes-PostCopy-*.csv` | Full scoped snapshot after at least one verified copy. |
| `ADUserExtensionAttributes-Copy-Results-*.csv` | Per-user copy status and checked mappings. |
| `ADUserExtensionAttributes-Copy-Summary-*.csv` | Copy totals and paths to related audit files. |
| `ADUserExtensionAttributes-PreClear-*.csv` | Mandatory full snapshot written before cleanup. |
| `ADUserExtensionAttributes-PostClear-*.csv` | Full scoped snapshot after at least one verified cleanup. |
| `ADUserExtensionAttributes-ClearSource-Results-*.csv` | Per-user cleanup status and fail-safe outcome. |
| `ADUserExtensionAttributes-ClearSource-Summary-*.csv` | Cleanup totals and paths to related audit files. |

Important status values include:

| Status | Meaning |
| --- | --- |
| `CopiedAndVerified` | Target values were written and verified. |
| `AlreadyCopied` | Every populated source already matched its target. |
| `ClearedAndVerified` | All populated pairs matched, then the sources were cleared and verified. |
| `SkippedSourceChanged` | Live source data differed from the imported snapshot. |
| `SkippedTargetConflict` | A mapped target had a different populated value. |
| `SkippedChangedAfterPreClearSnapshot` | A mapped value changed after the mandatory cleanup snapshot, so the user was skipped. |
| `SkippedVerificationFailed` | Cleanup fail-safe comparison failed; no source was cleared for the user. |
| `PostChangeVerificationFailed` | A write completed but the read-back verification failed. Investigate immediately. |
| `WhatIf` | The change passed preconditions but AD was not modified. |
| `NoSourceValues` | No source values needed copying or clearing. |

## Rollback notes

- Copy is non-destructive to the source attributes. If copied target values must be removed, use the pre-change snapshot to identify exactly which targets were blank before the run, then perform a separately reviewed target cleanup. Do not clear all target attributes indiscriminately.
- Cleanup creates a mandatory `PreClear` snapshot containing the original source and target values. If restoration is required, use `ObjectGuid` from that snapshot and restore only the affected source fields after a change review.
- Automatic rollback is intentionally not included because a later legitimate attribute change could otherwise be overwritten. Treat restoration as a separate, approved operation using the snapshot as evidence.

## Limitations

- This script does not trigger or monitor Entra Connect / Cloud Sync. Allow for the normal directory synchronisation cycle before downstream validation.
- It processes AD user objects only, not contacts, groups, computers, or mailboxes that do not have corresponding AD users.
- Exact comparisons are case-sensitive and preserve whitespace. This is intentional for cleanup safety.
- Large directories can take time because change operations read each user before and after a write.
- A `PostChangeVerificationFailed` status cannot undo a completed AD write; use the pre-change snapshot for investigation and recovery.

## Change history

| Date | Change |
| --- | --- |
| 2026-08-19 | Initial export, CSV-driven copy, verification, and fail-safe cleanup implementation. |
| 2026-08-20 | Added one-user test scoping and strict one-row test CSV validation. |
