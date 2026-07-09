# Entra Provisioning Readiness Scripts

This folder contains two report-first PowerShell scripts created to support Workday to Active Directory provisioning readiness and GAL duplicate investigation.

## Scripts

| Script | Purpose |
|---|---|
| `Invoke-HRToADEmployeeIDMatch.ps1` | Imports an HR / Workday CSV, matches workers to AD users, optionally writes `employeeID`, and can report Workday-vs-AD attribute mismatches before provisioning is enabled. |
| `Find-MailContactADUserDuplicates.ps1` | Audits AD for potential duplicate GAL entries where a MailContact and AD User / MailUser appear to represent the same person. |

## Safety model

Both scripts are designed to be safe for discovery and readiness reporting.

`Invoke-HRToADEmployeeIDMatch.ps1` only writes to AD when `-Apply` is explicitly supplied. By default, it only attempts to update the `employeeID` attribute on matched AD user objects. It does not update Workday-managed attributes such as names, title, department, manager, telephone numbers, address fields, preferred language, or OU placement.

MailContact fallback matches in `Invoke-HRToADEmployeeIDMatch.ps1` are report-only and are not update-eligible for `employeeID`. MailUsers are AD user objects in hybrid Exchange environments and can be update-eligible when they are matched or approved.

Approved MailContacts can be deleted only when `-DeleteApprovedMailContacts` and `-Apply` are supplied together. For each row, the script requires `ReviewDecision` to be `Approved`, resolves `ApprovedADDistinguishedName`, confirms the resolved object is a mail-enabled contact, confirms the HR row is represented by a single matched AD user object, and skips deletion if the DN resolves to an AD User, MailUser, non-mail-enabled contact, no object, multiple objects, a duplicated approval, or no single representing AD user.

`Find-MailContactADUserDuplicates.ps1` is always report-only and makes no AD changes.

## Requirements

- Windows PowerShell or PowerShell 7 with the Active Directory module available.
- Domain connectivity and permissions to read AD user/contact attributes.
- Optional: `ImportExcel` PowerShell module for `.xlsx` output.

If `ImportExcel` is missing, the scripts try to install it for the current user unless `-DoNotInstallMissingModules` is specified. If installation fails, the scripts fall back to CSV output.

## Invoke-HRToADEmployeeIDMatch.ps1

### Default HR / Workday CSV columns

The script defaults are aligned to the current Workday export column names:

```powershell
[string]$CsvEmailColumn      = "PrimaryWorkEmail"
[string]$CsvFirstNameColumn  = "PreferredFirstName"
[string]$CsvLastNameColumn   = "PreferredLastName"
[string]$CsvJobTitleColumn   = "BusinessTitle"
[string]$CsvEmployeeIdColumn = "WorkerID"
[string]$CsvDepartmentColumn = "Department"
[string]$CsvManagerReferenceColumn = "ManagerReference"
[string]$CsvManagerEmailColumn = "PrimaryWorkEmail_Manager"
```

These can be overridden at runtime if the HR export uses different column names.

For last-name matching and mismatch checks, `PreferredLastName` is used when populated. If `PreferredLastName` is blank for a row, the script falls back to the `LastName` column for that row. Report output includes `HR_LastNameSource` so reviewed rows show which source column was used.

### Switch reference

| Switch | Type | Default | Explanation |
|---|---|---|---|
| `-Apply` | `[switch]` | Off | Writes eligible `employeeID` updates to matched AD Users and MailUsers. Without this switch, the script is a dry run. |
| `-CheckSystemMismatch` | `[switch]` | Off | Creates a report-only comparison of selected HR columns against attributes on the matched AD object. |
| `-CheckMailContactReview` | `[switch]` | Off | Enables the slower extra MailContact lookup for rows that already matched an AD User or MailUser. |
| `-ReserveExactMatches` | `[switch]` | Off | Enables the slower pre-reservation scan that stops exact AD matches elsewhere in the file appearing as potential matches on other rows. |
| `-ApprovedOnly` | `[switch]` | Off | Processes only rows where `ReviewDecision` is `Approved`; resolves `ApprovedADDistinguishedName` directly and skips WorkerID, email, name, and MailContact matching. Writes still require `-Apply`. |
| `-DeleteApprovedMailContacts` | `[switch]` | Off | Deletes approved MailContacts only when also used with `-Apply`; each target is verified as a mail-enabled contact and represented by one matched AD user. |
| `-DoNotInstallMissingModules` | `[switch]` | Off | Stops the script trying to install `ImportExcel`; CSV fallback is used if the module is unavailable. |

### Matching order

1. Match AD user by email-style values:
   - `mail`
   - `userPrincipalName`
   - `proxyAddresses`
   - `targetAddress`
2. If no AD user email match exists, match AD user by:
   - first name
   - last name
   - title
3. If no exact AD user name match exists, report name-variant AD user candidates where:
   - title matches
   - first names match exactly, by token, by prefix, or by similarity
   - last names match exactly or by token
4. If no title-matching name variant exists, report lower-confidence name-variant candidates where title differs or is missing.
5. Name-variant matching considers `givenName`, `sn`, `displayName`, `cn`, `name`, and the CN portion of `distinguishedName` where available.
6. If no AD user match exists, check MailContacts and MailUsers using the same email/name/title and name-variant approach.

### Common examples

#### Fast readiness report

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx"
```

This is the recommended first run for a large HR export. It is a dry run and makes no AD changes.

#### Readiness report with slower MailContact duplicate review

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -CheckMailContactReview
```

Use this when you specifically want to find MailContacts for rows that already matched AD users. It can be slow on large exports because it performs extra MailContact lookups.

#### Conservative review run with exact-match reservation

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -ReserveExactMatches
```

Use this when you want the slower pre-reservation pass that keeps exact AD matches from appearing as potential matches for other rows.

#### System mismatch report

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -SystemMismatchReportPath "C:\Temp\Workday-AD-System-Mismatches.xlsx" `
  -CheckSystemMismatch
```

This remains report-only. It writes the main readiness report and a separate system mismatch report.

Expected workbook outputs:

- Main report: `HR AD Match`
- MailContact review, when MailContacts are found: `MailContact Review`
- Separate system mismatch report: `System Mismatches`

When Excel output is available, the main report workbook is recreated on each run so worksheets from earlier runs with different switches do not remain in the file.

If `-SystemMismatchReportPath` is not supplied, the script creates a separate mismatch report beside the main report using the main file name plus `-System-Mismatches`.

The console output includes timestamped phase markers such as loading the AD module, importing the CSV, reserving AD matches, matching HR users, checking MailContacts for review, and exporting reports. If a run appears to stop, the last timestamped phase shown identifies the quiet or slow section to investigate.

For performance on large HR exports, the default run skips the slower exact-match pre-reservation scan. This does not change the primary matching order, but it can allow a later exact AD user to appear as a potential candidate on an earlier row. Use `-ReserveExactMatches` only when you want the slower, more conservative potential-match suppression.

`-CheckSystemMismatch` does not perform a separate user match. It compares the HR row against the AD object selected by the main match logic. Manual `ApprovedADDistinguishedName` decisions are used first, then WorkerID to AD `employeeID`, then email and name-based matching.

The current system mismatch comparison is deliberately limited to:

- `PreferredFirstName` -> AD `givenName`
- `PreferredLastName` fallback to `LastName` -> AD `sn`
- `BusinessTitle` -> AD `title`
- `Department` -> AD `department`
- `ManagerReference` -> AD `manager`

Email-style values are still used for matching, but they are excluded from the system mismatch report because mail attributes may only be managed during account creation.

For manager mismatches, the HR `ManagerReference` value is treated as the manager's WorkerID. The script resolves it to a proposed AD manager DN by first looking for an AD user whose `employeeID` equals that WorkerID. If that is not available, it finds the manager's own HR row and resolves that row to AD by email, then by first name, last name, and title. If the manager reference cannot resolve to exactly one AD user, the mismatch report flags it as `Manager reference unresolved` rather than guessing.

The mismatch report includes the HR first name, last name, business title, department, and manager reference; the selected AD user's UPN, current AD `employeeID`, given name, surname, job title, department, and manager DN; the current AD value; the proposed HR or resolved DN value; and a `WhatWillChange` summary for each differing attribute.

When `-CheckSystemMismatch` is used, the console output includes a compact `System mismatch summary` that counts matched users with HR/AD attribute mismatches and matched users with no HR/AD attribute mismatches. The system mismatch check is report-only and only compares HR columns against AD attributes on the AD object selected by the main match logic.

Worker ID, ManagerReference, and `employeeID` values are exported as text in the Excel workbook so leading zeros, such as `001234`, are preserved.

### MailContact review

When a MailContact is found, the main workbook includes a `MailContact Review` sheet. It separates two scenarios:

- `MailContact found - no matched AD user in this run`: the worker may be missing an AD account, or the correct AD user still needs to be manually approved/matched.
- `MailContact cleanup candidate - represented by matched AD user`: the worker already has a matched AD User or MailUser, so the MailContact is likely a duplicate GAL object and can be reviewed for deletion.

By default, this sheet reports MailContacts that were already found by the normal matching flow. To also search for MailContact duplicates for rows that already matched an AD User or MailUser, add `-CheckMailContactReview`. That extra check can be slow on large exports because it performs additional MailContact lookups.

The review sheet includes the MailContact distinguished name, how it was matched, the HR row details, and the matched AD user details where one exists.

### Apply EmployeeID updates

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -Apply
```

Only matched and update-eligible AD user objects are updated. Hybrid MailUsers are treated as AD user objects for `employeeID` updates. Rows matched by `WorkerID + AD employeeID` are reported as `Matched - By WorkerID`; because AD `employeeID` already matches the HR WorkerID, no write is attempted and they are excluded from the unapproved / still requiring review summary. MailContacts, ambiguous matches, and potential title-mismatch matches are skipped.
When `-Apply` is used, the main workbook also includes an `EmployeeID Apply Results` sheet. It lists AD User and MailUser rows as `Updated EmployeeID`, `Already had EmployeeID`, or `Skipped`, with the AD object identity, `EmployeeIDBefore`, `EmployeeIDAfter`, action result, and skip reason. MailContacts are excluded from this sheet because they cannot receive `employeeID` updates. `Already had EmployeeID` means the AD user already had the requested WorkerID value in `employeeID`, so no write was attempted.
If Excel export is unavailable, the apply results are written beside the main CSV as `*.EmployeeID-Apply-Results.csv`.

For the fastest apply pass after manual review, use a reviewed CSV with `-ApprovedOnly -Apply`:

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Reviewed-HR-AD-Match.csv" `
  -ReportPath "C:\Temp\Reviewed-HR-AD-Apply-Results.xlsx" `
  -ApprovedOnly `
  -Apply
```

`-ApprovedOnly` processes only rows where `ReviewDecision` is `Approved`, resolves `ApprovedADDistinguishedName` directly, and skips WorkerID, email, name, and MailContact matching. If `-ApprovedOnly` is used without `-Apply`, the script remains a dry run and reports what would be updated.
Potential name-variant rows are also skipped and are intended for manual review only.
Potential name-variant title-mismatch rows are lower-confidence and are also skipped.
When name-variant matching finds multiple candidates for the same CSV row, the script keeps one report row and lists the candidate count and details in `PotentialMatchCount` and `PotentialMatches`.
Name-variant matching ignores weak surname particles such as `de`, `da`, `van`, and `von`; these words are not enough on their own to create a surname match.
Before producing potential name-variant candidates, the script always reserves approved AD user review decisions from the input. Use `-ReserveExactMatches` when you also want the slower scan that reserves exact email or exact name-and-title matches elsewhere in the file. Reserved AD users are not reused as potential matches for other CSV rows.
Fuzzy first-name matching is intentionally conservative; prefix and token matches are preferred, and pure similarity matches must be strong.
If a row is clearly the wrong person, leave it unapproved or set `ReviewDecision` to a non-approved value such as `Rejected` and add context in `ApprovalNotes`.
If the same AD user is already matched to an earlier CSV row, later update-eligible matches to that same AD user are reported as `Duplicate - AD User Already Matched` and skipped. When duplicate AD user matches exist, they are also exported to a dedicated `Duplicate AD User Matches` worksheet in the main readiness workbook. If Excel export is unavailable, CSV fallback writes `*.Duplicate-AD-User-Matches.csv`.

### Delete approved MailContacts

When MailContacts have been replaced by AD accounts, use the reviewed CSV to approve the MailContact distinguished names for cleanup:

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Reviewed-HR-AD-Match.csv" `
  -ReportPath "C:\Temp\Reviewed-HR-AD-Apply-Results.xlsx" `
  -DeleteApprovedMailContacts `
  -Apply
```

The CSV must include `ReviewDecision` and `ApprovedADDistinguishedName`. Only rows where `ReviewDecision` is `Approved` are considered. Before deletion, the script resolves `ApprovedADDistinguishedName`, confirms the object is a mail-enabled contact, and confirms the HR row is already represented by a single matched AD user object. The main workbook includes a `MailContact Delete Results` sheet only when both `-DeleteApprovedMailContacts` and `-Apply` are used. The sheet shows deleted and skipped outcomes, including the matched AD user that made the MailContact safe to delete.

### ReviewDecision workflow

The `HR AD Match` report includes manual review columns:

- `ReviewDecision`
- `ApprovedADDistinguishedName`
- `ApprovedBy`
- `ApprovalNotes`

Use these columns only for rows that should be manually accepted, such as a single reviewed potential name-variant match. To approve a row for a later `-Apply` run:

1. Set `ReviewDecision` to `Approved`.
2. Set `ApprovedADDistinguishedName` to the exact AD user distinguished name that should receive the `employeeID`.
3. Fill `ApprovedBy` and `ApprovalNotes` for audit context.
4. Save the reviewed `HR AD Match` sheet as CSV.
5. Rerun the script using that reviewed CSV as `-CsvPath`.

The script also accepts the report-style `HR_` column names when reading a reviewed `HR AD Match` CSV.

Use `ReviewDecision` value `Ignore` for reviewed rows where no action should be taken, such as workers who need new AD accounts or rows where no acceptable AD user exists. Ignored rows do not require `ApprovedADDistinguishedName`, are not matched, are never updated by `-Apply`, and are excluded from the unapproved summary.

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Reviewed-HR-AD-Match.csv" `
  -ReportPath "C:\Temp\Reviewed-HR-AD-Apply-Results.xlsx" `
  -Apply
```

Approved review rows are update-eligible when `ReviewDecision` is `Approved` and `ApprovedADDistinguishedName` resolves to exactly one AD user object, including a hybrid MailUser. If the approved distinguished name resolves to a MailContact, the row is treated as reviewed and approved, but no `employeeID` update is attempted.
Each approved distinguished name can be used on one row only. If the same `ApprovedADDistinguishedName` appears on multiple approved rows, all rows using that duplicate DN are marked `ReviewDecision Invalid - Duplicate ApprovedADDistinguishedName` and none of them are processed by `-Apply`.
The same duplicate-approved-DN protection is also used by `-DeleteApprovedMailContacts`; duplicated approvals are skipped until the approved DN is unique.
Unresolved, multiple, blank, duplicate, or unsupported approved distinguished names remain invalid approval rows and continue to appear in the unapproved summary.
At the end of each run, the PowerShell console also shows an `Unapproved / still requiring review` summary that excludes accepted approved and ignored rows but still includes invalid approvals that need correction.

## Find-MailContactADUserDuplicates.ps1

This script audits potential GAL duplicates where a MailContact and an AD User / MailUser may represent the same person.

### Switch reference

| Switch | Type | Default | Explanation |
|---|---|---|---|
| `-DoNotInstallMissingModules` | `[switch]` | Off | Stops the script trying to install `ImportExcel`; CSV fallback is used if the module is unavailable. |
| `-IncludeTitleMismatches` | `[switch]` | Off | Includes lower-confidence fallback matches where the name matches but title differs or is missing. |
| `-EnabledUsersOnly` | `[switch]` | Off | Compares MailContacts only with enabled AD user objects. |

### Matching order

1. Email-style values:
   - `mail`
   - `userPrincipalName`
   - `proxyAddresses`
   - `targetAddress`
2. Fallback person match:
   - first name
   - last name
   - title

### Common examples

#### Basic duplicate report

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx"
```

Use this first for a normal report-only duplicate audit.

#### Include title mismatches

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx" `
  -IncludeTitleMismatches
```

Use this when you want the broader review set that includes lower-confidence name matches.

#### Only check enabled AD users

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx" `
  -EnabledUsersOnly
```

Use this when disabled AD users should not be considered as duplicate candidates.

#### Scoped OU search

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx" `
  -SearchBase "OU=Users,DC=example,DC=com"
```

Use this when you want to limit the audit to a known OU or domain partition.

Expected workbook sheets:

- `Duplicate Candidates`
- `Summary`

## Notes

These scripts are intended to support discovery, readiness, and design validation before enabling or changing Microsoft Entra inbound provisioning from Workday to Active Directory. Review all report output before making any directory changes.
