# Entra Provisioning Readiness Scripts

This folder contains two report-first PowerShell scripts created to support Workday to Active Directory provisioning readiness and GAL duplicate investigation.

## Scripts

| Script | Purpose |
|---|---|
| `Invoke-HRToADEmployeeIDMatch.ps1` | Imports an HR / Workday CSV, matches workers to AD users, optionally writes `employeeID`, and can report Workday-vs-AD attribute mismatches before provisioning is enabled. |
| `Find-MailContactADUserDuplicates.ps1` | Audits AD for potential duplicate GAL entries where a MailContact and AD User / MailUser appear to represent the same person. |

## Safety model

Both scripts are designed to be safe for discovery and readiness reporting.

`Invoke-HRToADEmployeeIDMatch.ps1` only writes to AD when `-Apply` is explicitly supplied, and even then it only attempts to update the `employeeID` attribute on matched AD user objects. It does not update Workday-managed attributes such as names, title, department, manager, telephone numbers, address fields, preferred language, or OU placement.

MailContact fallback matches in `Invoke-HRToADEmployeeIDMatch.ps1` are report-only and are not update-eligible. MailUsers are AD user objects in hybrid Exchange environments and can be update-eligible when they are matched or approved.

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
[string]$CsvManagerEmailColumn = "PrimaryWorkEmail_Manager"
```

These can be overridden at runtime if the HR export uses different column names.

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

### Example dry run

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx"
```

### Dry run with system mismatch reporting

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -SystemMismatchReportPath "C:\Temp\Workday-AD-System-Mismatches.xlsx" `
  -CheckSystemMismatch
```

Expected workbook outputs:

- Main report: `HR AD Match`
- Separate system mismatch report: `System Mismatches`

If `-SystemMismatchReportPath` is not supplied, the script creates a separate mismatch report beside the main report using the main file name plus `-System-Mismatches`.

`-CheckSystemMismatch` does not perform a separate user match. It compares the HR row against the AD object selected by the main match logic. Manual `ApprovedADDistinguishedName` decisions are used first, then WorkerID to AD `employeeID`, then email and name-based matching.

The current system mismatch comparison is deliberately limited to:

- `PreferredFirstName` -> AD `givenName`
- `PreferredLastName` -> AD `sn`
- `PrimaryWorkEmail` -> AD `mail`, `proxyAddresses`, or `targetAddress`
- `BusinessTitle` -> AD `title`

The mismatch report includes the selected user's UPN, current AD `employeeID`, current AD value, proposed HR value, and a `WhatWillChange` summary for each differing attribute.

Worker ID values are exported as text in the Excel workbook so leading zeros, such as `001234`, are preserved.

### Apply EmployeeID updates

```powershell
.\Invoke-HRToADEmployeeIDMatch.ps1 `
  -CsvPath "C:\Temp\Workday-HR-Export.csv" `
  -ReportPath "C:\Temp\Workday-AD-Readiness-Report.xlsx" `
  -Apply
```

Only matched and update-eligible AD user objects are updated. Hybrid MailUsers are treated as AD user objects for `employeeID` updates. MailContacts, ambiguous matches, and potential title-mismatch matches are skipped.
Potential name-variant rows are also skipped and are intended for manual review only.
Potential name-variant title-mismatch rows are lower-confidence and are also skipped.
When name-variant matching finds multiple candidates for the same CSV row, the script keeps one report row and lists the candidate count and details in `PotentialMatchCount` and `PotentialMatches`.
Name-variant matching ignores weak surname particles such as `de`, `da`, `van`, and `von`; these words are not enough on their own to create a surname match.
Before producing potential name-variant candidates, the script reserves AD users that are exact email matches, exact name-and-title matches, or approved AD user review decisions anywhere in the input. Reserved AD users are not reused as potential matches for other CSV rows.
Fuzzy first-name matching is intentionally conservative; prefix and token matches are preferred, and pure similarity matches must be strong.
If a row is clearly the wrong person, leave it unapproved or set `ReviewDecision` to a non-approved value such as `Rejected` and add context in `ApprovalNotes`.
If the same AD user is already matched to an earlier CSV row, later update-eligible matches to that same AD user are reported as `Duplicate - AD User Already Matched` and skipped.

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
Unresolved, multiple, blank, duplicate, or unsupported approved distinguished names remain invalid approval rows and continue to appear in the unapproved summary.
At the end of each run, the PowerShell console also shows an `Unapproved / still requiring review` summary that excludes accepted approved and ignored rows but still includes invalid approvals that need correction.

## Find-MailContactADUserDuplicates.ps1

This script audits potential GAL duplicates where a MailContact and an AD User / MailUser may represent the same person.

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

### Example run

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx"
```

### Include title mismatches

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx" `
  -IncludeTitleMismatches
```

### Only check enabled AD users

```powershell
.\Find-MailContactADUserDuplicates.ps1 `
  -ReportPath "C:\Temp\MailContact-ADUser-Duplicate-Report.xlsx" `
  -EnabledUsersOnly
```

Expected workbook sheets:

- `Duplicate Candidates`
- `Summary`

## Notes

These scripts are intended to support discovery, readiness, and design validation before enabling or changing Microsoft Entra inbound provisioning from Workday to Active Directory. Review all report output before making any directory changes.
