<#
.SYNOPSIS
    Exports and safely migrates Active Directory user extension attributes.

.DESCRIPTION
    Exports extensionAttribute1 through extensionAttribute15 for every Active Directory user in
    scope to a timestamped CSV file. The script can then import one of its snapshots and copy the
    populated source attributes to the mapped target attributes without clearing the source.

    Mapping supplied for this migration:
        extensionAttribute1 -> extensionAttribute10
        extensionAttribute2 -> extensionAttribute15
        extensionAttribute3 -> extensionAttribute9
        extensionAttribute4 -> extensionAttribute14
        extensionAttribute5 -> extensionAttribute13
        extensionAttribute6 -> extensionAttribute11
        extensionAttribute7 -> extensionAttribute12

    Source cleanup is a separate operation. For each user, -ClearSource clears populated source
    attributes only when every populated extensionAttribute1 through extensionAttribute7 has an
    exact, ordinal match in its mapped target attribute. If any populated pair does not match,
    none of that user's source attributes are cleared.

    Copy and cleanup operations support -WhatIf and -Confirm. Timestamped pre-change snapshots,
    detailed results, summaries, and post-change snapshots are written for auditing.

.PARAMETER Export
    Exports all 15 extension attributes for every user in scope. Use with -Copy to create a fresh
    snapshot and immediately use that CSV as the copy input.

.PARAMETER Copy
    Imports source values from a script-generated CSV snapshot and copies populated values into
    the mapped target attributes. Source attributes are not cleared.

.PARAMETER ClearSource
    Runs the isolated cleanup operation. This switch cannot be combined with -Export or -Copy.
    Before clearing anything, the script creates a fresh snapshot of all users in scope. A user is
    changed only if all of that user's populated source/target pairs match exactly.

.PARAMETER InputCsvPath
    Path to a CSV snapshot created by this script. Required with -Copy unless -Export is specified
    in the same command.

.PARAMETER SearchBase
    Optional distinguished name of the OU or container to process. If omitted, the current domain
    naming context is used. During CSV import, users currently outside this SearchBase are skipped.

.PARAMETER TestUser
    Limits export, copy, or cleanup to one AD user for controlled testing. Supply an identity
    accepted by Get-ADUser -Identity, such as sAMAccountName, distinguished name, ObjectGUID, or SID.
    With -Copy, the selected CSV must contain exactly one row and its ObjectGuid must match this user.

.PARAMETER OutputFolder
    Folder for timestamped snapshots, result files, and summary files. The folder is created when
    necessary. Generated files may contain personal data and should be stored securely.

.PARAMETER Server
    Optional domain controller or AD LDS instance to query and update.

.PARAMETER Credential
    Optional credential for Active Directory operations.

.PARAMETER OverwriteTarget
    Allows -Copy to overwrite a populated mapped target when it differs from the CSV source value.
    Without this explicit switch, a conflicting target causes the entire user to be skipped.

.PARAMETER DisableProgress
    Suppresses progress bars.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -Export -OutputFolder "C:\Secure\ExtensionAttributeMigration"

    Exports all 15 extension attributes for every user in the current domain.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -Export -Copy -WhatIf -Verbose

    Creates a fresh snapshot and previews the copy operation without changing Active Directory.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -Export -TestUser "test.user" -OutputFolder "C:\Secure\ExtensionAttributeMigration\Test"

    Exports all 15 extension attributes for only the selected test user.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -Copy -TestUser "test.user" -InputCsvPath "C:\Secure\ExtensionAttributeMigration\Test\ADUserExtensionAttributes-TestUser-Export-20260820-100000000.csv" -WhatIf

    Validates that the test CSV contains only the selected user, then previews copying that user.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -Copy -InputCsvPath "C:\Secure\ADUserExtensionAttributes-Export-20260819-190000000.csv"

    Imports a previous script-generated snapshot and interactively confirms user changes.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -ClearSource -WhatIf -Verbose

    Creates a fresh pre-cleanup snapshot and previews which users pass the fail-safe.

.EXAMPLE
    .\Move-ADUserExtensionAttributes.ps1 -ClearSource -Confirm:$false

    Clears verified source values without per-user confirmation. Use only after reviewing a
    -WhatIf run and its result report.

.NOTES
    Required module:
        ActiveDirectory

    Required access:
        Read access to user objects and extensionAttribute1 through extensionAttribute15.
        Write access to extensionAttribute9 through extensionAttribute15 for -Copy.
        Clear/write access to extensionAttribute1 through extensionAttribute7 for -ClearSource.

    PowerShell:
        Windows PowerShell 5.1 or PowerShell 7 on Windows with RSAT Active Directory tools.

    Scope:
        This script targets on-premises Active Directory attributes. It is not intended for
        cloud-only Exchange Online CustomAttribute values or Entra ID-only extension properties.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [switch]$Export,

    [switch]$Copy,

    [switch]$ClearSource,

    [string]$InputCsvPath,

    [string]$SearchBase,

    [Alias("TestUserIdentity")]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$TestUser,

    [string]$OutputFolder = ".\Output",

    [string]$Server,

    [System.Management.Automation.PSCredential]$Credential,

    [switch]$OverwriteTarget,

    [switch]$DisableProgress
)

$ErrorActionPreference = "Stop"
$script:SnapshotSchemaVersion = "1.0"
$script:AttributeMapping = [ordered]@{
    extensionAttribute1 = "extensionAttribute10"
    extensionAttribute2 = "extensionAttribute15"
    extensionAttribute3 = "extensionAttribute9"
    extensionAttribute4 = "extensionAttribute14"
    extensionAttribute5 = "extensionAttribute13"
    extensionAttribute6 = "extensionAttribute11"
    extensionAttribute7 = "extensionAttribute12"
}
$script:AllExtensionAttributes = @(1..15 | ForEach-Object { "extensionAttribute$_" })
$script:SnapshotProperties = @("mail") + $script:AllExtensionAttributes

function Import-ActiveDirectoryModule {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        return
    }
    catch {
        $initialError = $_
    }

    if ($PSVersionTable.PSVersion.Major -ge 6 -and $IsWindows) {
        try {
            Import-Module ActiveDirectory -UseWindowsPowerShell -ErrorAction Stop
            return
        }
        catch {
            throw "The ActiveDirectory module could not be imported. Install RSAT Active Directory tools and try again. Original error: $($initialError.Exception.Message)"
        }
    }

    throw "The ActiveDirectory module could not be imported. Install RSAT Active Directory tools and try again. Original error: $($initialError.Exception.Message)"
}

function ConvertTo-ExactString {
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [string]$Value
}

function Test-ExactStringEqual {
    param(
        [AllowNull()]
        $Left,

        [AllowNull()]
        $Right
    )

    $leftString = ConvertTo-ExactString -Value $Left
    $rightString = ConvertTo-ExactString -Value $Right
    return [string]::Equals($leftString, $rightString, [System.StringComparison]::Ordinal)
}

function Get-ConnectionParameter {
    param(
        [AllowEmptyString()]
        [string]$ServerValue,

        [AllowNull()]
        [System.Management.Automation.PSCredential]$CredentialValue
    )

    $parameters = @{}

    if (-not [string]::IsNullOrWhiteSpace($ServerValue)) {
        $parameters["Server"] = $ServerValue
    }

    if ($null -ne $CredentialValue) {
        $parameters["Credential"] = $CredentialValue
    }

    return $parameters
}

function Get-ScopedUser {
    param(
        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters,

        [AllowEmptyString()]
        [string]$SearchBaseValue,

        [AllowEmptyString()]
        [string]$TestUserIdentityValue
    )

    $parameters = @{
        Properties  = $script:SnapshotProperties
        ErrorAction = "Stop"
    }

    foreach ($key in $ConnectionParameters.Keys) {
        $parameters[$key] = $ConnectionParameters[$key]
    }

    if (-not [string]::IsNullOrWhiteSpace($TestUserIdentityValue)) {
        $parameters["Identity"] = $TestUserIdentityValue
        Write-Verbose "Querying the selected test user '$TestUserIdentityValue'."
        $selectedUser = Get-ADUser @parameters

        if (-not (Test-UserInSearchBase -User $selectedUser -SearchBaseValue $SearchBaseValue)) {
            throw "The selected test user '$TestUserIdentityValue' is outside the specified SearchBase '$SearchBaseValue'."
        }

        return @($selectedUser)
    }

    $parameters["Filter"] = "*"
    if (-not [string]::IsNullOrWhiteSpace($SearchBaseValue)) {
        $parameters["SearchBase"] = $SearchBaseValue
    }

    Write-Verbose "Querying Active Directory users."
    return @(Get-ADUser @parameters | Sort-Object -Property SamAccountName)
}

function Get-UserByObjectGuid {
    param(
        [Parameter(Mandatory)]
        [guid]$ObjectGuid,

        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters
    )

    $parameters = @{
        Identity    = $ObjectGuid
        Properties  = $script:SnapshotProperties
        ErrorAction = "Stop"
    }

    foreach ($key in $ConnectionParameters.Keys) {
        $parameters[$key] = $ConnectionParameters[$key]
    }

    return Get-ADUser @parameters
}

function Test-UserInSearchBase {
    param(
        [Parameter(Mandatory)]
        $User,

        [AllowEmptyString()]
        [string]$SearchBaseValue
    )

    if ([string]::IsNullOrWhiteSpace($SearchBaseValue)) {
        return $true
    }

    $userDistinguishedName = ConvertTo-ExactString -Value $User.DistinguishedName
    $normalisedSearchBase = $SearchBaseValue.Trim()
    return $userDistinguishedName.EndsWith(",$normalisedSearchBase", [System.StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-SnapshotRow {
    param(
        [Parameter(Mandatory)]
        $User,

        [Parameter(Mandatory)]
        [string]$SnapshotPurpose,

        [Parameter(Mandatory)]
        [string]$GeneratedUtc
    )

    $row = [ordered]@{
        SnapshotSchemaVersion = $script:SnapshotSchemaVersion
        SnapshotPurpose       = $SnapshotPurpose
        SnapshotGeneratedUtc  = $GeneratedUtc
        ObjectGuid            = $User.ObjectGuid.ToString("D")
        SID                   = if ($null -ne $User.SID) { $User.SID.ToString() } else { "" }
        SamAccountName        = ConvertTo-ExactString -Value $User.SamAccountName
        UserPrincipalName     = ConvertTo-ExactString -Value $User.UserPrincipalName
        Name                  = ConvertTo-ExactString -Value $User.Name
        Mail                  = ConvertTo-ExactString -Value $User.mail
        Enabled               = [bool]$User.Enabled
        DistinguishedName     = ConvertTo-ExactString -Value $User.DistinguishedName
    }

    foreach ($attribute in $script:AllExtensionAttributes) {
        $row[$attribute] = ConvertTo-ExactString -Value $User.$attribute
    }

    return [PSCustomObject]$row
}

function Export-UserSnapshot {
    param(
        [Parameter(Mandatory)]
        [array]$Users,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$SnapshotPurpose
    )

    $generatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    $rows = @(
        foreach ($user in $Users) {
            ConvertTo-SnapshotRow -User $user -SnapshotPurpose $SnapshotPurpose -GeneratedUtc $generatedUtc
        }
    )

    if (@($rows).Count -eq 0) {
        throw "No Active Directory users were found in scope. No snapshot was written."
    }

    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -WhatIf:$false -Confirm:$false
    Write-Verbose "Exported $(@($rows).Count) users to '$Path'."

    return [PSCustomObject]@{
        Path      = $Path
        UserCount = @($rows).Count
    }
}

function Export-ScopedSnapshot {
    param(
        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters,

        [AllowEmptyString()]
        [string]$SearchBaseValue,

        [AllowEmptyString()]
        [string]$TestUserIdentityValue,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$SnapshotPurpose
    )

    $users = @(Get-ScopedUser -ConnectionParameters $ConnectionParameters -SearchBaseValue $SearchBaseValue -TestUserIdentityValue $TestUserIdentityValue)
    return Export-UserSnapshot -Users $users -Path $Path -SnapshotPurpose $SnapshotPurpose
}

function Import-ValidatedSnapshot {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $rows = @(Import-Csv -LiteralPath $resolvedPath)

    if (@($rows).Count -eq 0) {
        throw "The input CSV '$resolvedPath' contains no user rows."
    }

    $requiredColumns = @("SnapshotSchemaVersion", "ObjectGuid") + @($script:AttributeMapping.Keys)
    $availableColumns = @($rows[0].PSObject.Properties.Name)
    $missingColumns = @($requiredColumns | Where-Object { $_ -notin $availableColumns })

    if (@($missingColumns).Count -gt 0) {
        throw "The input CSV is missing required columns: $([string]::Join(', ', $missingColumns)). Use a snapshot created by this script."
    }

    $wrongVersions = @(
        $rows |
            Where-Object { $_.SnapshotSchemaVersion -ne $script:SnapshotSchemaVersion } |
            Select-Object -ExpandProperty SnapshotSchemaVersion -Unique
    )
    if (@($wrongVersions).Count -gt 0) {
        throw "The input CSV uses an unsupported SnapshotSchemaVersion: $([string]::Join(', ', $wrongVersions)). Expected $($script:SnapshotSchemaVersion)."
    }

    $invalidObjectGuids = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $rows) {
        $parsedGuid = [guid]::Empty
        if (-not [guid]::TryParse([string]$row.ObjectGuid, [ref]$parsedGuid)) {
            $invalidObjectGuids.Add([string]$row.ObjectGuid)
        }
    }

    if ($invalidObjectGuids.Count -gt 0) {
        throw "The input CSV contains invalid ObjectGuid values: $([string]::Join(', ', @($invalidObjectGuids)))."
    }

    $duplicateObjectGuids = @(
        $rows |
            Group-Object -Property ObjectGuid |
            Where-Object { $_.Count -gt 1 } |
            Select-Object -ExpandProperty Name
    )
    if (@($duplicateObjectGuids).Count -gt 0) {
        throw "The input CSV contains duplicate ObjectGuid values: $([string]::Join(', ', $duplicateObjectGuids)). No changes were made."
    }

    return [PSCustomObject]@{
        Path = $resolvedPath
        Rows = $rows
    }
}

function Test-TestUserSnapshot {
    param(
        [Parameter(Mandatory)]
        [array]$SnapshotRows,

        [Parameter(Mandatory)]
        [string]$TestUserIdentityValue,

        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters,

        [AllowEmptyString()]
        [string]$SearchBaseValue
    )

    if (@($SnapshotRows).Count -ne 1) {
        throw "Test-user copy requires an input CSV containing exactly one user row. The selected CSV contains $(@($SnapshotRows).Count) rows. No changes were made."
    }

    $selectedUsers = @(Get-ScopedUser -ConnectionParameters $ConnectionParameters -SearchBaseValue $SearchBaseValue -TestUserIdentityValue $TestUserIdentityValue)
    if (@($selectedUsers).Count -ne 1) {
        throw "The test-user identity '$TestUserIdentityValue' did not resolve to exactly one AD user. No changes were made."
    }

    $selectedUser = $selectedUsers[0]
    $csvObjectGuid = [guid]$SnapshotRows[0].ObjectGuid
    if ($csvObjectGuid -ne $selectedUser.ObjectGuid) {
        throw "The selected test CSV contains ObjectGuid '$csvObjectGuid', but -TestUser '$TestUserIdentityValue' resolves to '$($selectedUser.ObjectGuid)'. No changes were made."
    }

    Write-Verbose "Validated that the test CSV contains only '$($selectedUser.SamAccountName)' ($($selectedUser.ObjectGuid))."
    return $true
}

function ConvertTo-OperationResult {
    param(
        [Parameter(Mandatory)]
        [string]$Operation,

        [AllowEmptyString()]
        [string]$ObjectGuid,

        [AllowEmptyString()]
        [string]$SamAccountName,

        [AllowEmptyString()]
        [string]$UserPrincipalName,

        [AllowEmptyString()]
        [string]$DistinguishedName,

        [Parameter(Mandatory)]
        [string]$Status,

        [string[]]$ChangedAttributes = @(),

        [string[]]$CheckedMappings = @(),

        [AllowEmptyString()]
        [string]$Details = ""
    )

    return [PSCustomObject]@{
        TimestampUtc       = (Get-Date).ToUniversalTime().ToString("o")
        Operation          = $Operation
        ObjectGuid         = $ObjectGuid
        SamAccountName     = $SamAccountName
        UserPrincipalName  = $UserPrincipalName
        DistinguishedName  = $DistinguishedName
        Status             = $Status
        ChangedAttributes  = [string]::Join(";", @($ChangedAttributes))
        CheckedMappings    = [string]::Join(";", @($CheckedMappings))
        Details            = $Details
    }
}

function Invoke-AttributeCopy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [array]$SnapshotRows,

        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters,

        [AllowEmptyString()]
        [string]$SearchBaseValue,

        [switch]$AllowTargetOverwrite,

        [switch]$SuppressProgress,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCmdlet]$CallingCmdlet,

        [switch]$IsWhatIf
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $totalRows = @($SnapshotRows).Count

    for ($index = 0; $index -lt $totalRows; $index++) {
        $row = $SnapshotRows[$index]
        $currentNumber = $index + 1

        if (-not $SuppressProgress) {
            $percentComplete = [math]::Min(100, [int](($currentNumber / $totalRows) * 100))
            Write-Progress -Activity "Copying AD extension attributes" -Status "Processing $currentNumber of $totalRows" -PercentComplete $percentComplete
        }

        $objectGuidText = ConvertTo-ExactString -Value $row.ObjectGuid
        $rowSamAccountName = ConvertTo-ExactString -Value $row.SamAccountName
        $rowUserPrincipalName = ConvertTo-ExactString -Value $row.UserPrincipalName
        $rowDistinguishedName = ConvertTo-ExactString -Value $row.DistinguishedName

        try {
            $objectGuid = [guid]$objectGuidText
            $currentUser = Get-UserByObjectGuid -ObjectGuid $objectGuid -ConnectionParameters $ConnectionParameters

            if (-not (Test-UserInSearchBase -User $currentUser -SearchBaseValue $SearchBaseValue)) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $rowSamAccountName -UserPrincipalName $rowUserPrincipalName -DistinguishedName $rowDistinguishedName -Status "SkippedOutsideSearchBase" -Details "The user is no longer within the specified SearchBase."))
                continue
            }

            $changedSources = [System.Collections.Generic.List[string]]::new()
            foreach ($sourceAttribute in $script:AttributeMapping.Keys) {
                if (-not (Test-ExactStringEqual -Left $row.$sourceAttribute -Right $currentUser.$sourceAttribute)) {
                    $changedSources.Add($sourceAttribute)
                }
            }

            if ($changedSources.Count -gt 0) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "SkippedSourceChanged" -CheckedMappings @($changedSources) -Details "One or more source values changed after the input snapshot. No target attributes were changed for this user."))
                continue
            }

            $replaceValues = @{}
            $targetConflicts = [System.Collections.Generic.List[string]]::new()
            $populatedMappings = [System.Collections.Generic.List[string]]::new()

            foreach ($sourceAttribute in $script:AttributeMapping.Keys) {
                $targetAttribute = $script:AttributeMapping[$sourceAttribute]
                $desiredValue = ConvertTo-ExactString -Value $row.$sourceAttribute

                if ([string]::IsNullOrEmpty($desiredValue)) {
                    continue
                }

                $mappingLabel = "$sourceAttribute->$targetAttribute"
                $populatedMappings.Add($mappingLabel)
                $currentTargetValue = ConvertTo-ExactString -Value $currentUser.$targetAttribute

                if (Test-ExactStringEqual -Left $desiredValue -Right $currentTargetValue) {
                    continue
                }

                if (-not [string]::IsNullOrEmpty($currentTargetValue) -and -not $AllowTargetOverwrite) {
                    $targetConflicts.Add($mappingLabel)
                    continue
                }

                $replaceValues[$targetAttribute] = $desiredValue
            }

            if ($targetConflicts.Count -gt 0) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "SkippedTargetConflict" -CheckedMappings @($targetConflicts) -Details "A mapped target already contains a different value. No target attributes were changed for this user. Use -OverwriteTarget only after reviewing the conflict."))
                continue
            }

            if ($populatedMappings.Count -eq 0) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "NoSourceValues" -Details "All source attributes in the input snapshot are empty."))
                continue
            }

            if ($replaceValues.Count -eq 0) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "AlreadyCopied" -CheckedMappings @($populatedMappings) -Details "Every populated source value already matches its mapped target."))
                continue
            }

            $attributesToChange = @($replaceValues.Keys | Sort-Object)
            $action = "Set and verify mapped target attributes: $([string]::Join(', ', $attributesToChange))"

            if (-not $CallingCmdlet.ShouldProcess($currentUser.DistinguishedName, $action)) {
                $status = if ($IsWhatIf) { "WhatIf" } else { "SkippedByShouldProcess" }
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status $status -ChangedAttributes $attributesToChange -CheckedMappings @($populatedMappings) -Details "Active Directory was not changed."))
                continue
            }

            $setParameters = @{
                Identity    = $objectGuid
                Replace     = $replaceValues
                ErrorAction = "Stop"
                Confirm     = $false
            }
            foreach ($key in $ConnectionParameters.Keys) {
                $setParameters[$key] = $ConnectionParameters[$key]
            }

            Set-ADUser @setParameters
            $verifiedUser = Get-UserByObjectGuid -ObjectGuid $objectGuid -ConnectionParameters $ConnectionParameters

            $verificationFailures = [System.Collections.Generic.List[string]]::new()
            foreach ($sourceAttribute in $script:AttributeMapping.Keys) {
                $desiredValue = ConvertTo-ExactString -Value $row.$sourceAttribute
                if ([string]::IsNullOrEmpty($desiredValue)) {
                    continue
                }

                $targetAttribute = $script:AttributeMapping[$sourceAttribute]
                if (-not (Test-ExactStringEqual -Left $desiredValue -Right $verifiedUser.$targetAttribute)) {
                    $verificationFailures.Add("$sourceAttribute->$targetAttribute")
                }
            }

            if ($verificationFailures.Count -gt 0) {
                $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $verifiedUser.SamAccountName -UserPrincipalName $verifiedUser.UserPrincipalName -DistinguishedName $verifiedUser.DistinguishedName -Status "PostChangeVerificationFailed" -ChangedAttributes $attributesToChange -CheckedMappings @($verificationFailures) -Details "The AD write completed, but one or more target values did not verify. Review this user before cleanup."))
                Write-Warning "Post-change verification failed for '$($verifiedUser.DistinguishedName)'."
                continue
            }

            $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $verifiedUser.SamAccountName -UserPrincipalName $verifiedUser.UserPrincipalName -DistinguishedName $verifiedUser.DistinguishedName -Status "CopiedAndVerified" -ChangedAttributes $attributesToChange -CheckedMappings @($populatedMappings) -Details "Mapped targets were updated and verified. Source attributes were left unchanged."))
        }
        catch {
            $results.Add((ConvertTo-OperationResult -Operation "Copy" -ObjectGuid $objectGuidText -SamAccountName $rowSamAccountName -UserPrincipalName $rowUserPrincipalName -DistinguishedName $rowDistinguishedName -Status "Failed" -Details $_.Exception.Message))
            Write-Warning "Copy failed for ObjectGuid '$objectGuidText': $($_.Exception.Message)"
        }
    }

    if (-not $SuppressProgress) {
        Write-Progress -Activity "Copying AD extension attributes" -Completed
    }

    return @($results)
}

function Invoke-SourceCleanup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [array]$Users,

        [Parameter(Mandatory)]
        [hashtable]$ConnectionParameters,

        [AllowEmptyString()]
        [string]$SearchBaseValue,

        [switch]$SuppressProgress,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCmdlet]$CallingCmdlet,

        [switch]$IsWhatIf
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $totalUsers = @($Users).Count

    for ($index = 0; $index -lt $totalUsers; $index++) {
        $snapshotUser = $Users[$index]
        $currentNumber = $index + 1

        if (-not $SuppressProgress) {
            $percentComplete = [math]::Min(100, [int](($currentNumber / $totalUsers) * 100))
            Write-Progress -Activity "Verifying and clearing source extension attributes" -Status "Processing $currentNumber of $totalUsers" -PercentComplete $percentComplete
        }

        $objectGuidText = $snapshotUser.ObjectGuid.ToString("D")

        try {
            $currentUser = Get-UserByObjectGuid -ObjectGuid $snapshotUser.ObjectGuid -ConnectionParameters $ConnectionParameters

            if (-not (Test-UserInSearchBase -User $currentUser -SearchBaseValue $SearchBaseValue)) {
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $snapshotUser.SamAccountName -UserPrincipalName $snapshotUser.UserPrincipalName -DistinguishedName $snapshotUser.DistinguishedName -Status "SkippedOutsideSearchBase" -Details "The user moved outside the specified SearchBase after the pre-cleanup snapshot. No source attributes were cleared."))
                continue
            }

            $changedAfterSnapshot = [System.Collections.Generic.List[string]]::new()
            $mappedAttributes = @($script:AttributeMapping.Keys) + @($script:AttributeMapping.Values)
            foreach ($attribute in @($mappedAttributes | Sort-Object -Unique)) {
                if (-not (Test-ExactStringEqual -Left $snapshotUser.$attribute -Right $currentUser.$attribute)) {
                    $changedAfterSnapshot.Add($attribute)
                }
            }

            if ($changedAfterSnapshot.Count -gt 0) {
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "SkippedChangedAfterPreClearSnapshot" -CheckedMappings @($changedAfterSnapshot) -Details "Mapped source or target values changed after the mandatory pre-cleanup snapshot. No source attributes were cleared for this user."))
                continue
            }

            $populatedSources = [System.Collections.Generic.List[string]]::new()
            $checkedMappings = [System.Collections.Generic.List[string]]::new()
            $verificationFailures = [System.Collections.Generic.List[string]]::new()
            $expectedTargetValues = @{}

            foreach ($sourceAttribute in $script:AttributeMapping.Keys) {
                $sourceValue = ConvertTo-ExactString -Value $currentUser.$sourceAttribute
                if ([string]::IsNullOrEmpty($sourceValue)) {
                    continue
                }

                $targetAttribute = $script:AttributeMapping[$sourceAttribute]
                $mappingLabel = "$sourceAttribute->$targetAttribute"
                $populatedSources.Add($sourceAttribute)
                $checkedMappings.Add($mappingLabel)
                $expectedTargetValues[$targetAttribute] = $sourceValue

                if (-not (Test-ExactStringEqual -Left $sourceValue -Right $currentUser.$targetAttribute)) {
                    $verificationFailures.Add($mappingLabel)
                }
            }

            if ($populatedSources.Count -eq 0) {
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "NoSourceValues" -Details "There are no populated source attributes to clear."))
                continue
            }

            if ($verificationFailures.Count -gt 0) {
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status "SkippedVerificationFailed" -CheckedMappings @($verificationFailures) -Details "At least one populated source value does not exactly match its mapped target. None of this user's source attributes were cleared."))
                continue
            }

            $attributesToClear = @($populatedSources | Sort-Object)
            $action = "Clear verified source attributes: $([string]::Join(', ', $attributesToClear))"

            if (-not $CallingCmdlet.ShouldProcess($currentUser.DistinguishedName, $action)) {
                $status = if ($IsWhatIf) { "WhatIf" } else { "SkippedByShouldProcess" }
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $currentUser.SamAccountName -UserPrincipalName $currentUser.UserPrincipalName -DistinguishedName $currentUser.DistinguishedName -Status $status -ChangedAttributes $attributesToClear -CheckedMappings @($checkedMappings) -Details "Active Directory was not changed. The fail-safe verification passed."))
                continue
            }

            $setParameters = @{
                Identity    = $currentUser.ObjectGuid
                Clear       = $attributesToClear
                ErrorAction = "Stop"
                Confirm     = $false
            }
            foreach ($key in $ConnectionParameters.Keys) {
                $setParameters[$key] = $ConnectionParameters[$key]
            }

            Set-ADUser @setParameters
            $verifiedUser = Get-UserByObjectGuid -ObjectGuid $currentUser.ObjectGuid -ConnectionParameters $ConnectionParameters

            $remainingSources = @(
                $attributesToClear |
                    Where-Object { -not [string]::IsNullOrEmpty((ConvertTo-ExactString -Value $verifiedUser.$_)) }
            )
            $changedTargets = [System.Collections.Generic.List[string]]::new()
            foreach ($targetAttribute in $expectedTargetValues.Keys) {
                if (-not (Test-ExactStringEqual -Left $expectedTargetValues[$targetAttribute] -Right $verifiedUser.$targetAttribute)) {
                    $changedTargets.Add($targetAttribute)
                }
            }

            if (@($remainingSources).Count -gt 0 -or $changedTargets.Count -gt 0) {
                $failedChecks = @($remainingSources) + @($changedTargets)
                $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $verifiedUser.SamAccountName -UserPrincipalName $verifiedUser.UserPrincipalName -DistinguishedName $verifiedUser.DistinguishedName -Status "PostChangeVerificationFailed" -ChangedAttributes $attributesToClear -CheckedMappings $failedChecks -Details "The cleanup write completed, but the post-change verification did not pass. Review the pre-cleanup snapshot and this user immediately."))
                Write-Warning "Post-cleanup verification failed for '$($verifiedUser.DistinguishedName)'."
                continue
            }

            $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $verifiedUser.SamAccountName -UserPrincipalName $verifiedUser.UserPrincipalName -DistinguishedName $verifiedUser.DistinguishedName -Status "ClearedAndVerified" -ChangedAttributes $attributesToClear -CheckedMappings @($checkedMappings) -Details "All populated source/target pairs matched before cleanup. Source values were cleared and the result was verified."))
        }
        catch {
            $results.Add((ConvertTo-OperationResult -Operation "ClearSource" -ObjectGuid $objectGuidText -SamAccountName $snapshotUser.SamAccountName -UserPrincipalName $snapshotUser.UserPrincipalName -DistinguishedName $snapshotUser.DistinguishedName -Status "Failed" -Details $_.Exception.Message))
            Write-Warning "Cleanup failed for ObjectGuid '$objectGuidText': $($_.Exception.Message)"
        }
    }

    if (-not $SuppressProgress) {
        Write-Progress -Activity "Verifying and clearing source extension attributes" -Completed
    }

    return @($results)
}

function Write-OperationReport {
    param(
        [Parameter(Mandatory)]
        [string]$Operation,

        [Parameter(Mandatory)]
        [array]$Results,

        [Parameter(Mandatory)]
        [string]$OutputFolderPath,

        [Parameter(Mandatory)]
        [string]$RunTimestamp,

        [AllowEmptyString()]
        [string]$InputSnapshotPath,

        [AllowEmptyString()]
        [string]$PreChangeSnapshotPath,

        [AllowEmptyString()]
        [string]$PostChangeSnapshotPath,

        [AllowEmptyString()]
        [string]$TestUserIdentity
    )

    $resultsPath = Join-Path -Path $OutputFolderPath -ChildPath "ADUserExtensionAttributes-$Operation-Results-$RunTimestamp.csv"
    $summaryPath = Join-Path -Path $OutputFolderPath -ChildPath "ADUserExtensionAttributes-$Operation-Summary-$RunTimestamp.csv"

    @($Results) | Export-Csv -LiteralPath $resultsPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false -Confirm:$false

    $changedStatuses = @("CopiedAndVerified", "ClearedAndVerified")
    $failureStatuses = @("Failed", "PostChangeVerificationFailed")
    $summary = [PSCustomObject]@{
        GeneratedUtc             = (Get-Date).ToUniversalTime().ToString("o")
        Operation                = $Operation
        TestUserIdentity         = $TestUserIdentity
        TotalUsers               = @($Results).Count
        ChangedAndVerified       = @($Results | Where-Object { $_.Status -in $changedStatuses }).Count
        AlreadyCopied            = @($Results | Where-Object { $_.Status -eq "AlreadyCopied" }).Count
        NoSourceValues           = @($Results | Where-Object { $_.Status -eq "NoSourceValues" }).Count
        WhatIf                   = @($Results | Where-Object { $_.Status -eq "WhatIf" }).Count
        Skipped                  = @($Results | Where-Object { $_.Status -like "Skipped*" }).Count
        Failed                   = @($Results | Where-Object { $_.Status -in $failureStatuses }).Count
        InputSnapshotPath        = $InputSnapshotPath
        PreChangeSnapshotPath    = $PreChangeSnapshotPath
        PostChangeSnapshotPath   = $PostChangeSnapshotPath
        ResultsPath              = $resultsPath
        SummaryPath              = $summaryPath
    }

    $summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false -Confirm:$false
    return $summary
}

if (-not ($Export -or $Copy -or $ClearSource)) {
    throw "Specify at least one operation: -Export, -Copy, or -ClearSource. Start with -Export or use -WhatIf for a change operation."
}

if ($ClearSource -and ($Export -or $Copy)) {
    throw "-ClearSource is an isolated cleanup operation and cannot be combined with -Export or -Copy. Run it as a separate command."
}

if ($Copy -and -not $Export -and [string]::IsNullOrWhiteSpace($InputCsvPath)) {
    throw "-Copy requires -InputCsvPath unless -Export is specified in the same command."
}

if ($Export -and $Copy -and -not [string]::IsNullOrWhiteSpace($InputCsvPath)) {
    throw "Do not specify -InputCsvPath with -Export -Copy. The new export is automatically used as the copy input."
}

if (-not $Copy -and -not [string]::IsNullOrWhiteSpace($InputCsvPath)) {
    throw "-InputCsvPath is valid only with -Copy."
}

if ($OverwriteTarget -and -not $Copy) {
    throw "-OverwriteTarget is valid only with -Copy."
}

Import-ActiveDirectoryModule
$connectionParameters = Get-ConnectionParameter -ServerValue $Server -CredentialValue $Credential
$resolvedOutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
$null = New-Item -Path $resolvedOutputFolder -ItemType Directory -Force -WhatIf:$false -Confirm:$false
$runTimestamp = Get-Date -Format "yyyyMMdd-HHmmssfff"
$scopePrefix = if ([string]::IsNullOrWhiteSpace($TestUser)) { "" } else { "TestUser-" }

if ($Export) {
    $baseSnapshotPurpose = if ($Copy) { "ExportAndCopyInput" } else { "Export" }
    $snapshotPurpose = "$scopePrefix$baseSnapshotPurpose"
    $exportPath = Join-Path -Path $resolvedOutputFolder -ChildPath "ADUserExtensionAttributes-$snapshotPurpose-$runTimestamp.csv"
    $exportResult = Export-ScopedSnapshot -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -TestUserIdentityValue $TestUser -Path $exportPath -SnapshotPurpose $snapshotPurpose

    if (-not $Copy) {
        [PSCustomObject]@{
            Operation     = "Export"
            TestUser      = $TestUser
            UsersExported = $exportResult.UserCount
            SnapshotPath  = $exportResult.Path
        }
        return
    }

    $InputCsvPath = $exportResult.Path
}

if ($Copy) {
    $validatedSnapshot = Import-ValidatedSnapshot -Path $InputCsvPath
    if (-not [string]::IsNullOrWhiteSpace($TestUser)) {
        $null = Test-TestUserSnapshot -SnapshotRows $validatedSnapshot.Rows -TestUserIdentityValue $TestUser -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase
    }

    $preChangeSnapshotPath = ""

    if ($Export) {
        $preChangeSnapshotPath = $validatedSnapshot.Path
    }
    else {
        $preCopyPurpose = "${scopePrefix}PreCopy"
        $preChangeSnapshotPath = Join-Path -Path $resolvedOutputFolder -ChildPath "ADUserExtensionAttributes-$preCopyPurpose-$runTimestamp.csv"
        $null = Export-ScopedSnapshot -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -TestUserIdentityValue $TestUser -Path $preChangeSnapshotPath -SnapshotPurpose $preCopyPurpose
    }

    $copyResults = @(
        Invoke-AttributeCopy -SnapshotRows $validatedSnapshot.Rows -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -AllowTargetOverwrite:$OverwriteTarget -SuppressProgress:$DisableProgress -CallingCmdlet $PSCmdlet -IsWhatIf:$WhatIfPreference
    )

    $copyChangedCount = @($copyResults | Where-Object { $_.Status -eq "CopiedAndVerified" }).Count
    $postChangeSnapshotPath = ""
    if ($copyChangedCount -gt 0) {
        $postCopyPurpose = "${scopePrefix}PostCopy"
        $postChangeSnapshotPath = Join-Path -Path $resolvedOutputFolder -ChildPath "ADUserExtensionAttributes-$postCopyPurpose-$runTimestamp.csv"
        $null = Export-ScopedSnapshot -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -TestUserIdentityValue $TestUser -Path $postChangeSnapshotPath -SnapshotPurpose $postCopyPurpose
    }

    Write-OperationReport -Operation "Copy" -Results $copyResults -OutputFolderPath $resolvedOutputFolder -RunTimestamp $runTimestamp -InputSnapshotPath $validatedSnapshot.Path -PreChangeSnapshotPath $preChangeSnapshotPath -PostChangeSnapshotPath $postChangeSnapshotPath -TestUserIdentity $TestUser
    return
}

if ($ClearSource) {
    $usersBeforeCleanup = @(Get-ScopedUser -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -TestUserIdentityValue $TestUser)
    $preClearPurpose = "${scopePrefix}PreClear"
    $preCleanupSnapshotPath = Join-Path -Path $resolvedOutputFolder -ChildPath "ADUserExtensionAttributes-$preClearPurpose-$runTimestamp.csv"
    $null = Export-UserSnapshot -Users $usersBeforeCleanup -Path $preCleanupSnapshotPath -SnapshotPurpose $preClearPurpose

    $cleanupResults = @(
        Invoke-SourceCleanup -Users $usersBeforeCleanup -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -SuppressProgress:$DisableProgress -CallingCmdlet $PSCmdlet -IsWhatIf:$WhatIfPreference
    )

    $cleanupChangedCount = @($cleanupResults | Where-Object { $_.Status -eq "ClearedAndVerified" }).Count
    $postCleanupSnapshotPath = ""
    if ($cleanupChangedCount -gt 0) {
        $postClearPurpose = "${scopePrefix}PostClear"
        $postCleanupSnapshotPath = Join-Path -Path $resolvedOutputFolder -ChildPath "ADUserExtensionAttributes-$postClearPurpose-$runTimestamp.csv"
        $null = Export-ScopedSnapshot -ConnectionParameters $connectionParameters -SearchBaseValue $SearchBase -TestUserIdentityValue $TestUser -Path $postCleanupSnapshotPath -SnapshotPurpose $postClearPurpose
    }

    Write-OperationReport -Operation "ClearSource" -Results $cleanupResults -OutputFolderPath $resolvedOutputFolder -RunTimestamp $runTimestamp -InputSnapshotPath "" -PreChangeSnapshotPath $preCleanupSnapshotPath -PostChangeSnapshotPath $postCleanupSnapshotPath -TestUserIdentity $TestUser
}
