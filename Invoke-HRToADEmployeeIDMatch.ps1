<#
.SYNOPSIS
    Match Workday/HR CSV workers to AD users, optionally update employeeID, and optionally report provisioning-impact mismatches.

.DESCRIPTION
    Default mode is dry-run. AD users are preferred. If no AD user is matched, MailContacts and MailUsers are checked and reported only.
    The only write action is Set-ADUser -EmployeeID and only when -Apply is specified.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)] [string]$CsvPath,
    [string]$ReportPath = ".\Workday-AD-Readiness-Report.xlsx",
    [string]$SystemMismatchReportPath,

    [string]$CsvEmailColumn = "PrimaryWorkEmail",
    [string]$CsvFirstNameColumn = "PreferredFirstName",
    [string]$CsvLastNameColumn = "PreferredLastName",
    [string]$CsvJobTitleColumn = "BusinessTitle",
    [string]$CsvEmployeeIdColumn = "WorkerID",

    [string]$CsvDisplayNameColumn = "DisplayName",
    [string]$CsvCompanyColumn = "Company",
    [string]$CsvDepartmentColumn = "Department",
    [string]$CsvManagerEmailColumn = "PrimaryWorkEmail_Manager",
    [string]$CsvManagerReferenceColumn = "ManagerReference",
    [string]$CsvStreetAddressColumn = "StreetAddress",
    [string]$CsvCityColumn = "City",
    [string]$CsvCountryNameColumn = "Country",
    [string]$CsvCountryCodeColumn = "CountryCode",
    [string]$CsvStateColumn = "State",
    [string]$CsvOfficeColumn = "Office",
    [string]$CsvPostalCodeColumn = "PostalCode",
    [string]$CsvTelephoneColumn = "TelephoneNumber",
    [string]$CsvFaxColumn = "Fax",
    [string]$CsvMobileColumn = "Mobile",
    [string]$CsvPreferredLanguageColumn = "PreferredLanguage",
    [string]$CsvTargetOUColumn = "TargetOU",
    [string]$CsvReviewDecisionColumn = "ReviewDecision",
    [string]$CsvApprovedADDistinguishedNameColumn = "ApprovedADDistinguishedName",
    [string]$CsvApprovedByColumn = "ApprovedBy",
    [string]$CsvApprovalNotesColumn = "ApprovalNotes",
    [string]$ApprovedReviewDecision = "Approved",
    [string]$IgnoredReviewDecision = "Ignore",

    [string]$SearchBase,
    [switch]$Apply,
    [switch]$CheckSystemMismatch,
    [switch]$DoNotInstallMissingModules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-Text { param([object]$Value) if ($null -eq $Value) { return "" }; return ($Value.ToString().Trim() -replace "\s+", " ").ToLowerInvariant() }
function Remove-SmtpPrefix { param([object]$Value) if ($null -eq $Value) { return "" }; return (($Value.ToString()) -replace "^smtp:", "" -replace "^SMTP:", "") }
function Escape-Ldap { param([string]$Value) if ($null -eq $Value) { return "" }; $v=$Value; $v=$v -replace "\\","\5c"; $v=$v -replace "\*","\2a"; $v=$v -replace "\(","\28"; $v=$v -replace "\)","\29"; $v=$v -replace "`0","\00"; return $v }
function Get-CsvValue { param([object]$Row,[string]$ColumnName) if ($Row.PSObject.Properties.Name -contains $ColumnName) { return $Row.PSObject.Properties[$ColumnName].Value }; $reportColumnName="HR_$ColumnName"; if ($Row.PSObject.Properties.Name -contains $reportColumnName) { return $Row.PSObject.Properties[$reportColumnName].Value }; return $null }
function Get-HrLastNameValue { param([object]$Row) $preferredLastName=Get-CsvValue $Row $CsvLastNameColumn; if(-not [string]::IsNullOrWhiteSpace(([string]$preferredLastName))){return $preferredLastName}; return (Get-CsvValue $Row "LastName") }
function Get-HrLastNameSource { param([object]$Row) $preferredLastName=Get-CsvValue $Row $CsvLastNameColumn; if(-not [string]::IsNullOrWhiteSpace(([string]$preferredLastName))){return $CsvLastNameColumn}; $lastName=Get-CsvValue $Row "LastName"; if(-not [string]::IsNullOrWhiteSpace(([string]$lastName))){return "LastName"}; return "" }
function Get-ObjectValue { param([object]$Object,[string]$Name) if($null -eq $Object -or $Object.PSObject.Properties.Name -notcontains $Name){return $null}; return $Object.PSObject.Properties[$Name].Value }
function Get-ParentDn { param([string]$Dn) if ([string]::IsNullOrWhiteSpace($Dn)) { return "" }; $p=$Dn -split ",",2; if ($p.Count -lt 2) { return "" }; return $p[1] }
function Get-CnFromDistinguishedName { param([string]$Dn) if([string]::IsNullOrWhiteSpace($Dn)){return ""}; if($Dn -match "^CN=([^,]+)"){return (($Matches[1] -replace "\\, ",", ") -replace "\\,",",")}; return "" }
function Normalize-EmailAddress { param([object]$Value) return (Normalize-Text (Remove-SmtpPrefix $Value)) }

function Get-ObjectEmailAddresses {
    param([object]$Object)
    $addresses=New-Object System.Collections.Generic.List[string]
    foreach($value in @((Get-ObjectValue $Object "mail"),(Get-ObjectValue $Object "targetAddress"))){
        $normalized=Normalize-EmailAddress $value
        if(-not [string]::IsNullOrWhiteSpace($normalized) -and -not $addresses.Contains($normalized)){$addresses.Add($normalized)}
    }
    foreach($proxyAddress in @((Get-ObjectValue $Object "proxyAddresses"))){
        $normalized=Normalize-EmailAddress $proxyAddress
        if(-not [string]::IsNullOrWhiteSpace($normalized) -and -not $addresses.Contains($normalized)){$addresses.Add($normalized)}
    }
    return @($addresses)
}

function Test-ObjectEmailMatches {
    param([object]$Object,[object]$Email)
    $normalizedEmail=Normalize-EmailAddress $Email
    if([string]::IsNullOrWhiteSpace($normalizedEmail)){return $false}
    return (@(Get-ObjectEmailAddresses $Object) -contains $normalizedEmail)
}

function Get-StringSimilarityPercent {
    param([object]$ReferenceValue,[object]$CandidateValue)
    $reference=Normalize-Text $ReferenceValue
    $candidate=Normalize-Text $CandidateValue
    if([string]::IsNullOrWhiteSpace($reference) -and [string]::IsNullOrWhiteSpace($candidate)){return 100}
    if([string]::IsNullOrWhiteSpace($reference) -or [string]::IsNullOrWhiteSpace($candidate)){return 0}
    if($reference -eq $candidate){return 100}
    $rows=$reference.Length+1
    $cols=$candidate.Length+1
    $distance=New-Object 'int[,]' $rows,$cols
    for($i=0;$i -lt $rows;$i++){$distance[$i,0]=$i}
    for($j=0;$j -lt $cols;$j++){$distance[0,$j]=$j}
    for($i=1;$i -lt $rows;$i++){
        for($j=1;$j -lt $cols;$j++){
            $cost=if($reference[$i-1] -eq $candidate[$j-1]){0}else{1}
            $delete=$distance[($i-1),$j]+1
            $insert=$distance[$i,($j-1)]+1
            $replace=$distance[($i-1),($j-1)]+$cost
            $distance[$i,$j]=[Math]::Min([Math]::Min($delete,$insert),$replace)
        }
    }
    $maxLength=[Math]::Max($reference.Length,$candidate.Length)
    return [Math]::Round(((1-($distance[($rows-1),($cols-1)]/$maxLength))*100),0)
}

function Get-NameTokens {
    param([object]$Value)
    $normalized=Normalize-Text $Value
    if([string]::IsNullOrWhiteSpace($normalized)){return @()}
    $weakNameTokens=@("a","al","ap","bin","binti","da","de","del","della","der","di","dos","du","el","ibn","la","le","mac","mc","of","the","van","von")
    return @($normalized -replace "[^a-z0-9]+"," " -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_.Length -ge 3 -and $weakNameTokens -notcontains $_ } | Select-Object -Unique)
}

function Test-NameTokenOverlap {
    param([object]$ReferenceValue,[object]$CandidateValue)
    $referenceTokens=@(Get-NameTokens $ReferenceValue)
    $candidateTokens=@(Get-NameTokens $CandidateValue)
    foreach($referenceToken in $referenceTokens){
        if($candidateTokens -contains $referenceToken){return $true}
    }
    return $false
}

function Test-FirstNamePotentialMatch {
    param([object]$ReferenceValue,[object]$CandidateValue)
    $reference=Normalize-Text $ReferenceValue
    $candidate=Normalize-Text $CandidateValue
    if([string]::IsNullOrWhiteSpace($reference) -or [string]::IsNullOrWhiteSpace($candidate)){return $false}
    if($reference -eq $candidate){return $true}
    if(Test-NameTokenOverlap $reference $candidate){return $true}

    $referenceTokens=@(Get-NameTokens $reference)
    $candidateTokens=@(Get-NameTokens $candidate)
    foreach($referenceToken in $referenceTokens){
        foreach($candidateToken in $candidateTokens){
            $shorterLength=[Math]::Min($referenceToken.Length,$candidateToken.Length)
            if($shorterLength -ge 4 -and ($referenceToken.StartsWith($candidateToken) -or $candidateToken.StartsWith($referenceToken))){return $true}
            if((Get-StringSimilarityPercent $referenceToken $candidateToken) -ge 80){return $true}
        }
    }
    return $false
}

function Test-NameTokenPotentialMatch {
    param([object]$ReferenceValue,[object]$CandidateValue,[int]$SimilarityThreshold=70)
    $reference=Normalize-Text $ReferenceValue
    $candidate=Normalize-Text $CandidateValue
    if([string]::IsNullOrWhiteSpace($reference) -or [string]::IsNullOrWhiteSpace($candidate)){return $false}
    if($reference -eq $candidate){return $true}
    if(Test-NameTokenOverlap $reference $candidate){return $true}

    $referenceTokens=@(Get-NameTokens $reference)
    $candidateTokens=@(Get-NameTokens $candidate)
    foreach($referenceToken in $referenceTokens){
        foreach($candidateToken in $candidateTokens){
            $shorterLength=[Math]::Min($referenceToken.Length,$candidateToken.Length)
            if($shorterLength -ge 4 -and ($referenceToken.StartsWith($candidateToken) -or $candidateToken.StartsWith($referenceToken))){return $true}
            if((Get-StringSimilarityPercent $referenceToken $candidateToken) -ge $SimilarityThreshold){return $true}
        }
    }
    return $false
}

function Test-LastNamePotentialMatch {
    param([object]$ReferenceValue,[object]$CandidateValue)
    $referenceTokens=@(Get-NameTokens $ReferenceValue)
    $candidateTokens=@(Get-NameTokens $CandidateValue)
    if($referenceTokens.Count -eq 0 -or $candidateTokens.Count -eq 0){return $false}
    $overlapCount=@($referenceTokens|Where-Object{$candidateTokens -contains $_}).Count
    if($overlapCount -ge 2){return $true}
    if($referenceTokens.Count -gt 1 -and $candidateTokens.Count -gt 1 -and $overlapCount -lt 2){return $false}
    return (Test-NameTokenPotentialMatch -ReferenceValue $ReferenceValue -CandidateValue $CandidateValue -SimilarityThreshold 78)
}

function Get-CandidateFirstNameValue {
    param([object]$Object)
    return (@(
        (Get-ObjectValue $Object "givenName"),
        (Get-ObjectValue $Object "displayName"),
        (Get-ObjectValue $Object "cn"),
        (Get-ObjectValue $Object "name"),
        (Get-CnFromDistinguishedName (Get-ObjectValue $Object "DistinguishedName"))
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join " "
}

function Get-CandidateLastNameValue {
    param([object]$Object)
    return (@(
        (Get-ObjectValue $Object "sn"),
        (Get-ObjectValue $Object "displayName"),
        (Get-ObjectValue $Object "cn"),
        (Get-ObjectValue $Object "name"),
        (Get-CnFromDistinguishedName (Get-ObjectValue $Object "DistinguishedName"))
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join " "
}

function Test-PotentialPersonNameMatch {
    param([string]$FirstName,[string]$LastName,[object]$Object)
    $candidateFirst=Get-CandidateFirstNameValue $Object
    $candidateLast=Get-CandidateLastNameValue $Object
    return ((Test-FirstNamePotentialMatch $FirstName $candidateFirst) -and (Test-LastNamePotentialMatch $LastName $candidateLast))
}

function Ensure-ImportExcelModule {
    param([switch]$DoNotInstallMissingModules)
    if (Get-Module -ListAvailable -Name ImportExcel) { Import-Module ImportExcel -ErrorAction Stop; return $true }
    if ($DoNotInstallMissingModules) { Write-Warning "ImportExcel not installed; CSV fallback will be used."; return $false }
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop }
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne "Trusted") { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop }
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        return $true
    } catch { Write-Warning "Could not install ImportExcel. CSV fallback will be used. Error: $($_.Exception.Message)"; return $false }
}

function Export-ReportWorksheet {
    param(
        [Parameter(ValueFromPipeline)] [object]$InputObject,
        [string]$Path,
        [string]$WorksheetName,
        [string]$TableName,
        [switch]$Append,
        [string[]]$NoNumberConversionColumns = @()
    )
    begin {
        $items=New-Object System.Collections.Generic.List[object]
    }
    process {
        foreach($item in @($InputObject)){
            if($null -ne $item){$items.Add($item)}
        }
    }
    end {
        $params=@{
            Path=$Path
            WorksheetName=$WorksheetName
            TableName=$TableName
            AutoSize=$true
            FreezeTopRow=$true
            BoldTopRow=$true
        }
        if($Append){$params["Append"]=$true}else{$params["ClearSheet"]=$true}
        $exportExcelCommand=Get-Command Export-Excel -ErrorAction SilentlyContinue
        if($NoNumberConversionColumns.Count -gt 0 -and $exportExcelCommand -and $exportExcelCommand.Parameters.ContainsKey("NoNumberConversion")){
            $params["NoNumberConversion"]=$NoNumberConversionColumns
        }
        $items | Export-Excel @params
    }
}

function Get-DirectoryObject {
    param([string]$LdapFilter,[string]$SearchBase)
    $props = @("objectClass","cn","name","mail","proxyAddresses","targetAddress","mailNickname","givenName","sn","displayName","company","department","manager","title","streetAddress","l","co","c","st","physicalDeliveryOfficeName","postalCode","telephoneNumber","facsimileTelephoneNumber","mobile","preferredLanguage","employeeID","userPrincipalName","sAMAccountName","enabled","distinguishedName")
    if ([string]::IsNullOrWhiteSpace($SearchBase)) { return @(Get-ADObject -LDAPFilter $LdapFilter -Properties $props) }
    return @(Get-ADObject -LDAPFilter $LdapFilter -SearchBase $SearchBase -Properties $props)
}

function Get-RecipientType {
    param([object]$Object)
    $classes=@((Get-ObjectValue $Object "objectClass"))
    if ($classes -contains "contact") { return "MailContact" }
    if (($classes -contains "user") -and -not [string]::IsNullOrWhiteSpace((Get-ObjectValue $Object "targetAddress"))) { return "MailUser" }
    if ($classes -contains "user") { return "ADUser" }
    return "Unknown"
}

function Test-IsADUserObject {
    param([object]$Object)
    return (@((Get-ObjectValue $Object "objectClass")) -contains "user")
}

function Find-ADUserByEmail {
    param([string]$Email,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($Email)) { return @() }
    $e=Escape-Ldap $Email; $smtp1=Escape-Ldap "smtp:$Email"; $smtp2=Escape-Ldap "SMTP:$Email"
    $filter="(&(objectCategory=person)(objectClass=user)(|(mail=$e)(userPrincipalName=$e)(proxyAddresses=$smtp1)(proxyAddresses=$smtp2)(targetAddress=$smtp1)(targetAddress=$smtp2)(targetAddress=$e)))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase)
}

function Find-ADUserByEmployeeId {
    param([string]$EmployeeId,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($EmployeeId)) { return @() }
    $e=Escape-Ldap $EmployeeId
    return @(Get-DirectoryObject -LdapFilter "(&(objectCategory=person)(objectClass=user)(employeeID=$e))" -SearchBase $SearchBase)
}

function Find-ADUserByName {
    param([string]$FirstName,[string]$LastName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) { return @() }
    $f=Escape-Ldap $FirstName; $l=Escape-Ldap $LastName
    return @(Get-DirectoryObject -LdapFilter "(&(objectCategory=person)(objectClass=user)(givenName=$f)(sn=$l))" -SearchBase $SearchBase)
}

function Find-ADUserByDistinguishedName {
    param([string]$DistinguishedName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return @() }
    $d=Escape-Ldap $DistinguishedName
    return @(Get-DirectoryObject -LdapFilter "(&(objectCategory=person)(objectClass=user)(distinguishedName=$d))" -SearchBase $SearchBase)
}

function Find-DirectoryObjectByDistinguishedName {
    param([string]$DistinguishedName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return @() }
    $d=Escape-Ldap $DistinguishedName
    return @(Get-DirectoryObject -LdapFilter "(distinguishedName=$d)" -SearchBase $SearchBase)
}

function Find-ADUserByLastNameAndTitle {
    param([string]$LastName,[string]$Title,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($LastName) -or [string]::IsNullOrWhiteSpace($Title)) { return @() }
    $l=Escape-Ldap $LastName; $t=Escape-Ldap $Title
    return @(Get-DirectoryObject -LdapFilter "(&(objectCategory=person)(objectClass=user)(sn=$l)(title=$t))" -SearchBase $SearchBase)
}

function Find-ADUserByPotentialNameAndTitle {
    param([string]$FirstName,[string]$LastName,[string]$Title,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName) -or [string]::IsNullOrWhiteSpace($Title)) { return @() }
    $t=Escape-Ldap $Title
    $lastTokens=@(Get-NameTokens $LastName)
    if($lastTokens.Count -eq 0){return @()}
    $surnameClauses=@()
    foreach($lastToken in $lastTokens){
        $escapedToken=Escape-Ldap $lastToken
        $surnameClauses+="(sn=*$escapedToken*)"
        $surnameClauses+="(displayName=*$escapedToken*)"
        $surnameClauses+="(cn=*$escapedToken*)"
        $surnameClauses+="(name=*$escapedToken*)"
    }
    $surnameFilter=$surnameClauses -join ""
    $filter="(&(objectCategory=person)(objectClass=user)(title=$t)(|$surnameFilter))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase | Where-Object { Test-PotentialPersonNameMatch -FirstName $FirstName -LastName $LastName -Object $_ })
}

function Find-ADUserByPotentialName {
    param([string]$FirstName,[string]$LastName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) { return @() }
    $lastTokens=@(Get-NameTokens $LastName)
    if($lastTokens.Count -eq 0){return @()}
    $surnameClauses=@()
    foreach($lastToken in $lastTokens){
        $escapedToken=Escape-Ldap $lastToken
        $surnameClauses+="(sn=*$escapedToken*)"
        $surnameClauses+="(displayName=*$escapedToken*)"
        $surnameClauses+="(cn=*$escapedToken*)"
        $surnameClauses+="(name=*$escapedToken*)"
    }
    $surnameFilter=$surnameClauses -join ""
    $filter="(&(objectCategory=person)(objectClass=user)(|$surnameFilter))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase | Where-Object { Test-PotentialPersonNameMatch -FirstName $FirstName -LastName $LastName -Object $_ })
}

function Find-MailRecipientByEmail {
    param([string]$Email,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($Email)) { return @() }
    $e=Escape-Ldap $Email; $smtp1=Escape-Ldap "smtp:$Email"; $smtp2=Escape-Ldap "SMTP:$Email"
    $filter="(&(|(objectClass=contact)(&(objectCategory=person)(objectClass=user)(targetAddress=*)))(|(mail=$e)(userPrincipalName=$e)(proxyAddresses=$smtp1)(proxyAddresses=$smtp2)(targetAddress=$smtp1)(targetAddress=$smtp2)(targetAddress=$e)))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase)
}

function Find-MailRecipientByName {
    param([string]$FirstName,[string]$LastName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) { return @() }
    $f=Escape-Ldap $FirstName; $l=Escape-Ldap $LastName
    $filter="(&(|(objectClass=contact)(&(objectCategory=person)(objectClass=user)(targetAddress=*)))(givenName=$f)(sn=$l))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase)
}

function Find-MailRecipientByLastNameAndTitle {
    param([string]$LastName,[string]$Title,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($LastName) -or [string]::IsNullOrWhiteSpace($Title)) { return @() }
    $l=Escape-Ldap $LastName; $t=Escape-Ldap $Title
    $filter="(&(|(objectClass=contact)(&(objectCategory=person)(objectClass=user)(targetAddress=*)))(sn=$l)(title=$t))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase)
}

function Find-MailRecipientByPotentialNameAndTitle {
    param([string]$FirstName,[string]$LastName,[string]$Title,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName) -or [string]::IsNullOrWhiteSpace($Title)) { return @() }
    $t=Escape-Ldap $Title
    $lastTokens=@(Get-NameTokens $LastName)
    if($lastTokens.Count -eq 0){return @()}
    $surnameClauses=@()
    foreach($lastToken in $lastTokens){
        $escapedToken=Escape-Ldap $lastToken
        $surnameClauses+="(sn=*$escapedToken*)"
        $surnameClauses+="(displayName=*$escapedToken*)"
        $surnameClauses+="(cn=*$escapedToken*)"
        $surnameClauses+="(name=*$escapedToken*)"
    }
    $surnameFilter=$surnameClauses -join ""
    $filter="(&(|(objectClass=contact)(&(objectCategory=person)(objectClass=user)(targetAddress=*)))(title=$t)(|$surnameFilter))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase | Where-Object { Test-PotentialPersonNameMatch -FirstName $FirstName -LastName $LastName -Object $_ })
}

function Find-MailRecipientByPotentialName {
    param([string]$FirstName,[string]$LastName,[string]$SearchBase)
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) { return @() }
    $lastTokens=@(Get-NameTokens $LastName)
    if($lastTokens.Count -eq 0){return @()}
    $surnameClauses=@()
    foreach($lastToken in $lastTokens){
        $escapedToken=Escape-Ldap $lastToken
        $surnameClauses+="(sn=*$escapedToken*)"
        $surnameClauses+="(displayName=*$escapedToken*)"
        $surnameClauses+="(cn=*$escapedToken*)"
        $surnameClauses+="(name=*$escapedToken*)"
    }
    $surnameFilter=$surnameClauses -join ""
    $filter="(&(|(objectClass=contact)(&(objectCategory=person)(objectClass=user)(targetAddress=*)))(|$surnameFilter))"
    return @(Get-DirectoryObject -LdapFilter $filter -SearchBase $SearchBase | Where-Object { Test-PotentialPersonNameMatch -FirstName $FirstName -LastName $LastName -Object $_ })
}

function Test-EmailMatch {
    param([object]$Object,[string]$Email)
    $needle=Normalize-Text $Email
    if ([string]::IsNullOrWhiteSpace($needle)) { return $false }
    foreach ($v in @((Get-ObjectValue $Object "mail"),(Get-ObjectValue $Object "userPrincipalName"),(Remove-SmtpPrefix (Get-ObjectValue $Object "targetAddress")))) { if ((Normalize-Text $v) -eq $needle) { return $true } }
    foreach ($p in @((Get-ObjectValue $Object "proxyAddresses"))) { if ((Normalize-Text (Remove-SmtpPrefix $p)) -eq $needle) { return $true } }
    return $false
}

function Format-PotentialMatchSummary {
    param([object[]]$Matches,[string]$ReferenceFirstName)
    $items=foreach($match in @($Matches)){
        $identity=@(
            (Get-ObjectValue $match "sAMAccountName"),
            (Get-ObjectValue $match "userPrincipalName"),
            (Get-ObjectValue $match "mail"),
            (Remove-SmtpPrefix (Get-ObjectValue $match "targetAddress")),
            (Get-ObjectValue $match "mailNickname")
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
        $similarity=Get-StringSimilarityPercent $ReferenceFirstName (Get-ObjectValue $match "givenName")
        @(
            "Type=$(Get-RecipientType $match)",
            "DisplayName=$(Get-ObjectValue $match "displayName")",
            "CN=$(Get-ObjectValue $match "cn")",
            "Name=$(Get-ObjectValue $match "name")",
            "FirstName=$(Get-ObjectValue $match "givenName")",
            "LastName=$(Get-ObjectValue $match "sn")",
            "Title=$(Get-ObjectValue $match "title")",
            "Identity=$identity",
            "FirstNameSimilarity=$similarity%",
            "DN=$(Get-ObjectValue $match "DistinguishedName")"
        ) -join "; "
    }
    return (@($items) -join " | ")
}

function New-Row {
    param([object]$HrRow,[object]$Object,[string]$MatchStatus,[string]$MatchedBy,[bool]$UpdateEligible,[string]$Action,[string]$ActionResult,[string]$Notes,[int]$CsvRow,[int]$PotentialMatchCount=0,[string]$PotentialMatches="")
    $hrEmail=Get-CsvValue $HrRow $CsvEmailColumn; $hrFirst=Get-CsvValue $HrRow $CsvFirstNameColumn; $hrLast=Get-HrLastNameValue $HrRow; $hrLastSource=Get-HrLastNameSource $HrRow; $hrTitle=Get-CsvValue $HrRow $CsvJobTitleColumn; $hrEmp=Get-CsvValue $HrRow $CsvEmployeeIdColumn
    $reviewDecision=Get-CsvValue $HrRow $CsvReviewDecisionColumn; $approvedDn=Get-CsvValue $HrRow $CsvApprovedADDistinguishedNameColumn; $approvedBy=Get-CsvValue $HrRow $CsvApprovedByColumn; $approvalNotes=Get-CsvValue $HrRow $CsvApprovalNotesColumn
    $type=if($Object){Get-RecipientType $Object}else{""}
    $row=[ordered]@{}
    $row["CsvRow"]=$CsvRow
    $row["HR_$CsvEmployeeIdColumn"]=$hrEmp; $row["HR_$CsvEmailColumn"]=$hrEmail; $row["HR_$CsvFirstNameColumn"]=$hrFirst; $row["HR_$CsvLastNameColumn"]=$hrLast; $row["HR_$CsvJobTitleColumn"]=$hrTitle
    $row["HR_LastNameSource"]=$hrLastSource
    $row["MatchStatus"]=$MatchStatus; $row["MatchedBy"]=$MatchedBy; $row["UpdateEligible"]=$UpdateEligible; $row["EmployeeIDNeedsUpdate"] = if($Object){(Normalize-Text (Get-ObjectValue $Object "employeeID")) -ne (Normalize-Text $hrEmp)}else{$false}
    $row["Action"]=$Action; $row["ActionResult"]=$ActionResult; $row["Notes"]=$Notes
    $row["ReviewDecision"]=$reviewDecision; $row["ApprovedADDistinguishedName"]=$approvedDn; $row["ApprovedBy"]=$approvedBy; $row["ApprovalNotes"]=$approvalNotes
    $row["PotentialMatchCount"]=$PotentialMatchCount; $row["PotentialMatches"]=$PotentialMatches
    $row["AD_ObjectType"]=$type; $row["AD_SamAccountName"]=Get-ObjectValue $Object "sAMAccountName"; $row["AD_UserPrincipalName"]=Get-ObjectValue $Object "userPrincipalName"; $row["AD_Mail"]=Get-ObjectValue $Object "mail"; $row["AD_TargetAddress"]=Get-ObjectValue $Object "targetAddress"; $row["AD_MailNickname"]=Get-ObjectValue $Object "mailNickname"; $row["AD_CN"]=Get-ObjectValue $Object "cn"; $row["AD_Name"]=Get-ObjectValue $Object "name"
    $row["AD_FirstName"]=Get-ObjectValue $Object "givenName"; $row["AD_LastName"]=Get-ObjectValue $Object "sn"; $row["AD_DisplayName"]=Get-ObjectValue $Object "displayName"; $row["AD_JobTitle"]=Get-ObjectValue $Object "title"; $row["AD_EmployeeID_Current"]=Get-ObjectValue $Object "employeeID"; $row["AD_Enabled"]=Get-ObjectValue $Object "Enabled"; $row["AD_DistinguishedName"]=Get-ObjectValue $Object "DistinguishedName"
    $row["EmailMatchesAD"] = if($Object){Test-EmailMatch $Object $hrEmail}else{$false}; $row["FirstNameMatchesAD"] = if($Object){(Normalize-Text (Get-ObjectValue $Object "givenName")) -eq (Normalize-Text $hrFirst)}else{$false}; $row["FirstNameSimilarityPercent"] = if($Object){Get-StringSimilarityPercent $hrFirst (Get-ObjectValue $Object "givenName")}else{$null}; $row["LastNameMatchesAD"] = if($Object){(Normalize-Text (Get-ObjectValue $Object "sn")) -eq (Normalize-Text $hrLast)}else{$false}; $row["JobTitleMatchesAD"] = if($Object){(Normalize-Text (Get-ObjectValue $Object "title")) -eq (Normalize-Text $hrTitle)}else{$false}
    return [pscustomobject]$row
}

function Add-MailRecipientFallback {
    param([int]$CsvRow,[object]$HrRow,[string]$Email,[string]$FirstName,[string]$LastName,[string]$Title,[string]$SearchBase,[System.Collections.Generic.List[object]]$Results)
    $matches=@()
    if (-not [string]::IsNullOrWhiteSpace($Email)) { $matches=@(Find-MailRecipientByEmail $Email $SearchBase) }
    if ($matches.Count -eq 1) { $t=Get-RecipientType $matches[0]; $Results.Add((New-Row $HrRow $matches[0] "Matched to $t - No AD user update" "MailRecipient Email" $false "No action" "Skipped" "No AD user was matched. A $t was matched by email. Report-only." $CsvRow)); return $true }
    if ($matches.Count -gt 1) { $Results.Add((New-Row $HrRow $null "Ambiguous - Multiple MailRecipient Email Matches" "MailRecipient Email" $false "No action" "Skipped" "Multiple MailContacts/MailUsers matched the HR email value. Manual review required." $CsvRow)); return $true }
    $nameMatches=@(Find-MailRecipientByName $FirstName $LastName $SearchBase)
    if ($nameMatches.Count -eq 0) {
        $partialMatches=@(Find-MailRecipientByPotentialNameAndTitle $FirstName $LastName $Title $SearchBase)
        if($partialMatches.Count -eq 1){
            $t=Get-RecipientType $partialMatches[0]
            $Results.Add((New-Row $HrRow $partialMatches[0] "Potential $t Match - Name Variant" "MailRecipient Partial/Token Name + JobTitle" $false "No action" "Skipped" "No AD user was matched. A $t matched job title plus partial or tokenized first/last name values. Manual review required." $CsvRow 1 (Format-PotentialMatchSummary $partialMatches $FirstName)))
        }elseif($partialMatches.Count -gt 1){
            $Results.Add((New-Row $HrRow $null "Potential MailRecipient Matches - Name Variant" "MailRecipient Partial/Token Name + JobTitle" $false "No action" "Skipped" "No AD user was matched. Multiple MailContacts/MailUsers matched job title plus partial or tokenized first/last name values. Review the PotentialMatches column." $CsvRow $partialMatches.Count (Format-PotentialMatchSummary $partialMatches $FirstName)))
        }
        if($partialMatches.Count -gt 0){return $true}

        $titleMismatchMatches=@(Find-MailRecipientByPotentialName $FirstName $LastName $SearchBase | Where-Object { (Normalize-Text (Get-ObjectValue $_ "title")) -ne (Normalize-Text $Title) })
        if($titleMismatchMatches.Count -eq 1){
            $t=Get-RecipientType $titleMismatchMatches[0]
            $Results.Add((New-Row $HrRow $titleMismatchMatches[0] "Potential $t Match - Name Variant Title Mismatch" "MailRecipient Partial/Token Name only" $false "No action" "Skipped" "No AD user was matched. A $t matched partial or tokenized first/last name values, but job title differs or is missing. Manual review required." $CsvRow 1 (Format-PotentialMatchSummary $titleMismatchMatches $FirstName)))
        }elseif($titleMismatchMatches.Count -gt 1){
            $Results.Add((New-Row $HrRow $null "Potential MailRecipient Matches - Name Variant Title Mismatch" "MailRecipient Partial/Token Name only" $false "No action" "Skipped" "No AD user was matched. Multiple MailContacts/MailUsers matched partial or tokenized first/last name values, but job title differs or is missing. Review the PotentialMatches column." $CsvRow $titleMismatchMatches.Count (Format-PotentialMatchSummary $titleMismatchMatches $FirstName)))
        }
        return ($titleMismatchMatches.Count -gt 0)
    }
    $titleMatches=@($nameMatches | Where-Object { (Normalize-Text (Get-ObjectValue $_ "title")) -eq (Normalize-Text $Title) })
    if ($titleMatches.Count -eq 1) { $t=Get-RecipientType $titleMatches[0]; $Results.Add((New-Row $HrRow $titleMatches[0] "Matched to $t - No AD user update" "MailRecipient FirstName + LastName + JobTitle" $false "No action" "Skipped" "No AD user was matched. A $t matched by name and title. Report-only." $CsvRow)); return $true }
    if ($titleMatches.Count -gt 1) { $Results.Add((New-Row $HrRow $null "Ambiguous - Multiple MailRecipient Name and Title Matches" "MailRecipient FirstName + LastName + JobTitle" $false "No action" "Skipped" "Multiple MailContacts/MailUsers matched name and title. Manual review required." $CsvRow)); return $true }
    if ($nameMatches.Count -eq 1) { $t=Get-RecipientType $nameMatches[0]; $Results.Add((New-Row $HrRow $nameMatches[0] "Potential $t Match - Title Mismatch" "MailRecipient FirstName + LastName only" $false "No action" "Skipped" "A single $t matched name but title differs. Manual review required." $CsvRow)); return $true }
    $Results.Add((New-Row $HrRow $null "Ambiguous - Multiple MailRecipient Name Matches" "MailRecipient FirstName + LastName" $false "No action" "Skipped" "Multiple MailContacts/MailUsers matched first name and last name. Manual review required." $CsvRow)); return $true
}

function Resolve-ManagerReferenceToADUser {
    param([object]$ManagerReference,[object[]]$HrRows,[string]$SearchBase)
    $reference=([string]$ManagerReference).Trim()
    if([string]::IsNullOrWhiteSpace($reference)){
        return [pscustomobject]@{Resolved=$true; DistinguishedName=""; Status="No manager reference supplied."; MatchCount=0}
    }

    $employeeIdMatches=@(Find-ADUserByEmployeeId $reference $SearchBase)
    if($employeeIdMatches.Count -eq 1){
        return [pscustomobject]@{Resolved=$true; DistinguishedName=(Get-ObjectValue $employeeIdMatches[0] "DistinguishedName"); Status="Resolved ManagerReference to AD user by employeeID."; MatchCount=1}
    }
    if($employeeIdMatches.Count -gt 1){
        return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="ManagerReference matched multiple AD users by employeeID."; MatchCount=$employeeIdMatches.Count}
    }

    $managerRows=@($HrRows | Where-Object { (Normalize-Text (Get-CsvValue $_ $CsvEmployeeIdColumn)) -eq (Normalize-Text $reference) })
    if($managerRows.Count -eq 0){
        return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="ManagerReference did not match an AD employeeID or a WorkerID row in the HR CSV."; MatchCount=0}
    }
    if($managerRows.Count -gt 1){
        return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="ManagerReference matched multiple WorkerID rows in the HR CSV."; MatchCount=$managerRows.Count}
    }

    $managerRow=$managerRows[0]
    $managerEmail=([string](Get-CsvValue $managerRow $CsvEmailColumn)).Trim()
    $managerFirst=([string](Get-CsvValue $managerRow $CsvFirstNameColumn)).Trim()
    $managerLast=([string](Get-HrLastNameValue $managerRow)).Trim()
    $managerTitle=([string](Get-CsvValue $managerRow $CsvJobTitleColumn)).Trim()

    if(-not [string]::IsNullOrWhiteSpace($managerEmail)){
        $emailMatches=@(Find-ADUserByEmail $managerEmail $SearchBase)
        if($emailMatches.Count -eq 1){
            return [pscustomobject]@{Resolved=$true; DistinguishedName=(Get-ObjectValue $emailMatches[0] "DistinguishedName"); Status="Resolved ManagerReference through manager HR row email."; MatchCount=1}
        }
        if($emailMatches.Count -gt 1){
            return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="Manager HR row email matched multiple AD users."; MatchCount=$emailMatches.Count}
        }
    }

    $nameMatches=@(Find-ADUserByName $managerFirst $managerLast $SearchBase)
    $titleMatches=@($nameMatches | Where-Object { (Normalize-Text (Get-ObjectValue $_ "title")) -eq (Normalize-Text $managerTitle) })
    if($titleMatches.Count -eq 1){
        return [pscustomobject]@{Resolved=$true; DistinguishedName=(Get-ObjectValue $titleMatches[0] "DistinguishedName"); Status="Resolved ManagerReference through manager HR row name and title."; MatchCount=1}
    }
    if($titleMatches.Count -gt 1){
        return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="Manager HR row name and title matched multiple AD users."; MatchCount=$titleMatches.Count}
    }

    return [pscustomobject]@{Resolved=$false; DistinguishedName=""; Status="ManagerReference matched an HR row, but that manager row did not resolve to a single AD user."; MatchCount=0}
}

function New-MismatchRows {
    param([int]$CsvRow,[object]$HrRow,[object]$User,[array]$Map,[object[]]$HrRows,[string]$SearchBase,[string]$MatchedBy,[string]$MatchStatus,[bool]$EmployeeIDAlreadyPresent,[string]$EmployeeIDActionGuidance)
    $rows=New-Object System.Collections.Generic.List[object]
    foreach($m in $Map){
        if(-not $m.Compare){continue}
        if($m.CsvColumn -eq $CsvLastNameColumn){
            $hr=Get-HrLastNameValue $HrRow
            if([string]::IsNullOrWhiteSpace(([string]$hr).Trim())){continue}
        }else{
            if(($HrRow.PSObject.Properties.Name -notcontains $m.CsvColumn) -and ($HrRow.PSObject.Properties.Name -notcontains "HR_$($m.CsvColumn)")){continue}
            $hr=Get-CsvValue $HrRow $m.CsvColumn
        }
        $managerResolution=$null
        if($m.CompareMode -eq "ManagerReferenceToDN"){
            $managerResolution=Resolve-ManagerReferenceToADUser -ManagerReference $hr -HrRows $HrRows -SearchBase $SearchBase
            $ad=Get-ObjectValue $User "manager"
            $proposed=$managerResolution.DistinguishedName
            $isMatch=if([string]::IsNullOrWhiteSpace(([string]$hr).Trim())){[string]::IsNullOrWhiteSpace(([string]$ad).Trim())}elseif($managerResolution.Resolved){(Normalize-Text $proposed) -eq (Normalize-Text $ad)}else{$false}
        }else{
            $ad=if($m.CompareMode -eq "Email"){(@(Get-ObjectEmailAddresses $User) -join "; ")}elseif($m.CompareMode -eq "ParentOU"){Get-ParentDn (Get-ObjectValue $User "DistinguishedName")}elseif($m.CompareMode -eq "ManagerEmailToDN"){Get-ObjectValue $User "manager"}else{Get-ObjectValue $User $m.ADAttribute}
            $proposed=$hr
            $isMatch=if($m.CompareMode -eq "Email"){Test-ObjectEmailMatches $User $hr}else{(Normalize-Text $hr) -eq (Normalize-Text $ad)}
        }
        if(-not $isMatch){
            $mismatchType=if($m.CompareMode -eq "ManagerReferenceToDN" -and $managerResolution -and -not $managerResolution.Resolved){"Manager reference unresolved"}else{"Value differs"}
            $whatWillChange=if($m.CompareMode -eq "ManagerReferenceToDN" -and $managerResolution -and -not $managerResolution.Resolved){"ManagerReference '$hr' could not be resolved to exactly one AD manager DN. Review before provisioning updates AD manager."}else{"AD $($m.ADAttribute) would change from '$ad' to '$proposed'."}
            $rows.Add([pscustomobject]@{
                CsvRow=$CsvRow
                HR_WorkerID=(Get-CsvValue $HrRow $CsvEmployeeIdColumn)
                HR_FirstName=(Get-CsvValue $HrRow $CsvFirstNameColumn)
                HR_LastName=(Get-HrLastNameValue $HrRow)
                HR_LastNameSource=(Get-HrLastNameSource $HrRow)
                HR_BusinessTitle=(Get-CsvValue $HrRow $CsvJobTitleColumn)
                HR_Department=(Get-CsvValue $HrRow $CsvDepartmentColumn)
                HR_ManagerReference=(Get-CsvValue $HrRow $CsvManagerReferenceColumn)
                AD_UserPrincipalName=(Get-ObjectValue $User "userPrincipalName")
                AD_EmployeeID_Current=(Get-ObjectValue $User "employeeID")
                AD_SamAccountName=(Get-ObjectValue $User "sAMAccountName")
                AD_GivenName=(Get-ObjectValue $User "givenName")
                AD_Surname=(Get-ObjectValue $User "sn")
                AD_JobTitle=(Get-ObjectValue $User "title")
                AD_Department=(Get-ObjectValue $User "department")
                AD_ManagerDN=(Get-ObjectValue $User "manager")
                ProposedManagerResolution=if($managerResolution){$managerResolution.Status}else{""}
                MatchedBy=$MatchedBy
                MatchStatus=$MatchStatus
                EmployeeIDAlreadyPresent=$EmployeeIDAlreadyPresent
                EmployeeIDActionGuidance=$EmployeeIDActionGuidance
                WorkdayAttribute=$m.WorkdayAttribute
                CsvColumn=$m.CsvColumn
                ADAttribute=$m.ADAttribute
                MappingType=$m.MappingType
                CurrentADValue=$ad
                ProposedHRValue=$proposed
                WhatWillChange=$whatWillChange
                MismatchType=$mismatchType
                ProvisioningImpact="Potential AD change when Entra Workday provisioning updates this attribute."
            })
        }
    }
    return $rows
}

function New-EmptyMismatchReportRow {
    [pscustomobject]@{
        CsvRow=$null
        HR_WorkerID=$null
        HR_FirstName=$null
        HR_LastName=$null
        HR_LastNameSource=$null
        HR_BusinessTitle=$null
        HR_Department=$null
        HR_ManagerReference=$null
        AD_UserPrincipalName=$null
        AD_EmployeeID_Current=$null
        AD_SamAccountName=$null
        AD_GivenName=$null
        AD_Surname=$null
        AD_JobTitle=$null
        AD_Department=$null
        AD_ManagerDN=$null
        ProposedManagerResolution=$null
        MatchedBy=$null
        MatchStatus=$null
        EmployeeIDAlreadyPresent=$null
        EmployeeIDActionGuidance=$null
        WorkdayAttribute=$null
        CsvColumn=$null
        ADAttribute=$null
        MappingType=$null
        CurrentADValue=$null
        ProposedHRValue=$null
        WhatWillChange=$null
        MismatchType=$null
        ProvisioningImpact=$null
        Result="No mismatches found."
    }
}

function New-EmployeeIdApplyReportRows {
    param([object[]]$MatchResults)
    $rows=New-Object System.Collections.Generic.List[object]
    foreach($result in @($MatchResults)){
        $workerId=Get-ObjectValue $result "HR_$CsvEmployeeIdColumn"
        $currentEmployeeId=Get-ObjectValue $result "AD_EmployeeID_Current"
        $actionResult=Get-ObjectValue $result "ActionResult"
        $hasEmployeeId=($actionResult -eq "Updated" -or $actionResult -eq "Already correct")
        $outcome=if($actionResult -eq "Updated"){"Updated EmployeeID"}elseif($actionResult -eq "Already correct"){"Already had EmployeeID"}else{"Skipped"}
        $employeeIdAfter=if($hasEmployeeId){$workerId}else{$currentEmployeeId}
        $rows.Add([pscustomobject]@{
            CsvRow=(Get-ObjectValue $result "CsvRow")
            ApplyOutcome=$outcome
            Action=(Get-ObjectValue $result "Action")
            ActionResult=$actionResult
            EmployeeIDNowPresent=$hasEmployeeId
            EmployeeIDBefore=$currentEmployeeId
            EmployeeIDAfter=$employeeIdAfter
            HR_WorkerID=$workerId
            HR_FirstName=(Get-ObjectValue $result "HR_$CsvFirstNameColumn")
            HR_LastName=(Get-ObjectValue $result "HR_$CsvLastNameColumn")
            HR_LastNameSource=(Get-ObjectValue $result "HR_LastNameSource")
            HR_BusinessTitle=(Get-ObjectValue $result "HR_$CsvJobTitleColumn")
            MatchStatus=(Get-ObjectValue $result "MatchStatus")
            MatchedBy=(Get-ObjectValue $result "MatchedBy")
            UpdateEligible=(Get-ObjectValue $result "UpdateEligible")
            AD_ObjectType=(Get-ObjectValue $result "AD_ObjectType")
            AD_SamAccountName=(Get-ObjectValue $result "AD_SamAccountName")
            AD_UserPrincipalName=(Get-ObjectValue $result "AD_UserPrincipalName")
            AD_DistinguishedName=(Get-ObjectValue $result "AD_DistinguishedName")
            SkippedReason=if($outcome -eq "Skipped"){Get-ObjectValue $result "Notes"}else{""}
            Notes=(Get-ObjectValue $result "Notes")
        })
    }
    return $rows
}

Write-Host "Importing HR CSV: $CsvPath"
Import-Module ActiveDirectory -ErrorAction Stop
if(-not(Test-Path $CsvPath)){throw "CSV path not found: $CsvPath"}
$hrRows=@(Import-Csv $CsvPath)
$required=@($CsvEmailColumn,$CsvFirstNameColumn,$CsvJobTitleColumn,$CsvEmployeeIdColumn)
$cols=@($hrRows[0].PSObject.Properties.Name)
foreach($c in $required){if(($cols -notcontains $c) -and ($cols -notcontains "HR_$c")){throw "Required CSV column not found: '$c' or 'HR_$c'. Available columns: $($cols -join ', ')"}}
if(($cols -notcontains $CsvLastNameColumn) -and ($cols -notcontains "HR_$CsvLastNameColumn") -and ($cols -notcontains "LastName") -and ($cols -notcontains "HR_LastName")){throw "Required CSV last-name column not found: '$CsvLastNameColumn', 'HR_$CsvLastNameColumn', 'LastName', or 'HR_LastName'. Available columns: $($cols -join ', ')"}

$map=@(
    [pscustomobject]@{WorkdayAttribute="FirstName";CsvColumn=$CsvFirstNameColumn;ADAttribute="givenName";MappingType="Create + update";CompareMode="Direct";Compare=$true},
    [pscustomobject]@{WorkdayAttribute="PreferredLastName";CsvColumn=$CsvLastNameColumn;ADAttribute="sn";MappingType="Create + update";CompareMode="Direct";Compare=$true},
    [pscustomobject]@{WorkdayAttribute="PrimaryWorkEmail";CsvColumn=$CsvEmailColumn;ADAttribute="mail/proxyAddresses/targetAddress";MappingType="Create only";CompareMode="Email";Compare=$false},
    [pscustomobject]@{WorkdayAttribute="HR Business Title";CsvColumn=$CsvJobTitleColumn;ADAttribute="title";MappingType="Create + update";CompareMode="Direct";Compare=$true},
    [pscustomobject]@{WorkdayAttribute="Department";CsvColumn=$CsvDepartmentColumn;ADAttribute="department";MappingType="Create + update";CompareMode="Direct";Compare=$true},
    [pscustomobject]@{WorkdayAttribute="ManagerReference";CsvColumn=$CsvManagerReferenceColumn;ADAttribute="manager";MappingType="Create + update";CompareMode="ManagerReferenceToDN";Compare=$true},
    [pscustomobject]@{WorkdayAttribute="parentDistinguishedName";CsvColumn=$CsvTargetOUColumn;ADAttribute="distinguishedName";MappingType="Create + update";CompareMode="ParentOU";Compare=$false}
)

$results=New-Object System.Collections.Generic.List[object]
$mismatches=New-Object System.Collections.Generic.List[object]
$systemMismatchMatchSources=New-Object System.Collections.Generic.List[object]
$reservedExactAdUserDns=@{}
$confirmedAdUserDns=@{}
$approvedReviewDnUseCounts=@{}
$approvedReviewDnUseRows=@{}

$approvedScanRowNum=1
$approvedScanTotalRows=[Math]::Max($hrRows.Count,1)
foreach($r in $hrRows){
    $approvedScanRowNum++
    $approvedScanCurrentRow=$approvedScanRowNum-1
    $approvedScanPercentComplete=[Math]::Min(100,[Math]::Round((($approvedScanCurrentRow / $approvedScanTotalRows)*100),0))
    Write-Progress -Activity "Scanning approved review decisions" -Status "$approvedScanPercentComplete% complete - Row $approvedScanCurrentRow of $($hrRows.Count) (CSV row ${approvedScanRowNum})" -PercentComplete $approvedScanPercentComplete
    $reviewDecision=Get-CsvValue $r $CsvReviewDecisionColumn
    $approvedDn=Get-CsvValue $r $CsvApprovedADDistinguishedNameColumn
    if((Normalize-Text $reviewDecision) -ne (Normalize-Text $ApprovedReviewDecision) -or [string]::IsNullOrWhiteSpace($approvedDn)){continue}

    $approvedUsageMatches=@(Find-DirectoryObjectByDistinguishedName $approvedDn $SearchBase)
    if($approvedUsageMatches.Count -eq 1){
        $approvedUsageKey=Normalize-Text (Get-ObjectValue $approvedUsageMatches[0] "DistinguishedName")
    }else{
        $approvedUsageKey=Normalize-Text $approvedDn
    }
    if([string]::IsNullOrWhiteSpace($approvedUsageKey)){continue}
    if(-not $approvedReviewDnUseCounts.ContainsKey($approvedUsageKey)){
        $approvedReviewDnUseCounts[$approvedUsageKey]=0
        $approvedReviewDnUseRows[$approvedUsageKey]=New-Object System.Collections.Generic.List[int]
    }
    $approvedReviewDnUseCounts[$approvedUsageKey]++
    $approvedReviewDnUseRows[$approvedUsageKey].Add($approvedScanRowNum)
}
Write-Progress -Activity "Scanning approved review decisions" -Completed

$duplicateApprovedDns=@{}
foreach($approvedUsageKey in @($approvedReviewDnUseCounts.Keys)){
    if($approvedReviewDnUseCounts[$approvedUsageKey] -gt 1){
        $duplicateApprovedDns[$approvedUsageKey]=($approvedReviewDnUseRows[$approvedUsageKey] -join ", ")
    }
}

$reservationScanRowNum=1
$reservationScanTotalRows=[Math]::Max($hrRows.Count,1)
foreach($r in $hrRows){
    $reservationScanRowNum++
    $reservationScanCurrentRow=$reservationScanRowNum-1
    $reservationScanPercentComplete=[Math]::Min(100,[Math]::Round((($reservationScanCurrentRow / $reservationScanTotalRows)*100),0))
    Write-Progress -Activity "Reserving exact AD matches" -Status "$reservationScanPercentComplete% complete - Row $reservationScanCurrentRow of $($hrRows.Count) (CSV row ${reservationScanRowNum})" -PercentComplete $reservationScanPercentComplete
    $email=([string](Get-CsvValue $r $CsvEmailColumn)).Trim()
    $first=([string](Get-CsvValue $r $CsvFirstNameColumn)).Trim()
    $last=([string](Get-HrLastNameValue $r)).Trim()
    $title=([string](Get-CsvValue $r $CsvJobTitleColumn)).Trim()
    $emp=([string](Get-CsvValue $r $CsvEmployeeIdColumn)).Trim()
    if([string]::IsNullOrWhiteSpace($emp)){continue}

    $reviewDecision=Get-CsvValue $r $CsvReviewDecisionColumn
    if((Normalize-Text $reviewDecision) -eq (Normalize-Text $IgnoredReviewDecision)){continue}
    $approvedDn=Get-CsvValue $r $CsvApprovedADDistinguishedNameColumn
    if((Normalize-Text $reviewDecision) -eq (Normalize-Text $ApprovedReviewDecision) -and -not [string]::IsNullOrWhiteSpace($approvedDn)){
        $approvedMatches=@(Find-DirectoryObjectByDistinguishedName $approvedDn $SearchBase)
        if($approvedMatches.Count -eq 1){
            $approvedReservationKey=Normalize-Text (Get-ObjectValue $approvedMatches[0] "DistinguishedName")
        }else{
            $approvedReservationKey=Normalize-Text $approvedDn
        }
        if($approvedReservationKey -and $duplicateApprovedDns.ContainsKey($approvedReservationKey)){continue}
        if($approvedMatches.Count -eq 1 -and (Test-IsADUserObject $approvedMatches[0])){
            $reservedDn=Normalize-Text (Get-ObjectValue $approvedMatches[0] "DistinguishedName")
            if($reservedDn){$reservedExactAdUserDns[$reservedDn]=$true}
            continue
        }
    }

    $employeeIdMatches=@(Find-ADUserByEmployeeId $emp $SearchBase)
    if($employeeIdMatches.Count -eq 1){
        $reservedDn=Normalize-Text (Get-ObjectValue $employeeIdMatches[0] "DistinguishedName")
        if($reservedDn){$reservedExactAdUserDns[$reservedDn]=$true}
        continue
    }

    $emailMatches=@(Find-ADUserByEmail $email $SearchBase)
    if($emailMatches.Count -eq 1){
        $reservedDn=Normalize-Text (Get-ObjectValue $emailMatches[0] "DistinguishedName")
        if($reservedDn){$reservedExactAdUserDns[$reservedDn]=$true}
        continue
    }

    $nameMatches=@(Find-ADUserByName $first $last $SearchBase)
    $titleMatches=@($nameMatches | Where-Object{(Normalize-Text (Get-ObjectValue $_ "title")) -eq (Normalize-Text $title)})
    if($titleMatches.Count -eq 1){
        $reservedDn=Normalize-Text (Get-ObjectValue $titleMatches[0] "DistinguishedName")
        if($reservedDn){$reservedExactAdUserDns[$reservedDn]=$true}
    }
}
Write-Progress -Activity "Reserving exact AD matches" -Completed

$rowNum=1
$totalRows=[Math]::Max($hrRows.Count,1)
$matchingProgressActivity=if($CheckSystemMismatch){"Matching HR users to AD and checking system mismatches"}else{"Matching HR users to AD"}
foreach($r in $hrRows){
    $rowNum++
    $email=([string](Get-CsvValue $r $CsvEmailColumn)).Trim(); $first=([string](Get-CsvValue $r $CsvFirstNameColumn)).Trim(); $last=([string](Get-HrLastNameValue $r)).Trim(); $title=([string](Get-CsvValue $r $CsvJobTitleColumn)).Trim(); $emp=([string](Get-CsvValue $r $CsvEmployeeIdColumn)).Trim()
    $currentRow=$rowNum-1
    $percentComplete=[Math]::Min(100,[Math]::Round((($currentRow / $totalRows)*100),0))
    Write-Progress -Activity $matchingProgressActivity -Status "$percentComplete% complete - Row $currentRow of $($hrRows.Count) (CSV row ${rowNum}): $email" -PercentComplete $percentComplete
    $reviewDecision=Get-CsvValue $r $CsvReviewDecisionColumn
    if((Normalize-Text $reviewDecision) -eq (Normalize-Text $IgnoredReviewDecision)){
        $results.Add((New-Row $r $null "Ignored - ReviewDecision" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is Ignore. No matching or employeeID update is attempted for this row." $rowNum))
        continue
    }
    if([string]::IsNullOrWhiteSpace($emp)){ $results.Add((New-Row $r $null "Skipped - Missing EmployeeID" "" $false "No action" "Skipped" "HR EmployeeID is blank." $rowNum)); continue }
    $user=$null; $status=""; $by=""; $eligible=$false; $notes=""
    $approvedDn=Get-CsvValue $r $CsvApprovedADDistinguishedNameColumn
    if((Normalize-Text $reviewDecision) -eq (Normalize-Text $ApprovedReviewDecision)){
        if([string]::IsNullOrWhiteSpace($approvedDn)){
            $results.Add((New-Row $r $null "ReviewDecision Invalid - Missing ApprovedADDistinguishedName" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is approved, but ApprovedADDistinguishedName is blank. Add the approved AD user distinguishedName before using -Apply." $rowNum))
            continue
        }
        $approvedMatches=@(Find-DirectoryObjectByDistinguishedName $approvedDn $SearchBase)
        if($approvedMatches.Count -eq 1){
            $approvedDuplicateKey=Normalize-Text (Get-ObjectValue $approvedMatches[0] "DistinguishedName")
            $approvedDuplicateObject=$approvedMatches[0]
        }else{
            $approvedDuplicateKey=Normalize-Text $approvedDn
            $approvedDuplicateObject=$null
        }
        if($approvedDuplicateKey -and $duplicateApprovedDns.ContainsKey($approvedDuplicateKey)){
            $duplicateRows=$duplicateApprovedDns[$approvedDuplicateKey]
            $results.Add((New-Row $r $approvedDuplicateObject "ReviewDecision Invalid - Duplicate ApprovedADDistinguishedName" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is approved, but the same ApprovedADDistinguishedName is used on multiple CSV rows ($duplicateRows). None of the duplicated approved rows are processed. Keep this distinguishedName approved on exactly one row." $rowNum))
            continue
        }
        if($approvedMatches.Count -eq 1){
            $approvedObject=$approvedMatches[0]
            $approvedObjectType=Get-RecipientType $approvedObject
            if(Test-IsADUserObject $approvedObject){
                $user=$approvedObject;$status="Matched - Approved ReviewDecision";$by="ReviewDecision + ApprovedADDistinguishedName";$eligible=$true;$notes="Manual review approved this $approvedObjectType distinguishedName for employeeID update."
            }elseif($approvedObjectType -eq "MailContact"){
                $user=$approvedObject;$status="Approved MailContact - No AD user update";$by="ReviewDecision + ApprovedADDistinguishedName";$eligible=$false;$notes="Manual review approved this MailContact as the accepted match. No employeeID update is attempted because contacts are not AD user objects."
            }else{
                $results.Add((New-Row $r $approvedObject "ReviewDecision Invalid - ApprovedADDistinguishedName Unsupported Object Type" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is approved, but ApprovedADDistinguishedName resolved to an unsupported object type: $approvedObjectType." $rowNum))
                continue
            }
        }elseif($approvedMatches.Count -eq 0){
            $results.Add((New-Row $r $null "ReviewDecision Invalid - ApprovedADDistinguishedName Not Found" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is approved, but ApprovedADDistinguishedName did not resolve to a directory object." $rowNum))
            continue
        }else{
            $results.Add((New-Row $r $null "ReviewDecision Invalid - Multiple ApprovedADDistinguishedName Matches" "ReviewDecision" $false "No action" "Skipped" "ReviewDecision is approved, but ApprovedADDistinguishedName resolved to multiple directory objects. Manual review required." $rowNum))
            continue
        }
    }
    $employeeIdMatches=@(Find-ADUserByEmployeeId $emp $SearchBase)
    if($null -eq $user){
        if($employeeIdMatches.Count -eq 1){$user=$employeeIdMatches[0];$status="Matched";$by="WorkerID + AD employeeID";$eligible=$true;$notes="Single AD user matched by WorkerID to AD employeeID."}
        elseif($employeeIdMatches.Count -gt 1){$results.Add((New-Row $r $null "Ambiguous - Multiple EmployeeID Matches" "WorkerID + AD employeeID" $false "No action" "Skipped" "Multiple AD users matched the HR WorkerID value in AD employeeID." $rowNum));continue}
    }
    $emailMatches=@(Find-ADUserByEmail $email $SearchBase)
    if($null -eq $user){
        if($emailMatches.Count -eq 1){$user=$emailMatches[0];$status="Matched";$by="Email";$eligible=$true;$notes="Single AD user matched by HR email address."}
        elseif($emailMatches.Count -gt 1){$results.Add((New-Row $r $null "Ambiguous - Multiple Email Matches" "Email" $false "No action" "Skipped" "Multiple AD users matched the HR email value." $rowNum));continue}
    }
    if($null -eq $user){
        $nameMatches=@(Find-ADUserByName $first $last $SearchBase)
        if($nameMatches.Count -eq 0){
            $partialAdMatches=@(Find-ADUserByPotentialNameAndTitle $first $last $title $SearchBase | Where-Object { $candidateDn=Normalize-Text (Get-ObjectValue $_ "DistinguishedName"); -not $confirmedAdUserDns.ContainsKey($candidateDn) -and -not $reservedExactAdUserDns.ContainsKey($candidateDn) })
            if($partialAdMatches.Count -gt 0){
                if($partialAdMatches.Count -eq 1){
                    $results.Add((New-Row $r $partialAdMatches[0] "Potential Match - Name Variant" "Partial/Token Name + JobTitle" $false "No action" "Skipped" "An AD user matched job title plus partial or tokenized first/last name values. Review likely spelling, preferred-name, shortened-name, or multi-part-name differences before taking action." $rowNum 1 (Format-PotentialMatchSummary $partialAdMatches $first)))
                }else{
                    $results.Add((New-Row $r $null "Potential Matches - Name Variant" "Partial/Token Name + JobTitle" $false "No action" "Skipped" "Multiple AD users matched job title plus partial or tokenized first/last name values. Review the PotentialMatches column before taking action." $rowNum $partialAdMatches.Count (Format-PotentialMatchSummary $partialAdMatches $first)))
                }
                continue
            }
            $titleMismatchAdMatches=@(Find-ADUserByPotentialName $first $last $SearchBase | Where-Object { $candidateDn=Normalize-Text (Get-ObjectValue $_ "DistinguishedName"); ((Normalize-Text (Get-ObjectValue $_ "title")) -ne (Normalize-Text $title)) -and -not $confirmedAdUserDns.ContainsKey($candidateDn) -and -not $reservedExactAdUserDns.ContainsKey($candidateDn) })
            if($titleMismatchAdMatches.Count -gt 0){
                if($titleMismatchAdMatches.Count -eq 1){
                    $results.Add((New-Row $r $titleMismatchAdMatches[0] "Potential Match - Name Variant Title Mismatch" "Partial/Token Name only" $false "No action" "Skipped" "An AD user matched partial or tokenized first/last name values, but job title differs or is missing. Review likely spelling, preferred-name, shortened-name, multi-part-name, or title differences before taking action." $rowNum 1 (Format-PotentialMatchSummary $titleMismatchAdMatches $first)))
                }else{
                    $results.Add((New-Row $r $null "Potential Matches - Name Variant Title Mismatch" "Partial/Token Name only" $false "No action" "Skipped" "Multiple AD users matched partial or tokenized first/last name values, but job title differs or is missing. Review the PotentialMatches column before taking action." $rowNum $titleMismatchAdMatches.Count (Format-PotentialMatchSummary $titleMismatchAdMatches $first)))
                }
                continue
            }
            if(-not(Add-MailRecipientFallback $rowNum $r $email $first $last $title $SearchBase $results)){ $results.Add((New-Row $r $null "No Match" "" $false "No action" "Skipped" "No AD user, MailContact, or MailUser matched." $rowNum))}
            continue
        }
        $titleMatches=@($nameMatches | Where-Object{(Normalize-Text (Get-ObjectValue $_ "title")) -eq (Normalize-Text $title)})
        if($titleMatches.Count -eq 1){$user=$titleMatches[0];$status="Matched";$by="FirstName + LastName + JobTitle";$eligible=$true;$notes="Email did not match, but a single AD user matched name and title."}
        elseif($titleMatches.Count -gt 1){$results.Add((New-Row $r $null "Ambiguous - Multiple Name and Title Matches" "FirstName + LastName + JobTitle" $false "No action" "Skipped" "Multiple AD users matched name and title." $rowNum));continue}
        elseif($nameMatches.Count -eq 1){$user=$nameMatches[0];$status="Potential Match - Title Mismatch";$by="FirstName + LastName only";$eligible=$false;$notes="Single AD user matched first and last name, but title differs."}
        else{$results.Add((New-Row $r $null "Ambiguous - Multiple Name Matches" "FirstName + LastName" $false "No action" "Skipped" "Multiple AD users matched name but no single title match." $rowNum));continue}
    }
    if($CheckSystemMismatch -and $user){
        $employeeIdAlreadyPresent=((Normalize-Text (Get-ObjectValue $user "employeeID")) -eq (Normalize-Text $emp))
        $matchedByWorkerId=($by -eq "WorkerID + AD employeeID")
        $employeeIdGuidance=if($matchedByWorkerId){"No further EmployeeID action needed - matched by WorkerID already present in AD employeeID."}elseif($employeeIdAlreadyPresent){"No further EmployeeID action needed - AD employeeID already matches HR WorkerID."}else{"WorkerID still needs to be applied - system mismatch comparison used a non-WorkerID match."}
        $systemMismatchMatchSources.Add([pscustomobject]@{
            CsvRow=$rowNum
            HR_WorkerID=$emp
            MatchedBy=$by
            MatchStatus=$status
            AD_UserPrincipalName=(Get-ObjectValue $user "userPrincipalName")
            AD_SamAccountName=(Get-ObjectValue $user "sAMAccountName")
            AD_EmployeeID_Current=(Get-ObjectValue $user "employeeID")
            EmployeeIDAlreadyPresent=$employeeIdAlreadyPresent
            EmployeeIDActionGuidance=$employeeIdGuidance
        })
        foreach($mm in @(New-MismatchRows -CsvRow $rowNum -HrRow $r -User $user -Map $map -HrRows $hrRows -SearchBase $SearchBase -MatchedBy $by -MatchStatus $status -EmployeeIDAlreadyPresent $employeeIdAlreadyPresent -EmployeeIDActionGuidance $employeeIdGuidance)){ $mismatches.Add($mm) }
    }
    $action="No action";$actionResult="Dry run"
    if($user -and $eligible -and (Test-IsADUserObject $user)){
        $matchedDn=Normalize-Text (Get-ObjectValue $user "DistinguishedName")
        if($matchedDn -and $confirmedAdUserDns.ContainsKey($matchedDn)){
            $firstMatch=$confirmedAdUserDns[$matchedDn]
            $firstMatchDetail=if($firstMatch -and $firstMatch.PSObject.Properties.Name -contains "CsvRow"){" Earlier match: CSV row $($firstMatch.CsvRow), HR WorkerID $($firstMatch.HR_WorkerID), MatchedBy $($firstMatch.MatchedBy)."}else{""}
            $results.Add((New-Row $r $user "Duplicate - AD User Already Matched" $by $false "No action" "Skipped" "This AD user was already matched to an earlier CSV row.$firstMatchDetail Review manually before approving or applying any employeeID update." $rowNum))
            continue
        }
    }
    if($user){
        if(-not $eligible){$actionResult="Skipped"}
        elseif((Normalize-Text (Get-ObjectValue $user "employeeID")) -eq (Normalize-Text $emp)){$actionResult="Already correct";$notes="$notes EmployeeID already matches."}
        elseif($Apply){ $userDn=Get-ObjectValue $user "DistinguishedName"; if($PSCmdlet.ShouldProcess($userDn,"Set employeeID to $emp")){ Set-ADUser -Identity $userDn -EmployeeID $emp; $action="Update employeeID";$actionResult="Updated" } }
        else{$action="Would update employeeID";$actionResult="Dry run only"}
    }
    if($user -and $eligible -and (Test-IsADUserObject $user)){
        $confirmedDn=Normalize-Text (Get-ObjectValue $user "DistinguishedName")
        if($confirmedDn){
            $confirmedAdUserDns[$confirmedDn]=[pscustomobject]@{
                CsvRow=$rowNum
                HR_WorkerID=$emp
                MatchedBy=$by
                AD_UserPrincipalName=(Get-ObjectValue $user "userPrincipalName")
                AD_SamAccountName=(Get-ObjectValue $user "sAMAccountName")
            }
        }
    }
    $results.Add((New-Row $r $user $status $by $eligible $action $actionResult $notes $rowNum))
}
Write-Progress -Activity $matchingProgressActivity -Completed

$employeeIdApplyResults=if($Apply){@(New-EmployeeIdApplyReportRows -MatchResults $results)}else{@()}
$duplicateAdUserMatches=@($results|Where-Object{$_.MatchStatus -eq "Duplicate - AD User Already Matched"})
$dir=Split-Path $ReportPath -Parent; if($dir -and -not(Test-Path $dir)){New-Item -Path $dir -ItemType Directory -Force | Out-Null}
if($CheckSystemMismatch -and [string]::IsNullOrWhiteSpace($SystemMismatchReportPath)){
    $reportDirectory=Split-Path $ReportPath -Parent
    if([string]::IsNullOrWhiteSpace($reportDirectory)){$reportDirectory="."}
    $reportFileName=[System.IO.Path]::GetFileNameWithoutExtension($ReportPath)
    $reportExtension=[System.IO.Path]::GetExtension($ReportPath)
    if([string]::IsNullOrWhiteSpace($reportExtension)){$reportExtension=".xlsx"}
    $SystemMismatchReportPath=Join-Path $reportDirectory "$reportFileName-System-Mismatches$reportExtension"
}
if($CheckSystemMismatch){
    $mismatchDir=Split-Path $SystemMismatchReportPath -Parent
    if($mismatchDir -and -not(Test-Path $mismatchDir)){New-Item -Path $mismatchDir -ItemType Directory -Force | Out-Null}
}
$xlsx=Ensure-ImportExcelModule -DoNotInstallMissingModules:$DoNotInstallMissingModules
if($xlsx){
    $employeeIdTextColumns=@("HR_$CsvEmployeeIdColumn","HR_WorkerID",$CsvEmployeeIdColumn,"AD_EmployeeID_Current","EmployeeIDBefore","EmployeeIDAfter","User_EmployeeID","HR_$CsvManagerReferenceColumn","HR_ManagerReference",$CsvManagerReferenceColumn)|Select-Object -Unique
    $results | Export-ReportWorksheet -Path $ReportPath -WorksheetName "HR AD Match" -TableName "HRADMatch" -NoNumberConversionColumns $employeeIdTextColumns
    if($duplicateAdUserMatches.Count -gt 0){
        $duplicateAdUserMatches | Export-ReportWorksheet -Path $ReportPath -WorksheetName "Duplicate AD User Matches" -TableName "DuplicateADUserMatches" -NoNumberConversionColumns $employeeIdTextColumns
        Write-Host "Duplicate AD user matches worksheet added to readiness report: $ReportPath" -ForegroundColor Green
    }
    if($Apply){
        $employeeIdApplyResults | Export-ReportWorksheet -Path $ReportPath -WorksheetName "EmployeeID Apply Results" -TableName "EmployeeIDApplyResults" -NoNumberConversionColumns $employeeIdTextColumns
    }
    Write-Host "Excel report written to: $ReportPath" -ForegroundColor Green
    if($CheckSystemMismatch){
        if($mismatches.Count -gt 0){
            $mismatches|Export-ReportWorksheet -Path $SystemMismatchReportPath -WorksheetName "System Mismatches" -TableName "SystemMismatches" -NoNumberConversionColumns $employeeIdTextColumns
        }else{
            New-EmptyMismatchReportRow|Export-ReportWorksheet -Path $SystemMismatchReportPath -WorksheetName "System Mismatches" -TableName "SystemMismatches" -NoNumberConversionColumns $employeeIdTextColumns
        }
        Write-Host "System mismatch report written to: $SystemMismatchReportPath" -ForegroundColor Green
    }
}else{
    $csv=[System.IO.Path]::ChangeExtension($ReportPath,".csv"); $results|Export-Csv $csv -NoTypeInformation -Encoding UTF8; Write-Warning "CSV report written to: $csv"
    if($duplicateAdUserMatches.Count -gt 0){
        $duplicateCsv=[System.IO.Path]::ChangeExtension($ReportPath,".Duplicate-AD-User-Matches.csv")
        $duplicateAdUserMatches|Export-Csv $duplicateCsv -NoTypeInformation -Encoding UTF8
        Write-Warning "Duplicate AD user matches CSV report written to: $duplicateCsv"
    }
    if($Apply){
        $applyCsv=[System.IO.Path]::ChangeExtension($ReportPath,".EmployeeID-Apply-Results.csv")
        $employeeIdApplyResults|Export-Csv $applyCsv -NoTypeInformation -Encoding UTF8
        Write-Warning "EmployeeID apply results CSV report written to: $applyCsv"
    }
    if($CheckSystemMismatch){
        $mismatchCsv=[System.IO.Path]::ChangeExtension($SystemMismatchReportPath,".csv")
        if($mismatches.Count -gt 0){$mismatches|Export-Csv $mismatchCsv -NoTypeInformation -Encoding UTF8}else{New-EmptyMismatchReportRow|Export-Csv $mismatchCsv -NoTypeInformation -Encoding UTF8}
        Write-Warning "System mismatch CSV report written to: $mismatchCsv"
    }
}

Write-Host "`nSummary:" -ForegroundColor Cyan
$results|Group-Object MatchStatus|Select-Object Name,Count|Format-Table -AutoSize
$acceptedReviewDecisions=@((Normalize-Text $ApprovedReviewDecision),(Normalize-Text $IgnoredReviewDecision))
$unapprovedResults=@($results|Where-Object{($acceptedReviewDecisions -notcontains (Normalize-Text $_.ReviewDecision)) -or ($_.MatchStatus -like "ReviewDecision Invalid*")})
Write-Host "`nUnapproved / still requiring review:" -ForegroundColor Cyan
if($unapprovedResults.Count -gt 0){
    $unapprovedResults|Group-Object MatchStatus|Select-Object Name,Count|Format-Table -AutoSize
}else{
    Write-Host "No unapproved rows remain." -ForegroundColor Green
}
if($CheckSystemMismatch){
    Write-Host "`nSystem mismatch match source:" -ForegroundColor Cyan
    if($systemMismatchMatchSources.Count -gt 0){
        $systemMismatchMatchSources|Group-Object EmployeeIDActionGuidance|Select-Object Name,Count|Format-Table -AutoSize
    }else{
        Write-Host "No matched AD users were available for system mismatch comparison." -ForegroundColor Yellow
    }
}
if(-not $Apply){Write-Host "Dry run complete. No AD changes were made." -ForegroundColor Yellow}else{Write-Host "Apply mode complete. Only employeeID updates were attempted against eligible AD users." -ForegroundColor Green}
