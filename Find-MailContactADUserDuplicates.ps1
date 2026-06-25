<#
.SYNOPSIS
    Find potential duplicate GAL entries where a MailContact and an AD User / MailUser appear to represent the same person.

.DESCRIPTION
    Report-only script. Checks MailContacts against AD users using email-style values first, then first name, last name and title.
    No AD objects are modified.
#>

[CmdletBinding()]
param(
    [string]$ReportPath = ".\MailContact-ADUser-Duplicate-Report.xlsx",
    [string]$SearchBase,
    [switch]$DoNotInstallMissingModules,
    [switch]$IncludeTitleMismatches,
    [switch]$EnabledUsersOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-Text { param([object]$Value) if ($null -eq $Value) { return "" }; return ($Value.ToString().Trim() -replace "\s+", " ").ToLowerInvariant() }
function Remove-SmtpPrefix { param([object]$Value) if ($null -eq $Value) { return "" }; return (($Value.ToString()) -replace "^smtp:", "" -replace "^SMTP:", "") }
function Escape-Ldap { param([string]$Value) if ($null -eq $Value) { return "" }; $v=$Value; $v=$v -replace "\\","\5c"; $v=$v -replace "\*","\2a"; $v=$v -replace "\(","\28"; $v=$v -replace "\)","\29"; $v=$v -replace "`0","\00"; return $v }

function Ensure-ImportExcelModule {
    param([switch]$DoNotInstallMissingModules)
    if (Get-Module -ListAvailable -Name ImportExcel) { Import-Module ImportExcel -ErrorAction Stop; return $true }
    if ($DoNotInstallMissingModules) { Write-Warning "ImportExcel not installed; CSV fallback will be used."; return $false }
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop }
        $repo=Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if($repo -and $repo.InstallationPolicy -ne "Trusted"){Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop}
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        return $true
    } catch { Write-Warning "Could not install ImportExcel. CSV fallback will be used. Error: $($_.Exception.Message)"; return $false }
}

function Get-ADObjectSafe {
    param([string]$LdapFilter,[string]$SearchBase)
    $props=@("objectClass","mail","proxyAddresses","targetAddress","mailNickname","givenName","sn","displayName","title","employeeID","userPrincipalName","sAMAccountName","enabled","distinguishedName","whenCreated","whenChanged")
    if([string]::IsNullOrWhiteSpace($SearchBase)){return @(Get-ADObject -LDAPFilter $LdapFilter -Properties $props)}
    return @(Get-ADObject -LDAPFilter $LdapFilter -SearchBase $SearchBase -Properties $props)
}

function Get-ObjectValue { param([object]$Object,[string]$Name) if($Object.PSObject.Properties.Name -notcontains $Name){return $null}; return $Object.PSObject.Properties[$Name].Value }

function Get-AllComparableEmails {
    param([object]$Object)
    $items=New-Object System.Collections.Generic.List[string]
    foreach($v in @((Get-ObjectValue $Object "mail"),(Get-ObjectValue $Object "userPrincipalName"),(Remove-SmtpPrefix (Get-ObjectValue $Object "targetAddress")))){ $n=Normalize-Text $v; if($n){$items.Add($n)} }
    foreach($p in @((Get-ObjectValue $Object "proxyAddresses"))){ $n=Normalize-Text (Remove-SmtpPrefix $p); if($n){$items.Add($n)} }
    return @($items|Sort-Object -Unique)
}

function Get-PrimaryComparableEmail {
    param([object]$Object)
    $all=@(Get-AllComparableEmails $Object)
    if($all.Count -gt 0){return $all[0]}
    return ""
}

function Test-EmailOverlap {
    param([object]$A,[object]$B)
    $aEmails=@(Get-AllComparableEmails $A); $bEmails=@(Get-AllComparableEmails $B)
    foreach($e in $aEmails){if($bEmails -contains $e){return $true}}
    return $false
}

function Get-RecipientType {
    param([object]$Object)
    $classes=@($Object.objectClass)
    if($classes -contains "contact"){return "MailContact"}
    if(($classes -contains "user") -and -not [string]::IsNullOrWhiteSpace($Object.targetAddress)){return "MailUser"}
    if($classes -contains "user"){return "ADUser"}
    return "Unknown"
}

function Find-UserObjectsByEmail {
    param([string]$Email,[string]$SearchBase)
    if([string]::IsNullOrWhiteSpace($Email)){return @()}
    $e=Escape-Ldap $Email; $smtp1=Escape-Ldap "smtp:$Email"; $smtp2=Escape-Ldap "SMTP:$Email"
    $enabled=if($EnabledUsersOnly){"(!(userAccountControl:1.2.840.113556.1.4.803:=2))"}else{""}
    $filter="(&(objectCategory=person)(objectClass=user)$enabled(|(mail=$e)(userPrincipalName=$e)(proxyAddresses=$smtp1)(proxyAddresses=$smtp2)(targetAddress=$smtp1)(targetAddress=$smtp2)(targetAddress=$e)))"
    return @(Get-ADObjectSafe -LdapFilter $filter -SearchBase $SearchBase)
}

function Find-UserObjectsByName {
    param([string]$FirstName,[string]$LastName,[string]$SearchBase)
    if([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)){return @()}
    $f=Escape-Ldap $FirstName; $l=Escape-Ldap $LastName
    $enabled=if($EnabledUsersOnly){"(!(userAccountControl:1.2.840.113556.1.4.803:=2))"}else{""}
    $filter="(&(objectCategory=person)(objectClass=user)$enabled(givenName=$f)(sn=$l))"
    return @(Get-ADObjectSafe -LdapFilter $filter -SearchBase $SearchBase)
}

function New-DuplicateReportRow {
    param([object]$Contact,[object]$User,[string]$MatchStatus,[string]$MatchedBy,[string]$Notes)
    [pscustomobject]@{
        MatchStatus=$MatchStatus; MatchedBy=$MatchedBy; Notes=$Notes
        EmailMatches=(Test-EmailOverlap $Contact $User)
        FirstNameMatches=((Normalize-Text $Contact.givenName) -eq (Normalize-Text $User.givenName))
        LastNameMatches=((Normalize-Text $Contact.sn) -eq (Normalize-Text $User.sn))
        TitleMatches=((Normalize-Text $Contact.title) -eq (Normalize-Text $User.title))
        Contact_ObjectType=(Get-RecipientType $Contact); Contact_DisplayName=$Contact.displayName; Contact_FirstName=$Contact.givenName; Contact_LastName=$Contact.sn; Contact_Title=$Contact.title; Contact_Mail=$Contact.mail; Contact_TargetAddress=$Contact.targetAddress; Contact_PrimaryComparableEmail=(Get-PrimaryComparableEmail $Contact); Contact_ProxyAddresses=(@($Contact.proxyAddresses)-join "; "); Contact_MailNickname=$Contact.mailNickname; Contact_DistinguishedName=$Contact.DistinguishedName; Contact_WhenCreated=$Contact.whenCreated; Contact_WhenChanged=$Contact.whenChanged
        User_ObjectType=(Get-RecipientType $User); User_SamAccountName=$User.sAMAccountName; User_UserPrincipalName=$User.userPrincipalName; User_DisplayName=$User.displayName; User_FirstName=$User.givenName; User_LastName=$User.sn; User_Title=$User.title; User_Mail=$User.mail; User_TargetAddress=$User.targetAddress; User_PrimaryComparableEmail=(Get-PrimaryComparableEmail $User); User_ProxyAddresses=(@($User.proxyAddresses)-join "; "); User_MailNickname=$User.mailNickname; User_EmployeeID=$User.employeeID; User_Enabled=$User.Enabled; User_DistinguishedName=$User.DistinguishedName; User_WhenCreated=$User.whenCreated; User_WhenChanged=$User.whenChanged
    }
}

function Add-UniqueResult {
    param([System.Collections.Generic.List[object]]$Results,[hashtable]$SeenPairs,[object]$Contact,[object]$User,[string]$MatchStatus,[string]$MatchedBy,[string]$Notes)
    if($null -eq $Results){throw "Internal error: Results collection was not supplied to Add-UniqueResult."}
    $key="$($Contact.DistinguishedName)|$($User.DistinguishedName)"
    if($SeenPairs.ContainsKey($key)){return}
    $SeenPairs[$key]=$true
    $Results.Add((New-DuplicateReportRow -Contact $Contact -User $User -MatchStatus $MatchStatus -MatchedBy $MatchedBy -Notes $Notes))
}

Write-Host "Starting MailContact to AD User duplicate audit..."
Import-Module ActiveDirectory -ErrorAction Stop
$filter="(&(objectClass=contact)(|(mail=*)(proxyAddresses=*)(targetAddress=*)(givenName=*)(sn=*)))"
$mailContacts=@(Get-ADObjectSafe -LdapFilter $filter -SearchBase $SearchBase)
Write-Host "MailContacts collected: $($mailContacts.Count)" -ForegroundColor Cyan

$results=New-Object System.Collections.Generic.List[object]
$seen=@{}
$i=0
foreach($contact in $mailContacts){
    $i++
    Write-Progress -Activity "Checking MailContacts against AD Users" -Status "MailContact $i of $($mailContacts.Count): $($contact.displayName)" -PercentComplete (($i/[Math]::Max($mailContacts.Count,1))*100)
    foreach($email in @(Get-AllComparableEmails $contact)){
        $emailMatches=@(Find-UserObjectsByEmail -Email $email -SearchBase $SearchBase)
        foreach($user in $emailMatches){Add-UniqueResult -Results $results -SeenPairs $seen -Contact $contact -User $user -MatchStatus "Likely Duplicate" -MatchedBy "Email" -Notes "MailContact and AD user share at least one comparable email value. Review whether the MailContact is still required."}
    }
    if(-not [string]::IsNullOrWhiteSpace($contact.givenName) -and -not [string]::IsNullOrWhiteSpace($contact.sn)){
        $nameMatches=@(Find-UserObjectsByName -FirstName $contact.givenName -LastName $contact.sn -SearchBase $SearchBase)
        $titleMatches=@($nameMatches|Where-Object{(Normalize-Text $_.title) -eq (Normalize-Text $contact.title)})
        foreach($user in $titleMatches){Add-UniqueResult -Results $results -SeenPairs $seen -Contact $contact -User $user -MatchStatus "Likely Duplicate" -MatchedBy "FirstName + LastName + Title" -Notes "MailContact and AD user matched on first name, last name, and title. Email values may differ. Review for duplicate GAL entry."}
        if($IncludeTitleMismatches){
            foreach($user in @($nameMatches|Where-Object{(Normalize-Text $_.title) -ne (Normalize-Text $contact.title)})){Add-UniqueResult -Results $results -SeenPairs $seen -Contact $contact -User $user -MatchStatus "Potential Duplicate - Title Mismatch" -MatchedBy "FirstName + LastName only" -Notes "MailContact and AD user matched on first name and last name, but title differs. Manual review required."}
        }
    }
}
Write-Progress -Activity "Checking MailContacts against AD Users" -Completed

$summary=@(
    [pscustomobject]@{Metric="MailContacts checked";Value=$mailContacts.Count},
    [pscustomobject]@{Metric="Duplicate candidate rows";Value=$results.Count},
    [pscustomobject]@{Metric="Unique MailContacts with candidate duplicate";Value=@($results|Select-Object -ExpandProperty Contact_DistinguishedName -Unique).Count},
    [pscustomobject]@{Metric="Unique AD users with candidate duplicate";Value=@($results|Select-Object -ExpandProperty User_DistinguishedName -Unique).Count},
    [pscustomobject]@{Metric="EnabledUsersOnly";Value=$EnabledUsersOnly},
    [pscustomobject]@{Metric="IncludeTitleMismatches";Value=$IncludeTitleMismatches}
)

$dir=Split-Path $ReportPath -Parent; if($dir -and -not(Test-Path $dir)){New-Item -Path $dir -ItemType Directory -Force|Out-Null}
$xlsx=Ensure-ImportExcelModule -DoNotInstallMissingModules:$DoNotInstallMissingModules
if($xlsx){
    if($results.Count -gt 0){$results|Export-Excel -Path $ReportPath -WorksheetName "Duplicate Candidates" -TableName "DuplicateCandidates" -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet}else{[pscustomobject]@{Result="No duplicate candidates found."}|Export-Excel -Path $ReportPath -WorksheetName "Duplicate Candidates" -TableName "DuplicateCandidates" -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet}
    $summary|Export-Excel -Path $ReportPath -WorksheetName "Summary" -TableName "Summary" -AutoSize -FreezeTopRow -BoldTopRow -Append
    Write-Host "Excel report written to: $ReportPath" -ForegroundColor Green
}else{
    $csv=[System.IO.Path]::ChangeExtension($ReportPath,".csv")
    if($results.Count -gt 0){$results|Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8}else{[pscustomobject]@{Result="No duplicate candidates found."}|Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8}
    $summaryCsv=[System.IO.Path]::ChangeExtension($ReportPath,".Summary.csv")
    $summary|Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    Write-Warning "CSV report written to: $csv"
}

Write-Host "`nSummary:" -ForegroundColor Cyan
$summary|Format-Table -AutoSize
if($results.Count -gt 0){Write-Host "`nCandidate breakdown:" -ForegroundColor Cyan; $results|Group-Object MatchStatus,MatchedBy|Select-Object Name,Count|Format-Table -AutoSize}
Write-Host "Audit complete. No AD changes were made." -ForegroundColor Yellow
