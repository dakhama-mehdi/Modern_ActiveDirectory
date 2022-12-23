
<#
.SYNOPSIS
    Generate graphed report for all Active Directory objects, Search and filter AD
.DESCRIPTION
    This Script help to manage and easy search request AD  from HTML page
.Requis 
    Script can be executed from Win10/11 or windows server 2012 or more
    required : RSAT Module AD and GPO 
    Required : PSWriteHTML Module
.PARAMETER CompanyLogo
    Enter URL or UNC path to your desired Company Logo for generated report.
    -CompanyLogo "\\Server01\Admin\Files\CompanyLogo.png"
.PARAMETER RightLogo
    Enter URL or UNC path to your desired right-side logo for generated report.
    -RightLogo "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png"
.PARAMETER ReportTitle
    Enter desired title for generated report.
    -ReportTitle "Active Directory _ Over HTML"
.PARAMETER Days
    Users that have not logged in [X] amount of days or more.
    -Days "30"
.PARAMETER UserCreatedDays
    Users that have been created within [X] amount of days.
    -UserCreatedDays "7"
.PARAMETER DaysUntilPWExpireINT
    Users password expires within [X] amount of days
    -DaysUntilPWExpireINT "7"
.PARAMETER ADModNumber
    Active Directory Objects that have been modified within [X] amount of days.
    -ADModNumber "5"  
.PARAMETER maxsearcher
    "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000.
    -$maxsearcher "300"
.PARAMETER maxsearchergroups
    "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200.
    -$maxsearchergroups "100"
.NOTES
    Version: 1.0.3
    Author: Bradley Wyatt
    Date: 12/4/2018
    Modified: JBear 12/5/2018
    Bradley Wyatt 12/8/2018
    jporgand 12/6/2018
    Version: 2.0.0
    Modified: Dakhama Mehdi 
    Date : 08/12/2022
#>

#region code

param (
	
	#Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$CompanyLogo = ".\cg13.png",
	#Logo that will be on the right side, UNC or URL

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$RightLogo = ".\scc.png",
	#Title of generated report

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$ReportTitle = "Active Directory Over HTML - NOVEA",
	#Location the report will be saved to

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
	[String]$ReportSavePath = "C:\Temp\AD_ovh.html",
	#Find users that have not logged in X Amount of days, this sets the days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have not logged on in more than [X] days. amount of days; Default: 30")]
	$Days = 30,
	#Get users who have been created in X amount of days and less

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have been created within [X] amount of days; Default: 7")]
	$UserCreatedDays = 7,
	#Get users whos passwords expire in less than X amount of days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users password expires within [X] amount of days; Default: 7")]
	$DaysUntilPWExpireINT = 7,
	#Get AD Objects that have been modified in X days and newer

	[Parameter(ValueFromPipeline = $true, HelpMessage = "AD Objects that have been modified within [X] amount of days; Default: 3")]
	$ADModNumber = 5,

    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearcher = 300,
    
    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearchergroups = 100
	
	#CSS template located C:\Program Files\WindowsPowerShell\Modules\ReportHTML\1.4.1.1\
	#Default template is orange and named "Sample"
)

function LastLogonConvert ($ftDate)
{
	
	$Date = [DateTime]::FromFileTime($ftDate)
	
	if ($Date -lt (Get-Date '1/1/1900') -or $date -eq 0 -or $date -eq $null)
	{
		
		"Never"
	}
	
	else
	{
		
		$Date
	}
	
} #End function LastLogonConvert

#Check for ReportHTML Module
$Mod = Get-Module -ListAvailable -Name "PSWriteHTML"

If ($null -eq $Mod)
{
	
	Write-Host "ReportHTML Module is not present, attempting to install it"
	
	Install-Module -Name PSWriteHTML -Force
	Import-Module PSWriteHTML -ErrorAction SilentlyContinue
} else { Import-Module PSWriteHTML}

#Array of default Security Groups
$DefaultSGs = $barcreateobject = $null
$DefaultSGs = @()	
$DefaultSGs += ([adsisearcher]"(&(groupType:1.2.840.113556.1.4.803:=1)(!(objectSID=S-1-5-32-546))(!(objectSID=S-1-5-32-545)))").findall().Properties.name
$DefaultSGs += ([adsisearcher] "(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))").FindAll().Properties.name

#region PScutom
$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupTypetable = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultGrouptable = New-Object 'System.Collections.Generic.List[System.Object]'
$EnabledDisabledUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ExpiringAccountsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$Unlockusers = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'
$OUGPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupMembershipTable = New-Object 'System.Collections.Generic.List[System.Object]'
$PasswordExpirationTable = New-Object 'System.Collections.Generic.List[System.Object]'
$PasswordExpireSoonTable = New-Object 'System.Collections.Generic.List[System.Object]'
$userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
$EnterpriseAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$NewCreatedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
$OUProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ADObjectTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ProtectedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputerProtectedTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputersEnabledTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultComputersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultUsersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPGroupsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GraphComputerOS = New-Object 'System.Collections.Generic.List[System.Object]'
$barcreateobject = New-Object 'System.Collections.Generic.List[System.Object]'

#endregion PScustom

#Get all users right away. Instead of doing several lookups, we will use this object to look up all the information needed.
$Alluserpropert = @(
'WhenCreated'
'DistinguishedName'
'ProtectedFromAccidentalDeletion'
'LastLogon'
'EmailAddress'
'LastLogonDate'
'PasswordExpired'
'PasswordLastSet'
'PasswordNeverExpires'
'PasswordNotRequired'
'AccountExpirationDate'
)

Write-Host get All users properties -ForegroundColor Green

$AllUsers = $null
$AllUsers = Get-ADUser -Filter * -Properties $Alluserpropert -ResultSetSize $maxsearcher


Write-Host get All GPO settings
$GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, id, ModificationTime, CreationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }

#region Dashboard
<###########################
         Dashboard
############################>

Write-Host "Working on Dashboard Report..." -ForegroundColor Green

$dte = (Get-Date).AddDays(-$ADModNumber)

#this function replace whenchanged object, because whenchanged dont return the real reason, if computer is logged value when changed is modified.
#Get deleted objets last admodnumber 

Get-ADObject -Filter { whenchanged -gt $dte -and isDeleted -eq $true -and (ObjectClass -eq 'user' -or  ObjectClass -eq 'computer' -or ObjectClass -eq 'group') }  -IncludeDeletedObjects -Properties ObjectClass,whenChanged | ForEach-Object {
	
    if ($_.ObjectClass -eq "GroupPolicyContainer")
	{
		
		$Name = $_.DisplayName
	}
	
	else
	{
		
		$Name = ($_.Name).split([Environment]::NewLine)[0]
	}
	
	$obj = [PSCustomObject]@{
		
		'Name'	      = $Name
		'Object Type' = $_.ObjectClass
		'When Changed' = $_.WhenChanged
	}
	
	$ADObjectTable.Add($obj)
}

if (($ADObjectTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No AD Objects have been deleted recently'
	}
	
	$ADObjectTable.Add($obj)
}


$ADRecycleBinStatus = (Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes

if ($ADRecycleBinStatus.Count -lt 1)
{
	
	$ADRecycleBin = "Disabled"
}
else
{
	
	$ADRecycleBin = "Enabled"
}

#Company Information
$ADInfo = Get-ADDomain
$ForestObj = Get-ADForest
$DomainControllerobj = Get-ADDomain
$Forest = $ADInfo.Forest
$InfrastructureMaster = $DomainControllerobj.InfrastructureMaster
$RIDMaster = $DomainControllerobj.RIDMaster
$PDCEmulator = $DomainControllerobj.PDCEmulator
$DomainNamingMaster = $ForestObj.DomainNamingMaster
$SchemaMaster = $ForestObj.SchemaMaster

$obj = [PSCustomObject]@{
	
	'Domain'			    = $Forest
	'AD Recycle Bin'	    = $ADRecycleBin
	'Infrastructure Master' = $InfrastructureMaster
	'RID Master'		    = $RIDMaster
	'PDC Emulator'		    = $PDCEmulator
	'Domain Naming Master'  = $DomainNamingMaster
	'Schema Master'		    = $SchemaMaster
}

$CompanyInfoTable.Add($obj)

if (($CompanyInfoTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: Could not get items for table'
	}
	$CompanyInfoTable.Add($obj)
}

#Get newly created users
$When = ((Get-Date).AddDays(-$UserCreatedDays)).Date

$AllUsers | Where-Object { $_.whenCreated -ge $When } | ForEach-Object {

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Creation Date' = $_.whenCreated
	}
	
	$NewCreatedUsersTable.Add($obj)
}

if (($NewCreatedUsersTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No new users have been recently created'
	}
	$NewCreatedUsersTable.Add($obj)
}



#Get Domain Admins
#search domain admins default group and entreprise andministrators

([adsisearcher] "(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))").FindAll().Properties | ForEach-Object {

#List group contains admins domain or entreprise or administrator 

 $sidstring = (New-Object System.Security.Principal.SecurityIdentifier($_["objectsid"][0], 0)).Value 

      if ($sidstring -like "*-512" ) {

      $admdomain = $_.name 
      }

      if ( $sidstring -like "*-519" ) {

      $admentreprise = $_.name
      }

      }

Get-ADGroupMember "$admdomain" | ForEach-Object {
	
	$Name = $_.Name
	$Type = $_.ObjectClass
	$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
	
	$obj = [PSCustomObject]@{
		
		'Name'    = $Name
		'Enabled' = $Enabled
		'Type'    = $Type
	}
	
	$DomainAdminTable.Add($obj)
}

if (($DomainAdminTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Domain Admin Members were found'
	}
	$DomainAdminTable.Add($obj)
}


#Get Enterprise Admins
Get-ADGroupMember "$admentreprise" -Server $SchemaMaster | ForEach-Object {

	
	$Name = $_.Name
	$Type = $_.ObjectClass
	$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
	
	$obj = [PSCustomObject]@{
		
		'Name'    = $Name
		'Enabled' = $Enabled
		'Type'    = $Type
	}
	
	$EnterpriseAdminTable.Add($obj)
}

if (($EnterpriseAdminTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: Enterprise Admin members were found'
	}
	$EnterpriseAdminTable.Add($obj)
}

$DefaultComputersOU = (Get-ADDomain).computerscontainer

Write-Host 'get All computer properties on default OU'

Get-ADComputer -Filter * -Properties OperatingSystem,Modified,PasswordLastSet,ProtectedFromAccidentalDeletion -SearchBase "$DefaultComputersOU"  | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Operating System' = $_.OperatingSystem
		'Modified Date' = $_.Modified
		'Password Last Set' = $_.PasswordLastSet
		'Protect from Deletion' = $_.ProtectedFromAccidentalDeletion
	}
	
	$DefaultComputersinDefaultOUTable.Add($obj)
}

if (($DefaultComputersinDefaultOUTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No computers were found in the Default OU'
	}
	$DefaultComputersinDefaultOUTable.Add($obj)
}

$DefaultUsersOU = (Get-ADDomain).UsersContainer 
Get-ADUser -Filter * -SearchBase $DefaultUsersOU -Properties Name,UserPrincipalName,Enabled,ProtectedFromAccidentalDeletion,EmailAddress,DistinguishedName | foreach-object {
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'UserPrincipalName' = $_.UserPrincipalName
		'Enabled' = $_.Enabled
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Last Logon' = (LastLogonConvert $_.lastlogon)
        'Last LogonDate' = ($_.LastLogonDate)
		'Email Address' = $_.EmailAddress
	}
	
	$DefaultUsersinDefaultOUTable.Add($obj)
}
if (($DefaultUsersinDefaultOUTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Users were found in the default OU'
	}
	$DefaultUsersinDefaultOUTable.Add($obj)
}


#Expiring Accounts, this is list all expiring Account and still enabel also expiring user soon 
Write-Host Expiring Accounts and not disabled
$dateexpiresoone = (Get-DAte).AddDays(7)
$expiredsoon = 0
$expired = (get-date)

$AllUsers | Where-Object {$_.AccountExpirationDate -lt $dateexpiresoone -and $_.AccountExpirationDate -ne $null -and $_.enabled -eq $true} | foreach-object {
	
    if ($_.AccountExpirationDate -gt $expired) { $expiredsoon++ } 
    else {

	$NameLoose = $_.Name
	$UPNLoose = $_.UserPrincipalName
	$ExpirationDate = $_.AccountExpirationDate
	$enabled = $_.Enabled
	
	$obj = [PSCustomObject]@{
		
		'Name'			    = $NameLoose
		'UserPrincipalName' = $UPNLoose
		'Expiration Date'   = $ExpirationDate
		'Enabled'		    = $enabled
	}
	
	$ExpiringAccountsTable.Add($obj)
}
}

if (($ExpiringAccountsTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Users were found to expire soon'
	}
	$ExpiringAccountsTable.Add($obj)
}

#Security Logs, this is not improve, you can replace Account with name on your langue, for exemple replace by 'compte' for french version
#We can replace it by event 4771 to list failed kerberos, this will be interesed, or listed 7 users logon on DC by RDP or openlocalsession

#Get-EventLog -Newest 7 -LogName "Security" -ComputerName $PDCEmulator | Where-Object { $_.Message -like "*Un compte*" }  | ForEach-Object {
Search-ADAccount -LockedOut -UsersOnly  | ForEach-Object { 

	
	$obj = [PSCustomObject]@{
		
		'name'    = $_.name
		'samaccountname'    = $_.samaccountname
		'lastlogondate ' = $_.lastlogondate 
        'distinguishedname' = $_.distinguishedname
	}
	
	$Unlockusers.Add($obj)
}

if (($Unlockusers).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No logon security events were found'
	}
	$Unlockusers.Add($obj)
}

#Tenant Domain
 Get-ADForest | Select-Object -ExpandProperty upnsuffixes | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'UPN Suffixes' = $_
		Valid		   = "True"
	}
	
	$DomainTable.Add($obj)
}
if (($DomainTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No UPN Suffixes were found'
	}
	$DomainTable.Add($obj)
}

Write-Host "Done!" -ForegroundColor White
#endregion Dashboard

#region groups
<###########################

		   Groups

############################>

Write-Host "Working on Groups Report..." -ForegroundColor Green

#Get groups and sort in alphabetical order
#list only group with members, this can be interresed on big domain with a lot of groups, you can remove the where if you are in small company
#I'm excluded the Exchange groups -ResultSetSize $maxsearchergroups
#$Groups = Get-ADGroup -Filter "name -notlike '*Exchange*'"  -Properties Member,ManagedBy,ProtectedFromAccidentalDeletion | where {$_.Member -ne $null}
$SecurityCount = 0
$MailSecurityCount = 0
$CustomGroup = 0
$DefaultGroup = 0
$Groupswithmemebrship = 0
$Groupswithnomembership = 0
$GroupsProtected = 0
$GroupsNotProtected = 0
$totalgroups = 0
$DistroCount = 0 

#Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups  -Properties Member,ManagedBy,ProtectedFromAccidentalDeletion | where {$_.Member -ne $null} | ForEach-Object {
Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups  -Properties Member,ManagedBy,ProtectedFromAccidentalDeletion  | ForEach-Object {

$totalgroups++

if  (!$_.member) { 

$Groupswithnomembership++

 }  
 
 else {

	$Groupswithmemebrship++

	$DefaultADGroup = 'False'
	$Type = New-Object 'System.Collections.Generic.List[System.Object]'
	#$Gemail = (Get-ADGroup $Group -Properties mail).mail
    $Gemail = $null

	if (($_.GroupCategory -eq "Security") -and ($Gemail -ne $Null))
	{
		
		$MailSecurityCount++
	}
	
	if (($_.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
	{
		
		$SecurityCount++
	} elseif ($_.GroupCategory -eq "Distribution") {

        $DistroCount++
    }    
	
	if ($_.ProtectedFromAccidentalDeletion -eq $True)
	{
		
		$GroupsProtected++
	}
	
	else
	{
		
		$GroupsNotProtected++
	}
	
	if ($DefaultSGs -contains $_.Name)
	{
		
		$DefaultADGroup = "True"
		$DefaultGroup++
	}
	
	else
	{
		
		$CustomGroup++
	}
	
	if ($_.GroupCategory -eq "Distribution")
	{
		
		$Type = "Distribution Group"
	}
	
	if (($_.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
	{
		
		$Type = "Security Group"
	}
	
	if (($_.GroupCategory -eq "Security") -and (($Gemail) -ne $Null))
	{
		
		$Type = "Mail-Enabled Security Group"
	}


	if ($_.Name -ne $admdomain)
	{
      
        $users = ($_.member -split (",") | ? {$_ -like "CN=*"}) -replace ("CN="," ") -join ","

	}	

	else
	{
		
		$Users = "Skipped Domain Users Membership"
	}


    $OwnerDN = ($_.ManagedBy -split (",") | ? {$_ -like "CN=*"}) -replace ("CN=","")

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.name
		'Type' = $Type
		'Members' = $users
		'Managed By' = $OwnerDN
		#'E-mail Address' = $GEmail
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Default AD Group' = $DefaultADGroup
	}
	
	$table.Add($obj)
}

}

if (($table).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Groups were found'
	}
	$table.Add($obj)
}

#TOP groups table
$obj1 = [PSCustomObject]@{
	
	'Total Groups' = $totalgroups
	'Mail-Enabled Security Groups' = $MailSecurityCount
	'Security Groups' = $SecurityCount
	'Distribution Groups' = $DistroCount
}

$TOPGroupsTable.Add($obj1)

$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Mail-Enabled Security Groups'
	'Count' = $MailSecurityCount
}

$GroupTypetable.Add($obj1)

$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Security Groups'
	'Count' = $SecurityCount
}

$GroupTypetable.Add($obj1)

$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Distribution Groups'
	'Count' = $DistroCount
}

$GroupTypetable.Add($obj1)

#Default Group Pie Chart
$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Default Groups'
	'Count' = $DefaultGroup
}

$DefaultGrouptable.Add($obj1)

$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Custom Groups'
	'Count' = $CustomGroup
}

$DefaultGrouptable.Add($obj1)

#Group Protection Pie Chart
$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Protected'
	'Count' = $GroupsProtected
}

$GroupProtectionTable.Add($obj1)

$obj1 = [PSCustomObject]@{
	
	'Name'  = 'Not Protected'
	'Count' = $GroupsNotProtected
}

$GroupProtectionTable.Add($obj1)

#Groups with membership vs no membership pie chart
$objmem = [PSCustomObject]@{
	
	'Name'  = 'With Members'
	'Count' = $Groupswithmemebrship
}

$GroupMembershipTable.Add($objmem)

$objmem = [PSCustomObject]@{
	
	'Name'  = 'No Members'
	'Count' = $Groupswithnomembership
}

$GroupMembershipTable.Add($objmem)

Write-Host "Done!" -ForegroundColor White
#endregion groups

#region OU
<###########################

    Organizational Units

############################>

Write-Host "Working on Organizational Units Report..." -ForegroundColor Green

#Get all OUs'
#$OUs = Get-ADOrganizationalUnit -Filter * -Properties ProtectedFromAccidentalDeletion -ResultSetSize 1
$OUwithLinked = 0
$OUwithnoLink = 0
$OUProtected = 0
$OUNotProtected = 0

#foreach ($OU in $OUs)
Get-ADOrganizationalUnit -Filter * -Properties ProtectedFromAccidentalDeletion -SearchScope OneLevel | ForEach-Object {
	
	$LinkedGPOs = New-Object 'System.Collections.Generic.List[System.Object]'
	
	if (($_.linkedgrouppolicyobjects).length -lt 1)
	{
		
		$LinkedGPOs = "None"
		$OUwithnoLink++
	}
	
	else
	{
		
		$OUwithLinked++
		$GPOslinks = $_.linkedgrouppolicyobjects
		
		foreach ($GPOlink in $GPOslinks)
		{
			
			$Split1 = $GPOlink -split "{" | Select-Object -Last 1
			$Split2 = $Split1 -split "}" | Select-Object -First 1
			$LinkedGPOs.Add((Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName)
		}
	}

	if ($_.ProtectedFromAccidentalDeletion -eq $True)
	{
		
		$OUProtected++
	}
	
	else
	{
		
		$OUNotProtected++
	}
	
	$LinkedGPOs = $LinkedGPOs -join ", "
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Linked GPOs' = $LinkedGPOs
		'Modified Date' = $_.WhenChanged
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
	}
	
	$OUTable.Add($obj)
}

if (($OUTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No OUs were found'
	}
	$OUTable.Add($obj)
}

#OUs with no GPO Linked
$obj1 = [PSCustomObject]@{
	
	'Name'  = "OUs with no GPO's linked"
	'Count' = $OUwithnoLink
}

$OUGPOTable.Add($obj1)

$obj2 = [PSCustomObject]@{
	
	'Name'  = "OUs with GPO's linked"
	'Count' = $OUwithLinked
}

$OUGPOTable.Add($obj2)

#OUs Protected Pie Chart
$obj1 = [PSCustomObject]@{
	
	'Name'  = "Protected"
	'Count' = $OUProtected
}

$OUProtectionTable.Add($obj1)

$obj2 = [PSCustomObject]@{
	
	'Name'  = "Not Protected"
	'Count' = $OUNotProtected
}

$OUProtectionTable.Add($obj2)

Write-Host "Done!" -ForegroundColor White
#endregion OU

#region Users
<###########################

           USERS

############################>

Write-Host "Working on Users Report..." -ForegroundColor Green

$UserEnabled = 0
$UserDisabled = 0
$UserPasswordExpires = 0
$UserPasswordNeverExpires = 0
$ProtectedUsers = 0
$NonProtectedUsers = 0
$totalusers = 0
$Userinactive = 0
$Createdlastdays = ((Get-Date).AddDays(-30)).Date
$lastcreatedusers = 0
$neverlogedenabled = 0

$UsersWIthPasswordsExpiringInUnderAWeek = 0
$UsersNotLoggedInOver30Days = 0
$AccountsExpiringSoon = 0

#Get users that haven't logged on in X amount of days, var is set at start of script
#$userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days

 $AllUsers | ForEach-Object {
 
    $totalusers ++

    $lastlog = (LastLogonConvert $_.lastlogon)


    #User Never loged and Enabled
    if (($_.lastlogondate -eq $null) -and ($_.Enabled -ne $false)) 
    {
     $neverlogedenabled++
    }

    #get days until password expired
	if ((($_.PasswordNeverExpires) -eq $False) -and (($_.Enabled) -ne $false))
	{
		
		#Get Password last set date
		$passwordSetDate = ($_.PasswordLastSet)
		
		if ($null -eq $passwordSetDate)
		{
			
			$daystoexpire = "User has never logged on"
		}
		
		else
		{
			
			#Check for Fine Grained Passwords
			$PasswordPol = (Get-ADUserResultantPasswordPolicy $_)
			
			if (($PasswordPol) -ne $null)
			{
				
				$maxPasswordAgePSO = ($PasswordPol).MaxPasswordAge
                $expireson = $passwordsetdate.AddDays($maxPasswordAgePSO.days)
                $maxPasswordAgePSO = $null
			} else {
		
			$expireson = $passwordsetdate.AddDays($maxPasswordAge)
            }

			$today = (Get-Date)
			
			#Gets the count on how many days until the password expires and stores it in the $daystoexpire var
			$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
		}
	}
	
	else
	{
		
		$daystoexpire = "N/A"
	}
	


	if (($_.Enabled -eq $True) -and ($lastlog -lt ((Get-Date).AddDays(-$Days))) -and ($_.LastLogon -ne $NULL))
	{
	    <#
		$obj = [PSCustomObject]@{
			
			'Name' = $_.Name
			'UserPrincipalName' = $_.UserPrincipalName
			'Enabled' = $_.Enabled
			'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
			'Last Logon' = $lastlog
            'Last LongonDate' = $_.LastLogonDate
			'Password Never Expires' = $_.PasswordNeverExpires
			'Days Until Password Expires' = $daystoexpire
		}
		#>

		#$userphaventloggedonrecentlytable.Add($obj)
        $userinactive++
	}

    #Get User created Last 30 days
    if ( $_.whenCreated -ge $createdlastdays ) 
    {

    $lastcreatedusers++

    $barcreated = ($_.Whencreated.ToString("yyyy/MM/dd"))

    $rec=$barcreateobject | where {$_.date -eq $barcreated } 

    if ($rec) {

            $rec.Nbr_users += 1
    } else {

       $obj = [PSCustomObject]@{
		
		'Nbr_users' = 1
        'Nbr_PC' = 0
		'Date' = $barcreated
	}

$barcreateobject.Add($obj)

}

    }

	
	#Items for protected vs non protected users
	if ($_.ProtectedFromAccidentalDeletion -eq $False)
	{
		
		$NonProtectedUsers++
	}
	
	else
	{
		
		$ProtectedUsers++
	}
	
	#Items for the enabled vs disabled users pie chart
	if (($_.PasswordNeverExpires) -ne $false)
	{
		
		$UserPasswordNeverExpires++
	}
	
	else
	{
		
		$UserPasswordExpires++
	}
	
	#Items for password expiration pie chart
	if (($_.Enabled) -ne $false)
	{
		
		$UserEnabled++
	}
	
	else
	{
		
		$UserDisabled++
	}
	
	$Name = $_.Name
	$UPN = $_.UserPrincipalName
	$Enabled = $_.Enabled
	$EmailAddress = $_.EmailAddress
    $LastLogon = $lastlog
    $LastLogonDate = $_.LastLogonDate
    $Created = $_.whencreated
    $OU_DN     = (($_.DistinguishedName -split (",") | ? {$_ -like "OU=*"}) -replace ("OU=","") -join ",")
	$AccountExpiration = $_.AccountExpirationDate
	$PasswordExpired = $_.PasswordExpired
	$PasswordLastSet = $_.PasswordLastSet
	$PasswordNeverExpires = $_.PasswordNeverExpires
	$daysUntilPWExpire = $daystoexpire
	
	$obj = [PSCustomObject]@{
		
		'Name'				      = $Name
		'UserPrincipalName'	      = $UPN
		'Enabled'				  = $Enabled
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Last Logon'			  = $LastLogon
        'Last Logon Date'         = $_.LastLogonDate
        'Created'                 = $Created
        'OU - DN'                 = $OU_DN
		'Email Address'		      = $EmailAddress
		'Account Expiration'	  = $AccountExpiration
		'Change Password Next Logon' = $PasswordExpired
		'Password Last Set'	      = $PasswordLastSet
		'Password Never Expires'  = $PasswordNeverExpires
		'Days Until Password Expires' = $daystoexpire
	}
	
	$usertable.Add($obj)
	
	if ($daystoexpire -lt $DaysUntilPWExpireINT -and $daystoexpire -ge 0)
	{
		
		$obj = [PSCustomObject]@{
			
			'Name'					      = $Name
			'Days Until Password Expires' = $daystoexpire
		}
		
		$PasswordExpireSoonTable.Add($obj)
	}
}

<#
if (($userphaventloggedonrecentlytable).Count -eq 0)
{
	$userphaventloggedonrecentlytable = [PSCustomObject]@{
		
		Information = "Information: No Users were found to have not logged on in $Days days or more"
	}
}

#>

if (($PasswordExpireSoonTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No users were found to have passwords expiring soon'
	}
	$PasswordExpireSoonTable.Add($obj)
}


if (($usertable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No users were found'
	}
	$usertable.Add($obj)
}

#Data for users enabled vs disabled pie graph
$objULic = [PSCustomObject]@{
	
	'Name'  = 'Enabled'
	'Count' = $UserEnabled
}

$EnabledDisabledUsersTable.Add($objULic)

$objULic = [PSCustomObject]@{
	
	'Name'  = 'Disabled'
	'Count' = $UserDisabled
}

$EnabledDisabledUsersTable.Add($objULic)

#Data for users password expires pie graph
$objULic = [PSCustomObject]@{
	
	'Name'  = 'Password Expires'
	'Count' = $UserPasswordExpires
}

$PasswordExpirationTable.Add($objULic)

$objULic = [PSCustomObject]@{
	
	'Name'  = 'Password Never Expires'
	'Count' = $UserPasswordNeverExpires
}

$PasswordExpirationTable.Add($objULic)

#Data for protected users pie graph
$objULic = [PSCustomObject]@{
	
	'Name'  = 'Protected'
	'Count' = $ProtectedUsers
}

$ProtectedUsersTable.Add($objULic)

$objULic = [PSCustomObject]@{
	
	'Name'  = 'Not Protected'
	'Count' = $NonProtectedUsers
}

$ProtectedUsersTable.Add($objULic)

<#
if ($null -ne (($userphaventloggedonrecentlytable).Information))
{
	$UHLONXD = "0"
	
}
Else
{
	$UHLONXD = $userphaventloggedonrecentlytable.Count
	
}
#>

#TOP User table
If ($null -eq (($ExpiringAccountsTable).Information))
{
	
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
		'Expiring Accounts' = $ExpiringAccountsTable.Count
		"Locked users" = $UHLONXD
	}
	
	$TOPUserTable.Add($objULic)
	
	
}
Else
{
	
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
		'Expiring Accounts' = "0"
		"Lock" = $UHLONXD
	}
	$TOPUserTable.Add($objULic)
}

Write-Host "Done!" -ForegroundColor White
#endregion Users

#region GPO
<###########################

	   Group Policy

############################>
Write-Host "Working on Group Policy Report..." -ForegroundColor Green

$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
<#
foreach ($GPO in $GPOs)
{
	
	$obj = [PSCustomObject]@{
		
		'Name' = $GPO.DisplayName
		'Status' = $GPO.GpoStatus
		'Created Date' = $GPO.CreationTime
		'User Version' = $GPO.UserVersion
		'Computer Version' = $GPO.ComputerVersion
	}
	
	$GPOTable.Add($obj)
}#>

#Get GPOs Not Linked
#region gponotlinked
$rootDSE = $adObjects = $linkedGPO = $null
# info: # gpLink est une chaine de caractère de la forme [LDAP://cn={C408C216-5CEE-4EE7-B8BD-386600DC01EA},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={C408C16-5D5E-4EE7-B8BD-386611DC31EA},cn=policies,cn=system,DC=domain,DC=com;0]

[System.Collections.Generic.List[PSObject]]$adObjects = @()
[System.Collections.Generic.List[PSObject]]$linkedGPO = @()

$rootDSE = Get-ADRootDSE

$domainAndOUS = Get-ADObject -LDAPFilter "(&(|(objectClass=organizationalUnit)(objectClass=domainDNS))(gplink=*))" -SearchBase "$($rootDSE.defaultNamingcontext)" -Properties gpLink
$sites = Get-ADObject -LDAPFilter "(&(objectClass=site)(gplink=*))" -SearchBase "$($rootDSE.configurationNamingContext)" -Properties gpLink

# construit une liste avec tous les gpLink existants
$adObjects.Add($domainAndOUS)
$adObjects.Add($sites)

    # Compare si GUID de la GPO existe dans les gpLink
    # gpLink est une chaine de caractère de la forme [LDAP://cn={C408C216-5CEE-4EE7-B8BD-386600DC01EA},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={C408C16-5D5E-4EE7-B8BD-386611DC31EA},cn=policies,cn=system,DC=domain,DC=com;0]
    
    $GPOs | ForEach-Object {
    
    if($adObjects.gpLink -match $_.id) {
        $linkedGPO.Add($_.DisplayName)
    }
}

# cast en array pour prendre en considération le cas où un seul object
$gponotlinked = ([array]($gpos | Where-Object {$_.DisplayName -notin $linkedGPO}) | select DisplayName,CreationTime,GposStatus)

 
#endrgion gponotlinked

if (($GPOs).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Group Policy Obejects were found'
	}
	$GPOTable.Add($obj)
}
Write-Host "Done!" -ForegroundColor White
#endregion GPO

#region Computers
<###########################

	   Computers

############################>
Write-Host "Working on Computers Report..." -ForegroundColor Green

$filtercomputer = @(
'OperatingSystem'
'OperatingSystemVersion'
'ProtectedFromAccidentalDeletion'
'lastlogondate'
'Created'
'PasswordLastSet'
'DistinguishedName'
)

$ComputersProtected = 0
$ComputersNotProtected = 0
$ComputerEnabled = 0
$ComputerDisabled = 0
$totalcomputers = 0
$ComputerNotSupported = 0
$lastcreatedpc = 0

#Only search for versions of windows that exist in the Environment

$OSClass = $WindowsRegex = $null
$WindowsRegex = "(Windows (Server )?(\d+|XP)?(\d+|Vista)?( R2)?).*"
$OSClass = @{}

Get-ADComputer -Filter * -Properties $filtercomputer -ResultSetSize $maxsearcher | ForEach-Object {

	$totalcomputers ++
	if ($_.ProtectedFromAccidentalDeletion -eq $True)
	{
		$ComputersProtected++
	} else 	{
		
		$ComputersNotProtected++
	}
	
	if ($_.Enabled -eq $True)
	{		
		$ComputerEnabled++
	} else 	{
		
		$ComputerDisabled++
	}

    #Computer Created last 30 jours 
    if ( $_.Created -ge $createdlastdays ) 
    {

    $lastcreatedpc++

    $barcreated = ($_.created.ToString("yyyy/MM/dd"))

    $rec=$barcreateobject | where {$_.date -eq $barcreated } 

    if ($rec) {

            $rec.Nbr_PC += 1
    } else {

       $obj = [PSCustomObject]@{
		
		'Nbr_users' = 0
        'Nbr_PC' = 1
		'Date' = $barcreated
	}

$barcreateobject.Add($obj)

}

    }


if (($_.OperatingSystem -match 'Windows Embedded Standard' -or $_.OperatingSystem -like '*7*')) { 

if ($_.OperatingSystemVersion -like '6.1*') {
$_.OperatingSystem = $null
$_.OperatingSystem = 'Windows 7'}
 } 
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Operating System' = $_.OperatingSystem
		'Created Date' = $_.Created
        'OU _ Patch'      = (($_.DistinguishedName -split (",") | ? {$_ -like "OU=*"}) -replace ("OU=","") -join ",")
		'Password Last Set' = $_.PasswordLastSet
        'Last Logon Date'   = $_.LastLogonDate
		'Protect from Deletion' = $_.ProtectedFromAccidentalDeletion
	}
	
	$ComputersTable.Add($obj)
	
    
    if ($_.OperatingSystem -match 'Windows 7'){
    $OSClass['Windows 7'] += 'Windows 7'.Count
    } elseif ($_.OperatingSystem -match $WindowsRegex ){ 
        $OSClass[$matches[1]] += $matches[1].Count
    } elseif ($null -ne $_.OperatingSystem) {
        $OSClass[$_.OperatingSystem] += $_.OperatingSystem.Count
    }   
    
    #Get unsuported Machines
   
   switch ($_.OperatingSystem ) {

    {$_ -match 2003 -or $_ -match 2008 -or $_ -match 2000 -or $_ -match 7 -or $_ -match 'Vista' -or $_ -match 'XP' } {$ComputerNotSupported++}
}

}

if (($ComputersTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No computers were found'
	}
	$ComputersTable.Add($obj)
}

#region Pie chart breaking down OS for computer obj
$GraphComputerOS =  $null
$GraphComputerOS = New-Object 'System.Collections.Generic.List[System.Object]'

$OSClass.GetEnumerator() | ForEach-Object {

$hashcomputer = [PSCustomObject]@{

	'Name'			    = $($_.key)
	'Count'	            = $($_.value)
}

$GraphComputerOS.Add($hashcomputer)

}
#endregion Pie chart

#Data for TOP Computers data table

$OSClass.Add("Total Computers",$totalcomputers)

$TOPComputersTable = [pscustomobject]$OSClass

#Data for protected Computers pie graph
$objULic = [PSCustomObject]@{
	
	'Name'  = 'Protected'
	'Count' = $ComputerProtected
}

$ComputerProtectedTable.Add($objULic)

$objULic = [PSCustomObject]@{
	
	'Name'  = 'Not Protected'
	'Count' = $ComputersNotProtected
}

$ComputerProtectedTable.Add($objULic)

#Data for enabled/vs Computers pie graph
$objULic = [PSCustomObject]@{
	
	'Name'  = 'Enabled'
	'Count' = $ComputerEnabled
}

$ComputersEnabledTable.Add($objULic)

$objULic = [PSCustomObject]@{
	
	'Name'  = 'Disabled'
	'Count' = $ComputerDisabled
}

$ComputersEnabledTable.Add($objULic)
#endregion Computers

$Allobjects =  $null

$totalcontacts = (Get-ADObject -Filter 'objectclass -eq "contact"').count

#$totalADgroups = (Get-ADGroup -Filter *).count 

$Allobjects  = New-Object 'System.Collections.Generic.List[System.Object]'

$Allobjects = @(
    [pscustomobject]@{Name='Groups';Count=$totalgroups}
    [pscustomobject]@{Name='Users'; Count=$totalusers}
    [pscustomobject]@{Name='Computers'; Count=$totalcomputers}
    [pscustomobject]@{Name='Contacts'; Count=$totalcontacts}
)

Write-Host "Done!" -ForegroundColor White

#endregion code 

$time = (get-date)

#region generatehtml

New-HTML -TitleText 'AD_OVH' {
   
    New-HTMLNavTop -Logo $CompanyLogo -MenuColorBackground 	gray  -MenuColor Black -HomeColorBackground gray  -HomeLinkHome   {
       
        New-NavTopMenu -Name 'Domains' -IconRegular address-book -IconColor black  {
        New-NavLink -IconSolid users -Name 'Groups' -InternalPageID 'Groups'
        New-NavLink -IconMaterial folder -Name 'OU' -InternalPageID 'OU'
        New-NavLink -IconSolid scroll -Name 'Group Policy' -InternalPageID 'GPO'
        }

        New-NavTopMenu -Name 'Objects' -IconSolid sitemap {
            New-NavLink -IconSolid user-tie -Name 'Users' -InternalPageID 'Users'
            New-NavLink -IconSolid laptop -Name 'Computers' -InternalPageID 'Computers'
        }

        New-NavTopMenu -Name 'About' -IconRegular chart-bar {
            New-NavLink -IconSolid chart-pie -Name 'Resume' -InternalPageID 'Resume'
        }
    } 
   
    New-HTMLTab -Name 'Dashboard' -IconRegular chart-bar  {
   
    New-HTMLTabStyle  -BackgroundColorActive teal
        
      New-HTMLSection -Invisible {
    
      New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $CompanyInfoTable -DisablePaging -DisableSelect -DisableSearch -DisableStateSave -DisableInfo 
            }
        }    

      New-HTMLSection -HeaderText 'Company Information' -HeaderTextAlignment center -HeaderTextColor WhiteSmoke -HeaderBackGroundColor DarkBlue  { 

        New-HTMLSection -Invisible  {
        New-HTMLToast  -Text "Disable Users  : $UserDisabled " -IconSolid user-slash -TextColor CoralRed -BarColorLeft CoralRed -IconColor CoralRed 
        }

        New-HTMLSection -Invisible {
        New-HTMLToast  -Text "Users not login in Last 90 Days : $userinactive" -IconSolid user-clock -TextColor CarrotOrange -BarColorLeft CarrotOrange -IconColor CarrotOrange
        }

        New-HTMLSection -Invisible  {
        New-HTMLToast -Text "Users Never Loged : $neverlogedenabled" -IconSolid house-user
        }

        New-HTMLSection -Invisible {
        New-HTMLToast  -Text "Administrateur du domain : $($DomainAdminTable.count)" -IconSolid user-edit -TextColor green -BarColorLeft green -IconColor green
        }

        New-HTMLSection -Invisible {
        New-HTMLToast -Text "Administrateur d'Entreprise : $($EnterpriseAdminTable.count)" -IconSolid user-tie
        }
               
     }
     
      New-HTMLSection -Invisible  { 

      New-HTMLSection -Invisible  {
        New-HTMLToast  -Text "Users/computer in RecycleBin : $($ADObjectTable.Count)" -IconSolid trash-alt -TextColor Teal -BarColorLeft Teal -IconColor Teal
        }

      New-HTMLSection -Invisible {
        New-HTMLToast  -Text "Computer Not Supported : $ComputerNotSupported" -IconSolid laptop-medical -TextColor CarrotOrange -BarColorLeft Brown -IconColor Brown
        }

      New-HTMLSection -Invisible  -Width "60%" {
      New-HTMLGage -Label 'Empty Groups' -MinValue 0 -MaxValue $totalgroups -Value $Groupswithnomembership -ValueColor Black -LabelColor Black -Pointer
        }

      New-HTMLSection -Invisible  {

       New-HTMLToast  -Text "Account Expired and Stiil Enabled : $($ExpiringAccountsTable.count)" -IconSolid umbrella-beach -BarColorLeft gold -IconColor gold
        }

      New-HTMLSection -Invisible  {
      New-HTMLToast  -Text "GPOs not Linked : $($gponotlinked.count) " -IconSolid scroll -TextColor deeppink -BarColorLeft deeppink -IconColor deeppink
        }
      
     }
         
      New-HTMLSection  -HeaderBackGroundColor teal -HeaderTextAlignment left  {

      New-HTMLSection -Name 'Created Machines / Users By date in last 30 Days' -Invisible  {
      
      New-HTMLPanel  {
                   New-HTMLChart -Title 'Created Machines / Users By date in last 30 Days' -TitleAlignment center -Height 280 {                 
                    New-ChartAxisX -Names ($barcreateobject).date
                    New-ChartLine -Name 'User created' -Value ($barcreateobject).Nbr_users
                    New-ChartLine -Name 'PC Created' -Value ($barcreateobject).Nbr_PC                    
                }
            }    
         }
      
      New-HTMLSection -HeaderBackGroundColor teal -Invisible -Width "70%" {    

      New-HTMLPanel  {

                New-HTMLChart -Title 'Created Objects VS Deleted' -TitleAlignment center -Height "100%" {
                    New-ChartToolbar -Download pan
                    New-ChartBarOptions -Gradient -Vertical
                    New-ChartLegend -Name 'Created users', 'Created Machines', 'Deleted Users/machines' 
                    New-ChartBar -Name 'Result Current 30 Days' -Value $lastcreatedusers, $lastcreatedpc, '5'
                }
            }  
      
      New-HTMLSection -Name 'Objects in Default OU'  -Width "80%"  {
            New-HTMLChart {
                New-ChartLegend -LegendPosition bottom
                New-ChartDonut -Name 'Users' -Value $DefaultUsersinDefaultOUTable.Count
                New-ChartDonut -Name 'Computers' -Value $DefaultComputersinDefaultOUTable.Count
            }
        }

            }     

        }
      
 New-HTMLSection -Invisible {

      New-HTMLSection -Name "Last Locked Users" -HeaderBackGroundColor teal  -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $Unlockusers -Buttons pdfHtml5
            }
    
      New-HTMLSection -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $NewCreatedUsersTable -DisableInfo -DisableSearch 
            }

     New-HTMLSection -Width "70%" -HeaderBackGroundColor Teal -Invisible {
      
      New-HTMLPanel -BackgroundColor Orange  -AlignContentText center  {
      New-HTMLText -LineBreak
      New-HTMLText -LineBreak
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-bell fa-5x" } 
      New-HTMLText -Text "Account Expired Soon " -Alignment center -FontSize 14
      New-HTMLText -Text $expiredsoon -FontSize 14
      }


      New-HTMLPanel -BackgroundColor Cyan  -AlignContentText center {
      New-HTMLText -LineBreak
      New-HTMLText -LineBreak
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-exclamation fa-5x" }
      New-HTMLText -Text "Password Expired Soon ." -FontSize 14
      New-HTMLText -Text $PasswordExpireSoonTable.Count -FontSize 14
        }
      }


        }


 New-HTMLSection -Name 'Objects in Default OUs'  {

      New-HTMLSection -Name 'AD Objects in Recycle Bin' -HeaderBackGroundColor teal  {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' 
                New-htmlTable -HideFooter -DataTable $ADObjectTable -PagingLength 10
           } 


      New-HTMLSection -Name 'Domain Administrators' -HeaderBackGroundColor teal -Width "50%"  {
                new-htmlTable -HideFooter -DataTable $DomainAdminTable -PagingLength 10 -DisableInfo -Buttons pdfHtml5
            }

      New-HTMLSection -Name 'Enterprise Administrators' -HeaderBackGroundColor teal -Width "50%" {
                new-htmlTable -HideFooter -DataTable $EnterpriseAdminTable -PagingLength 10 -DisableInfo -Buttons pdfHtml5
            }

        }    
                   
   
          }
              
    New-HTMLPage -Name 'Groups' {
        New-HTMLTab -Name 'Groups' -IconSolid user-alt   {

       New-HTMLSection -Name 'Groups Overivew' -HeaderBackGroundColor Teal -HeaderTextAlignment left {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $TOPGroupsTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }          
          
       New-HTMLSection -Name 'Active Directory Groups' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                new-htmlTable -HideFooter -DataTable $Table
            }
        }
        
       New-HTMLSection -Name 'Objects in Default OUs' -HeaderBackGroundColor teal -HeaderTextAlignment left {
            New-HTMLSection -Name 'Domain Administrators' -HeaderBackGroundColor teal  {
                new-htmlTable -HideFooter -DataTable $DomainAdminTable 
                
            }
            New-HTMLSection -Name 'Enterprise Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $EnterpriseAdminTable
            }
}                  

       New-HTMLSection -HeaderText 'Active Directory Groups Chart' -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Types' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    $GroupTypetable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Custom vs Default Groups' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    $DefaultGrouptable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Membership' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette3 
                    $GroupMembershipTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Protected From Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4
                    $GroupProtectionTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }                
        }
    }

    New-HTMLPage -Name 'OU' {
        New-HTMLTab -Name 'Organizational Units' -IconRegular folder {          
          
       New-HTMLSection -Name 'Organizational Units infos' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                new-htmlTable -HideFooter -DataTable $OUTable
            }
        }
      
                
       New-HTMLSection -HeaderText "Organizational Units Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'OU Gpos Links' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    $OUGPOTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'Organizations Units Protected from deletion' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    $OUProtectionTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }                

    }

    }

    New-HTMLPage -Name 'GPO' {
        New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       New-HTMLSection -Name 'Informations"' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                New-HTMLTable  -DataTable $GPOs 
            }
        }
       
       New-HTMLSection -Name 'Information' -Invisible {

       New-HTMLSection -Name 'GPOs Not Linked Details' {
              New-HTMLTable -DataTable $gponotlinked -PagingLength 10
       }

       New-HTMLSection -Name 'Linked Vs Inliked GPOs' -Width "50%"  {
            New-HTMLChart -Gradient {
                New-ChartLegend -LegendPosition bottom
                New-ChartDonut -Name 'Inlinked' -Value $gponotlinked.Count -Color Silver
                New-ChartDonut -Name 'linked' -Value $GPOs.Count -Color CarrotOrange
            }
        }
    }
    }
    }

    New-HTMLPage -Name 'Users' {

        New-HTMLTab -Name 'Users' -IconSolid audio-description  {
        
       New-HTMLSection -Name 'Users Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons  -DataTable $TOPUserTable -DisableSearch
            }
        }
       
       New-HTMLSection -Name 'Active Directory Users' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -DataTable $UserTable -DefaultSortColumn Name -HideFooter
            }
        }        
        
       New-HTMLSection -Name 'Expiring Items' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            New-HTMLSection -Name "Users Locked" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                New-HTMLTable -HideFooter -DataTable $Unlockusers
            }
            New-HTMLSection -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                New-HTMLTable -HideFooter -DataTable $NewCreatedUsersTable
            }

        }

       New-HTMLSection -Name 'Accounts' -HeaderBackGroundColor teal -HeaderTextAlignment left {

       New-HTMLSection -Name "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $PasswordExpireSoonTable
            }
       New-HTMLSection -Name "Accounts Expiring Soon" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $ExpiringAccountsTable
            }

        }

       New-HTMLSection -HeaderText "Users Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Enable Vs Disable Users' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette2
                    $EnabledDisabledUsersTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Password Expiration' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    $PasswordExpirationTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Users Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    $ProtectedUsersTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }
    }


    }

    New-HTMLPage -Name 'Computers' {

        New-HTMLTab -Name 'Computers' -IconBrands microsoft {
        
       New-HTMLSection -Name 'Computers Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                New-HTMLTable -HideFooter -HideButtons -DataTable $TOPComputersTable
            }
        }
       
         New-HTMLSection -Name 'Computers' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel -Invisible {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -DataTable $ComputersTable  
                #New-HTMLTab  -DataTable $ComputersTable -DateTimeSortingFormat 'yyyy-MM-dd' -HideFooter 
                            }
            }

          New-HTMLSection -HeaderText 'Computers Charts' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette10 -Mode light
                    $ComputerProtectedTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Enabled Vs Disabled' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4 -Mode light
                    $ComputersEnabledTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            }

         New-HTMLSection -HeaderText 'Computers Operating System Breakdown' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
                New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Operating Systems' -TitleAlignment center  { 
                    New-ChartTheme  -Mode light
                    $GraphComputerOS.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }                    
                }
            }
         
        }


    }


    }

    New-HTMLPage -Name 'Resume'  {
    
    New-HTMLTab -Name 'Resume' {     

       New-HTMLSection -Name 'Graphes' -Invisible {

            New-HTMLSection -Name 'Nombres d objets' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $Allobjects
            }
            New-HTMLSection -HeaderText 'All Members' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Pourcent By AD Objects' -TitleAlignment center -Height 300  {
                    New-ChartTheme  -Mode light                    
		    $Allobjects.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }


            }

        }

       New-HTMLSection -Name 'About' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
         New-HTMLList {
              New-HTMLListItem -Text 'Resume All objects AD' 
              New-HTMLListItem -Text "Generated date $time"
              New-HTMLListItem -Text 'Active Directory _ OverHTML  Version : 2.0  Author Dakhama Mehdi - Date : 08/12/2022<br> 
              <br> Inspired ADReportHTLM Version : 1.0.3 Author: Bradley Wyatt - Date: 12/4/2018 [thelazyadministrato](https://www.thelazyadministrator.com/)<br>
              <br> Thanks : JBear,jporgand<br>
              <br> Credit : Mahmoud Hatira, Zouhair sarouti<br>
              <br> Thanks : Boss PrzemyslawKlys - Module PSWriteHTML- [Evotec](https://evotec.xyz) '
              } -FontSize 12
            }
            
          New-HTMLPanel {
            New-HTMLImage -Source $RightLogo 
        } 
        }   
    }

    }    

    
} -ShowHTML -Online -FilePath $ReportSavePath

#endregion generatehtml
