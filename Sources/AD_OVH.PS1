<#
.SYNOPSIS
    Generate graphed report for all Active Directory objects, Search and filter AD
.DESCRIPTION
    This Script help to manage and easy search request AD  from HTML page
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

#region Code 

param (
	
	#Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$CompanyLogo = "https://github.com/dakhama-mehdi/AD_OVH/blob/main/Images/AD_OVH.png?raw=true",
	#Logo that will be on the right side, UNC or URL

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$RightLogo = "",
	#Title of generated report

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$ReportTitle = "Active Directory _ Over HTML",
	#Location the report will be saved to

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Temp\")]
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

	[Parameter(ValueFromPipeline = $true, HelpMessage = "AD Objects that have been deleted")]
	$ADModNumber = 5,

    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearcher = 300,
    
    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearchergroups = 100
	
	#CSS template located C:\Program Files\WindowsPowerShell\Modules\PswriteHTML\
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
	
	Write-Host "PSWriteHTML Module is not present, attempting to install it"
	
	Install-Module -Name PSWriteHTML -Force
	Import-Module PSWriteHTML -ErrorAction SilentlyContinue
} else { Import-Module PSWriteHTML}

#Array of default Security Groups, work with all languages

$DefaultSGs = $null
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
$securityeventtable = New-Object 'System.Collections.Generic.List[System.Object]'
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

Write-Host get All users properties
$AllUsers = $null
$AllUsers = Get-ADUser -Filter * -Properties $Alluserpropert -ResultSetSize $maxsearcher

Write-Host get All GPO settings
$GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, ModificationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }

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
	
	'Domain'		    = $Forest
	'AD Recycle Bin'	    = $ADRecycleBin
	'Infrastructure Master'     = $InfrastructureMaster
	'RID Master'		    = $RIDMaster
	'PDC Emulator'		    = $PDCEmulator
	'Domain Naming Master'      = $DomainNamingMaster
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

Write-Host 'get All computers properties on default OU'

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

Write-Host 'get All users properties on default OU'

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

$AllUsers | Where-Object {$_.AccountExpirationDate -lt $dateexpiresoone -and $_.AccountExpirationDate -ne $null -and $_.enabled -eq $true} | foreach-object {
	
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

if (($ExpiringAccountsTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Users were found to expire soon'
	}
	$ExpiringAccountsTable.Add($obj)
}


#Security Logs, this is not improve, you can replace Account with name on your langue, for exemple replace by 'compte' for french version
#We can replace it by event 4771 to list failed kerberos, this will be interesed, or listed 7 users logon on DC by RDP or openlocalsession

Get-EventLog -Newest 7 -LogName "Security" -ComputerName $PDCEmulator | Where-Object { $_.Message -like "*An Account*" }  | ForEach-Object {
	
	$TimeGenerated = $_.TimeGenerated
	$EntryType = $_.EntryType
	$Recipient = $_.Message
	
	$obj = [PSCustomObject]@{
		
		'Time'    = $TimeGenerated
		'Type'    = $EntryType
		'Message' = $Recipient
	}
	
	$SecurityEventTable.Add($obj)
}

if (($SecurityEventTable).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No logon security events were found'
	}
	$SecurityEventTable.Add($obj)
}

#Tenant Domain
 Get-ADForest | Select-Object -ExpandProperty upnsuffixes | ForEach-Object{
	
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
#I'm excluded the Exchange groups
#$Groups = Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups -Properties Member,ManagedBy,ProtectedFromAccidentalDeletion | where {$_.Member -ne $null}
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

#if you are on big company, you can exclude the groups without membership it will be more fast and interessing
Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups -Properties Member,ManagedBy,ProtectedFromAccidentalDeletion  | ForEach-Object {

	$totalgroups ++

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
	} } elseif ($_.GroupCategory -eq "Distribution") {

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
	
	if ($DefaultSGs -contains $Group.Name)
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
	
		if (!($users))
		{
			
			$Groupswithnomembership++
		}
		
		else
		{
			
			$Groupswithmemebrship++
			
		}
	}	
	else
	{
		
		$Users = "Skipped Domain Users Membership"
	}

    $OwnerDN = ($group.ManagedBy -split (",") | ? {$_ -like "CN=*"}) -replace ("CN=","")

	
	$obj = [PSCustomObject]@{
		
		'Name' = $Group.name
		'Type' = $Type
		'Members' = $users
		'Managed By' = $Manager
		'E-mail Address' = $GEmail
		'Protected from Deletion' = $Group.ProtectedFromAccidentalDeletion
		'Default AD Group' = $DefaultADGroup
	}
	
	$table.Add($obj)
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
$OUwithLinked = 0
$OUwithnoLink = 0
$OUProtected = 0
$OUNotProtected = 0

#Get OU, and info, i have chose onelevel to skip loops when the are many OUs, on small company with few OU you can chose subtree on searchscope
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

$UsersWIthPasswordsExpiringInUnderAWeek = 0
$UsersNotLoggedInOver30Days = 0

#Get users that haven't logged on in X amount of days, var is set at start of script
$userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days

 $AllUsers | ForEach-Object {
 
    $totalusers ++

    $lastlog = (LastLogonConvert $_.lastlogon)


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
			#Note if not use the PSO, you can comment this or change by $passwordpol = $null
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
		
		$userphaventloggedonrecentlytable.Add($obj)
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

if (($userphaventloggedonrecentlytable).Count -eq 0)
{
	$userphaventloggedonrecentlytable = [PSCustomObject]@{
		
		Information = "Information: No Users were found to have not logged on in $Days days or more"
	}
}
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
if ($null -ne (($userphaventloggedonrecentlytable).Information))
{
	$UHLONXD = "0"
	
}
Else
{
	$UHLONXD = $userphaventloggedonrecentlytable.Count
	
}
#TOP User table
If ($null -eq (($ExpiringAccountsTable).Information))
{
	
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
		'Expiring Accounts' = $ExpiringAccountsTable.Count
		"Users Haven't Logged on in $Days Days or more" = $UHLONXD
	}
	
	$TOPUserTable.Add($objULic)
	
	
}
Else
{
	
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
		'Expiring Accounts' = "0"
		"Users Haven't Logged on in $Days Days or more" = $UHLONXD
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

foreach ($GPO in $GPOs)
{
	
	$obj = [PSCustomObject]@{
		
		'Name' = $GPO.DisplayName
		'Status' = $GPO.GpoStatus
		'Modified Date' = $GPO.ModificationTime
		'User Version' = $GPO.UserVersion
		'Computer Version' = $GPO.ComputerVersion
	}
	
	$GPOTable.Add($obj)
}
if (($GPOTable).Count -eq 0)
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
'Modified'
'PasswordLastSet'
'DistinguishedName'
)

#$Computers = Get-ADComputer -Filter * -Properties $filtercomputer -ResultSetSize $maxsearcher
$ComputersProtected = 0
$ComputersNotProtected = 0
$ComputerEnabled = 0
$ComputerDisabled = 0
$totalcomputers = 0

#Only search for versions of windows that exist in the Environment

$OSClass = $WindowsRegex = $null
$WindowsRegex = "(Windows (Server )?(\d+|XP)?(\d+|Vista)?( R2)?).*"
$OSClass = @{}

#foreach ($Computer in $Computers)
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

#Fix bug when Win7 showed as 'Windows Embedded Standard'
if (($_.OperatingSystem -match 'Windows Embedded Standard' -or $_.OperatingSystem -like '*7*')) { 

if ($_.OperatingSystemVersion -like '6.1*') {
$_.OperatingSystem = $null
$_.OperatingSystem = 'Windows 7'}
 } 
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Operating System' = $_.OperatingSystem
		'Modified Date' = $_.Modified
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

	'Name'		    = $($_.key)
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

$totalADgroups = (Get-ADGroup -Filter *).count 

$Allobjects  = New-Object 'System.Collections.Generic.List[System.Object]'

$Allobjects = @(
    [pscustomobject]@{Name='Groups';Count=$totalADgroups}
    [pscustomobject]@{Name='Users'; Count=$totalusers}
    [pscustomobject]@{Name='Computers'; Count=$totalcomputers}
    [pscustomobject]@{Name='Contacts'; Count=$totalcontacts}
)

Write-Host "Done!" -ForegroundColor White

#endregion code 

$time = (get-date)

New-HTML -TitleText 'AD_Report-Over-HTML' -FilePath $ReportSavePath -Show  {

    New-HTMLLogo -LeftLogoString $CompanyLogo        
        
    New-HTMLTab -Name 'Dashboard' -IconRegular chart-bar {
    New-HTMLTabStyle  -BackgroundColorActive teal
    
       New-HTMLSection -Name 'Company Information' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                New-htmlTable -HideFooter -HideButtons -DataTable $CompanyInfoTable -DisablePaging -DisableSelect -DisableSearch -DisableStateSave -DisableInfo 
            }
        }    
       New-HTMLSection -Name 'Groups' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            New-HTMLSection -Name 'Domain Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $DomainAdminTable
            }
            New-HTMLSection -Name 'Enterprise Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $EnterpriseAdminTable
            }

        }
       New-HTMLSection -Name 'Objects in Default OUs' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            New-HTMLSection -Name 'Computers' -HeaderBackGroundColor teal  {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $DefaultComputersinDefaultOUTable 
            }
            New-HTMLSection -Name 'Users' -HeaderBackGroundColor teal {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $DefaultUsersinDefaultOUTable
            }

        }   
             
       New-HTMLSection -Name 'AD Objects Deleted in Last 5 Days' -HeaderBackGroundColor teal -HeaderTextAlignment left {
               New-HTMLPanel {
                new-htmlTable -HideFooter -DataTable $ADObjectTable
            }

        }
       New-HTMLSection -Name 'Expiring Items' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            New-HTMLSection -Name "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" -HeaderBackGroundColor teal {
                New-htmlTable -HideFooter -DataTable $PasswordExpireSoonTable 
            }
            New-HTMLSection -Name 'Accounts Expiring Soon' -HeaderBackGroundColor teal {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $ExpiringAccountsTable
            }

        }

       New-HTMLSection -Name 'Accounts' -HeaderBackGroundColor teal -HeaderTextAlignment left  {

            New-HTMLSection -Name "Users Haven't Logged on in $Days Days or more" -HeaderBackGroundColor teal  {
                New-htmlTable -HideFooter -DataTable $userphaventloggedonrecentlytable  
            }
            New-HTMLSection -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor teal {
                New-htmlTable -HideFooter -DataTable $NewCreatedUsersTable
            }

        }
       New-HTMLSection -Name 'Security Logs' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                New-htmlTable -HideFooter -HideButtons -DataTable $securityeventtable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }       
       New-HTMLSection -Name 'UPN Suffixes' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $DomainTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }    
    }

    New-HTMLTab -Name 'Groups' -IconSolid user-alt {
        
       New-HTMLSection -Name 'Groups Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left {
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

            New-HTMLSection -Name "Users Haven't Logged on in $Days Days or more" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                New-HTMLTable -HideFooter -DataTable $userphaventloggedonrecentlytable 
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

    New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       Section -Name 'Users Overivew"' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $GPOTable 
            }
        }

    }

    New-HTMLTab -Name 'Computers' -IconBrands microsoft {
        
       Section -Name 'Computers Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         Panel {
                New-HTMLTable -HideFooter -HideButtons -DataTable $TOPComputersTable
            }
        }
       
         Section -Name 'Computers' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel -Invisible {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -DataTable $ComputersTable   
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
                New-HTMLChart -Title 'Computers Operating Systems' -TitleAlignment center  { 
                    New-ChartTheme -Palette palette5 -Mode light
                    $GraphComputerOS.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }                    
                }
            }
         
        }


    }

    New-HTMLTab -Name 'Resume'  {     

       Section -Name 'Graphes' -HeaderBackGroundColor teal -HeaderTextAlignment left  {

            Section -Name 'Nombres d objets' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $Allobjects
            }
            New-HTMLSection -HeaderText 'All Members' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Pourcent By AD Objects' -TitleAlignment center -Height 300  {
                    New-ChartTheme -Palette palette10 -Mode light
                    
		    $Allobjects.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }


            }

        }

       Section -Name 'About' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
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
    
