
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

function show-adovh {

param (
	
	#Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$CompanyLogo = ".\rre.png",
	#Logo that will be on the right side, UNC or URL

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$RightLogo = ".\rr.png",
	#Title of generated report

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$ReportTitle = "Active Directory Over HTML",
	#Location the report will be saved to

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
	[String]$ReportSavePath = "C:\Temp\test\AD_ovh.html",
	#Find users that have not logged in X Amount of days, this sets the days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have not logged on in more than [X] days. amount of days; Default: 90")]
	$Days = 90,
	#Get users who have been created in X amount of days and less

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have been created within [X] amount of days; Default: 7")]
	$UserCreatedDays = 7,
	#Get users whos passwords expire in less than X amount of days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users password expires within [X] amount of days; Default: 7")]
	$DaysUntilPWExpireINT = 7,
	#Get AD Objects that have been modified in X days and newer

    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearcher = 300,
    
    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on bigg company we can chose a small value like 20 or 200; Default: 10000")]
	$maxsearchergroups = 100
	
	#CSS template located C:\Program Files\WindowsPowerShell\Modules\ReportHTML\1.4.1.1\
	#Default template is orange and named "Sample"
)

#region get infos
function LastLogonConvert ($ftDate)
{
	
	$Date = [DateTime]::FromFileTime($ftDate)
	
	if ($Date -lt (Get-Date '1/1/1900') -or $date -eq 0 -or $date -eq $null)
	{
		
		""
	}
	
	else
	{
		
		$Date
	}
	
} #End function LastLogonConvert

#Check for ReportHTML Module
$Mod = Get-Module -ListAvailable -Name "PSWriteHTML"

if ($null -eq $Mod)
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
$Groupsnomembers = New-Object 'System.Collections.Generic.List[System.Object]'
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
'Description'
'PasswordNeverExpires'
'PasswordNotRequired'
'AccountExpirationDate'
)

Write-Host get All users properties -ForegroundColor Green



Write-Host Get All GPO settings -ForegroundColor Green
$GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, id, ModificationTime, CreationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }
#endregion get infos

#region Dashboard
<###########################
         Dashboard
############################>

Write-Host "Working on Dashboard Report..." -ForegroundColor Green

$dte = (Get-Date).AddDays(-30)
$usercomputerdeleted = 0
$deletedobject = 0

#this function replace whenchanged object, because whenchanged dont return the real reason, if computer is logged value when changed is modified.
#Get deleted objets last admodnumber 

Get-ADObject -Filter {isDeleted -eq $true -and (ObjectClass -eq 'user' -or  ObjectClass -eq 'computer' -or ObjectClass -eq 'group') }  -IncludeDeletedObjects -Properties ObjectClass,whenChanged | ForEach-Object {
	
    if ($_.ObjectClass -eq "GroupPolicyContainer")
	{
		
		$Name = $_.DisplayName
	}
	
	elseif ($_.objectclass -ne "Group")
	{
		$usercomputerdeleted++
        $Name = ($_.Name).split([Environment]::NewLine)[0]
           
	} else {
        $Name = ($_.Name).split([Environment]::NewLine)[0]

}
	if ($_.whenChanged -ge $dte ) 
    {
    $deletedobject++
    }

	$obj = [PSCustomObject]@{
		
		'Name'	      = $Name
		'Object Type' = $_.ObjectClass
		'When Changed' = $_.WhenChanged
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

$ForestObj = Get-ADForest
$DomainControllerobj = Get-ADDomain
$Forest = $DomainControllerobj.Forest
$InfrastructureMaster = ($DomainControllerobj.InfrastructureMaster).Split('.')[0]
$RIDMaster = ($DomainControllerobj.RIDMaster).Split('.')[0]
$PDCEmulator = ($DomainControllerobj.PDCEmulator).Split('.')[0]
$DomainNamingMaster = ($ForestObj.DomainNamingMaster).Split('.')[0]
$SchemaMaster = ($ForestObj.SchemaMaster).Split('.')[0]



#Get Domain Admins
#search domain admins default group and entreprise andministrators
#This is disapriecied because i dont wont to list the sensible informations

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

Get-ADGroupMember -identity "$admdomain" -Recursive -Server $PDCEmulator| ForEach-Object {
	
	$Name = $_.Name
	$Type = $_.ObjectClass
	#$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
	
	$obj = [PSCustomObject]@{
		
		'Name'    = $Name
		#'Enabled' = $Enabled
		'Type'    = $Type
	}
	
	$DomainAdminTable.Add($obj)
}


#Get Enterprise Admins
Get-ADGroupMember -identity "$admentreprise" -Recursive -Server $SchemaMaster | ForEach-Object {

	
	$Name = $_.Name
	$Type = $_.ObjectClass
	#$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
	
	$obj = [PSCustomObject]@{
		
		'Name'    = $Name
		#'Enabled' = $Enabled
		'Type'    = $Type
	}
	
	$EnterpriseAdminTable.Add($obj)
}


$DefaultComputersOU = (Get-ADDomain).computerscontainer
$DefaultComputersinDefaultOU = 0

Write-Host 'get All computer properties on default OU'

Get-ADComputer -Filter * -Properties OperatingSystem,created,PasswordLastSet -SearchBase "$DefaultComputersOU"  | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Operating System' = $_.OperatingSystem
		'Modified Date' = $_.Modified
		'Password Last Set' = $_.PasswordLastSet
	}
	
	$DefaultComputersinDefaultOUTable.Add($obj)
    $DefaultComputersinDefaultOU++
}


$DefaultUsersOU = (Get-ADDomain).UsersContainer 
Get-ADUser -Filter 'enabled -eq $true' -SearchBase $DefaultUsersOU -Properties Name,UserPrincipalName,Enabled,LastLogon | foreach-object {
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'UserPrincipalName' = $_.UserPrincipalName
		'Enabled' = $_.Enabled
		'Last Logon' = (LastLogonConvert $_.lastlogon)
	}
	
	$DefaultUsersinDefaultOUTable.Add($obj)
}

#Security Logs, this is not improve, you can replace Account with name on your langue, for exemple replace by 'compte' for french version
#We can replace it by event 4771 to list failed kerberos, this will be interesed, or listed 7 users logon on DC by RDP or openlocalsession

Write-Host get last locked users

Search-ADAccount -LockedOut -UsersOnly  | ForEach-Object { 
	
	$obj = [PSCustomObject]@{
		
		'name'    = $_.name
		'samaccountname'    = $_.samaccountname
		'lastlogondate ' = $_.lastlogondate 
        'distinguishedname' = $_.distinguishedname
	}
	
	$Unlockusers.Add($obj)
}

#Tenant Domain
$ForestObj | Select-Object -ExpandProperty upnsuffixes | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'UPN Suffixes' = $_
		Valid		   = "True"
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

Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups -Properties Member,ManagedBy,created,ProtectedFromAccidentalDeletion  | ForEach-Object {

$totalgroups++
$OwnerDN = $null

if  (!$_.member) { 

$Groupswithnomembership++
    
    if ($($_.ManagedBy)) {
    $OwnerDN = ($_.ManagedBy -split (",") | ? {$_ -like "CN=*"}) -replace ("CN=","")
    }

	$obj = [PSCustomObject]@{
		
		'Name' = $_.name
		'Type' = $_.GroupCategory
		'Managed By' = $OwnerDN
        'Created' = ($_.created.ToString("yyyy/MM/dd"))
		'Default AD Group' = $DefaultADGroup
        'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
	}
	
	$Groupsnomembers.Add($obj)

 }  
 
 else {

	$Groupswithmemebrship++

	$DefaultADGroup = 'False'
	$Type = New-Object 'System.Collections.Generic.List[System.Object]'


	if ($_.GroupCategory -eq "Security")
	{
		
		$SecurityCount++
        $Type = "Security Group"

	} elseif ($_.GroupCategory -eq "Distribution") {

        $DistroCount++
        $Type = "Distribution Group"
    }    
	
	if ($_.ProtectedFromAccidentalDeletion -eq $True)
	{
		
		$GroupsProtected++
	}
	
	else
	{
		
		$GroupsNotProtected++
	}
	
    if ($DefaultSGs -notcontains $_.Name)
	{
		$CustomGroup++
        $users = ($_.member -split (",") | ? {$_ -like "CN=*"}) -replace ("CN="," ") -join ","

	}
	
	else
	{
		$DefaultADGroup = "True"
		$DefaultGroup++
        $Users = "Skipped Domain Users Membership"

	}


    $OwnerDN = ($_.ManagedBy -split (",") | ? {$_ -like "CN=*"}) -replace ("CN=","")

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.name
		'Type' = $Type
		'Members' = $users
		'Managed By' = $OwnerDN
        'Created' = ($_.created.ToString("yyyy/MM/dd"))
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Default AD Group' = $DefaultADGroup
	}
	
	$table.Add($obj)
}

}


#TOP groups table
$obj1 = [PSCustomObject]@{
	
	'Total Groups' = $totalgroups
	'Groups with members' = $Groupswithmemebrship
	'Security Groups' = $SecurityCount
	'Distribution Groups' = $DistroCount
}

$TOPGroupsTable.Add($obj1)

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
$ExpiringAccountsTable = 0
$PasswordExpireSoonTable = 0

#Get newly created users
$When = ((Get-Date).AddDays(-$UserCreatedDays)).Date

#Get expxired account and still enabled
$dateexpiresoone = (Get-DAte).AddDays(7)
$expiredsoon = 0
$expired = (get-date)

$UsersWIthPasswordsExpiringInUnderAWeek = 0
$UsersNotLoggedInOver30Days = 0
$AccountsExpiringSoon = 0

#Get users that haven't logged on in X amount of days, var is set at start of script
#$userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days

$AllUsers = $null

Get-ADUser -Filter * -Properties $Alluserpropert -ResultSetSize $maxsearcher | ForEach-Object {
 
    $totalusers++

    $lastlog = (LastLogonConvert $_.lastlogon)

    #New created users
    if ( $_.whenCreated -ge $When ) {

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Creation Date' = $_.whenCreated
	}
	
	$NewCreatedUsersTable.Add($obj)
}

    #Expired users and still enabled
    if ($_.AccountExpirationDate -lt $dateexpiresoone -and $_.AccountExpirationDate -ne $null -and $_.enabled -eq $true) {
    	
    if ($_.AccountExpirationDate -gt $expired) { $expiredsoon++ } 

    else {

    $ExpiringAccountsTable++
}
}

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
			
			#$daystoexpire = "User has never logged on"
            $daystoexpire = 000

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
		
        $daystoexpire = 00

	}	


	if (($_.Enabled -eq $True) -and ($lastlog -lt ((Get-Date).AddDays(-$Days))) -and ($_.LastLogon -ne $NULL))
	{
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
	
<#
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
#>
	
	$obj = [PSCustomObject]@{
		
		'Name'				      = $_.Name
		'UserPrincipalName'	      = $_.UserPrincipalName
		'Enabled'				  = $_.Enabled
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Last Logon'			  = $lastlog
        'Last Logon Date'         = $_.LastLogonDate
        'Created'                 = $_.whencreated
        'OU - DN'                 = (($_.DistinguishedName -split (",") | ? {$_ -like "OU=*"}) -replace ("OU=","") -join ",")
		'Email Address'		      = $_.EmailAddress
		'Account Expiration'	  = $_.AccountExpirationDate
		'Change Password Next Logon' = $PasswordExpired
        'Description'             =  $_.description
		'Password Last Set'	      = $_.PasswordLastSet
		'Password Never Expires'  = $_.PasswordNeverExpires
        'Days Until password expired' = $daystoexpire   
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


#TOP User table
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
		'Expiring Accounts' = $ExpiringAccountsTable
	}	
	$TOPUserTable.Add($objULic)

Write-Host "Done!" -ForegroundColor White
#endregion Users

#region GPO
<###########################

	   Group Policy

############################>
Write-Host "Working on Group Policy Report..." -ForegroundColor Green

$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'

#Get GPOs Not Linked
#region gponotlinked
$rootDSE = $adObjects = $linkedGPO = $null
# info: # gpLink est une chaine de caractère de la forme [LDAP://cn={C408C216-5CEE-4EE7-B8BD-386600DC01EA},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={C408C16-5D5E-4EE7-B8BD-386611DC31EA},cn=policies,cn=system,DC=domain,DC=com;0]

[System.Collections.Generic.List[PSObject]]$adObjects = @()
[System.Collections.Generic.List[PSObject]]$linkedGPO = @()

$configuration = ($DomainControllerobj.SubordinateReferences | ? {$_ -like '*configuration*' }).trim()

$domainAndOUS = Get-ADObject -LDAPFilter "(&(|(objectClass=organizationalUnit)(objectClass=domainDNS))(gplink=*))" -SearchBase "$($DomainControllerobj.DistinguishedName)" -Properties gpLink
$sites = Get-ADObject -LDAPFilter "(&(objectClass=site)(gplink=*))" -SearchBase "$configuration" -Properties gpLink

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
$gponotlinked = ([array]($gpos | Where-Object {$_.DisplayName -notin $linkedGPO}) | select DisplayName,CreationTime,GpoStatus)

 
#endregion gponotlinked

if (($GPOs).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Group Policy Obejects were found'
	}
	$GPOTable.Add($obj)
}
Write-Host "Done!" -ForegroundColor White
#endregion GPO

#region Printers
<###########################

		   Printers

############################>

Write-Host "Working on Printers Report..." -ForegroundColor Green
$printersnbr = 0

$printers = Get-AdObject -filter "objectCategory -eq 'printqueue'" -Properties description,drivername,created,location | select name,description,drivername,created,location

$printersnbr = $printers.count

if (!$printers) {

$printers = "No printers Found"

}

#endregion Printers

#region Computers
<###########################

	   Computers

############################>
Write-Host "Working on Computers Report..." -ForegroundColor Green

function getwindowsbuild {
[CmdletBinding()]
param( [string] $OperatingSystemVersion )

switch ( $OperatingSystemVersion )
{
    '10.0 (22621)' {$Build ="2202" }
    '10.0 (19045)' {$Build ="2202" }
    '10.0 (22000)' {$Build ="2102" }
    '10.0 (19044)' {$Build ="2102" }
    '10.0 (19043)' {$Build ="2101" }
    '10.0 (19042)' {$Build ="2002" }
    '10.0 (18362)' {$Build ="1903" }
    '10.0 (17763)' {$Build ="1809" }
    '10.0 (17134)' {$Build ="1803" }
    '10.0 (16299)' {$Build ="1709" }
    '10.0 (15063)' {$Build ="1703" }
    '10.0 (14393)' {$Build ="1607" }
    '10.0 (10586)' {$Build ="1511" }
    '10.0 (10240)' {$Build ="1507" }
    '10.0 (18898)' {$Build ="00"   }
    'Default'      {$Build = "000" }
}

return $Build
}

$filtercomputer = @(
'OperatingSystem'
'OperatingSystemVersion'
'ProtectedFromAccidentalDeletion'
'lastlogondate'
'Created'
'PasswordLastSet'
'DistinguishedName'
'ipv4address'
)

$ComputersProtected = 0
$ComputersNotProtected = 0
$ComputerEnabled = 0
$ComputerDisabled = 0
$totalcomputers = 0
$ComputerNotSupported = 0
$lastcreatedpc = 0
$ComputerProtected = 0
$endofsupportwin = 0
$allwin1011 = 0

#Only search for versions of windows that exist in the Environment

$OSClass = $WindowsRegex = $null
$WindowsRegex = "(Windows (Server )?(\d+|XP)?(\d+|Vista)?( R2)?).*"
$OSClass = @{}

Get-ADComputer -Filter * -Properties $filtercomputer -ResultSetSize $maxsearcher | ? {$_.distinguishedname -notlike '*OU=Domain Controllers*'} |  ForEach-Object {

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

if (($_.OperatingSystem -like '*Windows 10*') –or ($_.OperatingSystem -like 'Windows 11*')) { 

$Winbuild = getwindowsbuild -OperatingSystemVersion $_.OperatingSystemVersion

if ($Winbuild -le '2002') { $endofsupportwin++ }
$allwin1011++
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
        'Build' = $Winbuild
        'IPv4Address' = $_.IPv4Address
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

   $Winbuild  = $null

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


Write-Host "End region Computers Report..." -ForegroundColor Green

#endregion Computers

#region Resume

$Allobjects =  $null

$totalcontacts = (Get-ADObject -Filter 'objectclass -eq "contact"').count

$Allobjects  = New-Object 'System.Collections.Generic.List[System.Object]'


$Allobjects = @(
    [pscustomobject]@{Name='Groups';Count=$totalgroups}
    [pscustomobject]@{Name='Users'; Count=$totalusers}
    [pscustomobject]@{Name='Computers'; Count=$totalcomputers}
    [pscustomobject]@{Name='Contacts'; Count=$totalcontacts}
    [pscustomobject]@{Name='Serveur Printer'; Count=$printersnbr}
)

Write-Host "Done!" -ForegroundColor White

#endregion Resume

#endregion code

$time = (get-date)

#region generatehtml

Write-Host "Working on HTML Report ..." -ForegroundColor Green


New-HTML -TitleText 'AD_OVH' {
   
    New-HTMLNavTop -Logo $CompanyLogo -MenuColorBackground 	gray  -MenuColor Black -HomeColorBackground gray  -HomeLinkHome   {
       
        New-NavTopMenu -Name 'Domains' -IconRegular address-book -IconColor black  {
        New-NavLink -IconSolid users -Name 'Groups' -InternalPageID 'Groups'
        New-NavLink -IconSolid users -Name 'Groups_Empty' -InternalPageID 'Groups_Empty'    
        New-NavLink -IconMaterial folder -Name 'OU' -InternalPageID 'OU'
        New-NavLink -IconSolid scroll -Name 'Group Policy' -InternalPageID 'GPO'
        New-NavLink -IconSolid print -Name 'Printers' -InternalPageID 'Printers'
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
 
      New-HTMLSection  -Name 'Block infos' -Invisible  {

      New-HTMLPanel -Margin 10 -Width "70%" {

      New-HTMLPanel -BackgroundColor silver  {
      New-HTMLText -TextBlock  {
      New-HTMLText -Text  "Domain : $Forest" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "AD Recycle Bin : $ADRecycleBin" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -LineBreak
      New-HTMLText -Text  "Infrastructure : $InfrastructureMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Rid Master : $RIDMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "PDC Emulator : $PDCEmulator" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Domain Naming : $DomainNamingMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Schema Master : $SchemaMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -LineBreak
      
      }
      }

      }

      New-HTMLPanel -Margin 10  {
      
      New-HTMLPanel -BackgroundColor lightgreen -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $UserDisabled -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Disable Users' -Alignment justify -FontSize 15 
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-slash fa-3x" } 
      } 
      }

      New-HTMLPanel -BackgroundColor yellowgreen -AlignContentText right -BorderRadius 10px {
      #New-HTMLTag -Tag 'span' -Attributes @{ class = "fas fa-user-clock fa-3x" } 
      #$Icon = "fas fa-user-clock"
      #New-HTMLTag -Tag 'span' -Attributes @{ class = $Icon; style = @{ 'font-size' = '30px'; 'margin' = '5px'; 'color' = 'green'; 'Alignment' = 'justify' } } 
      New-HTMLText -TextBlock { 
      New-HTMLText -Text $userinactive -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Users not login in Last 90 Days' -Alignment justify -FontSize 15 
      New-HTMLTag -Tag 'span' -Attributes @{ class = "fas fa-user-clock fa-3x" } 
      }
        }
        
         }

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightpink  -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $neverlogedenabled -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Users Never Loged' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-house-user fa-3x" } 
      }
      }

      New-HTMLPanel -BackgroundColor palevioletred  -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $usercomputerdeleted -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Users/computer in RecycleBin' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-trash-alt fa-3x" } 
      }
      }

}

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightblue  -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($DomainAdminTable.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Administrateur du domain' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-edit fa-3x" } 
      }
        }


      New-HTMLPanel -BackgroundColor steelblue -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($EnterpriseAdminTable.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Administrateur d Entreprise' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-tie fa-3x" } 
      }
      }
      

      }     

      New-HTMLPanel -Margin 10  {

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $ComputerNotSupported -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Computer Not Supported' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-laptop-medical fa-3x" } 
      }
      }

      New-HTMLPanel -BackgroundColor orange -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($gponotlinked.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'GPOs not Linked' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-scroll fa-3x" } 
      }
      }

}

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor paleturquoise -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($ExpiringAccountsTable) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Account Expired and Stiil Enabled ' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-umbrella-beach fa-3x" } 
      }
      }
      
      New-HTMLPanel -BackgroundColor mediumaquamarine -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $expiredsoon -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Account Expired Soon' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-bell fa-3x" } 
      }
      }

      }  

        }
         
      New-HTMLSection  -HeaderBackGroundColor teal -HeaderTextAlignment left  {

      New-HTMLSection -Name 'Created Machines / Users By date in last 30 Days' -Invisible  {
      
      New-HTMLPanel  {
                   New-HTMLChart -Title 'Created Machines / Users By date in last 30 Days' -TitleAlignment center -Height 280 {                 
                    New-ChartAxisX -Names $(($barcreateobject).date)
                    New-ChartLine -Name 'User created' -Value $(($barcreateobject).Nbr_users)
                    New-ChartLine -Name 'PC Created' -Value $(($barcreateobject).Nbr_PC)                  
                }
            }    
         }
      
      New-HTMLSection -HeaderBackGroundColor teal -Invisible -Width "70%" {    

      New-HTMLPanel  {

                New-HTMLChart -Title 'Created Objects VS Deleted' -TitleAlignment center -Height "100%" {
                    New-ChartToolbar -Download pan
                    New-ChartBarOptions -Gradient -Vertical
                    New-ChartLegend -Name 'Created users', 'Created Machines', 'Deleted Users/machines' 
                    New-ChartBar -Name 'Result Current 30 Days' -Value $lastcreatedusers, $lastcreatedpc, $deletedobject
                }
            }  
      
      New-HTMLSection -Name 'Objects in Default OU'  -Width "80%"  {
            New-HTMLChart -Gradient  {
                New-ChartLegend -LegendPosition bottom 
                New-ChartDonut -Name 'Users' -Value $DefaultUsersinDefaultOUTable.Count
                New-ChartDonut -Name 'Computers' -Value $DefaultComputersinDefaultOU
            }
        }

            }     

        }
      
   New-HTMLSection -Invisible {

      New-HTMLSection -Name "Last Locked Users" -HeaderBackGroundColor DarkGreen  -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $Unlockusers -HideButtons -DisableSearch
            }

      New-HTMLSection -Name 'UPN Suffix' -HeaderTextAlignment center -HeaderBackGroundColor Black -Width "60%"  {
                New-HTMLTable -DataTable $DomainTable -HideButtons -DisableInfo -DisableSearch -HideFooter -TextWhenNoData 'Information: No UPN Suffixes were found'
      }
    
      New-HTMLSection -Width "60%" -HeaderBackGroundColor Teal -name 'Groups Without members'  {
      
 
      New-HTMLGage -Label 'Empty Groups' -MinValue 0 -MaxValue $totalgroups -Value $Groupswithnomembership -ValueColor Black -LabelColor Black -Pointer
      
            }

      New-HTMLSection -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor DarkBlue {
                new-htmlTable -HideFooter -DataTable $NewCreatedUsersTable -DisableInfo -HideButtons -PagingLength 6 -DisableSearch -TextWhenNoData 'Information: No new users have been recently created'
            }



        }


 New-HTMLSection -Name 'Objects in Default OUs' -Invisible  {

      New-HTMLSection -Name 'AD Objects in Recycle Bin' -HeaderBackGroundColor skyblue -Width "70%" {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-htmlTable -HideFooter -DataTable $ADObjectTable -PagingLength 12 -Buttons csvHtml5 
           } 


      New-HTMLSection -Name 'Computers in default OU' -HeaderBackGroundColor teal   {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-htmlTable -HideFooter -DataTable $DefaultComputersinDefaultOUTable -PagingLength 12 -HideButtons 
            }

      New-HTMLSection -Name 'Users in Default OU' -HeaderBackGroundColor brown {
               New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
               New-HTMLTable -HideFooter -DataTable $DefaultUsersinDefaultOUTable -PagingLength 12 -HideButtons 
            }

        }    
                   
   
          }

    New-HTMLPage -Name 'Groups' {

        New-HTMLTab -Name 'Groups' -IconSolid user-alt   {

       New-HTMLSection -Invisible {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $TOPGroupsTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch 
            }
        }          
          
       New-HTMLSection -Name 'Active Directory Groups With Members' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                new-htmlTable -HideFooter -DataTable $Table -TextWhenNoData 'Information: No Groups were found'
            }
        }
        
       New-HTMLSection -HeaderText 'Active Directory Groups Chart' -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Types' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                     New-ChartPie -Name 'Security Groups' -Value $SecurityCount
                     New-ChartPie -Name 'Distribution Groups' -Value $DistroCount                                    
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Custom vs Default Groups' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name 'Custom Groups' -Value $CustomGroup
                    New-ChartPie -Name 'Default Groups' -Value $DefaultGroup
                  }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Membership' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette3 
                    New-ChartPie -Name 'With Members' -Value $Groupswithmemebrship
                    New-ChartPie -Name 'No Members' -Value $Groupswithnomembership  
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Protected From Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4
                    New-ChartPie -Name 'Not Protected' -Value $GroupsNotProtected
                    New-ChartPie -Name 'Protected' -Value $GroupsProtected                   
                }
            }

        } 
               
       }                
     }
    
    New-HTMLPage -Name 'Groups_Empty' {

       New-HTMLTab -Name 'Groups Without Members' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $Groupsnomembers
            }
        }


    }
    }    

    New-HTMLPage -Name 'OU' {
     
       New-HTMLTab -Name 'Organizational Units' -IconRegular folder {          
          
       New-HTMLSection -Name 'Organizational Units infos' -Invisible {
         New-HTMLPanel {
                new-htmlTable -HideFooter -DataTable $OUTable -TextWhenNoData 'Information: No OUs were found'
            }
        }      
                
       New-HTMLSection -HeaderText "Organizational Units Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'OU Gpos Links' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    New-ChartPie -Name "OUs with GPO's linked" -Value $OUwithLinked
                    New-ChartPie -Name "OUs with no GPO's linked" -Value $OUNotProtected                                      
                }
            }

            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'Organizations Units Protected from deletion' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Protected" -Value $OUProtected
                    New-ChartPie -Name "Not Protected" -Value $OUwithnoLink
                }
            }

        }                

    }

    }

    New-HTMLPage -Name 'GPO' {
        New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $GPOs
            }
        }
       
       New-HTMLSection -Invisible {

       New-HTMLSection -name 'Inlinked Details' -HeaderBackGroundColor Teal {
              New-HTMLTable -DataTable $gponotlinked 
       }

       New-HTMLSection -Name 'Linked Vs Inliked GPOs' -HeaderBackGroundColor Teal  {
            New-HTMLChart {
                New-ChartLegend -LegendPosition bottom 
                New-ChartBarOptions -Gradient
                New-ChartDonut -Name 'Linked' -Value $gponotlinked.Count -Color silver
                New-ChartDonut -Name 'Inlinked' -Value $GPOs.Count -Color orange
            }
        }


    }


    }
    }

    New-HTMLPage -Name 'Printers' {

       New-HTMLTab -Name 'Server Printer' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $printers
            }
        }


    }
    }

    New-HTMLPage -Name 'Users' {

       New-HTMLTab -Name 'Users' -IconSolid audio-description  {
        
       New-HTMLSection -Name 'Users Overivew' -Invisible  {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons  -DataTable $TOPUserTable -DisableSearch
            }
        }
       
       New-HTMLSection -Name 'Active Directory Users' -HeaderBackGroundColor Teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                New-HTMLTable -DataTable $UserTable -DefaultSortColumn Name -HideFooter
            }
        }                
       
       New-HTMLSection -HeaderText "Users Charts" -HeaderBackGroundColor DarkBlue -HeaderTextAlignment left  {
           
            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Enable Vs Disable Users' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette2
                    New-ChartPie -Name "Enabled" -Value $UserEnabled
                    New-ChartPie -Name "Disabled" -Value $UserDisabled                    
                }
            }

             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Password Expiration' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Password Never Expires" -Value $UserPasswordNeverExpires
                    New-ChartPie -Name "Password Expires" -Value $UserPasswordExpires 
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Users Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Protected" -Value $ProtectedUsers
                    New-ChartPie -Name "Not Protected" -Value $NonProtectedUsers 
                }
            }

        }
    }


    }

    New-HTMLPage -Name 'Computers' {
    New-HTMLTab -Name 'Computers' -IconBrands microsoft {
        
       New-HTMLSection -Name 'Computers Overivew' -Invisible  {
         New-HTMLPanel {
                New-HTMLTable -HideFooter -HideButtons -DataTable $TOPComputersTable
            }
        }
       
         New-HTMLSection -Name 'Computers' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel -Invisible {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                New-HTMLTable -DataTable $ComputersTable  
                            }
            }

          New-HTMLSection -HeaderText 'Computers Charts' -HeaderBackGroundColor DarkBlue -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette10 -Mode light
                    New-ChartPie -Name 'Protected' -Value $ComputerProtected
                    New-ChartPie -Name 'Not Protected' -Value $ComputersNotProtected                              
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Enabled Vs Disabled' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4 -Mode light
                    New-ChartPie -Name 'Enabled' -Value $ComputerEnabled
                    New-ChartPie -Name 'Disabled' -Value $ComputerDisabled                  
                }
            }

            }

          New-HTMLSection -Invisible {

         New-HTMLSection -name 'Potentiel Win10 End of support' {
                       
           New-HTMLChart {
                New-ChartLegend -LegendPosition bottom
                New-ChartDonut -Name 'Endofsupport' -Value $endofsupportwin
                New-ChartDonut -Name 'Windows 10/11 supported' -Value ($allwin1011 - $endofsupportwin)
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
    }         
    
    New-HTMLPage -Name 'Resume'  {
    
    New-HTMLTab -Name 'Resume' {     

    New-HTMLSection -Invisible {

      New-HTMLSection  -HeaderBackGroundColor Teal -Invisible  {

      New-HTMLPanel -Margin 10 {
      
      New-HTMLPanel -BackgroundColor lightgreen -AlignContentText right {
      New-HTMLText -Text $Allobjects[0].count -Alignment left -FontSize 40 -FontWeight bold 
      New-HTMLText -Text $Allobjects[0].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-users fa-3x" } 
      }

      New-HTMLText -LineBreak 

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right {
      New-HTMLText -Text $Allobjects[1].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[1].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user fa-3x" } 

        }
        
         }

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightblue  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[2].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[2].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-laptop fa-3x" } 
      }

      New-HTMLText -LineBreak 
      
      New-HTMLPanel -BackgroundColor lightpink  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[3].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[3].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-address-card fa-3x" } 
      }

      }    
      
      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor khaki  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[4].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[4].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-print fa-3x" } 
      }

      New-HTMLText -LineBreak 
      
      }   
        }
      
      New-HTMLSection -HeaderText 'All Members' -Invisible {
     
             New-HTMLPanel -Width "70%" {
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

}

Export-ModuleMember -Function show-adovh
