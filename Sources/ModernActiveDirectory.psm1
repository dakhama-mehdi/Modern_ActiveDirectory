function Get-ADModernReport{

<#
    .SYNOPSIS
    New Experience to Manage Active Directory over interactive HTML.

    .DESCRIPTION
    This Module help to create a Dynamic Web Report to manage Active Directory.

    .EXAMPLE
    Create a sample report multipages for test, note by default only 300 objects will be listed.
    Get-ADModernReport 
    .EXAMPLE
    Create a report for illimited objects
    Get-ADModernReport -illimitedsearch
    .EXAMPLE
    Create onepage report and save in specific folder, We can change the name output file like Mycompany.html
    Get-ADModernReport -SavePath C:\myfolder\ -htmlonepage
    .EXAMPLE
    Create report with companylogo and limited listed groups to 3000 and object to 5000
    Get-ADModernReport -CompanyLogo C:\myfolder\ADWeb.PNG -maxsearchergroups 3000 -maxsearcher 5000

    .NOTES
    #>


 [CmdletBinding()]
param (
    #Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$CompanyLogo = "https://raw.githubusercontent.com/dakhama-mehdi/Modern_ActiveDirectory/main/Pictures/SmallLogo.png",

    #Logo that will be on the right side, UNC or URL
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$RightLogo = "https://raw.githubusercontent.com/dakhama-mehdi/Modern_ActiveDirectory/main/Pictures/Rightlogo-1.png",

    #Title of generated report
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$ReportTitle = "Active Directory Over HTML",

    #Location the report will be saved	
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
	[String]$SavePath = $env:TEMP,

   	#Find users that have not logged in X Amount of days, this sets the days	
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have not logged on in more than [X] days. amount of days; Default: 90")]
	$Days = 90,

	#Get users who have been created in X amount of days and less	
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have been created within [X] amount of days; Default: 7")]
	$UserCreatedDays = 7,

	#Get users whos passwords expire in less than X amount of days
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users password expires within [X] amount of days; Default: 7")]
	$DaysUntilPWExpireINT = 7,

    #MAX AD Objects to search, for quick test on big company we can chose a small value like 20 or 200
    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on big company we can chose a small value like 20 or 200; Default: 300")]
	$maxsearcher = 300,
    
    #MAX AD Objects to search, for quick test on big company we can chose a small value like 20 or 200.
    [Parameter(ValueFromPipeline = $true, HelpMessage = "MAX AD Objects to search, for quick test on big company we can chose a small value like 20 or 200; Default: 300")]
	$maxsearchergroups = 300,

    #OU Level Searching, Will be slow if there are a lot of OU; Default: Onelevel"
    [Parameter(ValueFromPipeline = $true, HelpMessage = "OU Level Searching, Will be slow if there are a lot of OU; Default: Onelevel")]
	[ValidateSet("Onelevel","Base","Subtree")]$OUlevelSearch = 'Onelevel',

    #Search on all AD objects whitout limit, can be long on Big Company.
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Illimited search max value")]
    [switch]$illimitedsearch,

    #Show admins members with admincount 1.
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Show admins members")]
    [switch]$Showadmin,

    #Generate One page, not recommanded for company with more 5K objects.
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Onelinepage")]
    [switch]$htmlonepage
)

if ($illimitedsearch.IsPresent) {
     $maxsearcher = 300000
     $maxsearchergroups = 100000
 }

#region get infos

function LastLogonConvert ($ftDate)
{	
	$Date = [DateTime]::FromFileTime($ftDate)	
	if ($Date -lt (Get-Date '1/1/1900') -or $date -eq 0 -or $null -eq $date)
	{		
		""
	} else {
    $Date  
    }
	
} #End function LastLogonConvert

#Check for Active Directory Module 
if (!(Get-Module -ListAvailable -Name "ActiveDirectory")) {

throw "AD RSAT Module is required, pls install it, operation aborted. to install module run : Install-WindowsFeature RSAT-AD-PowerShell" 

}

#Check for ReportHTML Module

if (!(Get-Module -ListAvailable -Name "PSWriteHTML"))
{	
	Write-Host "ReportHTML Module is not present, attempting to install it" -ForegroundColor Red	
	Install-Module -Name PSWriteHTML -Force
	Import-Module PSWriteHTML -ErrorAction SilentlyContinue

} else { Import-Module PSWriteHTML -ErrorAction Stop }

#Array of default Security Groups
$DefaultSGs = $barcreateobject = $null
$DefaultSGs = @()	
$DefaultSGs = ([adsisearcher]"(|(&(groupType:1.2.840.113556.1.4.803:=1))(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*)))").FindAll().Properties.name
#region PScutom
$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$Groupsnomembers = New-Object 'System.Collections.Generic.List[System.Object]'
$OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ExpiringAccountsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$Unlockusers = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ADObjectTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultComputersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultUsersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPGroupsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GraphComputerOS = New-Object 'System.Collections.Generic.List[System.Object]'
$barcreateobject = New-Object 'System.Collections.Generic.List[System.Object]'
$NewCreatedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
#endregion PScustom

#Check for GPMC module
$GPMOD = Get-Module -ListAvailable -Name "GroupPolicy"
if ($null -eq $GPMOD)
{	
	Write-Host "GPMC Feature is not present, Pls install it to get info. To install module : Install-windowsfeature GPMC" -ForegroundColor Red
    $nogpomod = $true
} else { 
Write-Host Get All GPO settings -ForegroundColor Green
$GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, id, ModificationTime, CreationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }
}

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
$EnterpriseAdminTable = $DomainAdminTable = 0

#Get Domain Admins and Entreprise admins
#search domain admins default group and entreprise andministrators
#This is disapriecied because i dont wont to list the sensible informations

([adsisearcher] "(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))").FindAll().Properties | ForEach-Object {

#List group contains admins domain or entreprise or administrator 

 $sidstring = (New-Object System.Security.Principal.SecurityIdentifier($_["objectsid"][0], 0)).Value 

      if ($sidstring -like "*-512" ) {
        $admindomaine = $_.name
        Get-ADGroupMember -identity "$admindomaine"  -Recursive  | ForEach-Object {	               
        $DomainAdminTable++
    }
      }
      if ( $sidstring -like "*-519" ) {
        $adminEnter = $_.name
        Get-ADGroupMember -identity "$adminEnter" -Recursive | ForEach-Object {                      
        $EnterpriseAdminTable++
        }
      }
    }

#Get objects in default OU

$DefaultComputersOU = (Get-ADDomain).computerscontainer
$DefaultComputersinDefaultOU = 0

Write-Host 'get All computer properties on default OU'

Get-ADComputer -Filter * -Properties OperatingSystem,created,PasswordLastSet -SearchBase "$DefaultComputersOU"  | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.Name
		'Enabled' = $_.Enabled
		'Operating System' = $_.OperatingSystem
		'Created' = $_.created
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

#Get Last Locked Users

Write-Host Get last locked users

Search-ADAccount -LockedOut -UsersOnly  | ForEach-Object { 
	
	$obj = [PSCustomObject]@{
		
		'name'    = $_.name
		'samaccountname'    = $_.samaccountname
		'lastlogondate ' = $_.lastlogondate 
	}
	
	$Unlockusers.Add($obj)
}

#Tenant Domain
$ForestObj | Select-Object -ExpandProperty upnsuffixes | ForEach-Object {
	
	$obj = [PSCustomObject]@{
		
		'UPN Suffixes'     = $_
		 'Valid'		   = "True"
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
$CustomGroup = 0
$DefaultGroup = 0
$Groupswithmemebrship = 0
$Groupswithnomembership = 0
$GroupsProtected = 0
$GroupsNotProtected = 0
$totalgroups = 0
$DistroCount = 0 
# Filter non privilelge groups
$Skipdefaultadmingroups = "(&(!(groupType:1.2.840.113556.1.4.803:=1))(!(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))))"
Get-ADGroup -LDAPFilter $Skipdefaultadmingroups -ResultSetSize $maxsearchergroups -Properties Member,ManagedBy,info,created,ProtectedFromAccidentalDeletion | Where-Object {$DefaultSGs -notcontains $_.Name} | ForEach-Object {

$totalgroups++
$OwnerDN = $null

if  (!$_.member) { 

$Groupswithnomembership++
    
    if ($($_.ManagedBy)) {
    $OwnerDN = ($_.ManagedBy -split (",") | Where-Object {$_ -like "CN=*"}) -replace ("CN=","")
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
	
		$CustomGroup++
        $users = ($_.member -split (",") | Where-Object {$_ -like "CN=*"}) -replace ("CN="," ") -join ","
        $OwnerDN = ($_.ManagedBy -split (",") | Where-Object {$_ -like "CN=*"}) -replace ("CN=","")

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.name
		'Type' = $Type
		'Members' = $users
		'Managed By' = $OwnerDN
        'Created' = ($_.created.ToString("yyyy/MM/dd"))
        'Remark' = $_.info
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
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
Get-ADOrganizationalUnit -Filter * -Properties ProtectedFromAccidentalDeletion,whenchanged -SearchScope $OUlevelSearch | ForEach-Object {
	
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
            if (!$nogpomod) {
			$LinkedGPOs.Add((Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName)
        }
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
$SkipReadPSO = $false

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
'AccountExpirationDate'
'msDS-PSOApplied'
)

#Get newly created users
$When = ((Get-Date).AddDays(-$UserCreatedDays)).Date

#Get expxired account and still enabled
$dateexpiresoone = (Get-DAte).AddDays(7)
$expiredsoon = 0
$expired = (get-date)

#Get users that haven't logged on in X amount of days, var is set at start of script
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days

$filterusers = "(!(admincount=*))"

if ($Showadmin.IsPresent) {
$filterusers = "(samaccountname=*)"
}

Get-ADUser -LDAPFilter $filterusers -Properties $Alluserpropert -ResultSetSize $maxsearcher | ForEach-Object {
 
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
    if ($_.AccountExpirationDate -lt $dateexpiresoone -and $null -ne $_.AccountExpirationDate -and $_.enabled -eq $true) {
    	
    if ($_.AccountExpirationDate -gt $expired) { $expiredsoon++ } 

    else {

    $ExpiringAccountsTable++
}
}

    #User Never loged and Enabled
    if (($null -eq $_.lastlogondate) -and ($_.Enabled -ne $false)) 
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
            if ($null -eq $_.lastlogondate) {
            #$daystoexpire = "User has never logged or must change password next logon, identified by code -999"
            $daystoexpire = -999 
            } else {
            $daystoexpire = -998
            }

		}
		
		else
		{
			
			#Check for Fine Grained Passwords
			
            if ($($_."msDS-PSOApplied") -and $SkipReadPSO -ne $true ) {

			$PasswordPol = (Get-ADUserResultantPasswordPolicy $_ -ErrorAction SilentlyContinue -ErrorVariable PSOeror).MaxPasswordAge.days

            if ($PSOeror) {
            Write-Host "Cannot read a PSO, pls check error or permissions" -ForegroundColor Yellow
            $PSOeror
            $SkipReadPSO = $true
            }

            $expireson = $passwordsetdate.AddDays($PasswordPol)
            $PasswordPol = $null

            }  else {
		
			$expireson = $passwordsetdate.AddDays($maxPasswordAge)
            }

			$today = (Get-Date)
			
			#Gets the count on how many days until the password expires and stores it in the $daystoexpire var
			$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
		}
	}
	
	else
	{
		#Users not need change passwords
        $daystoexpire = 0

	}	


	if (($_.Enabled -eq $True) -and ($lastlog -lt ((Get-Date).AddDays(-$Days))) -and ($NULL -ne $_.LastLogon))
	{
        $userinactive++
	}

    #Get User created Last 30 days
    if ( $_.whenCreated -ge $createdlastdays ) 
    {

    $lastcreatedusers++

    $barcreated = ($_.Whencreated.ToString("yyyy/MM/dd"))

    $rec= $barcreateobject | Where-Object {$_.date -eq $barcreated } 

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
	

	$obj = [PSCustomObject]@{
		
		'Name'				      = $_.Name
		'UserPrincipalName'	      = $_.UserPrincipalName
		'Enabled'				  = $_.Enabled
		'Protected from Deletion' = $_.ProtectedFromAccidentalDeletion
		'Last Logon'			  = $lastlog
        'Last Logon Date'         = $_.LastLogonDate
        'Created'                 = $_.whencreated
        'OU'                      = (($_.DistinguishedName -split (",") | Where-Object {$_ -like "OU=*"}) -replace ("OU=","") -join ",")
		'Email Address'		      = $_.EmailAddress
		'Account Expiration'	  = $_.AccountExpirationDate
        'Description'             =  $_.description
		'Password Last Set'	      = $_.PasswordLastSet
		'Password Never Expired'  = $_.PasswordNeverExpires
        'Days Until password expired' = $daystoexpire   
	}
	
	$usertable.Add($obj)
	
	if ($daystoexpire -lt $DaysUntilPWExpireINT -and $daystoexpire -ge 0)
	{
		
		
		$PasswordExpireSoonTable++
	}
}


#TOP User table
	$objULic = [PSCustomObject]@{
		'Total Users' = $totalusers
		"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable
		'Expiring Accounts' = $ExpiringAccountsTable
	}	
	$TOPUserTable.Add($objULic)

Write-Host "Done!" -ForegroundColor White
#endregion Users

#region GPO
<###########################

	   Group Policy

############################>


if (!$nogpomod) {
Write-Host "Working on Group Policy Report..." -ForegroundColor Green

$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'

#Get GPOs Not Linked
#region gponotlinked
$adObjects = $linkedGPO = $null
# info: # gpLink est une chaine de caractere de la forme [LDAP://cn={C408C216-5CEE-4EE7-B8BD-386600DC01EA},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={C408C16-5D5E-4EE7-B8BD-386611DC31EA},cn=policies,cn=system,DC=domain,DC=com;0]

[System.Collections.Generic.List[PSObject]]$adObjects = @()
[System.Collections.Generic.List[PSObject]]$linkedGPO = @()

$configuration = ($DomainControllerobj.SubordinateReferences | Where-Object {$_ -like '*configuration*' }).trim()

$domainAndOUS = Get-ADObject -LDAPFilter "(&(|(objectClass=organizationalUnit)(objectClass=domainDNS))(gplink=*))" -SearchBase "$($DomainControllerobj.DistinguishedName)" -Properties gpLink
$sites = Get-ADObject -LDAPFilter "(&(objectClass=site)(gplink=*))" -SearchBase "$configuration" -Properties gpLink

# construit une liste avec tous les gpLink existants
$adObjects.Add($domainAndOUS)
$adObjects.Add($sites)

    # Compare si GUID de la GPO existe dans les gpLink
    # gpLink est une chaine de caractere de la forme [LDAP://cn={C408C216-5CEE-4EE7-B8BD-386600DC01EA},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={C408C16-5D5E-4EE7-B8BD-386611DC31EA},cn=policies,cn=system,DC=domain,DC=com;0]
    
    $GPOs | ForEach-Object {
    
    if($adObjects.gpLink -match $_.id) {
        $linkedGPO.Add($_.DisplayName)
    }
}

# cast en array pour prendre en considération le cas où un seul object
$gponotlinked = ([array]($gpos | Where-Object {$_.DisplayName -notin $linkedGPO}) | Select-Object DisplayName,CreationTime,GpoStatus)

if (!$gponotlinked) { $gponotlinked = 0 }
 
#endregion gponotlinked

if (($GPOs).Count -eq 0)
{
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Group Policy Obejects were found'
	}
	$GPOTable.Add($obj)
}
Write-Host "Done!" -ForegroundColor White
}
#endregion GPO

#region Printers
<###########################

		   Printers

############################>

Write-Host "Working on Printers Report..." -ForegroundColor Green
$printersnbr = 0

$printers = Get-AdObject -filter "objectCategory -eq 'printqueue'" -Properties description,drivername,created,location | Select-Object name,description,drivername,created,location

$printersnbr = ($printers.name).count

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

Get-ADComputer -LDAPFilter "(!(userAccountControl:1.2.840.113556.1.4.803:=8192))" -Properties $filtercomputer -ResultSetSize $maxsearcher |  ForEach-Object {

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

    $rec=$barcreateobject | Where-Object {$_.date -eq $barcreated } 

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


 if (($_.OperatingSystem -match 'Windows Embedded Standard')) { 
    if ($_.OperatingSystemVersion -like '6.1*') {
    $_.OperatingSystem = 'Windows Embedded 7 Standard'}
 } 

 if ($_.OperatingSystem -Like 'Windows Server® 2008 *') { 
    $_.OperatingSystem = $_.OperatingSystem -replace '®'}

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
        'OU'           = (($_.DistinguishedName -split (",") | Where-Object {$_ -like "OU=*"}) -replace ("OU=","") -join ",")
		'Password Last Set' = $_.PasswordLastSet
        'Last Logon Date'   = $_.LastLogonDate
		'Protect from Deletion' = $_.ProtectedFromAccidentalDeletion
        'Build' = $Winbuild
        'IPv4Address' = $_.IPv4Address
	}
	
	$ComputersTable.Add($obj)	
    
    if ($_.OperatingSystem -Like 'Windows*7*'){
    $OSClass['Windows 7'] += 'Windows 7'.Count
    } elseif ($_.OperatingSystem -Like 'Windows 8*' -or $_.OperatingSystem -Like 'Windows Embedded 8*') {
    $OSClass['Windows 8'] += 'Windows 8'.Count
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

if ($barcreateobject.Count -eq 1) {
 $obj = [PSCustomObject]@{		
		'Nbr_users' = 0
        'Nbr_PC' = 0
		'Date' = (Get-Date).AddDays(+1)
	}
$barcreateobject.Add($obj)
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
$contacts = (Get-ADObject -Filter 'objectclass -eq "contact"').name
$totalcontacts = ($contacts).count
$Allobjects  = New-Object 'System.Collections.Generic.List[System.Object]'
$totalusers = $totalusers + (Get-ADUser -LDAPFilter "(admincount=*)").count
$totalcomputers = $totalcomputers + (Get-ADComputer -LDAPFilter "(userAccountControl:1.2.840.113556.1.4.803:=8192)"| select name ).name.count
$totalgroups = $totalgroups + ($DefaultSGs.count)

$Allobjects = @(
    [pscustomobject]@{Name='Groups';Count=$totalgroups}
    [pscustomobject]@{Name='Users'; Count=$totalusers}
    [pscustomobject]@{Name='Computers'; Count=$totalcomputers}
    [pscustomobject]@{Name='Contacts'; Count=$totalcontacts}
    [pscustomobject]@{Name='Printer server'; Count=$printersnbr}
)

Write-Host "Done!" -ForegroundColor White

#endregion Resume

#endregion code

$SavePath = $SavePath + '\ADModern.html'

if ($htmlonepage.IsPresent ) {

HTMLOnePage

} else {

HTMLMultiPage

}

}

function HTMLMultiPage {
#region generatehtml
$time = (get-date)
Write-Host "Working on HTML Report ..." -ForegroundColor Green
New-HTML -TitleText 'AD_ModernReport' -ShowHTML -Online -FilePath $SavePath {   
    New-HTMLNavTop -Logo $CompanyLogo -MenuColorBackground 	gray  -MenuColor Black -HomeColorBackground gray -HomeLinkHome {
       
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
   
     New-HTMLTabStyle  -BackgroundColorActive Teal   
 
      New-HTMLSection  -Name 'Block infos' -Invisible  {

      New-HTMLPanel -Margin 10 -Width "80%" {

      New-HTMLPanel -BackgroundColor silver  {
      New-HTMLText -TextBlock  {
      New-HTMLText -Text  "Domain : $Forest" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "AD Recycle Bin : $ADRecycleBin" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text "FSMO Roles" -Alignment center -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Infra : $InfrastructureMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Rid : $RIDMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "PDC  : $PDCEmulator" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Naming : $DomainNamingMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Schema : $SchemaMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -LineBreak
      
      }
      }

      }

      New-HTMLPanel -Margin 10  {
      
      New-HTMLPanel -BackgroundColor lightgreen -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $UserDisabled -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Disabled Users' -Alignment justify -FontSize 15 
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-slash fa-3x" } 
      } 
      }

      New-HTMLPanel -BackgroundColor yellowgreen -AlignContentText right -BorderRadius 10px {
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
      New-HTMLText -Text 'Users Never logged' -Alignment left -FontSize 15
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
      New-HTMLText -Text $($DomainAdminTable) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'domain admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-edit fa-3x" } 
      }
        }


      New-HTMLPanel -BackgroundColor steelblue -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($EnterpriseAdminTable) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Enterprise Admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-tie fa-3x" } 
      }
      }
      

      }     

      New-HTMLPanel -Margin 10  {

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $ComputerNotSupported -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Computers Not Supported' -Alignment left -FontSize 15
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
      New-HTMLText -Text 'Expired Account and still Enabled' -Alignment left -FontSize 15
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
      
      New-HTMLSection -HeaderBackGroundColor Teal -Invisible -Width "70%" {    

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
                    New-ChartPie -Name 'Default Groups' -Value ($DefaultSGs.count)
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

       New-HTMLSection -name 'Unlinked Details' -HeaderBackGroundColor Teal {
              New-HTMLTable -DataTable $gponotlinked 
       }

       New-HTMLSection -Name 'Linked Vs Unliked GPOs' -HeaderBackGroundColor Teal  {
            New-HTMLChart {
                New-ChartLegend -LegendPosition bottom 
                New-ChartBarOptions -Gradient
                New-ChartDonut -Name 'Unlinked' -Value $gponotlinked.Count -Color silver
                New-ChartDonut -Name 'Linked' -Value $GPOs.Count -Color orange
            }
        }


    }


    }
    }
    New-HTMLPage -Name 'Printers' {

       New-HTMLTab -Name 'Printer server' -IconSolid print {
        
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
                    New-ChartPie -Name "Password Never Expired" -Value $UserPasswordNeverExpires
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

         New-HTMLSection -name 'Potential Win10 End of support' {
                       
           New-HTMLChart {
                New-ChartLegend -LegendPosition bottom
                New-ChartDonut -Name 'Endofsupport' -Value $endofsupportwin
                New-ChartDonut -Name 'Windows 10/11 supported' -Value ($allwin1011 - $endofsupportwin)
            }

         }


         New-HTMLSection -HeaderText 'Computers Operating System Distribution' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
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

      New-HTMLPanel -Margin 10  {
      
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
              New-HTMLListItem -Text "Generated date : $time"
              New-HTMLListItem -Text 'Modern Active Directory _ Version : 1.0.9 _ Release : 03/2023' 
              New-HTMLListItem -Text 'Author : Dakhama Mehdi<br> 
              <br> Inspired ADReportHTLM Bradley Wyatt [thelazyadministrato](https://www.thelazyadministrator.com/)<br>
              <br> Credit : Thirrey Demon-Barcelo, Mattieu Souin, Mahmoud Hatira, Zouhair sarouti<br>
              <br> Thanks : Boss Przemyslaw Klys - Module PSWriteHTML- [Evotec](https://evotec.xyz)'
              } -FontSize 14
            }         
    New-HTMLPanel {
            New-HTMLImage -Source $RightLogo 
        } 
        }   
    }

    }    
} 
#endregion generatehtml
}

function HTMLOnePage {
#region generatehtml
$time = (get-date)
Write-Host "Working on HTML Report ..." -ForegroundColor Green
New-HTML -TitleText 'AD_ModernReport' -ShowHTML -Online -FilePath $SavePath {
New-HTMLHeader {
        New-HTMLSection -Invisible  {
            New-HTMLPanel -Invisible  {
               # New-HTMLImage -Source $CompanyLogo -AlternativeText 'My other text' -Width '20%'
               New-HTMLText  -Text "Modern Active Directory"  -FontSize 22 -Color White
            } -AlignContentText left -BackGroundColor Teal
        } 
    }      
 New-HTMLTab -Name 'Dashboard' -IconRegular chart-bar  {   
     New-HTMLTabStyle  -BackgroundColorActive teal    
      New-HTMLSection  -Name 'Block infos' -Invisible  {
      New-HTMLPanel -Margin 10 -Width "80%" {
      New-HTMLPanel -BackgroundColor silver  {
      New-HTMLText -TextBlock  {
      New-HTMLText -Text  "Domain : $Forest" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "AD Recycle Bin : $ADRecycleBin" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text "FSMO Roles" -Alignment center -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Infra : $InfrastructureMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Rid : $RIDMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "PDC  : $PDCEmulator" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Naming : $DomainNamingMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Schema : $SchemaMaster" -Alignment justify -FontSize 15 -FontWeight bold
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
      New-HTMLText -TextBlock { 
      New-HTMLText -Text $userinactive -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Users not loged in Last 90 Days' -Alignment justify -FontSize 15 
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
      New-HTMLText -Text 'Domain Admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-edit fa-3x" } 
      }
        }


      New-HTMLPanel -BackgroundColor steelblue -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($EnterpriseAdminTable.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Entreprise Admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-tie fa-3x" } 
      }
      }      

      }     

      New-HTMLPanel -Margin 10  {

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $ComputerNotSupported -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Computesr Not Supported' -Alignment left -FontSize 15
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
      New-HTMLText -Text 'Expired Account and still Enabled' -Alignment left -FontSize 15
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
                    New-ChartPie -Name 'Default Groups' -Value ($DefaultSGs.count)
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
 New-HTMLTab -Name 'Groups Without Members' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $Groupsnomembers
            }
        }


    }                
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
 New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       New-HTMLSection -Name 'Informations"'  {
                new-htmlTable  -DataTable $GPOs
            }
   
       
       New-HTMLSection {

       New-HTMLSection -name 'Unlinked Details' -HeaderBackGroundColor Teal {
              New-HTMLTable -DataTable $gponotlinked 
       }

       New-HTMLSection -Name 'Linked Vs Unliked GPOs' -HeaderBackGroundColor Teal  {
            New-HTMLChart {
                New-ChartLegend -LegendPosition bottom 
                New-ChartBarOptions -Gradient
                New-ChartDonut -Name 'Unlinked' -Value $gponotlinked.Count -Color silver
                New-ChartDonut -Name 'linked' -Value $GPOs.Count -Color orange
            }
        }


    }


    }          
 New-HTMLTab -Name 'Printer server' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $printers
            }
        }


    }  
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


         New-HTMLSection -HeaderText 'Computers Operating System Distrubiton' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
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
 New-HTMLTab -Name 'Resume' {        
    New-HTMLSection -HeaderBackGroundColor teal -Name 'All Members' -HeaderTextAlignment left  {
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
              New-HTMLListItem -Text "Generated date : $time"
              New-HTMLListItem -Text 'Modern Active Directory _ Version : 1.0.9 _ Release : 03/2023' 
              New-HTMLListItem -Text 'Author : Dakhama Mehdi<br> 
              <br> Inspired ADReportHTLM Bradley Wyatt [thelazyadministrato](https://www.thelazyadministrator.com/)<br>
              <br> Credit : Thirrey Demon-Barcelo, Mattieu Souin, Mahmoud Hatira, Zouhair sarouti<br>
              <br> Thanks : Boss Przemyslaw Klys - Module PSWriteHTML- [Evotec](https://evotec.xyz)'
              } -FontSize 14
            }            
       New-HTMLPanel {
            New-HTMLImage -Source $RightLogo 
        } 
        }   
    } 
   
} 

#endregion generatehtml
}
Export-ModuleMember -Function Get-ADModernReport

# SIG # Begin signature block
# MIIr0gYJKoZIhvcNAQcCoIIrwzCCK78CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBbshdxLumJ8Fg0
# O7F25HIHblxKRrfsMM6AWzjryHmvR6CCJOowggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggWNMIIEdaADAgECAhAOmxiO
# +dAt5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAi
# BgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAw
# MDBaFw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERp
# Z2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsb
# hA3EMB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iT
# cMKyunWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGb
# NOsFxl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclP
# XuU15zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCr
# VYJBMtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFP
# ObURWBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTv
# kpI6nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWM
# cCxBYKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls
# 5Q5SUUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBR
# a2+xq4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6
# MIIBNjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qY
# rhwPTzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8E
# BAMCAYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCg
# v0NcVec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQT
# SnovLbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh
# 65ZyoUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSw
# uKFWjuyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAO
# QGPFmCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjD
# TZ9ztwGpn1eqXijiuZQwggYaMIIEAqADAgECAhBiHW0MUgGeO5B5FSCJIRwKMA0G
# CSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExp
# bWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBSb290
# IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5NTlaMFQxCzAJBgNVBAYT
# AkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28g
# UHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0GCSqGSIb3DQEBAQUAA4IB
# jwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjIztNsfvxYB5UXeWUzCxEe
# AEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NVDgFigOMYzB2OKhdqfWGV
# oYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/36F09fy1tsB8je/RV0mIk
# 8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05ZwmRmTnAO5/arnY83jeNzh
# P06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm+qxp4VqpB3MV/h53yl41
# aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUedyz8rNyfQJy/aOs5b4s+
# ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz44MPZ1f9+YEQIQty/NQd/
# 2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBMdlyh2n5HirY4jKnFH/9g
# Rvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQYMBaAFDLrkpr/NZZILyhA
# QnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritUpimqF6TNDDAOBgNVHQ8B
# Af8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcD
# AzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsGA1UdHwREMEIwQKA+oDyG
# Omh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5n
# Um9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsGAQUFBzAChjpodHRwOi8v
# Y3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RSNDYu
# cDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG
# 9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURhw1aVcdGRP4Wh60BAscjW
# 4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0ZdOaWTsyNyBBsMLHqafvIh
# rCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajjcw5+w/KeFvPYfLF/ldYp
# mlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNcWbWDRF/3sBp6fWXhz7Dc
# ML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalOhOfCipnx8CaLZeVme5yE
# Lg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJszkyeiaerlphwoKx1uHRz
# NyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z76mKnzAfZxCl/3dq3dUNw
# 4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5JKdGvspbOrTfOXyXvmPL6
# E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHHj95Ejza63zdrEcxWLDX6
# xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2Bev6SivBBOHY+uqiirZt
# g0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/L9Uo2bC5a4CH2RwwggZO
# MIIEtqADAgECAhBgu3yFSQ7/ewyXbnO66QF/MA0GCSqGSIb3DQEBDAUAMFQxCzAJ
# BgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNl
# Y3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwHhcNMjIwOTI3MDAwMDAw
# WhcNMjMwOTI3MjM1OTU5WjBlMQswCQYDVQQGEwJGUjEmMCQGA1UECAwdUHJvdmVu
# Y2UtQWxwZXMtQ8O0dGUtZOKAmUF6dXIxFjAUBgNVBAoMDURBS0hBTUEgTUVIREkx
# FjAUBgNVBAMMDURBS0hBTUEgTUVIREkwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQCeqJVocM6guJh46zi0xvEc99tdwGM54OKAnMwgFgTIzRCWf8v2Im6X
# ROG+2VX6WhdincTRtbGpVeguLyWyQkfiKwzhigRmGr9l4iXROQUBv/dd/bywmHRu
# rWtiQ1iEcT68xp4d9vgHRYu4oLyv0Lrkn9mfnRDi1QxsIdrAixecvd/5Iyp8tv39
# mjSR9GeALxpNy13qedHd0gtggUrrtsTao583EUQhbQJi+2xqZrNrrMvmqYSEyV0R
# QMtYBkwZ2b7HgCOfXGlX3s08YKaoXF/6R4Zz2quhEtJjv0ge5SL6Ek5J4PGS+Syw
# B6O7stflVGtW7jr/OoaWD+3g5+fkg9DaRmhlYDGd/exMc88Er9ewKCA0FpC+KHE3
# 3Gra3SFXQjRLp7WBWdtJU44+b6l/GZ6qecsfR6T0MLDEsTNU5X4mPzBYiPG+FRxC
# Ca7d4hihJL6cjXmHHUiWcgbr9Vqe/PM+I+2mHz71Ss16nIGfyC6aBfqLUp0tN8Oz
# 0i5EOGV7mfjRqtyQmd8DqbE1bxa3EWqVdvVIbA4OG3uFiO3mK7UDSZaJHkAyAFpD
# aWHPAjwZ6MJsdiBkrrkxdX0saDGYhuNpm9cxbN58OFXkiwrtNE3I+qFRCvgNXB2x
# UV7sad5GVK3p6dIp5Hj3SKE0xBjA+5rb/YPjDqVgIT97nprmZtfBlQIDAQABo4IB
# iTCCAYUwHwYDVR0jBBgwFoAUDyrLIIcouOxvSK4rVKYpqhekzQwwHQYDVR0OBBYE
# FOftFEnJB1xi+ppnWpBm+3qIQlU+MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8E
# AjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMEoGA1UdIARDMEEwNQYMKwYBBAGyMQEC
# AQMCMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMAgGBmeB
# DAEEATBJBgNVHR8EQjBAMD6gPKA6hjhodHRwOi8vY3JsLnNlY3RpZ28uY29tL1Nl
# Y3RpZ29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNybDB5BggrBgEFBQcBAQRtMGsw
# RAYIKwYBBQUHMAKGOGh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1B1Ymxp
# Y0NvZGVTaWduaW5nQ0FSMzYuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5z
# ZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAYEAjX8plWtErXDsSXX06XnK9OD6
# /aDKDZ/EVuMo74OmANrihNH/dWzyLBBjL6wZEizeC3qz7BBlfi0EOEEQeFj/0tph
# vYfCIupCdXwdClCbbJQXZ8hHB4r8lHiwxpR1+XgEaFrdeMANRkkVgEf/FBZ9I4sQ
# 9o4XWjN6UhfLD5JcwXcKh/pFIKLhTMRFKZXmsyEFFB7HNvaddQuy/EbW9YZcQtio
# JCtAjC7UPrPCNzKZam9DyPQDblHbyQn2Bsb0STHqlEtkS4MY8JyIIm1xKiHnNCBI
# /f3VWUMtjliNaKkchK8gGqEwSQ2a1jvb2dg0UDL/4YJsb8dTaVCYY/PI18Es+bbg
# G+ck0Vh4XN/n4UCe2/FMQdHFegapEJ01OsshB4tRL99r6f3x0WHY8aSDo29f1Ikj
# Ol6MrNwT21vlmi6RYiZvzENMP4Uj1PY1vt1ik+OCWYCSnCnp/r7vdOv9asKGU+Xm
# jJ0r6DgfUM/L7z6T44Fhm61juvMrJtTcfJvb2Ot+MIIGrjCCBJagAwIBAgIQBzY3
# tyRUfNhHrP0oZipeWzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAw
# WhcNMzcwMzIyMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBT
# SEEyNTYgVGltZVN0YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAxoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQV
# Ql+kiPNo+n3znIkLf50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY
# 3+vaPcQXf6sZKz5C3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB7
# 20RBidx8ald68Dd5n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71
# ORJn7w6lY2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW
# 8LAxE6lXKZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7
# W+yB3iIU2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qq
# lnNCaJ+2RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWc
# ZklSUPRR8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI
# /rOP015LdhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27Ktd
# RnXiYKNYCQEoAA6EVO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0w
# ggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WM
# aiCPnshvMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RH
# NC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIw
# CwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbY
# IULhsBguEE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6
# hCAlNDFnzbYSlm/EUExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6
# q4XNQ1/tYLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/
# KFSJ8NWKcXZl2szwcqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/E
# jaDnmPv7pp1yr8THwcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNT
# rDsdCEkPlM05et3/JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRx
# ktcma+Q4c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7
# K9h+8kaddSweJywm228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrd
# VcA6KYawmKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C
# +dAjfwAL5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QV
# SucTDh3bNzgaoSv27dZ8/DCCBsAwggSooAMCAQICEAxNaXJLlPo8Kko9KQeAPVow
# DQYJKoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hB
# MjU2IFRpbWVTdGFtcGluZyBDQTAeFw0yMjA5MjEwMDAwMDBaFw0zMzExMjEyMzU5
# NTlaMEYxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDEkMCIGA1UEAxMb
# RGlnaUNlcnQgVGltZXN0YW1wIDIwMjIgLSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAz+ylJjrGqfJru43BDZrboegUhXQzGias0BxVHh42bbySVQxh
# 9J0Jdz0Vlggva2Sk/QaDFteRkjgcMQKW+3KxlzpVrzPsYYrppijbkGNcvYlT4Dot
# jIdCriak5Lt4eLl6FuFWxsC6ZFO7KhbnUEi7iGkMiMbxvuAvfTuxylONQIMe58ty
# SSgeTIAehVbnhe3yYbyqOgd99qtu5Wbd4lz1L+2N1E2VhGjjgMtqedHSEJFGKes+
# JvK0jM1MuWbIu6pQOA3ljJRdGVq/9XtAbm8WqJqclUeGhXk+DF5mjBoKJL6cqtKc
# tvdPbnjEKD+jHA9QBje6CNk1prUe2nhYHTno+EyREJZ+TeHdwq2lfvgtGx/sK0YY
# oxn2Off1wU9xLokDEaJLu5i/+k/kezbvBkTkVf826uV8MefzwlLE5hZ7Wn6lJXPb
# wGqZIS1j5Vn1TS+QHye30qsU5Thmh1EIa/tTQznQZPpWz+D0CuYUbWR4u5j9lMNz
# IfMvwi4g14Gs0/EH1OG92V1LbjGUKYvmQaRllMBY5eUuKZCmt2Fk+tkgbBhRYLqm
# gQ8JJVPxvzvpqwcOagc5YhnJ1oV/E9mNec9ixezhe7nMZxMHmsF47caIyLBuMnnH
# C1mDjcbu9Sx8e47LZInxscS451NeX1XSfRkpWQNO+l3qRXMchH7XzuLUOncCAwEA
# AaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB
# /wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwH
# ATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUYore
# 0GH8jzEU7ZcLzT0qlBTfUpwwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRp
# bWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAVaoqGvNG83hXNzD8
# deNP1oUj8fz5lTmbJeb3coqYw3fUZPwV+zbCSVEseIhjVQlGOQD8adTKmyn7oz/A
# yQCbEx2wmIncePLNfIXNU52vYuJhZqMUKkWHSphCK1D8G7WeCDAJ+uQt1wmJefkJ
# 5ojOfRu4aqKbwVNgCeijuJ3XrR8cuOyYQfD2DoD75P/fnRCn6wC6X0qPGjpStOq/
# CUkVNTZZmg9U0rIbf35eCa12VIp0bcrSBWcrduv/mLImlTgZiEQU5QpZomvnIj5E
# IdI/HMCb7XxIstiSDJFPPGaUr10CU+ue4p7k0x+GAWScAMLpWnR1DT3heYi/HAGX
# yRkjgNc2Wl+WFrFjDMZGQDvOXTXUWT5Dmhiuw8nLw/ubE19qtcfg8wXDWd8nYive
# QclTuf80EGf2JjKYe/5cQpSBlIKdrAqLxksVStOYkEVgM4DgI974A6T2RUflzrgD
# QkfoQTZxd639ouiXdE4u2h4djFrIHprVwvDGIqhPm73YHJpRxC+a9l+nJ5e6li6F
# V8Bg53hWf2rvwpWaSxECyIKcyRoFfLpxtU56mWz06J7UWpjIn7+NuxhcQ/XQKuji
# Yu54BNu90ftbCqhwfvCXhHjjCANdRyxjqCU4lwHSPzra5eX25pvcfizM/xdMTQCi
# 2NYBDriL7ubgclWJLCcZYfZ3AYwxggY+MIIGOgIBATBoMFQxCzAJBgNVBAYTAkdC
# MRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVi
# bGljIENvZGUgU2lnbmluZyBDQSBSMzYCEGC7fIVJDv97DJduc7rpAX8wDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgq2GmvcFcMet8JxoKVjA1kB5nZGrKA18mGI7jio4FsTcw
# DQYJKoZIhvcNAQEBBQAEggIAWxuD3eI9rDuZovj6y/7M0UnCbQHq5TvyYSLuQDDc
# yoHy+OMA5SN3B3+RzyVwcJsX9kQUPxQvSQacfbCkZeg/DGNGJa+Im2Q2kgHwdsy6
# Cvp3KI+iVbFAMbQN5L67GPgYcroYTQ+uVS4JGCpe9H1ruIoscoyz7iW0/vp09y+d
# fs1yEd9HqLmidAV7cWGHLGSMWD1veCjlKnWQAv337sMR4OirMYxOD2cyiBh00yRU
# pDDuQD1alv++kHuj2JIgwu7LFM+O3ePgLig9GzB+OJyUD27/xSCOEFL/nLSazIId
# ozGohGgxIfzKCxLGdHvPIcdc1KlN/cxmeS3N+zkjtQ24pEhWUTB69arOKP5d+0xh
# XaQYAcODosSC/0RBY9TbPIIgEAIl6ZXiHqxgohAdEMwUN+wzhiDExbX/0pos0ZAg
# 25Ry9qLD+1h98NmMpWBOOw18C4fRurmGNyF5YiU5e3Pz3PtFAxrBju7HAO+8paR9
# XGjmSSRsZqwpB7JfjDum1PjOfkFi4CbcqevyLyn3izalGHmzDF+cBnNNrlFVOegK
# MJEN70szdb69Gxzr5bIMLTKa7hURMguJX2yUgHuoC8tfUM7afwrUB4EbsxcDXhKl
# rAHtyqyKqwQHVDbPRAnQFWeQ8WIfBV8NmhbK0cIeJ8kahNWHWd8sAqT4gaP/H+01
# UPehggMgMIIDHAYJKoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBjMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRy
# dXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhAMTWlyS5T6
# PCpKPSkHgD1aMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMjMwMzIzMjIwMDAxWjAvBgkqhkiG9w0BCQQx
# IgQgy/go9QztoZdaXIsDaGw+/Aj3eQHhTBUl+TwAEcS1c7wwDQYJKoZIhvcNAQEB
# BQAEggIARxXWpN+IDUcNuAqdkka+49Bz4VRja2VmfIQ2yTuyZsuiE0giBQaf1xma
# r0e0xEUpyLTinHRy2Iob3JHr5Ou36cp0XwEQZOzHKFrjkIbDnE6RtFHFwBvtZz6m
# Ff3qgPjZJV5MDCp7qMh4EfHKZ/O9u4qP6xYJy8C5SGONwPnSMdZQD7MoIdM6m2s2
# 4cIcehtfVC963jpIzsS+tEc8HQN9mL8BAMA8eKskHcWn5iexJ6Pfn5nnnxW01WYx
# Fu+HZYL+YJOH2gfv341rSpgeEkdTbeAxxgZs7oNjRGCMiFhEcKCf9Es4dfiRJX/q
# 0FsIidFXjo4lCOYouOJPUm4HPEB/QCU12yVZ/eeKn+FM72tFpMlqm9UbnwGua9p5
# ECGiJj7uXNhp9TSNhblCZJAXbZHvDUmbYHuzFsr61dH9oh5S7kqkdofEW9dtMSy/
# i87+luUAy/GNjKrQswylTt2qNT1JE9BM+t6cMC3hYdJGXlntu6tmIqDEaFFNBCgh
# gIQipYwm0+95JRestIrkj1ZfvbQ0gHosxT8uuxoszRpBkxmZS/JiZ8sJMD2+l2o0
# vz2FGwwFZFrQCX/HVXAPZ5fLCn29qxkdbrkKIkyXdocEzca0ijZwx7nKMDgjwq9S
# 4LhiRR+LfgsxtGc+JZTzlJJ5qi+WFU/XqqhGUBu6ktkaCJOfvng=
# SIG # End signature block
