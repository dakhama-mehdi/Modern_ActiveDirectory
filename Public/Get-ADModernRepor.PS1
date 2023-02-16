function Get-ADModernReport{

<#
    .SYNOPSIS
    New Experience to Manage Active Directory over interactive HTML.

    .DESCRIPTION
    This Module help to create a Dynamic Web Report to manage Active Directory.

    .EXAMPLE
    Create a sample report multipages for test, note by default only 300 objects will be listed.
    Get-ADWebReport 
    .EXAMPLE
    Create a report for illimited objects
    Get-ADWebReport -illimitedsearch
    .EXAMPLE
    Create onepage report and save in specific folder, We can change the name output file like Mycompany.html
    Get-ADWebReport -SavePath C:\myfolder\ADWeb.HTML -htmloneline
    .EXAMPLE
    Create report with companylogo and limited listed groups to 3000 and object to 5000
    Get-ADWebReport -CompanyLogo C:\myfolder\ADWeb.PNG -maxsearchergroups 3000 -maxsearcher 5000

    .NOTES
    Based on https://powersnippets.com/convertto-flatobject/
    #>


 [CmdletBinding()]
param (
    #Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$CompanyLogo = "https://github.com/dakhama-mehdi/ActiveDirectory_ModernReport/blob/main/Pictures/SmallLogo.png?raw=true",

    #Logo that will be on the right side, UNC or URL
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$RightLogo = "https://github.com/dakhama-mehdi/ActiveDirectory_ModernReport/blob/main/Pictures/Rightlogo.png?raw=true",

    #Title of generated report
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$ReportTitle = "Active Directory Over HTML",

    #Location the report will be saved	
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
	[String]$SavePath = "C:\Temp",

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
    [switch]$htmloneline
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

throw "AD RSAT Module is required, pls install it, operation aborted" 

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
$DefaultSGs += ([adsisearcher]"(&(groupType:1.2.840.113556.1.4.803:=1)(!(objectSID=S-1-5-32-546))(!(objectSID=S-1-5-32-545)))").findall().Properties.name
$DefaultSGs += ([adsisearcher] "(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))").FindAll().Properties.name

#region PScutom
$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$Groupsnomembers = New-Object 'System.Collections.Generic.List[System.Object]'
$OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ExpiringAccountsTable = New-Object 'System.Collections.Generic.List[System.Object]'
$Unlockusers = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'
$EnterpriseAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
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
	
	Write-Host "GPMC Feature is not present, Pls install it to get info" -ForegroundColor Red
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



#Get Domain Admins and Entreprise admins
#search domain admins default group and entreprise andministrators
#This is disapriecied because i dont wont to list the sensible informations

([adsisearcher] "(&(objectCategory=group)(admincount=1)(iscriticalsystemobject=*))").FindAll().Properties | ForEach-Object {

#List group contains admins domain or entreprise or administrator 

 $sidstring = (New-Object System.Security.Principal.SecurityIdentifier($_["objectsid"][0], 0)).Value 

      if ($sidstring -like "*-512" ) {

        $admindomaine = $_.name
        Get-ADGroupMember -identity "$admindomaine"  -Recursive  | ForEach-Object {
	       
        $obj = [PSCustomObject]@{
            
            'Name'    = "$_.Name"
            'Type'    = $_.ObjectClass
        }
        
        $DomainAdminTable.Add($obj)
    }
      }

      if ( $sidstring -like "*-519" ) {

        $adminEnter = $_.name
        Get-ADGroupMember -identity "$adminEnter" -Recursive | ForEach-Object {
           
            $obj = [PSCustomObject]@{
                
                'Name'    = "$_.Name"
                'Type'    = $_.ObjectClass
            }
            
            $EnterpriseAdminTable.Add($obj)
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
}