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