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