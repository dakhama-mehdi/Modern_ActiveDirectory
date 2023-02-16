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