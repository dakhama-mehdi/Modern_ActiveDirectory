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

Get-ADGroup -Filter "name -notlike '*Exchange*'" -ResultSetSize $maxsearchergroups -Properties Member,ManagedBy,info,created,ProtectedFromAccidentalDeletion  | ForEach-Object {

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
        $users = ($_.member -split (",") | Where-Object {$_ -like "CN=*"}) -replace ("CN="," ") -join ","

	}
	
	else
	{
		$DefaultADGroup = "True"
		$DefaultGroup++
        $Users = "Skipped Domain Users Membership"

	}


    $OwnerDN = ($_.ManagedBy -split (",") | Where-Object {$_ -like "CN=*"}) -replace ("CN=","")

	
	$obj = [PSCustomObject]@{
		
		'Name' = $_.name
		'Type' = $Type
		'Members' = $users
		'Managed By' = $OwnerDN
        'Created' = ($_.created.ToString("yyyy/MM/dd"))
        'Remark' = $_.info
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