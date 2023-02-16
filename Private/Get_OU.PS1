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