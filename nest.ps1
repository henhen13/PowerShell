# Ensure the Active Directory module is imported
Import-Module ActiveDirectory

# Define the group names
$salesManagerGroup = "Salgssjef"
$salesRepGroup = "Salgsrepresentanter"
$itManagerGroup = "IT Sjef"
$itTeamGroup = "IT Team"

# Define the OU path
$ouPath = "OU=Grupper,OU=FSI-HENHEN,DC=fsi-henhen,DC=com"

# Create the Sales group if it does not exist
if (-not (Get-ADGroup -Filter { Name -eq "Salg" } -ErrorAction SilentlyContinue)) {
    New-ADGroup -Name "Salg" -GroupScope Global -GroupCategory Security -Description "Sales group" -Path $ouPath
    Write-Host "Created group: Salg in $ouPath"
}

# Create the IT group if it does not exist
if (-not (Get-ADGroup -Filter { Name -eq "IT" } -ErrorAction SilentlyContinue)) {
    New-ADGroup -Name "IT" -GroupScope Global -GroupCategory Security -Description "IT group" -Path $ouPath
    Write-Host "Created group: IT in $ouPath"
}

# Create nested groups for Sales if they do not exist
foreach ($group in @($salesManagerGroup, $salesRepGroup)) {
    if (-not (Get-ADGroup -Filter { Name -eq $group } -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $group -GroupScope Global -GroupCategory Security -Description "$group group" -Path $ouPath
        Write-Host "Created group: $group in $ouPath"
    } else {
        Write-Host "Group already exists: $group"
    }
}

# Create nested groups for IT if they do not exist
foreach ($group in @($itManagerGroup, $itTeamGroup)) {
    if (-not (Get-ADGroup -Filter { Name -eq $group } -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $group -GroupScope Global -GroupCategory Security -Description "$group group" -Path $ouPath
        Write-Host "Created group: $group in $ouPath"
    } else {
        Write-Host "Group already exists: $group"
    }
}

# Nest the Sales groups
Add-ADGroupMember -Identity "Salg" -Members $salesManagerGroup, $salesRepGroup
Write-Host "Added nested groups to Salg"

# Nest the IT groups
Add-ADGroupMember -Identity "IT" -Members $itManagerGroup, $itTeamGroup
Write-Host "Added nested groups to IT"

# Move specified users to the Sales representatives group
$usernamesToMove = @("erijoh", "perped", "bjolse") # Adjust these usernames accordingly

foreach ($username in $usernamesToMove) {
    $adUser = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction SilentlyContinue
    if ($adUser) {
        Add-ADGroupMember -Identity $salesRepGroup -Members $adUser.SamAccountName
        Write-Host "Moved $username to group: $salesRepGroup"
    } else {
        Write-Host "User not found: $username"
    }
}

Write-Host "Nested groups created and users moved successfully."
