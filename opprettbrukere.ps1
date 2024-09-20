# Ensure the Active Directory module is imported
Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the file path for the Excel file
$excelFilePath = "C:\Users\Administrator\Documents\Example Users and Permissions.xlsx"

# Import the Excel data
$users = Import-Excel -Path $excelFilePath -WorksheetName 'kontoer'

# Function to replace Norwegian characters
function Replace-NorwegianChars {
    param (
        [string]$inputString
    )
    $outputString = $inputString -replace 'æ', 'e' `
                                    -replace 'ø', 'o' `
                                    -replace 'å', 'a'
    return $outputString
}

# Function to generate username
function Generate-Username {
    param ($firstName, $lastName)

    # Validate input
    if (-not $firstName -or -not $lastName) {
        Write-Host "First name or last name is missing. Cannot generate username."
        return $null
    }

    # Take the first 3 letters of first and last name
    $username = ($firstName.Substring(0, 3) + $lastName.Substring(0, 3)).ToLower()

    # Replace Norwegian characters in the username
    $username = Replace-NorwegianChars $username

    # Check if the username exists in AD
    $existingUser = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction SilentlyContinue

    # If the username exists, modify the username by adjusting the 3rd letter of the first name
    if ($existingUser) {
        $username = ($firstName.Substring(0, 2) + $firstName.Substring(3, 1) + $lastName.Substring(0, 3)).ToLower()
        $username = Replace-NorwegianChars $username
    }

    return $username
}

# Loop through the users and create each user in Active Directory
foreach ($user in $users) {
    try {
        # Generate username based on first and last name
        $samAccountName = Generate-Username -firstName $user.Fornavn -lastName $user.Etternavn

        # Check if the username was generated successfully
        if (-not $samAccountName) {
            Write-Host "Skipping user due to invalid username generation."
            continue
        }

        # Prepare user properties
        $userPrincipalName = "$samAccountName@fsi-henhen.com"
        $password = ConvertTo-SecureString "InitialP@ssword123" -AsPlainText -Force  # Set an initial password

        # Get phone number and initials
        $phoneNumber = $user.mobiltelefon # Adjust this to match the correct column name in your Excel file
        $initials = $user.Initial # Adjust this to match the correct column name in your Excel file
        
        # Determine the Organizational Unit (OU) based on the department (Avdeling)
        switch ($user.Avdeling) {
            "Administrasjon" { $OU = "OU=Administrasjon, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            "Salg" { $OU = "OU=Salg, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            "Utvikling" { $OU = "OU=Utvikling, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            "Kundesupport" { $OU = "OU=Kundesupport, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            "HR" { $OU = "OU=HR, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            "IT" { $OU = "OU=IT, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
            default { $OU = "OU=Administrasjon, OU=FSI-HENHEN, DC=fsi-henhen,DC=com" }
        }

        # Create the new user in Active Directory with OU and password settings
        New-ADUser -Name "$($user.Fornavn) $($user.Etternavn)" `
                   -GivenName $user.Fornavn `
                   -Surname $user.Etternavn `
                   -UserPrincipalName $userPrincipalName `
                   -SamAccountName $samAccountName `
                   -Path $OU `
                   -AccountPassword $password `
                   -Enabled $true `
                   -ChangePasswordAtLogon $true `
                   -EmailAddress $userPrincipalName `
                   -Description "Department: $($user.Avdeling)" `
                   -OfficePhone $phoneNumber `
                   -Initials $initials
                   
        Write-Host "Created user: $userPrincipalName with SamAccountName: $samAccountName in OU: $OU"
        
        # Add user to the correct AD group based on the department
        switch ($user.Avdeling) {
            "Administrasjon" { Add-ADGroupMember -Identity "Administrasjon" -Members $samAccountName }
            "Salg" { Add-ADGroupMember -Identity "Salg" -Members $samAccountName }
            "Utvikling" { Add-ADGroupMember -Identity "Utvikling" -Members $samAccountName }
            "Kundesupport" { Add-ADGroupMember -Identity "Kundesupport" -Members $samAccountName }
            "HR" { Add-ADGroupMember -Identity "HR" -Members $samAccountName }
            "IT" { Add-ADGroupMember -Identity "IT" -Members $samAccountName }

            default { Write-Host "No specific group assignment for: $($user.Avdeling)" }
        }

        Write-Host "Assigned $userPrincipalName to the group: $($user.Avdeling)"

    } catch {
        Write-Host "Error creating user: $($user.Fornavn) $($user.Etternavn). Error: $_"
    }
}

Write-Host "User creation process completed."
