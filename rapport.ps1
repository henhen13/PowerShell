# Ensure the Active Directory module is imported
Import-Module ActiveDirectory

# Define the output file path
$outputFilePath = "C:\Users\Administrator\Documents\UserReport.csv"

# Initialize an array to store user information
$userReport = @()

# Retrieve all users from Active Directory
try {
    $users = Get-ADUser -Filter * -Property Name, Department, TelephoneNumber, EmailAddress, LastLogonDate

    foreach ($user in $users) {
        # Create a custom object for each user
        $userInfo = [PSCustomObject]@{
            Name          = $user.Name
            Department    = $user.Department
            PhoneNumber   = $user.TelephoneNumber
            EmailAddress  = $user.EmailAddress
            LastLoginTime = $user.LastLogonDate
        }

        # Add the user information to the report array
        $userReport += $userInfo
    }
} catch {
    Write-Host "Error retrieving users: $_"
}

# Export the collected data to a CSV file
try {
    $userReport | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "User report exported successfully to $outputFilePath"
} catch {
    Write-Host "Error exporting report: $_"
}
