# Import necessary modules
Import-Module ActiveDirectory

# Ensure PSWindowsUpdate module is installed
try {
    Import-Module PSWindowsUpdate -ErrorAction Stop
} catch {
    Write-Host "Error importing PSWindowsUpdate module: $_"
    exit
}

# Define the AD group for servers
$groupDN = "CN=Servere,OU=Grupper,OU=FSI-HENHEN,DC=fsi-henhen,DC=com"  # Update with your actual group name

# Query the members of the group to get the list of server names
$servers = Get-ADGroupMember -Identity $groupDN | Where-Object { $_.objectClass -eq 'computer' } | ForEach-Object { $_.Name }

# If no servers are found in the group, exit the script
if (-not $servers) {
    Write-Host "No servers found in the group $groupDN."
    exit
}

$updateRequired = $false

# Function to check the W32Time service and install updates
function CheckAndUpdate {
    param (
        [string]$server,
        [ref]$updateRequired
    )

    try {
        # Check the status of W32Time service
        $service = Get-Service -Name W32Time -ComputerName $server -ErrorAction Stop
        
        if ($service.Status -eq 'Running') {
            Write-Host "$($server): W32Time service is running."

            # Install software updates
            Write-Host "Installing updates on $server..."
            Invoke-Command -ComputerName $server -ScriptBlock {
                Install-WindowsUpdate -AcceptAll -AutoReboot -ErrorAction Stop
            }
            
            $updateRequired.Value = $true
        } else {
            Write-Host "$($server): W32Time service is not running."
        }
    } catch {
        Write-Host "Error processing $($server): $_"
    }
}

# Function to schedule a reboot
function ScheduleReboot {
    param (
        [string]$server,
        [DateTime]$rebootTime = $(Get-Date).AddMinutes(1)  # Default to 1 minute from now if no time is specified
    )

    try {
        # Create a scheduled task for rebooting
        Invoke-Command -ComputerName $server -ScriptBlock {
            param ($rebootTime)

            $action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /t 0"
            $trigger = New-ScheduledTaskTrigger -Once -At $rebootTime
            $taskName = "Reboot_$env:COMPUTERNAME"

            Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -User "SYSTEM" -ErrorAction Stop
            Write-Host "Scheduled reboot for $($env:COMPUTERNAME) at $($rebootTime)."
        } -ArgumentList $rebootTime
    } catch {
        Write-Host "Error scheduling reboot for $($server): $_"
    }
}

# Function to create a weekly scheduled task to run the maintenance script
function SetupWeeklyTask {
    param (
        [string]$taskName,
        [string]$scriptPath,  # Full path of the script to run
        [DateTime]$startTime = $(Get-Date).AddDays(1).Date.AddHours(2)  # Default start time at 2 AM tomorrow
    )

    try {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File `"$scriptPath`""
        $trigger = New-ScheduledTaskTrigger -Weekly -At $startTime
        
        Invoke-Command -ComputerName $servers -ScriptBlock {
            param ($taskName, $action, $trigger)

            Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -User "SYSTEM" -ErrorAction Stop
            Write-Host "Scheduled weekly maintenance task '$taskName' to run at $($trigger.StartBoundary)."
        } -ArgumentList $taskName, $action, $trigger
    } catch {
        Write-Host "Error creating weekly scheduled task '$taskName': $_"
    }
}

# Loop through each server and perform maintenance tasks
foreach ($server in $servers) {
    CheckAndUpdate -server $server -updateRequired ([ref]$updateRequired)

    if ($updateRequired) {
        # Ask if the user wants to schedule a reboot
        $response = Read-Host "Do you want to schedule a reboot for $($server)? (Y/N)"
        if ($response -eq 'Y') {
            # Ask for a scheduled time if reboot is requested
            $rebootTimeInput = Read-Host "Enter reboot time (format: 'yyyy-MM-dd HH:mm') or press Enter for immediate reboot"
            if ([string]::IsNullOrEmpty($rebootTimeInput)) {
                # Default to immediate reboot
                ScheduleReboot -server $server
            } else {
                try {
                    $rebootTime = [DateTime]::ParseExact($rebootTimeInput, "yyyy-MM-dd HH:mm", $null)
                    ScheduleReboot -server $server -rebootTime $rebootTime
                } catch {
                    Write-Host "Invalid date format. Skipping scheduled reboot for $($server)."
                }
            }
        }
    }
}

# Provide the option to schedule the script to run weekly
$scriptPath = "C:\Path\To\Your\Script.ps1"  # Replace with the full path to your script
$weeklyResponse = Read-Host "Do you want to schedule this script to run weekly? (Y/N)"
if ($weeklyResponse -eq 'Y') {
    SetupWeeklyTask -taskName "WeeklyServerMaintenance" -scriptPath $scriptPath
}

Write-Host "Remote maintenance tasks completed."
