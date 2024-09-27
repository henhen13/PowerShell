Welcome to the MA Chang Industries PowerShell Scripts repository! This repository contains several PowerShell scripts designed to manage user accounts, generate reports, and perform remote maintenance tasks on Windows Servers.

## Contents

1. **User Creation Script**: 
   - **File**: `opprettbrukere.ps1`
   - **Description**: This script reads user data from an Excel file, creates new users in Active Directory, assigns them to appropriate groups, sets initial passwords, and enforces password changes at the next login. It also handles nested groups for the Sales and IT departments.
   
2. **User Report Generation Script**: 
   - **File**: `rapport.ps1`
   - **Description**: This script generates a report of all users, capturing their name, department, phone number, email address, and last login time. The data is exported to a CSV file for easy access and analysis.

3. **Remote Maintenance Script**: 
   - **File**: `mtc.ps1`
   - **Description**: This script performs maintenance tasks on multiple Windows Servers, including checking the status of the W32Time service, installing updates if the service is running, and scheduling reboots as needed. It also sets up a scheduled task to run weekly.
  
4. **CSV-file with DHCP-leases**: 
   - **File**: `dhcpsave.ps1`
   - **Description**: This script outputs a csv file with retrived dhcp-leases and write the selected properties in the code(you can add more if you'd like) 

## Prerequisites

- PowerShell version 5.1 or higher
- Active Directory module for Windows PowerShell
- ImportExcel module for reading Excel files
- PSWindowsUpdate module for managing Windows updates

## How to Use

1. **User Creation**:
   - Update the path to the Excel file in the `Create-Users.ps1` script.
   - Run the script to create users in Active Directory.

2. **User Report Generation**:
   - Update the output file path in the `Generate-UserReport.ps1` script.
   - Run the script to generate a user report in CSV format.

3. **Remote Maintenance**:
   - Update the list of servers in the `Remote-Maintenance.ps1` script.
   - Run the script to check service status, install updates, and schedule reboots.
