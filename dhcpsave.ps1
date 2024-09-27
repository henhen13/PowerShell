# Define the output file path
$outputFilePath = "C:\Windows\Temp\DHCPLeases.csv"

# Retrieve DHCP leases
$dhcpLeases = Get-DhcpServerv4Lease -ScopeId 172.16.11.0 -ComputerName HENHEN-DC01

# Select required properties and create custom objects
$leaseData = $dhcpLeases | Select-Object ScopeId, IPAddress, HostName, ClientID, AddressState

# Export to CSV with semicolon delimiter
$leaseData | Export-Csv -Path $outputFilePath -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Host "DHCP leases have been exported to $outputFilePath"
