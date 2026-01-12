# Prompt for required inputs
$domainName = Read-Host "Enter the domain name (e.g., ad.xyz.edu)"
$adminUser = Read-Host "Enter the domain admin username (e.g., AdminUsername)"
$newHostname = Read-Host "Enter new computer name (leave blank to keep current name)"

# Get secure password input
$securePassword = Read-Host "Enter domain admin password" -AsSecureString

# Construct full domain user
$domainUser = "$domainName\$adminUser"

# Create PSCredential object
$credential = New-Object System.Management.Automation.PSCredential ($domainUser, $securePassword)

# Rename computer if new hostname provided
if (-not [string]::IsNullOrWhiteSpace($newHostname) -and $newHostname -ne $env:COMPUTERNAME) {
    try {
        Rename-Computer -NewName $newHostname -Force
        Write-Host "Renamed computer to $newHostname."
        Start-Sleep -Seconds 60  # Wait for rename restart
    } catch {
        Write-Host "Failed to rename computer: $_"
        exit 1
    }
}

# Use updated computer name
$computerName = if ([string]::IsNullOrWhiteSpace($newHostname)) { $env:COMPUTERNAME } else { $newHostname }

# Join computer to domain
Write-Host "Joining $computerName to $domainName..."
try {
    Add-Computer -ComputerName $computerName -DomainName $domainName -Credential $credential
    Write-Host "Successfully joined $computerName to $domainName. Please restart your PC"
    Start-Sleep -Seconds 60 #wait for join
} catch {
    Write-Host "Error joining the computer to the domain: $_"
    exit 1
}

Write-Host "Process complete. $computerName is now joined to $domainName."
