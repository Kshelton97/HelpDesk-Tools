# Prompt for required inputs
$domainName = Read-Host "Enter the domain name (e.g., ad.domainname.edu)"
$adminUser = Read-Host "Enter the domain admin username (e.g., AdminUsername)"
$computerName = Read-Host "Enter the hostname of the computer (leave blank for current computer)"
$newHostname = Read-Host "Enter new computer name (leave blank to keep current name)"

# Get secure password input
$securePassword = Read-Host "Enter domain admin password" -AsSecureString

# Construct full domain user
$domainUser = "$domainName\$adminUser"

# Create PSCredential object
$credential = New-Object System.Management.Automation.PSCredential ($domainUser, $securePassword)

# Use current computer name if not specified
if ([string]::IsNullOrWhiteSpace($computerName)) {
    $computerName = $env:COMPUTERNAME
}

Write-Host "Preparing to remove $computerName from $domainName..."

# Remove computer from domain
try {
    Remove-Computer -Credential $credential -Force -Restart
    Write-Host "Successfully removed $computerName from $domainName. Restarting..."
} catch {
    Write-Host "Error removing the computer from the domain: $_"
    exit 1
}

# Wait for reboot to finish (adjust delay as needed)
Start-Sleep -Seconds 90

# Rename computer if new hostname provided
if (-not [string]::IsNullOrWhiteSpace($newHostname) -and $newHostname -ne $env:COMPUTERNAME) {
    try {
        Rename-Computer -NewName $newHostname -Force -Restart
        Write-Host "Renamed computer to $newHostname. Restarting..."
        Start-Sleep -Seconds 90  # Wait for rename restart
        $computerName = $newHostname  # Update variable for next step
    } catch {
        Write-Host "Failed to rename computer: $_"
        exit 1
    }
}

# Rejoin computer to domain
Write-Host "Rejoining $computerName to $domainName..."
try {
    Add-Computer -ComputerName $computerName -DomainName $domainName -Credential $credential -Restart -Force
    Write-Host "Successfully rejoined $computerName to $domainName. Restarting..."
} catch {
    Write-Host "Error rejoining the computer to the domain: $_"
    exit 1
}

Write-Host "Process complete. $computerName has been renamed and rejoined to $domainName."
