<# Enable RPC over Named Pipes for printing via Local GPO (Registry-backed)
   - Outgoing: use RPC over Named Pipes
   - Incoming: allow RPC over Named Pipes and TCP
   Run as Administrator
#>

# Admin check
$admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $admin) { Write-Error "Run this script in an elevated PowerShell."; exit 1 }

$base = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
New-Item -Path $base -Force | Out-Null

# Outgoing RPC: named pipes
New-ItemProperty -Path $base -Name 'RpcUseNamedPipeProtocol' -PropertyType DWord -Value 1 -Force | Out-Null

# Listener protocols: named pipes + TCP (0x7)
New-ItemProperty -Path $base -Name 'RpcProtocols' -PropertyType DWord -Value 0x7 -Force | Out-Null

# Optional: enforce Kerberos for listener (domain environments). Comment out if not needed.
# New-ItemProperty -Path $base -Name 'ForceKerberosForRpc' -PropertyType DWord -Value 1 -Force | Out-Null

# Apply
Write-Host "Restarting Print Spooler..."
Restart-Service -Name Spooler -Force

# Refresh local policy application
Start-Process -FilePath "gpupdate.exe" -ArgumentList "/target:computer /force" -Wait

# Show resulting values
Get-ItemProperty -Path $base | Select-Object RpcUseNamedPipeProtocol, RpcProtocols
Write-Host "Done. If issues persist, reboot is safe."
