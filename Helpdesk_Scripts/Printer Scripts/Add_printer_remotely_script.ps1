<# Enable RPC over Named Pipes for printing via Local GPO (Registry-backed)
   - Outgoing: use RPC over Named Pipes
   - Incoming: allow RPC over Named Pipes and TCP
   Run as Administrator
#>
#param(
 # [string]$PrinterPath = '\\adsprint-prd-1\CNAS_OLM_2352-Canon_C5860'
#)


# Prompt user for a file path
$PrinterPath = Read-Host "Enter the printer file path"

# Optional: show what was entered
Write-Host "PrinterPath set to: $PrinterPath"

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

<# Enable Printer RPC over Named Pipes + connect to \\adsprint-prd-1\CNAS_OLM_2352-Canon_C5860
   Run in elevated PowerShell
#>

# --- Admin check ---
$admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $admin) { Write-Error "Run this script in an elevated PowerShell."; exit 1 }

# --- Configure Local GPO-backed registry for RPC over Named Pipes ---
$base = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
New-Item -Path $base -Force | Out-Null
New-ItemProperty -Path $base -Name 'RpcUseNamedPipeProtocol' -PropertyType DWord -Value 1 -Force | Out-Null   # Outgoing: Named Pipes
New-ItemProperty -Path $base -Name 'RpcProtocols'            -PropertyType DWord -Value 0x7 -Force | Out-Null   # Listener: NP + TCP

# --- Apply policy + restart spooler ---
Restart-Service -Name Spooler -Force
Start-Process gpupdate.exe -ArgumentList "/target:computer /force" -Wait

# --- Connect to shared printer ---
Write-Host "Connecting to printer $PrinterPath ..."
$connected = $false

try {
    Add-Printer -ConnectionName $PrinterPath -ErrorAction Stop
    $connected = $true
} catch {
    Write-Warning "Add-Printer failed ($($_.Exception.Message)). Falling back to PrintUIEntry..."
    $args = '/in', '/n', $PrinterPath
    $proc = Start-Process -FilePath 'rundll32.exe' -ArgumentList @('printui.dll,PrintUIEntry', $args) -PassThru -Wait -WindowStyle Hidden
    if ($proc.ExitCode -eq 0) { $connected = $true }
}

# --- Verify connection ---
try {
    $srv  = ($PrinterPath -replace '^\\\\([^\\]+)\\.*','$1')
    $share= ($PrinterPath -replace '^\\\\[^\\]+\\','')
    $wmi  = Get-CimInstance Win32_Printer -Filter "Network = TRUE" |
            Where-Object { $_.ServerName -ieq ('\\' + $srv) -and $_.ShareName -ieq $share }

    if ($connected -and $wmi) {
        Write-Host "✅ Connected to $PrinterPath as printer '$($wmi.Name)'."
    } else {
        Write-Error "❌ Could not confirm a connection to $PrinterPath."
    }
} catch {
    Write-Warning "Verification skipped ($($_.Exception.Message))."
}

# --- Show summary ---
Get-ItemProperty -Path $base | Select-Object RpcUseNamedPipeProtocol, RpcProtocols
Write-Host "Done."

Get-Printer | Select-Object -ExpandProperty Name

