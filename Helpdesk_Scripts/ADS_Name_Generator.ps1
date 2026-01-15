#This SAcript is designed to connect to a csv spreadsheet that reads the entry of the firstr textbox you fill out and compares it to column A of the spreadsheet looking for a match. 
#If there's a match then it will store the content of column B on the same row where the match was found and store that as the $Org variable.
#It will do that same thing with column D and E respectively to create the $Dept variable
#Finaly it reads your serial number and grabs the last 6 digits and generates a hostname for your Organization.
#
# Path to your CSV file (change as needed or make it a parameter)
$csvFile = "D:\User\FilePath"

# Check if the CSV file exists
if (-not (Test-Path $csvFile -PathType Leaf)) {
    Write-Host "Error: CSV file not found at '$csvFile'" -ForegroundColor Red
    Write-Host "Please make sure the file exists and the path is correct." -ForegroundColor Yellow
    exit 1
}

# Try to import the CSV with error handling
try {
    # Assuming the CSV has NO headers → we assign our own
    $csvData = Import-Csv -Path $csvFile -Header A,B,C,D,E -ErrorAction Stop
    
    # Basic validation: make sure we actually loaded something
    if ($csvData.Count -eq 0) {
        Write-Host "Error: The CSV file is empty." -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host "Error reading CSV file:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "Possible causes:" -ForegroundColor Yellow
    Write-Host " • File is not a valid CSV" -ForegroundColor Yellow
    Write-Host " • File is open in another program" -ForegroundColor Yellow
    Write-Host " • File is corrupted or has invalid encoding" -ForegroundColor Yellow
    exit 1
}

Write-Host "CSV loaded successfully ($($csvData.Count) rows)" -ForegroundColor Green

# ────────────────────────────────────────────────
# Get ORG#
# ────────────────────────────────────────────────
$orgInput = Read-Host "Please enter the ORG#"

$orgRow = $csvData | Where-Object { $_.A -eq $orgInput } | Select-Object -First 1

if (-not $orgRow) {
    Write-Host "Error: ORG# '$orgInput' not found in column A" -ForegroundColor Red
    exit 1
}

$Org = $orgRow.B
if ([string]::IsNullOrWhiteSpace($Org)) {
    Write-Host "Warning: Column B is empty for ORG# '$orgInput'" -ForegroundColor Yellow
    $Org = "UNKNOWN_ORG"
}

# ────────────────────────────────────────────────
# Get Dept#
# ────────────────────────────────────────────────
$deptInput = Read-Host "Please enter the Dept# - ex.) D01099"

$deptRow = $csvData | Where-Object { $_.D -eq $deptInput } | Select-Object -First 1

if (-not $deptRow) {
    Write-Host "Error: Dept# '$deptInput' not found in column D" -ForegroundColor Red
    exit 1
}

$Dept = $deptRow.E
if ([string]::IsNullOrWhiteSpace($Dept)) {
    Write-Host "Warning: Column E is empty for Dept# '$deptInput'" -ForegroundColor Yellow
    $Dept = "UNKNOWN_DEPT"
}

# ────────────────────────────────────────────────
# Get last 6 digits of serial number
# ────────────────────────────────────────────────
try {
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $serialFull = $bios.SerialNumber

    if ([string]::IsNullOrWhiteSpace($serialFull)) {
        Write-Host "Error: Could not retrieve BIOS serial number (empty)" -ForegroundColor Red
        exit 1
    }

    if ($serialFull.Length -lt 6) {
        Write-Host "Warning: Serial number is only $($serialFull.Length) characters long" -ForegroundColor Yellow
        $Serial = $serialFull.PadLeft(6, '0')  # or you could use the full value
    } else {
        $Serial = $serialFull.Substring($serialFull.Length - 6)
    }
}
catch {
    Write-Host "Error retrieving BIOS serial number:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# ────────────────────────────────────────────────
# Build and display final string
# ────────────────────────────────────────────────
$finalString = "$Org - $Dept - $Serial"

Write-Host "`nResult:" -ForegroundColor Cyan
Write-Host $finalString -ForegroundColor White -BackgroundColor DarkBlue

Write-Host ""
