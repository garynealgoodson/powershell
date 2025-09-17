# ================================================================================================
# PurposeGet Windows Autopilot HWID and store in NinjaOne Custom Field
# ===============================================================================================
# Prerequisites: New Custom Field & Add Script to NinjaOne Automations
# ================================================================================================
# 1. Add a new custom field in NinjaOne:
#   Type: Multi-line text
#   Name: autopilotHwid
#   Label: Autopilot HWID
#   Custom field is required: No
#   Inheritance: 
#        Device: Yes
#   Permissions:
#       Automations: Write only
#       Technician access: Read only
# 2. Add the script into NinjaOne Automations 
#  Name: Get Windows Autopilot HWID
#  Description: Get Windows Autopilot HWID and store in NinjaOne Custom Field
#  Language: PowerShell
#  Operating System: Windows
#  Architecture: x64
#  Run as: System
# ================================================================================================
# Generate AutopilotHWID.csv
# ================================================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$hwidPath   = "C:\HWID"
$outputFile = Join-Path $hwidPath "AutopilotHWID.csv"
$installDir = "C:\ProgramData\MSPScripts"

# Ensure working directories exist
New-Item -Type Directory -Path $hwidPath   -Force | Out-Null
New-Item -Type Directory -Path $installDir -Force | Out-Null

# Make sure Get-WindowsAutopilotInfo is saved locally
$scriptPath = Join-Path $installDir "Get-WindowsAutopilotInfo.ps1"
if (-not (Test-Path $scriptPath)) {
    Save-Script -Name Get-WindowsAutopilotInfo -Path $installDir -Force
}

# Run the script by full path
& $scriptPath -OutputFile $outputFile

# ================================
# Read CSV data
# ================================
$data      = Import-Csv -Path $outputFile
$serial    = $data.'Device Serial Number'
$productId = $data.'Windows Product ID'
$hash      = $data.'Hardware Hash'

# Escape quotes for SQL safety
$serial     = $serial    -replace "'", "''"
$productId  = $productId -replace "'", "''"
$hash       = $hash      -replace "'", "''"

# ================================================================================================
# Combined CSV-style output variable and output to custom field
# ================================================================================================
$autopilotLine = "$serial,$productId,$hash"

Ninja-Property-Set autopilotHwid $autopilotLine

# ================================================================================================
# Cleanup HWID folder
# ================================================================================================
if (Test-Path $hwidPath) {
    Remove-Item -Path $hwidPath -Recurse -Force
}