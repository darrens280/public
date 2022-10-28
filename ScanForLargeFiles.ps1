#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Scans a selected folder/s for large files, and outputs to CSV
.DESCRIPTION
    User selects a folder (or multiple folders) from the C:\ which script will then scan for files larger than a minimum size (5 MB is default minimum size), 
    and then outputs the results to a CSV file in the C:\TEMP folder, following which the output file will be opened automatically on completion of the script.
.EXAMPLE
    .\ScanForLargeFiles.ps1 -minimumFileSizeInMB "10" -outputFolder "c:\myTempFolder" -outputFileName "LargeFiles_Report"
.PARAMETER outputFolder
    This is the folder on the C: where the output CSV file will be stored. Defaults to C:\TEMP if not supplied at runtime.
.PARAMETER outputFileName
    This is the name of the output file, that gets stored in the output folder. Defaults to LargeFiles_Report_<YYYYMMDD_HHMMSS>.csv if not supplied at runtime.
.PARAMETER minimumFileSizeInMB
    This is the minimum file size that will be scanned for when the script is run. Defaults to 5MB if not supplied at runtime.

#>

##########################################################################################################################
# PARAMETERS

param(
    [Parameter(Mandatory=$false)][string]
    $outputFolder = "C:\TEMP",

    [Parameter(Mandatory=$false)][int]
    $minimumFileSizeInMB = 5,

    [Parameter(Mandatory=$false)][string]
    $outputFileName = "LargeFiles_Report"
)

##########################################################################################################################
# SCRIPT BODY

$settingBefore = Get-ExecutionPolicy -Scope LocalMachine
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine

$ErrorActionPreference = 'SilentlyContinue'
$dateAndTime           = Get-Date -Format yyyyMMdd_HHmmss
$minimumFileSize       = $minimumFileSizeInMB*1024*1024
$outputFile            = "$($outputFolder)\$($outputFileName)_$($dateAndTime).csv"
$targetFolders         = (Get-ChildItem -Path "C:\" | Select-Object FullName | Out-GridView -Title "Select folder to scan... and click OK" -PassThru).FullName

if (!(Test-Path $outputFolder)) {New-Item -Path $outputFolder -ItemType Directory | Out-Null}

Write-Output "--> Scanning for files larger than $($minimumFileSize / 1MB) MB. Please wait..."

foreach ($folder in $targetFolders) {
    Get-ChildItem -Path $folder -Recurse `
        | Where-Object { !$_.PSIsContainer -and $_.Length -gt $minimumFileSize } `
        | Select-Object -Property FullName,@{Name='SizeMB';Expression={[System.Math]::Round($_.Length / 1MB,2)}} `
        | Sort-Object { $_.SizeMB } -Descending `
        | Export-Csv $outputFile -NoTypeInformation -Force -Append
}

Write-Output "--> Done. Opening output file: $($outputFile)"

if (Test-Path $outputFile) { Start-Process $outputFile } `
else {throw "Output file does not exist."}

Set-ExecutionPolicy -ExecutionPolicy $settingBefore -Scope LocalMachine
$ErrorActionPreference='Continue'