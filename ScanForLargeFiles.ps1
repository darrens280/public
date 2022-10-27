<#
.SYNOPSIS
    Scans a specific folder for large files, and outputs to CSV
.DESCRIPTION
    User selects a folder from the C:\ which script will then scan files larger than 5MB (default minimum size), and then outputs the results to a CSV file,
    in the C:\TEMP folder, following which the output file will be opened automatically on completion of the script.
.EXAMPLE
    .\ScanForLargeFiles.ps1
.PARAMETER outputFolder
    This is the folder on the C: where the output CSV file will be stored. Defaults to C:\TEMP if not supplied at runtime.
.PARAMETER outputFileName
    This is the name of the output file, that gets stored in the output folder. Defaults to LargeFiles_Report.csv if not supplied at runtime.
.PARAMETER minimumFileSize
    This is the minimum file size that will be scanned for when the script is run. Defaults to 5MB if not supplied at runtime.

#>


param(
    [Parameter(Mandatory=$false)][string]
    $outputFolder = "C:\TEMP",

    [Parameter(Mandatory=$false)][int]
    $minimumFileSize = 5*1024*1024,

    [Parameter(Mandatory=$false)][string]
    $outputFileName = "LargeFiles_Report"
)

$settingBefore = Get-ExecutionPolicy -Scope LocalMachine
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine

$ErrorActionPreference = 'SilentlyContinue'
$dateAndTime           = Get-Date -Format yyyyMMdd_HHmmss
$outputFile            = "$($outputFolder)\$($outputFileName)_$($dateAndTime).csv"
$targetFolder          = (Get-ChildItem -Path "C:\" | Select-Object FullName | Out-GridView -Title "Select folder to scan... and click OK" -PassThru).FullName

if (!(Test-Path $outputFolder)) {New-Item -Path $outputFolder -ItemType Directory | Out-Null}

Write-Output "--> Scanning $($targetFolder) for files larger than $($minimumFileSize / 1MB) MB. Please wait..."

Get-ChildItem -Path $targetFolder -Recurse `
    | Where-Object { !$_.PSIsContainer -and $_.Length -gt $minimumFileSize } `
    | Select-Object -Property FullName,@{Name='SizeMB';Expression={[System.Math]::Round($_.Length / 1MB,2)}} `
    | Sort-Object { $_.SizeMB } -Descending `
    | Export-Csv $outputFile -NoTypeInformation -Force

Write-Output "--> Done. Opening output file: $($outputFile)"

if (Test-Path $outputFile) { Start-Process $outputFile } `
else {throw "Output file does not exist."}

Set-ExecutionPolicy -ExecutionPolicy $settingBefore -Scope LocalMachine
$ErrorActionPreference='Continue'