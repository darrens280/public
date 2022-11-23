<#
.SYNOPSIS
  Get-EdgeEnterpriseMSI

.DESCRIPTION
  Imports all device configurations in a folder to a specified tenant

.PARAMETER Channel
  Channel to download, Valid Options are: Dev, Stable, EdgeUpdate, Policy.

.PARAMETER Platform
  Platform to download, Valid Options are: Windows or MacOS, if using channel "Policy" this should be set to "any"
  Defaults to Windows if not set.

.PARAMETER Architecture
  Architecture to download, Valid Options are: x86, x64, arm64, if using channel "Policy" this should be set to "any"
  Defaults to x64 if not set.

.PARAMETER Version
  If set the script will try and download a specific version. If not set it will download the latest.

.PARAMETER Folder
  Specifies the Download folder

.NOTES
  https://www.deploymentresearch.com/using-powershell-to-download-edge-chromium-for-business/

  https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/deploy-edge

.EXAMPLE
  
  Download the latest version for the Stable channel and overwrite any existing file
  .\Get-EdgeEnterpriseMSI.ps1 -Channel Stable -Folder D:\SourceCode\PowerShell\Div -Force

#>

############################################################################################################################################
# PARAMETERS

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false, HelpMessage = 'Channel to download, Valid Options are: Dev, Stable, EdgeUpdate, Policy')]
  [ValidateSet('Dev', 'Stable', 'EdgeUpdate', 'Policy')]
  [string]$Channel = "Stable",
  
  [Parameter(Mandatory = $false, HelpMessage = 'Folder where the file will be downloaded')]
  [ValidateNotNullOrEmpty()]
  [string]$TargetDownloadFolder = $ENV:TEMP,

  [Parameter(Mandatory = $false, HelpMessage = 'Platform to download, Valid Options are: Windows or MacOS')]
  [ValidateSet('Windows', 'MacOS', 'any')]
  [string]$Platform = "Windows",

  [Parameter(Mandatory = $false, HelpMessage = "Architecture to download, Valid Options are: x86, x64, arm64, any")]
  [ValidateSet('x86', 'x64', 'arm64', 'any')]
  [string]$Architecture = "x64",

  [parameter(Mandatory = $false, HelpMessage = "Specifies which version to download")]
  [ValidateNotNullOrEmpty()]
  [string]$ProductVersion

)

############################################################################################################################################
# FUNCTIONS

function Set-Parameters {
  # Validating parameters to reduce user errors
  if ($Channel -eq "Policy" -and ($Architecture -ne "Any" -or $Platform -ne "Any")) {
    Write-Warning ("Channel 'Policy' requested, but either 'Architecture' and/or 'Platform' is not set to 'Any'. Setting Architecture and Platform to 'Any'")

    $global:Architecture = "Any"
    $global:Platform     = "Any"
  } 
  elseif ($Channel -ne "Policy" -and ($Architecture -eq "Any" -or $Platform -eq "Any")) {
    throw "If Channel isn't set to Policy, architecture and/or platform can't be set to 'Any'"
  }
  elseif ($Channel -eq "EdgeUpdate" -and ($Architecture -ne "x86" -or $Platform -eq "Windows")) {
    Write-Warning ("Channel 'EdgeUpdate' requested, but either 'Architecture' is not set to x86 and/or 'Platform' is not set to 'Windows'. Setting Architecture to 'x86' and Platform to 'Windows'")

    $global:Architecture = "x86"
    $global:Platform     = "Windows"
  }
}

function Set-TLSSecurity {
  Write-Verbose "--> Enabling connection over TLS for better compability on servers..."
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

  # Test if HTTP status code 200 is returned from URI
  try {
    Invoke-WebRequest $edgeEnterpriseMSIUri -UseBasicParsing | Where-Object StatusCode -match 200 | Out-Null
  }
  catch {
    throw "Unable to get HTTP status code 200 from '$($edgeEnterpriseMSIUri)'. Does the URL still exist?"
  }
}

function Get-AllAvailableDownloads {
  Write-Verbose "--> Getting available files from: '$($edgeEnterpriseMSIUri)'"
  # Try to get JSON data from Microsoft
  try {
    $response = Invoke-WebRequest -Uri $edgeEnterpriseMSIUri -Method Get -ContentType "application/json" -UseBasicParsing -ErrorVariable InvokeWebRequestError
    $jsonObj  = ConvertFrom-Json $([String]::new($response.Content))
    Write-Verbose "--> Successfully retrieved data"
  }
  catch {
    throw "Could not get MSI data: $InvokeWebRequestError"
  }

  # Alternative is to use Invoke-RestMethod to get a Json object directly
  # $jsonObj = Invoke-RestMethod -Uri "https://edgeupdates.microsoft.com/api/products?view=enterprise" -UseBasicParsing

  $selectedIndex = [array]::indexof($jsonObj.Product, "$Channel")

  if (-not $ProductVersion) {
    try {
      Write-Verbose "--> No version specified, therefore getting the latest for: '$($Channel)'"
      $selectedVersion = (([Version[]](($jsonObj[$selectedIndex].Releases | Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform }).ProductVersion) | Sort-Object -Descending)[0]).ToString(4)
    
      Write-Output "--> Latest $($Channel) version is: '$($selectedVersion)'"
      $global:selectedObject = $jsonObj[$selectedIndex].Releases | Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform -and $_.ProductVersion -eq $selectedVersion }
    }
    catch {
      throw "Unable to get object from Microsoft. Check your parameters and refer to script help."
    }
  }
  else {
    Write-Output "--> Matching $ProductVersion on channel $Channel"
    $global:selectedObject = ($jsonObj[$selectedIndex].Releases | Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform -and $_.ProductVersion -eq $ProductVersion })

    if (-not $selectedObject) {
      throw "No version matching $ProductVersion found in $channel channel for $Architecture architecture."
    }
    else {
      Write-Output "--> Found matching version`n"
    }
  }
}

function Get-Download {
  Set-Parameters
  Set-TLSSecurity
  Get-AllAvailableDownloads

  if (Test-Path $TargetDownloadFolder) {
    foreach ($artifacts in $selectedObject.Artifacts) {
      # Not showing the progress bar in Invoke-WebRequest is quite a bit faster than default
      $ProgressPreference = 'SilentlyContinue'
      
      Write-Verbose "--> Starting download of: $($artifacts.Location)"
      Write-Output "--> Downloading to: '$($TargetDownloadFolder)'"
      # Work out file name
      $global:fileName = Split-Path $artifacts.Location -Leaf
      $global:downloadedFilePath = "$TargetDownloadFolder\$fileName"

      if (!(Test-Path $downloadedFilePath -ErrorAction SilentlyContinue)) {

        Write-Output "--> Starting download..."
        try {
          Invoke-WebRequest -Uri $artifacts.Location -OutFile $downloadedFilePath -UseBasicParsing
        }
        catch {
          throw "Attempted to download file, but failed: $error[0]"
        }
      }
      if (((Get-FileHash -Algorithm $artifacts.HashAlgorithm -Path $downloadedFilePath).Hash) -eq $artifacts.Hash) {
        Write-Verbose "--> Checksum verified"
      }
      else {
        Write-Warning "Checksum mismatch!"
        Write-Warning "Expected Hash: $($artifacts.Hash)"
        Write-Warning "Downloaded file Hash: $((Get-FileHash -Algorithm $($artifacts.HashAlgorithm) -Path $downloadedFilePath).Hash)`n"
      }
    }
  }
  else {
    throw "Folder $TargetDownloadFolder does not exist"
  }
}

function Install-MSEdge {

  Write-Output "--> Installing: $($fileName)..."
  $msiExitCode    = (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $downloadedFilePath /qn /l* $TargetDownloadFolder\msedge_msi_install.log" -Wait -PassThru ).ExitCode
  if ($msiExitCode -ne 0) {
      Write-Output "ERROR - $filename installer returned exit code $msiExitCode"
      throw "Installation aborted"
  }
  else {
    Get-Service "edgeupdate" | Set-Service -StartupType Manual -ErrorAction SilentlyContinue | Out-Null
    Get-Service "edgeupdatem" | Set-Service -StartupType Manual -ErrorAction SilentlyContinue | Out-Null
    Get-ScheduledTask | Where-Object {$_.TaskName -like "MicrosoftEdgeUpdate*"} | Unregister-ScheduledTask -Confirm:$false -ea 0

    Remove-Item $downloadedFilePath -Force -ea 0 | Out-Null
    Write-Output "--> Done"
  }
}

function Set-DesktopShortcut {
    # Create Desktop Shortcut for All Users
    $TargetFile = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if (Test-Path $TargetFile) {
        $ShortcutFile = "$env:Public\Desktop\Microsoft Edge.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
    }
}

############################################################################################################################################
# VARIABLES

$ErrorActionPreference = "Stop"
$edgeEnterpriseMSIUri  = 'https://edgeupdates.microsoft.com/api/products?view=enterprise'

############################################################################################################################################
# SCRIPT BODY

try {

  Get-Download
  Install-MSEdge
  Set-DesktopShortcut

}
catch {
  throw "Error installing Microsoft Edge"
}

Write-Output "--> Done"