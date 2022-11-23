#Requires -RunAsAdministrator
#######################################################################################################################################
# FUNCTIONS

function Install-NugetPackageProvider {
    # Install Nuget Package Provider (required for installing Modules)
    # Open Powershell (as Admin)
    # https://stackoverflow.com/questions/51406685/powershell-how-do-i-install-the-nuget-provider-for-powershell-on-a-unconnected
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Verbose -Force
}

function Install-PowerShell {
    <#
        Install PowerShell 7.x.x
        https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.2
    #>

    param(
        [Parameter(Mandatory=$false)]
        [string]$version = "7.3.0",

        [Parameter(Mandatory=$false)]
        [string]$msiFileName = "PowerShell-$($version)-win-x64.msi"
    )

    $downloadURL        = "https://github.com/PowerShell/PowerShell/releases/download/v$($version)/$($msiFileName)"
    $downloadedFilePath = "$env:TEMP\$msiFileName"

    # Download file
    #(New-Object System.Net.WebClient).DownloadFile($downloadURL, "$downloadedFilePath")
    Write-Output "--> Downloading..."
    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -Uri $downloadURL -OutFile $downloadedFilePath -UseBasicParsing
    $ProgressPreference = "Continue"

    # Install the downloaded file
    # #msiexec.exe /package $msiFileName /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1
    if (Test-Path -Path $downloadedFilePath) {
        Write-Output "--> Installing: $($msiFileName)..."
        $msiExitCode = (Start-Process -FilePath $downloadedFilePath -ArgumentList '/quiet' -Wait -PassThru).ExitCode 
        if ($msiExitCode -ne 0) {
            Write-Output "Installer returned exit code $msiExitCode"
            throw "Installation aborted"
        }
        Remove-Item $downloadedFilePath -Force | Out-Null
    }
    else {
        throw "Downloaded file '$($downloadedFilePath)' not found."
    }
}

function Install-AzCLI {
    $downloadURL        = "https://aka.ms/installazurecliwindows"
    $downloadedFilePath = "$env:TEMP\AzureCLI.msi"

    Write-Output "--> Downloading..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $downloadURL -OutFile $downloadedFilePath -UseBasicParsing
    $ProgressPreference = 'Continue'

    if (Test-Path -Path $downloadedFilePath) {
        Write-Output "--> Installing..."
        Start-Process msiexec.exe -ArgumentList "/I $downloadedFilePath /quiet" -Wait
        Remove-Item $downloadedFilePath -Force | Out-Null
    }
    else {
        throw "Downloaded file '$($downloadedFilePath)' not found."
    }
}

function Install-Chrome {
    # https://www.snel.com/support/install-chrome-in-windows-server/
    $LocalTempDir = $env:TEMP
    $ChromeInstaller = "ChromeInstaller.exe"
    (New-Object System.Net.WebClient).DownloadFile('http://dl.google.com/chrome/install/375.126/chrome_installer.exe', "$LocalTempDir\$ChromeInstaller")
    & "$LocalTempDir\$ChromeInstaller" /silent /install
    $Process2Monitor = "ChromeInstaller"
    Do { $ProcessesFound = Get-Process | ?{$Process2Monitor -contains $_.Name} | Select-Object -ExpandProperty Name; 
    If ($ProcessesFound) { "Still running: $($ProcessesFound -join ', ')" | Write-Host; Start-Sleep -Seconds 2 } else { rm "$LocalTempDir\$ChromeInstaller" -ErrorAction SilentlyContinue -Verbose } } Until (!$ProcessesFound)

}

Function Get-FilesFromAzure {
    param
    (  
      [Parameter(Mandatory=$true)][string]$storageAccountName,
      [Parameter(Mandatory=$true)][string]$storageFileShareName,
      [Parameter(Mandatory=$true)][string]$storageAccessKey,
      [Parameter(Mandatory=$false)][string]$fileNameToDownload,
      [Parameter(Mandatory=$false)][string]$downloadPath = $env:TEMP
    )
  
    $loggedIn = Get-AzContext
    if ($loggedIn) {
  
      Write-Output "--> Downloading files from Azure Storage..."
      Write-Output "--> Subscription            = '$(($loggedIn).Subscription.Name)'"
      Write-Output "--> Storage Account Name    = '$($storageAccountName)'"
      Write-Output "--> Storage File Share Name = '$($storageFileShareName)'"
  
      $storageHostName = $storageAccountName + ".file.core.windows.net"
      
      Invoke-Expression -Command ("cmdkey /add:$storageHostName /user:AZURE\$storageAccountName /pass:$storageAccessKey")
  
      # Mapping Drive
      $password = ConvertTo-SecureString -String $storageAccessKey -AsPlainText -Force
      $credential = New-Object System.Management.Automation.PSCredential -ArgumentList "AZURE\$storageAccountName", $password
      $root = ("\\" + $storageHostName + "\" + $storageFileShareName)
  
      Write-Output "--> Full source path       = '$($root)'"
      New-PSDrive -Name "U" -PSProvider FileSystem -Root $root -Credential $credential | Out-Null
  
      if (!(Test-Path $downloadPath)) { New-Item -ItemType Directory -Path $downloadPath | Out-Null }
  
      if ($fileNameToDownload) {
        Write-Output "--> Downloading: '$($fileNameToDownload)' to '$($downloadPath)'..."
        Copy-Item -Path "U:\$fileNameToDownload" -Destination $downloadPath -Recurse -Force
      }
      else {
        Write-Output "--> Downloading all files from share '$($storageFileShareName)' to: '$($downloadPath)'..."
        Copy-Item -Path "U:\*" -Destination $downloadPath -Recurse -Force
      }
    }
    else {
      throw "You need to be logged in to Azure before downloading files. Use Connect-AzAccount to login."
    }
}

function Upload-FileToAzureStorage {
    param (
        [Parameter(Mandatory=$true)][string]$storageAccountName,
        [Parameter(Mandatory=$true)][string]$storageAccountResourceGroup,
        [Parameter(Mandatory=$true)][string]$storageFileShareName,
        [Parameter(Mandatory=$true)][string]$fileName,
        [Parameter(Mandatory=$false)][string]$folder

    )
    
    $loggedIn = Get-AzContext
    if ($loggedIn) {
        Write-Output "--> Uploading $($fileName) to Azure Storage..."
        $storageAccessKey = (Get-AzStorageAccountKey -ResourceGroupName $storageAccountResourceGroup -AccountName $storageAccountName)[0].value
        $storageContext   = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccessKey
        
        if (!$folder) { 
            Set-AzStorageFileContent -Context $storageContext -ShareName $storageFileShareName -Source $fileName -force 
        }
        else { 
            Set-AzStorageFileContent -Context $storageContext -ShareName $storageFileShareName -Source $fileName -force -Path $folder 
        }
    }
    else {
        throw "You need to be logged in to Azure before downloading files. Use Connect-AzAccount to login."
    }
}

#######################################################################################################################################

# Upload Files
Connect-AzAccount
Upload-FileToAzureStorage -storageAccountName $storageAccountName -storageAccountResourceGroup $storageAccountResourceGroup -storageFileShareName $storageFileShareName -fileName "C:\Temp\MicrosoftEdgeEnterpriseX64.msi"

# Download Files
Connect-AzAccount
$storageAccountName          = "aaa"
$storageAccountResourceGroup = "bbb"
$storageFileShareName        = "ccc"
$fileName                    = "ddd.exe"
$storageAccessKey            = (Get-AzStorageAccountKey -ResourceGroupName $storageAccountResourceGroup -AccountName $storageAccountName)[0].value

Get-FilesFromAzure -storageAccountName $storageAccountName -storageFileShareName $storageFileShareName -storageAccessKey $storageAccessKey -fileNameToDownload $fileName