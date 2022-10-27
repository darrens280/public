#######################################################################################################################################
# FUNCTIONS

function Install-NugetPackageProvider {
    # Install Nuget Package Provider (required for installing Modules)
    # Open Powershell (as Admin)
    # https://stackoverflow.com/questions/51406685/powershell-how-do-i-install-the-nuget-provider-for-powershell-on-a-unconnected
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Verbose
}

function Install-PowerShell {
    <#
        Install PowerShell 7.x.x
        https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.2
    #>

    param(
        [Parameter(Mandatory=$false)]
        [string]$version = "7.2.7",

        [Parameter(Mandatory=$false)]
        [string]$msiFileName = "PowerShell-$($version)-win-x64.msi"
    )

    $downloadURL        = "https://github.com/PowerShell/PowerShell/releases/download/v$($version)/$($msiFileName)"
    $LocalTempDir       = $env:TEMP
    $downloadedFilePath = "$LocalTempDir\$msiFileName"

    # Download file
    #(New-Object System.Net.WebClient).DownloadFile($downloadURL, "$LocalTempDir\$msiFileName")
    Invoke-WebRequest -Uri $downloadURL -OutFile $downloadedFilePath -UseBasicParsing

    # Install the downloaded file
    # #msiexec.exe /package $msiFileName /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1
    if (!(Test-Path -Path $downloadedFilePath)) {
        Write-Output "--> Installing: $($msiFileName)..."
        $msiExitCode = (Start-Process -FilePath $downloadedFilePath -ArgumentList '/quiet' -Wait -PassThru).ExitCode 
        if ($msiExitCode -ne 0) {
            Write-Output "Installer returned exit code $msiExitCode"
            throw "Installation aborted"
        }
    }
    else {
        throw "Downloaded file not found."
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
#######################################################################################################################################