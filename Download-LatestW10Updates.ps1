<#  
.NAME   
    Download-LatestW10Updates.ps1
.SYNOPSIS
    Downloads the latest cumulative updates for Windows 10 to a local folder.
.DESCRIPTION
    Downloads the latest cumulative updates for Windows 10 to a local folder on SystemDrive\Win10Cumu. 
    There will be a subfolder for each Windows 10 release, the Servicing Stack Update, and the zero day patch (ZDP).
.EXAMPLE
    Download-LatestW10Updates.ps1 [-OperatingSystem] WindowsClient|WindowsServer [-Architecture] x86|x64
.NOTES
    This script is an automation to fetch the latest cumulative Windows 10/Windows Server updates for each release. 
    It is built around the module "LatestUpdate" by Aaron Parker available in the PoweshellGallery. 
    
    The array $Windows10Builds contains the list of Windows 10 Builds, those will be translated to Windows 10
    release IDs by a switch statement for easier readablity of the download paths.

    Created by Markus Michalski, 11.04.2018
    Updated by Markus Michalski, 11.07.2018 - update LatestUpdate module to latest version (if avail.), create Servicing Stack Update folders for each release
    Updated by Markus Michalski, 21.01.2019 - Download Servicing Stack and Flash Player Updates
    Updated by Markus Michalski, 24.01.2019 - Added parameter to select update packages for Server, or Windows Client (Client is default)
    Updated by Markus Michalski, 19.02.2019 - Added parameter validate set for tab completion of valid parameter values, filter output to "Cumulative" to prevent download of Delta Updates
    Updated by Markus Michalski, 03.07.2019 - Updated parameters/cmdlets due to changes in LatestUpdate module (compatibility)
.LINK
    Link to the Module in the PowerShell Gallery
    https://www.powershellgallery.com/packages/LatestUpdate
#>

# Enable -WhatIf and -Verbose output
[CmdletBinding(SupportsShouldProcess = $True)]

Param(
    [parameter(Mandatory = $False, HelpMessage = "Select update packages for Windows Server, or Windows Client (default).")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("WindowsClient", "WindowsServer")]
    [string] $OperatingSystem = "WindowsClient",
    
    [parameter(Mandatory = $False, HelpMessage = "Select architecture (x64 is default).")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("x86", "x64")]
    [string] $Architecture = "x64"
)

switch ($OperatingSystem) {
    "WindowsClient" {
        $Windows10Builds = @(17763, 18362, 18363) ; 
        $Windows10ServerOrClientNote = "*Windows 10*" ; 
        $Windows10ServerOrClient = "WindowsClient" ; 
        $Windows10CumuPath = "$env:SystemDrive\Win10Cumu" 
    }
    "WindowsServer" {
        $Windows10Builds = @(14393, 17763) ; 
        $Windows10ServerOrClientNote = "*Windows Server*" ; 
        $Windows10ServerOrClient = "WindowsServer" ; 
        $Windows10CumuPath = "$env:SystemDrive\WinSrvCumu" 
    }
}

If (-not (Get-InstalledModule -Name LatestUpdate -ErrorAction SilentlyContinue)) {
    # Install the needed module, if missing
    Install-Module LatestUpdate -Force
}
else {
    # Update the needed module to always use the newest release
    Update-Module LatestUpdate -Force
}

foreach ($Windows10Build in $Windows10Builds) {
    switch ($Windows10Build) {
        # Translate the Windows10Buildnumber to the release for easier readability
        14393 { $Windows10Release = '1607' }
        15063 { $Windows10Release = '1703' }
        16299 { $Windows10Release = '1709' }
        17134 { $Windows10Release = '1803' }
        17763 { $Windows10Release = '1809' }
        18362 { $Windows10Release = '1903' }
        18363 { $Windows10Release = '1909' }
        Default { $Windows10Release = $Windows10Build }
    }
    # Flash Update, create the path for downloading the latest update into the respecive folder for each release
    if (-not (Test-Path $Windows10CumuPath\$Windows10Release\FLASH)) {
        mkdir "$Windows10CumuPath\$Windows10Release\FLASH"
    }
    # Servicing Stack Update, create the path for downloading the latest update into the respecive folder for each release
    if (-not (Test-Path $Windows10CumuPath\$Windows10Release\SSU)) {
        mkdir "$Windows10CumuPath\$Windows10Release\SSU"
    }
    # Cumulative Update for .NET Framework, create the path for downloading the latest update into the respecive folder for each release
    if (-not (Test-Path $Windows10CumuPath\$Windows10Release\DOTNET)) {
        mkdir "$Windows10CumuPath\$Windows10Release\DOTNET"
    }
    # Zero Day Patch, create the path for downloading the latest update into the respecive folder for each release
    if (-not (Test-Path $Windows10CumuPath\$Windows10Release\ZDP)) {
        mkdir "$Windows10CumuPath\$Windows10Release\ZDP"
    }
    # Download the latest Flash update to the Flash folder
    Get-LatestAdobeFlashUpdate | 
    Where-Object { ($PSItem.Architecture -eq $Architecture) -and ($PSItem.Version -match $Windows10Release) } | 
    Save-LatestUpdate -Path "$Windows10CumuPath\$Windows10Release\FLASH"

    # Download the latest Servicing Stack update to the SSU folder
    Get-LatestServicingStackUpdate -Version $Windows10Release | 
    Where-Object { ($PSItem.Architecture -eq $Architecture) -and ($PSItem.Note -like $Windows10ServerOrClientNote) } |
    Save-LatestUpdate -Path "$Windows10CumuPath\$Windows10Release\SSU"
    
    # Download the latest .NET Framework update to the DOTNET folder (Client only!)
    If ($Windows10ServerOrClient -eq 'WindowsClient') {
        Get-LatestNetFrameworkUpdate -OperatingSystem $Windows10ServerOrClient | 
        Where-Object { ($PSItem.Architecture -eq $Architecture) -and ($PSItem.Note -like $Windows10ServerOrClientNote) -and ($PSItem.Version -eq $Windows10Release) } |
        Save-LatestUpdate -Path "$Windows10CumuPath\$Windows10Release\DOTNET"
    }
   
    # Download the latest update to the zero day patch (ZDP) folder
    Get-LatestCumulativeUpdate -Version $Windows10Release -OperatingSystem $Windows10ServerOrClient | 
    Where-Object { ($PSItem.Architecture -eq $Architecture) } |
    Save-LatestUpdate -Path "$Windows10CumuPath\$Windows10Release\ZDP"
}

