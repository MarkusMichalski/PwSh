<#
.SYNOPSIS
    Use this PowerShell script for offline servicing of Windows 10 & Windows Server 2016 images.
.DESCRIPTION
    This script will update the WIM file located in the path, installing the current zero-day-patch. The ZDP is expected to be in the ZDP subdirectory of the path.
    If exists this script to install any other packages (LIPs, .NET Framework, FODv2, SSU), always keep in mind, that the ZDP must be installed as the last package.
.PARAMETERS
    I will write something here. Someday.
.EXAMPLE
    .\Invoke-WinOfflineServicing.ps1
.NOTES
    Script name:    Invoke-WinOfflineServicing.ps1
    Version:        1.0.0.5
    Author:         Markus Michalski
    DateCreated:    2017-06-07
    DateModified:   2019-07-30
#>

#region Parameters
Param(
    [Parameter(HelpMessage = "Enter the path and name of the Windows Image file: ")][string]$MyImageFile = '.\install.wim',
    [Parameter(HelpMessage = "Enter the mount directory/path for the WIM file, this path will be created and deleted automatically: ")][string]$MyMountDir = 'C:\MyMount',
    [Parameter(HelpMessage = "Enter the scratch directory path for the WIM file, this path will be created and deleted automatically: ")][string]$MyScratchDir = 'C:\MyScratch',
    [Parameter(HelpMessage = "Enter the path to the language packs that should be integrated, the package name is optional: ")][string]$MyLPDir = '.\LPs',
    [Parameter(HelpMessage = "Enter the path to the feature on demand packages that should be integrated, the package name is optional: ")][string]$MyFODDir = '.\FOD',
    [Parameter(HelpMessage = "Enter the path to the Flash Player update packages that should be integrated, the package name is optional: ")][string]$MyFlashDir = '.\FLASH',
    [Parameter(HelpMessage = "Enter the path to the .NET Framework update packages that should be integrated, the package name is optional: ")][string]$MyDOTNETDir = '.\DOTNET',
    [Parameter(HelpMessage = "Enter the path to the servicing stack update that should be integrated, the package name is optional: ")][string]$MySSUDir = '.\SSU',
    [Parameter(HelpMessage = "Enter the path to the zero-day-patch that should be integrated, the package name is optional: ")][string]$MyZDPDir = '.\ZDP'
)
#endregion

#region Functions
function Test-Administrator {
    Write-Verbose "Checking script running elevated."
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function NewTemporaryDirectory($Path) {
    Write-Verbose "Checking path: $Path"

    if (Test-Path $Path) {
        Write-Warning "Path: $Path already exists."
    }
    else {
        try {
            New-Item -ItemType Directory -Path $Path | Out-Null
            Write-Host "Created path:" $Path -ForegroundColor Green
        }
        catch {
            Throw $error[0].Exception
        }
    }
}

function RemoveTemporaryDirectory($Path) {
    Write-Verbose "Deleting path: $Path"

    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Force -Recurse | Out-Null
            Write-Host "Path: $Path deleted." -ForegroundColor Green
        }
        catch {
            Throw $error[0].Exception
        }
    }
    else {
        Write-Warning "Path: $Path doesn't exists."
    }
}

function MountWim($WimIndex) {
    Get-WindowsImage -ImagePath $MyImageFile -Index $WimIndex | Select-Object ImageIndex, ImageName, Version, Languages, @{N = 'Size (GB)'; E = {[Math]::Round($_.ImageSize / 1GB, 2)}} | Format-Table -AutoSize
    Write-Verbose "Mounting image with index $WimIndex to $MyMountDir using $MyScratchDir as scratch space."
    try {
        Mount-WindowsImage -Path $MyMountDir -ImagePath $MyImageFile -Index $WimIndex -ScratchDirectory $MyScratchDir
        Write-Host "Image mounted to: $MyMountDir" -ForegroundColor Green
    }
    catch {
        Throw $error[0].Exception
    }
}

function IntegratePackage($PackPath) {
    Write-Verbose "Adding package(s) from $PackPath to image with index $MyWimIndex mounted in $MyMountDir."
    try {
        Add-WindowsPackage -PackagePath $PackPath -Path $MyMountDir -ScratchDirectory $MyScratchDir -NoRestart
        Write-Host "Added package(s) from $PackPath to: $MyMountDir" -ForegroundColor Green
    }
    catch {
        Dismount-WindowsImage -Path $MyMountDir -Discard
        Throw $error[0].Exception
    }
}

function UnmountWim {
    Write-Verbose "Unmounting image from $MyMountDir. "
    try {
        Dismount-WindowsImage -Path $MyMountDir -Save
        Write-Host "Image unmounted from: $MyMountDir" -ForegroundColor Green
    }
    catch {
        Throw $error[0].Exception
    }
}

Function ShowWimContent {
    $WimContent = $NULL
    $WimContentIndex = (Get-WindowsImage -ImagePath $MyImageFile).imageindex
    ForEach ($WimContentObject in $WimContentIndex) {
        $WimContent += @(Get-WindowsImage -ImagePath $MyImageFile -index $WimContentObject | Select-Object ImageIndex, ImageName, Version, Languages, @{N = 'Size (GB)'; E = {[Math]::Round($_.ImageSize / 1GB, 2)}})
    }
    $WimContent | Format-Table -AutoSize
}
#endregion

#region Main

# Is the script running with Admin credentials?
switch (Test-Administrator) {
    $True { Write-Host "Script is running elevated." -ForegroundColor Green }
    $False { Write-Warning “You do not have Administrator rights to run this script!`nPlease re-run this script using an elevated account!”; Break }
}

# Is the INSTALL.WIM available?
switch (Test-Path $MyImageFile) {
    $True { $MyWimIndexes = (Get-WindowsImage -ImagePath $MyImageFile).ImageIndex }
    $False { Write-Warning "Can not find $MyImageFile, unable to continue."; Break }
}

# Show INSTALL.WIM content
ShowWimContent

# Do offline Servicing for each image found within the provided INSTALL.WIM
ForEach ($MyWimIndex in $MyWimIndexes) {
    $ProgressData = @{
        Activity = "Offline servicing $MyImageFile..."
        Status = "Integrating package(s) in image: $MyWimIndex of $($MyWimIndexes.Count)"
        PercentComplete = (($MyWimIndex / $MyWimIndexes.Count) * 100)
    }
    Write-Progress @ProgressData
    NewTemporaryDirectory($MyMountDir)
    NewTemporaryDirectory($MyScratchDir)
    MountWim($MyWimIndex)
    switch (Test-Path $MyLPDir\*.cab) {
        # Apply the Language Packs (if available)
        $True { IntegratePackage($MyLPDir) }
        $False { Write-Verbose "No files found for Language Packs, skipping." }
    }
    switch (Test-Path $MyFODDir\*.cab) {
        # Apply the Features On Demand (if available)
        $True { IntegratePackage($MyFODDir) }
        $False { Write-Verbose "No files found for Features On Demand, skipping." }
    }    
    switch (Test-Path $MyFlashDir\*.msu) {
        # Apply the Flash Player udpate (if available)
        $True { IntegratePackage($MyFlashDir) }
        $False { Write-Verbose "No files found for Flash Player update, skipping." }
    }
    switch (Test-Path $MyDOTNETDir\*.msu) {
        # Apply the .NET Framework udpate (if available)
        $True { IntegratePackage($MyDOTNETDir) }
        $False { Write-Verbose "No files found for .NET Framework update, skipping." }
    }
    switch (Test-Path $MySSUDir\*.msu) {
        # Apply the Servicing Stack Update (if available)
        $True { IntegratePackage($MySSUDir) }
        $False { Write-Verbose "No file found for Servicing Stack Update, skipping." }
    }
    switch (Test-Path $MyZDPDir\*.msu) {
        # Apply the Zero Day Patch (if available)
        $True { IntegratePackage($MyZDPDir) }
        $False { Write-Verbose "No file found for Zero Day Patch, skipping." }
    }
    UnmountWim
    RemoveTemporaryDirectory($MyScratchDir)
    RemoveTemporaryDirectory($MyMountDir)
}

# Show INSTALL.WIM content
ShowWimContent
#endregion