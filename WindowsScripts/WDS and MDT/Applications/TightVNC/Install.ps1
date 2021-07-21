﻿function Get-TightVNCVersion {
    [cmdletbinding()]
    [outputtype([string])]
    $uri = "https://www.tightvnc.com/download.php"
    $web = wget -UseBasicParsing -Uri $uri
    $m = $web.ToString() -split "[`r`n]" | Select-String "Download TightVNC for Windows" | Select-Object -First 1
    $m = $m -replace "<((?!@).)*?>"
    $m = $m.Replace('Download TightVNC for Windows (Version','')
    $m = $m.Replace(')','')
    $m = $m.Replace(' ','')
    $Version = $m
    Write-Output $Version    
}

# PowerShell Wrapper for MDT, Standalone and Chocolatey Installation - (C)2015 xenappblog.com 
# Example 1: Start-Process "XenDesktopServerSetup.exe" -ArgumentList $unattendedArgs -Wait -Passthru
# Example 2 Powershell: Start-Process powershell.exe -ExecutionPolicy bypass -file $Destination
# Example 3 EXE (Always use ' '):
# $UnattendedArgs='/qn'
# (Start-Process "$PackageName.$InstallerType" $UnattendedArgs -Wait -Passthru).ExitCode
# Example 4 MSI (Always use " "):
# $UnattendedArgs = "/i $PackageName.$InstallerType ALLUSERS=1 /qn /liewa $LogApp"
# (Start-Process msiexec.exe -ArgumentList $UnattendedArgs -Wait -Passthru).ExitCode

Clear-Host
Write-Verbose "Setting Arguments" -Verbose
$StartDTM = (Get-Date)

Write-Verbose "Installing Modules" -Verbose
if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
if (!(Get-Module -ListAvailable -Name Evergreen)) {Install-Module Evergreen -Force | Import-Module Evergreen}
Update-Module Evergreen

$Vendor = "Misc"
$Product = "TightVNC"
$PackageName = "TightVNC-win64"
#$Evergreen = Get-NotepadPlusPlus | Where-Object {$_.Architecture -eq "x64"}
$Version = $(Get-TightVNCVersion)
$URL = "https://www.tightvnc.com/download/$Version/tightvnc-$Version-gpl-setup-64bit.msi"
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$LogPS = "${env:SystemRoot}" + "\Temp\$Vendor $Product $Version PS Wrapper.log"
$LogApp = "${env:SystemRoot}" + "\Temp\$PackageName.log"
$Destination = "${env:ChocoRepository}" + "\$Vendor\$Product\$Version\$packageName.$installerType"
$ProgressPreference = 'SilentlyContinue'
$MST = "TightVNC.mst"
$UnattendedArgs = "/i $PackageName.$InstallerType TRANSFORMS=$MST ALLUSERS=1 /qn /liewa $LogApp"

Start-Transcript $LogPS | Out-Null
 
If (!(Test-Path -Path $Version)) {New-Item -ItemType directory -Path $Version | Out-Null}

Copy-Item $MST -Destination $Version -Recurse -Force
 
CD $Version
 
Write-Verbose "Downloading $Vendor $Product $Version" -Verbose
If (!(Test-Path -Path $Source)) {Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $Source}
        
Write-Verbose "Starting Installation of $Vendor $Product $Version" -Verbose
(Start-Process msiexec.exe -ArgumentList $UnattendedArgs -Wait -Passthru).ExitCode

Write-Verbose "Customization" -Verbose

Write-Verbose "Stop logging" -Verbose
$EndDTM = (Get-Date)
Write-Verbose "Elapsed Time: $(($EndDTM-$StartDTM).TotalSeconds) Seconds" -Verbose
Write-Verbose "Elapsed Time: $(($EndDTM-$StartDTM).TotalMinutes) Minutes" -Verbose
Stop-Transcript