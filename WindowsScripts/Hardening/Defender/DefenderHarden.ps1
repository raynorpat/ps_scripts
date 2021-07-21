Write-Host "Enabling Windows Defender Protections and Features" -ForegroundColor Green -BackgroundColor Black

# Download WDEP xml
$wc = New-Object System.Net.WebClient
$wc.DownloadFile('https://raw.githubusercontent.com/raynorpat/ps_scripts/master/WindowsHardening/Defender/DOD_EP_V3.xml','C:\temp\Windows Defender\DOD_EP_V3.xml')

# Enable Windows Defender Exploit Protection
Write-Host "Enabling Windows Defender Exploit Protections..."
Set-ProcessMitigation -PolicyFilePath "C:\temp\Windows Defender\DOD_EP_V3.xml"

# Download WDAC xml
$wc = New-Object System.Net.WebClient
$wc.DownloadFile('https://raw.githubusercontent.com/raynorpat/ps_scripts/master/WindowsHardening/Defender/WDAC_V1_Recommended_Audit.xml','C:\temp\Windows Defender\WDAC_V1_Recommended_Audit.xml')

# Enable Windows Defender Application Control
#https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-defender-application-control/select-types-of-rules-to-create
Write-Host "Enabling Windows Defender Application Control..."
$PolicyPath = "C:\temp\Windows Defender\WDAC_V1_Recommended_Audit.xml"
ForEach ($PolicyNumber in (1..10)) {
    Write-Host "Importing WDAC Policy Option $PolicyNumber"
    Set-RuleOption -FilePath $PolicyPath -Option $PolicyNumber
}

#https://www.powershellgallery.com/packages/WindowsDefender_InternalEvaluationSetting
#https://social.technet.microsoft.com/wiki/contents/articles/52251.manage-windows-defender-using-powershell.aspx
#https://docs.microsoft.com/en-us/powershell/module/defender/set-mppreference?view=windowsserver2019-ps

# Enable real-time monitoring
Write-Host " -Enabling real-time monitoring"
Set-MpPreference -DisableRealtimeMonitoring $false
# Enable cloud-deliveredprotection
Write-Host " -Enabling cloud-deliveredprotection"
Set-MpPreference -MAPSReporting Advanced
# Enable sample submission
Write-Host " -Enabling sample submission"
Set-MpPreference -SubmitSamplesConsent Always
# Enable checking signatures before scanning
Write-Host " -Enabling checking signatures before scanning"
Set-MpPreference -CheckForSignaturesBeforeRunningScan 1
# Enable behavior monitoring
Write-Host " -Enabling behavior monitoring"
Set-MpPreference -DisableBehaviorMonitoring $false
# Enable IOAV protection
Write-Host " -Enabling IOAV protection"
Set-MpPreference -DisableIOAVProtection $false
# Enable script scanning
Write-Host " -Enabling script scanning"
Set-MpPreference -DisableScriptScanning $false
# Enable removable drive scanning
Write-Host " -Enabling removable drive scanning"
Set-MpPreference -DisableRemovableDriveScanning $false
# Enable Block at first sight
Write-Host " -Enabling Block at first sight"
Set-MpPreference -DisableBlockAtFirstSeen $false
# Enable potentially unwanted apps
Write-Host " -Enabling potentially unwanted apps"
Set-MpPreference -PUAProtection 1
# Enable archive scanning
Write-Host " -Enabling archive scanning"
Set-MpPreference -DisableArchiveScanning $false
# Enable email scanning
Write-Host " -Enabling email scanning"
Set-MpPreference -DisableEmailScanning $false
# Enable File Hash Computation
Write-Host " -Enabling File Hash Computation"
Set-MpPreference -EnableFileHashComputation $true
# Enable Intrusion Prevention System
Write-Host " -Enabling Intrusion Prevention System"
Set-MpPreference -DisableIntrusionPreventionSystem $false
# Enable SSH Parcing
Write-Host " -Enabling SSH Parcing"
Set-MpPreference -DisableSshParsing $false
# Enable TLS Parcing
Write-Host " -Enabling TLS Parcing"
Set-MpPreference -DisableSshParsing $false
# Enable SSH Parcing
Write-Host " -Enabling SSH Parcing"
Set-MpPreference -DisableSshParsing $false
# Enable DNS Parcing
Write-Host " -Enabling DNS Parcing"
Set-MpPreference -DisableDnsParsing $false
Set-MpPreference -DisableDnsOverTcpParsing $false
# Enable DNS Sinkhole 
Write-Host " -Enabling DNS Sinkhole"
Set-MpPreference -EnableDnsSinkhole $true
# Enable Network Protection and setting to block mode
Write-Host " -Enabling Network Protection and setting to block mode"
Set-MpPreference -EnableNetworkProtection Enabled
# Enable Sandboxing for Windows Defender
Write-Host " -Enabling Sandboxing for Windows Defender"
setx /M MP_FORCE_USE_SANDBOX 1 | Out-Null
# Set cloud block level to 'High'
Write-Host " -Setting cloud block level to 'High'"
Set-MpPreference -CloudBlockLevel High
# Set cloud block timeout to 1 minute
Write-Host " -Setting cloud block timeout to 1 minute"
Set-MpPreference -CloudExtendedTimeout 50
# Schedule signature updates every 8 hours
Write-Host " -Scheduling signature updates every 8 hours"
Set-MpPreference -SignatureUpdateInterval 8
# Randomize Scheduled Task Times
Write-Host " -Randomizing Scheduled Task Times"
Set-MpPreference -RandomizeScheduleTaskTimes $true

# Enable Controlled Folder Access and setting to block mode
Write-Host " -Enabling Controlled Folder Access and setting to block mode"
Set-MpPreference -EnableControlledFolderAccess $true

# Dismiss Microsoft Defender offer in the Windows Security about signing in into a Microsoft account
Write-Host "Disabling Account Prompts"
if (!(Test-Path -Path "HKCU:\SOFTWARE\Microsoft\Windows Security Health\State\AccountProtection_MicrosoftAccount_Disconnected")) {
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows Security Health\State" -Name "AccountProtection_MicrosoftAccount_Disconnected" -PropertyType "DWORD" -Value "1" -Force
} else {
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows Security Health\State" -Name "AccountProtection_MicrosoftAccount_Disconnected" -PropertyType "DWORD" -Value "1" -Force
}

# Enable Cloud-delivered Protections
Write-Host "Enabling Cloud-delivered Protection"
Set-MpPreference -MAPSReporting Advanced
Set-MpPreference -SubmitSamplesConsent SendAllSamples

#https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/enable-attack-surface-reduction
#https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction
Write-Host "Enabling... Windows Defender Attack Surface Reduction Rules"
Write-Host " -Block executable content from email client and webmail"
Add-MpPreference -AttackSurfaceReductionRules_Ids BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block all Office applications from creating child processes"
Add-MpPreference -AttackSurfaceReductionRules_Ids D4F940AB-401B-4EFC-AADC-AD5F3C50688A -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block Office applications from creating executable content"
Add-MpPreference -AttackSurfaceReductionRules_Ids 3B576869-A4EC-4529-8536-B80A7769E899 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block Office applications from injecting code into other processes"
Add-MpPreference -AttackSurfaceReductionRules_Ids 75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block JavaScript or VBScript from launching downloaded executable content"
Add-MpPreference -AttackSurfaceReductionRules_Ids D3E037E1-3EB8-44C8-A917-57927947596D -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block execution of potentially obfuscated scripts"
Add-MpPreference -AttackSurfaceReductionRules_Ids 5BEB7EFE-FD9A-4556-801D-275E5FFC04CC -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block Win32 API calls from Office macros"
Add-MpPreference -AttackSurfaceReductionRules_Ids 92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block executable files from running unless they meet a prevalence, age, or trusted list criterion"
Add-MpPreference -AttackSurfaceReductionRules_Ids 01443614-cd74-433a-b99e-2ecdc07bfc25 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block credential stealing from the Windows local security authority subsystem"
Add-MpPreference -AttackSurfaceReductionRules_Ids 9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block persistence through WMI event subscription"
Add-MpPreference -AttackSurfaceReductionRules_Ids e6db77e5-3df2-4cf1-b95a-636979351e5b -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block process creations originating from PSExec and WMI commands"
Add-MpPreference -AttackSurfaceReductionRules_Ids d1e49aac-8f56-4280-b9ba-993a6d77406c -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block untrusted and unsigned processes that run from USB"
Add-MpPreference -AttackSurfaceReductionRules_Ids b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block Office communication application from creating child processes"
Add-MpPreference -AttackSurfaceReductionRules_Ids 26190899-1602-49e8-8b27-eb1d0a1ce869 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block Adobe Reader from creating child processes"
Add-MpPreference -AttackSurfaceReductionRules_Ids 7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block persistence through WMI event subscription"
Add-MpPreference -AttackSurfaceReductionRules_Ids e6db77e5-3df2-4cf1-b95a-636979351e5b -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Block abuse of exploited vulnerable signed drivers"
Add-MpPreference -AttackSurfaceReductionRules_Ids 56a863a9-875e-4185-98a7-b882c64b5ce5 -AttackSurfaceReductionRules_Actions Enabled
Write-Host " -Use advanced protection against ransomware"
Add-MpPreference -AttackSurfaceReductionRules_Ids c1db55ab-c21a-4637-bb3f-a12568109d35 -AttackSurfaceReductionRules_Actions Enabled

# Update Signatures
Write-Host "Updating Signatures..."
Update-MpSignature -UpdateSource MicrosoftUpdateServer
Update-MpSignature -UpdateSource MMPC

# Print current Defender configuration
Write-Host "Printing Current Windows Defender Configuration"
Get-MpComputerStatus ; Get-MpPreference ; Get-MpThreat ; Get-MpThreatDetection
