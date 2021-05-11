#requires -Version 3.0
<#

    .SYNOPSIS
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication.

    .DESCRIPTION
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication. This report can be customized to report on Backup, Replication,
    Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
    infrastructure details like repositories, proxies and license status. Work
    through the User Variables to determine what you would like to see and
    determine if you would like to save the results to a file or have them
    emailed to you.

    .EXAMPLE
    .\MyVeeamReport.ps1
    Run script from (an elevated) PowerShell console  
  
    .NOTES
    Author: Shawn Masterson, edited by Pat Raynor
    Last Updated: May 2021
    Version: 9.5.4.1
  
    Requires:
    Veeam Backup & Replication v10 (full or console install)

#> 

#region User-Variables
# VBR Server (Server Name, FQDN or IP)
$vbrServer = "VEEAMBDR1"
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = 24
# Report Title
$rptTitle = "BDR Veeam Report"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97

# Location of Veeam executable (Veeam.Backup.Shell.exe)
$veeamExePath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Shell.exe"

# Save HTML output to a file
$saveHTML = $false
# HTML File output path and filename
$pathHTML = "C:\Backup\MyVeeamReport_$(Get-Date -format MMddyyyy_hhmmss).htm"
# Launch HTML file after creation
$launchHTML = $false

# Email configuration
$sendEmail = $true
$emailHost = "mail.xxx.com"
$emailPort = 25
$emailEnableSSL = $false
$emailUser = ""
$emailPass = ""
$emailFrom = ""
$emailTo = ""
# Send HTML report as attachment (else HTML report is body)
$emailAttach = $false
# Email Subject 
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$modeSubject = $true
# Append VBR Server name to Email Subject
$vbrSubject = $true
# Append Date and Time to Email Subject
$dtSubject = $true

# Show VM Backup Protection Summary (across entire infrastructure)
$showSummaryProtect = $false
# Show VMs with No Successful Backups within RPO ($reportMode)
$showUnprotectedVMs = $false
# Show VMs with Successful Backups within RPO ($reportMode)
# Also shows VMs with Only Backups with Warnings within RPO ($reportMode)
$showProtectedVMs = $true
# Exclude VMs from Missing and Successful Backups sections
# $excludevms = @("vm1","vm2","*_replica")
$excludeVMs = @("")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) folder(s)
# $excludeFolder = @("folder1","folder2","*_testonly")
$excludeFolder = @("")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) datacenter(s)
# $excludeDC = @("dc1","dc2","dc*")
$excludeDC = @("")
# Exclude Templates from Missing and Successful Backups sections
$excludeTemp = $false

# Show VMs Backed Up by Multiple Jobs within time frame ($reportMode)
$showMultiJobs = $true

# Show Backup Session Summary
$showSummaryBk = $true
# Show Backup Job Status
$showJobsBk = $true
# Show Backup Job Size (total)
$showBackupSizeBk = $true
# Show detailed information for Backup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBk = $true
# Show all Backup Sessions within time frame ($reportMode)
$showAllSessBk = $true
# Show all Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksBk = $true
# Show Running Backup Jobs
$showRunningBk = $true
# Show Running Backup Tasks
$showRunningTasksBk = $true
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBk = $true
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBk = $true
# Show Successful Backup Sessions within time frame ($reportMode)
$showSuccessBk = $true
# Show Successful Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBk = $true
# Only show last Session for each Backup Job
$onlyLastBk = $true
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")

# Show Running Restore VM Sessions
$showRestoRunVM = $true
# Show Completed Restore VM Sessions within time frame ($reportMode)
$showRestoreVM = $true

# Show Replication Session Summary
$showSummaryRp = $false
# Show Replication Job Status
$showJobsRp = $true
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $true
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $false
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $false
# Show Running Replication Jobs
$showRunningRp = $false
# Show Running Replication Tasks
$showRunningTasksRp = $false
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $true
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $true
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $true
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $false
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $true
# Show Backup Copy Job Status
$showJobsBc = $true
# Show Backup Copy Job Size (total)
$showBackupSizeBc = $true
# Show detailed information for Backup Copy Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBc = $true
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $false
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $false
# Show Idle Backup Copy Sessions
$showIdleBc = $false
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $false
# Show Working Backup Copy Jobs
$showRunningBc = $false
# Show Working Backup Copy Tasks
$showRunningTasksBc = $false
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $true
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $true
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $true
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $false
# Only show last Session for each Backup Copy Job
$onlyLastBc = $false
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")

# Show Tape Backup Session Summary
$showSummaryTp = $false
# Show Tape Backup Job Status
$showJobsTp = $true
# Show detailed information for Tape Backup Sessions (Avg Speed, Total(GB), Read(GB), Transferred(GB))
$showDetailedTp = $true
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $false
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $false
# Show Waiting Tape Backup Sessions
$showWaitingTp = $false
# Show Idle Tape Backup Sessions
$showIdleTp = $false
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $false
# Show Working Tape Backup Jobs
$showRunningTp = $false
# Show Working Tape Backup Tasks
$showRunningTasksTp = $false
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $false
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $false
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $false
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $false
# Only show last Session for each Tape Backup Job
$onlyLastTp = $false
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show all Tapes
$showTapes = $false
# Show all Tapes by (Custom) Media Pool
$showTpMp = $false
# Show all Tapes by Vault
$showTpVlt = $false
# Show all Expired Tapes
$showExpTp = $false
# Show Expired Tapes by (Custom) Media Pool
$showExpTpMp = $false
# Show Expired Tapes by Vault
$showExpTpVlt = $false
# Show Tapes written to within time frame ($reportMode)
$showTpWrt = $false

# Show Agent Backup Session Summary
$showSummaryEp = $true
# Show Agent Backup Job Status
$showJobsEp = $true
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $true
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $true
# Show Running Agent Backup jobs
$showRunningEp = $true
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $true
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $true
# Only show last session for each Agent Backup Job
$onlyLastEp = $true
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show SureBackup Session Summary
$showSummarySb = $false
# Show SureBackup Job Status
$showJobsSb = $false
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $false
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $false
# Show Running SureBackup Jobs
$showRunningSb = $false
# Show Running SureBackup Tasks
$showRunningTasksSb = $false
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $false
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $false
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $false
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $false
# Only show last Session for each SureBackup Job
$onlyLastSb = $false
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $false
# Show Proxy Info
$showProxy = $false
# Show Repository Info
$showRepo = $true
# Show Repository Permissions for Agent Jobs
$showRepoPerms = $false
# Show Replica Target Info
$showReplicaTarget = $false
# Show Veeam Services Info (Windows Services)
$showServices = $false
# Show only Services that are NOT running
$hideRunningSvc = $false
# Show License expiry info
$showLicExp = $true

# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 5
$repoWarn = 10
# Replica Target Free Space Remaining %
$replicaCritical = 5
$replicaWarn = 10
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
#endregion
 
#region VersionInfo
$MVRversion = "9.5.4.1"
# Version 9.5.4.1 - raynorpat
# Tweaks to repository free space thresholds
# Replace WMI registry check for license info with a direct grab with Get-VBRInstalledLicense, fixes licensing check with Veeam 10 and later
# Edited link at bottom of report to link to this GitHub repo
#
# Version 9.5.4 - raynorpat
# Added -WarningAction SilentlyContinue | where {$_.BackupPlatform.Platform -ne 'ELinuxPhysical' -and $_.BackupPlatform.Platform -ne 'EEndPoint'} to Get-VBRJob call
#  - prevents issues from Veeam v10 from grabbing deprecated agent jobs
# Removed check for Veeam v9 or lower
# Changed Get-VBREP calls to use the more basic Get-VBRComputerBackupJob calls for use with Veeam 10 Agent backups
#
# Version 9.5.3 - SM
# Updated property changes introduced in VBR 9.5 Update 3
# Version 9.5.1.1 - SM
# Minor bug fixes:
# Removed requires VBR snapin
# Fixed HourstoCheck variable in Get-VMsBackupStatus function
# Version 9.5.1 - SM
# Updated HTML formatting - thanks for the inspiration Nick!
# Report header and email subject now reflect results (Failed/Warning/Success)
# Added report section - VMs Backed Up by Multiple Jobs within RPO
# Added report section - Repository Permissions for Agent Jobs
# Added Description field for Agent Job Status to identify type of Agent
# Added Next Run field for Agent Job Status (Fixed in VBR 9.5 Update 1)
# Added Next Run field for Configuration Backup Status (Fixed in VBR 9.5 Update 1)
# Added more details to VMs with No Successful/Successful/with Warnings within RPO
# Appended date and time to email attachment file name
# Added ability to append date and time to email subject
# Added ability to send email via SSL/TLS
# Renamed Endpoints to Agents
#
# Version 9.0.3 - SM
# Added report section - VM Backup Protection Summary (across entire infrastructure)
# Split report section - Split out VMs with only Backups with Warnings within RPO to separate from Successful
# Added report section - Backup Job Size (total)
# Added report section - All Backup Sessions
# Added report section - All Backup Tasks
# Added report section - Running Backup Tasks
# Added report section - Backup Tasks with Warnings or Failures
# Added report section - Successful Backup Tasks
# Added report section - Replication Job/Session Summary
# Added report section - Replication Job Status
# Added report section - All Replication Sessions
# Added report section - All Replication Tasks
# Added report section - Running Replication Jobs
# Added report section - Running Replication Tasks
# Added report section - Replication Job/Sessions with Warnings or Failures
# Added report section - Replication Tasks with Warnings or Failures
# Added report section - Successful Replication Jobs/Sessions
# Added report section - Successful Replication Tasks
# Added report section - Backup Copy Session Summary
# Added report section - Backup Copy Job Status
# Added report section - Backup Copy Job Size (total)
# Added report section - All Backup Copy Sessions
# Added report section - All Backup Copy Tasks
# Added report section - Idle Backup Copy Sessions
# Added report section - Pending Backup Copy Tasks
# Added report section - Working Backup Copy Jobs
# Added report section - Working Backup Copy Tasks
# Added report section - Backup Copy Sessions with Warnings or Failures
# Added report section - Backup Copy Tasks with Warnings or Failures
# Added report section - Successful Backup Copy Sessions
# Added report section - Successful Backup Copy Tasks
# Added report section - Tape Backup Session Summary
# Added report section - Tape Job Status
# Added report section - All Tape Backup Sessions
# Added report section - All Tape Backup Tasks
# Added report section - Waiting Tape Backup Sessions
# Added report section - Idle Tape Backup Sessions
# Added report section - Pending Tape Backup Tasks
# Added report section - Working Tape Backup Jobs
# Added report section - Working Tape Backup Tasks
# Added report section - Tape Backup Sessions with Warnings or Failures
# Added report section - Tape Backup Tasks with Warnings or Failures
# Added report section - Successful Tape Backup Sessions
# Added report section - Successful Tape Backup Tasks
# Added report section - All Tapes
# Added report section - All Tapes by (Custom) Media Pool
# Added report section - All Tapes by Vault
# Added report section - All Expired Tapes
# Added report section - Expired Tapes by (Custom) Media Pool - Thanks to Patrick IRVING & Olivier Dubroca!
# Added report section - Expired Tapes by Vault
# Added report section - All Tapes written to within time frame ($reportMode)
# Added report section - Endpoint Backup Job Size (total)
# Added report section - All Endpoint Backup Sessions
# Added report section - SureBackup Session Summary
# Added report section - SureBackup Job Status
# Added report section - All SureBackup Sessions
# Added report section - All SureBackup Tasks
# Added report section - Running SureBackup Jobs
# Added report section - Running SureBackup Tasks
# Added report section - SureBackup Sessions with Warnings or Failures
# Added report section - SureBackup Tasks with Warnings or Failures
# Added report section - Successful SureBackup Sessions
# Added report section - Successful SureBackup Tasks
# Added report section - Configuration Backup Status
# Added report section - Scale Out Repository Info - Thanks to Patrick IRVING & Olivier Dubroca!
# Added exclusion for Templates to VM Backup Protection sections
# Added Last Start and End times to VMs with Successful/Warning Backups
# Added Dedupe and Compression to Backup/Backup Copy/Replication session detailed info
# Added ability to report only on particular jobs (backup/replica/backup copy/tape/surebackup/endpoint)
# Added Mode/Type and Maximum Tasks to Proxy and Repository Info
# Filtered some heavy lifting commands to only run when/if needed
# Converted durations from Mins to HH:MM:SS
# Added html formatting of cells (vertical-align: middle;text-align:center;)
# Lots of misc tweaks/cleanup
#
# Version 9.0.2 - SM
# Fixed issue with Proxy details reported when using IP address instead of server names
# Fixed an issue where services were reported multiple times per server
#
# Version 9.0.1 - SM
# Initial version for VBR v9
# Updated version to follow VBR version (VeeamMajorVersion.VeeamMinorVersion.MVRVersion)
# Fixed Proxy Information (change in property names in v9)
# Rewrote Repository Info to use newly available properties (yay!)
# Updated Get-VMsBackupStatus to remove obsolete commandlet warning (Thanks tsightler!)
# Added ability to run from console only install
# Added ability to include VBR server in report title and email subject
# Rewrote License Info gathering to allow remote info gathering
# Misc minor tweaks/cleanup
#
# Version 2.0 - SM
# Misc minor tweaks/cleanup
# Proxy host IP info now always returns IPv4 address
# Added ability to query Veeam database for Repository size info
#   Big thanks to tsightler - http://forums.veeam.com/powershell-f26/get-vbrbackuprepository-why-no-size-info-t27296.html
# Added report section - Backup Job Status
# Added option to show detailed Backup Job/Session information (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB))
# Added report section - Running VM Restore Sessions
# Added report section - Completed VM Restore Sessions
# Added report section - Endpoint Backup Results Summary
# Added report section - Endpoint Backup Job Status
# Added report section - Running Endpoint Backup Jobs
# Added report section - Endpoint Backup Jobs/Sessions with Warnings or Failures
# Added report section - Successful Endpoint Backup Jobs/Sessions
#
# Version 1.4.1 - SM
# Fixed issue with summary counts
# Version 1.4 - SM
# Misc minor tweaks/cleanup
# Added variable for report width
# Added variable for email subject
# Added ability to show/hide all report sections
# Added Protected/Unprotected VM Count to Summary
# Added per object details for sessions w/no details
# Added proxy host name to Proxy Details
# Added repository host name to Repository Details
# Added section showing successful sessions
# Added ability to view only last session per job
# Added Cluster field for protected/unprotected VMs
# Added catch for cifs repositories greater than 4TB as erroneous data is returned
# Added % Complete for Running Jobs
# Added ability to exclude multiple (vCenter) folders from Missing and Successful Backups section
# Added ability to exclude multiple (vCenter) datacenters from Missing and Successful Backups section
# Tweaked license info for better reporting across different date formats
#
# Version 1.3 - SM
# Now supports VBR v8
# For VBR v7, use report version 1.2
# Added more flexible options to save and launch file 
#
# Version 1.2 - SM
# Added option to show VMs Successfully backed up
#
# Version 1.1.4 - SM
# Misc tweaks/bug fixes
# Reconfigured HTML a bit to help with certain email clients
# Added cell coloring to highlight status
# Added $rptTitle variable to hold report title
# Added ability to send report via email as attachment
#
# Version 1.1.3 - SM
# Added Details to Sessions with Warnings or Failures
#
# Version 1.1.2 - SM
# Minor tweaks/updates
# Added Veeam version info to header
#
# Version 1.1.1 - Shawn Masterson
# Based on vPowerCLI v6 Army Report (v1.1) by Thomas McConnell
# http://www.vpowercli.co.uk/2012/01/23/vpowercli-v6-army-report/
# http://pastebin.com/6p3LrWt7
#
# Tweaked HTML header (color, title)
#
# Changed report width to 1024px
#
# Moved hard-coded path to exe/dll files to user declared variables ($veeamExePath/$veeamDllPath)
#
# Adjusted sorting on all objects
#
# Modified info group/counts
#   Modified - Total Jobs = Job Runs
#   Added - Read (GB)
#   Added - Transferred (GB)
#   Modified - Warning = Warnings
#   Modified - Failed = Failures
#   Added - Failed (last session)
#   Added - Running (currently running sessions)
# 
# Modified job lines
#   Renamed Header - Sessions with Warnings or Failures
#   Fixed Write (GB) - Broke with v7
#   
# Added support license renewal
#   Credit - Gavin Townsend  http://www.theagreeablecow.com/2012/09/sysadmin-modular-reporting-samreports.html
#   Original  Credit - Arne Fokkema  http://ict-freak.nl/2011/12/29/powershell-veeam-br-get-total-days-before-the-license-expires/
#
# Modified Proxy section
#   Removed Read/Write/Util - Broke in v7 - Workaround unknown
# 
# Modified Services section
#   Added - $runningSvc variable to toggle displaying services that are running
#   Added - Ability to hide section if no results returned (all services are running)
#   Added - Scans proxies and repositories as well as the VBR server for services
#
# Added VMs Not Backed Up section
#   Credit - Tom Sightler - http://sightunseen.org/blog/?p=1
#   http://www.sightunseen.org/files/vm_backup_status_dev.ps1
#   
# Modified $reportMode
#   Added ability to run with any number of hours (8,12,72 etc)
#   Added bits to allow for zero sessions (semi-gracefully)
#
# Added Running Jobs section
#   Added ability to toggle displaying running jobs
#
# Added catch to ensure running v7 or greater
#
#
# Version 1.1
# Added job lines as per a request on the website
#
# Version 1.0
# Clean up for release
#
# Version 0.9
# More cmdlet rewrite to improve perfomace, credit to @SethBartlett
# for practically writing the Get-vPCRepoInfo
#
# Version 0.8
# Added Read/Write stats for proxies at requests of @bsousapt
# Performance improvement of proxy tear down due to rewrite of cmdlet
# Replaced 2 other functions
# Added Warning counter, .00 to all storage returns and fetch credentials for
# remote WinLocal repos
#
# Version 0.7
# Added Utilisation(Get-vPCDailyProxyUsage) and Modes 24, 48, Weekly, and Monthly
# Minor performance tweaks 
#endregion

#region Connect
# Load Veeam Snapin
If (!(Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue)) {
  If (!(Add-PSSnapin -PassThru VeeamPSSnapIn)) {
    Write-Error "Unable to load Veeam snapin" -ForegroundColor Red
    Exit
  }
}

# Connect to VBR server
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer){
  Disconnect-VBRServer
  Try {
    Connect-VBRServer -server $vbrServer -ErrorAction Stop
  } Catch {
    Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
    exit
  }
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
  $showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
  $showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
  $showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
  $showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
  $showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
  $showTaskWFBc + $showSuccessBc + $showTaskSuccessBc) {
  $allJobs = Get-VBRJob -WarningAction SilentlyContinue | where {$_.BackupPlatform.Platform -ne 'ELinuxPhysical' -and $_.BackupPlatform.Platform -ne 'EEndPoint'}
}
# Get all Backup Jobs
$allJobsBk = @($allJobs | ?{$_.JobType -eq "Backup"})
# Get all Replication Jobs
$allJobsRp = @($allJobs | ?{$_.JobType -eq "Replica"})
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | ?{$_.JobType -eq "BackupSync"})
# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
  $showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
  $showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
  $allJobsTp = @(Get-VBRTapeJob)
}
# Get all Agent Backup Jobs
$allJobsEp = @()
If ($showSummaryEp + $showJobsEp + $showAllSessEp + $showRunningEp +
  $showWarnFailEp + $showSuccessEp) {
  $allJobsEp = @(Get-VBRComputerBackupJob)
}
# Get all SureBackup Jobs
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb + 
  $showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb + 
  $showSuccessSb + $showTaskSuccessSb) {
  $allJobsSb = @(Get-VSBJob)
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
  $allSess = Get-VBRBackupSession
}
# Get all Restore Sessions
$allSessResto = @()
If ($showRestoRunVM + $showRestoreVM) {
  $allSessResto = Get-VBRRestoreSession
}
# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
  Foreach ($tpJob in $allJobsTp){
    $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
    $allSessTp += $tpSessions
  }
}
# Get all Agent Backup Sessions
$allSessEp = @()
If ($allJobsEp) {
  $allSessEp = Get-VBRComputerBackupJobSession
}
# Get all SureBackup Sessions
$allSessSb = @()
If ($allJobsSb) {
  $allSessSb = Get-VSBSession
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
  $jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | ?{$_.JobType -eq "Backup"})
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | ?{$_.JobType -eq "BackupSync"})
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | ?{$_.JobType -eq "EndpointBackup"})

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool
# Get all Media Vaults
$mediaVaults = Get-VBRTapeVault
# Get all Tapes
$mediaTapes = Get-VBRTapeMedium
# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary
# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob
# Get VBR Server object
$vbrServerObj = Get-VBRLocalhost
# Get all Proxies
$proxyList = Get-VBRViProxy
# Get all Repositories
$repoList = Get-VBRBackupRepository
$repoListSo = Get-VBRBackupRepository -ScaleOut
# Get all Tape Servers
$tapesrvList = Get-VBRTapeServer

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
  $HourstoCheck = 720
} Elseif ($reportMode -eq "Weekly") {
  $HourstoCheck = 168
} Else {
  $HourstoCheck = $reportMode
}

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($backupJob -ne $null -and $backupJob -ne "") {
  $allJobsBkTmp = @()
  $sessListBkTmp = @()
  $backupsBkTmp = @()
  Foreach ($bkJob in $backupJob) {
    $allJobsBkTmp += $allJobsBk | ?{$_.Name -like $bkJob}
    $sessListBkTmp += $sessListBk | ?{$_.JobName -like $bkJob}
    $backupsBkTmp += $backupsBk | ?{$_.JobName -like $bkJob}
  }
  $allJobsBk = $allJobsBkTmp | sort Id -Unique
  $sessListBk = $sessListBkTmp | sort Id -Unique
  $backupsBk = $backupsBkTmp | sort Id -Unique
}
If ($onlyLastBk) {
  $tempSessListBk = $sessListBk
  $sessListBk = @()
  Foreach($job in $allJobsBk) {
    $sessListBk += $tempSessListBk | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0
$sessListBk | %{$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBk | %{$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsBk = @($sessListBk | ?{$_.Result -eq "Success"})
$warningSessionsBk = @($sessListBk | ?{$_.Result -eq "Warning"})
$failsSessionsBk = @($sessListBk | ?{$_.Result -eq "Failed"})
$runningSessionsBk = @($sessListBk | ?{$_.State -eq "Working"})
$failedSessionsBk = @($sessListBk | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather VM Restore Sessions within timeframe
$sessListResto = @($allSessResto | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or !($_.IsCompleted)})
# Get VM Restore Session information
$completeResto = @($sessListResto | ?{$_.IsCompleted})
$runningResto = @($sessListResto | ?{!($_.IsCompleted)})

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica"})
If ($replicaJob -ne $null -and $replicaJob -ne "") {
  $allJobsRpTmp = @()
  $sessListRpTmp = @()
  Foreach ($rpJob in $replicaJob) {
    $allJobsRpTmp += $allJobsRp | ?{$_.Name -like $rpJob}
    $sessListRpTmp += $sessListRp | ?{$_.JobName -like $rpJob}
  }
  $allJobsRp = $allJobsRpTmp | sort Id -Unique
  $sessListRp = $sessListRpTmp | sort Id -Unique
}
If ($onlyLastRp) {
  $tempSessListRp = $sessListRp
  $sessListRp = @()
  Foreach($job in $allJobsRp) {
    $sessListRp += $tempSessListRp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Replication Session information
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | %{$totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRp | %{$totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRp = @($sessListRp | ?{$_.Result -eq "Success"})
$warningSessionsRp = @($sessListRp | ?{$_.Result -eq "Warning"})
$failsSessionsRp = @($sessListRp | ?{$_.Result -eq "Failed"})
$runningSessionsRp = @($sessListRp | ?{$_.State -eq "Working"})
$failedSessionsRp = @($sessListRp | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and $_.JobType -eq "BackupSync"})
If ($bcopyJob -ne $null -and $bcopyJob -ne "") {
  $allJobsBcTmp = @()
  $sessListBcTmp = @()
  $backupsBcTmp = @()
  Foreach ($bcJob in $bcopyJob) {
    $allJobsBcTmp += $allJobsBc | ?{$_.Name -like $bcJob}
    $sessListBcTmp += $sessListBc | ?{$_.JobName -like $bcJob}
    $backupsBcTmp += $backupsBc | ?{$_.JobName -like $bcJob}
  }
  $allJobsBc = $allJobsBcTmp | sort Id -Unique
  $sessListBc = $sessListBcTmp | sort Id -Unique
  $backupsBc = $backupsBcTmp | sort Id -Unique
}
If ($onlyLastBc) {
  $tempSessListBc = $sessListBc
  $sessListBc = @()
  Foreach($job in $allJobsBc) {
    $sessListBc += $tempSessListBc | ?{$_.Jobname -eq $job.name -and $_.BaseProgress -eq 100} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Copy Session information
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | %{$totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBc | %{$totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsBc = @($sessListBc | ?{$_.State -eq "Idle"})
$successSessionsBc = @($sessListBc | ?{$_.Result -eq "Success"})
$warningSessionsBc = @($sessListBc | ?{$_.Result -eq "Warning"})
$failsSessionsBc = @($sessListBc | ?{$_.Result -eq "Failed"})
$workingSessionsBc = @($sessListBc | ?{$_.State -eq "Working"})

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
If ($tapeJob -ne $null -and $tapeJob -ne "") {
  $allJobsTpTmp = @()
  $sessListTpTmp = @()
  Foreach ($tpJob in $tapeJob) {
    $allJobsTpTmp += $allJobsTp | ?{$_.Name -like $tpJob}
    $sessListTpTmp += $sessListTp | ?{$_.JobName -like $tpJob}
  }
  $allJobsTp = $allJobsTpTmp | sort Id -Unique
  $sessListTp = $sessListTpTmp | sort Id -Unique
}
If ($onlyLastTp) {
  $tempSessListTp = $sessListTp
  $sessListTp = @()
  Foreach($job in $allJobsTp) {
    $sessListTp += $tempSessListTp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Tape Backup Session information
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | %{$totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListTp | %{$totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsTp = @($sessListTp | ?{$_.State -eq "Idle"})
$successSessionsTp = @($sessListTp | ?{$_.Result -eq "Success"})
$warningSessionsTp = @($sessListTp | ?{$_.Result -eq "Warning"})
$failsSessionsTp = @($sessListTp | ?{$_.Result -eq "Failed"})
$workingSessionsTp = @($sessListTp | ?{$_.State -eq "Working"})
$waitingSessionsTp = @($sessListTp | ?{$_.State -eq "WaitingTape"})

# Gather all Agent Backup Sessions within timeframe
$sessListEp = $allSessEp | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")}
If ($epbJob -ne $null -and $epbJob -ne "") {
  $allJobsEpTmp = @()
  $sessListEpTmp = @()
  $backupsEpTmp = @()
  Foreach ($eJob in $epbJob) {
    $allJobsEpTmp += $allJobsEp | ?{$_.Name -like $eJob}
    $backupsEpTmp += $backupsEp | ?{$_.JobName -like $eJob}
  }
  Foreach ($job in $allJobsEpTmp) {
    $sessListEpTmp += $sessListEp | ?{$_.JobId -eq $job.Id}
  }
  $allJobsEp = $allJobsEpTmp | sort Id -Unique
  $sessListEp = $sessListEpTmp | sort Id -Unique
  $backupsEp = $backupsEpTmp | sort Id -Unique
}
If ($onlyLastEp) {
  $tempSessListEp = $sessListEp
  $sessListEp = @()
  Foreach($job in $allJobsEp) {
    $sessListEp += $tempSessListEp | ?{$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Agent Backup Session information
$successSessionsEp = @($sessListEp | ?{$_.Result -eq "Success"})
$warningSessionsEp = @($sessListEp | ?{$_.Result -eq "Warning"})
$failsSessionsEp = @($sessListEp | ?{$_.Result -eq "Failed"})
$runningSessionsEp = @($sessListEp | ?{$_.State -eq "Working"})

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped"})
If ($surebJob -ne $null -and $surebJob -ne "") {
  $allJobsSbTmp = @()
  $sessListSbTmp = @()
  Foreach ($SbJob in $surebJob) {
    $allJobsSbTmp += $allJobsSb | ?{$_.Name -like $SbJob}
    $sessListSbTmp += $sessListSb | ?{$_.JobName -like $SbJob}
  }
  $allJobsSb = $allJobsSbTmp | sort Id -Unique
  $sessListSb = $sessListSbTmp | sort Id -Unique
}
If ($onlyLastSb) {
  $tempSessListSb = $sessListSb
  $sessListSb = @()
  Foreach($job in $allJobsSb) {
    $sessListSb += $tempSessListSb | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get SureBackup Session information
$successSessionsSb = @($sessListSb | ?{$_.Result -eq "Success"})
$warningSessionsSb = @($sessListSb | ?{$_.Result -eq "Warning"})
$failsSessionsSb = @($sessListSb | ?{$_.Result -eq "Failed"})
$runningSessionsSb = @($sessListSb | ?{$_.State -ne "Stopped"})

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
  $rptMode = "RPO: $reportMode Hrs"
} Else {
  $rptMode = "RPO: $reportMode"
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

# Append Report Mode to Email subject
If ($modeSubject) {
  If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
    $emailSubject = "$emailSubject (Last $reportMode Hrs)"
  } Else {
    $emailSubject = "$emailSubject ($reportMode)"
  }
}

# Append VBR Server to Email subject
If ($vbrSubject) {
  $emailSubject = "$emailSubject - $vbrServer"
}

# Append Date and Time to Email subject
If ($dtSubject) {
  $emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions
 
Function Get-VBRProxyInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Proxy
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param ([PsObject]$inputObj)
      $ping = new-object system.net.networkinformation.ping
      $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
      If ($inputObj.Host.Name -match $isIP) {
        $IPv4 = $inputObj.Host.Name
      } Else {
        $DNS = [Net.DNS]::GetHostEntry("$($inputObj.Host.Name)")
        $IPv4 = ($DNS.get_AddressList() | Where {$_.AddressFamily -eq "InterNetwork"} | Select -First 1).IPAddressToString
      }
      $pinginfo = $ping.send("$($IPv4)")           
      If ($pinginfo.Status -eq "Success") {
        $hostAlive = "Alive"
        $response = $pinginfo.RoundtripTime
      } Else {
        $hostAlive = "Dead"
        $response = $null
      }
      If ($inputObj.IsDisabled) {
        $enabled = "False"
      } Else {
        $enabled = "True"
      }   
      $tMode = switch ($inputObj.Options.TransportMode) {
        "Auto" {"Automatic"}
        "San" {"Direct SAN"}
        "HotAdd" {"Hot Add"}
        "Nbd" {"Network"}
        default {"Unknown"}   
      }
      $vPCFuncObject = New-Object PSObject -Property @{
        ProxyName = $inputObj.Name
        RealName = $inputObj.Host.Name.ToLower()
        Disabled = $inputObj.IsDisabled
        pType = $inputObj.ChassisType
        Status  = $hostAlive
        IP = $IPv4
        Response = $response
        Enabled = $enabled
        maxtasks = $inputObj.Options.MaxTasksCount
        tMode = $tMode
      }
      Return $vPCFuncObject
    }
  }
  Process {
    Foreach ($p in $Proxy) {
      $outputObj = Build-Object $p
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }   
}

Function Get-VBRRepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        Target = $name
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        MaxTasks = $maxtasks
        rType = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($r in $Repository) {
      # Refresh Repository Size Info
      [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
      $rType = switch ($r.Type) {
        "WinLocal" {"Windows Local"}
        "LinuxLocal" {"Linux Local"}
        "CifsShare" {"CIFS Share"}
        "DataDomain" {"Data Domain"}
        "ExaGrid" {"ExaGrid"}
        "HPStoreOnce" {"HP StoreOnce"}
        default {"Unknown"}   
      }
      $outputObj = Build-Object $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.info.CachedFreeSpace $r.Info.CachedTotalSpace $r.Options.MaxTaskCount $rType
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }
}

Function Get-VBRSORepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $rname, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        SoTarget = $name
        Target = $rname
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        MaxTasks = $maxtasks
        rType = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($rs in $Repository) {
      ForEach ($rp in $rs.Extent) {
        $r = $rp.Repository 
        # Refresh Repository Size Info
        [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)           
        $rType = switch ($r.Type) {
          "WinLocal" {"Windows Local"}
          "LinuxLocal" {"Linux Local"}
          "CifsShare" {"CIFS Share"}
          "DataDomain" {"Data Domain"}
          "ExaGrid" {"ExaGrid"}
          "HPStoreOnce" {"HP StoreOnce"}
          default {"Unknown"}     
        }
        $outputObj = Build-Object $rs.Name $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.info.CachedFreeSpace $r.Info.CachedTotalSpace $r.Options.MaxTaskCount $rType
        $outputAry += $outputObj
      }
    } 
  }
  End {
    $outputAry
  }
}

function Get-RepoPermissions {
  $outputAry = @()
  $repoEPPerms = $script:repoList | get-vbreppermission
  $repoEPPermsSo = $script:repoListSo | get-vbreppermission
  ForEach ($repo in $repoEPPerms) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = (Get-VBRBackupRepository | where {$_.Id -eq $repo.RepositoryId}).Name
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  ForEach ($repo in $repoEPPermsSo) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = "[SO] $((Get-VBRBackupRepository -ScaleOut | where {$_.Id -eq $repo.RepositoryId}).Name)"
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  $outputAry
}

Function Get-VBRReplicaTarget {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline=$true)]
    [PSObject[]]$InputObj
  )
  BEGIN {
    $outputAry = @()
    $dsAry = @()
    If (($Name -ne $null) -and ($InputObj -eq $null)) {
      $InputObj = Get-VBRJob -Name $Name
    }
  }
  PROCESS {
    Foreach ($obj in $InputObj) {
      If (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
        $esxi = $obj.GetTargetHost()
        $dtstr =  $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName    
        $objoutput = New-Object -TypeName PSObject -Property @{
          Target = $esxi.Name
          Datastore = $obj.ViReplicaTargetOptions.DatastoreName
          StorageFree = [Math]::Round([Decimal]$dtstr.FreeSpace/1GB,2)
          StorageTotal = [Math]::Round([Decimal]$dtstr.Capacity/1GB,2)
          FreePercentage = [Math]::Round(($dtstr.FreeSpace/$dtstr.Capacity)*100)
        }
        $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
        $outputAry = $outputAry + $objoutput
      } Else {
        return
      }
    }
  }
  END {
    $outputAry | Select Target, Datastore, StorageFree, StorageTotal, FreePercentage
  }
}
 
Function Get-VeeamVersion {
  Try {
    $veeamExe = Get-Item $veeamExePath
    $VeeamVersion = $veeamExe.VersionInfo.ProductVersion
    Return $VeeamVersion
  } Catch {
    Write-Host "Unable to Locate Veeam executable, check path - $veeamExePath" -ForegroundColor Red
    exit  
  }
} 
 
Function Get-VeeamSupportDate {
  param (
    [string]$vbrServer
  ) 
  # Query for license info
  Try{
    $license = Get-VBRInstalledLicense
    $expirationDate = $license.ExpirationDate
    #$wmi = get-wmiobject -list "StdRegProv" -namespace root\default -computername $vbrServer -ErrorAction Stop
    #$hklm = 2147483650
    #$bKey = "SOFTWARE\Veeam\Veeam Backup and Replication\license"
    #$bValue = "Lic1"
    #$regBinary = ($wmi.GetBinaryValue($hklm, $bKey, $bValue)).uValue
    #$regBinary = (Get-Item 'HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication\license').GetValue('Lic1')
    
    #$veeamLicInfo = [string]::Join($null, ($regBinary | % { [char][int]$_; }))
    # Convert Binary key
    #$pattern = "expiration date\=\d{1,2}\/\d{1,2}\/\d{1,4}"
    #$expirationDate = [regex]::matches($VeeamLicInfo, $pattern)[0].Value.Split("=")[1]
    #$datearray = $expirationDate -split '/'
    #$expirationDate = Get-Date -Day $datearray[0] -Month $datearray[1] -Year $datearray[2]
    $totalDaysLeft = ($expirationDate - (get-date)).Totaldays.toString().split(",")[0]
    $totalDaysLeft = [int]$totalDaysLeft
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate = $expirationDate.ToShortDateString()
      DaysRemain = $totalDaysLeft
    }
  } Catch{
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate = "License Date Check Failed"
      DaysRemain = "License Date Check Failed"
    }
  }
  $objoutput
} 

Function Get-VeeamWinServers {
  $vservers=@{}
  $outputAry = @()
  $vservers.add($($script:vbrServerObj.Name),"VBRServer")
  Foreach ($srv in $script:proxyList) {
    If (!$vservers.ContainsKey($srv.Host.Name)) {
      $vservers.Add($srv.Host.Name,"ProxyServer")
    }
  }
  Foreach ($srv in $script:repoList) {
    If ($srv.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($srv.gethost().Name)) {
      $vservers.Add($srv.gethost().Name,"RepoServer")
    }
  }
  Foreach ($rs in $script:repoListSo) {
    ForEach ($rp in $rs.Extent) {
      $r = $rp.Repository 
      $rName = $($r.GetHost()).Name
      If ($r.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($rName)) {
        $vservers.Add($rName,"RepoSoServer")
      }
    }
  }  
  Foreach ($srv in $script:tapesrvList) {
    If (!$vservers.ContainsKey($srv.Name)) {
      $vservers.Add($srv.Name,"TapeServer")
    }
  }  
  $vservers = $vservers.GetEnumerator() | Sort-Object Name
  Foreach ($vserver in $vservers) {
    $outputAry += $vserver.Name
  }
  return $outputAry
}

Function Get-VeeamServices {
  param (
    [PSObject]$inputObj
  )   
  $outputAry = @()
  Foreach ($obj in $InputObj) {    
    $output = @()
    Try {
      $output = Get-Service -computername $obj -Name "*Veeam*" -exclude "SQLAgent*" |
        Select @{Name="Server Name"; Expression = {$obj.ToLower()}}, @{Name="Service Name"; Expression = {$_.DisplayName}}, Status
    } Catch {
      $output = New-Object PSObject -Property @{
        "Server Name" = $obj.ToLower()
        "Service Name" = "Unable to connect"
        Status = "Unknown"
      }
    }   
    $outputAry += $output  
  }
  $outputAry
}

Function Get-VMsBackupStatus {
  $outputary = @()
  # Convert exclusion list to simple regular expression
  $excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $vms=@{}
  # Build a hash table of all VMs.  Key is either Job Object Id (for any VM ever in a Veeam job) or vCenter ID+MoRef
  # Assume unprotected (!), and populate Cluster, DataCenter, and Name fields for hash key value
  Find-VBRViEntity | 
    Where-Object {$_.Type -eq "Vm" -and $_.VmFolderName -notmatch $excludefolder_regex} |
    Where-Object {$_.Name -notmatch $excludevms_regex} |
    Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
    ForEach {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.Path.Split("\")[2], $_.Name, "1/11/1911", "1/11/1911","", $_.VmFolderName))}
  If (!$script:excludeTemp) {
    Find-VBRViEntity -VMsandTemplates |
      Where-Object {$_.Type -eq "Vm" -and $_.IsTemplate -eq "True" -and $_.VmFolderName -notmatch $excludefolder_regex} |
      Where-Object {$_.Name -notmatch $excludevms_regex} |
      Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
      ForEach {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.VmHostName, "[template] $($_.Name)", "1/11/1911", "1/11/1911","", $_.VmFolderName))}
  }
  # Find all backup task sessions that have ended in the last x hours
  $vbrtasksessions = (Get-VBRBackupSession | 
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
    Get-VBRTaskSession | Where-Object {$_.Status -notmatch "InProgress|Pending"}
  # Compare VM list to session list and update found VMs status
  If ($vbrtasksessions) {
    Foreach ($vmtask in $vbrtasksessions) {
      If($vms.ContainsKey($vmtask.Info.ObjectId)) {
        If ((Get-Date $vmtask.Progress.StartTimeLocal) -ge (Get-Date $vms[$vmtask.Info.ObjectId][5])) {
          If ($vmtask.Status -eq "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=""
          } ElseIf ($vms[$vmtask.Info.ObjectId][0] -ne "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=($vmtask.GetDetails()).Replace("<br />","ZZbrZZ")
          }
        } ElseIf ($vms[$vmtask.Info.ObjectId][0] -match "Warning|Failed" -and $vmtask.Status -eq "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=""
        }
      }
    }
  }    
  Foreach ($vm in $vms.GetEnumerator()) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Status = $vm.Value[0]
      Name = $vm.Value[4]
      vCenter = $vm.Value[1]
      Datacenter = $vm.Value[2]
      Cluster = $vm.Value[3]
      StartTime = $vm.Value[5]
      StopTime = $vm.Value[6]
      Details = $vm.Value[7]
      Folder = $vm.Value[8]
    }
    $outputAry += $objoutput
  }
  $outputAry
}

function Get-Duration {
  param ($ts)
  $days = ""
  If ($ts.Days -gt 0) {
    $days = "{0}:" -f $ts.Days
  }
  "{0}{1}:{2,2:D2}:{3,2:D2}" -f $days,$ts.Hours,$ts.Minutes,$ts.Seconds
}

function Get-BackupSize {
  param ($backups)
  $outputObj = @()
  Foreach ($backup in $backups) {
    $backupSize = 0
    $dataSize = 0
    $files = $backup.GetAllStorages()
    Foreach ($file in $Files) {
      $backupSize += [math]::Round([long]$file.Stats.BackupSize/1GB, 2)
      $dataSize += [math]::Round([long]$file.Stats.DataSize/1GB, 2)
    }         
    $repo = If ($($script:repoList | Where {$_.Id -eq $backup.RepositoryId}).Name) {
              $($script:repoList | Where {$_.Id -eq $backup.RepositoryId}).Name
            } Else {
              $($script:repoListSo | Where {$_.Id -eq $backup.RepositoryId}).Name
            }
    $vbrMasterHash = @{
      JobName = $backup.JobName
      VMCount = $backup.VmCount
      Repo = $repo
      DataSize = $dataSize
      BackupSize = $backupSize
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    $outputObj += $vbrMasterObj
  }
  $outputObj
}
Function Get-MultiJobs {
  $outputAry = @()
  $vmMultiJobs = (Get-VBRBackupSession | 
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) | 
    Get-VBRTaskSession | Select Name, @{Name="VMID"; Expression = {$_.Info.ObjectId}}, JobName -Unique | Group Name, VMID | where {$_.Count -gt 1} | Select -ExpandProperty Group
  ForEach ($vm in $vmMultiJobs) {
    $objID = $vm.VMID
    $viEntity = Find-VBRViEntity -name $vm.Name | Where {$_.FindObject().Id -eq $objID}  
    If ($viEntity -ne $null) {
      $objoutput = New-Object -TypeName PSObject -Property @{
        Name = $vm.Name
        vCenter = $viEntity.Path.Split("\")[0]
        Datacenter = $viEntity.Path.Split("\")[1]
        Cluster = $viEntity.Path.Split("\")[2]
        Folder = $viEntity.VMFolderName
        JobName = $vm.JobName
      }
      $outputAry += $objoutput
    } Else { #assume Template
      $viEntity = Find-VBRViEntity -VMsAndTemplates -name $vm.Name | Where {$_.FindObject().Id -eq $objID}
      If ($viEntity -ne $null) {
        $objoutput = New-Object -TypeName PSObject -Property @{
          Name = "[template] " + $vm.Name
          vCenter = $viEntity.Path.Split("\")[0]
          Datacenter = $viEntity.Path.Split("\")[1]
          Cluster = $viEntity.VmHostName
          Folder = $viEntity.VMFolderName
          JobName = $vm.JobName
        }
      }
      If ($objoutput) {
        $outputAry += $objoutput
      }    
    }
  }  
  $outputAry
}
#endregion
 
#region Report
# Get Veeam Version
$VeeamVersion = Get-VeeamVersion

# HTML Stuff
$headerObj = @"
<html>
    <head>
        <title>$rptTitle</title>
            <style>  
              body {font-family: Tahoma; background-color:#ffffff;}
              table {font-family: Tahoma;width: $($rptWidth)%;font-size: 12px;border-collapse:collapse;}
              <!-- table tr:nth-child(odd) td {background: #e2e2e2;} -->
              th {background-color: #e2e2e2;border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff;border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@
 
$bodyTop = @"
    <body>
        <center>
            <table>
                <tr>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 0px 0px;">VBR v$VeeamVersion</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 2px 5px;">$rptMode</td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">MVR v$MVRversion</td>
                </tr>
            </table>
"@
 
$subHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead02 = @"
</td>
                </tr>
             </table>
"@

$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
						    </tr>
            </table>
"@

$footerObj = @"
<table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">My Veeam Report maintained at <a href="https://github.com/raynorpat/ps_scripts/tree/master/VeeamReport" target="_blank">https://github.com/raynorpat/ps_scripts</a></td>
                </tr>
            </table>
        </center>
    </body>
</html>
"@

#Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showProtectedVMs) {
  $vmStatus = Get-VMsBackupStatus
}
# VMs Missing Backups
$missingVMs = @($vmStatus | ?{$_.Status -match "!|Failed"})
ForEach ($VM in $missingVMs) {
  If ($VM.Status -eq "!") {
    $VM.Details = "No Backup Task has completed"
    $VM.StartTime = ""
    $VM.StopTime = ""
  }
}
# VMs Successfuly Backed Up
$successVMs = @($vmStatus | ?{$_.Status -eq "Success"})
# VMs Backed Up w/Warning
$warnVMs = @($vmStatus | ?{$_.Status -eq "Warning"})

# Get VM Backup Protection Summary
$bodySummaryProtect = $null
$sumprotectHead = $subHead01
If ($showSummaryProtect) {
  If (@($successVMs).Count -ge 1) {
    $percentProt = 1
    $sumprotectHead = $subHead01suc
  }
  If (@($warnVMs).Count -ge 1) {
    $percentWarn = "*"
    $sumprotectHead = $subHead01war
  } Else {
    $percentWarn = ""
  }
  If (@($missingVMs).Count -ge 1) {
    $percentProt = (@($warnVMs).Count + @($successVMs).Count) / (@($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count)
    $sumprotectHead = $subHead01err
  }  
  $vbrMasterHash = @{
    WarningVM = @($warnVMs).Count
    ProtectedVM = @($successVMs).Count
    FailedVM = @($missingVMs).Count
    PercentProt = "{0:P2}{1}" -f $percentProt,$percentWarn
    
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  $summaryProtect =  $vbrMasterObj | Select @{Name="% Protected"; Expression = {$_.PercentProt}},
    @{Name="Fully Protected VMs"; Expression = {$_.ProtectedVM}},
    @{Name="Protected VMs w/Warnings"; Expression = {$_.WarningVM}},
    @{Name="Unprotected VMs"; Expression = {$_.FailedVM}}
  $bodySummaryProtect = $summaryProtect | ConvertTo-HTML -Fragment
  $bodySummaryProtect = $sumprotectHead + "VM Backup Protection Summary" + $subHead02 + $bodySummaryProtect
}

# Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs) {  
  If ($missingVMs -ne $null) {
    $missingVMs = $missingVMs | Sort vCenter, Datacenter, Cluster, Name | Select Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-HTML -Fragment
    $bodyMissing = $subHead01err + "VMs with No Successful Backups within RPO" + $subHead02 + $missingVMs
  }
}

# Get VMs Backed Up w/Warnings
$bodyWarning = $null
If ($showProtectedVMs) {    
  If ($warnVMs -ne $null) {
    $warnVMs = $warnVMs | Sort vCenter, Datacenter, Cluster, Name | Select Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-HTML -Fragment
    $bodyWarning = $subHead01war + "VMs with only Backups with Warnings within RPO" + $subHead02 + $warnVMs
  }
}

# Get VMs Successfuly Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {    
  If ($successVMs -ne $null) {
    $successVMs = $successVMs | Sort vCenter, Datacenter, Cluster, Name | Select Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}} | ConvertTo-HTML -Fragment
    $bodySuccess = $subHead01suc + "VMs with Successful Backups within RPO" + $subHead02 + $successVMs
  }
}

# Get VMs Backed Up by Multiple Jobs
$bodyMultiJobs = $null
If ($showMultiJobs) {    
  $multiJobs = @(Get-MultiJobs)
  If ($multiJobs.Count -gt 0) {
    $bodyMultiJobs = $multiJobs | Sort vCenter, Datacenter, Cluster, Name | Select Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Job Name"; Expression = {$_.JobName}} | ConvertTo-HTML -Fragment
    $bodyMultiJobs = $subHead01err + "VMs Backed Up by Multiple Jobs within RPO" + $subHead02 + $bodyMultiJobs
  }
}

# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsBk).Count
    "Sessions" = If ($sessListBk) {@($sessListBk).Count} Else {0}
    "Read" = $totalReadBk
    "Transferred" = $totalXferBk
    "Successful" = @($successSessionsBk).Count
    "Warning" = @($warningSessionsBk).Count
    "Fails" = @($failsSessionsBk).Count
    "Running" = @($runningSessionsBk).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBk) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBk =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryBk = $arrSummaryBk | ConvertTo-HTML -Fragment
  If ($arrSummaryBk.Failed -gt 0) {
      $summaryBkHead = $subHead01err
  } ElseIf ($arrSummaryBk.Warnings -gt 0) {
      $summaryBkHead = $subHead01war
  } ElseIf ($arrSummaryBk.Successful -gt 0) {
      $summaryBkHead = $subHead01suc
  } Else {
      $summaryBkHead = $subHead01
  }
  $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
}

# Get Backup Job Status
$bodyJobsBk = $null
If ($showJobsBk) {
  If ($allJobsBk.count -gt 0) {
    $bodyJobsBk = @()
    Foreach($bkJob in $allJobsBk) {
      $bodyJobsBk += $bkJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($bkJob.IsRunning) {
            $currentSess = $runningSessionsBk | ?{$_.JobName -eq $bkJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name) {
            $($repoList | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          } Else {
            $($repoListSo | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          }
        }},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){"Unknown"}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBk = $bodyJobsBk | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
  }
}

# Get Backup Job Size
$bodyJobSizeBk = $null
If ($showBackupSizeBk) {
  If ($backupsBk.count -gt 0) {
    $bodyJobSizeBk = Get-BackupSize -backups $backupsBk | Sort JobName | Select @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBk = $subHead01 + "Backup Job Size" + $subHead02 + $bodyJobSizeBk
  }
}

# Get all Backup Sessions
$bodyAllSessBk = $null
If ($showAllSessBk) {
  If ($sessListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllSessBk = $sessListBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk  | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }      
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    } Else {
      $arrAllSessBk = $sessListBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    }
  }
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
  If ($runningSessionsBk.count -gt 0) {
    $bodyRunningBk = $runningSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
  }
} 

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
  $sessWF = @($warningSessionsBk + $failsSessionsBk)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBk) {
      $headerWF = "Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedBk) {
      $arrSessWFBk = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }      
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    } Else {
      $arrSessWFBk = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }      
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    }
  }
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
  If ($successSessionsBk.count -gt 0) {
    If ($onlyLastBk) {
      $headerSucc = "Successful Backup Jobs"
    } Else {
      $headerSucc = "Successful Backup Sessions"
    }
    If ($showDetailedBk) {
      $bodySessSuccBk = $successSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    } Else {
      $bodySessSuccBk = $successSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Tasks from Sessions within time frame
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | ?{$_.Status -eq "Success"})
$wfTasksBk = @($taskListBk | ?{$_.Status -match "Warning|Failed"})
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | ?{$_.Status -match "Pending|InProgress"}

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
  If ($taskListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllTasksBk = $taskListBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }      
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    } Else {
      $arrAllTasksBk = $taskListBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }      
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    }
  }
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
  If ($runningTasksBk.count -gt 0) {
    $bodyTasksRunningBk = $runningTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
  }
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
  If ($wfTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrTaskWFBk = $wfTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }      
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    } Else {
      $arrTaskWFBk = $wfTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }      
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    }
  }
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
  If ($successTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $bodyTaskSuccBk = $successTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    } Else {
      $bodyTaskSuccBk = $successTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    }
  }
}

# Get Running VM Restore Sessions
$bodyRestoRunVM = $null
If ($showRestoRunVM) {
  If ($($runningResto).count -gt 0) {
    $bodyRestoRunVM = $runningResto | Sort CreationTime | Select @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}}, @{Name="Start Time"; Expression = {$_.CreationTime}},        
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}},
      @{Name="Reason"; Expression = {$_.Info.Reason}} | ConvertTo-HTML -Fragment
    $bodyRestoRunVM = $subHead01 + "Running VM Restore Sessions" + $subHead02 + $bodyRestoRunVM 
  }
}

# Get Completed VM Restore Sessions
$bodyRestoreVM = $null
If ($showRestoreVM) {
  If ($($completeResto).count -gt 0) {
    $arrRestoreVM = $completeResto | Sort CreationTime | Select @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}},
      @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},        
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},        
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}}, @{Name="Reason"; Expression = {$_.Info.Reason}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyRestoreVM = $arrRestoreVM | ConvertTo-HTML -Fragment
    If ($arrRestoreVM.Result -match "Failed") {
      $restoreVMHead = $subHead01err
    } ElseIf ($arrRestoreVM.Result -match "Warning") {
      $restoreVMHead = $subHead01war
    } ElseIf ($arrRestoreVM.Result -match "Success") {
      $restoreVMHead = $subHead01suc
    } Else {
      $restoreVMHead = $subHead01
    }    
    $bodyRestoreVM = $restoreVMHead + "Completed VM Restore Sessions" + $subHead02 + $bodyRestoreVM 
  }
}

# Get Replication Summary Info
$bodySummaryRp = $null
If ($showSummaryRp) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsRp).Count
    "Sessions" = If ($sessListRp) {@($sessListRp).Count} Else {0}
    "Read" = $totalReadRp
    "Transferred" = $totalXferRp
    "Successful" = @($successSessionsRp).Count
    "Warning" = @($warningSessionsRp).Count
    "Fails" = @($failsSessionsRp).Count
    "Running" = @($runningSessionsRp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastRp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryRp =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryRp = $arrSummaryRp | ConvertTo-HTML -Fragment
  If ($arrSummaryRp.Failed -gt 0) {
      $summaryRpHead = $subHead01err
  } ElseIf ($arrSummaryRp.Warnings -gt 0) {
      $summaryRpHead = $subHead01war
  } ElseIf ($arrSummaryRp.Successful -gt 0) {
      $summaryRpHead = $subHead01suc
  } Else {
      $summaryRpHead = $subHead01
  }
  $bodySummaryRp = $summaryRpHead + "Replication Results Summary" + $subHead02 + $bodySummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
If ($showJobsRp) {
  If ($allJobsRp.count -gt 0) {
    $bodyJobsRp = @()
    Foreach($rpJob in $allJobsRp) {
      $bodyJobsRp += $rpJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($rpJob.IsRunning) {
            $currentSess = $runningSessionsRp | ?{$_.JobName -eq $rpJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Info.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
         }},
        @{Name="Target"; Expression = {$(Get-VBRServer | Where {$_.Id -eq $rpJob.Info.TargetHostId}).Name}},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $rpJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsRp = $bodyJobsRp | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsRp = $subHead01 + "Replication Job Status" + $subHead02 + $bodyJobsRp
  }
}

# Get Replication Sessions
$bodyAllSessRp = $null
If ($showAllSessRp) {
  If ($sessListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllSessRp = $sessListRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }      
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    } Else {
      $arrAllSessRp = $sessListRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    }
  }
}

# Get Running Replication Jobs
$bodyRunningRp = $null
If ($showRunningRp) {
  If ($runningSessionsRp.count -gt 0) {
    $bodyRunningRp = $runningSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningRp = $subHead01 + "Running Replication Jobs" + $subHead02 + $bodyRunningRp
  }
} 

# Get Replication Sessions with Warnings or Failures
$bodySessWFRp = $null
If ($showWarnFailRp) {
  $sessWF = @($warningSessionsRp + $failsSessionsRp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastRp) {
      $headerWF = "Replication Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Replication Sessions with Warnings or Failures"
    }
    If ($showDetailedRp) {
      $arrSessWFRp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    } Else {
      $arrSessWFRp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    }
  }
}

# Get Successful Replication Sessions
$bodySessSuccRp = $null
If ($showSuccessRp) {
  If ($successSessionsRp.count -gt 0) {
    If ($onlyLastRp) {
      $headerSucc = "Successful Replication Jobs"
    } Else {
      $headerSucc = "Successful Replication Sessions"
    }
    If ($showDetailedRp) {
      $bodySessSuccRp = $successSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    } Else {
      $bodySessSuccRp = $successSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Replication Tasks from Sessions within time frame
$taskListRp = @()
$taskListRp += $sessListRp | Get-VBRTaskSession
$successTasksRp = @($taskListRp | ?{$_.Status -eq "Success"})
$wfTasksRp = @($taskListRp | ?{$_.Status -match "Warning|Failed"})
$runningTasksRp = @()
$runningTasksRp += $runningSessionsRp | Get-VBRTaskSession | ?{$_.Status -match "Pending|InProgress"}

# Get Replication Tasks
$bodyAllTasksRp = $null
If ($showAllTasksRp) {
  If ($taskListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllTasksRp = $taskListRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    } Else {
      $arrAllTasksRp = $taskListRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    }
  }
}

# Get Running Replication Tasks
$bodyTasksRunningRp = $null
If ($showRunningTasksRp) {
  If ($runningTasksRp.count -gt 0) {
    $bodyTasksRunningRp = $runningTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningRp = $subHead01 + "Running Replication Tasks" + $subHead02 + $bodyTasksRunningRp
  }
}

# Get Replication Tasks with Warnings or Failures
$bodyTaskWFRp = $null
If ($showTaskWFRp) {
  If ($wfTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrTaskWFRp = $wfTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    } Else {
      $arrTaskWFRp = $wfTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    }
  }
}

# Get Successful Replication Tasks
$bodyTaskSuccRp = $null
If ($showTaskSuccessRp) {
  If ($successTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $bodyTaskSuccRp = $successTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    } Else {
      $bodyTaskSuccRp = $successTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    }
  }
}

# Get Backup Copy Summary Info
$bodySummaryBc = $null
If ($showSummaryBc) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListBc) {@($sessListBc).Count} Else {0}
    "Read" = $totalReadBc
    "Transferred" = $totalXferBc
    "Successful" = @($successSessionsBc).Count
    "Warning" = @($warningSessionsBc).Count
    "Fails" = @($failsSessionsBc).Count
    "Working" = @($workingSessionsBc).Count
    "Idle" = @($idleSessionsBc).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBc) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBc =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryBc = $arrSummaryBc | ConvertTo-HTML -Fragment
  If ($arrSummaryBc.Failures -gt 0) {
      $summaryBcHead = $subHead01err
  } ElseIf ($arrSummaryBc.Warnings -gt 0) {
      $summaryBcHead = $subHead01war
  } ElseIf ($arrSummaryBc.Successful -gt 0) {
      $summaryBcHead = $subHead01suc
  } Else {
      $summaryBcHead = $subHead01
  }
  $bodySummaryBc = $summaryBcHead + "Backup Copy Results Summary" + $subHead02 + $bodySummaryBc
}

# Get Backup Copy Job Status
$bodyJobsBc = $null
If ($showJobsBc) {
  If ($allJobsBc.count -gt 0) {
    $bodyJobsBc = @()
    Foreach($BcJob in $allJobsBc) {
      $bodyJobsBc += $BcJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Type"; Expression = {$_.TypeToString}},             
        @{Name="Status"; Expression = {
          If ($BcJob.IsRunning) {
            $currentSess = $BcJob.FindLastSession()
            If ($currentSess.State -eq "Working") {
              $csessPercent = $currentSess.Progress.Percents
              $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
              $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
              $cStatus
            } Else {
              $currentSess.State
            }
          } Else {
            "Stopped"
          }             
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $BcJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBc = $bodyJobsBc | Sort "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsBc = $subHead01 + "Backup Copy Job Status" + $subHead02 + $bodyJobsBc
  }
}

# Get Backup Copy Job Size
$bodyJobSizeBc = $null
If ($showBackupSizeBc) {
  If ($backupsBc.count -gt 0) {
    $bodyJobSizeBc = Get-BackupSize -backups $backupsBc | Sort JobName | Select @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBc = $subHead01 + "Backup Copy Job Size" + $subHead02 + $bodyJobSizeBc
  }
}

# Get All Backup Copy Sessions
$bodyAllSessBc = $null
If ($showAllSessBc) {
  If ($sessListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllSessBc = $sessListBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    } Else {
      $arrAllSessBc = $sessListBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    }
  }
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
  If ($idleSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerIdle = "Idle Backup Copy Jobs"
    } Else {
      $headerIdle = "Idle Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessIdleBc = $idleSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},                 
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    } Else {
      $bodySessIdleBc = $idleSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    }
  }
}

# Get Working Backup Copy Jobs
$bodyRunningBc = $null
If ($showRunningBc) {
  If ($workingSessionsBc.count -gt 0) {
    $bodyRunningBc = $workingSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBc = $subHead01 + "Working Backup Copy Sessions" + $subHead02 + $bodyRunningBc
  }
}

# Get Backup Copy Sessions with Warnings or Failures
$bodySessWFBc = $null
If ($showWarnFailBc) {
  $sessWF = @($warningSessionsBc + $failsSessionsBc)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBc) {
      $headerWF = "Backup Copy Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Copy Sessions with Warnings or Failures"
    }
    If ($showDetailedBc) {
      $arrSessWFBc = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    } Else {
      $arrSessWFBc = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    }
  }
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
  If ($successSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerSucc = "Successful Backup Copy Jobs"
    } Else {
      $headerSucc = "Successful Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessSuccBc = $successSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    } Else {
      $bodySessSuccBc = $successSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Copy Tasks from Sessions within time frame
$taskListBc = @()
$taskListBc += $sessListBc | Get-VBRTaskSession
$successTasksBc = @($taskListBc | ?{$_.Status -eq "Success"})
$wfTasksBc = @($taskListBc | ?{$_.Status -match "Warning|Failed"})
$pendingTasksBc = @($taskListBc | ?{$_.Status -eq "Pending"})
$runningTasksBc = @($taskListBc | ?{$_.Status -eq "InProgress"})

# Get All Backup Copy Tasks
$bodyAllTasksBc = $null
If ($showAllTasksBc) {
  If ($taskListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllTasksBc = $taskListBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    } Else {
      $arrAllTasksBc = $taskListBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    }
  }
}

# Get Pending Backup Copy Tasks
$bodyTasksPendingBc = $null
If ($showPendingTasksBc) {
  If ($pendingTasksBc.count -gt 0) {
    $bodyTasksPendingBc = $pendingTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingBc = $subHead01 + "Pending Backup Copy Tasks" + $subHead02 + $bodyTasksPendingBc
  }
}

# Get Working Backup Copy Tasks
$bodyTasksRunningBc = $null
If ($showRunningTasksBc) {
  If ($runningTasksBc.count -gt 0) {
    $bodyTasksRunningBc = $runningTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBc = $subHead01 + "Working Backup Copy Tasks" + $subHead02 + $bodyTasksRunningBc
  }
}

# Get Backup Copy Tasks with Warnings or Failures
$bodyTaskWFBc = $null
If ($showTaskWFBc) {
  If ($wfTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrTaskWFBc = $wfTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    } Else {
      $arrTaskWFBc = $wfTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    }
  }
}

# Get Successful Backup Copy Tasks
$bodyTaskSuccBc = $null
If ($showTaskSuccessBc) {
  If ($successTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $bodyTaskSuccBc = $successTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    } Else {
      $bodyTaskSuccBc = $successTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    }
  }
}

# Get Tape Backup Summary Info
$bodySummaryTp = $null
If ($showSummaryTp) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListTp) {@($sessListTp).Count} Else {0}
    "Read" = $totalReadTp
    "Transferred" = $totalXferTp
    "Successful" = @($successSessionsTp).Count
    "Warning" = @($warningSessionsTp).Count
    "Fails" = @($failsSessionsTp).Count
    "Working" = @($workingSessionsTp).Count
    "Idle" = @($idleSessionsTp).Count
    "Waiting" = @($waitingSessionsTp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastTp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryTp =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}}, @{Name="Waiting"; Expression = {$_.Waiting}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryTp = $arrSummaryTp | ConvertTo-HTML -Fragment
  If ($arrSummaryTp.Failures -gt 0) {
      $summaryTpHead = $subHead01err
  } ElseIf ($arrSummaryTp.Warnings -gt 0 -or $arrSummaryTp.Waiting -gt 0) {
      $summaryTpHead = $subHead01war
  } ElseIf ($arrSummaryTp.Successful -gt 0) {
      $summaryTpHead = $subHead01suc
  } Else {
      $summaryTpHead = $subHead01
  }
  $bodySummaryTp = $summaryTpHead + "Tape Backup Results Summary" + $subHead02 + $bodySummaryTp
}

# Get Tape Backup Job Status
$bodyJobsTp = $null
If ($showJobsTp) {
  If ($allJobsTp.count -gt 0) {
    $bodyJobsTp = @()
    Foreach($tpJob in $allJobsTp) {
      $bodyJobsTp += $tpJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Job Type"; Expression = {$_.Type}},@{Name="Media Pool"; Expression = {$_.Target}},
        @{Name="Status"; Expression = {$_.LastState}},
        @{Name="Next Run"; Expression = {
          If ($_.ScheduleOptions.Type -eq "AfterNewBackup") {"<Continious>"}
          ElseIf ($_.ScheduleOptions.Type -eq "AfterJob") {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $tpJob.ScheduleOptions.JobId}).Name + "]"}
          ElseIf ($_.NextRun) {$_.NextRun}
          Else {"<not scheduled>"}}},
        @{Name="Last Result"; Expression = {If ($_.LastResult -eq "None"){""}Else{$_.LastResult}}}
    }
    $bodyJobsTp = $bodyJobsTp | Sort "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsTp = $subHead01 + "Tape Backup Job Status" + $subHead02 + $bodyJobsTp
  }
}

# Get Tape Backup Sessions
$bodyAllSessTp = $null
If ($showAllSessTp) {
  If ($sessListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllSessTp = $sessListTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }      
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    } Else {
      $arrAllSessTp = $sessListTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }      
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    }
  
    # Due to issue with getting details on tape sessions, we may need to get session info again :-(
    If (($showWaitingTp -or $showIdleTp -or $showRunningTp -or $showWarnFailTp -or $showSuccessTp) -and $showDetailedTp) {
      # Get all Tape Backup Sessions
      $allSessTp = @()
      Foreach ($tpJob in $allJobsTp){
        $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
        $allSessTp += $tpSessions
      }
      # Gather all Tape Backup Sessions within timeframe
      $sessListTp = @($allSessTp | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
      If ($tapeJob -ne $null -and $tapeJob -ne "") {
        $allJobsTpTmp = @()
        $sessListTpTmp = @()
        Foreach ($tpJob in $tapeJob) {
          $allJobsTpTmp += $allJobsTp | ?{$_.Name -like $tpJob}
          $sessListTpTmp += $sessListTp | ?{$_.JobName -like $tpJob}
        }
        $allJobsTp = $allJobsTpTmp | sort Id -Unique
        $sessListTp = $sessListTpTmp | sort Id -Unique
      }
      If ($onlyLastTp) {
        $tempSessListTp = $sessListTp
        $sessListTp = @()
        Foreach($job in $allJobsTp) {
          $sessListTp += $tempSessListTp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
        }
      }
      # Get Tape Backup Session information
      $idleSessionsTp = @($sessListTp | ?{$_.State -eq "Idle"})
      $successSessionsTp = @($sessListTp | ?{$_.Result -eq "Success"})
      $warningSessionsTp = @($sessListTp | ?{$_.Result -eq "Warning"})
      $failsSessionsTp = @($sessListTp | ?{$_.Result -eq "Failed"})
      $workingSessionsTp = @($sessListTp | ?{$_.State -eq "Working"})
      $waitingSessionsTp = @($sessListTp | ?{$_.State -eq "WaitingTape"})
    }
  }
}

# Get Waiting Tape Backup Jobs
$bodyWaitingTp = $null
If ($showWaitingTp) {
  If ($waitingSessionsTp.count -gt 0) {
    $bodyWaitingTp = $waitingSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyWaitingTp = $subHead01war + "Waiting Tape Backup Sessions" + $subHead02 + $bodyWaitingTp
  }
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
  If ($idleSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerIdle = "Idle Tape Backup Jobs"
    } Else {
      $headerIdle = "Idle Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessIdleTp = $idleSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},                 
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    } Else {
      $bodySessIdleTp = $idleSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    }
  }
}

# Get Working Tape Backup Jobs
$bodyRunningTp = $null
If ($showRunningTp) {
  If ($workingSessionsTp.count -gt 0) {
    $bodyRunningTp = $workingSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningTp = $subHead01 + "Working Tape Backup Sessions" + $subHead02 + $bodyRunningTp
  }
}

# Get Tape Backup Sessions with Warnings or Failures
$bodySessWFTp = $null
If ($showWarnFailTp) {
  $sessWF = @($warningSessionsTp + $failsSessionsTp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastTp) {
      $headerWF = "Tape Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Tape Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedTp) {
      $arrSessWFTp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    } Else {
      $arrSessWFTp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    }
  }
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
  If ($successSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerSucc = "Successful Tape Backup Jobs"
    } Else {
      $headerSucc = "Successful Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessSuccTp = $successSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    } Else {
      $bodySessSuccTp = $successSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Tape Backup Tasks from Sessions within time frame
$taskListTp = @()
$taskListTp += $sessListTp | Get-VBRTaskSession
$successTasksTp = @($taskListTp | ?{$_.Status -eq "Success"})
$wfTasksTp = @($taskListTp | ?{$_.Status -match "Warning|Failed"})
$pendingTasksTp = @($taskListTp | ?{$_.Status -eq "Pending"})
$runningTasksTp = @($taskListTp | ?{$_.Status -eq "InProgress"})

# Get Tape Backup Tasks
$bodyAllTasksTp = $null
If ($showAllTasksTp) {
  If ($taskListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllTasksTp = $taskListTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }  
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    } Else {
      $arrAllTasksTp = $taskListTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }  
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    }
  }
}

# Get Pending Tape Backup Tasks
$bodyTasksPendingTp = $null
If ($showPendingTasksTp) {
  If ($pendingTasksTp.count -gt 0) {
    $bodyTasksPendingTp = $pendingTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingTp = $subHead01 + "Pending Tape Backup Tasks" + $subHead02 + $bodyTasksPendingTp
  }
}

# Get Working Tape Backup Tasks
$bodyTasksRunningTp = $null
If ($showRunningTasksTp) {
  If ($runningTasksTp.count -gt 0) {
    $bodyTasksRunningTp = $runningTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningTp = $subHead01 + "Working Tape Backup Tasks" + $subHead02 + $bodyTasksRunningTp
  }
}

# Get Tape Backup Tasks with Warnings or Failures
$bodyTaskWFTp = $null
If ($showTaskWFTp) {
  If ($wfTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrTaskWFTp = $wfTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    } Else {
      $arrTaskWFTp = $wfTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    }
  }
}

# Get Successful Tape Backup Tasks
$bodyTaskSuccTp = $null
If ($showTaskSuccessTp) {
  If ($successTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $bodyTaskSuccTp = $successTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    } Else {
      $bodyTaskSuccTp = $successTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    }
  }
}

# Get all Tapes
$bodyTapes = $null
If ($showTapes) {
  $expTapes = @($mediaTapes)
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort Name | ConvertTo-HTML -Fragment
    $bodyTapes = $subHead01 + "All Tapes" + $subHead02 + $expTapes
  }
}

# Get all Tapes in each Custom Media Pool
$bodyTpPool = $null
If ($showTpMp) {
  ForEach ($mp in ($mediaPools | ?{$_.Type -eq "Custom"} | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.MediaPoolId -eq $mp.Id)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpPool += $subHead01 + "All Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes in each Vault
$bodyTpVlt = $null
If ($showTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.Location.VaultId -eq $vlt.Id)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | ?{$_.Id -eq $poolId}).Name     
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort Name | ConvertTo-HTML -Fragment
      $bodyTpVlt += $subHead01 + "All Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Expired Tapes
$bodyExpTp = $null
If ($showExpTp) {
  $expTapes = @($mediaTapes | where {($_.IsExpired -eq $True)})
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort Name | ConvertTo-HTML -Fragment
    $bodyExpTp = $subHead01 + "All Expired Tapes" + $subHead02 + $expTapes
  }
}

# Get Expired Tapes in each Custom Media Pool
$bodyTpExpPool = $null
If ($showExpTpMp) {
  ForEach ($mp in ($mediaPools | ?{$_.Type -eq "Custom"} | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.MediaPoolId -eq $mp.Id -and $_.IsExpired -eq $True)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpPool += $subHead01 + "Expired Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get Expired Tapes in each Vault
$bodyTpExpVlt = $null
If ($showExpTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.Location.VaultId -eq $vlt.Id -and $_.IsExpired -eq $True)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | ?{$_.Id -eq $poolId}).Name     
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpVlt += $subHead01 + "Expired Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes written to within time frame
$bodyTpWrt = $null
If ($showTpWrt) {
  $expTapes = @($mediaTapes | ?{$_.LastWriteTime -ge (Get-Date).AddHours(-$HourstoCheck)})
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort "Last Write" | ConvertTo-HTML -Fragment
    $bodyTpWrt = $subHead01 + "All Tapes Written" + $subHead02 + $expTapes
  }
}

# Get Agent Backup Summary Info
$bodySummaryEp = $null
If ($showSummaryEp) {
  $vbrEpHash = @{
    "Sessions" = If ($sessListEp) {@($sessListEp).Count} Else {0}
    "Successful" = @($successSessionsEp).Count
    "Warning" = @($warningSessionsEp).Count
    "Fails" = @($failsSessionsEp).Count
    "Running" = @($runningSessionsEp).Count
  }
  $vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
  If ($onlyLastEp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryEp =  $vbrEPObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryEp = $arrSummaryEp | ConvertTo-HTML -Fragment
  If ($arrSummaryEp.Failures -gt 0) {
      $summaryEpHead = $subHead01err
  } ElseIf ($arrSummaryEp.Warnings -gt 0) {
      $summaryEpHead = $subHead01war
  } ElseIf ($arrSummaryEp.Successful -gt 0) {
      $summaryEpHead = $subHead01suc
  } Else {
      $summaryEpHead = $subHead01
  }
  $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
}

# Get Agent Backup Job Status
$bodyJobsEp = $null
If ($showJobsEp) {
  If ($allJobsEp.count -gt 0) {
    $bodyJobsEp = $allJobsEp | Sort Name | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Description"; Expression = {$_.Description}},
      @{Name="Enabled"; Expression = {$_.JobEnabled}},
      @{Name="Retention Policy"; Expression = {$_.RetentionPolicy}}, @{Name="Backup Type"; Expression = {$_.BackupType}},
      @{Name="OS"; Expression = {$_.OSPlatform}} | ConvertTo-HTML -Fragment
    $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
  }
}

# Get Agent Backup Job Size
$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
  If ($backupsEp.count -gt 0) {
    $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort JobName | Select @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
  }
}

# Get Agent Backup Sessions
$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
  If ($sessListEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $arrAllSessEp += $sessListEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="State"; Expression = {$_.State}},@{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    }
    $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllSessEp.Result -match "Failed") {
        $allSessEpHead = $subHead01err
      } ElseIf ($arrAllSessEp.Result -match "Warning") {
        $allSessEpHead = $subHead01war
      } ElseIf ($arrAllSessEp.Result -match "Success") {
        $allSessEpHead = $subHead01suc
      } Else {
        $allSessEpHead = $subHead01
      }               
    $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
  }
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
  If ($runningSessionsEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $bodyRunningEp += $runningSessionsEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
    }               
    $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
  }
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
  $sessWFEp = @($warningSessionsEp + $failsSessionsEp)
  If ($sessWFEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerWFEp = "Agent Backup Jobs with Warnings or Failures"
    } Else {
      $headerWFEp = "Agent Backup Sessions with Warnings or Failures"
    }
    Foreach($job in $allJobsEp) {
      $arrSessWFEp += $sessWFEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrSessWFEp.Result -match "Failed") {
        $sessWFEpHead = $subHead01err
      } ElseIf ($arrSessWFEp.Result -match "Warning") {
        $sessWFEpHead = $subHead01war
      } ElseIf ($arrSessWFEp.Result -match "Success") {
        $sessWFEpHead = $subHead01suc
      } Else {
        $sessWFEpHead = $subHead01
      }             
    $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
  }
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
  If ($successSessionsEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerSuccEp = "Successful Agent Backup Jobs"
    } Else {
      $headerSuccEp = "Successful Agent Backup Sessions"
    }
    Foreach($job in $allJobsEp) {
      $bodySessSuccEp += $successSessionsEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment             
    $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
  }
}

# Get SureBackup Summary Info
$bodySummarySb = $null
If ($showSummarySb) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListSb) {@($sessListSb).Count} Else {0}
    "Successful" = @($successSessionsSb).Count
    "Warning" = @($warningSessionsSb).Count
    "Fails" = @($failsSessionsSb).Count
    "Running" = @($runningSessionsSb).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastSb) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummarySb =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummarySb = $arrSummarySb | ConvertTo-HTML -Fragment
  If ($arrSummarySb.Failures -gt 0) {
      $summarySbHead = $subHead01err
  } ElseIf ($arrSummarySb.Warnings -gt 0) {
      $summarySbHead = $subHead01war
  } ElseIf ($arrSummarySb.Successful -gt 0) {
      $summarySbHead = $subHead01suc
  } Else {
      $summarySbHead = $subHead01
  }
  $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
}

# Get SureBackup Job Status
$bodyJobsSb = $null
If ($showJobsSb) {
  If ($allJobsSb.count -gt 0) {
    $bodyJobsSb = @()
    Foreach($SbJob in $allJobsSb) {
      $bodyJobsSb += $SbJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($_.GetLastState() -eq "Working") {
            $currentSess = $_.FindLastSession()
            $csessPercent = $currentSess.CompletionPercentage
            $cStatus = "$($csessPercent)% completed"
            $cStatus
          } Else {
            $_.GetLastState()
          }             
        }},
        @{Name="Virtual Lab"; Expression = {$(Get-VSBVirtualLab | Where {$_.Id -eq $SbJob.VirtualLabId}).Name}},
        @{Name="Linked Jobs"; Expression = {$($_.GetLinkedJobs()).Name -join ","}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $SbJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.GetLastResult() -eq "None"){""}Else{$_.GetLastResult()}}}
    }
    $bodyJobsSb = $bodyJobsSb | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
  }
}

# Get SureBackup Sessions
$bodyAllSessSb = $null
If ($showAllSessSb) {
  If ($sessListSb.count -gt 0) {
    $arrAllSessSb = $sessListSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    $bodyAllSessSb = $arrAllSessSb | ConvertTo-HTML -Fragment
    If ($arrAllSessSb.Result -match "Failed") {
        $allSessSbHead = $subHead01err
      } ElseIf ($arrAllSessSb.Result -match "Warning") {
        $allSessSbHead = $subHead01war
      } ElseIf ($arrAllSessSb.Result -match "Success") {
        $allSessSbHead = $subHead01suc
      } Else {
        $allSessSbHead = $subHead01
      }
    $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
    }
}

# Get Running SureBackup Jobs
$bodyRunningSb = $null
If ($showRunningSb) {
  If ($runningSessionsSb.count -gt 0) {
    $bodyRunningSb = $runningSessionsSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
      @{Name="% Complete"; Expression = {$_.Progress}} | ConvertTo-HTML -Fragment
    $bodyRunningSb = $subHead01 + "Running SureBackup Jobs" + $subHead02 + $bodyRunningSb
  }
} 

# Get SureBackup Sessions with Warnings or Failures
$bodySessWFSb = $null
If ($showWarnFailSb) {
  $sessWF = @($warningSessionsSb + $failsSessionsSb)
  If ($sessWF.count -gt 0) {
    If ($onlyLastSb) {
      $headerWF = "SureBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "SureBackup Sessions with Warnings or Failures"
    }
    $arrSessWFSb = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}}, Result
    $bodySessWFSb = $arrSessWFSb | ConvertTo-HTML -Fragment
    If ($arrSessWFSb.Result -match "Failed") {
        $sessWFSbHead = $subHead01err
      } ElseIf ($arrSessWFSb.Result -match "Warning") {
        $sessWFSbHead = $subHead01war
      } ElseIf ($arrSessWFSb.Result -match "Success") {
        $sessWFSbHead = $subHead01suc
      } Else {
        $sessWFSbHead = $subHead01
      }
    $bodySessWFSb = $sessWFSbHead + $headerWF + $subHead02 + $bodySessWFSb
    }
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
  If ($successSessionsSb.count -gt 0) {
    If ($onlyLastSb) {
      $headerSucc = "Successful SureBackup Jobs"
    } Else {
      $headerSucc = "Successful SureBackup Sessions"
    }
    $bodySessSuccSb = $successSessionsSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result | ConvertTo-HTML -Fragment
    $bodySessSuccSb = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSb
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
$taskListSb = @()
$taskListSb += $sessListSb | Get-VSBTaskSession
$successTasksSb = @($taskListSb | ?{$_.Info.Result -eq "Success"})
$wfTasksSb = @($taskListSb | ?{$_.Info.Result -match "Warning|Failed"})
$runningTasksSb = @()
$runningTasksSb += $runningSessionsSb | Get-VSBTaskSession | ?{$_.Status -ne "Stopped"}

# Get SureBackup Tasks
$bodyAllTasksSb = $null
If ($showAllTasksSb) {
  If ($taskListSb.count -gt 0) {
    $arrAllTasksSb = $taskListSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Status"; Expression = {$_.Status}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Info.FinishTime}}},
      @{Name="Duration (HH:MM:SS)"; Expression = {
        If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))
        } Else {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)
        }
      }},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {
          If ($_.Info.Result -eq "notrunning") {
            "None"
          } Else {
            $_.Info.Result
          }
      }}
    $bodyAllTasksSb = $arrAllTasksSb | Sort "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllTasksSb.Result -match "Failed") {
        $allTasksSbHead = $subHead01err
      } ElseIf ($arrAllTasksSb.Result -match "Warning") {
        $allTasksSbHead = $subHead01war
      } ElseIf ($arrAllTasksSb.Result -match "Success") {
        $allTasksSbHead = $subHead01suc
      } Else {
        $allTasksSbHead = $subHead01
      }
    $bodyAllTasksSb = $allTasksSbHead + "SureBackup Tasks" + $subHead02 + $bodyAllTasksSb
  }
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
  If ($runningTasksSb.count -gt 0) {
    $bodyTasksRunningSb = $runningTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningSb = $subHead01 + "Running SureBackup Tasks" + $subHead02 + $bodyTasksRunningSb
  }
}

# Get SureBackup Tasks with Warnings or Failures
$bodyTaskWFSb = $null
If ($showTaskWFSb) {
  If ($wfTasksSb.count -gt 0) {
    $arrTaskWFSb = $wfTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyTaskWFSb = $arrTaskWFSb | Sort "Start Time" | ConvertTo-HTML -Fragment
    If ($arrTaskWFSb.Result -match "Failed") {
        $taskWFSbHead = $subHead01err
      } ElseIf ($arrTaskWFSb.Result -match "Warning") {
        $taskWFSbHead = $subHead01war
      } ElseIf ($arrTaskWFSb.Result -match "Success") {
        $taskWFSbHead = $subHead01suc
      } Else {
        $taskWFSbHead = $subHead01
      }
    $bodyTaskWFSb = $taskWFSbHead + "SureBackup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFSb
  }
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
  If ($successTasksSb.count -gt 0) {
    $bodyTaskSuccSb = $successTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}} | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTaskSuccSb = $subHead01suc + "Successful SureBackup Tasks" + $subHead02 + $bodyTaskSuccSb
  }
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
  $vbrConfigHash = @{
    "Enabled" = $configBackup.Enabled
    "Status" = $configBackup.LastState
    "Target" = $configBackup.Target
    "Schedule" = $configBackup.ScheduleOptions
    "Restore Points" = $configBackup.RestorePointsToKeep
    "Encrypted" = $configBackup.EncryptionOptions.Enabled
    "Last Result" = $configBackup.LastResult
    "Next Run" = $configBackup.NextRun
  }
  $vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
  $bodySummaryConfig = $vbrConfigObj | Select Enabled, Status, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Last Result" | ConvertTo-HTML -Fragment  
  If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) {
    $configHead = $subHead01war
  } ElseIf ($configBackup.LastResult -eq "Success") {
    $configHead = $subHead01suc
  } ElseIf ($configBackup.LastResult -eq "Failed") {
    $configHead = $subHead01err
  } Else {
    $configHead = $subHead01
  }  
  $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
  If ($proxyList -ne $null) {
    $arrProxy = $proxyList | Get-VBRProxyInfo | Select @{Name="Proxy Name"; Expression = {$_.ProxyName}},
      @{Name="Transport Mode"; Expression = {$_.tMode}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Proxy Host"; Expression = {$_.RealName}}, @{Name="Host Type"; Expression = {$_.pType}},
      Enabled, @{Name="IP Address"; Expression = {$_.IP}},
      @{Name="RT (ms)"; Expression = {$_.Response}}, Status
    $bodyProxy = $arrProxy | Sort "Proxy Host" |  ConvertTo-HTML -Fragment
    If ($arrProxy.Status -match "Dead") {
      $proxyHead = $subHead01err
    } ElseIf ($arrProxy -match "Alive") {
      $proxyHead = $subHead01suc
    } Else {
      $proxyHead = $subHead01
    }    
    $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
  }
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
  If ($repoList -ne $null) {
    $arrRepo = $repoList | Get-VBRRepoInfo | Select @{Name="Repository Name"; Expression = {$_.Target}},
      @{Name="Type"; Expression = {$_.rType}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Host"; Expression = {$_.RepoHost}}, @{Name="Path"; Expression = {$_.Storepath}},
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"} 
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}
      }
    $bodyRepo = $arrRepo | Sort "Repository Name" | ConvertTo-HTML -Fragment       
    If ($arrRepo.status -match "Critical") {
      $repoHead = $subHead01err
    } ElseIf ($arrRepo.status -match "Warning|Unknown") {
      $repoHead = $subHead01war
    } ElseIf ($arrRepo.status -match "OK") {
      $repoHead = $subHead01suc
    } Else {
      $repoHead = $subHead01
    }    
    $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
  }
}

# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
  If ($repoListSo -ne $null) {
    $arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select @{Name="Scale Out Repository Name"; Expression = {$_.SOTarget}},
      @{Name="Member Repository Name"; Expression = {$_.Target}}, @{Name="Type"; Expression = {$_.rType}},
      @{Name="Max Tasks"; Expression = {$_.MaxTasks}}, @{Name="Host"; Expression = {$_.RepoHost}},
      @{Name="Path"; Expression = {$_.Storepath}}, @{Name="Free (GB)"; Expression = {$_.StorageFree}},
      @{Name="Total (GB)"; Expression = {$_.StorageTotal}}, @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}
      }
    $bodySORepo = $arrSORepo | Sort "Scale Out Repository Name", "Member Repository Name" | ConvertTo-HTML -Fragment
    If ($arrSORepo.status -match "Critical") {
      $sorepoHead = $subHead01err
    } ElseIf ($arrSORepo.status -match "Warning|Unknown") {
      $sorepoHead = $subHead01war
    } ElseIf ($arrSORepo.status -match "OK") {
      $sorepoHead = $subHead01suc
    } Else {
      $sorepoHead = $subHead01
    }
    $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
  }
}

# Get Repository Agent User Permissions
$bodyRepoPerms = $null
If ($showRepoPerms){
  If ($repoList -ne $null -or $repoListSo -ne $null) {
    $bodyRepoPerms = Get-RepoPermissions | Select Name, "Encryption Enabled", "Permission Type", Users | Sort Name | ConvertTo-HTML -Fragment
    $bodyRepoPerms = $subHead01 + "Repository Permissions for Agent Jobs" + $subHead02 + $bodyRepoPerms
  }
}

# Get Replica Target Info
$bodyReplica = $null
If ($showReplicaTarget) {
  If ($allJobsRp -ne $null) {
    $repTargets = $allJobsRp | Get-VBRReplicaTarget | Select @{Name="Replica Target"; Expression = {$_.Target}}, Datastore,
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $replicaCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
        ElseIf ($_.FreePercentage -lt $replicaWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}
        }
      } | Sort "Replica Target"
    $bodyReplica = $repTargets | ConvertTo-HTML -Fragment
    If ($repTargets.status -match "Critical") {
      $reptarHead = $subHead01err
    } ElseIf ($repTargets.status -match "Warning|Unknown") {
      $reptarHead = $subHead01war
    } ElseIf ($repTargets.status -match "OK") {
      $reptarHead = $subHead01suc
    } Else {
      $reptarHead = $subHead01
    }    
    $bodyReplica = $reptarHead + "Replica Target Details" + $subHead02 + $bodyReplica
  }
}

# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
  $vServers = Get-VeeamWinServers
  $vServices = Get-VeeamServices $vServers
  If ($hideRunningSvc) {$vServices = $vServices | ?{$_.Status -ne "Running"}}
  If ($vServices -ne $null) {
    $vServices = $vServices | Select "Server Name", "Service Name",
      @{Name="Status"; Expression = {If ($_.Status -eq "Stopped"){"Not Running"} Else {$_.Status}}}
    $bodyServices = $vServices | Sort "Server Name", "Service Name" | ConvertTo-HTML -Fragment
    If ($vServices.status -match "Not Running") {
      $svcHead = $subHead01err
    } ElseIf ($vServices.status -notmatch "Running") {
      $svcHead = $subHead01war
    } ElseIf ($vServices.status -match "Running") {
      $svcHead = $subHead01suc
    } Else {
      $svcHead = $subHead01
    }
    $bodyServices = $svcHead + "Veeam Services (Windows)" + $subHead02 + $bodyServices        
  }
}

# Get License Info
$bodyLicense = $null
If ($showLicExp) {
  $arrLicense = Get-VeeamSupportDate $vbrServer | Select @{Name="Expiry Date"; Expression = {$_.ExpDate}},
    @{Name="Days Remaining"; Expression = {$_.DaysRemain}}, `
    @{Name="Status"; Expression = {
      If ($_.DaysRemain -lt $licenseCritical) {"Critical"}
      ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"}
      ElseIf ($_.DaysRemain -eq "Failed") {"Failed"}
      Else {"OK"}}
    }  
  $bodyLicense = $arrLicense | ConvertTo-HTML -Fragment
  If ($arrLicense.Status -eq "OK") {
    $licHead = $subHead01suc
  } ElseIf ($arrLicense.Status -eq "Warning") {
    $licHead = $subHead01war
  } Else {
    $licHead = $subHead01err
  }
  $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}

# Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb
  
If ($bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb) {
  $htmlOutput += $HTMLbreak
}
  
$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess

If ($bodyMissing + $bodySuccess + $bodyWarning) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMultiJobs

If ($bodyMultiJobs) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk

If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM

If ($bodyRestoRunVM + $bodyRestoreVM) {
  $htmlOutput += $HTMLbreak
  }

$htmlOutput += $bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp

If ($bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc

If ($bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp

If ($bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt

If ($bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp

If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb

If ($bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ","<br />")
# Remove trailing HTMLbreak
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)","$($footerObj)")
# Add color to output depending on results
#Green
$htmlOutput = $htmlOutput.Replace("<td>Running<","<td style=""color: #00b051;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<","<td style=""color: #00b051;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<","<td style=""color: #00b051;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<","<td style=""color: #00b051;"">Success<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td>Warning<","<td style=""color: #ffc000;"">Warning<")
#Red
$htmlOutput = $htmlOutput.Replace("<td>Not Running<","<td style=""color: #ff0000;"">Not Running<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<","<td style=""color: #ff0000;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<","<td style=""color: #ff0000;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<","<td style=""color: #ff0000;"">Dead<")
# Color Report Header and Tag Email Subject
If ($htmlOutput -match "#FB9895") {
  # If any errors paint report header red
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#FB9895")
  $emailSubject = "[Failed] $emailSubject"
} ElseIf ($htmlOutput -match "#ffd96c") {
  # If any warnings paint report header yellow
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#ffd96c")
  $emailSubject = "[Warning] $emailSubject"
} ElseIf ($htmlOutput -match "#00b050") {
  # If any success paint report header green
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#00b050")
  $emailSubject = "[Success] $emailSubject"
} Else {
  # Else paint gray
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#626365")
}
#endregion

#region Output
# Send Report via Email
If ($sendEmail) {
  $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
  $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
  $smtp.EnableSsl = $emailEnableSSL
  $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
  $msg.Subject = $emailSubject
  If ($emailAttach) {
    $body = "Veeam Report Attached"
    $msg.Body = $body
    $tempFile = "$env:TEMP\$($rptTitle)_$(Get-Date -format MMddyyyy_hhmmss).htm"
    $htmlOutput | Out-File $tempFile
    $attachment = new-object System.Net.Mail.Attachment $tempFile
    $msg.Attachments.Add($attachment)
  } Else {
    $body = $htmlOutput
    $msg.Body = $body
    $msg.isBodyhtml = $true
  }       
  $smtp.send($msg)
  If ($emailAttach) {
    $attachment.dispose()
    Remove-Item $tempFile
  }
}

# Save HTML Report to File
If ($saveHTML) {       
  $htmlOutput | Out-File $pathHTML
  If ($launchHTML) {
    Invoke-Item $pathHTML
  }
}
#endregion
