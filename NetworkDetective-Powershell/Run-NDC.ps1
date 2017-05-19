# set variables from commandline
param (
    [switch]$IsServerAD,
    [switch]$wantHIPAA,
    [switch]$wantPCI,

    [string]$ADUserCred = "@DomainName@\@DomainAdmin@",
    [string]$ADUserPswd = "@DomainAdminPwd@",

    [string]$LocalUser = "Test",
    [string]$LocalPswd = "Password1",

    [string]$NDConnectorID = "ConectorIDHere!"
)

Write-Output "`n"
Write-Output "`n"
Write-Output "`n"

# generate ip range based on computer ip
$iptest = ((ipconfig | findstr [0-9].\.)[0]).Split()[-1] -replace ".{3}$"
$ipRanges = $iptest + "0" + "-" + $iptest + "255"

# output parameters so we know whats going on...
Write-Output "`n"
Write-Output "Parameters:`r"
if($PSBoundParameters.ContainsKey('wantPCI')) {
    Write-Output "  Will Perform PCI security scan.`r"
}
if($PSBoundParameters.ContainsKey('wantHIPAA')) {
    Write-Output "  Will Perform HIPAA security scan.`r"
}
if($PSBoundParameters.ContainsKey('IsServerAD')) {
    Write-Output "  Running in Active Directory environment.`r"
    Write-Output "  ADUserCred = " $ADUserCred " `r"
    Write-Output "  ADUserPswd = " $ADUserPswd " `r"
} else {
    Write-Output "  Running in local user environment.`r"
    Write-Output "  LocalUser = " $LocalUser " `r"
    Write-Output "  LocalPswd = " $LocalPswd " `r"
}
Write-Output "  ipRanges = " $ipRanges " `r"
Write-Output "  NDConnectorID = " $NDConnectorID " `r"
Write-Output "`n"
Write-Output "`n"

# make temporary folders
Write-Output "making temporary folders... `n"
mkdir -force C:\ndc
mkdir -force C:\ndc\results
if($PSBoundParameters.ContainsKey('wantPCI')) {
mkdir -force C:\ndc\pci_results
}
if($PSBoundParameters.ContainsKey('wantHIPAA')) {
mkdir -force C:\ndc\hipaa_results
}

# grab the network detective toolkit and save into temp folder
Write-Output "downloading network detective toolkit... `n"
$url = "https://s3.amazonaws.com/networkdetective/download/NetworkDetectiveDataCollectorNoRun.exe"
$output = "C:\ndc\NDDCNoRun.exe"

$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $output)

# change directory to temp folder
cd C:\ndc

# extract the toolkit
Write-Output "extracting network detective toolkit... `n"
.\NDDCNoRun.exe /auto

# build the command line depending on run mode and save it to file
if($PSBoundParameters.ContainsKey('IsServerAD')) {
    "-eventlogs
    -sql
    -internet
    -net
    -dhcp

    -ad
    -credsuser
    $ADUserCred
    -credsepwd
    $ADUserPswd

    -ipranges
    $ipRanges

    -threads
    20
    -snmp
    public
    -snmptimeout
    10
    -logfile
    ndfRun.log

    -testports
    -testurls
    -wifi
    -policies
    -screensaver
    -usb
    -nozip

    -scantype
    ndc,ldc,sdc,sdcnet
    -skipspeedchecks
    " | out-file -filepath "C:\ndc\run.ndp"
} else {
    "-eventlogs
    -sql
    -internet
    -net
    -dhcp

    -credsuser
    $LocalUser
    -credspwd
    $LocalPswd

    -ipranges
    $ipRanges

    -threads
    20
    -snmp
    public
    -snmptimeout
    10
    -logfile
    ndfRun.log

    -testports
    -testurls
    -wifi
    -policies
    -screensaver
    -usb
    -nozip

    -scantype
    ndc,ldc,sdc,sdcnet
    -skipspeedchecks
    " | out-file -filepath "C:\ndc\run.ndp"
}

# wait for filesystem to catch up
Start-Sleep -s 10

Write-Output "`n`nStarting scan... `n"

$start_time = Get-Date
Write-Output "Start Time: $((Get-Date).ToString('hh:mm:ss'))`n`n"

# run network detective data collector
Write-Output "Network Detective scan... `n"
.\nddc.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\results"
Start-Sleep -s 2
.\ndconnector.exe -ID $NDConnectorID -d C:\ndc\results -zipname $env:computername-NDC

# run pci detective data collector
if($PSBoundParameters.ContainsKey('wantPCI')) {
    Write-Output "PCI compliance Detective scan... `n"
    .\pcidc.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\pci_results"
    Start-Sleep -s 2
    .\ndconnector.exe -ID $NDConnectorID -d C:\ndc\pci_results -zipname $env:computername-PCI
}

# run hipaa detective data collector
if($PSBoundParameters.ContainsKey('wantHIPAA')) {
    Write-Output "HIPAA compliance Detective scan... `n"
    .\hipaadc.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\hipaa_results"
    Start-Sleep -s 2
    .\ndconnector.exe -ID $NDConnectorID -d C:\ndc\hipaa_results -zipname $env:computername-HIPAA
}

# go back to C:
cd C:\
Start-Sleep -s 5

# remove temporary folders
Write-Output "removing temporary folders... `n"
Remove-Item -Path C:\ndc -Recurse

# and we are complete!
$end_time = Get-Date
Write-Output "Time taken to finish data collection: $((Get-Date).Subtract($start_time).Seconds) second(s)"
Write-Output "End Time: $((Get-Date).ToString('hh:mm:ss'))`n`n"
