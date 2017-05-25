# set variables from commandline
param (
    [switch]$IsServerAD,
	[switch]$wantLocal,
    [switch]$wantHIPAA,
    [switch]$wantPCI,

    [string]$ADUserCred = "DomainName\DomainAdmin",
    [string]$ADUserPswd = "DomainAdminPwd",

    [string]$LocalUser = "Test",
    [string]$LocalPswd = "Password1",

    [string]$NDConnectorID = "ConectorIDHere!"
)

# some helper functions to help make the ip range
# from: http://blog.tyang.org/2012/02/09/powershell-script-calculate-first-and-last-ip-of-a-subnet/
function ConvertTo-Binary ($strDecimal)
{
	$strBinary = [Convert]::ToString($strDecimal, 2)
	if ($strBinary.length -lt 8)
	{
		while ($strBinary.length -lt 8)
		{
			$strBinary = "0"+$strBinary
		}
	}
	Return $strBinary
}
function Convert-IP-To-Binary ($strIP)
{
	$strBinaryIP = $null

	$arrSections = @()
	$arrSections += $strIP.split(".")
	foreach ($section in $arrSections)
	{
		if ($strBinaryIP -ne $null)
		{
			$strBinaryIP = $strBinaryIP+"."
		}
		$strBinaryIP = $strBinaryIP+(ConvertTo-Binary $section)
			
	}

	Return $strBinaryIP
}
Function Convert-SubnetMask-To-Binary ($strSubnetMask)
{
	$strBinarySubnetMask = $null

	$arrSections = @()
	$arrSections += $strSubnetMask.split(".")
	foreach ($section in $arrSections)
	{
		if ($strBinarySubnetMask -ne $null)
		{
			$strBinarySubnetMask = $strBinarySubnetMask+"."
		}
		$strBinarySubnetMask = $strBinarySubnetMask+(ConvertTo-Binary $section)
			
	}

	Return $strBinarySubnetMask
}
Function Convert-BinaryIPAddress ($BinaryIP)
{
	$FirstSection = [Convert]::ToInt64(($BinaryIP.substring(0, 8)),2)
	$SecondSection = [Convert]::ToInt64(($BinaryIP.substring(8,8)),2)
	$ThirdSection = [Convert]::ToInt64(($BinaryIP.substring(16,8)),2)
	$FourthSection = [Convert]::ToInt64(($BinaryIP.substring(24,8)),2)
	$strIP = "$FirstSection`.$SecondSection`.$ThirdSection`.$FourthSection"
	Return $strIP
}

Write-Output "`n"
Write-Output "`n"
Write-Output "`n"

# generate ip range based on computer ip
$configs=gwmi win32_networkadapterconfiguration | where {$_.ipaddress -ne $null -and $_.defaultipgateway -eq $null}
if ($configs -ne $null)
{
    # we want the first good adapter's ip, that will do
	$ipaddr = $configs[0].IPAddress[0]
	$maskaddr = $configs[0].IPSubnet[0]
}
else
{
	# put a default here...
	$ipaddr = "192.168.1.5"
	$maskaddr = "255.255.255.0"
}

$BinarySubnetMask = (Convert-SubnetMask-To-Binary $maskaddr).replace(".", "")
$BinaryNetworkAddressSection = $BinarySubnetMask.replace("1", "")
$BinaryNetworkAddressLength = $BinaryNetworkAddressSection.length
$CIDR = 32 - $BinaryNetworkAddressLength
$iAddressWidth = [System.Math]::Pow(2, $BinaryNetworkLength)
$iAddressPool = $iAddressWidth -2
$BinaryIP = (Convert-IP-To-Binary $ipaddr).Replace(".", "")
$BinaryIPNetworkSection = $BinaryIP.substring(0, $CIDR)
$BinaryIPAddressSection = $BinaryIP.substring($CIDR, $BinaryNetworkAddressLength)
	
# starting IP
$FirstAddress = $BinaryNetworkAddressSection -replace "0$", "1"
$BinaryFirstAddress = $BinaryIPNetworkSection + $FirstAddress
$strFirstIP = Convert-BinaryIPAddress $BinaryFirstAddress
	
# ending IP
$LastAddress = ($BinaryNetworkAddressSection -replace "0", "1") -replace "1$", "0"
$BinaryLastAddress = $BinaryIPNetworkSection + $LastAddress
$strLastIP = Convert-BinaryIPAddress $BinaryLastAddress

# put together the ip range
$ipRanges = ($strFirstIP) + "-" + ($strLastIP)

# output parameters so we know whats going on...
Write-Output "`n"
Write-Output "Parameters:`r"
if($PSBoundParameters.ContainsKey('wantPCI')) {
    Write-Output "  Will Perform PCI security scan.`r"
}
if($PSBoundParameters.ContainsKey('wantHIPAA')) {
    Write-Output "  Will Perform HIPAA security scan.`r"
}
if($PSBoundParameters.ContainsKey('wantLocal')) {
    Write-Output "  Will Perform Local Only scan.`r"
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
mkdir -force C:\ndc\ndresults
mkdir -force C:\ndc\secresults
mkdir -force C:\ndc\hipaaresults
mkdir -force C:\ndc\pciresults

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
    "-common

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
	-ndttimeout
	1
    -logfile
    ndfRun.log

    -testports
    -testurls
    -wifi
    -policies
    -screensaver
    -usb
    -nozip
	-sdfdir
	C:\ndc\secresults	
	-sdfbase
	$env:computername-SDF
	
    -scantype
    ndc,ldc,sdc,sdcnet
	-pcitype
	net	
    " | out-file -filepath "C:\ndc\run.ndp"
}

# wait for filesystem to catch up
Start-Sleep -s 10

Write-Output "`n`nStarting scan... `n"

$start_time = Get-Date
Write-Output "Start Time: $((Get-Date).ToString('hh:mm:ss'))`n`n"

# see if we are doing a local only scan
if($PSBoundParameters.ContainsKey('wantLocal')) {
	# run network detective data collector
	Write-Output "Network Detective local only scan... `n"
	.\nddc.exe -local -silent -outdir "C:\ndc\ndresults"
	
	# send results to network detective collector
	Start-Sleep -s 2
	.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\ndresults" -zipname $env:computername-ND
	
	# run security data collector
	Write-Output "Security data collector scan... `n"
	.\sddc.exe -common -sdfdir "C:\ndc\secresults"
	
	# send results to network detective collector
	Start-Sleep -s 2
	.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\secresults" -zipname $env:computername-SDF
	
	# run pci data collector
	if($PSBoundParameters.ContainsKey('wantPCI')) {
		Write-Output "PCI compliance Detective scan... `n"
		.\pcidc.exe -outdir "C:\ndc\pciresults"
	
		# send results to network detective collector
		Start-Sleep -s 2
		.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\pciresults" -zipname $env:computername-PCI
	}
	
	# run hipaa data collector
	if($PSBoundParameters.ContainsKey('wantHIPAA')) {
		Write-Output "HIPAA compliance Detective scan... `n"
		.\hipaadc.exe -hipaadeep -outdir "C:\ndc\hipaaresults"
		
		# send results to network detective collector
		Start-Sleep -s 2
		.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\hipaaresults" -zipname $env:computername-HIPAA	
	}
} else {
	# run network detective data collector
	Write-Output "Network Detective Active Directory scan... `n"
	.\nddc.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\ndresults"
	
	# send results to network detective collector
	Start-Sleep -s 2
	.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\ndresults" -zipname $env:computername-ND
	
	# run pci data collector
	if($PSBoundParameters.ContainsKey('wantPCI')) {
		Write-Output "PCI compliance Detective scan... `n"
		.\pcicmdline.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\pciresults"
		
		# send results to network detective collector
		Start-Sleep -s 2
		.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\pciresults" -zipname $env:computername-PCI
	}
	
	# run hipaa data collector
	if($PSBoundParameters.ContainsKey('wantHIPAA')) {
		Write-Output "HIPAA compliance Detective scan... `n"
		.\hipaacmdline.exe -file "C:\ndc\run.ndp" -outdir "C:\ndc\hipaaresults"
		
		# send results to network detective collector
		Start-Sleep -s 2
		.\ndconnector.exe -ID $NDConnectorID -d "C:\ndc\hipaaresults" -zipname $env:computername-HIPAA	
	}
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
