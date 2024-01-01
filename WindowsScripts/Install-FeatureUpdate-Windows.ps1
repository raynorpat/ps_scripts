# Force TLS 1.2 as default communication
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Check the major version and build number of Windows that we are running
$winMajorVer = [System.Environment]::OSVersion.Version.Major
$winBuildVer = [System.Environment]::OSVersion.Version.Build

if($winBuildVer -gt 22000)
{
    # NOTE: Windows 11 builds start at 22000 for the initial RTM release of 11 21H2

    # Create a temp directory
    $dir = 'C:\_temp'
    mkdir $dir
    Write-Host "Created temp directory"

    # Spawn up a webclient object and download the Windows 10 Upgrade Assistant tool and save into our temp dir
    $webClient = New-Object System.Net.WebClient
    $url = 'https://go.microsoft.com/fwlink/?linkid=2171764'
    $file = "$($dir)\Windows11InstallationAssistant.exe"
    $webClient.DownloadFile($url,$file)
    Write-Host "Downloaded Windows 11 Installation Assistant"

    # Kick off a quiet upgrade process
    Start-Process -FilePath $file -ArgumentList "/quietinstall /skipeula /auto upgrade /NoRestartUI /copylogs $dir"
    Write-Host "Kicked off Windows 11 Installation Assistant install process"
}
elseif($winMajorVer -eq 10)
{
    # Create a temp directory
    $dir = 'C:\_temp'
    mkdir $dir
    Write-Host "Created temp directory"

    # Spawn up a webclient object and download the Windows 10 Upgrade Assistant tool and save into our temp dir
    $webClient = New-Object System.Net.WebClient
    $url = 'https://go.microsoft.com/fwlink/?LinkID=799445'
    $file = "$($dir)\Win10Upgrade.exe"
    $webClient.DownloadFile($url,$file)
    Write-Host "Downloaded Windows 10 Upgrade Assistant"
    
    # Kick off a quiet upgrade process
    Start-Process -FilePath $file -ArgumentList "/quietinstall /skipeula /auto upgrade /copylogs $dir"
    Write-Host "Kicked off Windows 10 Upgrade Assistant install process"
}
else
{
    # This is not Windows 10 or 11!
    Write-Host "This is not a Windows 10 OS"
}
