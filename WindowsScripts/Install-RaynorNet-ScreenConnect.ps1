# The ScreenConnect full Windows service name
$serviceName = 'ScreenConnect Client (def9adccddfe2fbf)'

# URL for ScreenConnect msi download
$url = "https://connect.raynornet.xyz/Bin/ConnectWiseControl.ClientSetup.msi?e=Access&y=Guest"

# URL for the screenconnect instance
$scdomain = "connect.raynornet.xyz"

# Check if ScreenConnect Service exists, restart if needed
If (Get-Service $serviceName -ErrorAction SilentlyContinue)
{
    If ((Get-Service $serviceName).Status -eq 'Running')
    {
        # do nothing
    }
    Else
    {
        Write-Host "$serviceName found, but it is not running for some reason."
        write-host "starting $servicename"
        start-service $serviceName
    }
}
Else
{
    # No ScreenConnect service detected, install
    Write-Host "$serviceName not found - need to install"
    [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
    (new-object System.Net.WebClient).DownloadFile($url,'C:\windows\temp\sc.msi')
    msiexec.exe /i c:\windows\temp\sc.msi /quiet /norestart
}

# parse out ScreenConnect client GUID
$Keys = Get-ChildItem HKLM:\System\ControlSet001\Services
$Guid = "Null";
$Items = $Keys | Foreach-Object {Get-ItemProperty $_.PsPath }
 
ForEach ($Item in $Items)
{
    if ($item.PSChildName -like "*ScreenConnect Client*")
    {
        $SubKeyName = $Item.PSChildName
        $Guid = (Get-ItemProperty "HKLM:\SYSTEM\ControlSet001\Services\$SubKeyName").ImagePath
    }
}

$GuidParser1 = $Guid -split "&s="
$GuidParser2 = $GuidParser1[1] -split "&k="
$Guid = $GuidParser2[0]
$ScreenConnectUrl = "https://$scdomain/Host#Access/All%20Machines//$Guid/Join"

# Write out ScreenConnect URL
Write-Host ScreenConnect URL Is: $ScreenConnectUrl