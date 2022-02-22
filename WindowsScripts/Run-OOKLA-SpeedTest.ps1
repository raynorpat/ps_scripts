#https://www.speedtest.net/apps/cli
cls

$DownloadURL = "https://install.speedtest.net/app/cli/ookla-speedtest-1.0.0-win64.zip"

# Location to save on the computer. Path must exist
$DOwnloadPath = "c:\SpeedTest.Zip"
$ExtractToPath = "c:\SpeedTest"
$SpeedTestEXEPath = "C:\SpeedTest\speedtest.exe"
# Log File Path
$LogPath = 'c:\SpeedTestLog.txt'

# Start Logging to a Text File
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path $LogPath -Append:$false
# Check for and delete existing log files

function RunTest()
{
    $test = & $SpeedTestEXEPath --accept-license
    $test
}

# Check if file exists
if (Test-Path $SpeedTestEXEPath -PathType leaf)
{
    Write-Host "SpeedTest EXE Exists, starting test" -ForegroundColor Green
    RunTest
}
else
{
    Write-Host "SpeedTest EXE Doesn't Exist, starting file download"

    wget $DownloadURL -outfile $DOwnloadPath

    # unzip the Ookla CLI client
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function Unzip
    {
        param([string]$zipfile, [string]$outpath)

        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }

    Unzip $DOwnloadPath $ExtractToPath
    RunTest
}

# Stop logging
Stop-Transcript
exit 0