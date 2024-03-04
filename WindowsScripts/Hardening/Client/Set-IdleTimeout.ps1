# Set the time limit for inactivity in minutes
$inactivityTimeLimit = 10

# Get the current session's configuration
$systemPoliciesKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name InactivityTimeoutSecs

# Check if the current session's screen saver timeout is already set to the desired time limit
if ($systemPoliciesKey.InactivityTimeoutSecs -eq ($inactivityTimeLimit * 60)) {
    Write-Host "Screen saver timeout is already set to $inactivityTimeLimit minutes."
} else {
    # Set the screen saver timeout to the desired time limit
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name InactivityTimeoutSecs -Value ($inactivityTimeLimit * 60) -Force
    Write-Host "Idle timeout set to $inactivityTimeLimit minutes."
}

# set power scheme settings to also set monitor idle to inactivity time
powercfg.exe /setacvalueindex SCHEME_CURRENT SUB_VIDEO VIDEOIDLE ($inactivityTimeLimit * 60)
powercfg.exe /setacvalueindex SCHEME_CURRENT SUB_VIDEO VIDEOCONLOCK ($inactivityTimeLimit * 60)
powercfg.exe /setactive SCHEME_CURRENT
