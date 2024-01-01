Enable-NetFireWallRule -DisplayName "Windows Management Instrumentation (DCOM-In)"
Enable-NetFireWallRule -DisplayGroup "Remote Event Log Management"
Enable-NetFireWallRule -DisplayGroup "Remote Service Management"
Enable-NetFireWallRule -DisplayGroup "Remote Volume Management"
Enable-NetFireWallRule -DisplayGroup "Windows Firewall Remote Management"
Enable-NetFireWallRule -DisplayGroup "Remote Scheduled Tasks Management"
Enable-PSRemoting
