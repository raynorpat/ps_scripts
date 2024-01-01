# delete all shared driver mappings
net use * /delete /yes

# grab user's desktop path
$DesktopPath = [Environment]::GetFolderPath("Desktop")

# create WScript object for creating shortcuts
$wshshell = New-Object -ComObject WScript.Shell

# save the server SMB passwords so the drives will persist on reboot
cmd.exe /C "cmdkey /add:`"192.168.100.2`" /user:`"localhost\raynorwr`" /pass:`"vunwz76QAtbmnH`""
cmd.exe /C "cmdkey /add:`"192.168.100.3`" /user:`"localhost\raynorwr`" /pass:`"GkSSjv9PixZg2a`""

# test for media share drive
if( !( Test-Path -Path "M:\" ) ) {
    # mount the drive
    cmd.exe /C "net use M: `"\\192.168.100.2\Media`" /persistent:Yes"
    #New-PSDrive -Name M -PSProvider FileSystem -Root "\\192.168.100.2\Media" -Persist

    # test for media drive shortcut access and create if neccessary
    if( !( Test-Path -Path $DesktopPath"\Shared Media Drive.lnk" ) ) {
        # create link to drive map on desktop for user
        $lnk = $wshshell.CreateShortcut( "$DesktopPath\Shared Media Drive.lnk" )
        $lnk.TargetPath = "M:\"
        $lnk.Save() 
    }
}

# test for tech share drive
if( !( Test-Path -Path "Y:\" ) ) {
    # mount the drive
    cmd.exe /C "net use Y: `"\\192.168.100.3\Tech`" /persistent:Yes"
    #New-PSDrive -Name Y -PSProvider FileSystem -Root "\\192.168.100.3\Tech" -Persist

    # test for media drive shortcut access and create if neccessary
    if( !( Test-Path -Path $DesktopPath"\Shared Tech Drive.lnk" ) ) {
        # create link to drive map on desktop for user
        $lnk = $wshshell.CreateShortcut( "$DesktopPath\Shared Tech Drive.lnk" )
        $lnk.TargetPath = "Y:\"
        $lnk.Save() 
    }
}
