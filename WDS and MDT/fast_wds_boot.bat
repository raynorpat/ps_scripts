net stop WDSSERVER
bcdedit /enum all /store C:\RemoteInstall\Boot\x64\default.bcd
bcdedit /store C:\RemoteInstall\Boot\x64\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpwindowsize 32
bcdedit /store C:\RemoteInstall\Boot\x64\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpblocksize 16384
bcdedit /store C:\RemoteInstall\Boot\x64\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpvarwindow Yes
bcdedit /enum all /store C:\RemoteInstall\Boot\x64uefi\default.bcd
bcdedit /store C:\RemoteInstall\Boot\x64uefi\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpwindowsize 32
bcdedit /store C:\RemoteInstall\Boot\x64uefi\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpblocksize 16384
bcdedit /store C:\RemoteInstall\Boot\x64uefi\default.bcd /set {68d9e51c-a129-4ee1-9725-2ab00a957daf} ramdisktftpvarwindow Yes
net start WDSSERVER
sc control wdsserver 129
pause