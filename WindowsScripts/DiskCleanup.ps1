# turn off Windows update service
Get-Service -Name wuauserv | Stop-Service -Force -verbose -ErrorAction SilentlyContinue

# clear out Windows update download folder
Get-ChildItem "C:\Windows\SoftwareDistribution\*" -Recurse -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue

# turn back on Windows update service
Get-Service -Name wuauserv | Start-Service -Verbose

# clean out Windows Temporary files
Get-ChildItem "C:\Windows\Temp\*" -Recurse -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue

# clean out local User Profile Temporary files, Internet History, and Internet Cookies
Get-ChildItem "C:\Users\*\AppData\Local\Temp" -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue
Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files" -Recurse -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue
Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Windows\History\*" -Recurse -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue
Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Windows\INetCookies\*" -Recurse -Force -verbose -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue

# run disk cleanup 
cleanmgr /sagerun:1 | out-Null
