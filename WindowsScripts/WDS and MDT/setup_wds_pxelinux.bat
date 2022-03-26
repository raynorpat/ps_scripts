wdsutil /set-server /bootprogram:boot\x86\pxelinux.com /architecture:x86
wdsutil /set-server /bootprogram:boot\x64\pxelinux.com /architecture:x64
wdsutil /set-server /N12bootprogram:boot\x86\pxelinux.com /architecture:x86
wdsutil /set-server /N12bootprogram:boot\x64\pxelinux.com /architecture:x64
wdsutil /set-server /bootprogram:boot\x64\syslinux.efi  /architecture:x64UEFI
wdsutil /set-server /N12bootprogram:boot\x64\syslinux.efi  /architecture:x64UEFI