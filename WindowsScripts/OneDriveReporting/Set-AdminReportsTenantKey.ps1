# Set's the SyncAdminReports key to the Microsoft Azure tenant association key
# you'll need to grab the key from https://config.office.com go to Settings and copy it out
param(
 [String]$TenantAssociationKey=''
)

$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
$registryKey = "SyncAdminReports"

# test for the full registry path, first
IF(!(Test-Path $registryPath))
{
    # create path if it doesn't exist
    New-Item -Path $registryPath -Force | Out-Null
}

# create SyncAdminReports key in the registry path
New-ItemProperty -Path $registryPath -Name $registryKey -Value $TenantAssociationKey -PropertyType String -Force | Out-Null
