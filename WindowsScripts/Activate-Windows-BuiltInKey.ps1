$service = get-wmiObject -query 'select * from SoftwareLicensingService'
$key = $service.OA3xOriginalProductKey
$service.InstallProductKey($key)
$service.RefreshLicenseStatus()