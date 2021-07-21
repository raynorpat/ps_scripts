$domain = "TestCorp.com"
$domainIP = "192.168.88.10"
$altDNS = "192.168.88.1"

$password = "Password2" | ConvertTo-SecureString -asPlainText -Force
$username = "$domain\ByteAdmin" 
$credential = New-Object System.Management.Automation.PSCredential($username,$password)

# set dns to domain controller ip
Set-DnsClientServerAddress -InterfaceAlias "Ethernet" -ServerAddresses $domainIP, $altDNS

# add computer to domain controller
Add-Computer -DomainName $domain -Credential $credential

# restart computer
#Restart-Computer