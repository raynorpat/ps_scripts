
# Get the local c: volume's encryption properties.
$volume = Get-WmiObject win32_EncryptableVolume `
  -Namespace root\CIMv2\Security\MicrosoftVolumeEncryption `
  -Filter "DriveLetter = 'C:'"

# If the volume is not encrypted, prepare it.
if ( $volume.encryptionmethod -eq 0 -or !$volume ) {
  
  # Is there not an encryptable volume? Make C: encryptable with bdehdcfg.
  if ( -not $volume ) {
      bdehdcfg -target default -quiet
  }
  
  # Is the volume encryptable? Encrypt it.
  if ( $volume ) {
      manage-bde -on c: -s -rp
  }
}
