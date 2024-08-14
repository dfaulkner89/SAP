$main = "C:\Users\faulkda001\OneDrive - PRYSMIAN GROUP\Documents\SAP\Shortcuts"
function Execute-EncryptedScript($path) {
  trap { "Decryption failed"; break }
  $raw = Get-Content $path
  $secure = ConvertTo-SecureString $raw
  $helper = New-Object system.Management.Automation.PSCredential("test", $secure)
  $plain = $helper.GetNetworkCredential().Password
  Invoke-Expression $plain
}

Execute-EncryptedScript $main\secure.bin
