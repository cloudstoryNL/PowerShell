$string = "XXACBBB12345"
$secureString = $string | ConvertTo-SecureString -Force -Asplaintext
(New-Object System.Management.Automation.PSCredential('username', $secureString)).GetNetworkCredential().Password