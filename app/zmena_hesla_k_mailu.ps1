$password = read-host -prompt "Zadej heslo k mailu"
write-host "$password je heslo"
$secure = ConvertTo-SecureString $password -force -asPlainText
$bytes = ConvertFrom-SecureString $secure
$bytes | out-file .\securepassword.txt
