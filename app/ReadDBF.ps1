# ODBC pozor na 64 a 32 bit verzi nutne konfigurovat a instalovat pro správnou verzi
# Na PC můžou být instalovány být instalován obě verze ODBC (AccessDatabaseEngine.exe a AccessDatabaseEngine_X64.exe)
# 64 veze se instaluje z prikazoveho radku admina s parametrem /silence
# V ovladacich panelech se nastavuje systemove DSN opet ve spravne bitove verzi (Zdroje dat ODBC 64 nebo 32bit)
# https://www.microsoft.com/en-us/download/details.aspx?id=13255https://www.microsoft.com/en-us/download/details.aspx?id=13255
#

<#--

#get-odbcdsn|Format-Table
#$CestaDBFadr = "W:\UCET\UBM\"
#$CestaDBFdata = "W:\UCET\UBM\ZAMEST.DBF"

$CestaDBFadr = "W:\UCET\UBM\"
$CestaDBFdata = "W:\UCET\UBM\ZAMESTt.DBF"

#priklad cteni dat z tabulky DBF
$conn = new-object System.Data.Odbc.OdbcConnection
$conn.connectionstring = "DSN=dbftreber; "+$CestaDBFadr+"; DriverID=277"
$conn.open()

# $cmd = new-object System.Data.Odbc.OdbcCommand("select * from c:\eVyplatnice-AdmWin\tmp\ZAMEST.DBF",$Conn)
$selectdbf = "select * from " + $CestaDBFdata
$cmd = new-object System.Data.Odbc.OdbcCommand($selectdbf,$Conn)
$DBFdata = new-object System.Data.ODBC.OdbcDataAdapter($cmd)
$dt = new-object System.Data.dataTable 
$DBFdata.fill($dt)|Out-Null
#$dt|select OSCZ,CZDRPOJ, ZALOH, ZMALE

#Write-Host "---------------"

#$dt|Where-Object {$_.OSCZ -like 30006}|Format-Table -AutoSize$dt|Where-Object {$_.OSC -like 1248}| Select-Object OSC,JMEN,PRIJM,EMAIL,COP

--#>



#$CestaDBFadr = "W:\UCET\UBM\"
#$CestaDBFdata = "W:\UCET\UBM\ZAMESTt.DBF"

$CestaDBFdata = "W:\UCET\UBM\ZAMESTt.DBF"
$selectdbf = "select * from " + $CestaDBFdata

$ConnString = "Provider=VFPOLEDB.1;Data Source="+$CestaDBFdata+";Codepage=437;Extended Properties=dBASE V;User ID=;Password=;Collating Sequence=MACHINE;"
$Conn = new-object System.Data.OleDb.OleDbConnection($connString)
$conn.open()

$cmd = new-object System.Data.OleDb.OleDbCommand($selectdbf,$Conn)
$da = new-object System.Data.OleDb.OleDbDataAdapter($cmd)
$dt = new-object System.Data.dataTable 
$da.fill($dt)
$dt|Where-Object {$_.OSC -like "93949" }| Select-Object OSC,JMEN,PRIJM,EMAIL,COP 
#$dt |Select-Object OSC,JMEN,PRIJM,EMAIL,COP|ft
#-------



$UserInDB=$dt|Where-Object {$_.OSC -like "93949" }| Select-Object OSC,JMEN,PRIJM,EMAIL,COP

$ble1=$UserInDB.cop.Split(' ').Where({$_.Trim() -ne ''})
$ble2=$UserInDB.email.Split(' ').Where({$_.Trim() -ne ''})

if ($ble1 -eq $null) {Write-Host 1}
if ($ble1 -eq "") {Write-Host 2}   
if ([string]::IsNullOrEmpty($ble1)) {Write-Host 3}
if ([string]::IsNullOrWhiteSpace($ble1)) {Write-Host 4}


if ($ble2 -eq $null) {Write-Host 1}
if ($ble2 -eq "") {Write-Host 2}   
if ([string]::IsNullOrEmpty($ble2)) {Write-Host 3}
if ([string]::IsNullOrWhiteSpace($ble2)) {Write-Host 4}


Write-Output "--------------"

if ($ble2){
    write-output Chyba 
} else {
    write-output OK?
}



if ($UserInDB.cop.Split(' ').Where({$_.Trim() -ne ''}) -eq '' -OR $UserInDB.email.Split(' ').Where({$_.Trim() -ne ''}) -eq ''){
    write-output Chyba 
} else {
    write-output OK?
}
  

  