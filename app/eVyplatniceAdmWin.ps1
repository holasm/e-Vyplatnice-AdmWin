<#
.SYNOPSIS
    ADMWIN - pro spole�nost Reality
    Rozesl�n� mzdov�ch v�platnic na mail zam�stnanc� z programu ADMWIN
.DESCRIPTION
    Mzdov� ��t�rna provede export soubor� z ADMWIN do definovan� slo�ky (eVyplatnice-AdmWin\IN). N�sledn� se spu�� eVyplatnice, kter� zpracuje exportovan� data a roze�le na mail maily pracovn�k�.

    1. Po spu�t�n� se na�te konfigurace zam�stnanc� - mail, heslo, osobn� ��slo z DBF - ZAMEST
    2. Exportovan� soubory mus� b�t ve form�tu PDF
    3. Generuj� a zahesluj� se soubory PDF 
    5. Odes�l� se mail dle konfigurace v DBF ADMWIN
    6. V�echny podkladov� i pracovn� soubory jsou v p��pad� �sp�n�ho i chybn�ho dokon�en� smaz�ny
    7. Zobraz� se log soubor v notepadu
    8. opakovan� spou�t�n� je mo�n� dal��m exportem ze ADMWIN - nap�. oprava chyb zachycen�ch p�i zpracov�n�
    9. opakovan� spou�t�n� po n�hl�m vypnut� v pr�b�hu zpracov�n� bez nutnosti dal��ho exportu je mo�n�, z�stanou-li soubory PDF ve slo�ce IN s konfigura�n�m souborem CFG
.EXAMPLE
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -File ".\eVyplatnice.ps1" 
.EXAMPLE
    Another example of how to use this cmdlet
.INPUTS
    Vyplata_pro_os_c_XXXXX_za_RRRR_MM.pdf - form�t souboru exportovan�ch ze syst�mu ADMWIN, zm�na n�zvu znamen� probl�m!
.OUTPUTS
    logovaci soubor otevreny v notepadu
#>

# nastaveni cest k souborum
$CestaApp = "c:\eVyplatnice-AdmWin\app\"
$Cestaksouborum = "c:\eVyplatnice-AdmWin\in\" # cesta k exportovanym vyplatnicim ze SAPu
$CestaLocalIO = "c:\eVyplatnice-AdmWin\"
$KonvertorHtm2Pdf = $CestaApp + "wkhtmltopdf.exe" # konvertor HTML do PDF
# $SouborLogo = $CestaApp + "logo_ZKL.gif" # logo ZKL

# DBF s daty pro odeslani vyplatnic ADMWIN - napevno nastaveny Reality slo�ka UBM
# pozor na delku znaku v ceste do 8 OK pak chybuje
$CestaDBFdata = "W:\UCET\UBM\ZAMESTt.DBF"
$selectdbf = "select * from " + $CestaDBFdata


# ��st konfigurace mailu
$smtp = "mail.zkl.cz"
$MailBox =  "automat@zkl.cz"
$Subject = "ZKL v�platnice - "
$Body = "Tento email je generov�n automaticky a nelze na n�j odpov�dat. V p��pad� dotazu kontaktujte mzdovou ��t�rnu ZKL."
$MailUserName = "automat"
$CestaPassword = $CestaApp +"securepassword.txt" 

# counter chyb
$PocetChyb = 0

# nacteni potrebnych knihoven
[System.Reflection.Assembly]::LoadFrom($CestaApp + "itextsharp.dll") > $null
[System.Reflection.Assembly]::LoadFrom($CestaApp + "mailkit.dll") > $null

function Logging {
    Param ($sFullPath, $LogZ)
    Write-Host  [$([DateTime]::Now)]"-"$LogZ
    Add-Content -Path $sFullPath -Value  [$([DateTime]::Now)]"-"$LogZ
}


# funkce pro pouziti knihovny itextsharp
function PSUsing {
    param (
        [System.IDisposable] $inputObject = $(throw "The parameter -inputObject is required."),
        [ScriptBlock] $scriptBlock = $(throw "The parameter -scriptBlock is required.")
    )
 
    Try {
        &$scriptBlock
    }
    Finally {
        if ($null -eq $inputObject.psbase) {
            $inputObject.Dispose()
        } else {
            $inputObject.psbase.Dispose()
        }
    }
}

function IsValidEmail { 
    param([string]$Email)
    $Regex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'

   try {
        $obj = [mailaddress]$Email
        if($obj.Address -match $Regex){
            return $True
        }
        return $False
    }
    catch {
        return $False
    } 
}

$ZamestnanciArray = @()

#Vytvoreni souboru Logu a zacatek logovani
$sFullPath = Join-Path -Path $CestaLocalIO -ChildPath "$(get-date -f yyyy-MM-dd).log"
if (!(Test-Path -Path $sFullPath)) {
    New-Item -Path $sFullPath -ItemType File -Force > $null
    $LogMsg = "Vytvo�en� log souboru - " + $sFullPath
    Logging -sFullPath $sFullPath -LogZ $LogMsg
} else {
    $LogMsg = "P�id�v�n� dat do log souboru (dal�� spu�t�n�) - " + $sFullPath
    Logging -sFullPath $sFullPath -LogZ $LogMsg
}

try {
    $HashPassword = Get-Content -Path $CestaPassword  -ErrorAction SilentlyContinue
    $SecureString = ConvertTo-SecureString -String $HashPassword -ErrorAction SilentlyContinue
    $Credentials = New-Object System.Management.Automation.PSCredential "User", $SecureString
    $MailPassword = $Credentials.GetNetworkCredential().Password
}
catch {
    $LogMsg = "Probl�m s heslem k ��tu pro odesl�n� p�es SSL. (.\zmena_hesla_k_mailu.ps1) " + $smtp +"/"+ $MailBox
    Logging -sFullPath $sFullPath -LogZ $LogMsg
    Start-Process 'notepad.exe' $sFullPath
    break
}

# Do zpracovani se zahrnou pouze PDF soubory
try {
    $vyplatnice = Get-ChildItem -path $Cestaksouborum -Filter *.pdf -ErrorAction Stop -ErrorVariable $ERR| Select-Object name
}
catch {
    $LogMsg ="Nepoda�ilo se na��st soubory v�platnic:"+ $_.Exception.Message
    Logging -sFullPath $sFullPath -LogZ $LogMsg
    # Read-Host -Prompt "Zpracovani dokonceno - stisknete ENTER pro uzavreni okna"
    Start-Process 'notepad.exe' $sFullPath
    break
}
    $CelkemVyplatnic = ($vyplatnice | Measure-Object).Count

# Pokud neexistuje zadna vyplatnice k odeslani tak ukon�uji
if (!($vyplatnice)) {
    $LogMsg ="Nic ke zpracov�n�... Nejprve spus�te export v�platnic." 
    Logging -sFullPath $sFullPath -LogZ $LogMsg
    #Read-Host -Prompt "Zpracovani dokonceno - stisknete ENTER pro uzavreni okna"
    Start-Process 'notepad.exe' $sFullPath
    break
}

###########
# Faze 1. nacteni dat o zamestnanc�ch
###########
$LogMsg = "ZAH�JENO ZPRACOVAN� A ODES�L�N� V�PLATNIC eMAILEM"
Logging -sFullPath $sFullPath -LogZ $LogMsg

# Spusteni SQL Query, nastaven� v�sledku a zachycen� chyby
Try 
{
    $CSVdata = $null
    try {
                
        $ConnString = "Provider=VFPOLEDB.1;Data Source="+$CestaDBFdata+";Codepage=437;Extended Properties=dBASE V;User ID=;Password=;Collating Sequence=MACHINE;"
        $Conn = new-object System.Data.OleDb.OleDbConnection($connString)
        $conn.open()

        $cmd = new-object System.Data.OleDb.OleDbCommand($selectdbf,$Conn)
        $da = new-object System.Data.OleDb.OleDbDataAdapter($cmd)
        $DBFdata = new-object System.Data.dataTable 
        $da.fill($DBFdata)
        # $DBFdata |Select-Object OSC,JMEN,PRIJM,EMAIL,COP
    }
    catch {
        $LogMsg ="Nepoda�ilo se na��st konfiguraci ke zpracovani v�platnic - z DBF. (nechyb� ovlada�e, je dostupn� soubor s daty?)"
        Logging -sFullPath $sFullPath -LogZ $LogMsg
        #Read-Host -Prompt "Zpracovani dokonceno - stisknete ENTER pro uzavreni okna"
        Start-Process 'notepad.exe' $sFullPath
        break
    }
    
    foreach ($UserInDB in $DBFdata) {
        
        $ErrDB = 0 # nastaveni chyby v DB, pokud chybi u zamestnace heslo nebo email pravda
        #Kontrola vyplneni poli
        
        <#
            if ($UserInDB.cop -eq $null) {Write-Host 1}
            if ($UserInDB.cop -eq "") {Write-Host 2}   
            if ([string]::IsNullOrEmpty($UserInDB.cop)) {Write-Host 3}
            if ([string]::IsNullOrWhiteSpace($UserInDB.cop)) {Write-Host 4}
        #>
        if ($UserInDB.cop) {Write-Host 2 $UserInDB.prijm $UserInDB.cop}   
        if (!($UserInDB.cop)) {Write-Host 0 $UserInDB.prijm $UserInDB.cop}   

        if (-not $UserInDB.cop -OR -not $UserInDB.email) {
            # Chybi vyplnene pole heslo nebo email
            $ErrDB = 1
        }
  
        <#
        $EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
        $DidItMatch = $UserInDB.email -match $EmailRegex
        if (!($DidItMatch)) {
            # $LogMsg = "Chybne zadana mailova adresa " + $UserInDB.email
            # Logging -sFullPath $sFullPath -LogZ $LogMsg
            # Write-Host "Chybne zadana mailova adresa " $UserInDB.email        
            $ErrDB = 1
        } #>
        
        If (!(IsValidEmail($UserInDB.email))) {
            # Write-Host "Chybne zadana mailova adresa " $UserInDB.email        
            $ErrDB = 1
        }

        $ZamestnanciArrayLine = new-object PSObject
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "oscis" -Value $UserInDB.OSC
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "heslo" -Value $UserInDB.COP
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "email" -Value $UserInDB.EMAIL
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "jmeno" -Value $UserInDB.JMEN
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "prijmeni" -Value $UserInDB.PRIJM
        $celejmeno = $UserInDB.PRIJM+" "+$UserInDB.JMEN
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "celejmeno" -Value $celejmeno
        
        $ZamestnanciArrayLine | Add-Member -MemberType NoteProperty -Name "ErrDB" -Value $ErrDB #chybovy flag - neni vyplneno nejake pole
        $ZamestnanciArray += $ZamestnanciArrayLine     
        
    }
}
Catch 
{
    $LogMsg ="Chyba p�i zpracov�n� DBF: " +  $_.Exception.Message
    Logging -sFullPath $sFullPath -LogZ $LogMsg
    Start-Process 'notepad.exe' $sFullPath
    break
}

#Zjistim komu se nevygenerovala vyplatnice a zaloguji
#foreach ($soubor in $CSVdata.ID) {
#    $nazev_souboru_vyplatnice = [System.String]::Concat($soubor,'.htm')
#    $existuje_vyplatnice=$vyplatnice|where-object {$_.Name -like $nazev_souboru_vyplatnice} 
#    if (!($existuje_vyplatnice)) {
#        $jmeno_zamestnance=$CSVdata|where-object {$_.ID -like $soubor}
#        $LogMsg ="INFO: Pro zam�stnance "+$jmeno_zamestnance.prijmeni+" "+$jmeno_zamestnance.jmeno+" osobn� ��slo "+$soubor+" nebyla ze SAPu vyexportov�na v�platnice."
#        Logging -sFullPath $sFullPath -LogZ $LogMsg
#    }
#}

###########
# Faze 2. konverze, odstran�n� puvodn�ch soubor� a odesl�n� mailem
###########
$StatusCounter = 0
$client = New-Object MailKit.Net.Smtp.SmtpClient

foreach ($soubor in $vyplatnice) {
    $StatusCounter +=1
    Write-Progress -Activity "Zpracov�n� v�platnic" -Status "$StatusCounter/$CelkemVyplatnic"  -PercentComplete (($StatusCounter/$CelkemVyplatnic)*100)
    
    $VstupniSoubor = $Cestaksouborum+$soubor.Name
    $SouborHTM = $CestaLocalIO+[System.Io.Path]::GetFileNameWithoutExtension($soubor.name)+"tmp.htm"
    $SouborPDFtmp = $CestaLocalIO+[System.Io.Path]::GetFileNameWithoutExtension($soubor.name)+"tmp.pdf"

    $SplitSoubor = [System.Io.Path]::GetFileNameWithoutExtension($soubor.name) -split "_"
    $OSCsoubor = $SplitSoubor[4] # osobni cislo z nazvu souboru

    $zamestnanec = $ZamestnanciArray| Where-Object {$_.oscis -like [System.Io.Path]::GetFileNameWithoutExtension($OSCsoubor)}
    if (!($zamestnanec)) {
        Remove-Item -Path $VstupniSoubor
        # V DB Nenalezen z�znam pracovnika pro zpracov�vanou vyplatnici  
        $LogMsg ="CHYBA: K v�platnici nelze p�ipojit data pro odesl�n�. Opravte datab�zi. - "+ $soubor.name
        Logging -sFullPath $sFullPath -LogZ $LogMsg
        $PocetChyb += 1
    } elseif ($zamestnanec.ErrDB -eq 1) {
        Remove-Item -Path $VstupniSoubor
        $LogMsg ="CHYBA: Opravte v datab�zi chybn� email nebo pr�zdn� heslo - "+$zamestnanec.prijmeni.Trim()+", "+ $soubor.name
        Logging -sFullPath $sFullPath -LogZ $LogMsg
        $PocetChyb += 1

    } else {
        # Osobni cislo nalezeno v DB pokracuji v odeslani

        
        # vytazeni obdobi z nazvu souboru
        $VyplObd = $SplitSoubor[7] 
        $Rok = $SplitSoubor[6] 
 
        if (!($Rok -match '\d+$')) # pokud rok nen� ��slo, uka� chybu
        { 
            $LogMsg ="CHYBA: Z�ejm� chybn� form�t vstupn�ho souboru (rok = "+$Rok+"). - V�platnice bude i tak odesl�na - "+ $soubor.name
            Logging -sFullPath $sFullPath -LogZ $LogMsg
            $PocetChyb += 1    
        }



        # uprava podoby vyplatnice
        #$txtvyplatnice = $txtvyplatnice.replace('#E8EAD8', '#FFFFFF') # zmena barvy pozadi
        #$txtvyplatnice = $txtvyplatnice.replace('size="2" color=#0273bc', 'size="3" color=#000000') # zmena velikosti a barvy pisma
        #$txtvyplatnice = $txtvyplatnice.replace('V&nbsp;Y&nbsp;P&nbsp;L&nbsp;A&nbsp;T&nbsp;N&nbsp;Í&nbsp;&nbsp;&nbsp;P&nbsp;Ø&nbsp;S&nbsp;K&nbsp;A&nbsp;', '<strong>V&nbsp;Ý&nbsp;P&nbsp;L&nbsp;A&nbsp;T&nbsp;N&nbsp;Í&nbsp;&nbsp;&nbsp;P&nbsp;Ø&nbsp;S&nbsp;K&nbsp;A&nbsp;</strong>') # oprava textu VYPLATNI PASKA
        #$txtvyplatnice = $txtvyplatnice.replace('<font style="font-family:monospaced">', '&nbsp;&nbsp;&nbsp;<img src="'+$SouborLogo+'" width="70" height="70"><font style="font-family:monospaced">') # vlozeni loga
        #$txtvyplatnice = $txtvyplatnice.replace('==', '&#9472;&#9472') # zmena oddelovace bloku z "====" na linku
        #$txtvyplatnice = $txtvyplatnice.replace('&#9472=', '&#9472;&#9472') # dokonceni zmeny oddelovace bloku z "====" na linku

        # vytvoreni souboru HTML pro konverzi
        #New-Item -Path $SouborHTM -ItemType File | out-null
        #Add-Content $SouborHTM $txtvyplatnice

        # sestaveni prikazu pro konverzi HTM do PDF 
        #$Command = $KonvertorHtm2Pdf + ' -q --enable-local-file-access ' + $SouborHTM + ' ' + $SouborPDFtmp
        # spusteni konverze
        #Invoke-Expression -command $Command 2>&1 | out-null
        
        $SouborPDF = $CestaLocalIO+$VyplObd.Trim()+"-"+$Rok+"-"+$OSCsoubor+".pdf"    
        Copy-Item -Path $VstupniSoubor -Destination  $SouborPDFtmp

        
        # heslo zamestnance
        $password = ($zamestnanec.heslo).trim()
        # zaheslovani PDF pomoci knihovny itextsharp.dll
        $file = New-Object System.IO.FileInfo $SouborPDFtmp
        $fileWithPassword = New-Object System.IO.FileInfo $SouborPDF
        PSUsing ( $fileStreamIn = $file.OpenRead() ) { 
            PSUsing ( $fileStreamOut = New-Object System.IO.FileStream($fileWithPassword.FullName,[System.IO.FileMode]::Create,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None) ) { 
                PSUsing ( $reader = New-Object iTextSharp.text.pdf.PdfReader $fileStreamIn ) {
                    [iTextSharp.text.pdf.PdfEncryptor]::Encrypt($reader, $fileStreamOut, $true, $password, $password, [iTextSharp.text.pdf.PdfWriter]::ALLOW_PRINTING)
                }
            }
        }

        # odstraneni zdrojove vyplatnice
        Remove-Item -Path $VstupniSoubor
        # odstraneni HTML vyplatnice
        # Remove-Item -Path $SouborHTM
        # odstraneni nezaheslovane PDF vyplatnice
        Remove-Item -Path $SouborPDFtmp

        
        #       SendMailKit $zamestnanec $obdobi $SouborPDF
        
        # odeslani mailu se zaheslovanym souborem 
        # nejprve se pripravi adresace a obsah emailu
        $message = new-object Net.Mail.MailMessage
        $message.From = $MailBox
        $message.To.Add(($zamestnanec.email).trim())
        $message.Subject = $Subject + $vyplobd + "/" + $Rok
        $message.Body = $Body
        $message.isBodyhtml = $true
        $message.Attachments.add($SouborPDF)

        if (!$client.IsAuthenticated) {
            if (!$client.IsConnected) {
                $client.Connect($smtp, "587")
                if (!$client.IsConnected) {
                    $LogMsg = "Nelze se p�ipojit k emailovemu serveru " + $smtp + " - "+ $_.Exception.Message
                    Logging -sFullPath $sFullPath -LogZ $LogMsg
                    $LogMsg = "B�h ukon�en p�ed�asn�, nebylo zpracov�no " + ($CelkemVyplatnic-$StatusCounter+1) + " z celkov�ch "+ $CelkemVyplatnic + " v�platnic."
                    Logging -sFullPath $sFullPath -LogZ $LogMsg
                    Start-Process 'notepad.exe' $sFullPath
                    break
                }
            }
            $client.Authenticate($MailUserName, $MailPassword)
            if (!$client.IsAuthenticated) {
                $LogMsg = "Nelze se p�ihl�sit k ��tu " + $MailBox + " na serveru " + $smtp + " - "+ $_.Exception.Message
                Logging -sFullPath $sFullPath -LogZ $LogMsg  
                $LogMsg = "B�h ukon�en p�ed�asn�, nebylo zpracov�no " + ($CelkemVyplatnic-$StatusCounter+1) + " z celkov�ch "+ $CelkemVyplatnic + " v�platnic"
                Logging -sFullPath $sFullPath -LogZ $LogMsg
                Start-Process 'notepad.exe' $sFullPath
                break
            }
        }

        try {
            $client.Send($message)
            $LogMsg ="V�platnice pro "+$zamestnanec.oscis+" - "+$zamestnanec.prijmeni.Trim()+" byla odesl�na na email "+ $zamestnanec.email
            Logging -sFullPath $sFullPath -LogZ $LogMsg
            }
        catch {
            $global:PocetChyb += 1
            $LogMsg = "Chyba odesil�n� mailu " + $zamestnanec.prijmeni.Trim() +" - "+ $_.Exception.Message 
            Logging -sFullPath $sFullPath -LogZ $LogMsg  
        }

        $message.Dispose()

        #odstraneni zaheslovane PDF po odeslani a po chybe
        Remove-Item -Path $SouborPDF
    }
}

$client.Disconnect($smtp)



if ($PocetChyb -gt 0) {
    $LogMsg = "Po�et chyb p�i zpracov�n� " + $PocetChyb
    Logging -sFullPath $sFullPath -LogZ $LogMsg
}

$LogMsg = "KONEC ZPRACOV�N�"
Logging -sFullPath $sFullPath -LogZ $LogMsg

Start-Process 'notepad.exe' $sFullPath
