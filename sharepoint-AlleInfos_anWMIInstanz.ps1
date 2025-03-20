<#
.Synopsis
   Abruf aller Windowsserver mit Betriebssystem im Klartext
.DESCRIPTION
   
.EXAMPLE
   
.EXAMPLE
   
#>
Set-StrictMode -Version latest
Function get-MSS20Shp_AllDataMSS
{
        Function get-MSS20ShP_ManagedServerService
                                                                                                                                                                                                                                                                {
        Param(
                [string]$siteUrl="https://firm1.berlin.de/teamraum/Managed-Server-Service/MSS-bereitstellen",
                [string]$listName = "Status Managed-Server-Service",
                [bool]$mainList=$false,
                [object]$credentials
               )
    
        $top = 100  # Anzahl der Elemente pro Seite
    
        ## API-Endpunkt für den Zugriff auf Listenelemente
        $apiUrl = "$siteUrl/_api/web/lists/getbytitle('$listName')/items?/$top=$top"
    
    

        ## Schalten auf weiter Seiten vorbereiten
        $hasNextPage = $true
        $nextPageUrl = $apiUrl
    

        ## Objectarray vorbereiten
        [object[]]$objJsons=$null
        ## Schleife zur Verarbeitung aller Seiten mit Fehlerbehandlung
        while ($hasNextPage) {
            try {
                # API-Aufruf
                $webRequest = Invoke-RestMethod -Uri $nextPageUrl -Method Get -Credential $credentials -Headers @{Accept = "application/json;odata=verbose"}
                ## Doppelten Schlüssel Id gegen IdTemp ersetzen (2 Schlüssel ID und ID)
                ## und und in ein Objekt konvertieren
                $objJson=$webRequest -Replace '"Id":','"IdTemp":' | ConvertFrom-Json
        
                ## Verarbeitung der Antwort
                $objJsons+=$objJson
        
        
                ## Überprüfen auf nächste Seite, solange es eine nächste Seite gibt,
                ## läuft die Schleife
                if ($objJson.d.__next) {
                    ## in der Variable $json.d.__next steht die uri der nächsten Seite drin
                    $nextPageUrl = $objJson.d.__next
                    #write $nextPageUrl  ## Testausgabe
                } else {
                    ## gibt es keine __next mehr, wird $hasNextPage auf false gesetzt 
                    ## und die Schleife wird verlassen
                    $hasNextPage = $false
                }
            }
            catch {
                #Write-Host "Fehler beim Abrufen der Daten: $_"
                $hasNextPage = $false
            }
        }

        $objJsonsRes=$null
        $objJsonsRes = $objJsons.d.results
        ## nur windows server herausfiltern
        if ($mainList) {
            $objjsonsWSV= $objJsonsRes | Where-Object {$_.Hostname -ne $null -and $_.Hostname -match "wsv"}
            ## Rückgabe mainlist
            return $objjsonsWSV
            }
        else {
        ## Rückgabe
        #Write-Host $objJsonsRes[0]
        return $objJsonsRes
        }
        }

        ## NTLM Authentifizierungsdaten
        #$passwdPathBPC="P:\Dokumente\Passwds"
        $passwdPathAdmTS="$($env:USERPROFILE)\Documents\Passwds"
        ## !!!Für ersten Start bei den folgenden Zeilen die Kommentarzeichen entfernen
            #if (!(Test-Path -Path $passwdPathAdmTS)) {mkdir $passwdPathAdmTS}
            #$pw=Read-Host "PW" -AsSecureString
            #$pw | ConvertFrom-SecureString | Set-Content "$($passwdPathAdmTS)\mypw.xml"
            $username=(Get-ADUser $($env:USERNAME) -properties Name,extensionAttribute5).extensionAttribute5
            $strPW=Get-Content "$($passwdPathAdmTS)\mypw.xml"
            $securePassword = ConvertTo-SecureString $strPW  -Force
            $credentials = New-Object System.Management.Automation.PSCredential($username, $securePassword)

        ## Sharepointliste(n)
        <#$listNames=@(
        "LU: Betriebssystem"
        "LU: Domäne"
        "LU: Funktion"
        "LU: Kunden"
        "LU: Portal"
        "LU: Test-Netze"
        "LU: vCPU"
        "LU: Verfahren"
        "LU: vRAM"
        "LU: vStorage"
        "LU-Vorlagen"
        )
        #>
        $listNames=@(
        "Status Managed-Server-Service"
        "LU: Betriebssystem"
        )


        $objListMSS=get-MSS20ShP_ManagedServerService -listName $listNames[0] -mainList:$true -credentials $credentials

        $objListBetriebssystem=get-MSS20ShP_ManagedServerService -listName $listNames[1]  -credentials $credentials -mainList:$false
 
        ## Ausgabe zusammenstellen
        $objAllInfos=$objListMSS | Select-Object -First 5 SAP_x002d_Auftragsnummer,Abonnementname,Hostname,Adresse,Netzadresse,Betriebssystem_ListeId,
            @{Name="OS";Expression={$el=$_;($objListBetriebssystem | Where-Object {$_.IDTemp -eq $el.Betriebssystem_ListeId}).Title }}
        $count=$objListMSS.count
        return $objAllInfos
}


function New-WMIClassInstanceFromShP
{
        
   param(  
           $AboName,
           $Verwalter,
           $Erstellungsdatum,
           $Bereitstellungsdatum,
           $Verfahren,
           $Kunde,
           $AnsprechpartnerKundeName,
           $AnsprechpartnerKundeFachbereich,
           $AnsprechpartnerKundeMail,
           $SAPAuftrag,
           $SAPPosition,
           $Funktion,
           $Tier = "1",
           $Stage,
           $CloudZelle,
           $Netz,
           $vLAN,
           $WartungsfensterTag,
           $WartungsfensterZeit,
           $BasisServices 
         )

        Write-Host $AboName
        $classname="ITDZ_MSSConfiguration"

        $test=@"
        New-CimInstance -ClassName "ITDZ_MSSConfiguration" `
				        -Property @{	Hostname                    	= $env:COMPUTERNAME;
								        Domain                      	= (Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty Domain);
								        AboName                     	= $AboName;
								        Verwalter						= "IB5";
								        Erstellungsdatum            	= (Get-CimInstance -ClassName win32_OperatingSystem).InstallDate;
								        Bereitstellungsdatum        	= [datetime]"2024-03-01";
								        Verfahren                   	= "Testverfahren 1";
								        Kunde                       	= "Testkunde 1";
								        AnsprechpartnerKundeName        = "Unser Ansprechpartner";
								        AnsprechpartnerKundeFachbereich = "IB 0815";
								        AnsprechpartnerKundeMail        = "ib0815@itdz-berlin.de";
								        SAPAuftrag          	        = $SAPAuftrag;
								        SAPPosition                 	= "1234";
								        Funktion                    	= "Application";
								        Tier                        	= "1";
								        Stage                       	= "Produktion";
								        CloudZelle                  	= "CC01CCW01";
								        Netz                        	= "127.0.0.0/27";
								        vLAN                        	= "1005";
								        WartungsfensterTag          	= "3.Mittwoch";
								        WartungsfensterZeit         	= "20:00"
								        BasisServices					= @("000-MSS-BS-AV-Trellix","000-MSS-BS-Monitoring-Checkmk")
				        }
"@
        Get-CimInstance -ClassName $classname
}


$allDataFromShp=get-MSS20Shp_AllDataMSS
$allDataFromShp[0]
## Auslesen der csv als objShPCsv

$allDataFromShp | ForEach-Object {
<#New-WMIClassInstanceFromShP -AboName $($_.Abonnementname) `
-Verwalter `
-Erstellungsdatum (Get-CimInstance -ClassName win32_OperatingSystem).InstallDate `
-Bereitstellungsdatum `
-Verfahren `
-Kunde `
-AnsprechpartnerKundeName `
-AnsprechpartnerKundeFachbereich `
-AnsprechpartnerKundeMail `
-SAPAuftrag $($_.SAP_x002d_Auftragsnummer) `
-SAPPosition `
-Funktion `
-Stage `
-CloudZelle `
-Netz `
-vLAN `
-Tier=1 `
-WartungsfensterTag `
-WartungsfensterZeit `
-BasisServices#>
New-WMIClassInstanceFromShP -AboName $($_.Abonnementname) `
-SAPAuftrag $($_.SAP_x002d_Auftragsnummer) 
}
