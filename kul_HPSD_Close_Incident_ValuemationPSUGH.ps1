#*****************************************************
# Skript, um ein Incident automatisch zu schließen, wenn das Alert geschlossen wird
# Kulla, ab 23.7.13
# Anpassung an Scom2012 ab 13.05.14
# Version: Testumgebung
# Randbedingungen: Start über eine  Subscription,            Subscriber,                Channel
# Kriterien: Resolution State=close (255), 
# nomail in CF 4, um Jojo-Effekte zu verhindern  
# Kul, ab 11/2016 Anpassung für Valuemationnutzung
#
# 20170124 HG  mehr Text bei "kein EventID" für die Ausgabe eingefügt (\Val_event_Kul.log)
# 20170124 HG  mehr Text bei "Fehler:" für die Ausgabe eingefügt (\sd_event_Kul.log)
# 20170124 HG  mehr Text bei "TicketID:" für die Ausgabe eingefügt (\sd_event_success.log)
# 20170124 HG  {}-Beseitigung bei AlertID an den Skript-Anfang gesetzt, weil ID sonst als Scriptblock erkannt wurde
# 20170124 HG  zweites get-scomalert entfernt
# 30.10.2017 kul (IN-0151059) Zeile 38 in if (!($alert.IsMonitorAlert) -and $alert.ResolvedBy -notlike 'Auto-resolve') {$ResolvedBy='System'} geändert
#*****************************************************
param (
$AlertID="6a6a8e7e-26d5-482e-8550-31c077241b8d",
$Ursache="empty",
$TicketID="EV-0004409",
$ResolvedBy="EPA-SCOMAA"
)
$ScomStatusPath = "S:\Status\Valuemation"

    ## Kontrolle, ob die Auslösung des Vorgangs auch von der Subscription kommt
    if ($Ursache -match "Channel") {
    ## Zusätzliche Zeichen bei der Alert-ID entfernen
    $AlertIDTemp=[string]($AlertID)
    $AlertID=$AlertIDTemp.Trim("{}")}

[string]$LogMsg=$null
# "Start:  $AlertID : $Ursache : $TicketID : $Owner : $ResolvedBy" | Add-Content "$($ScomStatusPath)\Val_event_Kul.log"

## Zugriff auf Som-PS-Console
$temp=Start-OperationsManagerClientShell
$alert=Get-SCOMAlert -Id $AlertID  ## Zeile 65!!
if (!($alert.IsMonitorAlert) -and $alert.ResolvedBy -notlike 'Auto-resolve') {$ResolvedBy='System'}  ## erspart das dritte 'OR'
$LogMsg="Start um $(get-date -Format dd.MM.yy::HH:mm:ss):  $AlertID : $Ursache : $TicketID :  $ResolvedBy : $(!($alert.IsMonitorAlert))"
#if ($TicketID -notlike "empty" -and ( $ResolvedBy -like "System" -or $ResolvedBy -match "ZKS-SCOMAA" ) ) {
if ($ResolvedBy -like "System" -or $ResolvedBy -match "ZKS-SCOMAA")  {
    
     ## SCOM-Rootmanagementserver ermitteln (PU oder TU)
    if ($env:USERDOMAIN -eq "Laendle-BW") {
        $Rootmanagementserver = "lndstu01sa5062"
        #$varHPSD1="10.127.155.53"
        #$varHPSD2="10.127.155.53:30980"
        }
    else {
        $Rootmanagementserver = "tlndstu01sa5060"
        #$varHPSD1="10.127.155.52"
        #$varHPSD2="10.127.155.52:30980"
        }

    ## Bibliotheken für Json-Serialization
    [void][System.Reflection.Assembly]::LoadFile("C:\inetpub\wwwroot\SCOMWebService\SCOMValueLib.dll");
    [void][System.Reflection.Assembly]::LoadFile("C:\inetpub\wwwroot\SCOMWebService\Newtonsoft.Json.dll");
    [void][reflection.assembly]::LoadWithPartialName("System.Net")
    [void][reflection.assembly]::LoadWithPartialName("System.IO")

    
    ## Konstanten für Root-Server und für den HPService-Desk
    set-variable -option constant rootMs $Rootmanagementserver
    
    $IncidentDate1	= '{0:dd.MM.yyyy HH:mm}' -f (get-date)	
    ## wird in diesem Fall nicht unbedingt benötigt, eigene Kennzeichnung
    $ci="vorn"
    

    
    ## Logging
    $LogMsg+="; Parameter: ",$AlertID,$Ursache,$TicketID
    #"Parameter: ",$AlertID,$Ursache,$TicketID,$Owner,$ArbGrpSuchcode | Add-Content $ScomStatusPath\Val_event_Kul.log 
    ## Inhalt des Lösungsfeldes erstellen 
    $solution=("Das Alert wurde vom System automatisch geschlossen, weil die Störung behoben wurde. " + '{0:dd.MM.yyyy HH:mm}' -f (get-date)+ " " + $Ursache)
    $LogMsg+="; Solution: ", $solution
    # "Solution: ", $solution | Add-Content $ScomStatusPath\Val_event_Kul.log
## Start des Webservice  auf Valuemation für EventClose
    ## Zugriff auf das zu schließende Alert
    #$AlertID='fd987bda-83f4-44e9-bfa0-d549018cff80' ## Test!!!

#    $alert=Get-SCOMAlert -Id $AlertID	# HG 24.1.17 nur einmal oben holen und dort die Klammern entfernen,  unnötig, kann nach dem 10.11. weg

    if ($alert.CustomField3 -like 'EV*') {
        ## Parameterzuordnung
        $objTicketOutput=New-Object -TypeName SCOMValueLib.Models.USU_JSON_Output
        $objTicketOutput.client="01"
        $objTicketOutput.wfName="bwl_WS_REST_closeEvent"
        $objTicketOutput.username="SCOM"
        $objTicketOutput.password="SCOM"
        $objTicketOutput.o_params.xtriggerId="$($Alert.Id)"
        $objTicketOutput.o_params.EventID="$($alert.CustomField3)"
        #$objTicketOutput.o_params.description="$newDescr"
        #$objTicketOutput.o_params.xsname="$XSname"
        #$objTicketOutput.o_params.tckShorttext="$($PAlert.Name)"
        #$objTicketOutput.o_params.dateReported="$(get-date -date $($PAlert.TimeRaised.ToLocalTime()) -Format 'yyyy-MM-dd HH:mm')"
        #$objTicketOutput.o_params.xSource="$((Get-SCOMManagementGroup).Name)"
    
        $strJsonAuto=$objTicketOutput.SerializeToJSON()
        $strJsonAuto= $strJsonAuto.Replace('"description":"","xsname":"","tckShorttext":"","dateReported":"","xSource":"","Status":"",','')

        $(get-date -format 'yyyyMMdd-HHmmss.f'),$strJsonAuto | Add-Content $ScomStatusPath\Val_event_Kul.log

        #$webAddr = "http://tpolstu01sa5060:15723/"
        ## Aufruf des Webservice, Unterscheidung PU und TU
        if ($env:USERDOMAIN -like 'Laendle-BW')
            {$webAddr = "http://10.127.191.1/vmweb/services/api/execwf"}
        else
            {$webAddr = "http://lndw5112.zd.lnd.net/vmweb-test/services/api/execwf"}
    

        $httpWebRequest = [System.Net.HttpWebRequest]::Create($webAddr)
        $httpWebRequest.ContentType = "application/json; charset=utf-8"
    
        $httpWebRequest.Method = "POST"
        $streamwriter=New-Object -type System.IO.StreamWriter $httpWebRequest.GetRequestStream()

        $strJson | Set-Content C:\temp\vs.txt
        $streamwriter.Write($strJsonAuto)
        $streamwriter.Flush()

        $httpResponse =[System.Net.HttpWebResponse]$httpWebRequest.GetResponse()
        $StreamReader=New-Object -type System.IO.StreamReader $httpResponse.GetResponseStream()

        $result = $StreamReader.ReadToEnd();
        ## 
        $objResult=$result | ConvertFrom-Json 
        #$objResult.value.clixml | set-Content $env:TEMP\clitest.xml
        #$endErg=Import-Clixml $env:TEMP\clitest.xml
        #del $env:TEMP\clitest.xml
        # $endErg
        #if ($($objResult.returncode) -notlike '00') {return $($objResult)}
        if ($objResult.result) {
        $endErg=$objResult.result
        $LogMsg+="; $($objResult)"
    
        $LogMsg | Add-Content $ScomStatusPath\Val_event_Kul.log

        }
        else
        {   $LogMsg+=$result
            $LogMsg | Add-Content $ScomStatusPath\Val_event_Kul.log}
        $LogMsg+=$result
        $httpResponse.Close()
        $StreamReader.Close()
        #return $endErg

        ## Rule-basiertes Alert wurde  geschlossen, entsprechender Hinweis ins Customfield2
        if (!($alert.IsMonitorAlert)) {
            $alert.CustomField2="Alert im Scom von $($alert.ResolvedBy) geschlossen"
            }
        ## Anti 'Jojo'
        $alert.CustomField4='nomail'
        $temp=$alert.update("CustomFields modified by Autoclose-Mechanism")

#       ("TicketiD " + $TicketID) | Add-Content  $ScomStatusPath\sd_event_success.log 
       ($(get-date -format 'yyyyMMdd-HHmmss.f') + " " + $Alert.ID + " TicketiD " + $TicketID) | Add-Content  $ScomStatusPath\sd_event_success.log 

#       "Fehler: ",$Error | Add-Content $ScomStatusPath\sd_event_Kul.log
	$error | foreach {
	  $Errtext = $(get-date -format 'yyyyMMdd-HHmmss.f') + " " + $Alert.ID 
	  $ErrText += " Scriptname: " + $_.InvocationInfo.ScriptName
	  $ErrText += " Fehlerzeile: " + $_.InvocationInfo.ScriptLineNumber
	  $ErrText += " Fehlertext: " + $_.exception.message
	  $ErrText | Add-Content $ScomStatusPath\sd_event_Kul.log	# HG 20170124
	}

    }
    else
      {$(get-date -format 'yyyyMMdd-HHmmss.f') + ' Keine EventID in CF3 für ' + $Alert.ID | Add-Content $ScomStatusPath\Val_event_Kul.log}	# HG 20170124
#    {'Keine EventID' | Add-Content $ScomStatusPath\Val_event_Kul.log}   
}
else
{write $TicketID}