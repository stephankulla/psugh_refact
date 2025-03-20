<#
.Synopsis
   Neue Server ins WAC
.DESCRIPTION
   Lange Beschreibung
.EXAMPLE
   Beispiel für die Verwendung dieses Cmdlets
.EXAMPLE
   Ein weiteres Beispiel für die Verwendung dieses Cmdlets
#>


## Server aus der SA auslesen

    #region dll und connection
    $HPSAObject=Connect-HPESASA
    $user=$HPSAObject.User
    $mySaCoreIP=$HPSAObject.CoreServer

    ## Connect to API
    $wsdlPrefix = "https://" + $mySaCoreIP + ":443/osapi/com/opsware/server/ServerService"
    $serverService = new-object OpswareWebServices.ServerService($user, $wsdlPrefix)

## Server einsammeln und filtern

    $filter = new-object OpswareWebServices.Filter
    $filter.objectType = "device"
    # Locate the Server reference using the WS-API, display the number or servers located.
    $serverRef_ = $serverService.findServerRefs( $filter )
    
    ## Status enthält das Erstellungsdatum, daraus eine Liste filtern
    $Serversa=$serverService.getServerVOs($serverRef_) | 
        Where-Object {$_.osFlavor -match 'win' -and $_.createdDate -ge $((get-date).AddDays(-10))} | Select-Object @{Name='Name';expression={$_.hostname}},createdDate
    ## ohne das Erstellungsdatum zu berücksichtigen, daraus eine Liste filtern
    #$Serversa=$serverService.getServerVOs($serverRef_) | 
    ##    Where-Object {$_.osFlavor -match 'win' -and $_.createdDate -ge $((get-date).AddDays(-10))} | Select-Object @{Name='Name';expression={$_.hostname}},createdDate

## Mit WACHosts vergleichen, neue werden der CSV-Datei hinzugefügt
$WACGateway='srv03wsv40123'
Invoke-Command -ComputerName $WACGateway -ScriptBlock { 
    Import-Module 'C:\Program Files\Windows Admin Center\PowerShell\Modules\ConnectionTools'
    Export-Connection "https://$($using:WACGateway)" -fileName 'D:\AdminCenterDaten\WAC_ServerListVorhanden.csv'}
$WAcList=Import-Csv -Path "\\$($WACGateway)\D$\AdminCenterDaten\WAC_ServerListVorhanden.csv" -Delimiter ','

## Liste der hinzugekommenen Server ermitteln
$ServerNew=Compare-Object -ReferenceObject $WAcList -DifferenceObject $Serversa -Property name | Where-Object {$_.SideIndicator -like '=>'} 
## ITS und ITDZ heraufiltern
$ServerNew=$ServerNew | Where-Object {$_.name -match '.its.' -or $_.name -match '.itdz.'}
        ## Alternative: nur die neuen Server werden für den Import genutzt und danach werden sie der zentralen CSV-Datei hinzugefügt

## WAC fähiges Objekt bauen
    $Tag='CloudServer'
    
    ## Array of objects
    [object[]]$myObjects=$null
    ## Objecte erstellen für spätere CSV    
    $ServerNew | ForEach-Object {
    [object]$myObject=$null
    $myObject= '' | Select-Object name,type,tags,groupId
    $myObject.name=$_.name;
    $myObject.type="msft.sme.connection-type.server";
    $myObject.tags=$Tag
    $myObjects+=$myObject
    }

## Export in CSV für den Import ins WAC
$myObjects | Export-Csv -Path "\\srvwsv35178\D$\AdminCenterDaten\WAC_NewServerList.csv" -NoTypeInformation
# $myObjects | Export-Csv -Path "\\cc03wsv3578\D$\AdminCenterDaten\WAC_AdmTSundCo.csv" -NoTypeInformation -Append


#$Port='6516'
if (!(Test-Path "\\$($WACGateway)\D$\AdminCenterDaten")) {mkdir "\\$($WACGateway)\D$\AdminCenterDaten"}

## Liste zum neuen Gateway kopiern
Copy-Item -Path "\\srv03wsv35178\D$\AdminCenterDaten\WAC_NewServerList.csv" -Destination "\\$($WACGateway)\D$\AdminCenterDaten\WAC_NewServerList.csv"
## Session auf dem Server mit WAC eröffnen bzw. reomotzugriff
## Import der CSV Datei
## Dabei SSO in Abhängigkeit von der Domäne setzen

Invoke-Command -ComputerName $WACGateway -ScriptBlock { 
    Import-Module 'C:\Program Files\Windows Admin Center\PowerShell\Modules\ConnectionTools'
    Import-Connection "https://$($using:WACGateway):$($Port)" -fileName 'D:\AdminCenterDaten\WAC_NewServerList.csv'}
    $Serverlist=Import-Csv -Path "\\cc03wsv3578\D$\AdminCenterDaten\WAC_NewServerList.csv" -Delimiter ','
    #$TargetServer= ($Serverlist.name) | ForEach-Object {$_.split('.')[0]} 

    ## SSO
    $gateway = $WACGateway # Machine where Windows Admin Center is installed
    $Serverlist | ForEach-Object {
        if ($_.name -match '.itdz.')  
        {$aktDC='ITDZDC001.dom1.verwalt.de'
        $node = ($_.name).split('.')[0] # Machine that you want to manage
        $gatewayObject = Get-ADComputer -Identity $gateway -Server $aktDC 
        $nodeObject = Get-ADComputer -Identity $node -Server $aktDC 
        Set-ADComputer  -Identity $nodeObject -PrincipalsAllowedToDelegateToAccount $gatewayObject -PassThru  }
        if ($_.name -match '.its.')  
        {$aktDC='itsdc101.dom2.verwalt.de'
        $node = ($_.name).split('.')[0] # Machine that you want to manage
        $gatewayObject = Get-ADComputer -Identity $gateway 
        $nodeObject = Get-ADComputer -Identity $node -Server $aktDC 
        Set-ADComputer -Identity $nodeObject -PrincipalsAllowedToDelegateToAccount $gatewayObject -PassThru}
    }
