<#
.Synopsis
   Server aus der SA nach CSV
.DESCRIPTION
   Server aus der SA auslesen, um sie ins WAC zu transportieren
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
    
    ## Neue Klasse, neues Objekt
    class clsWAC {
    [string] $name
    [string] $type
    [string] $tags
    [string] $groupId
    }
    
    $objWAC=New-Object -TypeName clsWac
    [object[]]$objWACs=$null
    [object[]]$objTags=$null
    if ($objTag) {$objTag=$null}
        
    ## Status enthält das Erstellungsdatum, daraus eine Liste filtern
    #$Serversa=$serverService.getServerVOs($serverRef_) | 
    #    Where-Object {$_.osFlavor -match 'win' -and $_.createdDate -ge $((get-date).AddDays(-10))} | Select-Object @{Name='Name';expression={$_.hostname}},createdDate
    ## ohne das Erstellungsdatum zu berücksichtigen, daraus eine Liste filtern
    
    $serverRef_ | ForEach-Object {
        $Serversa=$null
        $custAttr=$null
        $ServerRef=$_
        $Serversa=$serverService.getServerVO($ServerRef) | 
            Where-Object {$_.osFlavor -match 'win' } #| Select-Object @{Name='Name';expression={$_.hostname}},createdDate
            if ($Serversa) {
            ## Attribut Wartungsfenster
                [string]$strWaFe=$(($serverService.getCustAttrKeys($ServerRef)).Where({$_.contains("Monat")}))
                if ($strWaFe) {$strWaTag=($strWaFe.Trim()).Split(" ")[0]
                $custWartFe=$serverService.getCustAttr($ServerRef,$strWaFe,$tru)  
                $strWartFeZeit="$($custWartFe.Split("-")[0]):00"}             
            ## Attribut Verfahren
                if ($serverService.getCustAttrKeys($ServerRef) | Where-Object {$_ -like "Verfahren"}) {
                $custAttr=$serverService.getCustAttr($ServerRef,"Verfahren",$true)
                }
     
    $objWAC=New-Object -TypeName clsWac
    $objWAC.name=$Serversa.hostName;
    $objWAC.type="msft.sme.connection-type.server";
    [string[]]$mytags=$null
    [string[]]$mytagsmecm=$null
    if ($custAttr) {$mytags="6-Ver:$custAttr"} 
       else {$mytags="unknown"}
    if ($custAttr) {$mytagsmecm="$custAttr"} 
       else {$mytagsmecm="unknown"}
    $mytags+="3-Dom:$(($Serversa.hostName).split('.')[1])"
    $mytagsmecm+=($Serversa.hostName).split('.')[1]
    $mytags+="2-Zone:$(($Serversa.hostName).Substring(0,4))"
    $mytagsmecm +=($Serversa.hostName).Substring(0,4)
    $mytags+="1-BS:$((($Serversa.osFlavor).trim()).split(" ")[2])"
    $mytagsmecm+=(($Serversa.osFlavor).trim()).split(" ")[2]
    $mytags+="4-WF:$strWaTag"
    $mytagsmecm+=$strWaTag
    $mytags+="5-WZ:$strWartFeZeit"
    $mytagsmecm+=$strWartFeZeit
    #$mytags+=$Serversa.use
    $mytagsmecm+=$Serversa.use
    $tagcount= $mytags.count
    [PSObject]$objTag=New-Object psobject -Property @{
    Servername = $Serversa.hostName
    Verfahren = $mytagsmecm[0]
    Domain = $mytagsmecm[1]
    OSVersion = $mytagsmecm[3]
    WartungsfensterTag=$mytagsmecm[4]
    WartungsfensterZeit=$mytagsmecm[5] 
    Stage=$mytagsmecm[6]
    }
    $objWAC.tags= $mytags | ForEach-Object {$tagcount--;if ($tagcount -gt 0) {"$($_) |"} else {$_}}
    #$objWAC.tags = "$($mytags[0]) | $($mytags[1]) | $($mytags[2]) | $($mytags[3])"
    $objWAC.groupId="global"
    $objWACs+=$objWAC
    $objTags+=$objTag
      }
    }

## Dateiname (Prod und Test)
$NameCSV="AllServer_For_WAC.csv"
$NameCSVMCM="AllServerTags.csv"

## Pfade Testumgebung
$pathCSVFolder="\\srvws40123\d$\AdminCenterDaten"
$pathCSV="$pathCSVFolder\$NameCSV"
$pathCSVMCM="$pathCSVFolder\$NameCSVMCM"

## Pfade Produktionsumgebung
$pathCSVFolderProd="\\srvwsv40965\d$\AdminCenterDaten"
$pathCSVProd="$pathCSVFolderProd\$NameCSV"
$pathCSVMCMProd="$pathCSVFolderProd\$NameCSVMCM"

## Test
$objWACs | Export-Csv -Path $pathCSV -NoTypeInformation #-Append
$objTags | Export-Csv -Path $pathCSVMCM -NoTypeInformation #-Append

## Prod
$objWACs | Export-Csv -Path $pathCSVProd -NoTypeInformation #-Append
$objTags | Export-Csv -Path $pathCSVMCMProd -NoTypeInformation #-Append

#$objWACs | ConvertTo-Json | Out-File -FilePath .\AllServer_For_WAC.json #-Append


#Import-Connection 'https://cc03wsv3578:6516' -fileName .\WACHosts.csv 