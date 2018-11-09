$WhereIndex = ""
$NetConnectionID = @()
$NetworkAdapter = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop
$NetworkAdapter | ForEach-Object {
    $NetworkAdapter | ForEach-Object {
        $Adapter = $_
        
        #region - Check for valid WMI entries
        #If the adapter is empty skip the rest of the work and move on to the next adapter
        $SkipAdapter = $false
        $Adapter,$Adapter.Index,$Adapter.PNPDeviceID,$Adapter.NetConnectionID | ForEach-Object {
            #Check the various properties of the adapter being inspected for validity 
            if ($_ -in "",$null){
                Write-Verbose -Message "One of the necessary WMI fields was empty. This adapter will be skipped."
                Write-Verbose -Message "The adapter is $Adapter"
                Write-Verbose -Message "Index is $($Adapter.Index)"
                Write-Verbose -Message "The Device ID is $($Adapter.PNPDeviceID)"
                Write-Verbose -Message "The net connection id is: $($Adapter.NetConnectionID)"
                $SkipAdapter = $true
                return
            }
        }
        
        #if object had a null value move on to the next adapter in the list
        if ($SkipAdapter -eq $true){
            Write-Verbose -Message "Skiping Adapter at index: $($Adapter.Index)"
            return
        }
        
        #endregion - Check for valid WMI entries
        
        #Get the first item in the PNPDevice item to see if it is a type that should be excluded.
        Write-Verbose -Message "Checking Adapter Type"
        $Type = $Adapter.PNPDeviceID.split("\")[0]
        if ($Type -notmatch "PCI|USB|ISAPNP|PCMCIA|VMBUS"){
            #Move on to the next item
            Write-Verbose -Message "This network adapter type will be excluded from configuration: $Type"
            return
        }
        
        #If there's a connection ID we'll add it to an array we can deal with later.
        #when this item is inspected we'll have to split it to get the data.
        Write-Verbose -Message "Building array of adapters that passed inspection."
        ###Adapter $($Adapter.NetConnectionID) at index $($Adapter.Index) passed initial inspection. Moving forward."
        
        #Build the array for inspection in the next block
        $NetConnectionID += "$($Adapter.Index);$($Adapter.NetConnectionID)"
        
        #Populate WhereIndex for the later query 
        if ($WhereIndex -in ""){
            $WhereIndex = "Index = $($Adapter.Index)"
            Write-Verbose -Message "Index portion of the statement will be: $WhereIndex"
        }
        else {
            $WhereIndex = -Join($WhereIndex," or Index = $($Adapter.Index)")
            Write-Verbose -Message "Index portion of the statement has been amended: $WhereIndex"
        }
    
    }#end the foreach-object where the query items are built
    
    #If there was something in the WhereIndex
    if ($WhereIndex -ne ""){
    
        #Query WMI again to get additional WMI information on each adapter.
        try {
            $AdapterConfigQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE $WhereIndex"
            $WMIQuery_AdapterConfig = Get-CimInstance -Query $AdapterConfigQuery -ErrorAction Stop
        }
        #WMI Error was detected
        catch {
            Write-Warning -Message "WMI Connection Error. Install Exiting."
            ###"WMI Connection Error. Install Exiting."
            #return 33002
        }
    
    }
}

$WhereIndex = "Index = 1"
$AdapterConfigQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE $WhereIndex"


#$AddDomain = Invoke-CimMethod -Query $AdapterConfigQuery -MethodName "SetDNSDomain" -Arguments @{DNSDomain = $DomainName} -ErrorAction SilentlyContinue
#$ChangeWINS = Invoke-CimMethod -Query $AdapterConfigQuery -MethodName "SetWINSServer" -Arguments @{WINSPrimaryServer = ""} -ErrorAction SilentlyContinue
#$ChangeWINS = (Get-CimInstance -Query $AdapterConfigQuery -ErrorAction SilentlyContinue).SetWINSServer('WINSPrimaryServer = ""')

#$ChangeWINS


#Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | ? {$_.Index -eq 1} | % {$_.SetWINSServer(",")}

#Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration | ? {$_.Index -eq 1} | % {$_.SetWINSServer($null,$null)}

#Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "Index='1'" | Invoke-WMIMethod -Name "SetWINSServer" -ArgumentList @("","")

#Invoke-CimMethod -Query $AdapterConfigQuery  -Name "SetWINSServer" -Arguments @{"";""}

$eatit = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | ? {$_.Index -eq 1} | Invoke-CimMethod -MethodName SetWINSServer -Arguments @{WINSPrimaryServer = "1.1.1.1"; WINSSecondaryServer= "1.1.1.1"}


#Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | ? {$_.Index -eq 1} | Format-List -Property * | Out-File -FilePath c:\temp\BeforeWINSChange.txt -append -force
#$NetworkAdapter | Where-Object {$_.DeviceID -eq "1"} | format-list -Property * | out-file -FilePath c:\temp\NetAdapter_beforeWINS.txt -append -force