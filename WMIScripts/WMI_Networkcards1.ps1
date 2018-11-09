####this query checks if there's an Ethernet Adapter that is listed as #2. It's just a throwaway script..
####Get-CimInstance -Query "Select * From Win32_NetworkAdapter Where AdapterType like '%802.3%' and Name like '%#2%'" -ErrorAction SilentlyContinue


$FullQuery = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction SilentlyContinue
#$FullQuery.count

Get-CimInstance -Query "SELECT * FROM Win32_NetworkAdapter" -ErrorAction SilentlyContinue -PipelineVariable Adapter | %{
    #Write-Host "Adapter = $($Adapter.name)"
    #$Adapter.PNPDeviceID
    $Adapter.NetConnectionID

    if (($Adapter -in "",$null) -or ($Adapter.PNPDeviceID -in "",$null)){
        Write-Host "Adapter = $($Adapter.name). It shouldn't have been in here."
        return
    }
    $Type = $Adapter.PNPDeviceID.split("\")[0]
    
    if ($type -match "PCI|USB|ISAPNP|PCMCIA|VMBUS"){return}

    #$Adapter.Index
    $Adapter.NetConnectionID
    #$Adapter.Description
    #$Adapter.IPEnabled
}