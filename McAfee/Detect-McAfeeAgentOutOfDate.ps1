#Check McAfee Agent Last ePO Update
##I Really don't like this logic. Should alter
##This was a result of McAfee Agent 5.0.6 and 5.5 having a service that wold randomly stop 

$32on64 = "HKLM:\SOFTWARE\Wow6432Node\Network Associates\TVD\Shared Components\Framework"
$Native = "HKLM:\SOFTWARE\Network Associates\TVD\Shared Components\Framework"
$Date = $null
$Current = (get-date).AddDays("-14")

if (test-path -Path $32on64){
    $Date = Get-ItemProperty -Path $32on64 -Name "LastUpdateCheck" -ErrorAction SilentlyContinue
    [datetime]$LastUpdateDate = [datetime]::ParseExact($Date.LastUpdateCheck,'yyyyMMddHHmmss',[datetime]::utc)
    if ($LastUpdateDate -le [datetime]$Current){
        "Update or Reinstall Agent"
    }
}
elseif (test-path -Path $Native){
    $Date = Get-ItemProperty -Path $Native -Name "LastUpdateCheck"
    if ([Datetime]$Date.LastUpdateCheck -le $Current){
    "Update or Reinstall Agent"
    }
}

