<#
Details: This is meant to be used to fool SCCM into thinking Office is installed for testing purposes
Author: Kevin Van Bogart
#>

#Initialize array splat
$Key2Spoof = @()

$Key2Spoof += @{
    Key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us'
    Name = ''
    Value = ''
}

$Key2Spoof += @{
    Key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    Name = 'Platform'
    Value = 'x86'
}

$Key2Spoof += @{
    Key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    Name = 'CDNBaseUrl'
    Value = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
}

if (!(Get-WSMUninstallData -DisplayName "*Office*") -and (!(Test-path -Path $($Key2Spoof[1].Key)))){

    $Key2Spoof | % {

        if (!(Test-path -Path $($_.Key))){
            try {
                $null = New-Item -Path $($_.Key) -Force -ea stop
                "Successfully created key: $($_.Key)"
            }
            catch {"Failed to create key: $($_.Key)"}
        }

        if ($_.Name){
            try {
                $null = Set-ItemProperty -Path $($_.Key) -Name $($_.Name) -Value $($_.Value) -Force -ErrorAction Stop
                "Successfully populated Name: $($_.Name) Value: $($_.Value)"
            }
            catch {"Failed to populate Name: $($_.Name) Value: $($_.Value)"}
        }
    }

}
elseif (!(Get-WSMUninstallData -DisplayName "*Office*") -and ((Test-path -Path $($Key2Spoof[1].Key)))){
    $Key2Spoof | % {
        if ((Get-Item -Path $($_.Key) -ea SilentlyContinue)){
            try {
                $null = Remove-Item -Path $($_.Key) -force -ea stop
                "Successfully Removed key $($_.Key)"
            }
            catch {"Failed to Remove Item key: $($_.Key)"}
        }
    }

}
elseif (Get-WSMUninstallData -DisplayName "*Office*"){
    "Office appears to be installed, no action will be taken"
}




