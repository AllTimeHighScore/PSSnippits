<#
This was boosted from GitHub

https://gist.github.com/joegasper/3fafa5750261d96d5e6edf112414ae18
#>


function ConvertFrom-DN {
    [cmdletbinding()]
    param(
    [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
    [ValidateNotNullOrEmpty()]
    [string[]]$DistinguishedName
    )
    process {
        foreach ($DN in $DistinguishedName) {
        Write-Verbose $DN
            foreach ( $item in ($DN.replace('\,','~').split(","))) {
                switch ($item.TrimStart().Substring(0,2)) {
                    'CN' {$CN = '/' + $item.Replace("CN=","")}
                    'OU' {$OU += ,$item.Replace("OU=","");$OU += '/'}
                    'DC' {$DC += $item.Replace("DC=","");$DC += '.'}
                }
            } 
            $CanonicalName = $DC.Substring(0,$DC.length - 1)
            for ($i = $OU.count;$i -ge 0;$i -- ){$CanonicalName += $OU[$i]}
            if ( $DN.Substring(0,2) -eq 'CN' ) {
                $CanonicalName += $CN.Replace('~','\,')
            }
            $qwer = [PSCustomObject]@{
			    'CanonicalName' = $CanonicalName;
		    }
            Write-Output $qwer

        }
    }
}

function ConvertFrom-CanonicalUser {
    [cmdletbinding()]
    param(
    [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
    [ValidateNotNullOrEmpty()] 
    [string]$CanonicalName
    )
    process {
        $obj = $CanonicalName.Replace(',','\,').Split('/')
        [string]$DN = "CN=" + $obj[$obj.count - 1]
        for ($i = $obj.count - 2;$i -ge 1;$i--){$DN += ",OU=" + $obj[$i]}
        $obj[0].split(".") | ForEach-Object { $DN += ",DC=" + $_}
        return $DN
    }
}

function ConvertFrom-CanonicalOU {
    [cmdletbinding()]
    param(
    [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
    [ValidateNotNullOrEmpty()] 
    [string]$CanonicalName
    )
    process {
        $obj = $CanonicalName.Replace(',','\,').Split('/')
        [string]$DN = "OU=" + $obj[$obj.count - 1]
        for ($i = $obj.count - 2;$i -ge 1;$i--){$DN += ",OU=" + $obj[$i]}
        $obj[0].split(".") | ForEach-Object { $DN += ",DC=" + $_}
        return $DN
    }
}