
<#
    This was boosted and slighly modified by KVB 
    
    Originally from GitHub
    https://gist.github.com/joegasper/3fafa5750261d96d5e6edf112414ae18

    Feeling cute,
    Maybe I take these later, mate them all into one and allow a master functionto take an AD object and spit out all both names.
    IDK...  Stupid arse meme... 
#>

function ConvertFrom-DN {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$DistinguishedName
    )
    process {
        foreach ($DN in $DistinguishedName){

            #Write for troubleshooting
            Write-Verbose $DN

            foreach ($item in ( $DN.replace('\,','~').split(",") )){

                switch ( $item.TrimStart().Substring(0,2) ){
                    'CN' {$CN = '/' + $item.Replace("CN=","")}
                    'OU' {$OU += ,$item.Replace("OU=",""); $OU += '/'}
                    'DC' {$DC += $item.Replace("DC=",""); $DC += '.'}
                }
            } #foreach

            $CanonicalName = $DC.Substring(0,$DC.length - 1)

            for ($i = $OU.count;$i -ge 0;$i -- ){$CanonicalName += $OU[$i]}

            if ( $DN.Substring(0,2) -eq 'CN' ){
                $CanonicalName += $CN.Replace('~','\,')
            }

            [PSCustomObject]@{'CanonicalName' = $CanonicalName}

        } #foreach
    }#Process
}