<#
    This was boosted and slighly modified by KVB 
    
    Originally from GitHub
    https://gist.github.com/joegasper/3fafa5750261d96d5e6edf112414ae18

    Top object is OU in this object.
#>
function ConvertFrom-CanonicalOU {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)] 
        [ValidateNotNullOrEmpty()] 
        [string]$CanonicalName
    )
    process {

        $obj = $CanonicalName.Replace(',','\,').Split('/')

        [string]$DN = -Join ("OU=",$obj[$obj.count - 1])

        #Sort the OUs
        for ($i = $obj.count - 2;$i -ge 1;$i--){
            $DN += -Join (",OU=",$obj[$i])
        }

        #Deal with the first object, the domain
        $obj[0].split(".") | ForEach-Object { $DN += -Join (",DC=",$_) }
        
        #Return Domain Name
        $DN
    }
}