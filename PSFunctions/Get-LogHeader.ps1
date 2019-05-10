function Get-LogHeader {
    <#
    .SYNOPSIS
        Used to create a marginally useful log header
    .DESCRIPTION
        Creates custom header for logs to be used in various scripts
    .EXAMPLE
        
    .INPUTS
        System.Object
        System.Management.Automation.InvocationInfo
        System.Version
    .OUTPUTS
        System.Array
    .NOTES
        Created on: 2018-10-12 14:13:05Z

        Created by: Kevin Van Bogart
        Inspired by Shawn's log header
    #>
        
    [CmdletBinding()]
    [OutputType([String[]])]
    param(
        #$MyInvocation used for the commandline
        [Parameter(Mandatory=$true,
                   Position=1)]
        [System.Management.Automation.InvocationInfo]$Invocation
    )
    
    #Function to find the commandline arguments
    $CommandLine = Get-WSMCommandLine -Invocation $Invocation
    $params = (($Invocation.MyCommand.Parameters.GetEnumerator() | 
    ForEach-Object {
        #Output only the BoundParameters that are present
        if ($_.key -in $Invocation.BoundParameters.Keys){
            
            #Rebuild the command line with the switch parameters e.g. -Force
            if ($_.value.SwitchParameter){
                -join ('-',$_.key)
            }
            #Rebuild the command line with the paramer names and values e.g. -Test 123
            else {
                "$(-join ('-',$_.key)) $($Invocation.BoundParameters[$_.Key])"
            }
        }
    #Find the Parameters and unbound Arguments
    })) + ' ' + $Invocation.UnboundArguments
    
    #Return the command line and the Parameters
    $CommandLine = "$($Invocation.InvocationName) $params"


    #We have to catch if we were invoked with the "&" character. 
    #In these cases there is no reliable location\.ps1 data so 
    #the timestring will have to be blank.
    if ($Invocation.InvocationName -ne "&"){ 
        $TimeStamp = (Get-Date ((Get-item -Path $Invocation.InvocationName).LastWriteTimeUtc) -Format 'yyyy-MM-dd HH:mm:ss zzz').ToString()
    }
    else {
        $TimeStamp = 'Unavailable'
    }

    $Win32OS = Get-CimInstance -ClassName Win32_OperatingSystem


    $Initialize = 'INITIALIZE'.PadRight(100,'-')
    $Action     = 'ACTION'.PadRight(100,'-')
    
#Cannot indent here-Strings
@"
$Initialize
 Commandline:        $CommandLine
 File Timestamp:     $TimeStamp
 System name:        $($env:COMPUTERNAME)
 Domain:             $(([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()).Name)
 User name:          $($env:USERNAME)
 OS Type:            $([environment]::OSVersion.VersionString)
 OS Architecture:    $($Win32OS.WMIOSArchitecture)
$Action
"@ -split "`n"

}