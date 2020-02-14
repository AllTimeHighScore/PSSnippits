<#
    .SYNOPSIS
        Legal Discovery Client Pamagement Script collects data on target devices and sends client scripts to collect the
         data for later processing. Formarly the 'motherScript'.
    .DESCRIPTION
        This script runs on a terminal or jump devices. It takes user input and creates a configuration file for the client side Legal Discovery script
            then copies the script and all supporting files to the target device, creates a scheduled task, and may even reboot the devices.
    .PARAMETER ParamInput
        This is the main switch that allows the script to be run without a config file.
            It is then completely dependent on other parameters to tell it what to do.
    .PARAMETER ConfigFile
        This is the primary method of operation for this script. parameters are fed into a JSON and then imported to this script.
    .PARAMETER PreferredSite
        This is the site you may prefer to utilize. The script will look for a destination that matches the site in the LegalSite.json
    .PARAMETER Drives
        The drives the script is going to backup
    .PARAMETER TechnicianID
        This is the User ID of the technician runing the Legal Discovery Process
    .PARAMETER DirEx
        Use this to List directories to exclude in the copy process
    .PARAMETER FileEx
        Use this to List Files and file types to exclude in the copy process
    .PARAMETER StopProcs
        Processes to stop that might interefere with the copy process
    .PARAMETER StopServices
        Services to stop that might interefere with the copy process
    .PARAMETER ManualDestination
        This is a manual override to allow a destiantion to be selected without relying on the Legalsite.json file
    .PARAMETER Silent
        This will prevent any popup messages or Splashscreens on the target machine. This will not suppress any cmd windows that Robocopy may display
    .EXAMPLE
        powershell.exe -executionpolicy bypass -file "C:\Temp\LegalDiscovery\LegalDiscovery_MotherScript.ps1"
    .INPUTS
        string[]
    .OUTPUTS
        LogFile - File Objects
    .NOTES
        Exit codes used by this wrapper:
            42001 - Module could not be loaded
            42002 - Admin rights needed
            42003 - Failed to obtain required paramters from dashboard
            42004 - User cancelation
            42005 - Failed to access target device
            ========== HISTORY ==========
            Author: Van Bogart, Kevin
            Created: 2019-09-23 13:18:48Z

#>
Param (
    #LogFile - File for log messages.
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$LogFile = "$PSScriptRoot\StorageProcess.log",

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [int]$LogSize = 5
)
Begin {

    #region - Load Modules
        $ModuleDir =  (Join-Path -Path $PSScriptRoot -ChildPath 'Modules')
        #Blatently lifted from wrapper and slightly modified...
        Set-Variable -Name BSCWSMModVersion -Value ([version]'1.0.2.0') -Option ReadOnly -Scope Script -Force
        #We'll try an import of our required module version from the system's module paths.
        #Importing is required because the .Path property of Get-Module, when using "-Listavailable",
        #returns the psD1 instead of the psM1. We need the latter for the sig check in the next block.
        Import-Module BSCWSMMod -RequiredVersion $BSCWSMModVersion -ErrorAction SilentlyContinue

        #Now we'll check all loaded modules and remove (from memory) any that are not our desired
        #version, or have broken signatures, leaving only the correct module provided it was
        #available for loading from a local module path and had an unbroken signature.
        $LoadedModule = Get-Module BSCWSMMod | ForEach-Object {
            if (($_.Version -ne $BSCWSMModVersion) -or ((Get-AuthenticodeSignature -FilePath $_.Path -ErrorAction SilentlyContinue).Status -ne 'Valid')){
                Remove-Module $_ -Force
            }
            else {$true} # baby don't load me, don't load me, no more...
        }

        #if there is no locally installed module, or the module has been altered
        #(breaking the signature), we'll import the one included in with the package
        if (!$LoadedModule){
            $SourceModulePath = "$ModuleDir\BSCWSMMod"
            Write-Verbose -Message 'Local module not found, or the local module has a broken signature. Loading from package...'
            #Try to import the module
            try {
                $LoadedModule = Import-Module -Name $SourceModulePath -Force -PassThru -ErrorAction Stop
            }
            #if importing didn't work, for whatever reason, we won't have
            #access to our required advanced functions and must exit.
            catch {
                $Host.SetShouldExit(42001)
                exit 42001
            }
        } # if (!$LoadedModule){

        # no try block here...too lazy
        $null = Get-Module -Name BSCWSMLDMod -ea SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue

        if (!(Get-Module -Name BSCWSMLDMod -ea SilentlyContinue)){
            $LegalDiscModPath = "$ModuleDir\BSCWSMLDMod"
            Write-Verbose -Message 'Logan Discovery module not found. Loading from package...'
            #Try to import the module
            try {
                $LDLoadedModule = Import-Module -Name $LegalDiscModPath -Force -PassThru -ErrorAction Stop
            }
            #If importing didn't work, for whatever reason, we won't have
            #access to our required advanced functions and must exit.
            catch {
                $Host.SetShouldExit(42001)
                exit 42001
            }
        } # if (!(Get-Module -Name BSCWSMLDMod -ea SilentlyContinue)){

    #endregion - Load Modules

    # Make some defaults here # not sure if I'm going to use these...
    $PSDefaultParameterValues["Write-WSMLogMessage:LogFile"] = $LogFile
    $PSDefaultParameterValues["Write-WSMLogMessage:LogSize"] = $LogSize

}
process {

    #region - isadmin
        # This is an export of the isadmin private function from the main module
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        $admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    #endregion - isadmin

    if ($principal.IsInRole($admin)){

        #region - Run Invoke-DashBoard
            #$DashParams = @{
                $LegalSites = (Get-Content -Path "\\stpnas08\legalinfo$\LegalSite.json" -ErrorAction SilentlyContinue | ConvertFrom-Json).Where({$_.Location -ne ''})
                $DirList = (Get-Content -Path "$PSScriptRoot\Configs\DirectoryExclusions.txt" -ErrorAction SilentlyContinue)
                $FileList = (Get-Content -Path "$PSScriptRoot\Configs\FileExclusions.txt" -ErrorAction SilentlyContinue)
                $ProcList = (Get-Content -Path "$PSScriptRoot\Configs\Processes.txt" -ErrorAction SilentlyContinue)
                $ServiceList = (Get-Content -Path "$PSScriptRoot\Configs\Services.txt" -ErrorAction SilentlyContinue)
            #}

            try {
                $DashBoard = Invoke-WSMDashBoard -LegalSites $LegalSites -DirList $DirList -FileList $FileList -Processes $ProcList -Services $ServiceList

                if ($DashBoard.Result -eq 'Cancel'){
                    Send-WSMMessage -Message "User Cancelled. Script Exiting." -Title 'DashBoard Cancelled' -Icon Warning
                    Invoke-WSMExitWithCode -ExitCode 42004
                }

                # Ensure the data is there
                if (($TargetComputer = $DashBoard.ComputerName) -in '',$null){
                    Send-WSMMessage -Message "Missing computer to target! Script Exiting!" -Title 'DashBoard Error' -Icon Error
                    Invoke-WSMExitWithCode -ExitCode 42003
                }

                if (($Drives = $DashBoard.Drives) -in '',$null){
                    Send-WSMMessage -Message "Missing Drives on the target computer to copy! Script Exiting!" -Title 'DashBoard Error' -Icon Error
                    Invoke-WSMExitWithCode -ExitCode 42003
                }
                elseif ($Drives -match 'Error'){
                    Send-WSMMessage -Message "Failed to contact target device! Ensure you have Domain Admin Rights and the device is on the network! Script Exiting!" -Title 'DashBoard Error' -Icon Error
                    Invoke-WSMExitWithCode -ExitCode 42005
                }

                # Check for the selected destination
                elseif ( ( ($Destination = $DashBoard.Destination) -in '',$null) -or ( !(Test-path -Path $DashBoard.Destination -ErrorAction SilentlyContinue) -and ($DashBoard.Destination -notmatch 'USB') ) ){
                    Send-WSMMessage -Message "Missing valid copy destination! Script Exiting!" -Title 'DashBoard Error' -Icon Error
                    Invoke-WSMExitWithCode -ExitCode 42003
                }

                #Create warning messages if any of the other items are missing.
                if (!($DashBoard.ProcList)){
                    Send-WSMMessage -Message "Warning! No processes were listed to stop. This may create or increase errors with the copy process" -Title 'DashBoard Warning' -Icon Warning
                }

                if (!($DashBoard.ServiceList)){
                    Send-WSMMessage -Message "Warning! No Services were listed to stop. This may create or increase errors with the copy process" -Title 'DashBoard Warning' -Icon Warning
                }

                if (!($DashBoard.DirExclusions)){
                    Send-WSMMessage -Message "Warning! No Directories will be excluded. This `'Will`' create and increase errors with the copy process" -Title 'DashBoard Warning' -Icon Warning
                }

                if (!($DashBoard.FileExclusions)){
                    Send-WSMMessage -Message "Warning! No files or file types will be excluded. This `'Will`' create and increase errors with the copy process" -Title 'DashBoard Warning' -Icon Warning
                }
            }
            catch [System.Management.Automation.ParameterBindingException]{
            #catch [System.Management.Automation.ParameterBindingValidationException]{
                Send-WSMMessage -Message "Missing required parameter! Script Exiting!" -Title 'DashBoard Error' -Icon Error
                Invoke-WSMExitWithCode -ExitCode 42003
            }
            catch {
                Send-WSMMessage -Message "Unexpected Error: $($_.exception.gettype().fullname). Script exiting!" -Title 'Target Device Inspection' -Icon Error
                Invoke-WSMExitWithCode -ExitCode 42003
            }

        #endregion - Run Invoke-DashBoard

        #region - Copy items to target device (Even USB if applicable)

            # Robocopy arguments
            $SourceDir = "$PSScriptRoot\ClientScript"
            if ($TargetComputer -eq $env:COMPUTERNAME){
                # Don't use temp env var in case user is not admin
                $DestDir = "$env:windir\Temp\LegalDiscovery"
            }
            elseif ($TargetComputer -eq 'LOCALHOST'){
                # Select a USB Drive
                $USBDriveSelect = Select-WSMSaveLocation
                if ($USBDriveSelect.Result -eq 'Success'){
                    $DestDir = "$(($USBDriveSelect).Path)\LegalDiscovery"

                    try {
                        $AddInitScript = Copy-Item -Path "$PSScriptRoot\InitializeLocalClientScript.exe" -Destination ($USBDriveSelect).Path -ErrorAction Stop
                    }
                    catch {
                        Send-WSMMessage -Message "Failed to transfer Initialization script. Script Exiting!" -Title 'USB Drive Prep' -Icon Error
                        Invoke-WSMExitWithCode -ExitCode 42004
                    }
                    Send-WSMMessage -Message "A USB Drive $DestDir selected." -Title 'USB Drive Prep'
                }
                else {
                    Send-WSMMessage -Message "A USB Drive must be selected to move forward. Script Exiting!" -Title 'USB Drive Prep' -Icon Error
                    Invoke-WSMExitWithCode -ExitCode 42004
                }
            }
            else {
                # The standard command
                $DestDir = "\\$TargetComputer\c$\Windows\Temp\LegalDiscovery"
            }

            # Assemble robocopy args
            $RetryRoboArgs = "$SourceDir $DestDir * /COPY:DAT /e /z /MT:100 /xjd /np /r:5 /w:5 /njh /njs /A-:SH"

            # Copy Items
            $TrueUpCopy = Start-Process -FilePath 'Robocopy.exe' -ArgumentList $RetryRoboArgs -PassThru
            $TrueUpCopy | Wait-Process -Timeout 600 -ErrorAction SilentlyContinue
        #endregion - Copy items to target device (Even USB if applicable)

        #region - Setup scheduled task
            
            if ($Destination -notmatch 'USB'){
                <#
                    The PowerShell cmdlts for creating scheduled tasks have a host of issues that have so far made it difficult to reliably configure for this task.
                    For now this script will leverage the built in legacy program schtasks to accoplish the task.
                #>

                # Prompt user for credentials - Always.

                #region - Get valid credentials
                <#
                    This was sourced from https://blogs.technet.microsoft.com/dsheehan/2018/06/23/confirmingvalidating-powershell-get-credential-input-before-use/
                #>

                    # Prompt for Credentials and verify them using the DirectoryServices.AccountManagement assembly.
                    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                    $Attempt = 1
                    $MaxAttempts = 3
                    $CredentialPrompt = "Enter credentials able to access all areas of targeted device and has write access to destination drive:"
                    $ValidAccount = $false

                    # Loop through prompting for and validating credentials, until the credentials are confirmed, or the maximum number of attempts is reached.
                    do {
                        # Blank any previous failure messages and then prompt for credentials with the custom message and the pre-populated domain\user name.
                        $FailureMessage = $Null
                        $Creds = Get-Credential -Message $CredentialPrompt
                        # Verify the credentials prompt wasn't bypassed.
                        if ($Creds){
                            # if the user name was changed, then switch to using it for this and future credential prompt validations.
                            if ($Creds.UserName -ne $UserName){
                                $UserName = $Creds.UserName
                            }
                            # Test the user name and password.
                            $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                            try {
                                $PrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ContextType,$UserDomain
                            } 
                            catch {
                                if ($_.Exception.InnerException -like "*The server could not be contacted*"){
                                    $FailureMessage = "Could not contact a server for the specified domain on attempt #$Attempt out of $MaxAttempts."
                                } 
                                else {$FailureMessage = "Unpredicted failure: `"$($_.Exception.Message)`" on attempt #$Attempt out of $MaxAttempts."}
                            }
                        
                            # if there wasn't a failure talking to the domain test the validation of the credentials, and if it fails record a failure message.
                            if (!($FailureMessage)){
                                $ValidAccount = $PrincipalContext.ValidateCredentials($UserName,$Creds.GetNetworkCredential().Password)
                                if (!($ValidAccount)){
                                    $FailureMessage = "Bad user name or password used on credential prompt attempt #$Attempt out of $MaxAttempts."
                                }
                            }
                        # Otherwise the credential prompt was (most likely accidentally) bypassed so record a failure message.
                        } 
                        else {$FailureMessage = "Credential prompt closed/skipped on attempt #$Attempt out of $MaxAttempts."}
                    
                        # if there was a failure message recorded above, display it, and update credential prompt message.
                        if ($FailureMessage){
                            Write-Warning -message "$FailureMessage"
                            $Attempt++
                            if ($Attempt -lt $MaxAttempts) {
                                $CredentialPrompt = "Authentication error. Please try again (attempt #$Attempt out of $MaxAttempts):"
                            } elseif ($Attempt -eq $MaxAttempts) {
                                $CredentialPrompt = "Authentication error. THIS IS YOUR LAST CHANCE (attempt #$Attempt out of $MaxAttempts):"
                            }
                        }
                    } until (($ValidAccount) -or ($Attempt -gt $MaxAttempts))
                

                #endregion - Get valid credentials

                # The only reason to send the command to the target device like this is to ensure the creds are encrypted.
                try {
                    # Send the command to the remote device
                    [string]$SetTask = Invoke-Command -ComputerName $TargetComputer -Credential $Creds -Scriptblock {
                        Param($Creds)

                        # Check for existing sched task
                        $null = Get-ScheduledTask -TaskName 'LegalDisc' -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

                        $TaskPath = 'Powershell.exe -executionpolicy bypass -File "C:\Windows\Temp\LegalDiscovery\LegalDiscovery_ClientScript.ps1"'
                        $SchTaskArgs = "/create /SC ONSTART /tn LegalDisc /tr ""$TaskPath"" /ru ""$($Creds.UserName)"" /rp ""$($Creds.GetNetworkCredential().password)"""
                        $CreateSchedTask = Start-Process -FilePath SchTasks.exe -ArgumentList $SchTaskArgs -PassThru
                        $CreateSchedTask | Wait-Process -Timeout 60 -ErrorAction SilentlyContinue

                        # Return the results in a more easily parsable object
                        (Get-ScheduledTask -TaskName 'LegalDisc' -ErrorAction SilentlyContinue).state

                    } -ArgumentList $Creds -ErrorAction Stop
                }
                catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
                    Send-WSMMessage -Message "Error: Credentials lack access to $TargetComputer. Script exiting!" -Title 'Bad Credentials' -Icon Error
                    Remove-Variable -name creds -ErrorAction SilentlyContinue
                    Invoke-WSMExitWithCode -ExitCode 1
                }

                if ($SetTask -eq 'Ready'){
                    Send-WSMMessage -Message 'Legal Discovery Scheduled Task Set Successfully!' -Title 'Success'
                }
                else {
                    Send-WSMMessage -Message "Error setting scheduled task. Result: $($SetTask.State) Script exiting!" -Title 'Scheduled Task Failure' -Icon Error
                    Remove-Variable -name creds -ErrorAction SilentlyContinue
                    Invoke-WSMExitWithCode -ExitCode 1
                }

                # Extract tech's user info before nulling the variable out.
                [string]$RunningTechID = $Creds.UserName

                # Inspect the user's AD record and attempt to populate the full name.
                [string]$RunningTechName = ((([adsisearcher]"(&(objectCategory=User)(samaccountname=$($Creds.UserName)))").findall()).properties).displayname
                if ($RunningTechName -in $null,$RunningTechName){
                    $RunningTechName = ''
                }

                # Clean the techs credentials out
                Remove-Variable -name creds -ErrorAction SilentlyContinue

            } # if ($Destination -notmatch 'USB'){
        #endregion - Setup scheduled task

        #region - send results
            
            # create custom object as a config.
            $ConfigObject = [PSCustomObject]@{
                TechID = $RunningTechID
                TechName = $RunningTechName
                PreferredSite = $DashBoard.Site
                Drives = $Drives
                DirEx = [string[]]$DashBoard.DirExclusions
                FileEx = [string[]]$DashBoard.FileExclusions
                StopProcs = [string[]]$DashBoard.ProcList
                StopServices = [string[]]$DashBoard.ServiceList
                ManualDest = $Destination
            }

            if ($DashBoard.Destination -notmatch 'USB'){

                #Send object to remote device
                $ConfigObject | ConvertTo-Json | Out-File -FilePath "\\$TargetComputer\c$\Windows\Temp\LegalDiscovery\Config.json" -Force -ErrorAction SilentlyContinue

                if ($TargetComputer -notin $null,'',$env:COMPUTERNAME){
                    Restart-Computer -ComputerName $TargetComputer -ErrorAction SilentlyContinue -Force
                    Send-WSMMessage -Message "$TargetComputer will be rebooted and data collection will begin. Please do not log into $TargetComputer while data is being collected for legal discovery" -Title 'Data Collection Preparation Complete'
                }
                elseif ($TargetComputer -in $env:COMPUTERNAME,'localhost') {
                    Send-WSMMessage -Message "Complete any tasks and reboot this device to begin data collection." -Title 'Data Collection Preparation Complete'
                }
            }
            else {
                #Need new GUI to send to USB
                $ConfigObject | ConvertTo-Json | Out-File -FilePath (Join-path -path $DestDir -ChildPath 'Config.json') -Force -ErrorAction SilentlyContinue
                "WSM TagFile: USB Drive" | Out-File -FilePath "$(($USBDriveSelect).Path)\LegalDisc.tag" -Force -ErrorAction SilentlyContinue
                Send-WSMMessage -Message "USB drive has been loaded with the Legal Discovery collection scripts." -Title 'Data Collection Preparation Complete'
            }

        #endregion - send results
    } # if (isAdmin){
    else {
        Send-WSMMessage -Message 'This tool must be run with administrative rights!' -Title 'Data Collection Preparation Arborted' -Icon Error
        Invoke-WSMExitWithCode -ExitCode 42002
    }
}
# SIG # Begin signature block
# MIIcigYJKoZIhvcNAQcCoIIcezCCHHcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUg4ecGsDFuFdphD2mKrJK/WsG
# i4egghe5MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFQjCC
# BCqgAwIBAgIQChqZGcgY1Mp84Z+CMF8p5jANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTE4MDIyMjAwMDAwMFoXDTIxMDIyNjEyMDAwMFowfzEL
# MAkGA1UEBhMCVVMxFjAUBgNVBAgTDU1hc3NhY2h1c2V0dHMxFDASBgNVBAcTC01h
# cmxib3JvdWdoMSAwHgYDVQQKExdCb3N0b24gU2NpZW50aWZpYyBDb3JwLjEgMB4G
# A1UEAxMXQm9zdG9uIFNjaWVudGlmaWMgQ29ycC4wggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQClLS1SCiz4y/YiZCMO5SvZn5dPm2RSCqjZGG4mzBcvzxo8
# vAHNb49i7IQF8qD1WmlsjqPsVunHf5UM2Wkc1R+gqfn4I7dkOcLaBMFC00d5rUOe
# QKHUrmsIXJzhSXMAntklFBwqohvUFl6qPbAJRr8THIljuSZbilaxBv8MVYkrR32B
# wzpApm2BR+/cvke4JbJGORpuOZw4DWgPMz0i4Nr8tLw63E74b9JZhv7IvGriugCT
# alOzRsLSCVMyL+7Dwgiv+7ajPJRULWNfDkkFFMVPMmETGs/zGXp+9+GtCdOGijJq
# xq3jBN7ckDSEv+ZFilrUXOg3JA4eiex3+lSWbKpJAgMBAAGjggHFMIIBwTAfBgNV
# HSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUBxYoD7RsSkRN
# oOc0IipVmldtepQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEy
# LWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5j
# b20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwD
# ATAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgG
# BmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAM
# BgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCdHHgxSUEv8wT+pv9t+PEq
# LVZe6aCS2/qcUwQEstQasalVSYLzy09lKNEiSeH5a37UZoUKV1YMdpK/LBl9IKQl
# ODKt32DThZUKHQ4SgaKMMd4wjg52TB5rBk0DqU7Jdju/gS19ac6zABV8tyGumGA6
# DnI/x4gD+sW5LHWLju7Xof7j3gu4Feq8z1scAQu3QEMCpZFMPk5bw6+h7HTi5Z6F
# 51QYhXUVTe2YYfd4wAJkca6cJItCxFBa1lXabFDpIlT/+CVIFqbYv789pABEeRur
# 2zybIadlJbKK5FOz3PE9FpYHGLUJICWrsrcKyrerR6B7cRBevOI36HEYAYvJ2F/t
# MIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENB
# LTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJV
# UzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFt
# cCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38
# fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu
# 3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75h
# n67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchho
# ubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kG
# LkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xi
# QpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0T
# AQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGy
# MIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGln
# aWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMA
# ZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8A
# bgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAA
# dABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAA
# dABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0A
# ZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQA
# eQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgA
# ZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1s
# AxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFa
# TSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJo
# dHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNy
# bDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2lj
# ZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0l
# fhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPve
# BgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IR
# Qaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0On
# cZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdm
# yKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQU
# Cjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0G
# CSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0
# IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAw
# MDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQg
# SUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJC
# laWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxAra
# r8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2
# P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VI
# xO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzy
# r5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8y
# FGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAy
# BggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUH
# AwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUF
# BwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnku
# aHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAA
# dABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQA
# dQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQA
# aQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIA
# ZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcA
# aABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQA
# IABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4A
# IABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMB
# Af8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8E
# ejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQft
# Ht+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkq
# hkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUf
# dDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp
# 5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPk
# KqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3W
# g+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN
# /CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGCBDswggQ3AgEBMIGGMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0ECEAoamRnIGNTKfOGfgjBfKeYwCQYFKw4DAhoFAKB4
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFFjEF0FOF+bpu+W+itFEaMB5AseAMA0GCSqGSIb3DQEBAQUABIIBAFP2wDiJ
# 9/JEEar9vQnp9ZOssE2WzBnOPJdfXs6MsEoac+nAiHrVqOhDD70OveNI7g4enS5W
# tkBaLBv8ngYy8LHNw60Vn3egqX+9kibvMWAkiwo6UN229IqhW6gdV2NPmvZReuEk
# sce3EuD6AbwlEPqE6Op+6EPDumhosfeKu7Hm5kYm6hoQ6u98dWt4Dh2xUsR9mdcc
# LOt4W0DbGNDyMc+D+FHcx39WX4eBAWZgZs+RkzFsdSXzWx5sdD7qieb1q9Th3iRL
# kIzTUPZevRpt9mm9xg9UrLsudCU5qvGyj7yYyqREF8a/EmknXPo9Lz1YCflmN8sB
# YjIpdKy668HXNnmhggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEC
# EAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIwMDIxMTIwMjAwMlowIwYJKoZIhvcN
# AQkEMRYEFGAq64t0MTAJqCtEJsVQ2HCiEVyeMA0GCSqGSIb3DQEBAQUABIIBAAIM
# 9MBwiR0VfRnHaz5dX1tCfsql+ZtPw7f02toYqGoOYEB4afhAbR1g9PiIvWqks5ML
# FI/0S1nEs2YbYbzQZj43iODAWPyE4fQQRwZxum3tITbC/J6vIRxf8ZZ8HOlWA9c4
# xKKWVHnNLZFCUac7TQvHNn2NZStKKIv2zb4Aw0eLIho/P2KFv76H6UaLX8b8rT7f
# YE5rVPk0WsE3KMSDeTI+0jk+lvUR+zQ22Mp/n/6B0+9V00Rpt6nhdiEw2g6AW0cR
# 2+kseKYSsqt92NwSDNAgIQ8heVjBUrAaGR6CNgvOW4cRKOdK+ivrn2prIJQaDB80
# D1zyARKCqGRROf3nagI=
# SIG # End signature block
