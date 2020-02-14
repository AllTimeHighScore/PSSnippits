<#
.SYNOPSIS
    Legal Discovery copy script
.DESCRIPTION
    This is the client side portion of the Legal Discovery copy script.
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
    This is the User ID of the technician running the Legal Discovery Process
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
    Config file would be formatted properly and in the same directory as the script.
    PowerShell.exe -executionpolicy bypass -file "C:\TEMP\LegalDiscTesting\LegalDiscovery_Client.ps1" -ConfigFile

    The config file parameter isn't actually needed. The script will look for the config file if 'ParamInput' isn't specified.

.INPUTS
    string[]
.OUTPUTS
    LogFile - File Objects
.NOTES
    Exit codes used by this wrapper:
        42001 - Module could not be loaded
        42002 - Admin rights needed #---------> not usig this 
        42003 - failed to load config file
        42004 - Failed to load LegalSites Json
        42005 - Failed to obtain required paramters from dashboard
        42006 - Failed to create destination path
        42007 - Failed to Access Destination path

    ========== HISTORY ==========
    Author: Van Bogart, Kevin
    Created: 2019-09-23 13:18:48Z

#>
param (
    # In case the config file is empty
    [Parameter(ParameterSetName='ParamInput')]
    [switch]$ParamInput,

    [Parameter(ParameterSetName='ConfigFile')]
    [switch]$ConfigFile,

    [Parameter(ParameterSetName='ParamInput', Mandatory=$true)]
    [string]$PreferredSite,

    #This is the 'source'
    [Parameter(ParameterSetName='ParamInput', Mandatory=$true)]
    [string[]]$Drives,

    [Parameter(ParameterSetName='ParamInput')]
    [string]$TechnicianID,

    [Parameter(ParameterSetName='ParamInput')]
    [string[]]$DirEx,

    [Parameter(ParameterSetName='ParamInput')]
    [string[]]$FileEx,

    [Parameter(ParameterSetName='ParamInput')]
    [string[]]$StopProcs,

    [Parameter(ParameterSetName='ParamInput')]
    [string[]]$StopServices,

    # This param will override any site selection
    [Parameter(ParameterSetName='ParamInput')]
    [string]$ManualDestination,

    #Disables the splashscreen, which is by default on.
    [switch]$Silent,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [ValidateScript({(Get-item -path $_ -ErrorAction SilentlyContinue).Extension -eq '.JSON'})]
    [string]$LegalSitesFile,


    #Set the package version
    [version]$ScriptVersion = '1.0.0.0'
)
Begin {
    $Return = 0
    $SchTN = 'LegalDisc'
    $USB = $false

    #region - Add Modules
        $ModuleDir = (Join-Path -Path $PSScriptRoot -ChildPath "Modules")
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

        #If there is no locally installed module, or the module has been altered
        #(breaking the signature), we'll import the one included in with the package
        if (!$LoadedModule){
            $SourceModulePath = "$ModuleDir\BSCWSMMod"
            Write-Verbose -Message 'Local module not found, or the local module has a broken signature. Loading from package...'
            #Try to import the module
            try {
                $LoadedModule = Import-Module -Name $SourceModulePath -Force -PassThru -ErrorAction Stop
            }
            #If importing didn't work, for whatever reason, we won't have
            #access to our required advanced functions and must exit.
            catch {
                #$Host.SetShouldExit(32001)
                $Return = 42001
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
                if ($LDLoadedModule.ExportedCommands.Values.name -notcontains 'Invoke-WSMValidatedCopySingleItems'){
                    $Return = 42001
                    $LegalDiscModuleMessage = "Failed to load the Legal Discovery module"
                }
                else {
                    $LegalDiscModuleMessage = "Legal Discovery module loaded: $LDLoadedModule"
                }
            }
            #If importing didn't work, for whatever reason, we won't have
            #access to our required advanced functions and must exit.
            catch {
                #$Host.SetShouldExit(32001)
                $Return = 42001
            }
        } # if (!(Get-Module -Name BSCWSMLDMod -ea SilentlyContinue)){

    #endregion - Add Modules

    #region - Set some basic values related to the functions and modules jsut loaded
        $InitData = Get-WSMInitData -PackageName $PackageName
        #Not sure if this location will change to a remote LD server at some point...
        # FileName is slightly different here because of the log file function.
        $LogFileNames = Get-WSMLogFileNames -LogLocation $PSScriptRoot -AppName "$($env:COMPUTERNAME)-LegalDiscovery"
        Set-Variable -Name LogFile -Value $LogFileNames.LogFile -Option ReadOnly -Scope Script -Force
        #Add a default param for package logging so we don't have to continually type "-LogFile $LogFile"
        $PSDefaultParameterValues["Write-WSMLogMessage:LogFile"] = $LogFile
        #Write the standard log header
        $WSMLogHeaderData = Get-WSMLogHeader -InitData $InitData -TemplateVersion $ScriptVersion -Invocation $MyInvocation
        $WSMLogHeaderData | Write-WSMLogMessage
        Write-WSMLogMessage -Message "BSCWSMMod module loaded: $LoadedModule"
        Write-WSMLogMessage -Message $LegalDiscModuleMessage
    #endregion - Set some basic values related to the functions and modules jsut loaded

    #region - Add functions without modules
        function Remove-WSMScheduledTask {
            <#
                .Notes
                Possibly the laziest function I've ever made....
            #>
            param (
                [Parameter(Mandatory=$true)]
                [string]$Taskname,
        
                [Parameter(Mandatory=$false)]
                $LogFile
            )
            Process {

                if ($Logfile){Write-WSMLogMessage -Message "About to remove Scheduled task $TaskName." -LogFile $logfile}

                try {
                    $DisableTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -ErrorAction Stop -Confirm:$false
                    if ($Logfile){Write-WSMLogMessage -Message "Scheduled task $TaskName has been unregistered. Result: $DisableTask" -LogFile $logfile}
                }
                catch {
                    if ($Logfile){Write-WSMLogMessage -Message "Error: Could not remove scheduled task $TaskName : $($_.Exception.Message)" -LogFile $logfile}
                }
            } # Process {    
        } # function Remove-WSMScheduledTask {...}
    #region - Add functions without modules

    $CompSys = Get-CimInstance -Namespace 'root/cimv2' -ClassName 'Win32_ComputerSystem' -ErrorAction SilentlyContinue
    $Baseline = (Get-ItemProperty -Path 'HKLM:\Software\SpecialKey\Baseline' -ErrorAction SilentlyContinue)
    $Company = (Get-ItemProperty -Path 'HKLM:\Software\SpecialKey\Company' -ErrorAction SilentlyContinue)
    $OS = (Get-CimInstance -Namespace 'root/cimv2' -ClassName 'Win32_OperatingSystem' -ErrorAction SilentlyContinue)

    if (!$ParamInput){
        try {
            $Config = Get-Content -Path "$Psscriptroot\Config.json" -raw | ConvertFrom-Json
        }
        catch {
            Write-WSMLogMessage -Message "Failed to load the config file. There will not be enough information available to begin the legal discovery process."
            $Return = 42003
        }
    } # if (!$ParamInput){

    # make sure we have a place to put the final report.
    if (!$PreferredSite -and (([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name) -notin '',$null)){
        $ReportSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    }
    elseif ($PreferredSite){
        $ReportSite = $PreferredSite
    }
    else {
        # Use STP as a default
        $ReportSite = 'STP-STPAUL'
    }

    # Gather info for later usage and output.
    $ComputerTXT = [PSCustomObject]@{
        ComputerName = $CompSys.Name
        Vendor = $CompSys.Manufacturer
        ComputerModel = $CompSys.Model
        OperatingSystem = "$($OS.Caption) $($OS.Version)"
        Language = $Baseline.Language
        ApplicationSuite = $Baseline.AppType
        Site = $ReportSite
    } # $ComputerTXT = [PSCustomObject]@{

    #Attempt to get info on any logged in user.
    try {
        $LogonInfo = (Get-CimInstance -classname Win32_LoggedOnUser -ErrorAction Stop |
        Where-Object {
            $_.Dependent.LogonId -in ( Get-CimInstance -ClassName Win32_LogonSession -ErrorAction Stop | Where-Object {$_.AuthenticationPackage -eq 'Kerberos'} ).LogonId
        }).Antecedent.Name
        Write-WSMLogMessage -Message "Successfully queried logon info"
    }
    catch {
        #try to get the info using the query tool
        Write-WSMLogMessage -Message "Failed to obtain login data via WMI. Will attempt the query tool."
        $LogonInfo = Invoke-Command -ScriptBlock {query user} | ForEach-Object {
            (($_.trim() -replace ">" -replace "(?m)^([A-Za-z0-9]{3,})\s+(\d{1,2}\s+\w+)", '$1  none  $2' -replace "\s{2,}", "," -replace "none", $null))
        } | ConvertFrom-Csv | Where-Object {$_.State -EQ 'Active'}
    }

    # query our primary user data. This should be available on most managed devices.
    if ( !($PrimaryUserID = (Get-ItemProperty -Path 'HKLM:\Software\SpecialKey\NetProfile' -ErrorAction SilentlyContinue).PrimaryUser) ){
        #May alter this when I revisit the old code, if there is anything we can do at all about that.
        $PrimaryUserID = 'N/A'
    }

    #Last logged on user
    [string]$LastLoggedOn = (Get-ChildItem -Path 'c:\Users' -Directory |
        Where-Object {$_.name -notmatch 'Public|Default|All Users'} |
            Sort-Object -Property LastWriteTime -Descending |
                Select-Object -First 1).name

    if ($LastLoggedOn){
        try {
            $ADLU = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$LastLoggedOn))").findall()).properties
            if ($ADLU.displayname -match '\{\}'){
                [string]$LastUserFullname = ($ADLU.displayname).replace('\{\}','')
            }
            else {
                [string]$LastUserFullname = $ADLU.displayname
            }
            Write-WSMLogMessage -Message "Successfully queried last user's Displayname ($LastUserFullname) from ($LastLoggedOn) in AD"
        }
        catch {
            Write-WSMLogMessage -Message "Failed to query Displayname from last user's ID ($LastLoggedOn) in AD: $($_.Exception.Message)"
            $LastUserFullname = 'N/A'
        }
    } # if ($LastLoggedOn){

    if ($ParamInput){
        $TechID = $TechnicianID
        $TechName = $null
    }
    else {
        $TechID = $Config.TechID
        $TechName = $Config.TechName
        $PreferredSite = $Config.PreferredSite

        # Check manual destination

        if ($Config.ManualDest -eq 'USB'){

            $USB = $true
            # Assign the USB as the destination
            [string]$ManualDestination = (Search-WSMUSBPorts | Where-Object {(Get-Item -Path "$($_.Name)\LegalDisc.tag" -ErrorAction SilentlyContinue).Exists -eq $true} | Select-Object -First 1).Name
            if (!$ManualDestination){
                $USBDriveErrMSG = "A site override to save data in USB has been input. However, You need to attach a large capacity USB drive with the `'LegalDisc.tag`' file in the root!`r`n"
                [string]$USBDriveErrMSG = -Join $USBDriveErrMSG, "The process has been broken. Please reformat the drive and start again from the Client Management script. Script Exiting!"
                Write-WSMLogMessage -Message $USBDriveErrMSG
                $null = Start-Process -FilePath 'Msg.exe' -ArgumentList '*','/server:LocalHost','/TIME:120',$USBDriveErrMSG
                $Return = 42007
            }
        } # if ($Config.ManualDest -eq 'USB'){
        elseif ($Config.ManualDest){
            $ManualDestination = $Config.ManualDest
            Write-WSMLogMessage -Message "A site override has been input. The script will deposit all logs and data to: $ManualDestination"
        } # elseif ($Config.ManualDest){
        else {
            # This is an option I mostly abandoned early one until the request for a USB option for devices that are offline. 
            Write-WSMLogMessage -Message "No Alternative location available"
        }
    } # else {

    # Build User Info custom object.
    $UserInfo = [PSCustomObject]@{
        LoggedOnUser = $LogonInfo
        LastLogonUserID = $LastLoggedOn
        LastUserFullname = $LastUserFullname
        TechnicianID = $TechID
        TechnicianName = $TechName
        PrimaryUserID = $PrimaryUserID
    } # $UserInfo = [PSCustomObject]@{

    # Look if primary user was populated, if not grab last user.
    if ($PrimaryUserID -in 'N/A',$null,''){
        $NameInPath = $LastLoggedOn
    }
    else {$NameInPath = $PrimaryUserID}

    if (!$ManualDestination){
        # Check the legalsite json to see where the preferred site maybe.
        # More needs to be done here to ensure a valid location has been selected.
        $LegalJson = Get-Content -Path $LegalSitesFile -ErrorAction Stop | ConvertFrom-Json

        if (  ($LegalJson | Where-Object {($_.Site -eq $Config.PreferredSite) -and ($_.Location -in $null,'') }) -or !($Config.PreferredSite) ){
            # try using the site the device is in to figure out where to send the data.
            if ( ($Destination = ($LegalJson | Where-Object { ($_.Site -match [regex]::Escape($ComputerTXT.Site)) -and ($_.Location -notin $null,'') }).Location ) ){
                Write-WSMLogMessage -Message "No valid preferred site pre-selected. Script has chosen: $Destination"
            }
            else {
                $Return = 42004
                $null = Start-Process -FilePath 'Msg.exe' -ArgumentList '*','/server:LocalHost','/TIME:120',"No valid site selected. legal discovery process cannot continue"
                Write-WSMLogMessage -Message "No valid site selected. legal discovery process cannot continue"
            }
        } # if ( ( ($Destination = ($LegalJson.$($Config.PreferredSite).Legal) ) -in $null,'' ) -or !($Config.PreferredSite) ){
        else {
            $Destination = ($LegalJson | Where-Object {($_.Site -eq $Config.PreferredSite) -and ($_.Location -in $null,'')}).Location
        }

        Write-WSMLogMessage -Message "Site Specific storage path located: $Destination"
    } # if (!$ManualDestination){
    else {
        # A very basic check to see if the path exists and is accessible
        if (Test-Path -Path $ManualDestination -errorAction SilentlyContinue){
            $Destination = $ManualDestination
            Write-wsmlogmessage -Message "Script will use this as the file destination: $Destination"
        }
    } # else {

        # Make sure we can actuall get to the destination
    if (!(Test-Path -Path $ManualDestination -errorAction SilentlyContinue)){
        Write-wsmlogmessage -Message "Error: Destination ($ManualDestination) is unreachable"
        $null = Start-Process -FilePath 'Msg.exe' -ArgumentList '*','/server:LocalHost','/TIME:120',"Error: Destination ($ManualDestination) is unreachable"
        $Return = 42007
    }

} # Begin {
Process {

    if ($Return -eq 0){

        # Give the target device a few seconds to initialize
        Start-sleep -seconds 30

        <#
        #region - Block inputs
            #
            #  Prevent anyone logging in and potentially locking files.
            #      There were issues with more advanced functions so I'm using this code from technet.
            #  This is not working across sessions similar to how the splashscreen was failing.
            #      I will leave the code in for now and revisit it at a later date.
            #
# Warning - Leave here-string with no left-hand margin or else either the here-string or the entire script will not work.
$signature = @"
[DllImport("user32.dll")]
public static extern bool BlockInput(bool fBlockIt);
"@
            try {
                $BlockInput = Add-Type -memberDefinition $signature -name Win32BlockInput -namespace Win32Functions -passThru -erroraction stop

                function Enable-BlockInput { $null = $BlockInput::BlockInput($true) }
                function Disable-BlockInput { $null = $BlockInput::BlockInput($false)}
                Enable-BlockInput
            }
            catch {
                Write-WSMLogMessage -Message "Failed to block user inputs. Script will continue..."
            }

        #endregion - Block inputs
        #>

        #region - Stop Processes\Services
            if (!$ParamInput){
                $StopProcs = $Config.StopProcs
                $StopServices = $Config.StopServices
            }

            # Stop any offending processes
            $null = Get-Process -Name $StopProcs -PipelineVariable $Proc -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $Null = Stop-process -name $Proc -force -ErrorAction Stop
                    Write-WSMLogMessage -message "Process $Proc has been stopped."
                }
                catch {
                    Write-WSMlogMessage -Message "Failed to stop process: $Proc. $($_.exception.Message)"
                }
            } # Get-Process -Name $StopProcs -PipelineVariable $Proc -ErrorAction SilentlyContinue | ForEach-Object {

            # Stop any offending services
            $PreServiceDisable = Get-Service -Name $StopServices -ErrorAction SilentlyContinue -PipelineVariable Svc | Foreach-object {
                # Build a custom object to report on the state
                [PSCustomObject]@{
                    Name = $Svc.Name
                    Status = $Svc.Status
                    StartMode = $Svc.StartType
                    DelayedStart = $((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$($Svc.Name)" -ErrorAction SilentlyContinue).DelayedAutostart)
                }

                # Stop the services
                try {
                    if ($Svc.Status -eq 'Running'){
                        $null = Stop-Service -name $Svc.Name -Force -ErrorAction Stop
                        Write-WSMLogMessage -Message "Set service $($Svc.Name) to: $($Svc.Status)"
                    }
                    else {
                        Write-WSMLogMessage -Message "Service $($Svc.Name) was already set to: $($Svc.Status)"
                    }
                }
                catch {
                    Write-WSMLogMessage -Message "Failed to stop service $($Svc.Name). $($_.exception.Message)"
                }

                # Disable the services
                try {
                    $null = Set-Service -name $Svc.Name -StartupType Disabled -ErrorAction Stop
                    Write-WSMLogMessage -Message "Service $($Svc.Name) has been disabled."
                }
                catch {
                    Write-WSMLogMessage -Message "Failed to disable service: $($Svc.Name). $($_.exception.Message)"
                }
            } # $PreServiceDisable = Get-Service -Name $Svc -ErrorAction SilentlyContinue | Foreach-object {
        #endregion - Stop Processes\Services

        #region - Create splash message object
            <#
                The pretty splashscreen isn't communicating with other sessions so I'm going full 'MVP' on it.
                I may replace this if I can circle back to it later and get the cross session coms working.
            #>

            try {
                # Suddenly having issues with this. So it has been swapped to the catch.
                $MsgSplat1 = @{
                    Title = 'Legal Discovery'
                    Message = 'This Device is currently running a legal discovery process and will reboot on completion. DO NOT LOG INTO THE DEVICE.'
                    ButtonSet = 0
                    Timeout = 21600
                    WaitResponse = $false
                }

                #Send
                $null = Send-WSMTSMessageBox @MsgSplat1
            }
            catch {
                $null = Start-Process -FilePath 'Msg.exe' -ArgumentList '*','/server:LocalHost','/TIME:21600','This Device is currently running a legal discovery process and will reboot on completion. DO NOT LOG INTO THE DEVICE!'
                #msg * /server:LocalHost /TIME:3600 'This Device is currently running a legal discovery process. DO NOT LOG INTO THE DEVICE.'
            }

        #endregion - Create splash message object

        #region - run the copy
            <#
                Find a random unused drive letter and assign it for the destination.
                This is important in case we run into long file paths that would otherwise cause issues.

                Recently discovered that if this fails simply checking for the
            #>
            
            if (!$USB){
                # Map the drive
                $MapDrive = (New-WSMRandomDriveLetter -Location $Destination -Persist)
                # Assign the drive to a static variable for slightly easier handling moving forward
                $DestDrive = $MapDrive.NewDrive
                Write-WSMLogMessage -Message "Destination `'$Destination`' mapped to `'$DestDrive`'"
            } # if (!$USB){
            else {
                $DestDrive = $ManualDestination -replace '\\',''
            }

            # Check for success
            if ( ($USB -eq $true) -or ( ($MapDrive.Result -match 'Success') -and ($MapDrive.NewDrive -notin '',$null) ) ){

                if (!$ParamInput){
                    $DirEx = $Config.DirEx
                    $FileEx = $Config.FileEx
                    $Drives = $Config.Drives
                } # if (!$ParamInput){

                # Designate the root copy directory
                $RootCopyDir = "$($env:COMPUTERNAME)_$($NameInPath)_$((Get-date).ToString('yyyyMMddHHmmss'))"
                $FinalLogLocation = "$Destination\$RootCopyDir"

                foreach ($Drive in $Drives){

                    # Filter out network drives and cleans the drive letter up
                    $DrvLetter = (Get-PSDrive -ErrorAction SilentlyContinue |
                        Where-Object {($_.Provider.name -eq 'Filesystem') -and ($_.DisplayRoot -notmatch '^\\') -and ($_.root -match [regex]::Escape($Drive))}).name

                    if ($DrvLetter){
                        # Check if we need to create dirs.
                        if ( !(Test-Path -Path "$Destination\$RootCopyDir\$DrvLetter" -ea silentlycontinue) ){
                            try {
                                $null = New-Item -Path "$Destination\$RootCopyDir\$DrvLetter" -ItemType "directory" -Force -ErrorAction Stop
                                Write-WSMLogMessage -Message "Successfully created $RootCopyDir : $($_.Exception.Message)"
                            }
                            catch {
                                #Bomb out....
                                #Disable-BlockInput

                                if (!$Silent){
                                    $hash.window.Dispatcher.Invoke("Normal",[action]{ $hash.window.close() })
                                    $Pwshell.EndInvoke($handle) | Out-Null
                                }
                                Write-WSMLogMessage -Message "Failed to create destinaton directory $("$Destination\$RootCopyDir\$DrvLetter"): $($_.Exception.Message)"
                                Invoke-WSMExitWithCode -ExitCode 42006
                            }
                        } # if ( !(Test-Path -Path "$Destination\$RootCopyDir\$DrvLetter" -ea silentlycontinue) ){

                        # Build the rest of the copy path
                        $FinalDest = "$DestDrive\$RootCopyDir\$DrvLetter"

                        # Build the log path location
                        $CopyLogLocation = "$DestDrive\$RootCopyDir"
                        Write-WSMLogMessage -message "About to begin inspection and file copy of drive $Drive"
                        Write-WSMLogMessage -message "Files will be move to: $FinalDest"
                        $CopyArguments = @{
                            Path = $Drive
                            Destination = $FinalDest
                            Recursive = $true
                            ExcludeDir = $DirEx
                            ExcludeFile = $FileEx
                            LogLocation = $($PSScriptRoot)
                            HashLogLocation = $CopyLogLocation
                            Logfile = "$Psscriptroot\WSMValidatedCopySingleItems_$((Get-date).ToString('yyyyMMddHHmmss')).log"
                            verbose = $true
                        }

                        # Log the copy paramters for referrence
                        $ParamMessage = "-Path $($CopyArguments.Path) -Destination $($CopyArguments.Destination) -Recursive $($CopyArguments.Recursive) -ExcludeDir $($CopyArguments.ExcludeDir) -ExcludeFile $($CopyArguments.ExcludeFile) -LogLocation $($CopyArguments.LogLocation) -Logfile $($CopyArguments.Logfile)"
                        Write-WSMLogMessage -Message "$Section Copy function parameters: $ParamMessage"
                        #$CopyReturn = Invoke-WSMValidatedCopy -Path $Drive -Destination $FinalDest -Recursive -ExcludeDir $DirEx -ExcludeFile $FileEx -LogLocation $CopyLogLocation -AdvancedExclusions -Logfile $LogFile -verbose
                        $CopyReturn = Invoke-WSMValidatedCopySingleItems @CopyArguments

                        # Keep the Returns separate for each DrvLetter
                        $CopyReturn | Out-File -FilePath "$CopyLogLocation\$($CompSys.Name)_$($DrvLetter)_CopyResult.txt" -Force
                    } # if ($DrvLetter){
                    else {
                        Write-WSMLogMessage -Message "Supplied Drive $Drive was not a locally mapped drive. No backup will take place on this drive."
                    }

                } # foreach ($Drive in $Config.Drives){
            } # if ($DestDrive = (New-WSMRandomDriveLetter -Location $Destination).NewDrive){
            else {
                Write-WSMLogMessage -Message "Failed to assign mapped network drive for the copy process"
            }
        #endregion - run the copy

        #region - Restart halted services
            $PreServiceDisable | ForEach-Object {
                $Svc = $_
                if ((Get-Service -Name $Svc.Name -ErrorAction SilentlyContinue).StartType -ne $Svc.StartMode){
                    try {
                        Get-Service -Name $Svc.Name -ErrorAction SilentlyContinue | Set-Service -StartupType $Svc.StartMode -ErrorAction Stop
                        Write-WSMLogMessage -Message "Successfully set service starttype back to $($Svc.StartMode)"
                    }
                    catch {
                        Write-WSMLogMessage -Message "Failed to set starttype for service $($Svc.Name) back to $($Svc.StartMode)."
                    }
                } # if ((Get-Service -Name $Svc.Name -ErrorAction SilentlyContinue).StartType -ne $Svc.StartType){

                if ($Svc.DelayedStart -notin $null,''){
                    # this assumes the value already exists and is ready for the data to be planted.
                    try {
                        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$($Svc.Name)" -Name 'DelayedAutostart' -Value $($Svc.DelayedStart) -Force -ErrorAction Stop
                        Write-WSMLogMessage -Message "Script set delayed starttype value for service $($Svc.Name) back to `'$($Svc.DelayedStart)`'."
                    }
                    catch {
                        Write-WSMLogMessage -Message "Failed to set delayed starttype value $($Svc.DelayedStart) for service $($Svc.Name)."
                    }
                } # if ($Svc.DelayedStart -notin $null,''){

                if (($Svc.Status -eq 'Running') -and ((get-Service -name $Svc.Name -ErrorAction SilentlyContinue).Status -ne 'Running')){

                    # Restart the service
                    try {
                        $null = Get-Service -Name $Svc.Name -ErrorAction Stop | Start-Service -ErrorAction Stop
                        Write-WSMLogMessage -Message "Successfully started service $($Svc.Name)"
                    }
                    catch {
                        Write-WSMLogMessage -Message "Failed to start service $($Svc.Name)"
                    }
                } # if (($Svc.Status -eq 'Running') -and ((get-Service -name $Svc.Name -ErrorAction SilentlyContinue).Status -ne 'Running')){
            } # $PreServiceDisable | ForEach-Object {
        #endregion - Restart halted services

        #region - Remove mapped drive
            if ( ($MapDrive.PreExisting0 -ne $true) -and ($MapDrive.NewDrive -notin '',$null) ){
                $BareLetter = $DestDrive  -replace ':','' -replace '\\',''
                try {
                    Remove-PSDrive -Name $BareLetter -Scope Global -Force -ErrorAction Stop
                    Write-WSMLogMessage -Message "Successfully removed mapped network drive."
                }
                catch {
                    Write-WSMLogMessage -Message "Failed to remove drive mapped for this session: $DestDrive. $($_.exception.message)"
                }
            }
        #endregion - Remove mapped drive

        Write-WSMLogMessage -Message "Finished primary discovery and copy functions."
            #Stop splashscreen
                <#
                    May Replace this if I can circle back to it later. The Splashscreen is not imperative.

                #>

        #region - Unblock inputs

            # enable access to the keyboard and mouse again
            #Disable-BlockInput
        #endregion - Unblock inputs

    } # if Return 0

} # Process {
end {
        #region - Write final reports

        #$SchTN = 'LegalDisc'
        #Write-WSMLogMessage -Message "About to remove Scheduled task $SchTN."
        #try {
        #    $DisableTask = Get-ScheduledTask -TaskName $SchTN -ErrorAction SilentlyContinue | Unregister-ScheduledTask -ErrorAction Stop -Confirm:$false
        #    Write-WSMLogMessage -Message "Scheduled task $SchTN has been unregistered. Result: $DisableTask"
        #}
        #catch {
        #    Write-WSMLogMessage -Message "Error: Could not remove scheduled task $SchTN : $($_.Exception.Message)"
        #}

        # remove scheduled task
        Remove-WSMScheduledTask -Taskname $SchTN -LogFile $LogFile

        Write-WSMLogMessage -Message "Script complete: $Return"
        if (Test-Path -Path $FinalLogLocation -ErrorAction SilentlyContinue){

            # Export the UserInfo in Json format - We may no longer require the old text format
            $UserInfo | ConvertTo-Json | Out-File -FilePath "$FinalLogLocation\UserInfo.JSON" -Force

            <#
                # Export the UserInfo file
                $UserInfo | Out-file -FilePath "$FinalLogLocation\UserInfo.txt" -Force

                # Export the Computer text file
                $ComputerTXT | Out-file -FilePath "$FinalLogLocation\Computer.txt" -Force
            #>

            # Export the Computer JSON file
            $ComputerTXT | ConvertTo-Json | Out-file -FilePath "$FinalLogLocation\Computer.JSON" -Force

            Write-WSMLogMessage -Message "Final confirmation files sent to `'$FinalLogLocation`'"

        } # if (Test-Path -Path "$FinalLogLocation" -ErrorAction SilentlyContinue){
        else {
            Write-WSMLogMessage -Message "Could not locate `'$FinalLogLocation`'"
        }

        #Currently not using NewUser
        $NewUserID = ''

        #Currently not using previous user ID either...
        $ComfrimationFile = [PSCustomObject]@{
            NewUserID = $NewUserID
            PreviousUserID = $LastLoggedOn
            PreviousUserName = $LastUserFullname
            LastLogonUserID = $LastLoggedOn
            ProfileUserIDs = (Get-ChildItem -Path 'c:\Users' -Directory -force) -Join '|'
            TechnicianID = $TechID
            TechnicianName = $TechName
            PrimaryUserID = $PrimaryUserID
            Company = $Company.Company
        }

        # Set confirmation location
        $ConfirmationFileLocation = "\\Whereever!!!"

        # Export the confirmation file
        try {
            $ComfrimationFile | Out-File -FilePath "$ConfirmationFileLocation\$RootCopyDir.txt" -Force -ErrorAction Stop
            Write-WSMLogMessage -Message "Successfully deposited confirmation log `'$RootCopyDir.txt`' in $ConfirmationFileLocation."
        }
        catch {
            $ComfrimationFile | Out-file -FilePath "$Destination\$RootCopyDir\$RootCopyDir.txt" -Force -ErrorAction SilentlyContinue # just forget about catching
            Write-WSMLogMessage -Message "Could not deposit confirmation file in standard location $ConfirmationFileLocation. File deposited here instead `'$Destination\$RootCopyDir`'"
        }

        #region - Final processes
            # Doesn't appear that it's possible to kill the previous msg box programatically, we'll have to wait for it to time out or for a user to click 'OK'.
            try {
                <#
                    $MsgSplat = @{
                        Title = 'Legal Discovery'
                        Message = 'The Data Collection process has completed, the device will no be rebooted'
                        ButtonSet = 0
                        Timeout = 300
                        WaitResponse = $true
                    }

                    #Send
                    $null = Send-WSMTSMessageBox @MsgSplat
                #>
                
            }
            catch {
                #$null = Start-Process -FilePath 'Msg.exe' -ArgumentList '*','/server:LocalHost','/TIME:300','/W','The Legal Discovery process has completed. This device is about to be rebooted!'
            }
            finally {Restart-Computer -ErrorAction SilentlyContinue}
            
            #msg * /server:LocalHost /TIME:300 /W 'The Legal Discovery process has completed. This device is about to be rebooted!'
            # After msg box has timed out, reboot...
            
        #endregion - Final processes
    #endregion - Write final reports
}
# SIG # Begin signature block
# MIIcigYJKoZIhvcNAQcCoIIcezCCHHcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU5cEodqQ8kvtBK+v3+37Ir6eu
# Q6Sgghe5MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# MRYEFA5hqzfNIgWYXACCMECp4SXdGSOJMA0GCSqGSIb3DQEBAQUABIIBAFVgfwFx
# wbY5CHViWzgz1eGyVrP+wXiG/ZgyQQUmK3e0qznwfKntzk3u1GfRa0wt5TV+Na7q
# 9rViIcRhrRYdevsHvJSbyg5DaUwQlzQPu06g2rnb3SeAoyHtYQbg7mxBDNCuM8Wx
# ulSb+lvRHzo9UOSOvgLuIxFnREJU61y9xYVMjnvGAsTkhQcbC9lNLid7gqBMe+jS
# D/+XTwDYED8pWezXdVW+FGCbPDkh1Ok1XQ3oVtvTvnBFKHQvZRy3ly5q1/AE1/ex
# Mg8sey5sVcHBOXkVZRmTQUK9+1BrV13vyxpyGcF0BjYyIYJaT3dbxdTIGe7VDXZu
# 5oi7+RUjr8FlzZChggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEC
# EAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIwMDIxMTIwMTkyMVowIwYJKoZIhvcN
# AQkEMRYEFGnin4lZpWgYow+tEQJ9R8rskwyBMA0GCSqGSIb3DQEBAQUABIIBAF5d
# wCmNv6EgsOxyQpumCet613eYRU7FwBdfK5EwgDs7JcV381BWC+lFTVB6mVhVfWjf
# 6VryAACtmiw19un9MlITsfjyGo0Y/c0JtOnjs6+qM7W+KNEXtU4rGsoqzye5ntnM
# ME/wvLXQ8TpinoJUcHECyDz+z+rC89lHig9jN8McTKNE3WfWiNKeXObFVV5wWQgQ
# UcnMGI95i1BfTpIkKdSy7ziZGuueEzPBye9VsKnYunq+xBQCXsYF3/Le27+SFrZO
# Nmicc+wTyhBjazuDYXdSfdGIeBBWHLXxWUyXnekAPzFr5sDfHO1XjyFc4j6HTcAq
# 6E4OjF3gOB0MlrfANPU=
# SIG # End signature block
