<#
    .Description
        Retrives Data from one location and moves it to the next while validating file hashes
    .Parameter ConfigPath
        Full path to the location JSON file containing operating configurations
    .Inputs
        Object (JSON File)
    .OutPuts
        PSCustomObject
    .Notes
        Exit codes used by this wrapper:
        42001 = Module load failure
        42002 = Site Location Json load failure
        42002 = Unable to contact final repository location

        ========== HISTORY ==========
        Author: Van Bogart, Kevin
        Created: 2019-11-27 12:05:24Z
#>
Param (
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [ValidateScript({(Get-item -path $_ -ErrorAction SilentlyContinue).Extension -eq '.JSON'})]
    [string]$LegalSitesFile,

    #This is for a UNC specifically because I needed it, but just needs to be a path if you want to change the validation...
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [ValidateScript({[bool]([System.Uri]$_).IsUnc})]
    [string]$Destination,

    #LogFile - File for log messages.
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$LogFile = 'StorageProcess.log',

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [int]$LogSize = 5
)
begin {
    $Return = 0 
    #region - Load Modules

        #Yeah, you need to figure out what was in there and create it if needed. 

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
                $Host.SetShouldExit(1)
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
                #$Host.SetShouldExit(32001)
                $Return = 42001
            }

        } # if (!(Get-Module -Name BSCWSMLDMod -ea SilentlyContinue)){
    #endregion - Load Modules

    if ($LogFile -eq 'StorageProcess.log'){
        $LogFile = "$PSScriptRoot\$LogFile"
        Write-verbose -message "LogFile: $LogFile"
    }

    # Make some defaults here
    $PSDefaultParameterValues["Write-WSMLogMessage:LogFile"] = $LogFile
    $PSDefaultParameterValues["Write-WSMLogMessage:LogSize"] = $LogSize

    # Bring in the legal sites
    try {
        $LegalSites = Get-Content -Path $LegalSitesFile -ErrorAction Stop | ConvertFrom-Json
    }
    catch {$Return = 42002}

    if ( !(Test-Path -Path  $Destination -ErrorAction SilentlyContinue) ){
        Write-WSMLogMessage -Message "Unable to contact final locaiton: $Destination. Script exiting!"
        $Return = 42002
    }

} # begin {
process {

    if ($Return -eq 0){
        #region - Map destination drive
            # We don't need to do this if the path is mapped on the server.
            $MapDrive = (New-WSMRandomDriveLetter -Location $Destination -Persist -ExcludedDriveLetters 'a','b','c')

            # Check for success
            if ( ($MapDrive.Result -match 'Success') -and ($MapDrive.NewDrive -notin '',$null) ){

                # Assign the drive to a static variable for slightly easier handling moving forward.
                $DestDrive = $MapDrive.NewDrive
                Write-WSMLogMessage -Message "Destination `'$Destination`' mapped to `'$DestDrive`'"
            }
        #endregion - Map destination drive

        foreach ($Site in ($LegalSites | Where-Object {$_.location -notin ''} | Select-Object -Property Location -Unique)){
            Get-ChildItem -Path $Site.location -Directory -ErrorAction SilentlyContinue -PipelineVariable BackUp |
                <#
                    Having the json version of the file indicates it's the new process that was used and that
                        the process completed as it's one of the last items generated by the client script
                #>
                Where-Object { (Get-Item -path "$($BackUp.FullName)\UserInfo.JSON" -ErrorAction SilentlyContinue).Exists -eq $true } |
                    ForEach-Object {

                        Write-WSMLogMessage -Message "Backup located $($BackUp.FullName)."

                        if ($CompFile = Get-Content -path "$($BackUp.FullName)\Computer.JSON" -ErrorAction SilentlyContinue){
                            #This should be populated..
                            $ComputerName = $CompFile.ComputerName
                        }

                        if (!$ComputerName -and ($Backup.name -match '_')){
                            $ComputerName = ($Backup.name).split('_') | Select-Object -First 1
                        }
                        else {
                            $ComputerName = "UnknownDevice-$(((New-Guid).guid).split('-') | Select-Object -First 1)"
                        }

                        # The script is now in a directory that presumably has at least one directory to work on
                        #region - Map Source drive
                            # We don't need to do this if the path is mapped on the server.
                            $MapSrcDrive = (New-WSMRandomDriveLetter -Location $Site.location -Persist -ExcludedDriveLetters 'a','b','c')

                            # Check for success
                            if ( ($MapSrcDrive.Result -match 'Success') -and ($MapSrcDrive.NewDrive -notin '',$null) ){

                                # Assign the drive to a static variable for slightly easier handling moving forward.
                                $SourceDrive = $MapSrcDrive.NewDrive
                                $RoboSource = "$($SourceDrive)\$($BackUp.Name)"
                                Write-WSMLogMessage -Message "Initial unresolve Robosource Location: $Robosource"
                                Write-WSMLogMessage -Message "Legal Site source folder `'$($Site.Location)`' mapped to `'$SourceDrive`'"
                            }
                        #endregion - Map Source drive

                        #region - copy the files

                        # Calculate an appropriate amount of time for the files transfer before terminating the robocopy process.
                        try {
                            [int]$TotalSize = "{0}" -f [math]::Round((Get-ChildItem -Path $RoboSource -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property 'Length' -Sum -ErrorAction Stop).Sum / 1MB)
                            if ($TotalSize -le 420024){
                                $FileCopyTime = 600 # 10 min
                                Write-WSMLogMessage -Message "The target dir is 1 GB or less. 10 minutes will be allocated for the filecopy to be completed."
                            }
                            else {
                                $FileCopyTime = ([math]::Round(($TotalSize/420024)*600)) # 10 min per GB
                                Write-WSMLogMessage -Message "The target dir is $TotalSize GB. 10 minutes for each GB will be allocated for the filecopy to be completed. Total $FileCopyTime."
                            }
                        }
                        catch {
                            Write-WSMLogMessage -Message "Failed to get filesize for $RoboSource."
                        }

                        # And this next code block was evading me for hours... MVP right here...
                        $ResolvedDestination = $($RoboSource.Replace($SourceDrive,$DestDrive))
                        Write-WSMLogMessage -Message "Resolved Robosource destination is: $ResolvedDestination"
                        $InitialRobocopyLog = "/LOG:$DestDrive\RobocopyLog_$($ComputerName)_$(Get-Date -Format 'yyyyMMddHHmmss').log"
                        $RoboArgs = "$RoboSource $ResolvedDestination *.* /COPY:DAT /Move /e /z /MT:10 /xjd /np /r:5 /w:5 /xa:S /A-:SH $InitialRobocopyLog "
                        Write-WSMLogMessage -Message "Robocopy command: Robocopy $RoboArgs"

                        # Copy the files
                        $Robocopy = Start-Process -FilePath 'Robocopy.exe' -ArgumentList $RoboArgs -PassThru
                        $Robocopy | Wait-Process -Timeout $FileCopyTime -ErrorAction SilentlyContinue

                        #endregion - copy the files

                        #region - Obtain hashes of moved items
                            # Only get information on the top level directories and nothing they contain.
                            Get-ChildItem -Path $ResolvedDestination -Directory -PipelineVariable BUDrive -ErrorAction SilentlyContinue | ForEach-Object {
                                try {
                                    [array]$AllHash = @()

                                    # Splat out the get child items parameters.
                                    $GCIParams = @{
                                        Path = "$($BUDrive.FullName)"
                                        Recurse = $true
                                        ErrorAction = 'SilentlyContinue'
                                        Force = $true
                                        PipelineVariable = 'FSO'
                                    }

                                    # Inspect each item in the targeted directory and begin bulding an array of objects to repsresent the files being handled.
                                    $AllHash += Get-ChildItem @GCIParams | Foreach-object {

                                        try {
                                            Get-FileHash -Path $FSO.Fullname -Algorithm MD5 -ErrorAction Stop |
                                                ForEach-Object {
                                                    [PSCustomObject]@{
                                                        FinalHash = $_.Hash
                                                        ArchivedFile = $_.path
                                                        FileName = $FSO.Name
                                                        FileSize = $FSO.Length
                                                    } # $AllHash += get-ChildItem -Path "$Path" -Recurse  | .....
                                            } # $AllHash += get-ChildItem -Path "$Path" -Recurse | Get-FileHash -Algorithm MD5 | ForEach-Object {
                                        }
                                        catch [Microsoft.PowerShell.Commands.WriteErrorException]{
                                            [PSCustomObject]@{
                                                FinalHash = "No Hash Available - File is being used by another process: $($FSO.name)"
                                                ArchivedFile = $FSO.Fullname
                                                FileName = $FSO.Name
                                                FileSize = $FSO.Length
                                            } # [PSCustomObject]@{
                                        }
                                        catch {
                                            [PSCustomObject]@{
                                                FinalHash = $_.exception.message
                                                ArchivedFile = $FSO.Fullname
                                                FileName = $FSO.Name
                                                FileSize = $FSO.Length
                                            } # [PSCustomObject]@{
                                        }

                                    } # $AllHash += Get-ChildItem @GCIParams | Foreach-object {
                                }
                                catch {
                                    Write-WSMLogMessage -Message "Failure: File $($FSO.FullName). Error: $($_.exception.GetType().fullname)"
                                }

                                $AllHash | Convertto-JSON | Out-File -FilePath "$ResolvedDestination\$($ComputerName)_FinalHash_$($BUDrive.Name)_$(Get-Date -Format 'yyyyMMddHHmmss').json"
                            } # Get-ChildItem -Path $($RoboSource.Replace($SourceDrive,$DestDrive)) -Directory -ErrorAction SilentlyContinue -PipelineVariable BUDrive | ForEach-Object {

                        #endregion - get hashes of moved items

                        <#
                            Possibly write something to run a verification check. Need tests first to generate the logs we're looking for.
                        #>

                        #region - Remove mapped drive
                            if ( ($MapSrcDrive.PreExisting -ne $true) -and ($MapSrcDrive.NewDrive -notin '',$null) ){
                                $BareSrcLetter = $SourceDrive  -replace ':','' -replace '\\',''
                                try {
                                    Remove-PSDrive -Name $BareSrcLetter -Scope Global -Force -ErrorAction Stop
                                    Write-WSMLogMessage -Message "Successfully removed mapped network drive."
                                }
                                catch {
                                    Write-WSMLogMessage -Message "Failed to remove drive mapped for this session: $SourceDrive. $($_.exception.message)"
                                }
                                finally {
                                    #Clear certain variables
                                    Remove-Variable -Name AllHash,BareSrcLetter,SourceDrive,RoboSource,RoboArgs,ResolvedDestination,MapSrcDrive,BackUp -ErrorAction SilentlyContinue
                                }
                            }
                        #endregion - Remove mapped drive

                        #Less than happy with the pipeline behavior.

                    } # ForEach-Object {
        } # foreach ($Site in ($LegalSites | Where-Object {$_.location -notin ''})){

        #region - Remove mapped drive
            if ( ($MapDrive.PreExisting -ne $true) -and ($MapDrive.NewDrive -notin '',$null) ){
                $BareLetter = $DestDrive  -replace ':','' -replace '\\',''
                try {
                    Remove-PSDrive -Name $BareLetter -Scope Global -Force -ErrorAction Stop
                    Write-WSMLogMessage -Message "Successfully removed mapped network drive."
                }
                catch {
                    Write-WSMLogMessage -Message "Failed to remove drive mapped for this session: $DestDrive. $($_.exception.message)"
                }
                finally {
                    #Clear certain variables
                    Remove-Variable -Name AllHash,DestDrive,BareLetter,RoboSource,RoboArgs,ResolvedDestination,MapSrcDrive,BackUp -ErrorAction SilentlyContinue
                }
            }
        #endregion - Remove mapped drive

    } # if ($Return = 0){
}
end {
    Write-WSMLogMessage -Message "Script complete. Result: $Return"
    $Return
}
