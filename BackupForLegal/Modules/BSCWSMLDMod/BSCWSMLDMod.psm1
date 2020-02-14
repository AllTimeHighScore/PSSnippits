#=================================================================
#region - Description
#=================================================================
#Author: Kevin Van Bogart
#Created: 2019-10-23 12:30:18Z
#
#Version 1.0.0.0
#=================================================================
#endregion - Description
#=================================================================
#=================================================================
#region - Define and Export Module Variables
#=================================================================
# This is the WSM Icon
Set-Variable -Name iconBase64 -Option ReadOnly -Value 'AAABAAEAICAQAgAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAMDAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP//AAD///8AAAAAAAAGZ3iHdmAAAAAAAAAAAAAGeP
        ////+HYAAAAAAAAAAGj//4h3iP//hgAAAAAAAAaP+Ih3d3d3j/9wAAAAAAB/+HeGZmZmd4iP92AAAAAH/4eHhmZmZnhneP9wAAAAb/h3ZHiIiHeHZnd/9wAABo+HhmZ3ZmZ4aHZneP9gAAf/d2Z4dmZmhmaHZnePgABv+HZnhniIiHhmaHZof/YAf/hkeGR3d4Znhm
        Z2Z3j3Bo93iIZ3hmeHZohmdmh3+Gb/aP/4//eP/3aPiP+P9v9m/2/////4f//3j4j/j/aPd/hv+P+I+Gf/+P+I/3/2j3f4b/f/iPh4//j/iP9/9o+H+G/3/4j/iP+Gj4j/f/aPh/hv9/+I/4j/do+I/4/2j3b/b/j/iPh///eP////9o92/3/4/4j4Zv/4f/+P/4b/
        YI+HhmdmiGZodmh2aHhn+GB/93ZodmiGaIeIRodGiPcAaPiGZodmiId3iHh2Z4/2AAf/d2ZodGhkZGiHZnePgAAGj4eGZoeHd3d4Zmh4+GAAAG/4d2Z4Z4iIiGZ3j/YAAAAG/4d2hmZmZmh3eP9wAAAAAG//iHd2ZmZ4d//3AAAAAAAGj/+Hd3d3iI/4YAAAAAAAAG
        eP//iIj//4dgAAAAAAAAAAZ4/////4dgAAAAAAAAAAAAAGZ3d2ZgAAAAAAD/4Af//4AB//4AAH/8AAA/+AAAD/AAAA/gAAAHwAAAA8AAAAOAAAABgAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAACAAAABgAAAAcAAAAPAAAAD4AAAB/
        AAAA/4AAAf/AAAP/4AAH//gAH///AH/w=='

        #Do not export this
#=================================================================
#endregion - Define and Export Module Variables
#=================================================================
#=================================================================
#region - Define PUBLIC Advanced functions
#=================================================================

function New-WSMRandomDriveLetter {
    <#
        .SYNOPSIS
            Assigns a random unused drive letter.
        .DESCRIPTION
            Assigns a random unused drive letter that should at least be in the scope of the current script.
            If the drive letter is needed in a global scope this function can be used to just genreate an unused drive letter.
        .PARAMETER Location
            Path to the source to be mapped
        .PARAMETER ExcludedDriveLetters
            manually specified drive letters not to assign
        .PARAMETER Logfile
            Logfile if one is specified.
        .EXAMPLE
            New-WSMRandomDriveLetter -Location '\\stpnas08\legalinfo$'
        .INPUTS
            System.String
        .OUTPUTS
            PSCustomObject
        .Notes
            This is basically pulled from a block I wrote for the CM cmdlets

            Last Updated:
    
            ========== HISTORY ==========
            Author: Kevin Van Bogart
            Created: 2019-10-25 09:18:47Z
    #>
    Param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [ValidateScript({[bool]([System.Uri]$_).IsUnc})]
        $Location,

        [Parameter(ValueFromPipeline=$true)]
        [switch]$Persist,

        [parameter()]
        [ValidateNotNullOrEmpty()]
        [String[]]$ExcludedDriveLetters,

        #File to log actions
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$LogFile

    )
    Begin {

    }
    Process {
        #Don't carry these over as they could technically differ between DTs.
        $DriveAssigned = $false
        $PreExisting = $false
        if ( !($AssignTempDrive = (Get-PSProvider).Drives.where({ $_.DisplayRoot -match [regex]::escape($Location) }) ) ){

            #Assign random unused drive letter for space inspection
            $Alphabet = 'abcdefghijklmnopqrstuvwxyz'.ToUpper().ToCharArray()

            #Remove any drives that might cause an issue
            $ExcludedDriveLetters | ForEach-Object {$Alphabet -replace $_,''}

            $Drives = (Get-PSProvider).where({$_.name -eq 'FileSystem'}).drives.name
            #Split out for ease of reading
            $Alphabet = $Alphabet | Where-Object {$Drives -notcontains $_}
            [string]$TempDriveLetter = Get-Random -InputObject $Alphabet 

            try {
                #Assign only leading path so this doesn't get out of control
                #Scope one should at least keep the drive letter in the scope of the parent script.

                $PSDriveParams = @{
                    Name = $TempDriveLetter
                    Root = $Location
                    PSProvider = 'FileSystem'
                    scope = 'Global'
                    ErrorAction = 'Stop'
                }

                if ($Persist){$PSDriveParams.Add('Persist',$true)}

                $AssignTempDrive = New-PSDrive @PSDriveParams

                $TempDrive = -Join ($AssignTempDrive,':')
                if ($LogFile){Write-WSMLogMessage -Message "This root drive will be used: $($TempDrive.Name)" -LogFile $LogFile}
                else {Write-Verbose -message "This root drive will be used: $($TempDrive.Name)"}
                $DriveAssigned = 'Success'
            }
            catch {
                if ($LogFile){Write-WSMLogMessage -Message "Failed to create drive letter to gold location. | Error: $($_.Exception.Message)" -LogFile $LogFile}
                else {Write-Verbose -message "Failed to create drive letter to gold location. | Error: $($_.Exception.Message)"}
                $DriveAssigned = 'Failed'
            }
        } # if ( !($AssignTempDrive = (Get-PSProvider).Drives.where({ $_.DisplayRoot -match ($Location).replace('\','\\') }) ) ){
        else {
            #Use the existing drive for the next part
            $TempDrive = -Join ($AssignTempDrive.name,':')
            $DriveAssigned = 'Success'
            $PreExisting = $true
            if ($LogFile){Write-WSMLogMessage -Message "This root drive will be used: $($TempDrive)" -LogFile $LogFile}
            else {Write-Verbose -message "This root drive will be used: $($TempDrive)"}
        }

        #The return
        [PSCustomObject]@{
            Result = $DriveAssigned
            NewDrive = $TempDrive
            PreExisting = $PreExisting
        }
    }

} # function New-WSMRandomDriveLetter {

function Invoke-WSMValidatedCopySingleItems {
    <#
        .SYNOPSIS
            A copy function, you can trust (It's Slow 'as all get out')
        .DESCRIPTION
            Copy items and validate their MD5 hashes in the process with a before and after check
        .PARAMETER Path
            Path to the source files to be copied
        .PARAMETER Destination
            Path to where file are to be copied
        .PARAMETER Recursive
            Copies subfolders, regardless if they are empty
        .PARAMETER ExcludeDir
            Directory to exclude or comma delimited list of directories that need to be excluded.
        .PARAMETER ExcludeFile
            Excludes files that match the specified names or paths. Note that FileName can include wildcard characters (*) no (?) This is because I'm using Regex.
            Comma delimited lists are also valid.
        .PARAMETER Logging
            Sets the level of logging done by the function
            Advanced (default):
                All Robocopy logs - Example: COMPUTERNAME_FileCopyList_yyyyMMddHHmmss.log
                                             COMPUTERNAME_RetryCopy_yyyyMMddHHmmss.log
                A log of any files that failed to copy properly - Example: COMPUTERNAME_FailedFileCopyList_yyyyMMddHHmmss.log
            Basic:
                All Robocopy logs - Example: COMPUTERNAME_FileCopyList_yyyyMMddHHmmss.log
                                             COMPUTERNAME_RetryCopy_yyyyMMddHHmmss.log
                A JSON of the file hashes - Example: COMPUTERNAME_FullHashList_yyyyMMddHHmmss.json (txt file containing message if hash check fails.)
                    This JSON has the most useful information
            Disabled:
                No Logging
    
            Note: Failure and Retry logs only exist when there are objects to populate them.
        .PARAMETER LogLocation
            Location logs are to be created
            Note: Default is The parent folder of the destination path.
        .PARAMETER HashLogLocation
            Location the full hash list needs to be deposited.
        .PARAMETER LogFile
            The file to log items to
        .EXAMPLE
            $FileEx = @("*.ost","*.sys","*.AppX","*.dat","*.dmp","*.exe","*.dll","*.mui","*.svg","UPPS.bin","BuildInfo.ini")
            
            $DirEx = @("OneDrive","OneDrive - Boston Scientific","C:\OneDriveTemp","C:\Windows","C:\MSOCache","C:\Quarantine","C:\Program Files", "C:\Program Files (x86)","C:\Oracle","C:\System Volume Information","C:\Config.Msi","C:\ProgramData")
            
            $Dest = '\\stpnas05\dsm\Public\SuperGuy\TestBackup\TestDest\'
            $RunLog = '\\stpnas05\dsm\Public\Restricted\VanBogK\Test_SingleCopy.log'
            Measure-Command -Expression {
                Invoke-WSMValidatedCopySingleItems -Path 'C:\' -Destination $Dest -Recursive -ExcludeDir $DirEx -ExcludeFile $FileEx -LogLocation '\\stpnas05\dsm\Public\Restricted\VanBogK\TestBackup' -LogFile $RunLog
            }
        .INPUTS
            System.String
        .OUTPUTS
            System.Object
        .NOTES
            Last Updated:
    
            ========== HISTORY ==========
            Author: Kevin Van Bogart
            Created: 2019-09-30 15:08:43Z
    #>
    Param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [ValidateScript({(Test-Path -Path $_ -PathType Container -ErrorAction SilentlyContinue) -or (Test-Path -Path $_ -PathType Leaf -ErrorAction SilentlyContinue)})]
        [string]$Path,

        [Parameter(Mandatory=$true,
        ValueFromPipeline=$false)]
        [ValidateScript({(Test-Path -Path $_ -PathType Container -ErrorAction SilentlyContinue) -or (Test-Path -Path $_ -PathType Container -IsValid -ErrorAction SilentlyContinue)})]
        [string]$Destination,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$Recursive,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({ ($_ -notmatch "$([regex]::Escape('*'))|$([regex]::Escape('?'))") -and (Test-Path -path $_ -IsValid -ErrorAction SilentlyContinue) })]
        [string[]]$ExcludeDir,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]]$ExcludeFile,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Basic', 'Advanced', 'Disabled')]
        [string]$Logging = 'Basic',

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$LogLocation = (Split-Path -Path $Destination -Parent),

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$HashLogLocation = (Split-Path -Path $Destination -Parent),

        #File to log actions
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [String]$LogFile

    )
    Begin {
        #Create repository for error messages
        if ($Recursive){$Recurse = $true}
        else {$Recurse = $false}

        $LogPath = $Path -replace '\\','' -replace ':',''
        #Check the source and build copy arguments

        #region - Prep any exclude paramters

                #region - Robocopy exclusions (Not being used in this version)
                    #Prep the exclusions so they can be used by regex
                    $ExcludePRM = ''

                    if ($ExcludeFile){
                        $EF = $ExcludeFile.ForEach({ "`"$_`""}) # Join the items in a format usable by robocopy
                        $ExcludePRM = " /xf $([string]$EF)"
                    }
                
                    if ($ExcludeDir){
                        $RoboDirEx = $ExcludeDir.ForEach({ "`"$_`""}) # Join the items 
                        if ($ExcludePRM -ne ''){
                            $ExcludePRM = "$ExcludePRM /xd $([string]$RoboDirEx)"
                        }
                        else {$ExcludePRM = "/xd $([string]$RoboDirEx)"}
                    } # if ($ExcludeDir){
                    
                    $ExcludeDir += $ExcludeFile.Where({Test-Path -path $_ -IsValid -ErrorAction SilentlyContinue}) # Ensure path is valid
                    $ExcludeFile = $ExcludeFile.Where({!(Test-Path -path $_ -IsValid -ErrorAction SilentlyContinue)}) # This should be files and not paths
                #endregion - Robocopy exclusions (Not being used in this version)

            # The files and dirs are almost the same, except the match section doesn't work well if there're not separated.
            # Clean the file entries for a search
            Foreach ($Ex in '?','*'){
                if ($ExcludeFile.count -gt 1){
                    $RegexFileExclusions = $ExcludeFile.ForEach({
                            if ($_ -match [regex]::Escape($Ex) ){
                                    [regex]::Escape($_.replace($Ex,''))
                            }
                            else {
                                [Regex]::Escape($_)
                            }
                    }) -Join '|'
                }
                elseif ($ExcludeFile.count -eq 1){
                    if ($ExcludeFile -match [regex]::Escape($Ex) ){
                        $RegexFileExclusions = [regex]::Escape($ExcludeFile.replace($Ex,''))
                    }
                    else {
                        $RegexFileExclusions = [Regex]::Escape($ExcludeFile)
                    }
                }
            } # Foreach ($Ex in '?','*'){

            # Inspect the Dirs
            if ($ExcludeDir.count -gt 1){
                $RegexDirExclusions = $($ExcludeDir.ForEach({[Regex]::Escape($_)})) -Join '|'
            }
            elseif ($ExcludeDir -eq 1){
                $RegexDirExclusions = [Regex]::Escape($ExcludeDir)
            }

            if ($LogFile){Write-WSMLogMessage -Message "$Section Robocopy Dir Exclusions: $ExcludeDir" -LogFile $LogFile}
            else {Write-verbose -message "$Section Robocopy Dir Exclusions: $ExcludeDir"}
            if ($LogFile){Write-WSMLogMessage -Message "$Section Robocopy File Exclusions: $ExcludeFile" -LogFile $LogFile}
            else {Write-verbose -message "$Section Robocopy File Exclusions: $ExcludeFile"}

            if ($LogFile){Write-WSMLogMessage -Message "$Section Regex Dir Exclusions: $RegexDirExclusions" -LogFile $LogFile}
            else {Write-verbose -message "$Section Regex Dir Exclusions: $RegexDirExclusions"}
            if ($LogFile){Write-WSMLogMessage -Message "$Section Regex File Exclusions: $RegexFileExclusions" -LogFile $LogFile}
            else {Write-verbose -message "$Section Regex File Exclusions: $RegexFileExclusions"}
        #endregion - Prep any exclude paramters

        if ($Logging -eq 'Disabled'){
            $InitialRobocopyLog = ''
            if ($LogFile){Write-WSMLogMessage -Message "$env:COMPUTERNAME: All logging was disabled in command line. No logs will be created for Robocopy in this session." -LogFile $LogFile}
            else {Write-verbose -message "$env:COMPUTERNAME: All logging was disabled in command line. No logs will be created for Robocopy in this session."}
        }
        else {
            $InitialRobocopyLog = "/LOG:$LogLocation\$($env:COMPUTERNAME)_FileCopyList_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').log"
        }

        # Verify the source location is valid.
        try {

            if (Test-Path -Path $Path -PathType leaf -ErrorAction stop){
                $ParentDir = Split-Path -Path $Path -Parent

                # Report state
                if ($LogFile){Write-WSMLogMessage -Message "$Section Source location exists and can be accessed: $ParentDir" -LogFile $LogFile}
                else {Write-verbose -message "$Section Source location exists and can be accessed: $ParentDir"}
            }
            # Shouldn't be anything other than a path after being filtered by test-path type check
            else {
                Write-Verbose -Message "Source location exists and can be accessed: $Path"
            }
        } # try {
            catch [System.UnauthorizedAccessException]{
                if ($LogFile){Write-WSMLogMessage -Message "$Section Source location may exist but access to it has been denied: $Path" -LogFile $LogFile}
                else {write-verbose -Message "$Section Source location may exist but access to it has been denied: $Path"}
            }
            catch {
                if ($LogFile){Write-WSMLogMessage -Message "$Section Could not access source location for unknown reason :$($_.exception.gettype().fullname) : $Path" -LogFile $LogFile}
                else {write-verbose -Message "$Section Could not access source location for unknown reason :$($_.exception.gettype().fullname) : $Path"}
            }

        <#
            Check the destination - may attempt to read acls on directory.
            This should also take care of the logfile directory...In the default location.
            Perhaps changing this to a custom object would help the output make more sense
        #>
        ForEach ($Dest in ($LogLocation, $Destination)){
            try {
                if (Test-Path -Path $Dest -PathType Container -ErrorAction Stop){
                    if ($LogFile){Write-WSMLogMessage -Message "$Section Destination path exists and is accessible: $Dest" -LogFile $LogFile}
                    else {Write-verbose -message "$Section Destination path exists and is accessible: $Dest"}
                }
                else {
                    New-Item -Path $Dest -ItemType Directory -force -ErrorAction Stop | Out-Null
                    if ($LogFile){Write-WSMLogMessage -Message "$Section Successfully created directory: $Dest" -LogFile $LogFile}
                    else {Write-verbose -message "$Section Successfully created directory: $Dest"}
                }
                $ValidDest = $true
            }
            catch [System.UnauthorizedAccessException] {
                if ($LogFile){Write-WSMLogMessage -Message "$Section Location may exist but access to it has been denied: $Dest" -LogFile $LogFile}
                else {Write-verbose -message "$Section Location may exist but access to it has been denied: $Dest"}
                $ValidDest = $false
                break
            }
            catch {
                if ($LogFile){Write-WSMLogMessage -Message "$Section Could not access or write to Location for unknown reason :$($_.exception.gettype().fullname) : $Dest" -LogFile $LogFile}
                else {Write-verbose -message "$Section Could not access or write to Location for unknown reason :$($_.exception.gettype().fullname) : $Dest"}
                $ValidDest = $false
                break
            }
        } # ForEach ($Dest in ($LogLocation, $Destination)){
    }
    Process {

        [array]$Failures = @()
        [array]$AllHash = @()

        if ($ValidDest -eq $true){

            # System files only (-|d){1}(-|a){1}(-|r){1}(-|h){1}(s){1}(-|l){1}
            # Symbolic links (d){1}(a){1}(r){1}(-){2}(l){1}
            [regex]$AttributeEx = '(d){1}(a){1}(r){1}(-){2}(l){1}|(-|d){1}(-|a){1}(-|r){1}(-|h){1}(s){1}(-|l){1}'
            [regex]$DirAttributeEx = '(d){1}(-|a){1}(-|r){1}(-|h){1}(-|s){1}(-|l){1}'

            try {
                # My tiny mind can't find another way around the an issue where the object will always -notmatch null data 
                if ($RegexDirExclusions -in '',$null){$RegexDirExclusions = 'THIS/-DIRECTORY/-DOES/-NOT/-EXIST'}
                if ($RegexFileExclusions -in '',$null){$RegexDirExclusions = '/.THERISNOWAYTHISISAREALEXTENSION'}

                $InitialRobocopyLog = "/LOG+:$LogLocation\$($env:COMPUTERNAME)_FileCopyList_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').log"

                # Filtering out the system and reparse point attributes will prevent copying locked system files, Symbolic links, and directories acting as junction points
                if ( ($Path -match '^([a-zA-Z]){1}(:){1}$')){
                    $GCIPath = "$Path\"
                    if ($LogFile){Write-WSMLogMessage -Message "$Section The source path supplied may require a trailing backslash: $path. The value is now $GCIPath." -LogFile $LogFile}
                    else {Write-verbose -message "$Section The source path supplied does may require a trailing backslash: $path. The value is now $GCIPath."}
                }
                elseif (Test-Path -path $Path -ErrorAction SilentlyContinue){
                    $GCIPath = $Path
                }
                else {
                    if ($LogFile){Write-WSMLogMessage -Message "$Section The source path supplied does not appear to be valid: $path. This function is about to fail." -LogFile $LogFile}
                    else {Write-verbose -message "$Section The source path supplied does not appear to be valid: $path. This function is about to fail."}
                }

                $GCIParams = @{
                    Path = "$GCIPath"
                    Recurse = $Recurse
                    ErrorAction = 'SilentlyContinue'
                    Force = $true
                    PipelineVariable = 'FSO'
                }

                $ParamMessage = "-Path $($GCIParams.Path) -Recurse $($GCIParams.Recurse) -ErrorAction $($GCIParams.ErrorAction) -Force $($GCIParams.Force) -PipelineVariable $($GCIParams.PipelineVariable)"
                if ($LogFile){Write-WSMLogMessage -Message "$Section Sending get-childItem these Parameters: $ParamMessage" -LogFile $LogFile}
                else {Write-verbose -message "$Section Sending get-childItem these Parameters: $ParamMessage"}

                $AllHash += Get-ChildItem @GCIParams | Where-Object {($FSO.Fullname -notmatch $RegexDirExclusions) -and ($FSO.name -notmatch $RegexFileExclusions)} | 
                        Foreach-object {

                            if ($FSO.mode -match $AttributeEx){
                                <#
                                    Get-ChilItem and the Where-Object filters were ignoring any attempt to filter by attributes conventionally
                                    To get around this a variable was built in-line so that symbolic links are not followed.
                                #>
                                $RegexDirExclusions = "$RegexDirExclusions|$([Regex]::Escape($FSO.Fullname))"
                            }
                            elseif ($FSO.mode -notmatch $DirAttributeEx){
                                $SrcHashObj = ''
                                try {
                                    $SrcHashObj = Get-FileHash -Path $FSO.Fullname -Algorithm MD5 -ErrorAction Stop
                                    #Write-Verbose -Message "First File $($FSO.Fullname)"
                                    #Write-Verbose -Message "First Path $($SrchashObj.Path)"
                                }
                                catch [Microsoft.PowerShell.Commands.WriteErrorException]{
                                    $SrcHashObj = [PSCustomObject]@{
                                        Hash = "No Hash Available - File is being used by another process: $($FSO.name)"
                                        path = $FSO.Fullname
                                    } # [PSCustomObject]@{
                                }
                                catch {
                                    $SrcHashObj = [PSCustomObject]@{
                                        Hash = $_.exception.message
                                        Path = $FSO.Fullname
                                    } # [PSCustomObject]@{
                                }

                                #region - Initial copy section
                                    #Robocopy item
                                    if ($FileCopyTime = [MATH]::Round((($Snagit.length)/1MB)/50) -lt 1){$FileCopyTime = 1}

                                    If ($Destination -notmatch "\\$"){
                                        $Destination = "$Destination\"
                                    }

                                    $RoboArgs = "$($FSO.DirectoryName) $(($FSO.DirectoryName).Replace($Path,$Destination)) $($FSO.Name) /COPY:DAT /z /MT:10 /xjd /njh /njs /np /r:5 /w:5 /xa:S /A-:SH $InitialRobocopyLog "

                                    # Copy source to remote destination
                                    $InitialCopy = Start-Process -FilePath 'Robocopy.exe' -ArgumentList $RoboArgs -PassThru
                                    $InitialCopy | Wait-Process -Timeout $FileCopyTime -ErrorAction SilentlyContinue

                                    #Log exit code when uninstall process finishes
                                    if ($InitialCopy.HasExited){
                                        #Write-Verbose -message "RoboCopy exited with code: $($InitialCopy.ExitCode)"
                                    }
                                    else {
                                        if ($LogFile){Write-WSMLogMessage -Message "$Section RoboCopy process exceeded runtime. Moving on..." -LogFile $LogFile}
                                        else {Write-verbose -message "$Section Sending get-childItem these Parameters: $ParamMessage"}
                                        #$InitialCopy.Kill()
                                    }
                                #endregion - Initial copy section

                                #region - get destination hash
                                    try {
                                        $DestHashObj = Get-FileHash -Path $(($FSO.FullName).Replace($Path,$Destination)) -Algorithm MD5 -ErrorAction Stop
                                    }
                                    catch [Microsoft.PowerShell.Commands.WriteErrorException]{
                                        $DestHashObj = [PSCustomObject]@{
                                            Hash = "No Hash Available - File is being used by another process: $($FSO.path)"
                                            Path = $(($FSO.FullName).Replace($Path,$Destination))
                                        } # [PSCustomObject]@{
                                    } # catch [Microsoft.PowerShell.Commands.WriteErrorException]{
                                    catch {
                                        $DestHashObj = [PSCustomObject]@{
                                            Hash = $_.exception.message
                                            Path = $(($FSO.FullName).Replace($Path,$Destination))
                                        } # [PSCustomObject]@{
                                    }
                                #endregion - get destination hash

                                #Create Final Object
                                [PSCustomObject]@{
                                    SourceHash = $SrcHashObj.Hash
                                    DestHash = $DestHashObj.Hash
                                    SourceFile = $SrcHashObj.path
                                    TargetDest = $DestHashObj.path
                                    FileName = $FSO.Name
                                    FileSize = $FSO.Length
                                }

                            } # else {
                        } # Foreach-object {
            }
            catch [System.ArgumentException]{
                if ($LogFile){Write-WSMLogMessage -Message "$Section Unsupported argument in exclusion check for Object: $($FSO.FullName). $($_.exception.GetType().fullname)" -LogFile $LogFile}
                else {Write-verbose -message "$Section Unsupported argument in exclusion check for Object: $($FSO.FullName). $($_.exception.GetType().fullname)"}
            }
            catch {
                if ($LogFile){Write-WSMLogMessage -Message "$Section Error in FileHash check: $($_.exception.GetType().fullname). For Object $($FSO.FullName)" -LogFile $LogFile}
                else {Write-verbose -message "$Section Error in FileHash check: $($_.exception.GetType().fullname). For Object $($FSO.FullName)"}
            }

            $CorruptObjects = $AllHash.where({$_.SourceHash -ne $_.DestHash})

            if ($LogFile){Write-WSMLogMessage -Message "$Section Corrupt count pre rety from 'AllHash' check var: $($CorruptObjects.count)" -LogFile $LogFile}
            else {Write-verbose -message "$Section Corrupt count pre rety from 'AllHash' check var: $($CorruptObjects.count)"}

            # Testing has shown this object to be updated as well, so I'm going to grab the count now (Probably because I run it against an updated allhash object)
            $CorruptCount = $CorruptObjects.count
            if ($LogFile){Write-WSMLogMessage -Message "$Section Corrupt count pre rety from 'CorruptCount' var: $CorruptCount" -LogFile $LogFile}
            else {Write-verbose -message "$Section Corrupt count pre rety from 'CorruptCount' var: $CorruptCount"}

            <#
                Feed any corrupted objects through this little block to give them another chance to be copied
            #>
            $RetryLogName = "$LogLocation\$($env:COMPUTERNAME)_RetryCopy_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').log"
            Foreach ($CO in $CorruptObjects){

                # if the file exists populate variables and attempt to copy.
                if ($CO.SourceFile -notin $null,''){
                    $SourceDir = Split-Path -path $CO.SourceFile -parent -ErrorAction SilentlyContinue
                    $DestDir = (Split-Path -path $CO.SourceFile -parent -ErrorAction SilentlyContinue).Replace($Path,$Destination)
                    $File = Split-Path -path $CO.SourceFile -Leaf -ErrorAction SilentlyContinue

                    if ($Logging -eq 'Disabled'){
                        $RetryRoboArgs = "$SourceDir $DestDir $File /COPY:DAT /e /z /MT:100 /xjd /r:5 /w:5 /njh /njs /A-:SH $ExcludePRM"
                    }
                    else {
                        $RetryRoboArgs = "$SourceDir $DestDir $File /COPY:DAT /e /z /MT:100 /xjd /r:5 /w:5 /xa:S /njh /A-:SH /LOG+:$RetryLogName $ExcludePRM"
                    }

                    $TrueUpCopy = Start-Process -FilePath 'Robocopy.exe' -ArgumentList $RetryRoboArgs -PassThru
                    $TrueUpCopy | Wait-Process -Timeout 600 -ErrorAction SilentlyContinue    

                    #Attempt to mute a slight bug with error suppression in the get-filehash cmdlet when a path is not found.
                    try {
                         if (Test-path -Path "$DestDir\$File" -ErrorAction SilentlyContinue){
                            if (($RetryHash = Get-FileHash -Path "$DestDir\$File" -Algorithm MD5).Hash -ne $CO.SourceHash){
                                $Failures += "$($CO.SourceFile),$DestDir\$File"
                            }
                            else {
                                $AllHash += $AllHash.where({$_.SourceFile -eq $CO.SourceFile}) | foreach-object {
                                        $_.DestHash = $RetryHash.Hash
                                        $_.TargetDest = $RetryHash.Path
                                }
                            } # else {
                        } # if (Test-path -Path "$DestDir\$File" -ErrorAction SilentlyContinue){
                    }
                    catch {
                        if ($LogFile){Write-WSMLogMessage -Message "$Section Error in FileHash check: $($_.exception.GetType().fullname)" -LogFile $LogFile}
                        else {Write-verbose -message "$Section Error in FileHash check: $($_.exception.GetType().fullname)"}
                    }
                } # if ($CO.SourceFile -notin $null,''){

                <#
                    Leaving blanks alone for now. Hopefully that has been largely eliminated with other changes.
                #>

            } # Foreach ($CO in $CorruptObjects){}

            $FinalCorruptCount = $AllHash.where({$_.SourceHash -ne $_.DestHash})
            if ($LogFile){Write-WSMLogMessage -Message "$Section Corrupt count post rety: $($FinalCorruptCount.count)" -LogFile $LogFile}
            else {Write-verbose -message "$Section Corrupt count post rety: $($FinalCorruptCount.count)"}
        } # if ($ValidDest -eq $true){
    } # Process {
    End {
        try {
            if ($AllHash -and ($Logging -ne 'Disabled')){
                $FullHashLogLocation = "$HashLogLocation\$($env:COMPUTERNAME)_FullHashList_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').json"
                $AllHash | ConvertTo-Json | Out-File -FilePath $FullHashLogLocation -Force -ErrorAction Stop
            }
            elseif (!$AllHash -and ($Logging -ne 'Disabled')){
                $FullHashLogLocation = "$HashLogLocation\$($env:COMPUTERNAME)_FullHashList_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').txt"
                "$env:COMPUTERNAME: No MD5 file hashes were exported. Re-examine parameters before attempting to copy again." | Out-File -FilePath $FullHashLogLocation -Force -ErrorAction Stop
            }
            else {
                if ($LogFile){Write-WSMLogMessage -Message "$Section function not set to export list of MD5 Hashes. Pass the logging parameter as 'Basic' or 'Advanced' to receive additional logging." -LogFile $LogFile}
                else {Write-verbose -message "$Section function not set to export list of MD5 Hashes. Pass the logging parameter as 'Basic' or 'Advanced' to receive additional logging."}
            }
        }
        catch {
            Write-WSMLogMessage "$Section failed to export hashlog result. $($_.Exception.message)"
        }

        #Sanitize in case these are empty
        if (!($InitialHashCount = ($AllHash.SourceFile).count)){$InitialHashCount = 0}
        if (!($RemoteHashCount = ($AllHash.TargetDest).count)){$RemoteHashCount = 0}
        if (!($InitalCorruptCount = $CorruptCount)){$InitalCorruptCount = 0}
        if (!($FinalCorruptCount = $FinalCorruptCount.count)){$FinalCorruptCount = 0}

        #Count and export list of failures if there are any.
        if (!($RetryCorruptCount = $Failures.count) -or ($Failures.count -eq 0)){
            $RetryCorruptCount = 0
        } # if (!($RetryCorruptCount = $Failures.count) -or ($Failures.count -eq 0)){
        else {
            #This log is really most usefull for troubleshooting.
            if ($Logging -eq 'Advanced'){
                try {
                    $FailedLogLocation = "$LogLocation\$($env:COMPUTERNAME)_FailedFileCopyList_$($LogPath)_$(Get-Date -Format 'yyyyMMddHHmmss').log"
                    $Failures | Out-File -FilePath $FailedLogLocation -Force -ErrorAction Stop
                }
                catch {
                    if ($LogFile){Write-WSMLogMessage -Message "$Section List of failed file copies...failed: $($_.exception.gettype().fullname)" -LogFile $LogFile}
                    else {Write-Verbose -Message "$Section List of failed file copies...failed: $($_.exception.gettype().fullname)"}
                }
            } # if ($Logging -eq 'Advanced'){
            else {
                if ($LogFile){Write-WSMLogMessage -Message "$Section function not set to export list of failed file copies. List will not be sent to exiting custom object in order to prevent overloading the buffer." -LogFile $LogFile}
                else {Write-Verbose -Message "$Section function not set to export list of failed file copies. List will not be sent to exiting custom object in order to prevent overloading the buffer."}
            }
        } # else {

        $PerPrepObj = New-Object System.Globalization.CultureInfo -ArgumentList "en-us",$false
        $PerPrepObj.NumberFormat.PercentDecimalDigits = 4 # Moves the decimal
        # Calculate success by hashes
        if (($AllHash.Source).Count -ge '1'){
            $PercentSuccess = ($AllHash.where({$_.SourceHash -eq $_.DestHash})).count / ($AllHash.SourceFile).count
        } # if ($Logging -eq 'Advanced'){
        else {$PercentSuccess = 0}
    
        # Calculate success by source and destination
        if (($AllHash.Source).Count -ge '1'){
            $PercentSuccessFile = ($AllHash.where({(Split-Path -Path $_.SourceFile -Leaf) -eq (Split-Path -path $_.TargetDest -Leaf)})).count / ($AllHash.SourceFile).count
        } # if ($Logging -eq 'Advanced'){
        else {$PercentSuccessFile = 0}

        # Export Dir and File exclusions to csv in advanced logging
        $Exclusions = [PSCustomObject]@{
            RobocopyDirExclusion = $ExcludeDir
            RoboCopyFileExclusion = $ExcludeFile
            RegexFileEscapes = $RegexFileExclusions
            RegexDirEscapes = $RegexDirExclusions
        } #  yes we could just dump out here via the pipeline, but I would like to play with the variable for some other tings in testing...

        # Export exclusions to csv
        $Exclusions | ConvertTo-Csv | Out-File -FilePath "$LogLocation\Exclusions_$(Get-Date -Format 'yyyyMMddHHmmss').csv" -Force 

        [PSCustomObject]@{
            HashMetricResult = $PercentSuccess.ToString('P',$PerPrepObj)
            FileMetricResult = $PercentSuccessFile.ToString('P',$PerPrepObj)
            InitialFailCount = $InitalCorruptCount
            FinalCorruptCount = $FinalCorruptCount
            RetryFailureCount = $RetryCorruptCount
            PreCopyHashCount = $InitialHashCount
            PostCopyHashCount = $RemoteHashCount
            LogFileSettings = $Logging
            CopyCommand = "Robocopy.exe $RoboArgs"
        } # [PSCustomObject]@{
    } # End {
} # function Invoke-WSMValidatedCopySingleItems {

function Enter-WSMDevice {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $DeviceDriveSelect = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Button]$Ok = $null
    [System.Windows.Forms.TextBox]$EnterDevice = $null
    [System.Windows.Forms.Label]$EnterDeviceLabel = $null
    [System.Windows.Forms.ListBox]$DriveListBox = $null
    [System.Windows.Forms.Label]$Label1 = $null
    [System.Windows.Forms.Button]$Select = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #$resources = Invoke-Expression (Get-Content -Path (Join-Path $PSScriptRoot 'Enter-WSMDevice.resources.psd1') -Raw)
    $Ok = (New-Object -TypeName System.Windows.Forms.Button)
    $EnterDevice = (New-Object -TypeName System.Windows.Forms.TextBox)
    $EnterDeviceLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DriveListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $Label1 = (New-Object -TypeName System.Windows.Forms.Label)
    $Select = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DeviceDriveSelect.SuspendLayout()

    #region - add icon
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    #endregion - add icon

    # Ok
    $Ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Ok.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]169,[System.Int32]196))
    $Ok.Name = [System.String]'Ok'
    $Ok.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $Ok.TabIndex = [System.Int32]4
    $Ok.Text = [System.String]'Ok'
    $Ok.UseCompatibleTextRendering = $true
    $Ok.UseVisualStyleBackColor = $true
    $Ok.Enabled = $false

    # EnterDevice
    $EnterDevice.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]21,[System.Int32]53))
    $EnterDevice.Name = [System.String]'EnterDevice'
    $EnterDevice.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]142,[System.Int32]21))
    $EnterDevice.TabIndex = [System.Int32]1

    # EnterDeviceLabel
    $EnterDeviceLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]21,[System.Int32]27))
    $EnterDeviceLabel.Name = [System.String]'EnterDeviceLabel'
    $EnterDeviceLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]150,[System.Int32]23))
    $EnterDeviceLabel.TabIndex = [System.Int32]2
    $EnterDeviceLabel.Text = [System.String]'Enter Device Name'
    $EnterDeviceLabel.UseCompatibleTextRendering = $true

    # DriveListBox
    $DriveListBox.FormattingEnabled = $true
    $DriveListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]21,[System.Int32]121))
    $DriveListBox.Name = [System.String]'DriveListBox'
    $DriveListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $DriveListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]223,[System.Int32]69))
    $DriveListBox.TabIndex = [System.Int32]3

    # Label1
    $Label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]21,[System.Int32]95))
    $Label1.Name = [System.String]'Label1'
    $Label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]128,[System.Int32]23))
    $Label1.TabIndex = [System.Int32]4
    $Label1.Text = [System.String]'Select Drives to Collect'
    $Label1.UseCompatibleTextRendering = $true

    # Select
    $Select.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]169,[System.Int32]51))
    $Select.Name = [System.String]'Select'
    $Select.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $Select.TabIndex = [System.Int32]2
    $Select.Text = [System.String]'Select'
    $Select.UseCompatibleTextRendering = $true
    $Select.UseVisualStyleBackColor = $true
    $Select.Add_Click({
        
        # Empty previous selections
        if ($DriveListBox.Items.count -ge 1){$DriveListBox.Items.Clear()}

        # Go out and check for drives on target devices
        if (Test-Connection -ComputerName $EnterDevice.Text -Protocol WSMan -Count 3 -ErrorAction SilentlyContinue){
            try {
                Get-CimInstance -Computername $EnterDevice.Text -ClassName Win32_LogicalDisk -ErrorAction Stop |
                    where-Object {$_.Drivetype -in '3','2'} |
                        Select-Object -Property 'DeviceID','DriveType','VolumeName','Description' | 
                            ForEach-Object { 
                                if ($DriveListBox.Items -notcontains $_.DeviceID){ 
                                    [void] $DriveListBox.Items.Add(($_.DeviceID).trim())
                                }
                            }
            }
            catch {
                [void] $DriveListBox.Items.Add("Error: $($_.exception.GetType().fullname)")
            }
        }
        else {
            [void] $DriveListBox.Items.Add('Error: Device Not Available')
        }

    }) # $Select.Add_Click({...
    $Select.Enabled = $false

    # CancelButton
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]88,[System.Int32]196))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]5
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true
    $DeviceDriveSelect.CancelButton = $CancelButton
    $DeviceDriveSelect.Controls.Add($CancelButton)

    # DeviceDriveSelect
    $DeviceDriveSelect.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]270,[System.Int32]255))
    $DeviceDriveSelect.Controls.Add($Select)
    $DeviceDriveSelect.Controls.Add($Label1)
    $DeviceDriveSelect.Controls.Add($DriveListBox)
    $DeviceDriveSelect.Controls.Add($EnterDeviceLabel)
    $DeviceDriveSelect.Controls.Add($EnterDevice)
    $DeviceDriveSelect.Controls.Add($Ok)
    $DeviceDriveSelect.Name = [System.String]'DeviceDriveSelect'
    $DeviceDriveSelect.Text = [System.String]'Select Target Device'
    $DeviceDriveSelect.ResumeLayout($false)
    $DeviceDriveSelect.PerformLayout()
    # Use Icon
    $DeviceDriveSelect.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    # Wait for a the user to select a drive before OK can be selected.
    $DriveListBox.add_SelectedIndexChanged({$Ok.Enabled = $true})

    $EnterDevice.add_TextChanged({$Select.Enabled = $true})

    # Run the dialog 
    $Result = $DeviceDriveSelect.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        if (($DriveListBox.SelectedItems).foreach({$_ -notmatch 'Error'}) ){
            [pscustomObject]@{
                Result = 'Success'
                ComputerName = $EnterDevice.Text
                Drives = $DriveListBox.SelectedItems
            }
            $DeviceDriveSelect.Dispose()
        }
        else {
            [pscustomObject]@{
                Result = 'Error'
                ComputerName = $EnterDevice.Text
                Drives = $DriveListBox.SelectedItems
            }
            $DeviceDriveSelect.Dispose()
        }
    } # if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        [pscustomObject]@{
            Result = 'Cancel'
            ComputerName = $EnterDevice.Text
            Drives = $DriveListBox.SelectedItems
        }
        $DeviceDriveSelect.Dispose()
    }

    Add-Member -InputObject $DeviceDriveSelect -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $DeviceDriveSelect -Name Ok -Value $Ok -MemberType NoteProperty
    Add-Member -InputObject $DeviceDriveSelect -Name EnterDevice -Value $EnterDevice -MemberType NoteProperty
    Add-Member -InputObject $DeviceDriveSelect -Name EnterDeviceLabel -Value $EnterDeviceLabel -MemberType NoteProperty
    Add-Member -InputObject $DeviceDriveSelect -Name DriveListBox -Value $DriveListBox -MemberType NoteProperty
    Add-Member -InputObject $DeviceDriveSelect -Name Label1 -Value $Label1 -MemberType NoteProperty
} # function Enter-WSMDevice {

function Enter-WSMDeviceManually {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $ManualDriveForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    [System.Windows.Forms.TextBox]$EnterDeviceName = $null
    [System.Windows.Forms.Label]$DeviceNameLabel = $null
    [System.Windows.Forms.ComboBox]$DriveComboBox = $null
    [System.Windows.Forms.Label]$DriveSelect = $null
    [System.Windows.Forms.ListBox]$DrivesListBox = $null
    [System.Windows.Forms.Button]$AddDriveButton = $null
    [System.Windows.Forms.Button]$RemoveDriveButton = $null
    [System.Windows.Forms.Label]$Label1 = $null
    [System.Windows.Forms.Label]$InstructionLabel = $null
    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $EnterDeviceName = (New-Object -TypeName System.Windows.Forms.TextBox)
    $DeviceNameLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DriveComboBox = (New-Object -TypeName System.Windows.Forms.ComboBox)
    $DriveSelect = (New-Object -TypeName System.Windows.Forms.Label)
    $DrivesListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $AddDriveButton = (New-Object -TypeName System.Windows.Forms.Button)
    $RemoveDriveButton = (New-Object -TypeName System.Windows.Forms.Button)
    $Label1 = (New-Object -TypeName System.Windows.Forms.Label)
    $InstructionLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $ManualDriveForm.SuspendLayout()

    #region - add icon
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    #endregion - add icon

    #OkButton
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]352,[System.Int32]246))
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]9
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true
    $OkButton.add_Click($OkButton_Click)
    $OkButton.Enabled = $false

    #CancelButton
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]260,[System.Int32]246))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]8
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true
    $ManualDriveForm.CancelButton = $CancelButton

    #EnterDeviceName
    $EnterDeviceName.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]45))
    $EnterDeviceName.Name = [System.String]'EnterDeviceName'
    $EnterDeviceName.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]21))
    $EnterDeviceName.TabIndex = [System.Int32]3
    $EnterDeviceName.Text = [System.String]'LOCALHOST'
    $EnterDeviceName.ReadOnly = $true
    #$EnterDeviceName.add_TextChanged($TextBox1_TextChanged)

    #DeviceNameLabel
    $DeviceNameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]24))
    $DeviceNameLabel.Name = [System.String]'DeviceNameLabel'
    $DeviceNameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]120,[System.Int32]18))
    $DeviceNameLabel.TabIndex = [System.Int32]1
    $DeviceNameLabel.Text = [System.String]'Enter Device Name'
    $DeviceNameLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $DeviceNameLabel.UseCompatibleTextRendering = $true
    #$DeviceNameLabel.add_Click($Label1_Click) #wtf is there a add_click here?

    #ComboBox1
    $DriveComboBox.FormattingEnabled = $true
    $DriveComboBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]192,[System.Int32]45))
    $DriveComboBox.Name = [System.String]'DriveComboBox'
    $DriveComboBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]21))
    $DriveComboBox.TabIndex = [System.Int32]4

    #DriveSelect
    $DriveSelect.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]192,[System.Int32]19))
    $DriveSelect.Name = [System.String]'DriveSelect'
    $DriveSelect.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]175,[System.Int32]23))
    $DriveSelect.TabIndex = [System.Int32]2
    $DriveSelect.Text = [System.String]'Select Drives for Data Collection'
    $DriveSelect.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $DriveSelect.UseCompatibleTextRendering = $true

    #DrivesListBox
    $DrivesListBox.FormattingEnabled = $true
    $DrivesListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]94))
    $DrivesListBox.Name = [System.String]'DrivesListBox'
    $DrivesListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]147))
    $DrivesListBox.TabIndex = [System.Int32]7

    #AddDriveButton
    $AddDriveButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]352,[System.Int32]43))
    $AddDriveButton.Name = [System.String]'AddDriveButton'
    $AddDriveButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $AddDriveButton.TabIndex = [System.Int32]5
    $AddDriveButton.Text = [System.String]'Add'
    $AddDriveButton.UseCompatibleTextRendering = $true
    $AddDriveButton.UseVisualStyleBackColor = $true
    $AddDriveButton.add_Click({
        if ($DriveComboBox.SelectedItem -ne $null){
            if ($DrivesListBox.Items -notcontains $DriveComboBox.SelectedItem){
                [void] $DrivesListBox.Items.Add($DriveComboBox.SelectedItem)
                $OkButton.Enabled = $true
            }
        }
    }) # $AddDriveButton.add_Click({

    $RemoveDriveButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]352,[System.Int32]68))
    $RemoveDriveButton.Name = [System.String]'RemoveDriveButton'
    $RemoveDriveButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $RemoveDriveButton.TabIndex = [System.Int32]5
    $RemoveDriveButton.Text = [System.String]'Remove'
    $RemoveDriveButton.UseCompatibleTextRendering = $true
    $RemoveDriveButton.UseVisualStyleBackColor = $true
    $RemoveDriveButton.add_Click({
        if ($DrivesListBox.SelectedItem -ne $null){
            [void] $DrivesListBox.Items.Remove($DrivesListBox.SelectedItem)
            if ($DrivesListBox.Items.count -lt 1){$OkButton.Enabled = $false}
        }
    })

    #Label1
    $Label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]72))
    $Label1.Name = [System.String]'Label1'
    $Label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]19))
    $Label1.TabIndex = [System.Int32]6
    $Label1.Text = [System.String]'Selected Drives'
    $Label1.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $Label1.UseCompatibleTextRendering = $true

    #InstructionLabel
    $InstructionLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]192,[System.Int32]94))
    $InstructionLabel.Name = [System.String]'InstructionLabel'
    $InstructionLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]235,[System.Int32]147))
    $InstructionLabel.TabIndex = [System.Int32]10
    $InstructionLabel.Text = [System.String]'Select drives that are local to the device and not attached externally such as an USB device. The data collection script will attempt to filter any devices that do not meet the required criteria.'
    $InstructionLabel.UseCompatibleTextRendering = $true

    #ManualDriveForm
    $ManualDriveForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]453,[System.Int32]289))
    $ManualDriveForm.Controls.Add($InstructionLabel)
    $ManualDriveForm.Controls.Add($Label1)
    $ManualDriveForm.Controls.Add($AddDriveButton)
    $ManualDriveForm.Controls.Add($RemoveDriveButton)
    $ManualDriveForm.Controls.Add($DrivesListBox)
    $ManualDriveForm.Controls.Add($DriveSelect)
    $ManualDriveForm.Controls.Add($DriveComboBox)
    $ManualDriveForm.Controls.Add($DeviceNameLabel)
    $ManualDriveForm.Controls.Add($EnterDeviceName)
    $ManualDriveForm.Controls.Add($CancelButton)
    $ManualDriveForm.Controls.Add($OkButton)
    $ManualDriveForm.Name = [System.String]'ManualDriveForm'
    $ManualDriveForm.Text = [System.String]'Manually Select Device and Drive (USB)'
    $ManualDriveForm.ResumeLayout($false)
    $ManualDriveForm.PerformLayout()
    $ManualDriveForm.CancelButton = $CancelButton
    
    # Use Icon
    $ManualDriveForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    $DriveArray = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')

    $DriveArray | ForEach-Object {
        $DriveComboBox.Items.add($_)
    }

    # Wait for a the user to select a drive before OK can be selected.

    # Run the dialog 
    $Result = $ManualDriveForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        [pscustomObject]@{
            Result = 'Success'
            ComputerName = $EnterDeviceName.Text
            Drives = "$($DrivesListBox.Items):"
        }
        $ManualDriveForm.Dispose()
    } # if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        [pscustomObject]@{
            Result = 'Cancel'
            ComputerName = 'No Data' # Null creates an Error
            Drives = 'No Data' # Null creates an Error
        }
        $ManualDriveForm.Dispose()
    }

    Add-Member -InputObject $ManualDriveForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name OkButton -Value $OkButton -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name EnterDeviceName -Value $EnterDeviceName -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name DeviceNameLabel -Value $DeviceNameLabel -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name ComboBox1 -Value $DriveComboBox -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name DriveSelect -Value $DriveSelect -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name DrivesListBox -Value $DrivesListBox -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name AddDriveButton -Value $AddDriveButton -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name RemoveDriveButton -Value $RemoveDriveButton -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name Label1 -Value $Label1 -MemberType NoteProperty
    Add-Member -InputObject $ManualDriveForm -Name InstructionLabel -Value $InstructionLabel -MemberType NoteProperty
} # function Enter-WSMDeviceManually {

function Edit-WSMExclusions {
    param (
        $FileList,
        $DirList
    )

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $ExclusionsForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.ListBox]$DirExListBox = $null
    [System.Windows.Forms.ListBox]$FileExListBox = $null
    [System.Windows.Forms.Button]$AddDirButton = $null
    [System.Windows.Forms.Button]$AddFileButton = $null
    [System.Windows.Forms.Button]$RemDirButton = $null
    [System.Windows.Forms.Button]$RemFileButton = $null
    [System.Windows.Forms.Button]$DefaultFileButton = $null
    [System.Windows.Forms.Button]$DefaultDirButton = $null
    [System.Windows.Forms.Button]$ApplyDirButton = $null
    [System.Windows.Forms.Button]$ApplyFileButton = $null
    [System.Windows.Forms.Button]$ClearDirsButton = $null
    [System.Windows.Forms.Button]$ClearFilesDir = $null
    [System.Windows.Forms.TextBox]$NewDirTextBox = $null
    [System.Windows.Forms.TextBox]$NewFileTextBox = $null
    [System.Windows.Forms.Label]$DirExLabel = $null
    [System.Windows.Forms.Label]$FileExLabel = $null
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Add objects
    $DirExListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $FileExListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $AddDirButton = (New-Object -TypeName System.Windows.Forms.Button)
    $AddFileButton = (New-Object -TypeName System.Windows.Forms.Button)
    $RemDirButton = (New-Object -TypeName System.Windows.Forms.Button)
    $RemFileButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DefaultFileButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DefaultDirButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ApplyDirButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ApplyFileButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ClearDirsButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ClearFilesDir = (New-Object -TypeName System.Windows.Forms.Button)
    $NewDirTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $NewFileTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $DirExLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $FileExLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ExclusionsForm.SuspendLayout()

    #region - add icon
        # This region is based on code found on a couple different sites as it falls pretty far outside my normal knowledge base            
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    #endregion - add icon

    # DirExListBox
    $DirExListBox.FormattingEnabled = $true
    $DirExListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]65))
    $DirExListBox.Name = [System.String]'DirExListBox'
    $DirExListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]217,[System.Int32]251))
    $DirExListBox.TabIndex = [System.Int32]0

    # FileExListBox
    $FileExListBox.FormattingEnabled = $true
    $FileExListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]65))
    $FileExListBox.Name = [System.String]'FileExListBox'
    $FileExListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]217,[System.Int32]251))
    $FileExListBox.TabIndex = [System.Int32]1
    $FileExListBox.add_SelectedIndexChanged($ListBox2_SelectedIndexChanged)

    # AddDirButton
    $AddDirButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]34))
    $AddDirButton.Name = [System.String]'AddDirButton'
    $AddDirButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $AddDirButton.TabIndex = [System.Int32]2
    $AddDirButton.Text = [System.String]'Add Directory'
    $AddDirButton.UseCompatibleTextRendering = $true
    $AddDirButton.UseVisualStyleBackColor = $true
    $AddDirButton.Add_Click({
        # Add items to the dir list box
        if ($DirExListBox.Items -notcontains $NewDirTextBox.Text){
            $DirExListBox.Items.add($NewDirTextBox.Text)
        }
    }) # $ApplyDirButton.Add_Click({...})

    # AddFileButton
    $AddFileButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]34))
    $AddFileButton.Name = [System.String]'AddFileButton'
    $AddFileButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $AddFileButton.TabIndex = [System.Int32]3
    $AddFileButton.Text = [System.String]'Add File'
    $AddFileButton.UseCompatibleTextRendering = $true
    $AddFileButton.UseVisualStyleBackColor = $true
    $AddFileButton.Add_Click({
        # Add items to the dir list box
        if ($FileExListBox.Items -notcontains $NewFileTextBox.Text){
            $FileExListBox.Items.add($NewFileTextBox.Text)
        }
    }) # $ApplyFileButton.Add_Click({...})

    # RemDirButton
    $RemDirButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]322))
    $RemDirButton.Name = [System.String]'RemDirButton'
    $RemDirButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $RemDirButton.TabIndex = [System.Int32]4
    $RemDirButton.Text = [System.String]'Remove Dir'
    $RemDirButton.UseCompatibleTextRendering = $true
    $RemDirButton.UseVisualStyleBackColor = $true
    $RemDirButton.Add_Click({
        while($DirExListBox.SelectedItems){
            $DirExListBox.Items.Remove($DirExListBox.SelectedItems[0])
        }
    })

    # RemFileButton
    $RemFileButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]322))
    $RemFileButton.Name = [System.String]'RemFileButton'
    $RemFileButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $RemFileButton.TabIndex = [System.Int32]5
    $RemFileButton.Text = [System.String]'Remove File'
    $RemFileButton.UseCompatibleTextRendering = $true
    $RemFileButton.UseVisualStyleBackColor = $true
    $RemFileButton.Add_Click({
        while($FileExListBox.SelectedItems){
            $FileExListBox.Items.Remove($FileExListBox.SelectedItems[0])
        }
    })

    # DefaultFileButton
    $DefaultFileButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]322))
    $DefaultFileButton.Name = [System.String]'DefaultFileButton'
    $DefaultFileButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $DefaultFileButton.TabIndex = [System.Int32]6
    $DefaultFileButton.Text = [System.String]'Defaults'
    $DefaultFileButton.UseCompatibleTextRendering = $true
    $DefaultFileButton.UseVisualStyleBackColor = $true
    $DefaultFileButton.Add_Click({
        if ($FileExListBox.Items.count -ge 1){$FileExListBox.Items.Clear()}
        # Revert files to default
        $FileList | ForEach-Object {
            $FileExListBox.Items.Add($_)
        }
    })

    # DefaultDirButton
    $DefaultDirButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]322))
    $DefaultDirButton.Name = [System.String]'DefaultDirButton'
    $DefaultDirButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $DefaultDirButton.TabIndex = [System.Int32]7
    $DefaultDirButton.Text = [System.String]'Defaults'
    $DefaultDirButton.UseCompatibleTextRendering = $true
    $DefaultDirButton.UseVisualStyleBackColor = $true
    $DefaultDirButton.Add_click({
        # Clear any changes because, 'Cancel'
        if ($DirExListBox.Items.count -ge 1){$DirExListBox.Items.Clear()}
        # Revert to default Dirs
        $DirList | ForEach-Object {
            $DirExListBox.Items.Add($_)
        }
    })

    # ApplyDirButton
    $ApplyDirButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]351))
    $ApplyDirButton.Name = [System.String]'ApplyDirButton'
    $ApplyDirButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ApplyDirButton.TabIndex = [System.Int32]8
    $ApplyDirButton.Text = [System.String]'Apply Directories'
    $ApplyDirButton.UseCompatibleTextRendering = $true
    $ApplyDirButton.UseVisualStyleBackColor = $true
    $ApplyDirButton.Add_click({
        if ($DirExListBox.SelectedItems -ge 1){$DirExListBox.SelectedItems.Clear()}
    })
    #$ApplyDirButton

    # ApplyFileButton
    $ApplyFileButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]351))
    $ApplyFileButton.Name = [System.String]'ApplyFileButton'
    $ApplyFileButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ApplyFileButton.TabIndex = [System.Int32]9
    $ApplyFileButton.Text = [System.String]'Apply Files'
    $ApplyFileButton.UseCompatibleTextRendering = $true
    $ApplyFileButton.UseVisualStyleBackColor = $true
    $ApplyFileButton.Add_click({
        if ($FileExListBox.SelectedItems -ge 1){$FileExListBox.SelectedItems.Clear()}
    })
    #$ApplyFileButton

    # ClearDirsButton
    $ClearDirsButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]351))
    $ClearDirsButton.Name = [System.String]'ClearDirsButton'
    $ClearDirsButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ClearDirsButton.TabIndex = [System.Int32]10
    $ClearDirsButton.Text = [System.String]'Clear Directories'
    $ClearDirsButton.UseCompatibleTextRendering = $true
    $ClearDirsButton.UseVisualStyleBackColor = $true
    $ClearDirsButton.Add_click({
        if ($DirExListBox.Items.count -ge 1){$DirExListBox.Items.Clear()}
    })

    # ClearFilesDir
    $ClearFilesDir.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]351))
    $ClearFilesDir.Name = [System.String]'ClearFilesDir'
    $ClearFilesDir.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ClearFilesDir.TabIndex = [System.Int32]11
    $ClearFilesDir.Text = [System.String]'Clear Files'
    $ClearFilesDir.UseCompatibleTextRendering = $true
    $ClearFilesDir.UseVisualStyleBackColor = $true
    $ClearFilesDir.Add_click({
        if ($FileExListBox.Items.count -ge 1){$FileExListBox.Items.Clear()}
    })

    # NewDirTextBox
    $NewDirTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]36))
    $NewDirTextBox.Name = [System.String]'NewDirTextBox'
    $NewDirTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]21))
    $NewDirTextBox.TabIndex = [System.Int32]12

    # NewFileTextBox
    $NewFileTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]36))
    $NewFileTextBox.Name = [System.String]'NewFileTextBox'
    $NewFileTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]21))
    $NewFileTextBox.TabIndex = [System.Int32]13

    # DirExLabel
    $DirExLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
    $DirExLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]9))
    $DirExLabel.Name = [System.String]'DirExLabel'
    $DirExLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]123,[System.Int32]24))
    $DirExLabel.TabIndex = [System.Int32]14
    $DirExLabel.Text = [System.String]'Directory Exclusions'
    $DirExLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $DirExLabel.UseCompatibleTextRendering = $true
    $DirExLabel.add_Click($DirExLabel_Click)

    # FileExLabel
    $FileExLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
    $FileExLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]9))
    $FileExLabel.Name = [System.String]'FileExLabel'
    $FileExLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]24))
    $FileExLabel.TabIndex = [System.Int32]15
    $FileExLabel.Text = [System.String]'File Exclusions'
    $FileExLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $FileExLabel.UseCompatibleTextRendering = $true
    $FileExLabel.add_Click($FileExLabel_Click)

    # OkButton
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]513,[System.Int32]452))
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]16
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true
    $OkButton.add_Click($OkButton_Click)

    # CancelButton
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]432,[System.Int32]452))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]17
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true

    # ExclusionsForm
    $ExclusionsForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]600,[System.Int32]487))
    $ExclusionsForm.Controls.Add($CancelButton)
    $ExclusionsForm.Controls.Add($OkButton)
    $ExclusionsForm.Controls.Add($FileExLabel)
    $ExclusionsForm.Controls.Add($DirExLabel)
    $ExclusionsForm.Controls.Add($NewFileTextBox)
    $ExclusionsForm.Controls.Add($NewDirTextBox)
    $ExclusionsForm.Controls.Add($ClearFilesDir)
    $ExclusionsForm.Controls.Add($ClearDirsButton)
    $ExclusionsForm.Controls.Add($ApplyFileButton)
    $ExclusionsForm.Controls.Add($ApplyDirButton)
    $ExclusionsForm.Controls.Add($DefaultDirButton)
    $ExclusionsForm.Controls.Add($DefaultFileButton)
    $ExclusionsForm.Controls.Add($RemFileButton)
    $ExclusionsForm.Controls.Add($RemDirButton)
    $ExclusionsForm.Controls.Add($AddFileButton)
    $ExclusionsForm.Controls.Add($AddDirButton)
    $ExclusionsForm.Controls.Add($FileExListBox)
    $ExclusionsForm.Controls.Add($DirExListBox)
    $ExclusionsForm.Name = [System.String]'ExclusionsForm'
    $ExclusionsForm.Text = [System.String]'Exclusions'
    $ExclusionsForm.ResumeLayout($false)
    $ExclusionsForm.PerformLayout()    
    # Use Icon
    $ExclusionsForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    # Create the rows for the grid
    # Add Dirs
    $DirList | ForEach-Object {
        $DirExListBox.Items.Add($_)
    }
    # Add Files
    $FileList | ForEach-Object {
        $FileExListBox.Items.Add($_)
    }

    # Run the dialog 
    $Result = $ExclusionsForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
            [pscustomObject]@{
                Result = 'Success'
                DirExclusions = $DirExListBox.Items
                FileExclusions = $FileExListBox.Items
            }
            $ExclusionsForm.Dispose()

    } # if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){

        # Clear any changes because, 'Cancel'
        if ($DirExListBox.Items.count -ge 1){$DirExListBox.Items.Clear()}
        # Revert to default Dirs
        $DirList | ForEach-Object {
            $DirExListBox.Items.Add($_)
        }
        if ($FileExListBox.Items.count -ge 1){$FileExListBox.Items.Clear()}
        # Revert files to default
        $FileList | ForEach-Object {
            $FileExListBox.Items.Add($_)
        }

        [pscustomObject]@{
            Result = 'Cancel'
            DirExclusions = $DirExListBox.Items
            FileExclusions = $FileExListBox.Items
        }
        $ExclusionsForm.Dispose()
    }

    # Add Zee Members
    Add-Member -InputObject $ExclusionsForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name DirExListBox -Value $DirExListBox -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name FileExListBox -Value $FileExListBox -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name AddDirButton -Value $AddDirButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name AddFileButton -Value $AddFileButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name RemDirButton -Value $RemDirButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name RemFileButton -Value $RemFileButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name DefaultFileButton -Value $DefaultFileButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name DefaultDirButton -Value $DefaultDirButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name ApplyDirButton -Value $ApplyDirButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name ApplyFileButton -Value $ApplyFileButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name ClearDirsButton -Value $ClearDirsButton -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name ClearFilesDir -Value $ClearFilesDir -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name NewDirTextBox -Value $NewDirTextBox -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name NewFileTextBox -Value $NewFileTextBox -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name DirExLabel -Value $DirExLabel -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name FileExLabel -Value $FileExLabel -MemberType NoteProperty
    Add-Member -InputObject $ExclusionsForm -Name OkButton -Value $OkButton -MemberType NoteProperty
    #Add-Member -InputObject $ExclusionsForm -Name CancelButton -Value $CancelButton -MemberType NoteProperty
} # function Edit-WSMExclusions {

function Edit-WSMServicesProcs {
    param (
        $Processes,
        $Services
    )

    #Dir to Service
    #File to Proc

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $SvcProcForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.ListBox]$SvcExListBox = $null
    [System.Windows.Forms.ListBox]$ProcExListBox = $null
    [System.Windows.Forms.Button]$AddSvcButton = $null
    [System.Windows.Forms.Button]$AddProcButton = $null
    [System.Windows.Forms.Button]$RemSvcButton = $null
    [System.Windows.Forms.Button]$RemProcButton = $null
    [System.Windows.Forms.Button]$DefaultProcButton = $null
    [System.Windows.Forms.Button]$DefaultSvcButton = $null
    [System.Windows.Forms.Button]$ApplySvcButton = $null
    [System.Windows.Forms.Button]$ApplyProcButton = $null
    [System.Windows.Forms.Button]$ClearSvcsButton = $null
    [System.Windows.Forms.Button]$ClearProcsDir = $null
    [System.Windows.Forms.TextBox]$NewSvcTextBox = $null
    [System.Windows.Forms.TextBox]$NewProcTextBox = $null
    [System.Windows.Forms.Label]$SvcExLabel = $null
    [System.Windows.Forms.Label]$ProcExLabel = $null
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Add objects
    $SvcExListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $ProcExListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $AddSvcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $AddProcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $RemSvcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $RemProcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DefaultProcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DefaultSvcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ApplySvcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ApplyProcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ClearSvcsButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ClearProcsDir = (New-Object -TypeName System.Windows.Forms.Button)
    $NewSvcTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $NewProcTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $SvcExLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $ProcExLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $SvcProcForm.SuspendLayout()

    #region - add icon
    # This region is based on code found on a couple different sites as it falls pretty far outside my normal knowledge base
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    #endregion - add icon

    # SvcExListBox
    $SvcExListBox.FormattingEnabled = $true
    $SvcExListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]65))
    $SvcExListBox.Name = [System.String]'SvcExListBox'
    $SvcExListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]217,[System.Int32]251))
    $SvcExListBox.TabIndex = [System.Int32]0

    # ProcExListBox
    $ProcExListBox.FormattingEnabled = $true
    $ProcExListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]65))
    $ProcExListBox.Name = [System.String]'ProcExListBox'
    $ProcExListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]217,[System.Int32]251))
    $ProcExListBox.TabIndex = [System.Int32]1
    $ProcExListBox.add_SelectedIndexChanged($ListBox2_SelectedIndexChanged)

    # AddSvcButton
    $AddSvcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]34))
    $AddSvcButton.Name = [System.String]'AddSvcButton'
    $AddSvcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $AddSvcButton.TabIndex = [System.Int32]2
    $AddSvcButton.Text = [System.String]'Add Service'
    $AddSvcButton.UseCompatibleTextRendering = $true
    $AddSvcButton.UseVisualStyleBackColor = $true
    $AddSvcButton.Add_Click({
        # Add items to the dir list box
        if ($SvcExListBox.Items -notcontains $NewSvcTextBox.Text){
            $SvcExListBox.Items.add($NewSvcTextBox.Text)
        }
    }) # $ApplySvcButton.Add_Click({...})

    # AddProcButton
    $AddProcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]34))
    $AddProcButton.Name = [System.String]'AddProcButton'
    $AddProcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $AddProcButton.TabIndex = [System.Int32]3
    $AddProcButton.Text = [System.String]'Add Process'
    $AddProcButton.UseCompatibleTextRendering = $true
    $AddProcButton.UseVisualStyleBackColor = $true
    $AddProcButton.Add_Click({
        # Add items to the dir list box
        if ($ProcExListBox.Items -notcontains $NewProcTextBox.Text){
            $ProcExListBox.Items.add($NewProcTextBox.Text)
        }
    }) # $ApplyProcButton.Add_Click({...})

    # RemSvcButton
    $RemSvcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]322))
    $RemSvcButton.Name = [System.String]'RemSvcButton'
    $RemSvcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $RemSvcButton.TabIndex = [System.Int32]4
    $RemSvcButton.Text = [System.String]'Remove Service'
    $RemSvcButton.UseCompatibleTextRendering = $true
    $RemSvcButton.UseVisualStyleBackColor = $true
    $RemSvcButton.Add_Click({
        while($SvcExListBox.SelectedItems){
            $SvcExListBox.Items.Remove($SvcExListBox.SelectedItems[0])
        }
    })

    # RemProcButton
    $RemProcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]322))
    $RemProcButton.Name = [System.String]'RemProcButton'
    $RemProcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $RemProcButton.TabIndex = [System.Int32]5
    $RemProcButton.Text = [System.String]'Remove Process'
    $RemProcButton.UseCompatibleTextRendering = $true
    $RemProcButton.UseVisualStyleBackColor = $true
    $RemProcButton.Add_Click({
        while($ProcExListBox.SelectedItems){
            $ProcExListBox.Items.Remove($ProcExListBox.SelectedItems[0])
        }
    })

    # DefaultProcButton
    $DefaultProcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]322))
    $DefaultProcButton.Name = [System.String]'DefaultProcButton'
    $DefaultProcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $DefaultProcButton.TabIndex = [System.Int32]6
    $DefaultProcButton.Text = [System.String]'Defaults'
    $DefaultProcButton.UseCompatibleTextRendering = $true
    $DefaultProcButton.UseVisualStyleBackColor = $true
    $DefaultProcButton.Add_Click({
        if ($ProcExListBox.Items.count -ge 1){$ProcExListBox.Items.Clear()}
        # Revert files to default
        $Processes | ForEach-Object {
            $ProcExListBox.Items.Add($_)
        }
    })

    # DefaultSvcButton
    $DefaultSvcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]322))
    $DefaultSvcButton.Name = [System.String]'DefaultSvcButton'
    $DefaultSvcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $DefaultSvcButton.TabIndex = [System.Int32]7
    $DefaultSvcButton.Text = [System.String]'Defaults'
    $DefaultSvcButton.UseCompatibleTextRendering = $true
    $DefaultSvcButton.UseVisualStyleBackColor = $true
    $DefaultSvcButton.Add_click({
        # Clear any changes because, 'Cancel'
        if ($SvcExListBox.Items.count -ge 1){$SvcExListBox.Items.Clear()}
        # Revert to default Dirs
        $Services | ForEach-Object {
            $SvcExListBox.Items.Add($_)
        }
    })

    # ApplySvcButton
    $ApplySvcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]160,[System.Int32]351))
    $ApplySvcButton.Name = [System.String]'ApplySvcButton'
    $ApplySvcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ApplySvcButton.TabIndex = [System.Int32]8
    $ApplySvcButton.Text = [System.String]'Apply Service'
    $ApplySvcButton.UseCompatibleTextRendering = $true
    $ApplySvcButton.UseVisualStyleBackColor = $true
    $ApplySvcButton.Add_click({
        if ($SvcExListBox.SelectedItems -ge 1){$SvcExListBox.SelectedItems.Clear()}
    })
    #$ApplySvcButton

    # ApplyProcButton
    $ApplyProcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]462,[System.Int32]351))
    $ApplyProcButton.Name = [System.String]'ApplyProcButton'
    $ApplyProcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ApplyProcButton.TabIndex = [System.Int32]9
    $ApplyProcButton.Text = [System.String]'Apply Process'
    $ApplyProcButton.UseCompatibleTextRendering = $true
    $ApplyProcButton.UseVisualStyleBackColor = $true
    $ApplyProcButton.Add_click({
        if ($ProcExListBox.SelectedItems -ge 1){$ProcExListBox.SelectedItems.Clear()}
    })
    #$ApplyProcButton

    # ClearSvcsButton
    $ClearSvcsButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]351))
    $ClearSvcsButton.Name = [System.String]'ClearSvcsButton'
    $ClearSvcsButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ClearSvcsButton.TabIndex = [System.Int32]10
    $ClearSvcsButton.Text = [System.String]'Clear Services'
    $ClearSvcsButton.UseCompatibleTextRendering = $true
    $ClearSvcsButton.UseVisualStyleBackColor = $true
    $ClearSvcsButton.Add_click({
        if ($SvcExListBox.Items.count -ge 1){$SvcExListBox.Items.Clear()}
    })

    # ClearProcsDir
    $ClearProcsDir.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]351))
    $ClearProcsDir.Name = [System.String]'ClearProcsDir'
    $ClearProcsDir.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]97,[System.Int32]23))
    $ClearProcsDir.TabIndex = [System.Int32]11
    $ClearProcsDir.Text = [System.String]'Clear Processes'
    $ClearProcsDir.UseCompatibleTextRendering = $true
    $ClearProcsDir.UseVisualStyleBackColor = $true
    $ClearProcsDir.Add_click({
        if ($ProcExListBox.Items.count -ge 1){$ProcExListBox.Items.Clear()}
    })

    # NewSvcTextBox
    $NewSvcTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]36))
    $NewSvcTextBox.Name = [System.String]'NewSvcTextBox'
    $NewSvcTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]21))
    $NewSvcTextBox.TabIndex = [System.Int32]12

    # NewProcTextBox
    $NewProcTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]36))
    $NewProcTextBox.Name = [System.String]'NewProcTextBox'
    $NewProcTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]21))
    $NewProcTextBox.TabIndex = [System.Int32]13

    # SvcExLabel
    $SvcExLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
    $SvcExLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]40,[System.Int32]9))
    $SvcExLabel.Name = [System.String]'SvcExLabel'
    $SvcExLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]123,[System.Int32]24))
    $SvcExLabel.TabIndex = [System.Int32]14
    $SvcExLabel.Text = [System.String]'Services to Stop'
    $SvcExLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $SvcExLabel.UseCompatibleTextRendering = $true
    $SvcExLabel.add_Click($SvcExLabel_Click)

    # ProcExLabel
    $ProcExLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
    $ProcExLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]342,[System.Int32]9))
    $ProcExLabel.Name = [System.String]'ProcExLabel'
    $ProcExLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]114,[System.Int32]24))
    $ProcExLabel.TabIndex = [System.Int32]15
    $ProcExLabel.Text = [System.String]'Processes to Stop'
    $ProcExLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $ProcExLabel.UseCompatibleTextRendering = $true
    $ProcExLabel.add_Click($ProcExLabel_Click)

    # OkButton
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]513,[System.Int32]452))
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]16
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true
    $OkButton.add_Click($OkButton_Click)

    # CancelButton
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]432,[System.Int32]452))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]17
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true

    # SvcProcForm
    $SvcProcForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]600,[System.Int32]487))
    $SvcProcForm.Controls.Add($CancelButton)
    $SvcProcForm.Controls.Add($OkButton)
    $SvcProcForm.Controls.Add($ProcExLabel)
    $SvcProcForm.Controls.Add($SvcExLabel)
    $SvcProcForm.Controls.Add($NewProcTextBox)
    $SvcProcForm.Controls.Add($NewSvcTextBox)
    $SvcProcForm.Controls.Add($ClearProcsDir)
    $SvcProcForm.Controls.Add($ClearSvcsButton)
    $SvcProcForm.Controls.Add($ApplyProcButton)
    $SvcProcForm.Controls.Add($ApplySvcButton)
    $SvcProcForm.Controls.Add($DefaultSvcButton)
    $SvcProcForm.Controls.Add($DefaultProcButton)
    $SvcProcForm.Controls.Add($RemProcButton)
    $SvcProcForm.Controls.Add($RemSvcButton)
    $SvcProcForm.Controls.Add($AddProcButton)
    $SvcProcForm.Controls.Add($AddSvcButton)
    $SvcProcForm.Controls.Add($ProcExListBox)
    $SvcProcForm.Controls.Add($SvcExListBox)
    $SvcProcForm.Name = [System.String]'SvcProcForm'
    $SvcProcForm.Text = [System.String]'Services and Processes'
    $SvcProcForm.ResumeLayout($false)
    $SvcProcForm.PerformLayout()    
    # Use Icon
    $SvcProcForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    # Create the rows for the grid
    # Add Dirs
    $Services | ForEach-Object {
        $SvcExListBox.Items.Add($_)
    }
    # Add Files
    $Processes | ForEach-Object {
        $ProcExListBox.Items.Add($_)
    }

    # Run the dialog 
    $Result = $SvcProcForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
            [pscustomObject]@{
                Result = 'Success'
                ServiceList = $SvcExListBox.Items
                ProcList = $ProcExListBox.Items
            }
            $SvcProcForm.Dispose()

    } # if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){

        # Clear any changes because, 'Cancel'
        if ($SvcExListBox.Items.count -ge 1){$SvcExListBox.Items.Clear()}
        # Revert to default Dirs
        $Services | ForEach-Object {
            $SvcExListBox.Items.Add($_)
        }
        if ($ProcExListBox.Items.count -ge 1){$ProcExListBox.Items.Clear()}
        # Revert files to default
        $Processes | ForEach-Object {
            $ProcExListBox.Items.Add($_)
        }

        [pscustomObject]@{
            Result = 'Cancel'
            ServiceList = $SvcExListBox.Items
            ProcList = $ProcExListBox.Items
        }
        $SvcProcForm.Dispose()
    }

    # Add Zee Members
    Add-Member -InputObject $SvcProcForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name SvcExListBox -Value $SvcExListBox -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ProcExListBox -Value $ProcExListBox -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name AddSvcButton -Value $AddSvcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name AddProcButton -Value $AddProcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name RemSvcButton -Value $RemSvcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name RemProcButton -Value $RemProcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name DefaultProcButton -Value $DefaultProcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name DefaultSvcButton -Value $DefaultSvcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ApplySvcButton -Value $ApplySvcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ApplyProcButton -Value $ApplyProcButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ClearSvcsButton -Value $ClearSvcsButton -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ClearProcsDir -Value $ClearProcsDir -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name NewSvcTextBox -Value $NewSvcTextBox -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name NewProcTextBox -Value $NewProcTextBox -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name SvcExLabel -Value $SvcExLabel -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name ProcExLabel -Value $ProcExLabel -MemberType NoteProperty
    Add-Member -InputObject $SvcProcForm -Name OkButton -Value $OkButton -MemberType NoteProperty
    #Add-Member -InputObject $SvcProcForm -Name CancelButton -Value $CancelButton -MemberType NoteProperty
} # function Edit-WSMServicesProcs {

function Enter-WSMLegalSite {
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $LegalSites
    )

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $SiteForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    [System.Windows.Forms.DataGridView]$SiteDisplay = $null

    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $SiteDisplay = (New-Object -TypeName System.Windows.Forms.DataGridView)
    ([System.ComponentModel.ISupportInitialize]$SiteDisplay).BeginInit()
    $SiteForm.SuspendLayout()

    #region - add icon
    # This region is based on code found on a couple different sites as it falls pretty far outside my normal knowledge base
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $IconImage = [System.Drawing.Image]::FromStream($stream, $true) # Sent to var to silence stream
    #endregion - add icon

    # OkButton
    $OkButton.AccessibleRole = [System.Windows.Forms.AccessibleRole]::None
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]387,[System.Int32]260))
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]0
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true

    # CancelButton
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]296,[System.Int32]260))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]1
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true

    # SiteDisplay
    $SiteDisplay.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $SiteDisplay.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]34,[System.Int32]30))
    $SiteDisplay.MultiSelect = $false
    $SiteDisplay.SelectionMode = 'FullRowSelect'
    $SiteDisplay.Name = [System.String]'SiteDisplay'
    $SiteDisplay.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]428,[System.Int32]209))
    $SiteDisplay.TabIndex = [System.Int32]2

    # Grids
    $SiteDisplay.ColumnCount = 2
    $SiteDisplay.ColumnHeadersVisible = $true
    $SiteDisplay.Columns[0].Name = 'Site'
    $SiteDisplay.Columns[0].Width = 200
    $SiteDisplay.Columns[0].ReadOnly = $true
    $SiteDisplay.Columns[1].Name = 'Location'
    $SiteDisplay.Columns[1].MinimumWidth = 250
    $SiteDisplay.Columns[1].ReadOnly = $true

    # Create the rows for the grid
    $LegalSites | ForEach-Object {
        $SiteDisplay.Rows.Add($_.Site,$_.Location)
    }

    # Form
    $SiteForm.AcceptButton = $OkButton
    $SiteForm.CancelButton = $CancelButton
    $SiteForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]517,[System.Int32]325))
    $SiteForm.Controls.Add($SiteDisplay)
    $SiteForm.Controls.Add($CancelButton)
    $SiteForm.Controls.Add($OkButton)
    $SiteForm.Text = [System.String]'Site Select'
    # Use Icon
    $SiteForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    $result = $SiteForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        [pscustomObject]@{
            Site = $SiteDisplay.SelectedRows.Cells[0].value
            Destination = $SiteDisplay.SelectedRows.Cells[1].value
        }
        $SiteForm.Dispose()
    } 

    ([System.ComponentModel.ISupportInitialize]$SiteDisplay).EndInit()
    $SiteForm.ResumeLayout($false)
    Add-Member -InputObject $SiteForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $SiteForm -Name Button1 -Value $OkButton -MemberType NoteProperty
    Add-Member -InputObject $SiteForm -Name Button2 -Value $CancelButton -MemberType NoteProperty
    Add-Member -InputObject $SiteForm -Name SiteDisplay -Value $SiteDisplay -MemberType NoteProperty
} # function Enter-WSMLegalSite {

function Invoke-WSMDashBoard {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $LegalSites,

        [Parameter(Mandatory=$false)]
        $Processes,

        [Parameter(Mandatory=$false)]
        $Services,
        
        [Parameter(Mandatory=$false)]
        $FileList,
        
        [Parameter(Mandatory=$false)]
        $DirList
    )

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $DashBoardForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Button]$DeviceButton = $null
    [System.Windows.Forms.Button]$ManualDriveButton = $null
    [System.Windows.Forms.Button]$SiteButton = $null
    [System.Windows.Forms.Button]$ExclusionsButton = $null
    [System.Windows.Forms.Button]$ProcButton = $null
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    [System.Windows.Forms.Label]$DeviceLabel = $null
    [System.Windows.Forms.Label]$SiteLabel = $null
    [System.Windows.Forms.Label]$ExLabel = $null
    [System.Windows.Forms.Label]$ProcLabel = $null
    [System.Windows.Forms.DataGridView]$SiteGridView = $null
    [System.Windows.Forms.Label]$DNLabel = $null
    [System.Windows.Forms.Label]$TDLabel = $null
    [System.Windows.Forms.Label]$LegalSSLabel = $null
    [System.Windows.Forms.Label]$Dirlabel = $null
    [System.Windows.Forms.Label]$FileLabel = $null
    [System.Windows.Forms.Label]$SvcLabel = $null
    [System.Windows.Forms.Label]$ProcessLabel = $null
    [System.Windows.Forms.ListBox]$DeviceListBox = $null
    [System.Windows.Forms.ListBox]$DriveListBox = $null
    [System.Windows.Forms.ListBox]$DirListBox = $null
    [System.Windows.Forms.ListBox]$FileListBox = $null
    [System.Windows.Forms.ListBox]$ServiceListBox = $null
    [System.Windows.Forms.ListBox]$ProcListBox = $null

    # This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $DeviceButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ManualDriveButton = (New-Object -TypeName System.Windows.Forms.Button)
    $SiteButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ExclusionsButton = (New-Object -TypeName System.Windows.Forms.Button)
    $ProcButton = (New-Object -TypeName System.Windows.Forms.Button)
    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $DeviceLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $SiteLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $ExLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $ProcLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $SiteGridView = (New-Object -TypeName System.Windows.Forms.DataGridView)
    $DNLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $TDLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $LegalSSLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $Dirlabel = (New-Object -TypeName System.Windows.Forms.Label)
    $FileLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $SvcLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $ProcessLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DeviceListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $DriveListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $DirListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $FileListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $ServiceListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $ProcListBox = (New-Object -TypeName System.Windows.Forms.ListBox)
    ([System.ComponentModel.ISupportInitialize]$SiteGridView).BeginInit()
    $DashBoardForm.SuspendLayout()

    #region - add icon
    # This region is based on code found on a couple different sites as it falls pretty far outside my normal knowledge base
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    #endregion - add icon

    # DeviceButton
    $DeviceButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]44))
    $DeviceButton.Name = [System.String]'DeviceButton'
    $DeviceButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]23))
    $DeviceButton.TabIndex = [System.Int32]0
    $DeviceButton.Text = [System.String]'Target Device'
    $DeviceButton.UseCompatibleTextRendering = $true
    $DeviceButton.UseVisualStyleBackColor = $true
    $DeviceButton.add_Click({
        if ($DvcReturn = Enter-WSMDevice){

            # Hopefully only reset if needed.
            if ($DeviceListBox.Items.count -ge 1){$DeviceListBox.Items.Clear()}
            if ($DriveListBox.Items.count -ge 1){$DriveListBox.Items.Clear()}

            [void] $DeviceListBox.Items.Add($DvcReturn.ComputerName)

            $DvcReturn.Drives | ForEach-Object {
                if ($DriveListBox.Items -notcontains $_){ 
                    [void] $DriveListBox.Items.Add($_)
                }
            } # $DvcReturn.Drives | ForEach-Object {

            if ($SiteReady -and ($DeviceListBox.Items.count -eq 1) -and ($DriveListBox.Items.count -ge 1)){
                    $OkButton.Enabled = $true
            }

        } # if ($DvcReturn = Enter-WSMDevice){
    }) # $DeviceButton.add_Click({

    # ManualDriveButton
    $ManualDriveButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]44))
    $ManualDriveButton.Name = [System.String]'ManualDriveButton'
    $ManualDriveButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]23))
    $ManualDriveButton.TabIndex = [System.Int32]0
    $ManualDriveButton.Text = [System.String]'Offline Device'
    $ManualDriveButton.UseCompatibleTextRendering = $true
    $ManualDriveButton.UseVisualStyleBackColor = $true
    $ManualDriveButton.add_Click({
        if ($DvcReturn = Enter-WSMDeviceManually){

            # Hopefully only reset if needed.
            if ($DeviceListBox.Items.count -ge 1){$DeviceListBox.Items.Clear()}
            if ($DriveListBox.Items.count -ge 1){$DriveListBox.Items.Clear()}

            [void] $DeviceListBox.Items.Add($DvcReturn.ComputerName)

            $DvcReturn.Drives | ForEach-Object {
                if ($DriveListBox.Items -notcontains $_){ 
                    [void] $DriveListBox.Items.Add($_)
                }
            } # $DvcReturn.Drives | ForEach-Object {

            if ($SiteReady -and ($DeviceListBox.Items.count -eq 1) -and ($DriveListBox.Items.count -ge 1)){
                    $OkButton.Enabled = $true
            }

        } # if ($DvcReturn = Enter-Device){
    }) # $ManualDriveButton.add_Click({

    # SiteButton
    $SiteButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]44))
    $SiteButton.Name = [System.String]'SiteButton'
    $SiteButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]23))
    $SiteButton.TabIndex = [System.Int32]1
    $SiteButton.Text = [System.String]'Legal Site'
    $SiteButton.UseCompatibleTextRendering = $true
    $SiteButton.UseVisualStyleBackColor = $true
    $SiteButton.add_Click({
        if ($SiteReturn = (Enter-WSMLegalSite -LegalSites $LegalSites)){
            # Clear previous data
            $SiteGridView.Rows.Clear()

            # Add new selection
            $SiteGridView.Rows.Add($SiteReturn.Site,$SiteReturn.Destination)

            # the list box items appear to be accessible anywhere in the function where the datagrid is a bit harder to target.
            if (($SiteReturn.Destination) -and ($DeviceListBox.Items.count -eq 1) -and ($DriveListBox.Items.count -ge 1)){
                $OkButton.Enabled = $true
            }
            elseif ($SiteReturn.Destination){
                Set-Variable -Scope script -name SiteReady -Value $true
            }

        } # if ($SiteReturn = (Enter-WSMLegalSite -LegalSites $LegalJson)){
    }) # $SiteButton.add_Click({

    # ExclusionsButton
    $ExclusionsButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]217))
    $ExclusionsButton.Name = [System.String]'ExclusionsButton'
    $ExclusionsButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]23))
    $ExclusionsButton.TabIndex = [System.Int32]2
    $ExclusionsButton.Text = [System.String]'Exclusions'
    $ExclusionsButton.UseCompatibleTextRendering = $true
    $ExclusionsButton.UseVisualStyleBackColor = $true
    $ExclusionsButton.add_Click({
        
        if ($ExReturns = Edit-WSMExclusions -FileList $FileList -DirList $DirList){
            if ($FileListBox.Items.count -ge 1){$FileListBox.Items.Clear()}
            # Revert files to default
            $ExReturns.FileExclusions | ForEach-Object {
                $FileListBox.Items.Add($_)
            }

            # Clear any existing entries because
            if ($DirListBox.Items.count -ge 1){$DirListBox.Items.Clear()}
            # Revert to default Dirs
            $ExReturns.DirExclusions | ForEach-Object {
                $DirListBox.Items.Add($_)
            }
        } # if ($ExReturns = Edit-WSMExclusions -FileList $FileList -DirList $DirList){
    }) # $ExclusionsButton.add_Click({...})

    # ProcButton
    $ProcButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]217))
    $ProcButton.Name = [System.String]'ProcButton'
    $ProcButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]23))
    $ProcButton.TabIndex = [System.Int32]3
    $ProcButton.Text = [System.String]'Procs and Svcs'
    $ProcButton.UseCompatibleTextRendering = $true
    $ProcButton.UseVisualStyleBackColor = $true
    $ProcButton.add_Click({
        if ($ProcReturns = Edit-WSMServicesProcs -Processes $ProcList -Services $ServiceList){
            if ($ProcListBox.Items.count -ge 1){$ProcListBox.Items.Clear()}
            # Revert files to default
            $ProcReturns.ProcList | ForEach-Object {
                $ProcListBox.Items.Add($_)
            }

            # Clear any existing entries because
            if ($ServiceListBox.Items.count -ge 1){$ServiceListBox.Items.Clear()}
            # Revert to default Dirs
            $ProcReturns.ServiceList | ForEach-Object {
                $ServiceListBox.Items.Add($_)
            }
        } # if ($ProcReturns = Edit-WSMServicesProcs -Processes $ProcList -Services $ServiceList){
    }) # $ProcButton.add_Click({

    # OkButton
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]544,[System.Int32]522))
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]4
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true
    $OkButton.Enabled = $false

    # CancelButton
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]463,[System.Int32]522))
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]5
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true

    # DeviceLabel
    $DeviceLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]22))
    $DeviceLabel.Name = [System.String]'DeviceLabel'
    $DeviceLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]124,[System.Int32]19))
    $DeviceLabel.TabIndex = [System.Int32]6
    $DeviceLabel.Text = [System.String]'Set Device Information'
    $DeviceLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $DeviceLabel.UseCompatibleTextRendering = $true

    # SiteLabel
    $SiteLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]25))
    $SiteLabel.Name = [System.String]'SiteLabel'
    $SiteLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]133,[System.Int32]16))
    $SiteLabel.TabIndex = [System.Int32]7
    $SiteLabel.Text = [System.String]'Select target Site'
    $SiteLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $SiteLabel.UseCompatibleTextRendering = $true

    # ExLabel
    $ExLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]194))
    $ExLabel.Name = [System.String]'ExLabel'
    $ExLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]257,[System.Int32]20))
    $ExLabel.TabIndex = [System.Int32]8
    $ExLabel.Text = [System.String]'Customize Directories and File Types to Exclude'
    $ExLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $ExLabel.UseCompatibleTextRendering = $true

    # ProcLabel
    $ProcLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]197))
    $ProcLabel.Name = [System.String]'ProcLabel'
    $ProcLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]219,[System.Int32]17))
    $ProcLabel.TabIndex = [System.Int32]9
    $ProcLabel.Text = [System.String]'Customize Processes and Services to Halt'
    $ProcLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $ProcLabel.UseCompatibleTextRendering = $true

    # SiteGridView
    $SiteGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $SiteGridView.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]94))
    $SiteGridView.Name = [System.String]'SiteGridView'
    $SiteGridView.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]219,[System.Int32]97))
    $SiteGridView.TabIndex = [System.Int32]11
    # Grids
    $SiteGridView.ColumnCount = 2
    $SiteGridView.ColumnHeadersVisible = $true
    $SiteGridView.Columns[0].Name = 'Site'
    $SiteGridView.Columns[0].Width = 100
    $SiteGridView.Columns[0].ReadOnly = $true
    $SiteGridView.Columns[1].Name = 'Location'
    $SiteGridView.Columns[1].MinimumWidth = 250
    $SiteGridView.Columns[1].ReadOnly = $true

    # DNLabel
    $DNLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]75))
    $DNLabel.Name = [System.String]'DNLabel'
    $DNLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]73,[System.Int32]16))
    $DNLabel.TabIndex = [System.Int32]18
    $DNLabel.Text = [System.String]'Device Name'
    $DNLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $DNLabel.UseCompatibleTextRendering = $true

    # TDLabel
    $TDLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]68))
    $TDLabel.Name = [System.String]'TDLabel'
    $TDLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
    $TDLabel.TabIndex = [System.Int32]20
    $TDLabel.Text = [System.String]'Target Drives'
    $TDLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $TDLabel.UseCompatibleTextRendering = $true

    # LegalSSLabel
    $LegalSSLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]68))
    $LegalSSLabel.Name = [System.String]'LegalSSLabel'
    $LegalSSLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]133,[System.Int32]23))
    $LegalSSLabel.TabIndex = [System.Int32]21
    $LegalSSLabel.Text = [System.String]'Legal Site Selected'
    $LegalSSLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $LegalSSLabel.UseCompatibleTextRendering = $true

    # Dirlabel
    $Dirlabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]244))
    $Dirlabel.Name = [System.String]'Dirlabel'
    $Dirlabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]119,[System.Int32]19))
    $Dirlabel.TabIndex = [System.Int32]22
    $Dirlabel.Text = [System.String]'Directories Exclusions'
    $Dirlabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $Dirlabel.UseCompatibleTextRendering = $true

    # FileLabel
    $FileLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]244))
    $FileLabel.Name = [System.String]'FileLabel'
    $FileLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]18))
    $FileLabel.TabIndex = [System.Int32]23
    $FileLabel.Text = [System.String]'File Exclusions'
    $FileLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $FileLabel.UseCompatibleTextRendering = $true

    # SvcLabel
    $SvcLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]239))
    $SvcLabel.Name = [System.String]'SvcLabel'
    $SvcLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
    $SvcLabel.TabIndex = [System.Int32]24
    $SvcLabel.Text = [System.String]'Services'
    $SvcLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $SvcLabel.UseCompatibleTextRendering = $true

    # ProcessLabel
    $ProcessLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]524,[System.Int32]239))
    $ProcessLabel.Name = [System.String]'ProcessLabel'
    $ProcessLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
    $ProcessLabel.TabIndex = [System.Int32]25
    $ProcessLabel.Text = [System.String]'Processes'
    $ProcessLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $ProcessLabel.UseCompatibleTextRendering = $true

    # DeviceListBox
    $DeviceListBox.FormattingEnabled = $true
    $DeviceListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]21,[System.Int32]96))
    $DeviceListBox.Name = [System.String]'DeviceListBox'
    $DeviceListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]96,[System.Int32]95))
    $DeviceListBox.TabIndex = [System.Int32]28

    # DriveListBox
    $DriveListBox.FormattingEnabled = $true
    $DriveListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]96))
    $DriveListBox.Name = [System.String]'DriveListBox'
    $DriveListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]95))
    $DriveListBox.TabIndex = [System.Int32]29

    # DirListBox
    $DirListBox.FormattingEnabled = $true
    $DirListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]266))
    $DirListBox.Name = [System.String]'DirListBox'
    $DirListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]212))
    $DirListBox.TabIndex = [System.Int32]30

    # FileListBox
    $FileListBox.FormattingEnabled = $true
    $FileListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]147,[System.Int32]266))
    $FileListBox.Name = [System.String]'FileListBox'
    $FileListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]212))
    $FileListBox.TabIndex = [System.Int32]31

    # ServiceListBox
    $ServiceListBox.FormattingEnabled = $true
    $ServiceListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]400,[System.Int32]266))
    $ServiceListBox.Name = [System.String]'ServiceListBox'
    $ServiceListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]212))
    $ServiceListBox.TabIndex = [System.Int32]32

    # ProcListBox
    $ProcListBox.FormattingEnabled = $true
    $ProcListBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]524,[System.Int32]266))
    $ProcListBox.Name = [System.String]'ProcListBox'
    $ProcListBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]95,[System.Int32]212))
    $ProcListBox.TabIndex = [System.Int32]33

    # DashBoardForm
    $DashBoardForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]645,[System.Int32]557))
    $DashBoardForm.Controls.Add($ProcListBox)
    $DashBoardForm.Controls.Add($ServiceListBox)
    $DashBoardForm.Controls.Add($FileListBox)
    $DashBoardForm.Controls.Add($DirListBox)
    $DashBoardForm.Controls.Add($DriveListBox)
    $DashBoardForm.Controls.Add($DeviceListBox)
    $DashBoardForm.Controls.Add($ProcessLabel)
    $DashBoardForm.Controls.Add($SvcLabel)
    $DashBoardForm.Controls.Add($FileLabel)
    $DashBoardForm.Controls.Add($Dirlabel)
    $DashBoardForm.Controls.Add($LegalSSLabel)
    $DashBoardForm.Controls.Add($TDLabel)
    $DashBoardForm.Controls.Add($DNLabel)
    $DashBoardForm.Controls.Add($SiteGridView)
    $DashBoardForm.Controls.Add($ProcLabel)
    $DashBoardForm.Controls.Add($ExLabel)
    $DashBoardForm.Controls.Add($SiteLabel)
    $DashBoardForm.Controls.Add($DeviceLabel)
    $DashBoardForm.Controls.Add($CancelButton)
    $DashBoardForm.Controls.Add($OkButton)
    $DashBoardForm.Controls.Add($ProcButton)
    $DashBoardForm.Controls.Add($ExclusionsButton)
    $DashBoardForm.Controls.Add($SiteButton)
    $DashBoardForm.Controls.Add($DeviceButton)
    $DashBoardForm.Controls.Add($ManualDriveButton)

    # Use Icon
    $DashBoardForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
    $DashBoardForm.Text = [System.String]'Legal Discovery Device Targeter'
    ([System.ComponentModel.ISupportInitialize]$SiteGridView).EndInit()
    $DashBoardForm.ResumeLayout($false)

    $Services | ForEach-Object {
        $ServiceListBox.Items.Add($_)
    }
    $Processes | ForEach-Object {
        $ProcListBox.Items.Add($_)
    }
    $DirList | ForEach-Object {
        $DirListBox.Items.Add($_)
    }
    $FileList | ForEach-Object {
        $FileListBox.Items.Add($_)
    }

    $result = $DashBoardForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        [pscustomObject]@{
            Result = 'Success'
            Site = $SiteGridView.Rows[0].Cells[0].Value
            Destination = $SiteGridView.Rows[0].Cells[1].Value
            ComputerName = $DeviceListBox.Items # should only ever be one anyway...
            Drives = $DriveListBox.Items
            ServiceList = $ServiceListBox.Items
            ProcList = $ProcListBox.Items
            DirExclusions = $DirListBox.Items
            FileExclusions = $FileListBox.Items
        }
        $DashBoardForm.Dispose()
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        [pscustomObject]@{
            Result = 'Cancel'
            Site = $SiteGridView.Rows[0].Cells[0].Value
            Destination = $SiteGridView.Rows[0].Cells[1].Value
            ComputerName = $DeviceListBox.Items # should only ever be one anyway...
            Drives = $DriveListBox.Items
            ServiceList = $ServiceListBox.Items
            ProcList = $ProcListBox.Items
            DirExclusions = $DirListBox.Items
            FileExclusions = $FileListBox.Items
        }
        $DashBoardForm.Dispose()
    }

    Add-Member -InputObject $DashBoardForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DeviceButton -Value $DeviceButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ManualDriveButton -Value $ManualDriveButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name SiteButton -Value $SiteButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ExclusionsButton -Value $ExclusionsButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ProcButton -Value $ProcButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name OkButton -Value $OkButton -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DeviceLabel -Value $DeviceLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name SiteLabel -Value $SiteLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ExLabel -Value $ExLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ProcLabel -Value $ProcLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name SiteGridView -Value $SiteGridView -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DNLabel -Value $DNLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name TDLabel -Value $TDLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name LegalSSLabel -Value $LegalSSLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name Dirlabel -Value $Dirlabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name FileLabel -Value $FileLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name SvcLabel -Value $SvcLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ProcessLabel -Value $ProcessLabel -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DeviceListBox -Value $DeviceListBox -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DriveListBox -Value $DriveListBox -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name DirListBox -Value $DirListBox -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name FileListBox -Value $FileListBox -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ServiceListBox -Value $ServiceListBox -MemberType NoteProperty
    Add-Member -InputObject $DashBoardForm -Name ProcListBox -Value $ProcListBox -MemberType NoteProperty
} # function Invoke-WSMDashBoard {

function Remove-WSMScheduledTask {
    <#
        .SYNOPSIS
            Finds and removes scheduled tasks.
        .DESCRIPTION
            Searches for and removes scheduled tasks with some error trapping. 
        .PARAMETER TaskName
            Name of the task to search for
        .PARAMETER Logfile
            The log file being used, if applicable
        .EXAMPLE
            Remove-WSMScheduledTask -Taskname 'LegalDisc' -LogFile $LogFile
        .INPUTS
            System.String
        .OUTPUTS
            Logs
        .Notes
            Last Updated:
    
            ========== HISTORY ==========
            Author: Kevin Van Bogart
            Created: 2020-01-15 15:32:44Z
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$TaskName,

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

function Search-WSMUSBPorts {
    # This was snagged from Technet and reformatted as it was sort of a PITA
    # https://social.technet.microsoft.com/Forums/windowsserver/en-US/09c9814a-38fa-4b16-bc8f-01329882a791/powershell-wmi-get-usb-storage-devices-only?forum=winserverpowershell
    $USBDrives = Get-CimInstance -class win32_diskdrive -ErrorAction SilentlyContinue | Where-Object {$_.InterfaceType -eq "USB"}
    $DriveLetters = $USBDrives | ForEach-Object {
        Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition" -ErrorAction SilentlyContinue
    } |  ForEach-Object {
            Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition" -ErrorAction SilentlyContinue
         } | ForEach-Object {
                $_.deviceid
            }
        
    # return
    Get-CimInstance Win32_Volume -ErrorAction SilentlyContinue | Where-Object {($DriveLetters -contains ($_.Name -replace '\\')) -and ($_.Label -notmatch '^Port|Replicator')} #| 
        #Select-Object Name,FreeSpace,Capacity
} # function Search-WSMUSBPorts {...}

function Select-WSMSaveLocation {
    $USBSelectForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.DataGridView]$USBDisplayGrid = $null
    [System.Windows.Forms.Button]$RefreshButton = $null
    [System.Windows.Forms.Button]$FormatButton = $null
    [System.Windows.Forms.Button]$OkButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null
    [System.Windows.Forms.Label]$GridViewLabel = $null
    [System.Windows.Forms.Label]$FormatLabel = $null
    [System.ComponentModel.IContainer]$components = $null
    #This was added for the Icon
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #region - add icon
        # This region is based on code found on a couple different sites as it falls pretty far outside my normal knowledge base
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length)
        $IconImage = [System.Drawing.Image]::FromStream($stream, $true) # Sent to var to silence stream
    #endregion - add icon

    $components = (New-Object -TypeName System.ComponentModel.Container)
    $USBDisplayGrid = (New-Object -TypeName System.Windows.Forms.DataGridView)
    $RefreshButton = (New-Object -TypeName System.Windows.Forms.Button)
    $FormatButton = (New-Object -TypeName System.Windows.Forms.Button)
    $OkButton = (New-Object -TypeName System.Windows.Forms.Button)
    $CancelButton = (New-Object -TypeName System.Windows.Forms.Button)
    $GridViewLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $FormatLabel = (New-Object -TypeName System.Windows.Forms.Label)
    ([System.ComponentModel.ISupportInitialize]$USBDisplayGrid).BeginInit()
    $USBSelectForm.SuspendLayout()

    # USBDisplayGrid
    $USBDisplayGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $USBDisplayGrid.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]48,[System.Int32]40))
    $USBDisplayGrid.Name = [System.String]'USBDisplayGrid'
    $USBDisplayGrid.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]349,[System.Int32]150))
    $USBDisplayGrid.TabIndex = [System.Int32]0
    $USBDisplayGrid.SelectionMode = 'FullRowSelect'
    # Grids
    $USBDisplayGrid.ColumnCount = 3
    $USBDisplayGrid.ColumnHeadersVisible = $true
    $USBDisplayGrid.Columns[0].Name = 'Name'
    $USBDisplayGrid.Columns[0].Width = 200
    $USBDisplayGrid.Columns[0].ReadOnly = $true
    $USBDisplayGrid.Columns[1].Name = 'Freespace'
    $USBDisplayGrid.Columns[1].MinimumWidth = 200
    $USBDisplayGrid.Columns[1].ReadOnly = $true
    $USBDisplayGrid.Columns[2].Name = 'Capacity'
    $USBDisplayGrid.Columns[2].MinimumWidth = 200
    $USBDisplayGrid.Columns[2].ReadOnly = $true

    # RefreshButton
    $RefreshButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]322,[System.Int32]196))
    $RefreshButton.Name = [System.String]'RefreshButton'
    $RefreshButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $RefreshButton.TabIndex = [System.Int32]1
    $RefreshButton.Text = [System.String]'Refresh'
    $RefreshButton.UseCompatibleTextRendering = $true
    $RefreshButton.UseVisualStyleBackColor = $true
    $RefreshButton.add_Click({
        # Clear the Grid
        $USBDisplayGrid.Rows.Clear()

        # Add whatever is found.
        Search-WSMUSBPorts | Foreach-Object {
            $FreeSpace = "$([math]::Round($_.FreeSpace/1GB,3)) GB"
            $Capacity = "$([math]::Round($_.Capacity/1GB,3)) GB"
            $USBDisplayGrid.Rows.Add($_.Name,$FreeSpace,$Capacity)
        }
    })

    # FormatButton
    $FormatButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]48,[System.Int32]196))
    $FormatButton.Name = [System.String]'FormatButton'
    $FormatButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $FormatButton.TabIndex = [System.Int32]3
    $FormatButton.Text = [System.String]'Format'
    $FormatButton.UseCompatibleTextRendering = $true
    $FormatButton.UseVisualStyleBackColor = $true
    $FormatButton.add_Click({
        #Add warning and then format the drive.

        if (($USBDisplayGrid.SelectedRows.Cells[0]).value -notin '',$null){
            
            # Clean the drive letter for use
            $DriveLetter = (($USBDisplayGrid.SelectedRows.Cells[0]).value) -replace '\\','' -replace ':',''

            # Format drive
            if ($DriveLetter -match 'C'){
                #BAD! NO FORMATTY FOR YOU!!!
                $null = Send-WSMMessage -Message "Error: This utility will not format drive $(($USBDisplayGrid.SelectedRows.Cells[0]).value)!" -Title 'Format Error!' -Icon Warning
            }
            else {
                $WarningMessage = "*****WARNING*****`r`n`rThis command will reformat the selected USB drive!`r`n`r$(($USBDisplayGrid.SelectedRows.Cells[0]).value)`r`n`rDouble-check this is the drive you intend to format!`r`n`rData recovery may not be possible if the wrong drive is formatted!!!"
                $PreFormatMessage = Send-WSMMessage -Message $WarningMessage -Title 'Format USB' -Icon Warning

                if ($PreFormatMessage -eq 'Ok'){
                    try {
                        $null = Format-Volume -DriveLetter $DriveLetter -FileSystem exFAT -Force -ErrorAction Stop
                        $null = Send-WSMMessage -Message "Successfully formatted drive `'$(($USBDisplayGrid.SelectedRows.Cells[0]).value)`'" -Title 'Format USB'
                    }
                    catch {
                        $null = Send-WSMMessage -Message "Failed to format drive `'$(($USBDisplayGrid.SelectedRows.Cells[0]).value)`'" -Title 'Format Error!' -Icon Warning
                    }
                }
                else {
                    $null = Send-WSMMessage -Message "Cancelled formatting drive `'$(($USBDisplayGrid.SelectedRows.Cells[0]).value)`'" -Title 'Format USB' -Icon Warning
                }
            }
        }
        else {
            $null = Send-WSMMessage -Message "Error: No USB drive was selected!" -Title 'Format Error!' -Icon Warning
        }
    }) # $FormatButton.add_Click({

    # OkButton
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OkButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]362,[System.Int32]265))
    $OkButton.Name = [System.String]'OkButton'
    $OkButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $OkButton.TabIndex = [System.Int32]4
    $OkButton.Text = [System.String]'Ok'
    $OkButton.UseCompatibleTextRendering = $true
    $OkButton.UseVisualStyleBackColor = $true

    # CancelButton
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]281,[System.Int32]265))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
    $CancelButton.TabIndex = [System.Int32]5
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.UseCompatibleTextRendering = $true
    $CancelButton.UseVisualStyleBackColor = $true
    $CancelButton.add_Click($Button1_Click)

    # GridViewLabel
    $GridViewLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]48,[System.Int32]10))
    $GridViewLabel.Name = [System.String]'GridViewLabel'
    $GridViewLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]349,[System.Int32]23))
    $GridViewLabel.TabIndex = [System.Int32]6
    $GridViewLabel.Text = [System.String]'Use refresh button to detect newly inserted drives. '
    $GridViewLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $GridViewLabel.UseCompatibleTextRendering = $true

    # FormatLabel
    $FormatLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]129,[System.Int32]196))
    $FormatLabel.Name = [System.String]'FormatLabel'
    $FormatLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]187,[System.Int32]23))
    $FormatLabel.TabIndex = [System.Int32]7
    $FormatLabel.Text = [System.String]'Format USB drive better results.'
    $FormatLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $FormatLabel.UseCompatibleTextRendering = $true

    # USBSelectForm
    $USBSelectForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]453,[System.Int32]300))
    $USBSelectForm.Controls.Add($FormatLabel)
    $USBSelectForm.Controls.Add($GridViewLabel)
    $USBSelectForm.Controls.Add($CancelButton)
    $USBSelectForm.Controls.Add($OkButton)
    $USBSelectForm.Controls.Add($FormatButton)
    $USBSelectForm.Controls.Add($RefreshButton)
    $USBSelectForm.Controls.Add($USBDisplayGrid)
    $USBSelectForm.Name = [System.String]'USBSelectForm'
    $USBSelectForm.Text = [System.String]'Select USB Drive'
    ([System.ComponentModel.ISupportInitialize]$USBDisplayGrid).EndInit()
    $USBSelectForm.ResumeLayout($false)
    $USBSelectForm.PerformLayout()
    # Use Icon
    $USBSelectForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    # Add whatever is found.
    Search-WSMUSBPorts | Foreach-Object {
        $FreeSpace = "$([math]::Round($_.FreeSpace/1GB,3)) GB"
        $Capacity = "$([math]::Round($_.Capacity/1GB,3)) GB"
        $USBDisplayGrid.Rows.Add($_.Name,$FreeSpace,$Capacity)
    }

    # Run the dialog
    $Result = $USBSelectForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        [pscustomObject]@{
            Result = 'Success'
            Path = ($USBDisplayGrid.SelectedRows.Cells[0]).value
        }
        $USBSelectForm.Dispose()
    } # if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        [pscustomObject]@{
            Result = 'Cancel'
            Path = 'NOT A DRIVE!!!'
        }
        $USBSelectForm.Dispose()
    }

    Add-Member -InputObject $USBSelectForm -Name base -Value $base -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name USBDisplayGrid -Value $USBDisplayGrid -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name RefreshButton -Value $RefreshButton -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name FormatButton -Value $FormatButton -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name OkButton -Value $OkButton -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name GridViewLabel -Value $GridViewLabel -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name FormatLabel -Value $FormatLabel -MemberType NoteProperty
    Add-Member -InputObject $USBSelectForm -Name components -Value $components -MemberType NoteProperty
} # function Select-WSMSaveLocation {

function Send-WSMMessage {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Title = 'Info',
        [ValidateSet('Stop', 'Warning', 'Question','None','Information','Hand','Exclamation','Error','Asterisk')]
        [string]$Icon = 'None'
    )
    process {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

        [System.Windows.Forms.MessageBox]::show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::$Icon
        )
    }
} # function Send-WSMMessage {

function Send-WSMTSMessageBox { 
    <# Send-WSMTSMessageBox  
    .SYNOPSIS    
        Send a message or prompt to the interactive user in any session with the ability to get the results. 
       
    .DESCRIPTION    
        Allows the administrator to send a message / prompt to an interactive user.  
       
    .EXAMPLE    
        "Send a message immediately w/o waiting for a responce." 
        Send-WSMTSMessageBox -Title "Email Problem" -Message "We are currently having delays and are working on the issue." 
         
        "Send a message waiting 60 seconds for a reponse of [Yes / No]." 
        $Result = Send-WSMTSMessageBox -Title "System Updated" -Message "System requires a reboot. Would you like to the reboot system now?" ` 
         -ButtonSet 4 -Timeout 60 -WaitResponse $true  
          
        ButtonSets 
        0 = OK 
        1 = Ok/Cancel 
        2 = Abort/Retry/Ignore 
        3 = Yes/No/Cancel 
        4 = Yes/No 
        5 = Retry/Cancel 
        6 = Cancel/Try Again/Continue     
         
    .OUTPUTS
        "" = 0 
        "Ok" = 1 
        "Cancel" = 2   
        "Abort" = 3 
        "Retry" = 4     
        "Ignore" = 5 
        "Yes" = 6 
        "No" = 7 
        "Try Again" = 10 
        "Continue" = 11 
        "Timed out" = 32000 
        "Not set to wait" = 32001  
    
    .NOTES    
        Author: Raymond H Clark 
        Twitter: @Rowdybullgaming 
     
        http://technet.microsoft.com/en-us/query/aa383488  
        http://technet.microsoft.com/en-us/query/aa383842 
        http://pinvoke.net/default.aspx/wtsapi32.WTSSendMessage 
    #>  
    param(
        [string]$Title = "Title", 
        [string]$Message = "Message", 
        [int]$ButtonSet = 0, 
        [int]$Timeout = 0, 
        [bool]$WaitResponse = $false
    )

$Signature = @"
[DllImport("wtsapi32.dll", SetLastError = true)] 
public static extern bool WTSSendMessage( 
    IntPtr hServer, 
    [MarshalAs(UnmanagedType.I4)] int SessionId, 
    String pTitle, 
    [MarshalAs(UnmanagedType.U4)] int TitleLength, 
    String pMessage, 
    [MarshalAs(UnmanagedType.U4)] int MessageLength, 
    [MarshalAs(UnmanagedType.U4)] int Style, 
    [MarshalAs(UnmanagedType.U4)] int Timeout, 
    [MarshalAs(UnmanagedType.U4)] out int pResponse, 
    bool bWait); 
     
    [DllImport("kernel32.dll")] 
    public static extern uint WTSGetActiveConsoleSessionId(); 
"@ 
    [int]$TitleLength = $Title.Length; 
    [int]$MessageLength = $Message.Length; 
    [int]$Response = 0; 
                             
    $MessageBox = Add-Type -memberDefinition $Signature -name "WTSAPISendMessage" -namespace "WTSAPI" -passThru    
    $SessionId = $MessageBox::WTSGetActiveConsoleSessionId() 
     
    $MessageBox::WTSSendMessage(0, $SessionId, $Title, $TitleLength, $Message, $MessageLength, $ButtonSet, $Timeout, [ref] $Response, $WaitResponse) | Out-Null 
     
    $Response 
} # function Send-WSMTSMessageBox {

#=================================================================
#endregion - Define PUBLIC Advanced functions
#=================================================================
#=================================================================
#region - Define PRIVATE Advanced functions
#=================================================================
#=================================================================
#endregion - Define PRIVATE Advanced functions
#=================================================================
#=================================================================
#region - Export Modules
#=================================================================
Export-ModuleMember -function *-WSM*
#=================================================================
#endregion - Export Modules
#=================================================================
# SIG # Begin signature block
# MIIcigYJKoZIhvcNAQcCoIIcezCCHHcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHhqoGmR2lc0I60ikinKLOL4v
# yJ+gghe5MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# MRYEFBupJ+35Z9eS5hQS3SV8Qh+LQ4O8MA0GCSqGSIb3DQEBAQUABIIBAGG6BRPX
# Q4eXbLoA8xplQV4HSIWqy6sogIvvWxEXYm6SUB8YCt5L0xDl25ADVyFlxT1qrjvx
# nr/7zunbiRcMbfCtqONQjSgbRdK2MxeSqfzZ7uuUv6IlRp9Y2beeG2VK1d0Sk/rM
# 5m4eO3NjAPboyfkgUlVVwkiv502uklKMJJA4XqKh70NhN0AkcXdIbzRP/otvMGA8
# 3sVIO/LYBcjm/817FSQ6IhRjCzFFplDUUEJhxN8cRVZfBivX6cjLZM236Celg15/
# d4ISfxB1CtKs5XVp0jOkrZNN1Gsp5X3SGHdWYr3M4Sam6DNAfcKXJ9b+0hoasgaF
# KFHr+DRU74Q6O4GhggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEC
# EAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIwMDIxMzE4NDcxNVowIwYJKoZIhvcN
# AQkEMRYEFAmA71lClVEzb7qFrjC1Shm6CUMrMA0GCSqGSIb3DQEBAQUABIIBAKLX
# bXQtny8P8k5NYMnmJ0IvNkQKeC80bc9LP1aPYB4+CqKZMq0HQlCniFL6nyY9DR7t
# YIrcm/FghxCfGMwDF6z8AiKJR6eeb1MboMmRUc1HIchGQhOQtQKXbVLaUhlpxcJ6
# Eh4T6dDzrutxlTUK1hycHIbZemSSbA2kcaoOSZzkdGQsnDDczgmQbZcEJMh+hDHQ
# RY02+K21oOi0uWThqfbiurqrTPctAOsqEem5CIi49ZRtXlvsOx3IfVqX5YxglAXI
# GtITNMA5rxjSAACzdJisOfVoi7PXUEvM/4mROYrjO2gHqnq8gRA/HwfVRr4GKVKC
# JxhgUgomaXk3/IVOccg=
# SIG # End signature block
