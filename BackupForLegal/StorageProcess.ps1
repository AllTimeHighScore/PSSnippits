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
    [ValidateScript({(Get-item -path $_ -ErrorAction SilentlyContinue).Extension -eq '.JSON'})]
    [string]$LegalSitesFile = '\\stpnas08\legalinfo$\LegalSite.json',

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

    $Destination = '\\stpnas07\dld$\Profiles'
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
# SIG # Begin signature block
# MIIcigYJKoZIhvcNAQcCoIIcezCCHHcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTYY3x0j6Ys8j/vh5Iff+R+W4
# oPKgghe5MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# MRYEFIA0/KDaHLnrvkxa8cDGkrIjNB2DMA0GCSqGSIb3DQEBAQUABIIBAFsSfUQP
# JBV7SNZZDKrLjpp3vRUgLcrC/2W2U8ue+1Yro3ZJpVFA+ebW3567PwS3p6lB0Tgj
# LvWD/kNmemIQkM4RmXgbWvGuxfLyn/PSU2i0T7QYK8hZaF1WuFrJIwtu4jdyJr8Z
# iXQHZXJ2OwxR/BRFjk/M9Y7LlvbrC5IeOJcSrEYJBLUJHO0n+zcBCMa9vapHxlQZ
# OgIVikWUMUC9uVJ/AuHW6dn+IjyYqd4XMnIG/OVKP/dzR8qCVt4wwz8eh7mucbt1
# Vw/bkXYtSzBwPvemMIVba4qw/bRtqtEdGi0Qxtoz+NLt1nHH2cC++nw/uSZqjXM+
# 4rzx7pg5jHaKakyhggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEC
# EAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIwMDIxMzIxMDAyNVowIwYJKoZIhvcN
# AQkEMRYEFNaT6bln30eel8ErMEhT7+6ehAViMA0GCSqGSIb3DQEBAQUABIIBADOX
# g0ifnS05K5auUuilMngAtsWN73o+Hamjicdk0CmrruBYbwLWwP+e2x6rOIxNZ8KB
# OSKbM0/k49p0KKOjjqx8vpygxGoP3lFF8sZr5ctLQkbmZe1bPAX1aMGpieZGxXzh
# 5+Dfn0tPlGrcNtIO0drwSV3mmjmFTf6giKaeSuDA2WoGcfYU/eq3DfjMIInftP4e
# lAMcb5MY//PnhAycoUvjmWSccpFCQrAeYeiBrYfB4lunA7Z/jhyIraU7rU+wdLsd
# F6pTcpGbSAFl+NYqOmzsPpztFPdYd/lGsNmLQ+i12yuursDXi9Lz+bkXhEvQosOO
# CRmgaXM5BNPO+SboZ1M=
# SIG # End signature block
