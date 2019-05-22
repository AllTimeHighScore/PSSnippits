<#
.SYNOPSIS
    Use the WMI Process run method to run MS Patch
.DESCRIPTION
    Use the WMI Process run method to get around Application Control and patch vulnerabilities
.EXAMPLE
        PS CM2:\> powershell.exe -executionpolicy bypass -file "C:\Users\<SUPERADMINPERSON>\Desktop\Git\PatchWithoutPS_1.0_PGK_2003XP\PatchWithoutPS_1.0_PGK_2003XP.ps1" -verbose
.INPUTS
    string[]
.OUTPUTS
    LogFile
.NOTES
    ========== HISTORY ==========
    Author: Van Bogart, Kevin
    Created: 2019-05-17 14:04:47Z

    Created for RDP vulnerability <KB4500331>
    Created by altering an older script of mine that did this via Scheduled Tasks

    ***You must have access to the target devices to run this***

    If I weren't in a pinch I would clean up the output a bit more. 
    There is a lot of logging that is not needed post testing. 
    The PS Custom Object should be the only thing needed.

    Utilizes a Run as Jobs wrapper a coworker <Shawn Leehane> built
        Can target a large amount for devices in short order. 
        Since the patch is small it shouldn't cause too much of a bandwidth issue.
        If this is used for parger patches you might want to throttle the jobs a bit.
#>


# Create an output var to store completed job data
$Output = @()

# Creatte a Run ID for this execution space
$JobName = ([guid]::newGuid().Guid).Split('-')[0]

# Set a log file here, if we want one
$LogFile = $null

# Set the max number of concurrent jobs
$MaxConcurrentJobs = 10

# Set the max number of hours we want each job to run before we give up and kill it
#$MaxRunHours = 120

# Set the max number of minutes we want each job to run before we give up and kill it
$MaxRunMinutes = 30

#logFile
$PostOutFile = "$PSScriptRoot\kb4500331_Patch_Install_Results_$((Get-Date).ToString("yyyymmddhhmmss")).log"

#Json - For easier parcing
$PostOutJson = "$PSScriptRoot\kb4500331_Patch_Install_Results_$((Get-Date).ToString("yyyymmddhhmmss")).JSON"

#Device list
$TargetSource = "$PSScriptRoot\TargetDevices.txt"

$PCList = Get-content -Path $TargetSource

#Other Vars
$Media = "$PSScriptRoot\Media"
$PauseDuration = 20 # for future use, mainly if updates are larger

#region - Log header - remove if you don't need it
@"
"****************************************" 
"Running Script $($MyInvocation.mycommand.name) $((Get-Date).ToString("MM/dd/yyyy"))" 
"Time $((Get-Date).ToString("HH:mm:ss"))"
"Computer name = $($env:COMPUTERNAME)"
"Invoking User = $($env:USERNAME)"
"****************************************" 
"@ -split "`n" | Out-File -FilePath $PostOutFile -Append -Force
#endregion - Log header - remove if you don't need it

#Loops through items passed to the 'process' parameter
$PCList | ForEach-Object {
    
    Write-Verbose -Message "Working on Machine: $_"
    "$Section Working on Machine: $_" | Out-File -FilePath $ResultFile -Append -Force

    #ScriptBlock is the script that will be triggered by the job associated with the process. 
    #This is referenced in the "ScriptBlock" parameter of the Start-Job cmdlet
    $ScriptBlock = {

        ########### For Testing for more advanced console output. ############
        #$VerbosePreference='continue'

        # Assign $_ to a var with the "using" variable so we can pull external variables into this scope
        $Comp = $using:_
        $Copy = $using:Media
        #$ResultFile = $using:Logfile
        $ResultFile = $using:PostOutFile
        $Seconds2Wait = $using:PauseDuration

        $OSName = "Unknown"
        $KBInspection = $false
        $UpdateComplete = 'Unverified'
        $OkayToProceed = $null

        if (Test-Connection -ComputerName $Comp -Count 3 -ErrorAction SilentlyContinue){

            #region - Gather system info

            try {
                Write-Verbose -Message "$($Comp): Attempting to establish WMI connection"
                #The KB shouldn't be there yet. I just want to give WMI a few chances to connect
                foreach ($try in 1..10){
                    if ($KBInspection = (Get-WmiObject -ComputerName $Comp -Class Win32_QuickFixEngineering -ea Stop | Where-Object {$_.hotfixid -match 'kb4500331'})){
                        break
                    }
                    Start-Sleep -Seconds 2
                }

                $WMIConnects = $true
                $Access = $true

                Write-Verbose -message "$($Comp): Checking WMI for OS type"
                #rely on the same try block

                foreach ($try in 1..10){                
                    if ($OSData = Get-WmiObject -ComputerName $Comp -Class Win32_OperatingSystem -ea Stop){
                        $OSName = $OSData.Caption
                        break
                    }
                    Start-Sleep -Seconds 2
                }

                if (!$OSName){
                    Write-Verbose -message "$($Comp): Checking Sysinfo"
                    #Sometimes the WMI return didn't parse in testing.
                    #If access is denied Sysinfo will hang PS at an invisible login message, The try catch for WMI should guide us around that.
                    $OSLines = systeminfo /S $Comp
                    #Sysinfo sometimes throws odd errors. I hope we've now seen them all...
                    if ($OSLines -match "Error|Invalid"){
                        #Try to get the info from Reg then.
                        Write-Verbose -Message "$($Comp): Sysinfo had an error - checking reg keys."
                        $OSName = (& reg query "\\$Comp\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2>&1)

                        if ($OSName -match "Access is Denied"){
                            "$Comp is not allowing access and cannot be managed by this script (or at least this user)." | Out-File -FilePath $ResultFile -Append -Force
                            $Access = $false
                        }
                    }
                    else {
                        $OSName = ($OSLines -match "OS Name:").split(':').trim()[1]
                    }
                }

                if (($OSLines -match 'kb4500331') -or $KBInspection){
                    Write-verbose -message "$($Comp): Compliant - The patch is already installed on this device" #| Out-File -FilePath $ResultFile -Append -Force
                    $UpdateComplete = 'Verified'
                }

            }
            catch [System.UnauthorizedAccessException]{
                "$($Comp): Access is denied" | Out-File -FilePath $ResultFile -Append -Force
                $Access = $false
            }
            catch {
                "$($Comp): WMI May be corrupt, expect other issues. $($_.Exception.GetType().FullName) ^ $($_.exception.message)" | Out-File -FilePath $ResultFile -Append -Force
                $WMIConnects = $false
            }
            #endregion - Gather system info

            #region - Check if OS is okay to be patched
            #"WMIConnected = $WMIConnects" | Out-File -FilePath $ResultFile -Append -Force
            #"$Comp is $OSName" | Out-File -FilePath $ResultFile -Append -Force
            if ($OSName -notmatch "Windows XP|2003"){
                #Log Failure
                "$($Comp): Device is not the target OS ($OSName), the script will not run on this box." | Out-File -FilePath $ResultFile -Append -Force
                $OkayToProceed = $false
            }
            elseif (($OSLines -match "RPC") -or ($OSName -match "ERROR: Invalid syntax")){
                $OkayToProceed = $true
                Write-verbose -message "$($Comp): Device may be in a bad state. Will attempt update regardless. (Don't expect much.)"
                "$($Comp): Device may be in a bad state. Will attempt update regardless. (Don't expect much.)" | Out-File -FilePath $ResultFile -Append -Force
            }
            else {
                Write-verbose -Message "$($Comp): Proper OS Detected ($OSName). Proceeding" #| Out-File -FilePath $ResultFile -Append -Force
                $OkayToProceed = $true
            }
            #region - Check if OS is okay to be patched

            #If the update hasn't been found already
            #"UpdateComplete: $UpdateComplete" | Out-File -FilePath $ResultFile -Append -Force
            #"OkayToProceed: $OkayToProceed" | Out-File -FilePath $ResultFile -Append -Force
            #"Access: $Access" | Out-File -FilePath $ResultFile -Append -Force

            if (($UpdateComplete -eq 'Unverified') -and ($OkayToProceed -eq $true) -and ($Access -eq $true)){
                Write-verbose -message "$($Comp): Passed three required Checks"

                $RunDate = (Get-Date).ToString("ddMMyyyymmss")
                if ($OSName -match "(Server 2003|XP).*x64"){
                    [version]$Version = '5.2.3790.6787'
                    $File2Copy = 'windowsserver2003-kb4500331-x64-custom-enu_e2fd240c402134839cfa22227b11a5ec80ddafcf.exe'
                    $Task = "C:\Windows\Temp\KB4500331\$File2Copy /quiet /norestart /log:%windir%\TEMP\kb4500331_Install_$RunDate.log"
                }
                elseif ($OSName -match "Server 2003"){
                    [version]$Version = '5.2.3790.6787'
                    $File2Copy = 'windowsserver2003-kb4500331-x86-custom-enu_62d416d73d413b590df86224b32a52e56087d4c0.exe'
                    $Task = "C:\Windows\Temp\KB4500331\$File2Copy /quiet /norestart /log:%windir%\TEMP\kb4500331_Install_$RunDate.log"
                }
                elseif ($OSName -match 'Windows XP'){
                    [version]$Version = '5.2.3790.6787'
                    $File2Copy = 'windowsxp-kb4500331-x86-custom-enu_d7206aca53552fececf72a3dee93eb2da0421188.exe'
                    $Task = "C:\Windows\Temp\KB4500331\$File2Copy /quiet /norestart /log:%windir%\TEMP\kb4500331_Install_$RunDate.log"
                }

                #Start attempting to drop the task

                #region - Copy files to local device

                    <#
                        #################################
                        #
                        # - Copy the files to the target device
                        # - Something in the copy section is still pumping out to the log file.
                        #       I have mixed feeling about that.
                        #
                        #################################
                    #>

                try {
                    $FileCopied = $true
                    Write-Verbose -Message "$($Comp): Attempting to copy files: $("$Copy\$File2Copy")" #| Out-File -FilePath $ResultFile -Append -Force
                    #[string](& robocopy $Copy "\\$Comp\C$\Windows\Temp\KB4500331" $File2Copy /S /MT:50 /r:30 /w:05 2>&1) | Out-String -OutVariable CopyFiles
                    $CopyFiles = (& robocopy $Copy "\\$Comp\C$\Windows\Temp\KB4500331" $File2Copy /S /MT:50 /r:30 /w:05 2>&1)
                    Write-Verbose -Message "$CopyFiles"

                    if ($CopyFiles -match "Invalid Drive|ERROR"){
                        Write-Warning -Message "$($Comp):Error - Cannot access the drive. Listing this device as a failure."
                        "$($Comp):Error -  Cannot access the drive. Listing this device as a failure: $Comp" | Out-File -FilePath $ResultFile -Append -Force
                        "$CopyFiles" | Out-File -FilePath $ResultFile -Append -Force
                        $FileCopied = $false
                    }

                    Write-Verbose -message "$($Comp): Waiting for $Seconds2Wait for the patch to copy over."
                    Start-Sleep -Seconds $Seconds2Wait
                }
                catch {
                    "$CopyFiles" | Out-File -FilePath $ResultFile -Append -Force
                    "$($Comp):Error - Failed to copy files: $($_.Exception.Message)" | Out-File -FilePath $ResultFile -Append -Force
                    $FileCopied = $false
                }
                #endregion - Copy files to local device

                #region - Run the update
                if ($FileCopied -eq $true){

                    #Install the patch
                    try {
                        $proc = Invoke-WmiMethod -ComputerName $Comp -Class Win32_Process -Name Create -ArgumentList "cmd.exe /c $Task" -ErrorAction Stop

                        #$proc = ([WMICLASS]"\\$Comp\ROOT\CIMV2:win32_process").Create("cmd.exe /c $Task")
                        Write-Verbose -Message "$($Comp): Patch applied: $proc"
                        $UpdateComplete = "PendingVerification"
                    }
                    catch {
                        Write-Warning -Message "$($Comp):Error - Starting patch installation process: $Proc"
                        "$($Comp):Error - Starting patch installation process: $($_.Exception.Message)" | Out-File -FilePath $ResultFile -Append -Force
                    }

                    #Check if patch installed
                    foreach ($Num in 1..20){
                        try {
                            if ($num -in 1,5,10,15,20){
                                Write-Warning -Message "$($Comp):...Checking WMI..."
                            }

                            if ((Get-WmiObject -ComputerName $Comp -Class Win32_QuickFixEngineering -ea SilentlyContinue | Where-Object {$_.hotfixid -match 'kb4500331'})){
                                $UpdateComplete = "Verified"
                                Write-Verbose -message "$($Comp):Success - Update verified, breaking from this device's session"
                                break
                            }

                            #Check the file version directly if we've hit the max retries, should probably move this outside the other try block.
                            if ($Num -eq 20){
                                Write-Warning -Message "$($Comp): Could not verify via Win32_QuickFixEngineering"

                                #Check the file version via WMI
                                $File = Get-WmiObject -Computername $Comp -Query "SELECT * FROM CIM_DataFile WHERE Name = 'C:\\Windows\\System32\\drivers\\termdd.sys'" -ErrorAction Stop
                                try {
                                    if ([version](($File.version).split(' ')[0]) -ge $Version){
                                        $UpdateComplete = "Verified"
                                    }
                                    else {$UpdateComplete = "RebootToVerify"}
                                }
                                catch [System.Management.Automation.RuntimeException]{
                                    Write-Warning -Message "$($Comp): ($($File.Version)) did not meet expected format."
                                    $UpdateComplete = "RebootToVerify"
                                }
                                catch {
                                    Write-Warning -Message "$($Comp): Unexpected error - $($_.Exception.GetType().FullName) ^ $($_.Exception.Message)."
                                    $UpdateComplete = "RebootToVerify"
                                }
                            }
                        }
                        catch {
                            Write-Warning -message "$($Comp):Error On Attempt $Num - Checking patch installation process: $($_.Exception.Message)"
                            if ($Num -eq 50){
                                "$($Comp):Error - Checking patch installation process: $($_.Exception.Message)" | Out-File -FilePath $ResultFile -Append -Force
                            }

                            #Declare this each time. It looks a little different so I can tell where the variable data was populated by. Not the best way, but whatever...
                            $UpdateComplete = "UnVerified - $($_.Exception.Message)"
                        }
                        #pause briefly before the next attempt
                        Start-Sleep -Seconds 5
                    }

                }

                #endregion - Run the update
            }#if ($UpdateComplete -eq 'Unverified'){

        }#if (Test-Connection -ComputerName $Comp -Count 3 -ErrorAction SilentlyContinue){
        else {
            Write-Warning -Message "$($Comp):Could not contact device!"
            "$($Comp):Error - Could not contact device!" | Out-File -FilePath $ResultFile -Append -Force
        }

        [pscustomobject]@{
            Device = $Comp
            WMIHealthy = $WMIConnects
            HasAccess = $Access
            OperatingSystem = $OSName
            FileCopied = $FileCopied
            Okay2Patch = $OkayToProceed
            PatchApplied = $UpdateComplete
        }

    }#ScriptBlock block
        
    #for each job in the list we'll start jobs, up to the max allowed, then wait 
    #for some to stop running, then start more until we're out of jobs to start
    Write-Verbose -Message "Starting job: $JobName"
    if ($LogFile){Write-WSMLogMessage -Message "$Section Starting job: $JobName" -LogFile $LogFile}
    
    Start-Job -Name $JobName -ScriptBlock $ScriptBlock
    
    #While the queue is full, we'll sleep 1/4 second and then recheck for open queue slots
    while ((Get-Job -Name $JobName | Where-Object {$_.State -match 'Running'}).Count -ge $MaxConcurrentJobs){
        Start-Sleep -Milliseconds 250
        
        #We'll look at all the current jobs and if any of them have been running 
        #for longer than the max allocated time, we'll put them into a stopped 
        #state (collect the data later) so we don't get hung up with a full queue
        Get-Job -Name $JobName | Where-Object {$_.State -match 'Running'} | ForEach-Object {
            #if ((Get-Date) -gt (Get-Date ($_.PSBeginTime)).AddHours($MaxRunHours)){
            if ((Get-Date) -gt (Get-Date ($_.PSBeginTime)).AddMinutes($MaxRunMinutes)){
                Write-Verbose -Message 'Stopping Job...'
                Stop-Job $_
            }
        }#foreach get-job that runs over time
    }#while

    # Receive any completed jobs
    Get-Job -Name $JobName | ForEach-Object {
        #If the state is completed, we'll "Receive" the job to collect the exit data. 
        #If it's in any other state, we'll spoof the data, then force remove the job.
        if ($_.State -match 'Completed') {
            #Receive the job to collect it's output data
            $Output += Receive-Job $_ -Verbose
            
            Write-Verbose -Message "Receiving job: $($_.Name)"
            if ($LogFile){Write-WSMLogMessage -Message "$Section Receiving job: $($_.Name)" -LogFile $LogFile}

            #Remove the job. At the end of the loop all jobs should be cleared
            Remove-Job $_ -Force -Verbose
        }
    }#foreach Get-Job
}#foreach

# Allow any remaining jobs to complete and collect their data
#While there are any jobs in the 'Running' state, and the expiration time has not elapsed, sleep for 1/4 second and check again
while ((Get-Job -Name $JobName -ErrorAction SilentlyContinue | Where-Object {$_.State -ne 'Stopped'}).Count -gt 0) {
    #We'll look at all the current jobs and if any of them have been running 
    #for longer than the max allocated time, we'll put them into a stopped 
    #state (collect the data later) so we don't get hung up with a full queue
    Get-Job -Name $JobName | Where-Object {$_.State -match 'Running'} | ForEach-Object {
        #if ((Get-Date) -gt (Get-Date ($_.PSBeginTime)).AddHours($MaxRunHours)){
        if ((Get-Date) -gt (Get-Date ($_.PSBeginTime)).AddMinutes($MaxRunMinutes)){
            Write-Verbose -Message 'Stopping Job...'
            Stop-Job $_ -Verbose
        }
    }#foreach get-job that runs over time

    # Receive any completed jobs
    Get-Job -Name $JobName | ForEach-Object {
        #If the state is completed, we'll "Receive" the job to collect the exit data. 
        #If it's in any other state, we'll spoof the data, then force remove the job.
        if ($_.State -match 'Completed') {
            #Receive the job to collect it's output data
            $Output += Receive-Job $_ -Verbose
            
            Write-Verbose -Message "Receiving job: $($_.Name)"
            if ($LogFile){Write-WSMLogMessage -Message "$Section Receiving job: $($_.Name)" -LogFile $LogFile}

            #Remove the job. At the end of the loop all jobs should be cleared
            Remove-Job $_ -Force -Verbose
        }
    }#foreach Get-Job

    # Do a little sleep before we cycle again to releive the processor
    Start-Sleep -Milliseconds 250
}#while

# Ones that didn't get done before
$Remaining += $PCList | Where-Object {$_ -notin $Output.ComputerName}

$Output | Out-File -FilePath $PostOutFile -Append -Force

$Output | ConvertTo-Json | Out-File -FilePath $PostOutJson -Append -Force