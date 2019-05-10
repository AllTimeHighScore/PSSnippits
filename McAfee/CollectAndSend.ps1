
function Update-McAfeeAgent {
    #Locate the cmdagent
    $CmdAgent = ""
    
    #Check McAfee InstallRoot value for install location so we can look for the CMDAgent.exe in the next block
    $AgentRegPath = (Get-ChildItem -Path 'HKLM:\SOFTWARE\McAfee','HKLM:\SOFTWARE\Wow6432Node\McAfee' -ErrorAction SilentlyContinue | Where-Object {($_.property -match 'Installroot')}).pspath
    
    #Look for the CMDAgent
    if ($CmdAgent = (Get-Item -Path "$(Get-ItemPropertyValue -Path $AgentRegPath -name InstallRoot)\cmdagent.exe" -ErrorAction SilentlyContinue).FullName){
        "$CmdAgent located."
    }
    else {
        Write-Warning -Message 'Error: cmdagent.exe was not located! Script Exiting!'
    }
    
    #attempt to update the policy
    '-p','-c','-e' | ForEach-Object {
        try {
            "Update Comand line: $CmdAgent $_"
            $AgentAction = Start-Process -FilePath $CmdAgent -ArgumentList $_ -PassThru -ErrorAction Stop
            #Wait a short while for the process to complete but move on if it doesn't. This is only nudging the update process along.
            $AgentAction | Wait-Process -Timeout 20 -ErrorAction Stop
            "Action succeeded. Result = $($AgentAction.ExitCode)"
            #Because we sometimes see weird (Read unexplained) returns on the -c we're going to give each action some space to process
            Start-Sleep -Seconds 5 -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning -Message "Action failed. Result = $($AgentAction.ExitCode)"
        }
    }#End Agent updating action
}