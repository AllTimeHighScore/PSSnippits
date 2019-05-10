<#Created by Kevin Van Bogart - 12/12/2017...#>
        #$Query = "SELECT * FROM CCM_TSExecutionRequest WHERE State = 'Running'"
        $Query = "SELECT * FROM CCM_TSExecutionRequest WHERE RunningState = 'NotifyExecution'"
        $TSRunning = Get-ciminstance -Query $Query -Namespace root\CCM\SoftMgmtAgent
        
        $MemberProgs = $null
        $TS_Apps = @()
        $TS_PKGs = ""
        $MemberProgs = $TSRunning.TS_MemberProgramID
        $PKGAdverts = $TSRunning.OptionalAdvertisements
        $TaskSequenceName = $TSRunning.MIFPackageName
        
        "member Programs: $MemberProgs"

        if ($MemberProgs){
            $Apps = $MemberProgs.split(" ").split('/') | Where-Object {$_ -match "App*"} -ErrorAction SilentlyContinue | Get-Unique
            "After Splitting Member progs: $Apps"
            
            foreach ($App in $Apps){$AppQuery = "Select * from CCM_Application where Id like '%$App'"
                $AppObject = Get-ciminstance -Query $AppQuery -Namespace 'ROOT\ccm\clientsdk'
                $AppName = "$($AppObject.Publisher.Replace(' ',''))_$($AppObject.Name.Replace(' ',''))_$($AppObject.SoftwareVersion)"
                $TS_Apps += $AppName
            }
        }
        
        #All the packages should have the same advertisement ID
        #We'll query that ID to find any packages that will be installed.
        if ($PKGAdverts){
            $Packages = get-ciminstance -query "SELECT * FROM CCM_SoftwareDistribution where ADV_AdvertisementID like '%$PKGAdverts'" -namespace "root\ccm\policy\machine\actualconfig"
            $TS_PKGs = $Packages.PKG_Name | Where-Object {$_ -notmatch $TaskSequenceName} | Get-Unique
        }

        "Task Sequence Name: $TaskSequenceName"

        "Optional Advertisements: $PKGAdverts"

        #List out Applicaiton Objects
        if ($TS_Apps){
            $TS_Apps | ForEach-Object {"Applications: $_"}
        }
        else {
            "No Application objects being installed"
        }
        
        #List out legacy Package objects
        if ($TS_PKGs){
            $TS_PKGs | ForEach-Object {"Packages: $_"}
        }
        else {
            "No Package Objects being installed"
        }
