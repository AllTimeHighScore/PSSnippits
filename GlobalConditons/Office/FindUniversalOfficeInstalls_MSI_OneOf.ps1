####################################
# - Returns the version of Office
# - MSI Only
# - Author: Kevin Van Bogart
# - Created 07-06-2018
####################################

#Check OS Architecture
$RegPaths = @()
$OfficeVersions = @()
$Type = $null
$Bitness = $null
$Family = $null

$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
if ([environment]::Is64BitOperatingSystem){$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}

#[regex]$Regex = "^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4})-[0-1]0{3}-0{7}FF1CE}$"
#[regex]$BitCheck ="^{[0-9,A-F]{2}[0-9]{6}-([0-9,A-F]{4}-){2}[0-1]0{3}-[0-1]0{6}FF1CE}$"

ForEach($RegPath in $RegPaths){
    
    #Spin through the product codes in the uninstall paths 
    Get-ChildItem -Path $RegPath -PipelineVariable 'ProductCode' | ForEach-Object {
        
        #Check The GUID
        Switch -Regex ($ProductCode.PSChildName){
            "^{[0-9,A-F]{2}[0-9]{6}-([0-9,A-F]{4}-){2}1{1}0{3}-[0-1]0{6}FF1CE}$" {
                $Family = ([version](Get-ItemProperty -Path "$RegPath\$($ProductCode.PSChildName)").DisplayVersion).major
                $Bitness = 64
                $Type = 'MSI'
            }
            "^{[0-9,A-F]{2}[0-9]{6}-([0-9,A-F]{4}-){2}0{1}0{3}-[0-1]0{6}FF1CE}$" {
                $Family = ([version](Get-ItemProperty -Path "$RegPath\$($ProductCode.PSChildName)").DisplayVersion).major
                $Bitness = 64
                $Type = 'MSI'
            }
        }#end switch
    }#end Chile item inspection
}#end foreach

#$Family = $OfficeVersions.major
#$Bitness

-Join($Type,$Bitness,$Family)