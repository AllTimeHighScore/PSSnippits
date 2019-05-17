#Check OS Architecture
$RegPaths = @()
$OfficeVersions = @()
#just screwing with switches, no value in using it with this specific function.
#switch ([environment]::Is64BitOperatingSystem){
#    $true {$RegPaths = @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')}
#    $false {$RegPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'}
#}

$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
if ([environment]::Is64BitOperatingSystem){$RegPaths += 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'}

#For the product code I'm just playing with 008C because it's present for C2R
#The actual base install is 0011

#This format {BRMMmmmm-PPPP-LLLL-p000-D000000FF1CE} is valid for all Office Installs from 2007 on.
#### -> https://support.microsoft.com/en-us/help/3120274/description-of-the-numbering-scheme-for-product-code-guids-in-office-2
#         Office ProPlus 2007 MSI x86 
#   "{90120000-0011-0000-0000-0000000FF1CE}"
#         Office ProPlus 2007 MSI x64 
#   "{90120000-0011-0000-1000-0000000FF1CE}"
#         Office ProPlus 2010 MSI x86 
#   "{90140000-0011-0000-0000-0000000FF1CE}"
#         Office ProPlus 2010 MSI x64 
#   "{90140000-0011-0000-1000-0000000FF1CE}"
#         Office ProPlus 2016 MSI x86 
#   "{90160000-0011-0000-0000-0000000FF1CE}" 
#         Office ProPlus 2016 MSI x64 
#   "{90160000-0011-0000-1000-0000000FF1CE}"
#         Office ProPlus 2019 MSI x86 
#   "{90190000-0011-0000-0000-0000000FF1CE}" ????????
#         Office ProPlus 2019 MSI x64 
#   "{90190000-0011-0000-1000-0000000FF1CE}" ????????


#Matches only the base install
#{BRMMmmmm-PPPP-LLLL-p000-D000000FF1CE}
#B = 9 = RTM. This is the first version that is shipped (the initial release).
#R = 0 = Volume license
#MM = Major version = ([1-9]{1})([0-9]{1}) Exp: 16(Office 2016), 14(Office 2010)
#mmmm = Minor version = (\d{4}) Exp: 0000 or 1234 or 9992
#PPPP = ProductID = 0{2}1{2} = 0011 = Microsoft Office Professional Plus Base Install
#LLLL = Language Code = (0{4}) = 0000 Note: We're only looking for the base which is 0000
#p000 = Bitness and Unallocated = [0-1]0{3} = 1000 Example: (first char: 1 is x64, 0 is x86, the rest are unused)
#D = We'll never have the debug so this should stay 0
#000000FF1CE = 0{7}FF1CE Note: Self-explanatory
[regex]$Regex = "^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}1{2}-(0{4})-[0-1]0{3}-0{7}FF1CE}$"

#Test on extensibility for c2r
#[regex]$Regex = "^{90([1-9]{1})([0-9]{1})(\d{4})-0{2}8C-(0{4})-[0-1]0{3}-0{7}FF1CE}$"
#matches any language, ends up getting too much crap
#[regex]$Regex = "{90([1-9]{1})([0-9]{1})(\d{4})-008C-[0-9]+-[0-1]000-0000000FF1CE}"

ForEach($RegPath in $RegPaths){
    (Get-ChildItem -Path $RegPath).PSChildName | ForEach-Object {
        if ($OfficeGUID = ($Regex.Matches($_)).value){
            $OfficeVersions += [version](Get-ItemProperty -Path "$RegPath\$OfficeGUID").DisplayVersion
        }
    }
}

$OfficeVersions[0].Major




#Shawn's suggestioned improvements.
$RegRefactorOriginal = "^{90[1-9][0-9]0{4}-0{2}\d{2}-0{4}-[0-1]0{3}-0{7}FF1CE}$"
$RegExact ="^{[0-9,A-F]{2}[0-9]{6}-([0-9,A-F]{4}-){2}[0-1]0{3}-[0-1]0{6}FF1CE}$" 