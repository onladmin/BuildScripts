mkdir c:\temp > $null 2>&1
Remove-Item c:\temp\scriptdownloads -recurse -force > $null 2>&1
mkdir c:\temp\scriptdownloads > $null 2>&1

$OfficeExe = 'https://github.com/Ray-MRQ/MRQ/raw/master/Install%20files/setup.exe'
$OfficeXMLInstall = 'https://github.com/Ray-MRQ/MRQ/raw/master/Regkeys_xmls/configuration-Office365-x86.xml'
$OfficeXMLUninstall = 'https://github.com/Ray-MRQ/MRQ/raw/master/Regkeys_xmls/configruation_uninstall.xml'
$OfficeXMLHomeUninstall = 'https://github.com/Ray-MRQ/MRQ/raw/master/Regkeys_xmls/configuration_uninstall_home.xml'
$OfficeXMLBuisnessUninstall = 'https://github.com/Ray-MRQ/MRQ/raw/master/Regkeys_xmls/configuration_uninstall_buisness.xml'

function start-officecheck {
start-officeuninstall-pro
start-officeuninstall-buisness
start-officeuninstall-home
start-officeinstall
}

function start-officeuninstall-pro {
Write-Output "Starting uninstall process for all Office apps."
Write-Output ''
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest $OfficeExe -outfile c:\temp\scriptdownloads\office365setup.exe
Invoke-WebRequest $OfficeXMLUninstall -outfile c:\temp\scriptdownloads\office365uninstall.xml
$ProgressPreference = 'Continue'
c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstall.xml
Write-Output ''
}
 
function start-officeuninstall-buisness {
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest $OfficeExe -outfile c:\temp\scriptdownloads\office365setup.exe
Invoke-WebRequest $OfficeXMLBuisnessUninstall -outfile c:\temp\scriptdownloads\office365uninstallbuisness.xml
$ProgressPreference = 'Continue'
c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstallbuisness.xml
}

function start-officeuninstall-home {
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest $OfficeExe -outfile c:\temp\scriptdownloads\office365setup.exe
Invoke-WebRequest $OfficeXMLHomeUninstall -outfile c:\temp\scriptdownloads\office365uninstallhome.xml
$ProgressPreference = 'Continue'
c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstallhome.xml
Write-Output "Office apps should be uninstalled."
Write-Output ''
}
    
function start-officeinstall {
Write-Output ''
Write-Output "Starting Office 365 install silently...."
Write-Output ''
Write-Output "Downloading and installing Office365"
Write-Output ''
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest $OfficeExe -outfile c:\temp\scriptdownloads\office365setup.exe
Invoke-WebRequest $OfficeXMLInstall -outfile c:\temp\scriptdownloads\configuration.xml
$ProgressPreference = 'Continue'
c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\configuration.xml
Write-Output "Office365 is now installed."
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3
[Net.ServicePointManager]::SecurityProtocol = "Tls, Tls11, Tls12, Ssl3"

start-officecheck
Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" -Destination "C:\Users\Public\Desktop\Outlook.lnk"
Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Destination "C:\Users\Public\Desktop\Excel.lnk"
Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk" -Destination "C:\Users\Public\Desktop\Powerpoint.lnk"
Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Destination "C:\Users\Public\Desktop\Word.lnk"
Remove-Item c:\temp\scriptdownloads -recurse -force > $null 2>&1 