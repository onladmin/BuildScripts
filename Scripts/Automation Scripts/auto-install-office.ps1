$OfficeExe = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/setup.exe'
$OfficeXMLInstall = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/configuration-Office365-x86.xml'
$OfficeXMLUninstall = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/configruation_uninstall.xml'
$OfficeXMLHomeUninstall = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/configuration_uninstall_home.xml'
$OfficeXMLBuisnessUninstall = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/configuration_uninstall_buisness.xml'

function start-officeinstall {
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $OfficeExe -outfile c:\temp\scriptdownloads\office365setup.exe
    Invoke-WebRequest $OfficeXMLHomeUninstall -outfile c:\temp\scriptdownloads\office365uninstallhome.xml
    Invoke-WebRequest $OfficeXMLInstall -outfile c:\temp\scriptdownloads\configuration.xml
    Invoke-WebRequest $OfficeXMLUninstall -outfile c:\temp\scriptdownloads\office365uninstall.xml
    Invoke-WebRequest $OfficeXMLBuisnessUninstall -outfile c:\temp\scriptdownloads\office365uninstallbuisness.xml
    $ProgressPreference = 'Continue'
    c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstallhome.xml | Out-Null
    c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstallbuisness.xml | Out-Null
    c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\office365uninstall.xml | Out-Null
    c:\temp\scriptdownloads\office365setup.exe /configure c:\temp\scriptdownloads\configuration.xml | Out-Null
    Write-Output "Office365 is now installed."
}

start-officeinstall
