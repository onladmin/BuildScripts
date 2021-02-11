#Automated Build script for KAR
# MQ 11/02/2021

$SoftwareInstallChrome = 'https://github.com/onladmin/MRQ/raw/master/Install%20files/GoogleChromeStandaloneEnterprise64.msi'
$SoftwareInstall7zip = 'https://github.com/onladmin/MRQ/raw/master/Install%20files/7z1604-x64.msi'
$SoftwareInstallJava = 'https://github.com/onladmin/MRQ/raw/master/Install%20files/jre1.8.0_26164.msi'
$SoftwareInstallZoomFile = 'https://github.com/onladmin/MRQ/raw/master/Install%20files/ZoomInstallerFull.msi'
$SoftwareInstallAdobeReader = 'https://onl-my.sharepoint.com/:u:/g/personal/mohammed_quashem_onlinesupport_co_uk/EcRWAKSO321GgevynQeMUzkBpZ-6wm-kHKs7_uScUdfZmw?e=4bhFXR&download=1'
$SoftwareInstallFactSect = 'https://support.factset.com/workstation/gr/64/'
$Office365Install = 'https://github.com/onladmin/BuildScripts/raw/master/Scripts/Automation%20Scripts/auto-install-office.ps1'
$MimecastInstall = 'https://github.com/onladmin/MRQ/raw/master/Install%20files/Mimecast%20for%20Outlook%207.0.1740.17532%20(32%20bit).msi'
$NeteXtenderInstall= 'https://github.com/onladmin/MRQ/raw/master/Install%20files/NetExtender.8.6.265.MSI'
$PhotoviewerInstall = 'https://github.com/onladmin/MRQ/raw/master/Regkeys_xmls/Restore_Windows_Photo_Viewer_ALL_USERS.reg'
$DefaultAppPre1909= 'https://github.com/onladmin/MRQ/raw/master/Regkeys_xmls/Pre1909DefaultAppAssociations.xml'
$DefaultApp = 'https://github.com/onladmin/MRQ/raw/master/Regkeys_xmls/2004AppAssociations.xml'
#$BloatwareRemoverWin10 = 'https://github.com/onladmin/MRQ/raw/master/Scripts/Uninstall_windows10_bloatware.ps1'
#$HPBloatwareRemover = 'https://github.com/onladmin/MRQ/raw/master/Scripts/HP-Bloatware-removal.ps1'
#$DellBloatwareRemover = 'https://github.com/onladmin/MRQ/raw/master/Scripts/Dell-Bloatware-Removal.ps1'


function start-software-install {
  $ProgressPreference = 'SilentlyContinue'
  Invoke-WebRequest $SoftwareInstallChrome -outfile c:\temp\scriptdownloads\chrome.msi
  Invoke-WebRequest $SoftwareInstall7zip -outfile c:\temp\scriptdownloads\7zip.msi
  Invoke-WebRequest $SoftwareInstallJava -outfile c:\temp\scriptdownloads\java.msi
  Invoke-WebRequest $SoftwareInstallFactSect -outfile c:\temp\scriptdownloads\factsect.msi
  Invoke-WebRequest $SoftwareInstallAdobeReader -outfile c:\temp\scriptdownloads\adobereader.zip
  Invoke-WebRequest $SoftwareInstallZoomFile -outfile c:\temp\scriptdownloads\ZoomInstaller.msi
  Invoke-WebRequest $MimecastInstall -outfile c:\temp\scriptdownloads\mimecast32bit.msi
  Invoke-WebRequest $NeteXtenderInstall -outfile c:\temp\scriptdownloads\netextender.msi
  Invoke-WebRequest $PhotoviewerInstall -outfile c:\temp\scriptdownloads\Photoviewer.reg
  Invoke-WebRequest $Office365Install -outfile c:\temp\scriptdownloads\office365install.ps1
  Expand-Archive -LiteralPath C:\temp\scriptdownloads\adobereader.zip -DestinationPath C:\temp\scriptdownloads\
  $ProgressPreference = 'Continue'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\chrome.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\7zip.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\java.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\factsect.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\ZoomInstaller.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\adobereader\acroread.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\mimecast32bit.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i C:\temp\scriptdownloads\netextender.msi /qn /norestart allusers=2'
  Invoke-Command {reg import C:\temp\scriptdownloads\Photoviewer.reg *>&1 | Out-Null}
  powershell c:\temp\scriptdownloads\office365install.ps1
  Write-Output "Installed 7zip, Java, Chrome, Factsect, Zoom, Mimecast, Net extender, Office365 app and Photo viewer and Adobe reader silently."
  
}
function start-shortcuts-default-apps {
  if ($WindowsVersion -le '1909') { 
    Invoke-WebRequest $DefaultAppPre1909 -outfile c:\temp\scriptdownloads\MyDefaultAppAssociations.xml 
    dism /online /Import-DefaultAppAssociations:"c:\temp\scriptdownloads\MyDefaultAppAssociations.xml" }
    else {
    Invoke-WebRequest $DefaultApp -outfile c:\temp\scriptdownloads\MyDefaultAppAssociations.xml
    dism /online /Import-DefaultAppAssociations:"c:\temp\scriptdownloads\MyDefaultAppAssociations.xml" }
    Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" -Destination "C:\Users\Public\Desktop\Outlook.lnk"
    Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Destination "C:\Users\Public\Desktop\Excel.lnk"
    Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk" -Destination "C:\Users\Public\Desktop\Powerpoint.lnk"
    Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Destination "C:\Users\Public\Desktop\Word.lnk"
}

function start-clearstartmenu {
  Write-Output "Modiying start menu & taskbar settings."
  Write-Output ''
  $START_MENU_LAYOUT = @"
  <LayoutModificationTemplate xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" Version="1" xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout" xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification">
      <LayoutOptions StartTileGroupCellWidth="6" />
      <DefaultLayoutOverride>
          <StartLayoutCollection>
              <defaultlayout:StartLayout GroupCellWidth="6" />
          </StartLayoutCollection>
      </DefaultLayoutOverride>
      <CustomTaskbarLayoutCollection PinListPlacement="Replace">
        <defaultlayout:TaskbarLayout>
          <taskbar:TaskbarPinList>
            <taskbar:DesktopApp DesktopApplicationLinkPath="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\File Explorer.lnk" />
          </taskbar:TaskbarPinList>
        </defaultlayout:TaskbarLayout>
      </CustomTaskbarLayoutCollection>
  </LayoutModificationTemplate>
"@

 $layoutFile="C:\Windows\StartMenuLayout.xml"
  #Delete layout file if it already exists
  If(Test-Path $layoutFile)
  {
      Remove-Item $layoutFile
  }
  #Creates the blank layout file
  $START_MENU_LAYOUT | Out-File $layoutFile -Encoding ASCII
  $regAliases = @("HKLM", "HKCU")
  #Assign the start layout and force it to apply with "LockedStartLayout" at both the machine and user level
  foreach ($regAlias in $regAliases){
      $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
      $keyPath = $basePath + "\Explorer" 
      IF(!(Test-Path -Path $keyPath)) { 
          New-Item -Path $basePath -Name "Explorer"
      }
      Set-ItemProperty -Path $keyPath -Name "LockedStartLayout" -Value 1
      Set-ItemProperty -Path $keyPath -Name "StartLayoutFile" -Value $layoutFile
  }
  Stop-Process -name explorer -Force
  Start-Sleep -s 5
  $wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('^{ESCAPE}')
  Start-Sleep -s 5
  #Enable the ability to pin items again by disabling "LockedStartLayout"
  foreach ($regAlias in $regAliases){
      $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
      $keyPath = $basePath + "\Explorer" 
      Set-ItemProperty -Path $keyPath -Name "LockedStartLayout" -Value 0
  }
  #Restart Explorer and delete the layout file
  Stop-Process -name explorer -Force
  # Uncomment the next line to make clean start menu default for all new users
  Import-StartLayout -LayoutPath $layoutFile -MountPath $env:SystemDrive\
  Remove-Item $layoutFile
}
function start-modifyuac {Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v PromptOnSecureDesktop /t reg_dword /d 0 /f}}

function start-enablesystemrestore {Enable-ComputerRestore -Drive "C:\"}

function start-setdefault-timzeone {
  set-timezone -id "GMT Standard Time" -passthru
  Get-Date -Format “dddd MM/dd/yyyy HH:mm K”
  Rename-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\00000409" -NewName "00000409-disabled"
  Set-WinDefaultInputMethodOverride -InputTip "0809:00000809"
}

function start-windows-update {
  if ($WindowsVerison -le $OldWindows) { 
    #Removes the Upgrader app if it's installed.
    C:\Windows10Upgrade\Windows10UpgraderApp.exe /ForceUninstall > $null 2>&1
    Remove-Item C:\Windows10Upgrade\*.* -recurse -force > $null 2>&1
    Clear-Host
    #
    Write-Output "Starting windows updates..."
    Write-Output "Please wait..."
    $dir = 'C:\temp\scriptdownloads\packages'
    Remove-Item $dir -recurse -force > $null 2>&1
    mkdir $dir > $null 2>&1
    $webClient = New-Object System.Net.WebClient
    $url = 'https://go.microsoft.com/fwlink/?LinkID=799445'
    $file = "$($dir)\Win10Upgrade.exe"
    $webClient.DownloadFile($url,$file)
    Start-Process -FilePath $file -ArgumentList '/skipeula /auto upgrade /copylogs $dir'
    Write-Output ''
    #start-windows-update-running-checker <may not be needed.
    Write-Output "Please re-run the bloatware remover after restarting as doing a feature update may add new bloatware back in."
    Write-Output ''
    Write-Output "The script is setup to exit after hitting enter."
    Write-Output ''
    Start-Sleep -Seconds 60
    exit
    }
    if ($LatestWindows -match $LatestWindows) {
    start-windows-update-nofeature
}}

function start-feature-update {
  Write-Output "Starting windows updates..."
  Install-PackageProvider -Name NuGet -Force -MinimumVersion 2.8.5.201 > $null 2>&1
  Install-Module -Name PSWindowsUpdate -Force > $null 2>&1
  Write-Output "Checking windows update status..."
  Install-WindowsUpdate -AcceptAll -Install -MicrosoftUpdate -Verbose | Out-File "c:\temp\$(get-date -f dd-MM-yyyy-HH-mm)-WindowsUpdate.log" -force
}

function start-bitlocker {
  $localmachine = $env:computername
  Write-Output "Staring encryption..."
  Write-Output ''
  manage-bde -protectors -add C: -RecoveryPassword > C:\temp\$localmachine.RecoveryPassword.txt
  manage-bde -protectors -add C: -tpm > C:\temp\$localmachine.RecoveryPassword.txt
  manage-bde -protectors -get C: > C:\temp\$localmachine.RecoveryPassword.txt
  manage-bde -on C:
  Write-Output ''
  Write-Output "Stored recovery key in C:\temp\"
}

################################################

function start-script {
  start-software-install
  start-shortcuts-default-apps
  start-clearstartmenu
  start-modifyuac
  start-enablesystemrestore
  start-setdefault-timzeone
  start-windows-update
}

function start-mainmenu {
  Write-Output "KAR Build Script."
  Write-Output "Choose option 1 for automated."
  Write-Output "Choose option 2 for the manual last steps."
  do { $myInput = (Read-Host 'Type an option').ToLower() } while ($myInput -notin @('1','2','3'))
if ($myinput -eq '1') {start-script}
if ($myinput -eq '2') {manual-script}
}

start-mainmenu