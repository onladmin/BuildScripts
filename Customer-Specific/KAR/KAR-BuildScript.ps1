#Automated Build script for KAR
# MQ 11/02/2021

$SoftwareInstallChrome = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/GoogleChromeStandaloneEnterprise64.msi'
$SoftwareInstall7zip = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/7z1604-x64.msi'
$SoftwareInstallJava = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/jre1.8.0_26164.msi'
$SoftwareInstallZoomFile = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/ZoomInstallerFull.msi'
$SoftwareInstallAdobeReader = 'https://onl-my.sharepoint.com/:u:/g/personal/mohammed_quashem_onlinesupport_co_uk/EcRWAKSO321GgevynQeMUzkBpZ-6wm-kHKs7_uScUdfZmw?e=4bhFXR&download=1'
$SoftwareInstallFactSect = 'https://support.factset.com/workstation/gr/64/'
$Office365Install = 'https://github.com/onladmin/BuildScripts/raw/master/Scripts/Automation%20Scripts/auto-install-office.ps1'
$MimecastInstall = 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/Mimecast%20for%20Outlook%207.0.1740.17532%20(32%20bit).msi'
$NeteXtenderInstall= 'https://github.com/onladmin/BuildScripts/raw/master/Install%20files/NetExtender.8.6.265.MSI'
$PhotoviewerInstall = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/Restore_Windows_Photo_Viewer_ALL_USERS.reg'
$DefaultAppPre1909= 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/Pre1909DefaultAppAssociations.xml'
$DefaultApp = 'https://github.com/onladmin/BuildScripts/raw/master/Regkeys_xmls/2004AppAssociations.xml'
#$BloatwareRemoverWin10 = 'https://github.com/onladmin/BuildScripts/raw/master/Scripts/Uninstall_windows10_bloatware.ps1'
$HPBloatwareRemover = 'https://github.com/onladmin/BuildScripts/raw/master/Scripts/HP-Bloatware-removal.ps1'
$DellBloatwareRemover = 'https://github.com/onladmin/BuildScripts/raw/master/Scripts/Dell-Bloatware-Removal.ps1'

############################################

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
  powershell c:\temp\scriptdownloads\office365install.ps1
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\chrome.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\7zip.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\java.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\factsect.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\ZoomInstaller.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\adobereader\acroread.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i c:\temp\scriptdownloads\mimecast32bit.msi /qn /norestart allusers=2'
  Start-Process msiexec.exe -Wait -ArgumentList '/i C:\temp\scriptdownloads\netextender.msi /qn /norestart allusers=2'
  Invoke-Command {reg import C:\temp\scriptdownloads\Photoviewer.reg *>&1 | Out-Null}
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

fucntion start-disablefirewall {
  Invoke-Command {reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Systray" /v HideSystray /t reg_dword /d 1 /f} > $null 2>&1 # Removes Windows Defender from taskbar
  Write-Output ''
  Set-NetFirewallProfile -Profile Domain -Enabled False
  Write-Output "Windows Firewall disabled for Domain network."
}

function start-disablessd-defrag {schtasks /Delete /TN "\Microsoft\Windows\Defrag\ScheduledDefrag"  /f}

################################################ Last step options

function start-rename-computer {
  Clear-Host
  do { $myInput = (Read-Host 'Would you like to re-name this computer? (Y/N)').ToLower() } while ($myInput -notin @('Y','N'))
  if ($myinput -eq 'Y') {
  Write-Output ''
  $ComputerName = Read-Host -Prompt 'Enter the PC name you would like to rename to'
  Rename-Computer -NewName "$ComputerName" -passthru
  Write-Output ''
  pause
  }
  else {
  Write-Output "Will not rename computer..."
  Clear-Host
  }}
  
function start-joindomain {
  Clear-Host
  do { $myInput = (Read-Host 'Would you like to join this PC to a domain? (Y/N)').ToLower() } while ($myInput -notin @('Y','N'))
  if ($myinput -eq 'Y') {
  Write-Output ''
  $DomainName = Read-Host -Prompt 'Enter the domain name to join this PC'
  Write-Output ''
  Write-Output "Use domain credentials to join the PC to domain. There will be a prompt on the screen to do this part."
  Write-Output ''
  Start-Sleep 5
  add-computer –domainname "$DomainName" -PassThru -Options JoinWithNewName,AccountCreate
  Write-Output ''
  Write-Output "The above may not be accurate. (Because a restart is required to update the information)."
  Write-Output ''
  Write-Output "Just to be sure..."
  Write-Output "Providing Hostname & Domain output..."
  Write-Output ''
  hostname
  systeminfo | findstr /B "Domain"
  Write-Output ''
  Write-Output "Please check if the above information is correct and then continue."
  pause
  }
  else {
  Write-Output "PC Will not be joined to the domain."
  Write-Output "Please continue."
  Clear-Host
}}
  
function start-power-config {
  Clear-Host
  Write-Output ''
  Write-Output "Set power options."
  Write-Output ''
  Write-Output "Laptop: Hiberate on power button and do nothing when the lid closes. Show Hibernate button on start menu."
  Write-Output "Desktop: Show hibernate button on start menu. Disable sleep."
  Write-Output ''
  do { $myInput = (Read-Host 'Change power settings?(laptop/desktop/N)').ToLower() } while ($myInput -notin @('laptop','desktop','N'))
  if ($myinput -eq 'laptop') {
  Write-Output ''
  Write-Output "Modifying Power settings for laptop..."
  Write-Output ''
  powercfg -setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 2
  powercfg -setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 2
  powercfg -setacvalueindex SCHEME_CURRENT sub_buttons lidaction 0
  powercfg -setdcvalueindex SCHEME_CURRENT sub_buttons lidaction 0
  powercfg -SetActive SCHEME_CURRENT
  Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings /v ShowHibernateOption /t reg_dword /d 1 /f}
  Write-Output ''
  Write-Output "Done."
  Write-Output "Please continue."
  Write-Output ''
  pause
  Clear-Host
  }
  if ($myinput -eq 'desktop') { 
  Write-Output ''
  Write-Output "Modifying Power settings for Desktop..."
  Write-Output ''
  Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings /v ShowHibernateOption /t reg_dword /d 1 /f}
  powercfg -change -standby-timeout-dc 0
  powercfg -change -standby-timeout-ac 0
  powercfg -SetActive SCHEME_CURRENT
  Write-Output ''
  Write-Output "Done."
  Write-Output "Please continue."
  Write-Output ''
  pause
  Clear-Host
  }
  if ($myinput -eq 'N') {
  Write-Output ''
  Write-Output "Not modifying power settings..."
  pause
  Clear-Host
}}

function start-bitlocker-updaterecovery {
    Clear-Host
    do { $myInput = (Read-Host 'Would you like to update Bitlocker recovery key?(Y/N)').ToLower() } while ($myInput -notin @('Y','N'))
    if ($myinput -eq 'Y') {
    Write-Output ''
    Write-Output "Updating Bitlocker recovery key to AD..."
    Write-Output ''
    $BitLocker = Get-BitLockerVolume -MountPoint $env:SystemDrive
    $RecoveryProtector = $BitLocker.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
    
    Backup-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $RecoveryProtector.KeyProtectorID
    #BackupToAAD-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $RecoveryProtector.KeyProtectorID
    Write-Output ''
    Write-Output "Done."
    Write-Output ''
    #Write-Output "If you get an error regarding BackUptoAAD this is for AzureActive Directory so it can be safely ignored."
    #Write-Output "Check that the recoverykey/password matches the one in the c:\temp if it is, remove the one in c:\temp and continue."
    Write-Output "If you get any errors double check the computer is in the correct OU."
    Write-Output "Please confirm the recovery key is in AD."
    pause}
    else {
    Write-Output "Not updating bitlocker recovery."
}}

function start-dellbloatwareremoval {
  Invoke-WebRequest $DellBloatwareRemover -outfile c:\temp\scriptdownloads\dellbloatwareremoval.ps1
  powershell c:\temp\scriptdownloads\dellbloatwareremoval.ps1
}

function start-hpbloatwareremoval {
  Invoke-WebRequest $HPBloatwareRemover -outfile c:\temp\scriptdownloads\hpbloatwareremoval.ps1
  powershell c:\temp\scriptdownloads\hpbloatwareremoval.ps1
}

################################################ Menu

function start-script {
  start-software-install
  start-shortcuts-default-apps
  start-clearstartmenu
  start-modifyuac
  start-enablesystemrestore
  start-setdefault-timzeone
  start-disablefirewall
  start-disablessd-defrag
  start-windows-update #make sure this is last
}

function lastscript {
start-rename-computer
start-joindomain
start-power-config
start-bitlocker-updaterecovery
start-dellbloatwareremoval
start-hpbloatwareremoval
}
function start-mainmenu {
  Write-Output ''
  Write-Output "KAR Build Script."
  Write-Output ''
  Write-Output "Choose option 1 for automated."
  Write-Output "Choose option 2 for the last steps."
  Write-Output ''
  do { $myInput = (Read-Host 'Type an option').ToLower() } while ($myInput -notin @('1','2','3'))
if ($myinput -eq '1') {start-script}
if ($myinput -eq '2') {last-script}
}

mkdir c:\temp > $null 2>&1
Remove-Item c:\temp\scriptdownloads -recurse -force > $null 2>&1
mkdir c:\temp\scriptdownloads > $null 2>&1

start-mainmenu

