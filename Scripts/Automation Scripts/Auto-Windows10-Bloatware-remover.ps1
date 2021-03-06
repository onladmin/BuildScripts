$AppList = "Microsoft.SkypeApp",          
           "Microsoft.ZuneMusic",
           "Microsoft.ZuneVideo",
           "Microsoft.Office.OneNote",
           "Microsoft.BingFinance",
           "Microsoft.BingNews",
           "Microsoft.BingWeather",
           "Microsoft.BingSports",
           "Microsoft.MicrosoftOfficeHub",
		   "Microsoft.Wallet",
		   "Microsoft.OneConnect",
		   "Microsoft.MSPaint",
		   "Microsoft.Print3D",
		   "Microsoft.Messaging",
		   "Microsoft.Microsoft3DViewer",
		   "Microsoft.Windows.Cortana",
		   "Microsoft.3DBuilder",
		   "Microsoft.WindowsAlarms",
		   "Microsoft.windowscommunicationsapps",
		   "Microsoft.Getstarted",
		   "Microsoft.WindowsMaps",
		   "Microsoft.MicrosoftSolitaireCollection",
		   "Microsoft.WindowsFeedbackHub",
		   "Microsoft.MixedReality.Portal",
		   "Microsoft.GetHelp",
		   "Microsoft.People",
		   "Microsoft.549981C3F5F10",
		   "Microsoft.549981cf5f10",
		   "SpotifyAB.SpotifyMusic",
		   "king.com.BubbleWitch3Saga",
		   "A278AB0D.DisneyMagicKingdoms",
		   "A278AB0D.MarchofEmpires",
		   "king.com.CandyCrushSodaSaga",
		   "SpotifyAB.SpotifyMusic",
		   "4DF9E0F8.Netflix",
		   "C27EB4BA.DropboxOEM",
		   "Amazon.com.Amazon",
		   "7EE7776C.LinkedInforWindows",
		   "PricelinePartnerNetwork.Booking.comEMEABigsavingso",
		   "Microsoft.BingTranslator",
		   "MixedRealityLearning"
		   
#
$RemoveXboxAppList = "Microsoft.Xbox.TCUI",
		   "Microsoft.XboxSpeechToTextOverlay",
		   "Microsoft.XboxApp",
		   "Microsoft.XboxGameOverlay",
		   "Microsoft.XboxGamingOverlay"
#

function start-allusers-bloatware-noxbox {

ForEach ($App in $AppList)
{
$PackageFullName = (Get-AppxPackage $App -allusers).PackageFullName
$ProPackageFullName = (Get-AppxProvisionedPackage -online | Where-Object {$_.Displayname -eq $App}).PackageName
if ($PackageFullName)
{
remove-AppxPackage -package $PackageFullName 
}
if ($ProPackageFullName)
{
Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName -allusers
}
}
#######################################################################
ForEach ($App in $RemoveXboxAppList)
{
$PackageFullName = (Get-AppxPackage $App -allusers).PackageFullName
$ProPackageFullName = (Get-AppxProvisionedPackage -online | Where-Object {$_.Displayname -eq $App}).PackageName
if ($PackageFullName)
{
remove-AppxPackage -package $PackageFullName 
}
if ($ProPackageFullName)
{
Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName -allusers
}
}
Write-Output "Completed."
}

function start-basic-bloatware-remover {
     #Stops Cortana from being used as part of your Windows Search Function
    $Search = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    If (Test-Path $Search) {
        Set-ItemProperty $Search AllowCortana -Value 0 
    }
    #Disables Web Search in Start Menu
    $WebSearch = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" BingSearchEnabled -Value 0 
    If (!(Test-Path $WebSearch)) {
        New-Item $WebSearch
    }
    Set-ItemProperty $WebSearch DisableWebSearch -Value 1 
""
Invoke-Command {reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" /f}
Invoke-Command {reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" /f}
Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Explorer\ /v DisableSearchBoxSuggestions /t reg_dword /d 1 /f}
Invoke-Command {reg add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced /v ShowTaskViewButton /t reg_dword /d 0 /f}
Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Search /v SearchboxTaskbarMode /t red_dword /d 0 /f}
Invoke-Command {reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Explore /v HidePeopleBar /t reg_dword /d 1 /f}
Remove-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\LastPass.lnk" -force
Remove-Item "C:\Program Files (x86)\Online Services\LastPass" -recurse -force > $null 2>&1 
""
Stop-Process -name explorer -Force
Clear-Host
}

$OriginalPref = $ProgressPreference # Default is 'Continue'
$ProgressPreference = "SilentlyContinue"
Write-Host "Starting Windows 10 Bloatware remover"
start-allusers-bloatware-noxbox
start-basic-bloatware-remover
$ProgressPreference = $OriginalPref