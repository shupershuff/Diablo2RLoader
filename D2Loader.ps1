<# 
Author: Shupershuff
Usage: 
Happy for you to make any modifications this script for your own needs providing:
- Any variants of this script are never sold.
- Any variants of this script published online should always be open source.
- Any variants of this script are never modifed to enable or assist in any game altering or malicious behaviour including (but not limited to): Bannable Mods, Cheats, Exploits, Phishing
Purpose:
	Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle."
	Script will import account details from CSV. Alternatively you can run script parameters (see Github readme): - AccountUsername, -PW, -Region, -All, -Batch, -ManualSettingSwitcher
Instructions: See GitHub readme https://github.com/shupershuff/Diablo2RLoader

Notes:
- Multiple failed attempts (eg wrong Password) to sign onto a particular Realm via this method may temporarily lock you out. You should still be able to get in via the battlenet client if this occurs.

Servers:
 NA - us.actual.battle.net
 EU - eu.actual.battle.net
 Asia - kr.actual.battle.net
 
Changes since 1.7.0 (next version edits):
Moved customconfig into accounts.csv for nohd mod users. Script will auto remove this from config.xml and add a column in accounts.csv and assign the value that was previously in the config file.
Minor changes to some text outputs.
Added "KR" as a region parameter option.
Removed incorrect description in config.xml for gamgepath. Script will automatically fix the description.
Removed a couple of script statements that weren't adding any value.
Added a wee bit more reliability for the TZ checker. If it fails to read image after 11 seconds it will try another engine.
Account Usage Statistics (time account has been active) feature.
Fixed username and password Parameters not passing through to game client properly when being launched from parameters.

1.8.0 to do list

fix up update part of the script, was disabled for testing
Script MF Statistics
Fix whatever I broke or poorly implemented in 1.8.0 :)
#>

param($AccountUsername,$PW,$Region,$All,$Batch,$ManualSettingSwitcher) #used to capture paramters sent to the script, if anyone even wants to do that.
$CurrentVersion = "1.7.9"
$MenuRefreshRate = 30 #How often the script refreshes in seconds.
###########################################################################################################################################
# Script itself
###########################################################################################################################################
$host.ui.RawUI.WindowTitle = "Diablo 2 Resurrected Loader"
if (($Null -ne $PW -or $Null -ne $AccountUsername) -and ($Null -ne $Batch -or $Null -ne $All)){#if someone sends through incompatible parameters, prioritise $All and $Batch (in that order).
	$PW = $Null
	$AccountUsername = $Null
	if ($Null -ne $Batch -and $Null -ne $All){
		$Batch = $Null
	}
}
if ($Null -ne $AccountUsername){
	$ScriptArguments = "-accountusername $AccountUsername"  #this passes the value back through to the script when it's relaunched as admin mode.
}
if ($Null -ne $PW){
	$ScriptArguments += " -PW $PW"  
}
if ($Null -ne $Region){
	$ScriptArguments += " -region $Region"
}
if ($Null -ne $All){
	$ScriptArguments += " -all $All"
	$Script:OpenAllAccounts = $true
}
if ($Null -ne $Batch){
	$ScriptArguments += " -batch $Batch"
	$Script:OpenBatches = $True
} 
if ($Null -ne $ManualSettingSwitcher){
	$ScriptArguments += " -ManualSettingSwitcher $ManualSettingSwitcher"
	$Script:AskForSettings = $True
}

#check if username was passed through via parameter
if ($ScriptArguments -ne $Null){
	$Script:ParamsUsed = $true
}
Else {
	$Script:ParamsUsed = $false
}

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $ScriptArguments"  -Verb RunAs;exit } #run script as admin

#set window size
[console]::WindowWidth=77;
[console]::WindowHeight=48;
[console]::BufferWidth=[console]::WindowWidth
$X = [char]0x1b #escape character for ANSI text colors
$ProgressPreference = "SilentlyContinue"
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))
$Script:StartTime = Get-Date
Function ReadKey([string]$message=$Null) {
    $key = $Null
    $Host.UI.RawUI.FlushInputBuffer()
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host -NoNewLine $message
    }
    while($Null -eq $key) {
        if (($timeOutSeconds -eq 0) -or $Host.UI.RawUI.KeyAvailable) {
            $key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
            if ($key_.KeyDown) {
                $key = $key_
            }
        } else {
            Start-Sleep -m 250  # Milliseconds
        }
    }
    $EnterKey = 13
	if ($key_.VirtualKeyCode -ne $EnterKey-and -not ($Null -eq $key)) {
        Write-Host ("$X[38;2;255;165;000;22m" + "$($key.Character)" + "$X[0m") -NoNewLine
    }
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host "" # newline
    }       
    return $(
        if ($Null -eq $key -or $key.VirtualKeyCode -eq $EnterKey) {
            ""
        } else {
            $key.Character
        }
    )
}

Function ReadKeyTimeout([string]$message=$Null, [int]$timeOutSeconds=0, [string]$Default=$Null) {
    $key = $Null
    $Host.UI.RawUI.FlushInputBuffer()
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host -NoNewLine $message
    }
    $Counter = $timeOutSeconds * 1000 / 250
    while($Null -eq $key -and ($timeOutSeconds -eq 0 -or $Counter-- -gt 0)) {
        if (($timeOutSeconds -eq 0) -or $Host.UI.RawUI.KeyAvailable) {
            $key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
            if ($key_.KeyDown) {
                $key = $key_
            }
        } else {
            Start-Sleep -m 250  # Milliseconds
        }
    }
    $EnterKey = 13
    if ($key_.VirtualKeyCode -ne $EnterKey-and -not ($Null -eq $key)) {
        Write-Host ("$X[38;2;255;165;000;22m" + "$($key.Character)" + "$X[0m")
    }
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host "" # newline
    }
	Write-Host #prevent follow up text from ending up on the same line.
    return $(
        if ($Null -eq $key -or $key.VirtualKeyCode -eq $EnterKey) {
            $Default
        } else {
            $key.Character
        }
    )
}

#Check for updates
try {
	$tagList = Invoke-RestMethod "https://api.github.com/repos/Shupershuff/Diablo2RLoader/tags" -ErrorAction Stop
	if ([version[]]$taglist.Name.Trim('v') -gt $Script:CurrentVersion) {
		$ReleaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases/latest"
		Write-Host
		Write-Host " Update available! See Github for latest version and info:" -foregroundcolor Yellow
		Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m"
		Write-Host
		$ReleaseInfo.body -split "`n" | ForEach-Object {
			$_ = " " + $_
			if ($_[1] -eq "-") {#for any line starting with a dash
				 $DashFormat = ($_ -replace "(.{1,73})(\s+|$)", "`$1`n").trimend()
				 $DashFormat -split "`n" | ForEach-Object {
					if ($_[1] -eq "-") {#for any line starting with a dash
						$_
					}
					else {
						($_ -replace "(.{1,73})(\s+|$)", "   `$1`n").trimend()
					}
				}
			}
			else {
				($_ -replace "(.{1,75})(\s+|$)", "`$1`n ").trimend()
			}
		}
		Write-Host
		Write-Host
		Do {
			Write-Host " Would you like to update? $X[38;2;255;165;000;22mY$X[0m/$X[38;2;255;165;000;22mN$X[0m: " -nonewline
			$ShouldUpdate = ReadKey
			if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes" -or $ShouldUpdate -eq "n" -or $ShouldUpdate -eq "no"){
				$UpdateResponseValid = $True
			} Else {
				Write-Host
				Write-Host " Invalid response. Choose $X[38;2;255;165;000;22mY$X[0m $X[38;2;231;072;086;22mor$X[0m $X[38;2;255;165;000;22mN$X[0m." -ForegroundColor red
				Write-Host
			}
		} Until ($UpdateResponseValid -eq $True)
		if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes"){#if user wants to update script, download .zip of latest release, extract to temporary folder and replace old D2Loader.ps1 with new D2Loader.ps1
			New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
			$ZipURL = $ReleaseInfo.zipball_url #get zip download URL	
			$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
			Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
			$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
			Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
			$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
			($FolderPath + "\D2Loader.ps1")
			($Script:WorkingDirectory + "D2Loaders.ps1")
			Copy-Item -Path ($FolderPath + "\D2Loader.ps1") -Destination ($Script:WorkingDirectory + "\D2Loader.ps1")
			Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
			Write-Host "Updated :)"
			& ($Script:WorkingDirectory + "\D2Loader.ps1")
			exit
		}
		$ReleaseInfo = $Null
	}
}
Catch {
	Write-Host
	Write-Host " Couldn't check for updates. GitHub API limit may have been reached..." -foregroundcolor Yellow
	Start-Sleep -milliseconds 2750
}

#Import Config XML
try {
	$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig
	#Write-Host "Config imported successfully." -foregroundcolor green
}
Catch {
	Write-Host ""
	Write-Host " Config.xml Was not able to be imported. This could be due to a typo or a special character such as `'&`' being incorrectly used." -foregroundcolor red
	Write-Host " The error message below will show which line in the clientconfig.xml is invalid:" -foregroundcolor red
	Write-Host (" " + $PSitem.exception.message) -foregroundcolor red
	Write-Host ""
	Pause
	exit
}
#Perform some validation on config.xml. Helps avoid errors for people who may be on older versions of the script and are updating. Will look to remove all of this in a future update.
if (Select-String -path $Script:WorkingDirectory\Config.xml -pattern "multiple game installs"){#Sort out an incorrect description text that will have been in folks config.xml for some time. This description was never valid and was from when the setting switcher feature was being developed and tested.
	write-host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = ";;`t`tNote, if using multiple game installs \(to keep client specific config persistent for each account\), ensure these are referenced in the CustomGamePath field in accounts.csv."
	$Pattern += ";;`t`tOtherwise you can use a single install instead by linking the path below.-->;;"
	$NewXML = [string]::join(";;",($XML.Split("`r`n")))
	$NewXML = $NewXML -replace $Pattern, "-->;;"
	$NewXML = $NewXML -replace ";;","`r`n"
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	write-host " Corrected the description for GamePath in config.xml." -foregroundcolor Green
	Start-Sleep -milliseconds 1500
}
if ($Script:Config.CommandLineArguments -ne $Null){
	Write-Host
	Write-Host " Config option 'CommandLineArguments' is being moved to accounts.csv" -foregroundcolor Yellow
	Write-Host " This is to enable different CMD arguments per account." -foregroundcolor Yellow
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = ";;\t<!--Optionally add any command line arguments that you'd like the game to start with-->;;\t<CommandLineArguments>.*?</CommandLineArguments>;;"
	$NewXML = [string]::join(";;",($XML.Split("`r`n")))
	$NewXML = $NewXML -replace $Pattern, ""
	$NewXML = $NewXML -replace ";;","`r`n"
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	$AddCMDArgsToCSV = $True
	$OriginalCommandLineArguments = $Script:Config.CommandLineArguments
	Write-Host " CommandLineArguments has been removed from config.xml" -foregroundcolor green
	Start-Sleep -milliseconds 1500
}

if ($Script:Config.CheckForNextTZ -eq $Null){
	Write-Host
	Write-Host " Config option 'CheckForNextTZ' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to skip looking for NextTZ updates." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DefaultRegion>"
	$Replacement = "</DefaultRegion>`n`n`t<!--Choose whether or not TZ checker should look online for Next TZ updates.`n`t"
	$Replacement += "Choose False if Next TZ data is unreliable and you want faster updates for current TZ.-->`n`t<CheckForNextTZ>True</CheckForNextTZ>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	Pause
}

if ($Script:Config.ManualSettingSwitcherEnabled -eq $Null){#not to be confused with the AutoSettingSwitcher.
	Write-Host
	Write-Host " Config option 'ManualSettingSwitcherEnabled' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to allow you to manually select which " -foregroundcolor Yellow
	Write-Host " config file you want to use for each account when launching." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</SettingSwitcherEnabled>"
	$Replacement = "</SettingSwitcherEnabled>`n`n`t<!--Can be used standalone or in conjunction with the standard setting switcher above.`n`t"
	$Replacement +=	"This enables the menu option to enter 's' and manually pick between settings config files. Use this if you want to select Awesome or Poo graphics settings for an account.`n`t"
	$Replacement +=	"Any setting file with a number after it will not be an available option (eg settings1.json will not be an option).`n`t"
	$Replacement +=	"To make settings option, you can load from, call the file settings.<name>.json eg(settings.Awesome Graphics.json) which will appear as `"Awesome Graphics`" in the menu.-->`n`t"
	$Replacement +=	"<ManualSettingSwitcherEnabled>False</ManualSettingSwitcherEnabled>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	Pause
}

if ($Script:Config.TrackAccountUseTime -eq $Null){
	Write-Host
	Write-Host " Config option 'TrackAccountUseTime' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to track time played per account." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</ManualSettingSwitcherEnabled>"
	$Replacement = "</ManualSettingSwitcherEnabled>`n`n`t<!--This allows you to roughly track how long you've used each account while using this script. "
	$Replacement += "Choose False if you want to disable this.-->`n`t<TrackAccountUseTime>True</TrackAccountUseTime>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	Pause
}

if ($Script:Config.EnableBatchFeature -eq $Null){
	Write-Host
	Write-Host " Config option 'EnableBatchFeature' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to allow opening accounts in batches." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DefaultRegion>"
	$Replacement = "</DefaultRegion>`n`n`t<!--Enable the ability to open a group of accounts. EG if you have 5 accounts but only regularly use 3 of them that you want to open up all at once.`n`t"
	$Replacement += "Requires accounts.csv to be updated. See readme for more details.-->`n`t<EnableBatchFeature>False</EnableBatchFeature>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	Pause
}

if ($Script:Config.DisableOpenAllAccountsOption -eq $Null){
	Write-Host
	Write-Host " Config option 'DisableOpenAllAccountsOption' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to disable the functionality for opening all accounts." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DefaultRegion>"
	$Replacement = "</DefaultRegion>`n`n`t<!--Disable the functionality of being able to open all accounts at once. This is for any crazy people who have a lot of accounts and want to prevent accidentally opening all at once.-->`n`t"
	$Replacement += "<DisableOpenAllAccountsOption>False</DisableOpenAllAccountsOption>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	Pause
}
$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig #import config.xml again for any updates made by the above.

if ($Script:Config.EnableBatchFeature -eq $true -or $Null -ne $Batch){
	$Script:EnableBatchFeature = $True
	$BatchOption = "b" #specified here as well as in the ChooseAccounts section so that this works when being passed as a parameter
}
else {
	$BatchOption = "$Null"
}

#check if there's any missing config.xml options, if so user has out of date config file.
$AvailableConfigs = #add to this if adding features.
"GamePath",
"DefaultRegion",
"CheckForNextTZ",
"ShortcutCustomIconPath"

$BooleanConfigs = 
"ConvertPlainTextPasswords",
"ManualSettingSwitcherEnabled",
"DisableOpenAllAccountsOption",
"EnableBatchFeature",
"AskForRegionOnceOnly",
"CreateDesktopShortcut",
"ForceWindowedMode",
"SettingSwitcherEnabled"

$AvailableConfigs = $AvailableConfigs + $BooleanConfigs

if($Script:Config.CheckForNextTZ -ne $true -and $Script:Config.CheckForNextTZ -ne $false){#if CheckForNextTZ config is invalid, set to false
	$Script:CheckForNextTZ = $false
} Else {
	$Script:CheckForNextTZ = $Script:Config.CheckForNextTZ
}

$ConfigXMLlist = ($Config | Get-Member | Where-Object {$_.membertype -eq "Property" -and $_.name -notlike "#comment"}).name
Write-Host
foreach ($Option in $AvailableConfigs){
	if ($Option -notin $ConfigXMLlist){
		Write-Host "Config.xml file is missing a config option for $Option." -foregroundcolor yellow
		Start-Sleep 1
		pause
	}
}
if ($Option -notin $ConfigXMLlist){
	Write-Host
	Write-Host "Make sure to grab the latest version of config.xml from GitHub" -foregroundcolor yellow
	Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m"
	Write-Host
	Pause
}
if ($Config.GamePath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$Script:GamePath = $Config.GamePath.replace("`"","")
}
else {
	$Script:GamePath = $Config.GamePath
}
foreach ($ConfigCheck in $BooleanConfigs){#validate all configs that require "True" or "False" as the setting.
	if ($Config.$ConfigCheck -ne $null -and ($Config.$ConfigCheck -ne $true -and $Config.$ConfigCheck -ne $false)){#if config is invalid
		Write-Host " Config option '$ConfigCheck' is invalid." -foregroundcolor yellow 
		Write-Host " Ensure this is set to either True or False." -foregroundcolor yellow
		Write-Host
		pause
	}
}
if ($Config.ShortcutCustomIconPath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath.replace("`"","")
}
else {
	$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath
}
$DefaultRegion = $Config.DefaultRegion
$AskForRegionOnceOnly = $Config.AskForRegionOnceOnly
$CreateDesktopShortcut = $Config.CreateDesktopShortcut
$ConvertPlainTextPasswords = $Config.ConvertPlainTextPasswords

#Check Windows Game Path for D2r.exe is accurate.
if((Test-Path -Path "$GamePath\d2r.exe") -ne $True){ 
	Write-Host " Gamepath is incorrect. Looks like you have a custom D2r install location!" -foregroundcolor red
	Write-Host " Edit the GamePath variable in the config file." -foregroundcolor red
	Pause
	exit
}

# Create Shortcut
if ($CreateDesktopShortcut -eq $True){
	$DesktopPath = [Environment]::GetFolderPath("Desktop")
	$ScriptName = $MyInvocation.MyCommand.Name #in case someone renames the script.
	$Targetfile = "-ExecutionPolicy Bypass -File `"$WorkingDirectory\$ScriptName`""
	$ShortcutFile = "$DesktopPath\D2R Loader.lnk"
	$WScriptShell = New-Object -ComObject WScript.Shell
	$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
	$Shortcut.TargetPath = "powershell.exe" 
	$Shortcut.Arguments = $TargetFile
	if ($ShortcutCustomIconPath.length -eq 0){
		$Shortcut.IconLocation = "$Script:GamePath\D2R.exe"
	}
	Else {
		$Shortcut.IconLocation = $ShortcutCustomIconPath
	}
	$Shortcut.Save()
}

#Check SetText.exe setup
if((Test-Path -Path ($workingdirectory + '\SetText\SetText.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	Write-Host
	Write-Host "First Time run!" -foregroundcolor Yellow
	Write-Host
	Write-Host "SetText.exe not in .\SetText\ folder and needs to be built."
	if((Test-Path -Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe") -ne $True){#check that .net4.0 is actually installed or compile will fail.
		Write-Host ".Net v4.0 not installed. This is required to compile the Window Renamer for Diablo." -foregroundcolor red
		Write-Host "Download and install it from Microsoft here:" -foregroundcolor red
		Write-Host "https://dotnet.microsoft.com/en-us/download/dotnet-framework/net40" #actual download link https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net40-web-installer
		Pause
		exit
	}
	Write-Host "Compiling SetText.exe from SetText.bas..."
	& "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" -target:winexe -out:"`"$WorkingDirectory\SetText\SetText.exe`"" "`"$WorkingDirectory\SetText\SetText.bas`"" | out-null #/verbose  #actually compile the bastard
	if((Test-Path -Path ($workingdirectory + '\SetText\SetText.exe')) -ne $True){#if it fails for some reason and settext.exe still doesn't exist.
		Write-Host "SetText Could not be built for some reason :/"
		Write-Host "Exiting"
		Pause
		Exit
	}
	Write-Host "Successfully built SetText.exe for Diablo 2 Launcher script :)" -foregroundcolor green
	Start-Sleep -milliseconds 4000 #a small delay so the first time run outputs can briefly be seen
}

#Check Handle64.exe downloaded and placed into correct folder
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))
if((Test-Path -Path ($workingdirectory + '\Handle\Handle64.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	Write-Host "Handle64.exe is in the .\Handle\ folder. See instructions for more details on setting this up." -foregroundcolor red
	Pause
	exit
}

#Import CSV
if ($Script:AccountUsername -eq $Null){#If no parameters sent to script.
	try {
		$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv" #import all accounts from csv
	}
	Catch {
		Write-Host
		Write-Host " Accounts.csv does not exist. Make sure you create this and populate with accounts first." -foregroundcolor red
		Write-Host " Script exiting..." -foregroundcolor red
		Start-Sleep 5
		Exit
	}
}

if ($Script:AccountOptionsCSV -ne $Null){
	#check Accounts.csv has been updated and doesn't contain the example account.
	if ($Script:AccountOptionsCSV -match "yourbnetemailaddress"){
		Write-Host
		Write-Host "You haven't setup accounts.csv with your accounts." -foregroundcolor red
		Write-Host "Add your account details to the CSV file and run the script again :)" -foregroundcolor red
		Write-Host
		Pause
		exit
	}
	if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Batches" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.7.0. If batch column doesn't exist, add it
		# Column does not exist, so add it to the CSV data
		$Script:AccountOptionsCSV | ForEach-Object {
			$_ | Add-Member -NotePropertyName "Batches" -NotePropertyValue $null
		}
		# Export the updated CSV data back to the file
		$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
	}
	if (-not ($Script:AccountOptionsCSV | Get-Member -Name "CustomLaunchArguments" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.8.0. If CustomLaunchArguments column doesn't exist, add it
		# Column does not exist, so add it to the CSV data
		$Script:AccountOptionsCSV | ForEach-Object {
			$_ | Add-Member -NotePropertyName "CustomLaunchArguments" -NotePropertyValue $OriginalCommandLineArguments
		}
		# Export the updated CSV data back to the file
		$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
		Write-Host " Added CustomLaunchArguments column to accounts.csv." -foregroundcolor green
		Write-Host
		Start-Sleep -milliseconds 1200
		pause
	}
	if (-not ($Script:AccountOptionsCSV | Get-Member -Name "TimeActive" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.8.0. If TimeActive column doesn't exist, add it
		# Column does not exist, so add it to the CSV data
		$Script:AccountOptionsCSV | ForEach-Object {
			$_ | Add-Member -NotePropertyName "TimeActive" -NotePropertyValue $null
		}
		# Export the updated CSV data back to the file
		$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
	}
	if ($ConvertPlainTextPasswords -ne $false){
		#Check CSV for Plain text Passwords, convert to encryptedstrings and replace values in CSV
		$NewCSV = Foreach ($Entry in $AccountOptionsCSV) {
			if ($Entry.PWisSecureString.length -gt 0 -and $Entry.PWisSecureString -ne $False){#if nothing needs converting, make sure existing entries still make it into the updated CSV
				$Entry
			}
			if (($Entry.PWisSecureString.length -eq 0 -or $Entry.PWisSecureString -eq "no" -or $Entry.PWisSecureString -eq $false) -and $Entry.PW.length -ne 0){
				$Entry.PW = ConvertTo-SecureString -String $Entry.PW -AsPlainText -Force
				$Entry.PW = $Entry.PW | ConvertFrom-SecureString
				$Entry.PWisSecureString = "Yes"
				Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
				Start-Sleep -milliseconds 100
				$Entry
				$CSVupdated = $true
			}
			if ($Entry.PW.length -eq 0){#if csv has account details but password field has been left blank
				Write-Host
				Write-Host (" The account " + $Entry.AccountLabel + " doesn't yet have a password defined.") -foregroundcolor yellow
				Write-Host
				$Entry.PW = read-host -AsSecureString " Enter the Battle.net password for"$Entry.AccountLabel
				$Entry.PW = $Entry.PW | ConvertFrom-SecureString
				$Entry.PWisSecureString = "Yes"
				Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
				Start-Sleep -milliseconds 100
				$Entry
				$CSVupdated = $true
			}
		}
		if ($CSVupdated -eq $true){
			Try {
				$NewCSV | Export-CSV "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update CSV file
				Write-Host " Accounts.csv updated: Passwords have been secured." -foregroundcolor green
				Start-Sleep -milliseconds 4000
			}
			Catch {
				Write-Host
				Write-Host " Couldn't update Accounts.csv, probably because the file is open & locked." -foregroundcolor red
				Write-Host " Please close accounts.csv and run the script again!" -foregroundcolor red
				Write-Host
				Pause
				Exit
			}
		}
	}
}

#Set Region Array
$Script:ServerOptions = @(
	[pscustomobject]@{Option='1';region='NA';region_server='us.actual.battle.net'}#Americas
	[pscustomobject]@{Option='2';region='EU';region_server='eu.actual.battle.net'}#Europe
	[pscustomobject]@{Option='3';region='Asia';region_server='kr.actual.battle.net'}
)
Function SetQualityRolls {
	#Set item quality array for randomizing quote colours. A stupid addition to script but meh.
	$QualityArray = @(#quality and chances for things to drop based on 0MF values in D2r (I think?)
		[pscustomobject]@{Type='HighRune';Probability=1}
		[pscustomobject]@{Type='Unique';Probability=50}
		[pscustomobject]@{Type='SetItem';Probability=124}
		[pscustomobject]@{Type='Rare';Probability=200}
		[pscustomobject]@{Type='Magic';Probability=588}
		[pscustomobject]@{Type='Normal';Probability=19036}
	)
	if ($Script:GemActivated -eq $True){#small but noticeable MF boost
		$QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 16384  # New probability value
		}
	}
	Else {
		$QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 19036  # Original probability value
		}
	}
	if ($Script:CowKingActivated -eq $True){#big MF boost
		$QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 2048  # New probability value
		}
	}
	if ($Script:PGemActivated -eq $True){#huuge MF boost
		$QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 192  # New probability value
		}
	}
	$QualityHash = @{}; 
	foreach ($Object in $QualityArray | select-object type,probability){#convert PSOobjects to hashtable for enumerator
		$QualityHash.add($Object.type,$Object.probability) #add each PSObject to hash
	}
	$Script:ItemLookup = foreach ($Entry in $QualityHash.GetEnumerator()){
		[System.Linq.Enumerable]::Repeat($Entry.Key, $Entry.Value)
	}
}
function Magic {#ANSI text colour formatting for "magic" quotes. The variable $X (for the escape character) is defined earlier in the script.
    process { Write-Host "  $X[38;2;65;105;225;48;2;1;1;1;4m$_$X[0m" }
}
function SetItem {
    process { Write-Host "  $X[38;2;0;225;0;48;2;1;1;1;4m$_$X[0m"}
}
function Unique {
    process { Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4m$_$X[0m"}
}
function Rare {
    process { Write-Host "  $X[38;2;255;255;0;48;2;1;1;1;4m$_$X[0m"}
}
function Normal {
    process { Write-Host "  $X[38;2;255;255;255;48;2;1;1;1;4m$_$X[0m"}
}
function HighRune {
	process { Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4m$_$X[0m"}
}	

function quoteroll {#stupid thing to draw a random quote but also draw a random quality.
	$Quality = get-random $ItemLookup
	Write-Host
	$LeQuote = (Get-Random -inputobject $Script:quotelist)
	$ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
	$DesiredIndent = 2  # indent spaces
	$ChunkSize = $ConsoleWidth - $DesiredIndent
	[RegEx]::Matches($LeQuote, ".{$ChunkSize}|.+").Groups.Value | ForEach-Object {
		write-output $_ | &$Quality
	}
}

$Script:QuoteList =
"Stay a while and listen..",
"My brothers will not have died in vain!",
"Not even death can save you from me.",
"Good Day!",
"You have quite a treasure there in that Horadric Cube.",
"There's nothing the right potion can't cure.",
"Well, what the hell do you want? Oh, it's you. Uh, hi there.",
"Your souls shall fuel the Hellforge!",
"What do you need?",
"Your presence honors me.",
"I'll put that to good use.",
"Good to see you!",
"Looking for Baal?",
"All who oppose me, beware",
"Greetings",
"Ner. Ner! Nur. Roah. Hork, Hork.",
"Greetings, stranger. I'm not surprised to see your kind here.",
"There is a place of great evil in the wilderness.",
"I shall make weapons from your bones",
"I am overburdened",
"This magic ring does me no good.",
"The siege has everything in short supply...except fools.",
"Beware, foul demons and beasts.",
"They'll never see me coming.",
"I will cleanse this wilderness.",
"I shall purge this land of the shadow.",
"I hear foul creatures about.",
"Ahh yes, ruins, the fate of all cities.",
"I'm never gunna give you up, never gunna let you down - Griswold, 1996.",
"I have no grief for him. Oblivion is his reward.",
"The catapults have been silenced.",
"The staff of kings, you astound me!",
"When - or if - I get to Lut Gholein, I'm going to find the largest bowl`nof Narlant weed and smoke 'til all earthly sense has left my body.",
"I've just about had my fill of the walking dead.",
"Oh I hate staining my hands with the blood of foul Sorcerers!",
"Damn it, I wish you people would just leave me alone!",
"Beware! Beyond lies mortal danger for the likes of you!",
"Only the darkest Magics can turn the sun black.",
"You are too late! HAA HAA HAA",
"You now speak to Ormus. He was once a great mage, but now lives like a`nrat in a sinking vessel",
"I knew there was great potential in you, my friend. You've done a`nfantastic job.",
"Hi there. I'm Charsi, the Blacksmith here in camp. It's good to see some`nstrong adventurers around here.",
"Whatcha need?",
"Good day to you partner!",
"Moomoo, moo, moo. Moo, Moo Moo Moo Mooo.",
"Moo.",
"Moooooooooooooo",
"So cold and damp under the earth...",
"Good riddance, Blood Raven.",
"I shall meet death head-on.",
"The land here is dead and lifeless.",
"Let the gate be opened!",
"It is good to know that the sun shines again.",
"Maybe now the world will have peace.",
"Eternal suffering would be too brief for you, Diablo!",
"All that's left of proud Tristram are ghosts and ashes.",
"Ahh, the slow torture of caged starvation.",
"What a waste of undead flesh.",
"Good journey, Mephisto. Give my regards to the abyss.",
"My my, what a messy little demon!",
"Ah, the familiar scent of death.",
"What evil taints the light of the sun?",
"Light guide my way in this accursed place.",
"I shall honor Tal Rasha's sacrifice by destroying all the Prime Evils.",
"Oops...did I do that?",
"Death becomes you Andariel.",
"You dark mages are all alike, obsessed with power.",
"Planting the dead. How odd.",
"Oh no, snakes. I hate snakes.",
"Who would have thought that such primitive beings could cause so much `ntrouble.",
"Hail to you champion"

$BannerLogo = @"

  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%#%%%%%%%%%%%%%%%%%%%%
  %%%%%%%#%%%%%%%/%%%%%%%%%%%#%%%%%%%%%%%%##%%%%%%%%%#/##%%%%%%%%%%%%%%%%%%
  %%#(%/(%%%//(/*#%%%%%%%%%%%###%%%%#%%%%%###%%%%%###(*####%##%%%#%%*%%%%%%
  %%%( **/*///%%%%%%%%%###%%#(######%%#########%####/*/*.,#((#%%%#*,/%%%%%%
  %%%#*/.,*/,,*/#/%%%/(*#%%#*(%%(*/%#####%###%%%#%%%(///*,**/*/*(.&,%%%%%%%
  %%%%%%// % ,***(./*/(///,/(*,./*,*####%#*/#####/,/(((/.*.,.. (@.(%%%%%%%%
  %%%%%%%#%* &%%#..,,,,**,.,,.,,**///.*(#(*.,.,,*,,,,.,*, .&%&&.*%%%%%%%%%%
  %%%%%%%%%%%#.@&&%&&&%%%%&&%%,.(((//,,/*,,*.%&%&&&&&&&&%%&@%,#%%%%%%%%%%%%
  %%%%%%%%%%%%%(.&&&&&%&&&%(.,,*,,,,,.,,,,.,.*,%&&&%&&%&&&@*##%%%%%%%%%%%%%
  %%%%%%%%%%%%%%# @@@&&&&&(  @@&&@&&&&&&&&&*..,./(&&&&&&&&*####%%%%%%%%%%%%
  %%%%%%%%%%%%%%# &@@&&&&&(*, @@@&.,,,,. %@@&&*.,(%&&&&&&&/%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%#.&@@&&&&&(*, @@@@,((#&&%#.&@@&&.*#&&@&&&&/#%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%*&@@@&&&&#*, @@@@,*(#%&&%#,@@@&@,(%&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%*&@&&&%&&(,. @@@@,(%%%%%%#/,@@@& *#&&@&&%(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%*&&&@&%&&(,. @@@@,%&%%%%%%(.@@@@ /#&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%,&&&&&%%&(*, @@@@,&&&&&&&%//@@@@./%&&&&@&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%(*&&&&&&&%(,, @@@@,%&&&#(/*.@@@@&./%&&&&@&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%,&&&&&&&%(,, @@@@,/##/(// @@&@@,/#&&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%(,&&&&&&&%(,, @@@@.*,,..*@@@&&*./#&&%&&&&&(%#%%#%%%%%%%%%%%
  %%%%%%%%%%%#%%#.&&&&&%%#* @@@&&&@@@&&%&&&% */*%&&%#&&&&&/((#%%%%%%%%%%%%%
  %%%%%%%%(#//*/.&&&#%#%#.@&& ..,,****,,*//((/*#%%%####%%%#/#/#%%%%%%%%%%%%
  %%%%%##***.,**////*(//,&.*/***.*/%%#%/%#*.***/*/***//**/(((/.,*(//*/(##%%
  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
"@

Function JokeMaster {
	#If you're not going to leech and provide any damage value in the Throne Room then at least provide entertainment value right?
	Write-Host "  Copy these mediocre jokes into the game while doing Baal Comedy Club runs`r`n  to mislead everyone into believing you have a personality:"
	Write-Host
	$JokeProviderRoll = get-random -min 1 -max 3 #Randomly roll for a joke provider
	do {
		if ($JokeProviderRoll -eq 1){
			if ((Get-Date).Month -eq 12 -or ((Get-Date).Month -eq 10 -and (Get-Date).Day -ge 24 -and (Get-Date).Day -le 31)) {#If December or A week leading up to Halloween.
				if ((Get-Date).Month -eq 12){ #If December
					try {
						$Joke = Invoke-RestMethod -uri https://v2.jokeapi.dev/joke/Christmas -Method GET -header $headers -ErrorAction Stop #get absolutely punishing xmas jokes during December
						$JokeObtained = $true
					}
					catch {
						Write-Host "  Couldn't Reach v2.jokeapi.dev :( Trying alternate provider..." -foregroundcolor Red
						$JokeObtained = $false
						$JokeProviderRoll = 2
					}
				}
				Else {#else if it's around Halloweeen
					try {
						$Joke = Invoke-RestMethod -uri https://v2.jokeapi.dev/joke/Spooky -Method GET -header $headers -ErrorAction Stop #get absolutely punishing 'spooky' jokes in the week leading up to Halloween.
						$JokeObtained = $true
					}
					Catch {
						Write-Host "  Couldn't Reach v2.jokeapi.dev :( Trying alternate provider..." -foregroundcolor Red
						$JokeObtained = $false
						$JokeProviderRoll = 2
					}
				}
			}
			else {#For non seasonal jokes.
				try {
					$Joke = Invoke-RestMethod -uri "https://v2.jokeapi.dev/joke/Miscellaneous,Dark,Pun?blacklistFlags=racist,sexist" -Method GET -header $headers -ErrorAction Stop #exclude any racist/sexist jokes as we're not 75+ years old. Exclude IT/Programmer jokes that folk won't understand and also exclude, xmas & halloween themed jokes
					$JokeObtained = $true
				}
				catch {
					Write-Host "  Couldn't Reach v2.jokeapi.dev :( Trying alternate provider..." -foregroundcolor Red
					$JokeObtained = $false
					$JokeProviderRoll = 2
				}
			}
			if ($JokeObtained -eq $true){
				if ($Joke.Type -eq "twopart"){#some jokes are come through as two parters with two variables
					$JokeSetup = ($Joke.setup -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
					$JokeDelivery = ($Joke.delivery -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
					Write-Host "   $X[38;2;255;165;000;22m$JokeSetup$X[0m"
					Write-Host "   $X[38;2;255;165;000;22m$JokeDelivery$X[0m"
				}
				else {#else if single liner joke.
					$SingleJoke =  ($joke.joke -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
					Write-Host "   $X[38;2;255;165;000;22m$SingleJoke$X[0m"
				}
				$JokeProvider = "v2.jokeapi.dev"
			}
		}
		if ($JokeProviderRoll -eq 2){
			try {
				$Joke = Invoke-RestMethod -uri https://official-joke-api.appspot.com/random_joke -Method GET -header $headers -ErrorAction Stop #get absolutely punishing xmas jokes during December
				$JokeObtained = $true
				$JokeSetup = ($Joke.setup -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
				$JokeDelivery = ($Joke.punchline -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
				Write-Host "   $X[38;2;255;165;000;22m$JokeSetup$X[0m"
				Write-Host "   $X[38;2;255;165;000;22m$JokeDelivery$X[0m"
				$JokeProvider = "official-joke-api.appspot.com"
			}
			catch {
				$JokeProviderRoll = 1
				Write-Host "  Couldn't Reach official-joke-api.appspot.com :( Trying alternate provider..." -foregroundcolor Red
			}
		}
	} until ($JokeObtained -eq $true)
	Write-Host
	Write-Host "  Joke courtesy of $JokeProvider"
	Write-Host
	Write-Host " " -nonewline
	Pause
}
Function DClone {
	$headers = @{
		"D2R-Contact" = "placeholderemail@email.com"
		"D2R-Platform" = "GitHub"
		"D2R-Repo" = "https://github.com/shupershuff/Diablo2RLoader"
	}
	$uri = "https://d2runewizard.com/api/diablo-clone-progress/all?token=Pzttbnf1LduTScqauozCLQ"
	$D2RWDCloneResponse = Invoke-RestMethod -Uri $uri -Method GET -header $headers
	$LadderStatus = $D2RWDCloneResponse.servers | where-object {$_.Server -match "^ladder"} | select server,progress | sort server
	$NonLadderStatus = $D2RWDCloneResponse.servers | where-object {$_.Server -match "nonladder"} | select server,progress | sort server
	$CurrentStatus = $D2RWDCloneResponse.servers | select @{Name='Server'; Expression={$_.server}},@{Name='Progress'; Expression={$_.progress}} | sort server

	$DCloneLadderTable = New-Object -TypeName System.Collections.ArrayList
	$DCloneNonLadderTable = New-Object -TypeName System.Collections.ArrayList
	foreach ($Status in $CurrentStatus){
		$DCloneLadderInfo = New-Object -TypeName psobject
		$DCloneNonLadderInfo = New-Object -TypeName psobject
		if ($Status.server -match "^ladder"){
			$DCloneLadderInfo  | Add-Member -MemberType NoteProperty -Name LadderServer -Value $Status.server
			$DCloneLadderInfo  | Add-Member -MemberType NoteProperty -Name LadderProgress -Value $Status.progress
			[VOID]$DCloneLadderTable.Add($DCloneLadderInfo)
		}
		Else {
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderServer -Value $Status.server
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderProgress -Value $Status.progress
			[VOID]$DCloneNonLadderTable.Add($DCloneNonLadderInfo)
		}
	}
	$Count = 0
	Do {
		if ($Count -eq 0){
			Write-Host
			Write-Host "                         Current DClone Status:"
			Write-Host
			Write-Host " ##########################################################################"
			Write-Host " #             Ladder               |             Non-Ladder              #"
			Write-Host " ###################################|######################################"
			Write-Host " #  Server                  Status  |  Server                     Status  #"
			Write-Host " #----------------------------------|-------------------------------------#"
		}
		if ($Count -eq 3){Write-Host " #----------------------------------|-------------------------------------#"}
		$LadderServer = ($DCloneLadderTable.LadderServer[$Count]).tostring()
		do {
			$LadderServer = ($LadderServer + " ")
		}
		until ($LadderServer.length -ge 26)
		$NonLadderServer = ($DCloneNonLadderTable.NonLadderServer[$Count]).tostring()
		do {
			$NonLadderServer = ($NonLadderServer + " ")
		}
		until ($NonLadderServer.length -ge 29)
		Write-Host (" #  " + $LadderServer + " " + $DCloneLadderTable.LadderProgress[$Count] + "    |  " + $NonLadderServer + " " + $DCloneNonLadderTable.NonLadderProgress[$Count]+ "    #")
		$Count = $Count + 1
		if ($Count -eq 6){
			Write-Host " #                                  |                                     #"
			Write-Host " ##########################################################################"
			Write-Host
			Write-Host "   DClone Status provided D2runewizard.com"
			Write-Host
		}
	}
	Until ($Count -eq 6)	
	#$D2RWDCloneResponse.servers | select @{Name='Server'; Expression={$_.server}},@{Name='Progress'; Expression={$_.progress}} | sort server
	Pause
}
Function OCRCheckerWithTimeout {#Used to timeout OCR requests that take too long so that another attempt can be made with a different OCR engine.
	param (
		[ScriptBlock] $ScriptBlock,
		[int] $TimeoutSeconds
	)
	$Script:OCRSuccess = $null
	$job = Start-Job -ScriptBlock $ScriptBlock
	$timer = [Diagnostics.Stopwatch]::StartNew()

	while ($job.State -eq "Running" -and $timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
		Start-Sleep -Milliseconds 10
	}
	if ($job.State -eq "Running") {
		Stop-Job -Job $job
		Write-Host "  Command timed out." $Engine -foregroundcolor red
		$Script:OCRSuccess = $False
	}
	elseif ($job.State -eq "Completed") {
		$result = Receive-Job -Job $job
		$Script:OCRSuccess = $True
		#Write-Host "  Command completed successfully."
		$result
	}
	else {
		$Script:OCRSuccess = $False
		Write-Host " OCR failure or error, couldn't read next TZ details." -foregroundcolor red
	}
	Remove-Job -Job $job
}
Function TerrorZone {
	# Get the current time data was pulled
	$TimeDataObtained = (Get-Date -Format 'h:mmtt')
	#Find Current TZ
	Write-Host "  Finding Current TZ Name..."
	$FlyURI = "https://d2rapi.fly.dev/" #Get TZ from d2rapi.fly.dev
	$D2FlyTZResponse = Invoke-WebRequest -Uri $FlyURI -Method GET -header $headers
	$CurrentTZ = ($D2FlyTZResponse | Select-String -Pattern '(?<=<\/a> - )(.*?)(?=<br>)' -AllMatches | ForEach-Object { $_.Matches.Value })[0]	
	$CurrentTZProvider = "d2rapi.fly.dev"
	
	if ($CheckForNextTZ -ne $False){ #Find Next TZ image and convert to Text.
		Write-Host "  Finding Next TZ Name..."
		$NextTZProvider = "thegodofpumpkin"
		$75988Pool = @(#choose from pool of various connection details for OCR.Space to reduce chances of the free limit being reached. Helps future proof script if it ever becomes popular.
			"75988770481418K",
			"75988177170718K",
			"75988024702278K",
			"75988372514768K",
			"75988875571688K",
			"75988639433138K",
			"75988074942078K",
			"75988074727148K",
			"75988080905848K",
			"75988156278288K"
		)
		$tokreg = $75988Pool | Get-Random
		$script:D2ROCRref = $null
		for ($i = $tokreg.Length - 1; $i -ge 0; $i--) {
			$script:D2ROCRref += $tokreg[$i]
		}
		$Script:OCRAttempts = 0
		do { ## Try each OCR engine until success. I've noticed that OCR.Space engines sometimes go down causing the script not being able to detect current TZ. This gives it a bit more reliability, albeit at the cost of time for each additional attempt.
			$Script:OCRAttempts ++
			$NextTZOCR = OCRCheckerWithTimeout -ScriptBlock {
				$global:Engine = (2,1,3)[$using:OCRAttempts-1]	 #try engines in order of the best quality: Engine 2, then 1 then 3.
				#write-host "  Attempt:" $using:OCRAttempts -foregroundcolor Yellow #debug
				#write-host "  Using engine:" $Engine -foregroundcolor Yellow #debug
				(((Invoke-WebRequest -Uri ("https://api.ocr.space/parse/imageurl?apikey=" + $Using:D2ROCRref + "&filetype=png&isCreateSearchablePdf=false&OCREngine=" + $Engine + "&scale=true&url=https://thegodofpumpkin.com/terrorzones/terrorzone.png") -ErrorAction Stop).Content | ConvertFrom-Json).ParsedResults.ParsedText)
			} -TimeoutSeconds 11 #allow up to 11 seconds per request. Standard requests are around 3 seconds but as it's a free service it can sometimes be slow. Most of the time when it takes longer than 10 seconds to retrieve OCR it fails. Default timeout from OCR.Space (when an engine is down) is actually 30 seconds before an error is thrown.
		} until ($true -eq $Script:OCRSuccess -or ($Script:OCRAttempts -eq 3 -and $OCRSuccess -eq $false))
		if ($Script:OCRAttempts -eq 1){#For Engine 2
			$NextTZ = [regex]::Match($NextTZOCR.Replace("`n", ""), "(?<=Next TerrorZone(?:\s)?(?:\(s\))?(?:\s)?(?:\.|::) ).*").Value #Regex for engine 2
		} Elseif ($Script:OCRAttempts -eq 2){#For Engine 1
			$NextTZ = ($NextTZOCR.trim() -split "`n")[-1].replace("Next TerrorZone(s)","").replace(".","").replace(":","").trim()  #Regex for engine 1
		} Else {#For Engine 3
			$NextTZ = $nexttzocr -replace "\r?\n", " " -replace "(?s).*Next TerrorZone\Ss\S\s","" #Regex for engine 3
		}
		if ($OCRSuccess -eq $False){
			$FailMessage = "Was unable to pull next TZ details :("
		}
		if ($NextTZ -eq ""){#if next tz variable is empty due to OCR not working.
			$FailMessage = "OCR failure, couldn't read next TZ details :("
			$OCRSuccess -eq $False
		}
	}
	if ($Script:OCRSuccess -eq $False -or $CheckForNextTZ -eq $False){
		Write-Host
		Write-Host "   Current TZ is: "  -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta #print next tz. The regex code helps format the string into an indented new line for longer TZ's.
		if ($CheckForNextTZ -ne $False){
			$FailMessage += "`n                             Attempts made: $Script:OCRAttempts`n                             Token used:      $tokreg"
			Write-Host "   Next TZ info unavailable: $FailMessage" -ForegroundColor Red
		}
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  Current TZ pulled from    $CurrentTZProvider"
		Write-Host
		Pause
	}
	else {
		Write-Host
		Write-Host "   Current TZ is:  " -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host "   Next TZ is:     " -nonewline;Write-Host ($NextTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  Current TZ pulled from:    $CurrentTZProvider"
		Write-Host "  Next TZ pulled from:       $NextTZProvider"
		Write-Host
		Pause
	}
}

Function Killhandle {#kudos the info in this post to save me from figuring it out: https://forums.d2jsp.org/topic.php?t=90563264&f=87
	& "$PSScriptRoot\handle\handle64.exe" -accepteula -a -p D2R.exe > $PSScriptRoot\d2r_handles.txt
	$proc_id_populated = ""
	$handle_id_populated = ""
	foreach($Line in Get-Content $PSScriptRoot\d2r_handles.txt) {
		$proc_id = $Line | Select-String -Pattern '^D2R.exe pid\: (?<g1>.+) ' | %{$_.Matches.Groups[1].value}
		if ($proc_id){
			$proc_id_populated = $proc_id
		}
		$handle_id = $Line | Select-String -Pattern '^(?<g2>.+): Event.*DiabloII Check For Other Instances' | %{$_.Matches.Groups[1].value}
		if ($handle_id){
			$handle_id_populated = $handle_id
		}
		if($handle_id){
			#Write-Host "Closing" $proc_id_populated $handle_id_populated
			& "$PSScriptRoot\handle\handle64.exe" -p $proc_id_populated -c $handle_id_populated -y
		}
	}
}

Function CheckActiveAccounts {#Note: only works for accounts loaded by the script
	#check if there's any open instances and check the game title window for which account is being used.
	try {	
		$Script:ActiveIDs = $Null
		$D2rRunning = $false
		$Script:ActiveIDs = New-Object -TypeName System.Collections.ArrayList
		$Script:ActiveIDs = (Get-Process | Where {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "- Diablo II: Resurrected -"} | Select-Object MainWindowTitle).mainwindowtitle.substring(0,1) #find all diablo 2 game windows and pull the account ID from the title
		$Script:D2rRunning = $true
		#Write-Host "Running Instances."
	}
	catch {#if the above fails then there are no running D2r instances.
		$Script:D2rRunning = $false
		#Write-Host "No Running Instances."
		$Script:ActiveIDs = ""
	}
	if ($Script:D2rRunning -eq $True){
		$Script:ActiveAccountsList = New-Object -TypeName System.Collections.ArrayList
		foreach ($ActiveID in $ActiveIDs){#Build list of active accounts that we can omit from being selected later
			$ActiveAccountDetails = $Script:AccountOptionsCSV | where-object {$_.id -eq $ActiveID}
			$ActiveAccount = New-Object -TypeName psobject
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ID -Value $ActiveAccountDetails.ID
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name AccountName -Value $ActiveAccountDetails.accountlabel
			[VOID]$Script:ActiveAccountsList.Add($ActiveAccount)
		}
	}
	else {
		$Script:ActiveAccountsList = $Null
	}
}
Function DisplayActiveAccounts {
	Write-Host
	if ($Script:EnableBatchFeature -eq $true){
		$LongestAccountLabelLength = $Script:AccountOptionsCSV.accountlabel | ForEach-Object { $_.Length } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum #find out how many batches there are so table can be properly indented.
		while ($AccountHeaderIndent.length -lt ($LongestAccountLabelLength -13)){#indent the header for batches based on how long the longest account name is
			$AccountHeaderIndent = $AccountHeaderIndent + " "
		}
	}
	if ($Script:ActiveAccountsList.id -ne ""){#if batch feature is enabled add a column to display batch number(s)
		if ($Script:Config.TrackAccountUseTime -eq $true){
			$PlayTimeHeader = "Hours Played   "
		}
		if ($Script:EnableBatchFeature -eq $true){
			$BatchesHeader = (""+ $AccountHeaderIndent + "Batch(es)")
		}
		Write-Host ("  ID   Region   " + $PlayTimeHeader + "Account Label   " + $BatchesHeader) #Header
	}
	else {
		Write-Host "  ID   Account Label"
	}
	$Pattern = "(?<=- \w+ \()([a-z]+)"#Regex pattern to pull the region characters out of the window title.
	foreach ($AccountOption in $Script:AccountOptionsCSV){
		$RegionDisplayPostIndent = ""
		$RegionDisplayPreIndent = ""
		if ($AccountOption.ID.length -ge 2){#keep table formatting looking lovely if some crazy user has 10+ accounts.
			$IDIndent = ""
		}
		else {
			$IDIndent = " "
		}
		if ($Script:EnableBatchFeature -eq $true){
			$Batches = $AccountOption.batches
			$AccountIndent = ""
			if ($LongestAccountLabelLength -lt 13){#If longest account labels are shorter than 13 characters, set this variable to the minimum (13) so indenting works properly.
				$LongestAccountLabelLength = 13
			}
			while (($AccountOption.accountlabel.length + $AccountIndent.length) -le $LongestAccountLabelLength){#keep adding indents until account label plus the indents matches the longest account name. Keeps table nice and neat.
				$AccountIndent = $AccountIndent + " "
			}
		}
		if ($Script:Config.TrackAccountUseTime -eq $true){
			try {
				$AcctPlayTime = (" " + ("{0:N2}" -f [TimeSpan]::Parse($AccountOption.TimeActive).totalhours) + "     ").replace(",","") #Convert timeactive string to time. Round Current play time to 2 decimal places. Remove comma if some supernerd puts in over 10000 hours.
			}
			catch {#if account hasn't been opened yet.
				$AcctPlayTime = " 0     "
			}
			if ($AcctPlayTime.length -lt 15){#formatting. Depending on the amount of characters for this variable push it out until it's 15 chars long.
				while ($AcctPlayTime.length -lt 15){
					$AcctPlayTime = " " + $AcctPlayTime
				}
			}
		}
		if ($AccountOption.id -in $Script:ActiveAccountsList.id){#if account is currently active
			$Windowname = (Get-Process | Where {$_.processname -eq "D2r" -and $_.MainWindowTitle -match ($AccountOption.id + " - Diablo II: Resurrected -")} | Select-Object MainWindowTitle).mainwindowtitle #Check active game instances to see which accounts are active. As this is based on checking window titles, this will only work for accounts opened from the script
			$CurrentRegion = [regex]::Match($WindowName, $Pattern).value #Check which region aka realm the active account is connected to.
			if ($CurrentRegion -eq "US"){$CurrentRegion = "NA"; $RegionDisplayPreIndent = " "; $RegionDisplayPostIndent = " "}
			if ($CurrentRegion -eq "KR"){$CurrentRegion = "Asia"}
			if ($CurrentRegion -eq "EU"){$CurrentRegion = "EU"; $RegionDisplayPreIndent = " "; $RegionDisplayPostIndent = " "}
			Write-Host ("   " + $AccountOption.ID + "    "  + $RegionDisplayPreIndent + $CurrentRegion + $RegionDisplayPostIndent + "    " + $AcctPlayTime  + $AccountOption.accountlabel + " - Account Active.") -foregroundcolor yellow
		}
		else {#if account isn't currently active
			Write-Host ("  " + $IDIndent + $AccountOption.ID + "      -     " + $AcctPlayTime + $AccountOption.accountlabel + "  " + $AccountIndent + $Batches) -foregroundcolor green
		}
	}
}

Function Menu {
	cls
	if($Script:ScriptHasBeenRun -eq $true){
		$Script:AccountUsername = $Null
		Write-Host "Account previously opened was:" -foregroundcolor yellow -backgroundcolor darkgreen
		$Lastopened = @(
			[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$Script:LastRegion}
		)
		Write-Host " " -NoNewLine
		Write-Host ("Account:  " + $Lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen 
		Write-Host " " -NoNewLine
		Write-Host "Region:  " $Lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
	}
	Else {
		Write-Host (" You have quite a treasure there in that Horadric multibox script v" + $Currentversion)
	}
	Write-Host $BannerLogo -foregroundcolor yellow
	QuoteRoll
	if ($Batch -eq $Null -and $Script:OpenAllAccounts -ne $True){#go through normal account selection screen if script hasn't been launched with parameters that already determine this.
		ChooseAccount
	}
	Else {
		CheckActiveAccounts
		$Script:PWmanualset = $False
		$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
		foreach ($AccountOption in $Script:AccountOptionsCSV){
			if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
				$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
			}
		}
	}
	if ($Batch -ne $Null -or $Script:OpenBatches -eq $true){
		$Script:AcceptableBatchIDs = $Null #reset value
		foreach ($ID in $Script:AccountOptionsCSV){
			if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
				$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',')
				$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',')
			}
		}
		$AcceptableBatchValues = ($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort) #Unique list of available batches that can be opened
		do {
			if ($Batch -ne $Null -and $Batch -notin $AcceptableBatchValues){#if batch specified in the parameter isn't valid
				$Script:BatchToOpen = $Batch
				$Batch = $null
				DisplayActiveAccounts
				Write-Host
				Write-Host " Batch specified in Parameter is either incorrect or all accounts in that" -foregroundcolor Yellow
				Write-Host " batch are already open. Adjust your parameter or manually specify below." -foregroundcolor Yellow
				Write-Host
			}
			if ($Batch -ne $Null -and $Batch -in $AcceptableBatchValues){
				$Script:BatchToOpen = $Batch
			}
			Else {
				Write-Host " Which Batch of accounts would you like to open: " -nonewline
				foreach ($Value in $AcceptableBatchValues){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
					if ($Value -ne $AcceptableBatchValues[-1]){
						Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
						if ($Value -ne $AcceptableBatchValues[-2]){Write-Host ", " -nonewline}
					}
					else {
						Write-Host " or $X[38;2;255;165;000;22m$Value$X[0m"
					}
				}
				if ($Batch -eq $Null){
					Write-Host " Or Press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				}
				$Script:BatchToOpen = readkey
				Write-Host
				Write-Host
			}
			if ($BatchToOpen -notin $AcceptableBatchValues + "c"){
				Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
				Write-Host
				$Script:BatchToOpen = ""
			}
		} until ($Script:BatchToOpen -in $AcceptableBatchValues + "c")
		if ($BatchToOpen -ne "c"){
			$Script:BatchedAccountIDsToOpen = New-Object -TypeName System.Collections.ArrayList
			foreach ($ID in $Script:AccountOptionsCSV){
				if ($ID.id -in $Script:AcceptableBatchIDs -and $ID.batches.split(',') -contains $Script:BatchToOpen.tostring()){
					$Script:BatchedAccountIDsToOpen = $Script:BatchedAccountIDsToOpen + $ID.id
				}
			}
		}
		Else {
			$Script:OpenBatches = $False
		}
	}
	if($Script:ParamsUsed -eq $false -and ($Script:RegionOption.length -ne 0 -or $Script:Region.length -ne 0)){
		if ($Script:AskForRegionOnceOnly -ne $true){
			$Script:Region = ""
			$Script:RegionOption = ""
			#Write-Host "region reset" -foregroundcolor yellow #debug
		}
	}
	if ($Script:BatchToOpen -ne "c"){#get next region unless the cancel option has been specified.
		if ($Script:region.length -eq 0){#if no region parameter has been set already.
			ChooseRegion
		}
		Else {#if region parameter has been set already.
			if ($Script:region -eq "NA" -or $Script:region -eq "us" -or $Script:region -eq 1){$Script:region = "us.actual.battle.net"}
			if ($Script:region -eq "EU" -or $Script:region -eq 2){$Script:region = "eu.actual.battle.net"}
			if ($Script:region -eq "Asia" -or $Script:region -eq "As" -or $Script:region -eq "KR" -or $Script:region -eq 3){$Script:region = "kr.actual.battle.net"}
			if ($Script:region -ne "us.actual.battle.net" -and $Script:region -ne "eu.actual.battle.net" -and $Script:region -ne "kr.actual.battle.net"){
				Write-Host " Region not valid. Please choose region" -foregroundcolor red
				ChooseRegion
			}
		}
	}
	if ($Script:OpenAllAccounts -eq $True){
		Write-Host
		Write-Host " Opening all accounts..."
		foreach ($ID in $Script:AcceptableValues){
			$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $ID}
			$Script:AccountID = $ID
			if ($id -eq $Script:AcceptableValues[-1]){
				$Script:LastAccount = $true
			}
			Write-Host
			Write-Host " Opening Account: $ID" 
			Processing
		}
		$Script:LastAccount = $False
		$Script:OpenAllAccounts = $False
		if ($Script:ParamsUsed -ne $True){
			Menu
		}
	}
	Else {
		if ($Script:OpenBatches -eq $True -and $Script:RegionOption -ne "c"){
			Write-Host
			if ($Script:BatchedAccountIDsToOpen.count -ge 2){# a bunch of lines just to make the script cleanly display which accounts are being opened.
				if ($Script:BatchedAccountIDsToOpen.count -eq 2){
					$BatchAccountsCommaSeparated = ($Script:BatchedAccountIDsToOpen[-2] + " and " + $Script:BatchedAccountIDsToOpen[-1])
				}
				else {
					$BatchAccountsCommaSeparated = $Script:BatchedAccountIDsToOpen[0..($Script:BatchedAccountIDsToOpen.count -3)] -join ", "
					$BatchAccountsCommaSeparated += (", " + $Script:BatchedAccountIDsToOpen[-2] + " and " + $Script:BatchedAccountIDsToOpen[-1])
				}
			}
			else {
				$BatchAccountsCommaSeparated = $Script:BatchedAccountIDsToOpen
			}
			Write-Host " Opening accounts $BatchAccountsCommaSeparated from batch $BatchToOpen..."
			foreach ($ID in $Script:BatchedAccountIDsToOpen){
				$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $ID}
				$Script:AccountID = $ID
				if ($id -eq $Script:BatchedAccountIDsToOpen[-1]){
					$Script:LastAccount = $true
				}
				Write-Host
				Write-Host " Opening Account: $ID..." 
				Processing
			}
			$Script:LastAccount = $False
			$Script:OpenBatches = $False
			Start-Sleep -milliseconds 500
			if ($Script:ParamsUsed -ne $True){
				Menu
			}
		}
		Else {
			if ($Script:BatchToOpen -ne "c"){
				Processing
			}
			if ($Script:ParamsUsed -ne $true){
				Menu
			}
			else {
				Write-Host "I'm quitting LOL"
				exit
			}
		}
	}
}
Function ChooseAccount {
	if ($Null -ne $Script:AccountUsername){ #if parameters have already been set.
		$Script:AccountOptionsCSV = @(
			[pscustomobject]@{PW=$Script:PW;acct=$Script:AccountUsername}
		)
	}
	else {#if no account parameters have been set already		
		do {
			if ($Script:AccountID -eq "t"){
				TerrorZone
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "d"){
				DClone
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "j"){
				JokeMaster
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "s"){
				if ($Script:AskForSettings -eq $True){
					Write-Host "  Manual Setting Switcher Disabled." -foregroundcolor Green
					$Script:AskForSettings = $False
				}
				else {
					Write-Host "  Manual Setting Switcher Enabled." -foregroundcolor Green
					$Script:AskForSettings = $True
				}
				Start-Sleep -milliseconds 1550
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "g"){#silly thing to replicate in game chat gem.
				if ($Script:GemActivated -ne $True){
					$GibberingGemstone = get-random -minimum 0 -maximum 4095
					if($GibberingGemstone -eq 69){#nice
						Write-Host "  Perfect Gem Activated" -ForegroundColor magenta
						Write-Host
						Write-Host "     OMG!" -foregroundcolor green
						$Script:PGemActivated = $True
						SetQualityRolls
						Start-Sleep -milliseconds 3750
					}
					else {
						if($GibberingGemstone -in 16..32){
							Write-Host "  $X[38;2;165;146;99;22mMoooooooo!$X[0m"
							$Script:CowKingActivated = $True
							SetQualityRolls
							Start-Sleep -milliseconds 850
						}
						else {
							Write-Host "  Gem Activated" -ForegroundColor magenta
						}
					}
					$Script:GemActivated = $True
					SetQualityRolls
				}
				Else {
					Write-Host "  Gem Deactivated" -ForegroundColor magenta
					$Script:GemActivated = $False
					SetQualityRolls
				}
				Start-Sleep -milliseconds 850
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "r"){#refresh
				cls
				if($Script:ScriptHasBeenRun -eq $true){
					Write-Host "Account previously opened was:"  -foregroundcolor yellow -backgroundcolor darkgreen
					$Lastopened = @(
						[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$Script:region}#Americas
					)
					Write-Host " " -NoNewLine
					Write-Host ("Account:  " + $Lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen 
					Write-Host " " -NoNewLine
					Write-Host "Region:  " $Lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
				}
				Write-Host $BannerLogo -foregroundcolor yellow
				QuoteRoll	
			}
			CheckActiveAccounts
			DisplayActiveAccounts
			if ($Script:Config.TrackAccountUseTime -eq $True){
				$OpenD2LoaderInstances = Get-WmiObject -Class Win32_Process | Where-Object { $_.name -eq "powershell.exe" -and $_.commandline -match "d2loader.ps1"} | select name,processid,creationdate | sort creationdate
				if ($OpenD2LoaderInstances.length -gt 1){#If there's more than 1 D2loader.ps1 script open, close until there's only 1 open to prevent the time played accumulating too quickly.
					Stop-Process -id $OpenD2LoaderInstances[0].processid -force #Closes oldest running d2loader script
				}
				$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv"
				foreach ($AccountID in $Script:ActiveAccountsList.id |sort){ #$Script:ActiveAccountsList.id
					$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date) #work out elapsed time to add to accounts.csv				
					$AccountToUpdate = $Script:AccountOptionsCSV | Where-Object {$_.ID -eq $accountID}
					if ($AccountToUpdate) {
						try {
							$AccountToUpdate.TimeActive = [TimeSpan]::Parse($AccountToUpdate.TimeActive) + $AdditionalTimeSpan
						}
						Catch {
							$AccountToUpdate.TimeActive = $AdditionalTimeSpan
						}
					}
					try {
						$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update accounts.csv with the new time played.
					}
					Catch {
						$WriteError = $true 
					}
				}
				$Script:StartTime = Get-Date #restart timer.
				if ($True -eq $WriteError){
					Write-host
					Write-host "  Couldn't update accounts.csv with playtime info." -ForegroundColor Red
					Write-host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
					start-sleep -milliseconds 2500
					$WriteError = $False
				}
			}
			$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
			foreach ($AccountOption in $Script:AccountOptionsCSV){
				if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
					$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
				}
			}
			$accountoptions = ($Script:AcceptableValues -join  ", ").trim()
			do {
				Write-Host
				$Script:OpenAllAccounts = $False
				if ($accountoptions.length -gt 0){#if there are unopened account options available
					if ($Script:Config.ManualSettingSwitcherEnabled -eq $true){
						$ManualSettingSwitcherOption = "s"
						$ManualSettingSwitcherMenuText = "'$X[38;2;255;165;000;22ms$X[0m' to toggle the Manual Setting Switcher,"
					}
					if ($Script:Config.DisableOpenAllAccountsOption -eq $true){#if end user has a stupid amount of accounts and wants to prevent accidentally opening up over 9000 accounts
						$AllAccountMenuText = ""
						$AllAccountMenuTextNoBatch = ""
					}
					else {
						$AllOption = "a"
						$AllAccountMenuText = "'$X[38;2;255;165;000;22ma$X[0m' to open All accounts, "	
						$AllAccountMenuTextNoBatch = " or '$X[38;2;255;165;000;22ma$X[0m' for All."
					}
					if ($Script:EnableBatchFeature -ne $true){
						Write-Host (" Select which account to sign into: " + "$X[38;2;255;165;000;22m$accountoptions$X[0m" + $AllAccountMenuTextNoBatch)
						Write-Host " Alternatively choose from the following menu options:"
						$BatchMenuText = ""
					}
					else {
						$Script:BatchToOpen = $null
						$BatchMenuText = "'$X[38;2;255;165;000;22mb$X[0m' to open a Batch of accounts,"
						$Script:AcceptableBatchIDs = $Null #reset value
						$AcceptableBatchValues = $null
						foreach ($ID in $Script:AccountOptionsCSV){
							if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
								$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',')
								$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',')
							}
						}
						$AcceptableBatchValues = ($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort) #Unique list of available batches that can be opened
						if ($AcceptableBatchValues -eq $null){
							$BatchOption = ""
							$BatchMenuText = ""
						}
						Else {
							$BatchOption = "b"
						}
						Write-Host " Enter the ID# of the account you want to sign into."
						Write-Host " Alternatively choose from the following menu options:"
						Write-Host ("  " + $AllAccountMenuText + $BatchMenuText)
					}
				}
				else {#if there aren't any available options, IE all accounts are open
					$AllOption = $Null
					$BatchOption = $Null
					Write-Host " All Accounts are currently open!" -foregroundcolor yellow
				}
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, '$X[38;2;255;165;000;22mt$X[0m' for TZ info, '$X[38;2;255;165;000;22md$X[0m' for DClone status, '$X[38;2;255;165;000;22mj$X[0m' for jokes,"
				Write-Host "  $ManualSettingSwitcherMenuText or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: " -nonewline
				$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
				if ($Script:AccountID -notin ($Script:AcceptableValues + "x" + "r" + "t" + "d" + "g" + "j" + $ManualSettingSwitcherOption + $AllOption + $BatchOption) -and $Null -ne $Script:AccountID){
					if ($Script:AccountID -eq "a" -and $Script:Config.DisableOpenAllAccountsOption -ne $true){
						Write-Host " Can't open all accounts as all of your accounts are already open doofus!" -foregroundcolor red
					}
					else {
						Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
					}
					$Script:AccountID = $Null
				}
			} until ($Null -ne $Script:AccountID)
			if ($Null -ne $Script:AccountID){
				if($Script:AccountID -eq "x"){
					Write-Host
					Write-Host "Good day to you partner :)"
					Start-Sleep -milliseconds 386
					Exit
				}
				$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
			}
		} until ($Script:AccountID -ne "r" -and $Script:AccountID -ne "t" -and $Script:AccountID -ne "d" -and $Script:AccountID -ne "g" -and $Script:AccountID -ne "j" -and $Script:AccountID -ne "s")
		if ($Script:AccountID -eq "a"){
			$Script:OpenAllAccounts = $True
		}
		if ($Script:AccountID -eq "b"){
			$Script:OpenBatches = $True
		}
	}
	if (($Null -ne $Script:AccountUsername -and ($Null -eq $Script:PW -or "" -eq $Script:PW) -or ($Script:AccountChoice.id.length -gt 0 -and $Script:AccountChoice.PW.length -eq 0))){#This is called when params are used but the password wasn't entered. Not used for -all or -batch
		$SecuredPW = read-host -AsSecureString " Enter the Battle.net password for $Script:AccountUsername"
		$Bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecuredPW)
		$Script:PW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($Bstr)
		$Script:PWmanualset = $true
	}
	else {
		$Script:PWmanualset = $false
	}
}
Function ChooseRegion {#AKA Realm. Not to be confused with the actual Diablo servers that host your games, which are all over the world :)
	Write-Host " Available regions are:"
	Write-Host "  Option  Region  Server Address"
	Write-Host "  ------  ------  --------------"
	foreach ($Server in $ServerOptions){
		if ($Server.region.length -eq 2){$Regiontablespacing = " "}
		if ($Server.region.length -eq 4){$Regiontablespacing = ""}
		Write-Host ("    " + $Server.option + "      " + $Regiontablespacing + $Server.region + $Regiontablespacing + "   " + $Server.region_server) -foregroundcolor green
	}
	Write-Host	
		do {
			Write-Host " Please select a region: $X[38;2;255;165;000;22m1$X[0m, $X[38;2;255;165;000;22m2$X[0m or $X[38;2;255;165;000;22m3$X[0m"
			Write-Host (" Alternatively select '$X[38;2;255;165;000;22mc$X[0m' to cancel or press enter for the default (" + $Script:defaultregion + "-" + ($Script:ServerOptions | Where-Object {$_.option -eq $Script:defaultregion}).region + "): ") -nonewline
			$Script:RegionOption = ReadKey
			Write-Host
			if ("" -eq $Script:RegionOption){
				$Script:RegionOption = $Script:DefaultRegion #default to NA
			}
			else {
				$Script:RegionOption = $Script:RegionOption.tostring()
			}
			if($Script:RegionOption -notin $Script:ServerOptions.option + "c"){
				Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
				$Script:RegionOption = ""
			}
		} until ("" -ne $Script:RegionOption)
	if ($Script:RegionOption -in 1..3 ){# if value is 1,2 or 3 set the region string.
		$Script:region = ($ServerOptions | where-object {$_.option -eq $Script:RegionOption}).region_server
		$Script:LastRegion = $Script:Region
	}
}

Function Processing {
	if ($Script:RegionOption -ne "c"){
		if (($Script:PW -eq "" -or $Script:PW -eq $Null) -and $Script:PWmanualset -eq 0){
			$Script:PW = $Script:AccountChoice.PW.tostring()
		}
		if (($ConvertPlainTextPasswords -ne $false -and $Script:ParamsUsed -ne $true) -or ($Script:ParamsUsed -eq $true -and ($Script:OpenBatches -eq $True -or $Script:OpenAllAccounts -eq $True))){#Convert password if it's enabled in config and script is being run normally *OR* Convert password if script is being run from paramters using either -all batch or -all (but not if -username is used instead)
			$Script:acct = $Script:AccountChoice.acct.tostring()
			$EncryptedPassword = $PW | ConvertTo-SecureString
			$PWobject = New-Object System.Management.Automation.PsCredential("N/A", $EncryptedPassword)
			$Script:PW = $PWobject.GetNetworkCredential().Password
		}
		else {
			if ($Script:AccountID -eq $Null){
				$Script:acct = $Script:AccountUsername
				$Script:AccountID = "1"
			}
		}
		try {
			$Script:AccountFriendlyName = $Script:AccountChoice.accountlabel.tostring()
		}
		Catch {
			$Script:AccountFriendlyName = $Script:AccountUsername
		}	

		#Open diablo with parameters
			# IE, this is essentially just opening D2r like you would with a shortcut target of "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected\D2R.exe" -username <yourusername -password <yourPW> -address <SERVERaddress>
		#if ($Script:AccountID-eq "1" -or $Script:AccountID -eq "2" -or $Script:AccountID -eq "3"){$ModArguments = "-mod nohd"}
		#$arguments = (" -username " + $Script:acct + " -password " + $Script:PW +" -address " + $Script:Region + " " + $Config.CommandLineArguments + " " + $modarguments).tostring()
		$arguments = (" -username " + $Script:acct + " -password " + $Script:PW +" -address " + $Script:Region + " " + $Script:AccountChoice.CustomLaunchArguments).tostring()
		if ($Config.ForceWindowedMode -eq $true){#starting with forced window mode sucks, but someone asked for it.
			$arguments = $arguments + " -w"
		}
		$Script:PW = $Null
		
		#Switch Settings file to load D2r from.
		if ($Config.SettingSwitcherEnabled -eq $True -and $Script:AskForSettings -ne $True){#if user has enabled the auto settings switcher.
			$SettingsProfilePath = ("C:\Users\" + $Env:UserName + "\Saved Games\Diablo II Resurrected\")
			$SettingsJSON = ($SettingsProfilePath + "Settings.json")
			foreach ($id in $Script:AccountOptionsCSV){#create a copy of settings.json file per account so user doesn't have to do it themselves
				if ((Test-Path -Path ($SettingsProfilePath+ "Settings" + $id.id +".json")) -ne $true){#if somehow settings<ID>.json doesn't exist yet make one from the current settings.json file.
					try {
						Copy-Item $SettingsJSON ($SettingsProfilePath + "Settings"+ $id.id + ".json") -ErrorAction Stop
					}
					catch {
						Write-Host
						Write-Host " Couldn't find settings.json in $SettingsProfilePath" -foregroundcolor red
						Write-Host " Start the game normally (via Bnet client) & this file will be rebuilt." -foregroundcolor red
						Pause
						exit
					}
				}
			}
			try {Copy-item ($SettingsProfilePath + "settings"+ $Script:AccountID + ".json") $SettingsJSON -ErrorAction Stop #overwrite settings.json with settings<ID>.json (<ID> being the account ID). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
				$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).accountlabel
				Write-Host (" Custom game settings (settings" + $Script:AccountID + ".json) being used for " + $CurrentLabel) -foregroundcolor green
				Start-Sleep -milliseconds 100
			}
			catch {
				Write-Host " Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -foregroundcolor red
				Pause
			}
		}
		if ($Script:AskForSettings -eq $True){#steps go through if user has toggled on the manual setting switcher ('s' in the menu).
			$SettingsProfilePath = ("C:\Users\" + $Env:UserName + "\Saved Games\Diablo II Resurrected\")
			$SettingsJSON = ($SettingsProfilePath + "Settings.json")
			$files = Get-ChildItem -Path $SettingsProfilePath -Filter "settings.*.json"
			$Counter = 1
			if ((Test-Path -Path ($SettingsProfilePath+ "Settings" + $id.id +".json")) -ne $true){#if somehow settings<ID>.json doesn't exist yet make one from the current settings.json file.
				try {
					Copy-Item $SettingsJSON ($SettingsProfilePath + "Settings"+ $id.id + ".json") -ErrorAction Stop
				}
				catch {
					Write-Host
					Write-Host " Couldn't find settings.json in $SettingsProfilePath" -foregroundcolor red
					Write-Host " Please start the game normally (via Bnet client) and this file will be rebuilt." -foregroundcolor red
					Pause
					exit
				}
			}
			$SettingsDefaultOptionArray = New-Object -TypeName System.Collections.ArrayList #Add in an option for the default settings file (if it exists, if the auto switcher has never been used it won't appear.
			$SettingsDefaultOption = New-Object -TypeName psobject
			$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
			$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "Name" -Value ("Default - settings"+ $Script:AccountID + ".json")
			$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value ("settings"+ $Script:AccountID + ".json")
			[VOID]$SettingsDefaultOptionArray.Add($SettingsDefaultOption)
			$SettingsFileOptions = New-Object -TypeName System.Collections.ArrayList
			foreach ($file in $files) {
				 $SettingsFileOption = New-Object -TypeName psobject
				 $Counter = $Counter + 1
				 $Name = $file.Name -replace '^settings\.|\.json$'
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "Name" -Value $Name
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
				 [VOID]$SettingsFileOptions.Add($SettingsFileOption)
			}
			if ($SettingsFileOptions -ne $Null){
				$SettingsFileOptions = $SettingsDefaultOptionArray + $SettingsFileOptions
				Write-Host
				Write-Host "  Settings options you can choose from are:"
				foreach ($Option in $SettingsFileOptions){
					Write-Host ("   " + $Option.ID + ". " + $Option.name) -foregroundcolor green
				}
				do {
					Write-Host "  Which Settings file would you like to load from: " -nonewline
					foreach ($Value in $SettingsFileOptions.ID){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
						if ($Value -ne $SettingsFileOptions.ID[-1]){
							Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
							if ($Value -ne $SettingsFileOptions.ID[-2]){Write-Host ", " -nonewline}
						}
						else {
							Write-Host " or $X[38;2;255;165;000;22m$Value$X[0m"
						}
					}
					if ($ManualSettingSwitcher -eq $Null){#if not launched from parameters
						Write-Host "  Or Press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
						$SettingsCancelOption = "c"
					}
					$SettingsChoice = Readkey
					if ($SettingsChoice -eq ""){$SettingsChoice = 1}
					Write-Host
					if($SettingsChoice.tostring() -notin $SettingsFileOptions.id + $SettingsCancelOption){
						Write-Host "  Invalid Input. Please enter one of the options above." -foregroundcolor red
						$SettingsChoice = ""
					}
				} until ($SettingsChoice.tostring() -in $SettingsFileOptions.id + $SettingsCancelOption)
				if ($SettingsChoice -ne "c"){
					$SettingsToLoadFrom = $SettingsFileOptions | where-object {$_.id -eq $SettingsChoice.tostring()}
					try {
						Copy-item ($SettingsProfilePath + $SettingsToLoadFrom.FileName) -Destination $SettingsJSON #-ErrorAction Stop #overwrite settings.json with settings<Name>.json (<Name> being the name of the config user selects). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
						$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).accountlabel
						Write-Host (" Custom game settings (" + $SettingsToLoadFrom.Name + ") being used for " + $CurrentLabel) -foregroundcolor green
						Start-Sleep -milliseconds 100
						Write-Host
					}
					catch {
						Write-Host " Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -foregroundcolor red
						Pause
					}
				}
			}
			Else {
				Write-Host
				Write-Host "  No Custom Settings files have been saved yet. Loading default settings." -foregroundcolor Yellow
				Write-Host "  See README for setup instructions." -foregroundcolor Yellow
				Write-Host
				Pause
			}
		}
		if ($SettingsChoice -ne "c"){
			#Start Game
			#write-output $arguments #debug
			Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments"
			Start-Sleep -milliseconds 1500 #give D2r a bit of a chance to start up before trying to kill handle
			#Close the 'Check for other instances' handle
			Write-Host " Attempting to close `"Check for other instances`" handle..."
			#$handlekilled = $true #debug
			$Output = killhandle | out-string
			if(($Output.contains("DiabloII Check For Other Instances")) -eq $true){
				$handlekilled = $true
				Write-Host " `"Check for Other Instances`" Handle closed." -foregroundcolor green
			}
			else {
				Write-Host " `"Check for Other Instances`" Handle was NOT closed." -foregroundcolor red
				Write-Host " Who even knows what happened. I sure don't." -foregroundcolor red
				Write-Host " You may need to kill this manually via procexp. Good luck hero." -foregroundcolor red
				Write-Host
				Pause
			}

			if ($handlekilled -ne $True){
				Write-Host " Couldn't find any handles to kill." -foregroundcolor red
				Write-Host " Game may not have launched as expected." -foregroundcolor red
				Pause
			}
			#Rename the Diablo Game window for easier identification of which account and region the game is.
			$rename = ($Script:AccountID + " - Diablo II: Resurrected - " + $Script:AccountFriendlyName + " (" + $Script:Region + ")")
			$Command = ('"'+ $WorkingDirectory + '\SetText\SetText.exe" "Diablo II: Resurrected" "' + $rename +'"')
			try {
				cmd.exe /c $Command
				#write-output $Command  #debug
				Write-Host " Window Renamed." -foregroundcolor green
				Start-Sleep -milliseconds 250
				if ($Script:LastAccount -eq $True -or ($Script:OpenAllAccounts -ne $True -and $Script:OpenBatches -ne $True)){
					Write-Host
					Write-Host "Good luck hero..." -foregroundcolor magenta
				}
			}
			catch {
				Write-Host " Couldn't rename window :(" -foregroundcolor red
				Pause
			}
			Start-Sleep -milliseconds 1000
			$Script:ScriptHasBeenRun = $true
		}
	}
}
cls
SetQualityRolls
Menu

#For Diablo II: Resurrected
#Dedicated to my cat Toby.

#Text Colors
#Color Hex	RGB Value	Description
#FFA500		255 165 000	Crafted items
#4169E1		065 105 225	Magic items
#FFFF00		255 255 000	Rare items
#00FF00		000 255 000	Set items
#A59263		165 146 099	Unique items
