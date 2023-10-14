<# 
Author: Shupershuff
Usage: 
Happy for you to make any modifications this script for your own needs providing:
- Any variants of this script are never sold.
- Any variants of this script published online should always be open source.
- Any variants of this script are never modifed to enable or assist in any game altering or malicious behaviour including (but not limited to): Bannable Mods, Cheats, Exploits, Phishing
Purpose:
	Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle."
	Script will import account details from CSV. Alternatively you can run script parameters (see Github readme): -AccountUsername, -PW, -Region, -All, -Batch, -ManualSettingSwitcher
Instructions: See GitHub readme https://github.com/shupershuff/Diablo2RLoader

Notes:
- Multiple failed attempts (eg wrong Password) to sign onto a particular Realm via this method may temporarily lock you out. You should still be able to get in via the battlenet client if this occurs.

Servers:
 NA - us.actual.battle.net
 EU - eu.actual.battle.net
 Asia - kr.actual.battle.net
 
Changes since 1.9.0 (next version edits):
Added Token based authentication. There are some setup steps and drawbacks to using this method but this can help when Blizzard Auth servers don't respond to authentication when launching via parameters.
Added notification capability so if I notice one of the auth servers going down I can advise users directly.
Added a new version of the SetText source code to enable renaming D2r windows based on processID. This prevents issues for people who load an account from Battle.net in addition to using the script.
Changed Current and Next TZ source to D2Emu.com for faster and more accurate results (no more OCR delays or inaccuracies).
Script now works if one game instance has been launched outside of this script via battlenet. This is for users who want to keep MFA enabled on one account and launch that account outside of the script.
Improved Game Window name by adjusting to show more useful information first (ID - Acct Label (region) - Diablo II: Resurrected).
Improved DClone checker by adding 1 second to DClone timeout as 2 seconds was a bit tight and caused errors.
Fixed unhandled DClone errors for timeouts.
Minor tidy ups.

##########TEST IMMEDIATELY AFTER RELEASE, test if settextv2.bas is downloaded and extracted.
#do before release, delete all .zip files in other releases.

1.10.1+ to do list
To reduce lines, Tidy up all the import/export csv bits for stat updates into a function rather than copy paste the same commands throughout the script.
To reduce lines, add repeated commands into functions
Unlikely - Possibly add Current and Next TZ status for Single player folk, but ONLY if it's an easy addition with an easy source.
Unlikely - ISboxer has CTRL + Alt + number as a shortcut to switch between windows. Investigate how this could be done. Would need an agent to detect key combos, Possibly via AutoIT or Autohotkey. Likely not possible within powershell and requires a separate project.
Fix whatever I broke or poorly implemented in 1.10.0 :)
#>

param($AccountUsername,$PW,$Region,$All,$Batch,$ManualSettingSwitcher) #used to capture parameters sent to the script, if anyone even wants to do that.
$CurrentVersion = "1.9.1"

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
#run script as admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $ScriptArguments"  -Verb RunAs;exit } 

#set window size
[console]::WindowWidth=77; #script has been designed around this width. Adjust at your own peril.
[console]::WindowHeight=48;
[console]::BufferWidth=[console]::WindowWidth
#set misc vars
$Script:X = [char]0x1b #escape character for ANSI text colors
$ProgressPreference = "SilentlyContinue"
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\')) #Set Current Directory path.
$Script:StartTime = Get-Date #Used for elapsed time. Is reset when script refreshes.
$Script:MOO = "%%%"
$MenuRefreshRate = 30 #How often the script refreshes in seconds.
$Script:ScriptFileName = Split-Path $MyInvocation.MyCommand.Path -Leaf #find the filename of the script in case a user renames it.
$Script:SessionTimer = 0 #set initial session timer to avoid errors in info menu.
$Script:NotificationHasBeenChecked = $False
#Baseline of acceptable characters for ReadKey functions. Used to prevents receiving inputs from folk who are alt tabbing etc.
$Script:AllowedKeyList = @(48,49,50,51,52,53,54,55,56,57) #0 to 9
$Script:AllowedKeyList += @(48,49,50,51,52,53,54,55,56,57) #0 to 9 on numpad
$Script:AllowedKeyList += @(65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90) # A to Z
$EnterKey = 13

Function ReadKey([string]$message=$Null,[bool]$NoOutput,[bool]$AllowAllKeys) {#used to receive user input
    $key = $Null
    $Host.UI.RawUI.FlushInputBuffer()
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host -NoNewLine $message
    }
	$AllowedKeyList = $Script:AllowedKeyList + @(13,27) #Add Enter & Escape to the allowedkeylist as acceptable inputs.
    while ($Null -eq $key) {
        if ($Host.UI.RawUI.KeyAvailable) {
            $key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
            if ($True -ne $AllowAllKeys){
				if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList) {
					$key = $key_
				}
			}
			else {
				if ($key_.KeyDown) {
					$key = $key_
				}
			}
        }
		else {
            Start-Sleep -m 200  # Milliseconds
        }
    }
	if ($key_.VirtualKeyCode -ne $EnterKey -and -not ($Null -eq $key) -and [bool]$NoOutput -ne $true) {
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

Function ReadKeyTimeout([string]$message=$Null, [int]$timeOutSeconds=0, [string]$Default=$Null) {#used to receive user input but times out after X amount of time
	$key = $Null
    $Host.UI.RawUI.FlushInputBuffer()
    if (![string]::IsNullOrEmpty($message)) {
        Write-Host -NoNewLine $message
    }
    $Counter = $timeOutSeconds * 1000 / 250
	$IgnoreKeyList = @(9,13,16,17,18,20,32,91,192) #Ignore Tab,Enter,Shift,Ctrl,Alt,Caps,Space,Windows Key,Tilde
    while ($Null -eq $key -and ($timeOutSeconds -eq 0 -or $Counter-- -gt 0)) {
        if (($timeOutSeconds -eq 0) -or $Host.UI.RawUI.KeyAvailable) {
            $key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
            if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList) {
                $key = $key_
            }
        }
		else {
            Start-Sleep -m 200  # Milliseconds
        }
    }	
    if ($key_.VirtualKeyCode -ne $EnterKey -and -not ($Null -eq $key)) {
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
Function PressTheAnyKey {#Used instead of Pause so folk can hit any key to continue
	write-host "  Press any key to continue..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	write-host
}
Function PressTheAnyKeyToExit {#Used instead of Pause so folk can hit any key to exit
	write-host "  Press Any key to exit..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	remove-job * -force
	Exit
}

if ((Test-Path -Path "$Script:WorkingDirectory\Stats.csv") -ne $true){#Create Stats CSV if it doesn't exist
	$CreateStatCSV = {} | Select "TotalGameTime","TimesLaunched","LastUpdateCheck","HighRunesFound","UniquesFound","SetItemsFound","RaresFound","MagicItemsFound","NormalItemsFound","Gems","CowKingKilled","PerfectGems" | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
	write-host " Stats.csv created!"
}
do {
	$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv" #Get current stats csv details
	if ($CurrentStats -ne $null){
		#Todo: In the Future add CSV validation checks
		$StatsCSVImportSuccess = $True
	}
	else {#Error out and exit if there's a problem with the csv.
			if ($StatsCSVRecoveryAttempt -lt 1){
				try {
					Write-Host " Issue with Stats.csv. Attempting Autorecovery from backup..." -foregroundcolor red
					Copy-Item -Path $Script:WorkingDirectory\Stats.backup.csv -Destination $Script:WorkingDirectory\Stats.csv
					Write-Host " Autorecovery successful!" -foregroundcolor Green
					$StatsCSVRecoveryAttempt ++
					PressTheAnyKey
				}
				Catch {
					$StatsCSVImportSuccess = $False
				}
			}
			Else {
				$StatsCSVRecoveryAttempt = 2
			}
			if ($StatsCSVImportSuccess -eq $False -or $StatsCSVRecoveryAttempt -eq 2){
				write-host
				write-host " Stats.csv is corrupted or empty." -foregroundcolor red
				write-host " Replace with data from stats.backup.csv or delete stats.csv" -foregroundcolor red
				write-host
				PressTheAnyKeyToExit
			}
		}
} until ($StatsCSVImportSuccess -eq $True)

if (-not ($CurrentStats | Get-Member -Name "LastUpdateCheck" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.8.1+. If LastUpdateCheck column doesn't exist, add it to the CSV data
	$CurrentStats | ForEach-Object {
		$_ | Add-Member -NotePropertyName "LastUpdateCheck" -NotePropertyValue "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
	}
}
elseif ($CurrentStats.LastUpdateCheck -eq "" -or $CurrentStats.LastUpdateCheck -like "*/*") {# If script has just been freshly downloaded or has the old Date format.
	$CurrentStats.LastUpdateCheck = "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
	$CurrentStats | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
}

#Only Check for updates if updates haven't been checked in last 8 hours. Reduces API requests.
if ($CurrentStats.LastUpdateCheck -lt (Get-Date).addHours(-8).ToString('yyyy.MM.dd HH:mm:ss')){# Compare current date and time to LastUpdateCheck date & time.
	try {
		$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
		$ReleaseInfo = ($Releases | sort id -desc)[0] #find release with the highest ID.
		$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		if ($Script:LatestVersion -gt $Script:CurrentVersion) {
			Write-Host
			Write-Host " Update available! See Github for latest version and info" -foregroundcolor Yellow -nonewline
			if ([version[]]$CurrentVersion -in ($Releases.Name.Trim('v') | sort -desc)[2..$releases.count]){
				Write-Host ".`n There have been several releases since your version." -foregroundcolor Yellow
				Write-Host " Checkout Github releases for fixes/features added. " -foregroundcolor Yellow
				Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/$X[0m"
			} Else {
				Write-Host ":`n $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m"
			}
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
			Write-Host; Write-Host
			Do {
				Write-Host " Your Current Version is v$CurrentVersion."
				Write-Host (" Would you like to update to v"+ $Script:LatestVersion + "? $X[38;2;255;165;000;22mY$X[0m/$X[38;2;255;165;000;22mN$X[0m: ") -nonewline
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
				Write-Host
				Write-Host " Updating... :)" -foregroundcolor green
				try {
					New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
				}
				Catch {#if folder already exists for whatever reason.
					Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
					New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
				}
				$ZipURL = $ReleaseInfo.zipball_url #get zip download URL	
				$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
				Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
				if ($releaseinfo.assets.browser_download_url -ne $Null){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.				
					Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
				}
				$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
				Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
				$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
				Copy-Item -Path ($FolderPath + "\D2Loader.ps1") -Destination ($Script:WorkingDirectory + "\" + $Script:ScriptFileName) #using $Script:ScriptFileName allows the user to rename the file if they want
				Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
				Write-Host " Updated :)" -foregroundcolor green
				Start-Sleep -milliseconds 850
				& ($Script:WorkingDirectory + "\" + $Script:ScriptFileName)	
				exit
			}
		}
		$CurrentStats.LastUpdateCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
		$Script:LatestVersionCheck = $CurrentStats.LastUpdateCheck
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update stats.csv with the new time played.		
	}
	Catch {
		Write-Host
		Write-Host " Couldn't check for updates. GitHub API limit may have been reached..." -foregroundcolor Yellow
		Start-Sleep -milliseconds 3500
	}
}
#Update (or replace missing) SetTextV2.bas file. This is an newer version of SetText (built for my by ChatGPT) that allows windows to be closed by process ID.
if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.bas')) -ne $True){#if SetTextv2.bas doesn't exist, download it.
		try {
			New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
		}
		Catch {#if folder already exists for whatever reason.
			Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
			New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
		}
		$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
		$ReleaseInfo = ($Releases | sort id -desc)[0] #find release with the highest ID.
		$ZipURL = $ReleaseInfo.zipball_url #get zip download URL	
		$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
		Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
		if ($releaseinfo.assets.browser_download_url -ne $Null){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.				
			Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
		}
		$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
		Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
		$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
		Copy-Item -Path ($FolderPath + "\SetText\SetTextv2.bas") -Destination ($Script:WorkingDirectory + "\SetText\SetTextv2.bas")
		write-host "  SetTextV2.bas was missing and was downloaded."
		Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
}

#Import Config XML
Function ImportXML {
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
		PressTheAnyKeyToExit
	}
}
ImportXML
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
if ($Script:Config.CommandLineArguments -ne $Null){#remove this config option as arguments are now stored in accounts.csv so that different arguments can be set for each account
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
	$Script:OriginalCommandLineArguments = $Script:Config.CommandLineArguments
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
	PressTheAnyKey
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
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	PressTheAnyKey
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
	ImportXML
	PressTheAnyKey
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
	PressTheAnyKey
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
	PressTheAnyKey
}
if ($Script:Config.DCloneTrackerSource -eq $Null){
	Write-Host
	Write-Host " Config option 'DCloneTrackerSource' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is a required config option to determine which source should be used" -foregroundcolor Yellow
	Write-Host " for obtaining current DClone status." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</TrackAccountUseTime>"
	$Replacement = "</TrackAccountUseTime>`n`n`t<!--Options are d2rapi.fly.dev, D2runewizard.com and diablo2.io.`n`t"
	$Replacement += "Default and recommended option is d2rapi.fly.dev as this pulls live data from the game as opposed to crowdsourced data.-->`n`t"
	$Replacement += "<DCloneTrackerSource>d2rapi.fly.dev</DCloneTrackerSource>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	ImportXML
	PressTheAnyKey
}
if ($Script:Config.DCloneAlarmLevel -eq $Null){
	Write-Host
	Write-Host " Config option 'DCloneAlarmLevel' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This field determines if alarms should activate for all DClone status " -foregroundcolor Yellow
	Write-Host " changes or just when DClone is about to walk." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DCloneTrackerSource>"
	$Replacement = "</DCloneTrackerSource>`n`n`t<!--Specify what Statuses you want to be alarmed on.`n`t"
	$Replacement +=	"Enter `"All`" to be alarmed of all status changes`n`t"
	$Replacement +=	"Enter `"Close`" to be only alarmed when status is 4/6, 5/6 or has just walked.`n`t"
	$Replacement +=	"Enter `"Imminent`" to be only alarmed when status is 5/6 or has just walked.`n`t"
	$Replacement +=	"Recommend setting to `"All`"-->`n`t"
	$Replacement +=	"<DCloneAlarmLevel>All</DCloneAlarmLevel>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	ImportXML
	PressTheAnyKey
}
if ($Script:Config.DCloneAlarmList -eq $Null){
	Write-Host
	Write-Host " Config option 'DCloneAlarmList' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This is an optional config option to enable both audible and text based" -foregroundcolor Yellow
	Write-Host " alarms for DClone Status changes." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DCloneTrackerSource>"
	$Replacement = "</DCloneTrackerSource>`n`n`t<!--Allow you to have the script audibly warn you of upcoming dclone walks.`n`t"
	$Replacement +=	"Specify as many of the following options as you like: SCL-NA, SCL-EU, SCL-KR, SC-NA, SC-EU, SC-KR, HCL-NA, HCL-EU, HCL-KR, HC-NA, HC-EU, HC-KR`n`t"
	$Replacement +=	"EG if you want to be notified for all Softcore ladder walks on all regions, enter <DCloneAlarmList>SCL-NA, SCL-EU, SCL-KR</DCloneAlarmList>`n`t"
	$Replacement +=	"If left blank this feature is disabled. Default is blank as this may be annoying for some people-->`n`t"
	$Replacement +=	"<DCloneAlarmList></DCloneAlarmList>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	ImportXML
	PressTheAnyKey
}
if ($Null -ne $Script:Config.DCloneAlarmList -and $Script:Config.DCloneAlarmList -ne ""){#validate data to prevent errors from typos
	$pattern = "^(HC|SC)(L?)-(NA|EU|KR)$" #set pattern: must start with HC or SC, optionally has L after it, must end in -NA -EU or -KR
	foreach ($Alarm in $Script:Config.DCloneAlarmList.split(",").trim()) {
		if ($Alarm -notmatch $pattern) {
			Write-Host
			Write-Host " $Alarm is not a valid Alarm entry."  -foregroundcolor Red
			Write-Host " See valid options in Config.xml"  -foregroundcolor Red
			Write-Host
			PressTheAnyKeyToExit
		}
	}
}
if ($Script:Config.DCloneAlarmVoice -eq $Null){
	Write-Host
	Write-Host " Config option 'DCloneAlarmVoice' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This config allows you to choose between a Woman or Man's robot voice." -foregroundcolor Yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</DCloneAlarmLevel>"
	$Replacement = "</DCloneAlarmLevel>`n`n`t<!--Specify what voice you want. Choose 'Paladin' for David (Man) or 'Amazon' for Zira (Woman).-->`n`t"
	$Replacement +=	"<DCloneAlarmVoice>Paladin</DCloneAlarmVoice>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	PressTheAnyKey
}
if ($Script:Config.AuthenticationMethod -eq $Null){
	Write-Host
	Write-Host " Config option 'AuthenticationMethod' missing from config.xml" -foregroundcolor Yellow
	Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
	Write-Host " This config allows you to use alternative authentication methods." -foregroundcolor Yellow
	Write-Host " You can change this if Blizzard Auth servers are having issues." -foregroundcolor yellow
	Write-Host " Added this missing option into .xml file :)" -foregroundcolor green
	Write-Host
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	$Pattern = "</GamePath>"
	$Replacement = "</GamePath>`n`n`t<!--Choose the authentication method you want the script to use.`n`t"
	$Replacement +=	"Options are Token and Parameter. Parameter uses standard UserName Password parameters.`n`t"
	$Replacement +=	"With the token option, it's not possible to launch all accounts at once or use the batching feature`n`t"
	$Replacement +=	"as you must make sure you get to the character select screen before launching another instance.`n`t"
	$Replacement +=	"Token option requires additional setup steps but is useful when Blizzard auth servers are having issues.`n`t"
	$Replacement +=	"See Readme for more details. Default and recommended option is Parameter-->`n`t"
	$Replacement +=	"<AuthenticationMethod>Parameter</AuthenticationMethod>" #add option to config file if it doesn't exist.
	$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
	$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
	Start-Sleep -milliseconds 1500
	ImportXML
	PressTheAnyKey
}
if ($Script:Config.AuthenticationMethod -ne "Token" -and $Script:Config.AuthenticationMethod -ne "Parameter"){
		Write-host
		Write-host " Error: AuthenticationMethod in config has an invalid option set." -Foregroundcolor red
		Write-host " Open config.xml and set this to either Parameter or Token." -Foregroundcolor red
		Write-host
		PressTheAnyKeyToExit
}

if ($Script:Config.DCloneAlarmList -ne ""){
	$ValidVoiceOptions = 
		"Amazon",
		"Paladin",
		"Woman",
		"Man",
		"Wench",
		"Bloke"
	if ($Script:Config.DCloneAlarmVoice -notin $ValidVoiceOptions){
		Write-host
		Write-host " Error: DCloneAlarmVoice in config has an invalid option set." -Foregroundcolor red
		Write-host " Open config.xml and set this to either Paladin or Amazon." -Foregroundcolor red
		Write-host
		PressTheAnyKeyToExit
	}
	if ($Script:Config.DCloneAlarmLevel -eq "All"){
		$Script:DCloneAlarmLevel = "1,2,3,4,5,6"
	}
	elseif ($Script:Config.DCloneAlarmLevel -eq "Close"){
		$Script:DCloneAlarmLevel = "1,4,5,6"
	}
	elseif ($Script:Config.DCloneAlarmLevel -eq "Imminent"){
		$Script:DCloneAlarmLevel = "1,5,6"
	}
	else {#if user has typo'd the config file or left it blank.
		$DCloneErrorMessage = ("  Error: DClone Alarm Levels have been misconfigured in config.xml. ###  Check that the value for DCloneAlarmLevel is entered correctly.").Replace("###", "`n")
		Write-host
		Write-host $DCloneErrorMessage -Foregroundcolor red
		Write-host
		PressTheAnyKeyToExit
	}
}

$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig #import config.xml again for any updates made by the above.
$D2CloneTrackerSource = $Script:Config.DCloneTrackerSource # Set tracker source variable
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
"SettingSwitcherEnabled",
"TrackAccountUseTime"

$AvailableConfigs = $AvailableConfigs + $BooleanConfigs

if ($Script:Config.CheckForNextTZ -ne $true -and $Script:Config.CheckForNextTZ -ne $false){#if CheckForNextTZ config is invalid, set to false
	$Script:CheckForNextTZ = $false
} Else {
	$Script:CheckForNextTZ = $Script:Config.CheckForNextTZ
}

$ConfigXMLlist = ($Config | Get-Member | Where-Object {$_.membertype -eq "Property" -and $_.name -notlike "#comment"}).name
Write-Host
foreach ($Option in $AvailableConfigs){#Config validation
	if ($Option -notin $ConfigXMLlist){
		Write-Host "Config.xml file is missing a config option for $Option." -foregroundcolor yellow
		Start-Sleep 1
		PressTheAnyKey
	}
}
if ($Option -notin $ConfigXMLlist){
	Write-Host
	Write-Host "Make sure to grab the latest version of config.xml from GitHub" -foregroundcolor yellow
	Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m"
	Write-Host
	PressTheAnyKey
}
if ($Config.GamePath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$Script:GamePath = $Config.GamePath.replace("`"","")
}
else {
	$Script:GamePath = $Config.GamePath
}
foreach ($ConfigCheck in $BooleanConfigs){#validate all configs that require "True" or "False" as the setting.
	if ($Config.$ConfigCheck -ne $Null -and ($Config.$ConfigCheck -ne $true -and $Config.$ConfigCheck -ne $false)){#if config is invalid
		Write-Host " Config option '$ConfigCheck' is invalid." -foregroundcolor yellow 
		Write-Host " Ensure this is set to either True or False." -foregroundcolor yellow
		Write-Host;	PressTheAnyKey
	}
}
if ($Config.ShortcutCustomIconPath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath.replace("`"","")
}
else {
	$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath
}
$AskForRegionOnceOnly = $Config.AskForRegionOnceOnly
$CreateDesktopShortcut = $Config.CreateDesktopShortcut
$Script:ConvertPlainTextPasswords = $Config.ConvertPlainTextPasswords

#Check Windows Game Path for D2r.exe is accurate.
if ((Test-Path -Path "$GamePath\d2r.exe") -ne $True){ 
	Write-Host " Gamepath is incorrect. Looks like you have a custom D2r install location!" -foregroundcolor red
	Write-Host " Edit the GamePath variable in the config file." -foregroundcolor red
	write-host
	PressTheAnyKeyToExit
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

#Check if SetTextv2.exe exists, if not, compile from SetTextv2.bas. SetTextv2.exe is what's used to rename the windows.
if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	Write-Host
	Write-Host " First Time run!" -foregroundcolor Yellow
	Write-Host
	Write-Host " SetTextv2.exe not in .\SetText\ folder and needs to be built."
	if ((Test-Path -Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe") -ne $True){#check that .net4.0 is actually installed or compile will fail.
		Write-Host " .Net v4.0 not installed. This is required to compile the Window Renamer for Diablo." -foregroundcolor red
		Write-Host " Download and install it from Microsoft here:" -foregroundcolor red
		Write-Host " https://dotnet.microsoft.com/en-us/download/dotnet-framework/net40" #actual download link https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net40-web-installer
		PressTheAnyKeyToExit
	}	
	Write-Host " Compiling SetTextv2.exe from SetTextv2.bas..."
	& "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" -target:winexe -out:"`"$WorkingDirectory\SetText\SetTextv2.exe`"" "`"$WorkingDirectory\SetText\SetTextv2.bas`"" | out-null #/verbose  #actually compile the bastard
	if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.exe')) -ne $True){#if it fails for some reason and settextv2.exe still doesn't exist.
		Write-Host " SetTextv2 Could not be built for some reason :/"
		PressTheAnyKeyToExit
	}
	Write-Host " Successfully built SetTextv2.exe for Diablo 2 Launcher script :)" -foregroundcolor green
	Start-Sleep -milliseconds 4000 #a small delay so the first time run outputs can briefly be seen
}

#Check Handle64.exe downloaded and placed into correct folder
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))
if ((Test-Path -Path ($workingdirectory + '\Handle\Handle64.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	Write-Host " Handle64.exe is in the .\Handle\ folder. See instructions for more details on setting this up." -foregroundcolor red
	PressTheAnyKeyToExit
}

Function ValidateTokenInput {
	param (
		[bool] $ManuallyEntered,
		[string] $TokenInput
	)
	do {
		$extractedInfo = $null
		if ($ManuallyEntered) {
			$TokenInput = Read-host " Enter your token URL or enter the token"
		}
		$pattern = "(?<=\?ST=|&ST=|^|http://localhost:0/\?ST=)([^&]+)"
		if ($tokeninput -match $pattern) {
			$extractedInfo = ($matches[1]).replace("http://localhost:0/?ST=","")
			return $extractedInfo
		}
		Else {
			Write-Host " Token details are incorrect." -foregroundcolor red
			if (!$ManuallyEntered) {
				Write-Host " Please review the setup instructions." -foregroundcolor red
				Write-Host
				PressTheAnyKeyToExit
			}
			Else {
				Write-Host " Please enter the full URL." -foregroundcolor red
			}
		}
	}
	until ($extractedInfo -ne $null)
}

#Import Account CSV
Function ImportCSV {
	do {
		if ($Script:AccountUsername -eq $Null){#If no parameters sent to script.
			try {
				$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv" #import all accounts from csv
			}
			Catch {
				Write-Host
				Write-Host " Accounts.csv does not exist. Make sure you create this and populate with accounts first." -foregroundcolor red
				PressTheAnyKeyToExit
			}
		}
		if ($Script:AccountOptionsCSV -ne $Null){
			#check Accounts.csv has been updated and doesn't contain the example account.
			if ($Script:AccountOptionsCSV -match "yourbnetemailaddress"){
				Write-Host
				Write-Host "You haven't setup accounts.csv with your accounts." -foregroundcolor red
				Write-Host "Add your account details to the CSV file and run the script again :)" -foregroundcolor red
				Write-Host
				PressTheAnyKeyToExit
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Batches" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.7.0. If batch column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "Batches" -NotePropertyValue $Null
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Token" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "TokenIsSecureString" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.7.0. If batch column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Token" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
					$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "Token" -NotePropertyValue $Null}
				}
				if (-not ($Script:AccountOptionsCSV | Get-Member -Name "TokenIsSecureString" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
					$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "TokenIsSecureString" -NotePropertyValue $Null}
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
			}
			
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "CustomLaunchArguments" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.8.0. If CustomLaunchArguments column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "CustomLaunchArguments" -NotePropertyValue $Script:OriginalCommandLineArguments
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
				Write-Host " Added CustomLaunchArguments column to accounts.csv." -foregroundcolor green
				Write-Host
				Start-Sleep -milliseconds 1200
				PressTheAnyKey
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "TimeActive" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {#For update 1.8.0. If TimeActive column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "TimeActive" -NotePropertyValue $Null
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
				Write-host " Added TimeActive column to accounts.csv." -foregroundcolor Green
				PressTheAnyKey
			}
			#Secure any plain text tokens. Ask for tokens on accounts that don't have any if Config is configured to use Token Authentication.
			$NewCSV = Foreach ($Entry in $Script:AccountOptionsCSV) {
				if ($Entry.TokenIsSecureString.length -gt 0 -and $Entry.TokenIsSecureString -ne $False){#if nothing needs converting, make sure existing entries still make it into the updated CSV
					$Entry
				}
				if (($Entry.TokenIsSecureString.length -eq 0 -or $Entry.TokenIsSecureString -eq "no" -or $Entry.TokenIsSecureString -eq $false) -and $Entry.Token.length -ne 0){#if account.csv has a Token and TokenIsSecureString isn't set to yes, convert Token to secure string and update CSV.
					ValidateTokenInput -TokenInput $Entry.Token
					$Entry.Token = ConvertTo-SecureString -String $Entry.Token -AsPlainText -Force
					$Entry.Token = $Entry.Token | ConvertFrom-SecureString
					$Entry.TokenIsSecureString = "Yes"
					Write-Host (" Secured Token for " + $Entry.AccountLabel) -foregroundcolor green
					Start-Sleep -milliseconds 100
					$Entry
					$TokensUpdated = $true
				}
				if ($Entry.Token.length -eq 0 -and $Script:Config.AuthenticationMethod -eq "Token"){#if csv has account details but Token field has been left blank
					Write-Host
					Write-Host (" The account " + $Entry.AccountLabel + " doesn't yet have a Token defined.") -foregroundcolor yellow
					Write-Host " See the readme for how to obtain Auth token." -foregroundcolor yellow
					Write-Host
					while ($Entry.Token.length -eq 0){#prevent empty entries as this will cause errors.
						$Entry.Token = ValidateTokenInput -ManuallyEntered $True
					}
					$Entry.Token = ConvertTo-SecureString -String $Entry.Token -AsPlainText -Force
					$Entry.Token = $Entry.Token | ConvertFrom-SecureString
					$Entry.TokenIsSecureString = "Yes"
					Write-Host (" Secured Token for " + $Entry.AccountLabel) -foregroundcolor green
					Start-Sleep -milliseconds 100
					$Entry
					$TokensUpdated = $true
				}
			}
			if ($Script:ConvertPlainTextPasswords -ne $False){
				#Check CSV for Plain text Passwords, convert to encryptedstrings and replace values in CSV
				$NewCSV = Foreach ($Entry in $Script:AccountOptionsCSV) {
					if ($Entry.PWisSecureString.length -gt 0 -and $Entry.PWisSecureString -ne $False){#if nothing needs converting, make sure existing entries still make it into the updated CSV
						$Entry
					}
					if (($Entry.PWisSecureString.length -eq 0 -or $Entry.PWisSecureString -eq "no" -or $Entry.PWisSecureString -eq $false) -and $Entry.PW.length -ne 0){#if account.csv has a password and PWisSecureString isn't set to yes, convert PW to secure string and update CSV.
						$Entry.PW = ConvertTo-SecureString -String $Entry.PW -AsPlainText -Force
						$Entry.PW = $Entry.PW | ConvertFrom-SecureString
						$Entry.PWisSecureString = "Yes"
						Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
						Start-Sleep -milliseconds 100
						$Entry
						$PWsUpdated = $true
					}
					if ($Entry.PW.length -eq 0){#if csv has account details but password field has been left blank
						Write-Host
						Write-Host (" The account " + $Entry.AccountLabel + " doesn't yet have a password defined.") -foregroundcolor yellow
						Write-Host
						while ($Entry.PW.length -eq 0){#prevent empty entries as this will cause errors.
							$Entry.PW = read-host -AsSecureString " Enter the Battle.net password for"$Entry.AccountLabel
						}
						$Entry.PW = $Entry.PW | ConvertFrom-SecureString
						$Entry.PWisSecureString = "Yes"
						Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
						Start-Sleep -milliseconds 100
						$Entry
						$PWsUpdated = $true
					}
				}
			}
			if ($PWsUpdated -eq $true -or $TokensUpdated -eq $True){#if CSV needs to be updated
				Try {
					$NewCSV | Export-CSV "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update CSV file
					if ($TokensUpdated -eq $True){
						Write-Host " Accounts.csv updated: Tokens have been secured." -foregroundcolor green
					}
					if ($PWsUpdated -eq $True){
						Write-Host " Accounts.csv updated: Passwords have been secured." -foregroundcolor green
					}
					Start-Sleep -milliseconds 4000
				}
				Catch {
					Write-Host
					Write-Host " Couldn't update Accounts.csv, probably because the file is open & locked." -foregroundcolor red
					Write-Host " Please close accounts.csv and run the script again!" -foregroundcolor red
					PressTheAnyKeyToExit
				}
			}
			$AccountCSVImportSuccess = $True
		}
		else {#Error out and exit if there's a problem with the csv.
			if ($AccountCSVRecoveryAttempt -lt 1){
				try {
					Write-Host " Issue with accounts.csv. Attempting Autorecovery from backup..." -foregroundcolor red
					Copy-Item -Path $Script:WorkingDirectory\Accounts.backup.csv -Destination $Script:WorkingDirectory\Accounts.csv
					Write-Host " Autorecovery successful!" -foregroundcolor Green
					$AccountCSVRecoveryAttempt ++
					PressTheAnyKey
				}
				Catch {
					$AccountCSVImportSuccess = $False
				}
			}
			Else {
				$AccountCSVRecoveryAttempt = 2
			}
			if ($AccountCSVImportSuccess -eq $False -or $AccountCSVRecoveryAttempt -eq 2){
				write-host
				write-host " There's an issue with accounts.csv." -foregroundcolor red
				write-host " Please ensure that this is filled out correctly and rerun the script." -foregroundcolor red
				write-host " Alternatively, rebuild CSV from scratch or restore from accounts.backup.csv" -foregroundcolor red
				write-host
				PressTheAnyKeyToExit
			}
		}
	} until ($AccountCSVImportSuccess -eq $True)
	$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	([int]$CurrentStats.TimesLaunched) ++
	if ($CurrentStats.TotalGameTime -eq ""){$CurrentStats.TotalGameTime = 0} #prevents errors from happening on first time run.
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
	}
	Catch {
		Write-host "  Couldn't update stats.csv" -foregroundcolor yellow
	}
	#Make Backup of CSV. 
	 # Added this in as I had BSOD on my PC and noticed that this caused the files to get corrupted.
	Copy-Item -Path ($Script:WorkingDirectory + "\stats.csv") -Destination ($Script:WorkingDirectory + "\stats.backup.csv")
	Copy-Item -Path ($Script:WorkingDirectory + "\accounts.csv") -Destination ($Script:WorkingDirectory + "\accounts.backup.csv")
}

#Set Region Array
$Script:ServerOptions = @(
	[pscustomobject]@{Option='1';region='NA';region_server='us.actual.battle.net'}#Americas
	[pscustomobject]@{Option='2';region='EU';region_server='eu.actual.battle.net'}#Europe
	[pscustomobject]@{Option='3';region='Asia';region_server='kr.actual.battle.net'}
)

Function SetQualityRolls {
	#Set item quality array for randomizing quote colours. A stupid addition to script but meh.
	$Script:QualityArray = @(#quality and chances for things to drop based on 0MF values in D2r (I think?)
		[pscustomobject]@{Type='HighRune';Probability=1}
		[pscustomobject]@{Type='Unique';Probability=50}
		[pscustomobject]@{Type='SetItem';Probability=124}
		[pscustomobject]@{Type='Rare';Probability=200}
		[pscustomobject]@{Type='Magic';Probability=588}
		[pscustomobject]@{Type='Normal';Probability=19036}
	)
	if ($Script:GemActivated -eq $True){#small but noticeable MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 16384  # New probability value
		}
	}
	Else {
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 19036  # Original probability value
		}
	}
	if ($Script:CowKingActivated -eq $True){#big MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 2048  # New probability value
			$Script:MOO = "MOO"
		}
	}
	if ($Script:PGemActivated -eq $True){#huuge MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 192  # New probability value
		}
	}
	$QualityHash = @{}; 
	foreach ($Object in $Script:QualityArray | select-object type,probability){#convert PSOobjects to hashtable for enumerator
		$QualityHash.add($Object.type,$Object.probability) #add each PSObject to hash
	}
	$Script:ItemLookup = foreach ($Entry in $QualityHash.GetEnumerator()){
		[System.Linq.Enumerable]::Repeat($Entry.Key, $Entry.Value) #This creates a hash table with 19036 normal items, 588 magic items, 200 rare items etc etc. Used later as a list to randomly pick from.
	}
}

function HighRune {
	process { Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4m$_$X[0m"}
}
function Unique {
    process { Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4m$_$X[0m"}
}
function SetItem {
    process { Write-Host "  $X[38;2;0;225;0;48;2;1;1;1;4m$_$X[0m"}
}
function Rare {
    process { Write-Host "  $X[38;2;255;255;0;48;2;1;1;1;4m$_$X[0m"}
}
function Magic {#ANSI text colour formatting for "magic" quotes. The variable $X (for the escape character) is defined earlier in the script.
    process { Write-Host "  $X[38;2;65;105;225;48;2;1;1;1;4m$_$X[0m" }
}
function Normal {
    process { Write-Host "  $X[38;2;255;255;255;48;2;1;1;1;4m$_$X[0m"}
}

function QuoteRoll {#stupid thing to draw a random quote but also draw a random quality.
	$Quality = get-random $Script:ItemLookup #pick a random entry from ItemLookup hashtable.
	Write-Host
	$LeQuote = (Get-Random -inputobject $Script:quotelist) #pick a random quote.
	$ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
	$DesiredIndent = 2  # indent spaces
	$ChunkSize = $ConsoleWidth - $DesiredIndent
	[RegEx]::Matches($LeQuote, ".{$ChunkSize}|.+").Groups.Value | ForEach-Object {
		write-output $_ | &$Quality #write the quote and write it in the quality colour
	}
	$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	if ($Quality -eq "HighRune"){([int]$CurrentStats.HighRunesFound) ++}
	if ($Quality -eq "Unique"){([int]$CurrentStats.UniquesFound) ++}
	if ($Quality -eq "SetItem"){([int]$CurrentStats.SetItemsFound) ++}
	if ($Quality -eq "Rare"){([int]$CurrentStats.RaresFound) ++}
	if ($Quality -eq "Magic"){([int]$CurrentStats.MagicItemsFound) ++}
	if ($Quality -eq "Normal"){([int]$CurrentStats.NormalItemsFound) ++}
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv
	}
	Catch {
		Write-host "  Couldn't update stats.csv" -foregroundcolor yellow
		Start-Sleep -Milliseconds 256
	}
}

Function Inventory {#Info screen
	#cls
	Write-Host
	Write-Host "          Stay a while and listen! Here's your D2r Loader info." -foregroundcolor yellow
	write-host;	write-host
	write-host "  $X[38;2;255;255;255;4mNote:$X[0m D2r Playtime is based on the time the script has been running"
	write-host "  whilst D2r is running. In other words, if you use this script when you're"
	write-host "  playing the game, it will give you a reasonable idea of the total time"
	write-host "  you've spent receiving disappointing drops from Mephisto :)"
	$QualityArraySum = 0
	$Script:QualityArray | ForEach-Object {
		$QualityArraySum += $_.Probability
	}
	$NormalProbability = ($QualityArray | where-object {$_.type -eq "Normal"} | Select Probability).probability
	write-host
	$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"	
	$Line1 = "                    --------------------------------"
	$Line2 = ("                   |  $X[38;2;255;255;255;22mD2r Playtime (Hours):$X[0m    " +  ((($time =([TimeSpan]::Parse($CurrentStats.TotalGameTime))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line3 = ("                   |  $X[38;2;255;255;255;22mCurrent Session (Hours):$X[0m " + ((($time =([TimeSpan]::Parse($Script:SessionTimer))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line4 = ("                   |  $X[38;2;255;255;255;22mScript Launch Counter:$X[0m   " + $CurrentStats.TimesLaunched)
	$Line5 = "                    --------------------------------"
	$Line6 = ("                   |  $X[38;2;255;165;000;22mHR's$X[0m Found:              " + $(if ($CurrentStats.HighRunesFound -eq "") {"0"} else {$CurrentStats.HighRunesFound}))
	$Line7 = ("                   |  $X[38;2;165;146;99;22mUnique$X[0m Items Found:      " + $(if ($CurrentStats.UniquesFound -eq "") {"0"} else {$CurrentStats.UniquesFound}))
	$Line8 = ("                   |  $X[38;2;0;225;0;22mSet$X[0m Items Found:         " + $(if ($CurrentStats.SetItemsFound -eq "") {"0"} else {$CurrentStats.SetItemsFound}))
	$Line9 = ("                   |  $X[38;2;255;255;0;22mRare$X[0m Items Found:        " + $(if ($CurrentStats.RaresFound -eq "") {"0"} else {$CurrentStats.RaresFound}))
	$Line10 = ("                   |  $X[38;2;65;105;225;22mMagic$X[0m Items Found:       " + $(if ($CurrentStats.MagicItemsFound -eq "") {"0"} else {$CurrentStats.MagicItemsFound}))
	$Line11 = ("                   |  $X[38;2;255;255;255;22mNormal$X[0m Items Found:      " + $(if ($CurrentStats.NormalItemsFound -eq "") {"0"} else {$CurrentStats.NormalItemsFound}))
	$Line12 = "                    --------------------------------"
	$Line13 = ("                   |  $X[38;2;165;146;99;22mCow King Killed:$X[0m         " + $(if ($CurrentStats.CowKingKilled -eq "") {"0"} else {$CurrentStats.CowKingKilled}))
	$Line14 = ("                   |  $X[38;2;255;0;255;22mGems Activated:$X[0m          " + $(if ($CurrentStats.Gems -eq "") {"0"} else {$CurrentStats.Gems}))
	$Line15 = ("                   |  $X[38;2;255;0;255;22mPerfect Gems:$X[0m            " + $(if ($CurrentStats.PerfectGems -eq "") {"0"} else {$CurrentStats.PerfectGems}))
	$Line16 = "                    --------------------------------"
	$Lines = @($Line1,$Line2,$Line3,$Line4,$Line5,$Line6,$Line7,$Line8,$Line9,$Line10,$Line11,$Line12,$Line13,$Line14,$Line15,$Line16)
	$LongestObject = $null
	# Loop through each object in the array to find longest line (for formatting)
	foreach ($Line in $Lines) {
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -gt $LongestLine) {
			$LongestLine = ($Line -replace '\[.*?22m', '' -replace '\[0m','').Length
		}
	}
	Foreach ($Line in $Lines){#Formatting nonsense to indent things nicely
		$Indent = ""
		$Dash = ""
		write-host $Line -nonewline
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -lt $LongestLine + 2){
			if ($line -notmatch "-"){
				while ((($Line -replace '\[.*?22m', '' -replace '\[0m','').Length + $Indent.length) -lt ($LongestLine + 2)){
					$Indent = $Indent + " "
				}
				write-host ($Indent + "|") -nonewline
				write-host
			}
			else {
				while (($Line.Length + $Dash.length) -lt ($LongestLine)){
					$Dash = $Dash + "-"
				}
				write-host $Dash -nonewline
				write-host
			}
		}
		Else {
			write-host " |"
		}
	}
	write-host
	write-host ("  Chance to find $X[38;2;65;105;225;22mMagic$X[0m quality quote or better: " + [math]::Round((($QualityArraySum - $NormalProbability + 1) * (1/$QualityArraySum) * 100),2) + "%" )
	write-host
	write-host ("  D2r Game Version:    " + (Get-Command "$GamePath\D2R.exe").FileVersionInfo.FileVersion)
	write-host ("  Script Auth Method:  " + $Script:Config.AuthenticationMethod)
	write-host "  Script Install Path: " -nonewline
	write-host ("`"$Script:WorkingDirectory`"" -replace "((.{1,52})(?:\\|\s|$)|(.{1,53}))", "`n                        `$1").trim() #add two spaces before any line breaks for indenting. Add line break for paths that are longer than 53 characters.	
	write-host
	write-host "  Your Script Version: v$CurrentVersion"
	write-host "  https://github.com/shupershuff/Diablo2RLoader/releases/tag/v$CurrentVersion"
	if($Script:LatestVersionCheck -eq $null -or $Script:LatestVersionCheck.tostring() -lt (Get-Date).addhours(-2).ToString('yyyy.MM.dd HH:mm:ss')){ #check for updates. Don't check if this has been checked in the couple of hours.
		try {
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
			$ReleaseInfo = ($Releases | sort id -desc)[0] #find release with the highest ID.
			$Script:LatestVersionCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
			$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		}
		Catch {
			write-verbose "  Couldn't check for updates :("
		}
	}
	
	if ($Null -ne $Script:LatestVersion -and $Script:LatestVersion -gt $Script:CurrentVersion) {
		write-host
		write-host "  Latest Script Version: v$LatestVersion" -foregroundcolor yellow
		write-host "  https://github.com/shupershuff/Diablo2RLoader/releases/latest" -foregroundcolor yellow
	}
	write-host
	if ($Script:NotificationsAvailable -eq $True){
		write-host "  $X[38;2;255;165;000;48;2;1;1;1;4mNotification:$X[0m" -nonewline
		Notifications -check $False
		$Script:NotificationHasBeenChecked = $True
		write-host

	}
	write-host
	PressTheAnyKey
}

Function Notifications {
	param (
		[bool] $Check
	)
	if ($Check -eq $True -and $Script:LastNotificationCheck -lt (Get-Date).addminutes(-30).ToString('yyyy.MM.dd HH:mm:ss')){#check for notifications once every 30mins
		try {
			$URI = "https://raw.githubusercontent.com/shupershuff/Diablo2RLoader/main/Notifications.txt"
			$Script:Notifications = Invoke-RestMethod -Uri $URI
			if ($Notifications.notification -ne "") {
				if ($Script:PrevNotification -ne $Notifications.notification){#if message has changed since last check
					$Script:PrevNotification = $Notifications.notification	
					$Script:NotificationHasBeenChecked = $False
					if ((get-date).tostring('yyyy.MM.dd HH:mm:ss') -lt $Notifications.ExpiryDate){
						$Script:NotificationsAvailable = $True
					}
				}
			}
			Else {
				$Script:NotificationsAvailable = $False
			}
			$Script:LastNotificationCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
		}
		Catch {
			#Write-Host "  Couldn't check for notifications." -foregroundcolor Yellow #Commented out, for debug only. If this fails in production don't show any errors/warnings.
		}
	}
	Elseif ($Check -eq $False) {
		write-host
		$Notifications.notification -split "`n" | ForEach-Object {
			$_ = " " + $_
			if ($_[1] -eq "-") {#for any line starting with a dash
				 $DashFormat = ($_ -replace "(.{1,73})(\s+|$)", "`$1`n").trimend()
				 $DashFormat -split "`n" | ForEach-Object {
					if ($_[1] -eq "-") {#for any line starting with a dash
						" " + $_
					}
					else {
						($_ -replace "(.{1,73})(\s+|$)", "    `$1`n").trimend() 
					}
				}
			}			
			else {
				($_ -replace "(.{1,74})(\s+|$)", " `$1`n ").trimend()
			}
		}
	}
	if ($Check -eq $True -and $Script:NotificationHasBeenChecked -eq $False -and $Script:NotificationsAvailable -eq $True){#only show message if user hasn't seen notification yet.
		write-host
		write-host "  $X[38;2;255;165;000;48;2;1;1;1;4mNotification available. See Info screen for details.$X[0m"
	}
}

$D2rLevels =@(
	@(1, "Rogue Encampment"),
	@(2, "Blood Moor"),
	@(3, "Cold Plains"),
	@(4, "Stony Field"),
	@(5, "Dark Wood"),
	@(6, "Black Marsh"),
	@(7, "Tamoe Highland"),
	@(8, "Den of Evil"),
	@(9, "Cave 1"),
	@(10, "Underground Passage 1"),
	@(11, "Hole 1"),
	@(12, "Pit 1"),
	@(13, "Cave 2"),
	@(14, "Underground Passage 2"),
	@(15, "Hole 2"),
	@(16, "Pit 2"),
	@(17, "Burial Grounds"),
	@(18, "Crypt"),
	@(19, "Mausoleum"),
	@(20, "Forgotten Tower"),
	@(21, "Tower Cellar 1"),
	@(22, "Tower Cellar 2"),
	@(23, "Tower Cellar 3"),
	@(24, "Tower Cellar 4"),
	@(25, "Tower Cellar 5"),
	@(26, "Monastery Gate"),
	@(27, "Outer Cloister"),
	@(28, "Barracks"),
	@(29, "Jail 1"),
	@(30, "Jail 2"),
	@(31, "Jail 3"),
	@(32, "Inner Cloister"),
	@(33, "Cathedral"),
	@(34, "Catacombs 1"),
	@(35, "Catacombs 2"),
	@(36, "Catacombs 3"),
	@(37, "Catacombs 4"),
	@(38, "Tristram"),
	@(39, "The Secret Cow Level"),
	@(40, "Lut Gholein"),
	@(41, "Rocky Waste"),
	@(42, "Dry Hills"),
	@(43, "Far Oasis"),
	@(44, "Lost City"),
	@(45, "Valley of Snakes"),
	@(46, "Canyon of the Magi"),
	@(47, "Sewers 1"),
	@(48, "Sewers 2"),
	@(49, "Sewers 3"),
	@(50, "Harem 1"),
	@(51, "Harem 2"),
	@(52, "Palace Cellar 1"),
	@(53, "Palace Cellar 2"),
	@(54, "Palace Cellar 3"),
	@(55, "Stony Tomb 1"),
	@(56, "Halls of the Dead 1"),
	@(57, "Halls of the Dead 2"),
	@(58, "Claw Viper Temple 1"),
	@(59, "Stony Tomb 2"),
	@(60, "Halls of the Dead 3"),
	@(61, "Claw Viper Temple 2"),
	@(62, "Maggot Lair 1"),
	@(63, "Maggot Lair 2"),
	@(64, "Maggot Lair 3"),
	@(65, "Ancient Tunnels"),
	@(66, "Tal Rashas Tomb 1"),
	@(67, "Tal Rashas Tomb 2"),
	@(68, "Tal Rashas Tomb 3"),
	@(69, "Tal Rashas Tomb 4"),
	@(70, "Tal Rashas Tomb 5"),
	@(71, "Tal Rashas Tomb 6"),
	@(72, "Tal Rashas Tomb 7"),
	@(73, "Tal Rashas Chamber"),
	@(74, "Arcane Sanctuary"),
	@(75, "Kurast Docks"),
	@(76, "Spider Forest"),
	@(77, "Great Marsh"),
	@(78, "Flayer Jungle"),
	@(79, "Lower Kurast"),
	@(80, "Kurast Bazaar"),
	@(81, "Upper Kurast"),
	@(82, "Kurast Causeway"),
	@(83, "Travincal"),
	@(84, "Archnid Lair"),
	@(85, "Spider Cavern"),
	@(86, "Swampy Pit 1"),
	@(87, "Swampy Pit 2"),
	@(88, "Flayer Dungeon 1"),
	@(89, "Flayer Dungeon 2"),
	@(90, "Swampy Pit 3"),
	@(91, "Flayer Dungeon 3"),
	@(92, "Sewers 1"),
	@(93, "Sewers 2"),
	@(94, "Ruined Temple"),
	@(95, "Disused Fane"),
	@(96, "Forgotten Reliquary"),
	@(97, "Forgotten Temple"),
	@(98, "Ruined Fane"),
	@(99, "Disused Reliquary"),
	@(100, "Durance of Hate 1"),
	@(101, "Durance of Hate 2"),
	@(102, "Durance of Hate 3"),
	@(103, "Pandemonium Fortress"),
	@(104, "Outer Steppes"),
	@(105, "Plains of Despair"),
	@(106, "City of the Damned"),
	@(107, "River of Flame"),
	@(108, "Chaos Sanctuary"),
	@(109, "Harrogath"),
	@(110, "Bloody Foothills"),
	@(111, "Frigid Highlands"),
	@(112, "Arreat Plateau"),
	@(113, "Crystalline Passage"),
	@(114, "Frozen River"),
	@(115, "Glacial Trail"),
	@(116, "Drifter Cavern"),
	@(117, "Frozen Tundra"),
	@(118, "The Ancients Way"),
	@(119, "Icy Cellar"),
	@(120, "Arreat Summit"),
	@(121, "Nihlathaks Temple"),
	@(122, "Halls of Anguish"),
	@(123, "Halls of Pain"),
	@(124, "Halls of Vaught"),
	@(125, "Abaddon"),
	@(126, "Pit of Acheron"),
	@(127, "Infernal Pit"),
	@(128, "Worldstone Keep 1"),
	@(129, "Worldstone Keep 2"),
	@(130, "Worldstone Keep 3"),
	@(131, "Throne of Destruction"),
	@(132, "Worldstone Keep")
)

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
"We live...AGAIN!",
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
"What's the matter, hero? Questioning your fortitude? I know we are.",
"This whole place is one big ale fog.",
"So, this is daylight... It's over-rated.",
"When - or if - I get to Lut Gholein, I'm going to find the largest bowl`nof Narlant weed and smoke 'til all earthly sense has left my body.",
"I've just about had my fill of the walking dead.",
"Oh I hate staining my hands with the blood of foul Sorcerers!",
"Damn it, I wish you people would just leave me alone!",
"Beware! Beyond lies mortal danger for the likes of you!",
"Beware! The evil is strong ahead.",
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
"Live, Laugh, Love - Andariel, 1264.",
"Oh no, snakes. I hate snakes.",
"Who would have thought that such primitive beings could cause so much `ntrouble.",
"Hail to you champion"

Function BannerLogo {
	$BannerLogo = @"

  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%#%%%%%%%%%%%%%%%%%%%%  
  %%%%%%%#%%%%%%%/%%%%%%%%%%%#%%%%%$Script:MOO%%%%##%%%%%%%%%#/##%%%%%%%%%%%%%%%%%%  
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
	if ($Script:PGemActivated -eq $True){
		Write-Host "  $X[38;2;255;165;0;22m$BannerLogo$X[0m"
	}
	Else {
		Write-Host $BannerLogo -foregroundcolor yellow
	}
}
Function JokeMaster {
	#If you're going to leech and not provide any damage value in the Throne Room then at least provide entertainment value right?
	Write-Host "  Copy these mediocre jokes into the game while doing Baal Comedy Club runs`r`n  to mislead everyone into believing you have a personality:"
	Write-Host
	$attempt = 0
	$JokeProviderRoll = get-random -min 1 -max 3 #Randomly roll for a joke provider
	do {
		$attempt ++
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
	} until ($JokeObtained -eq $true -or $attempt -eq 3)
	Write-Host;	Write-Host "  Joke courtesy of $JokeProvider"
	Write-Host;	Write-Host
	PressTheAnyKey
}

Function DClone {# Display DClone Status.
	param (
		[bool] $DisableOutput,
		[String] $D2CloneTrackerSource,
		[String] $Taglist,
		[object] $DCloneChanges,
		[String] $DCloneAlarmLevel
	)
	if ($D2CloneTrackerSource -eq "d2rapi.fly.dev"){
		$URI = "https://d2rapi.fly.dev/dclone"
		try {
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $D2CloneTrackerSource -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET
			} -TimeoutSeconds 3
			$CurrentStatus = $D2RDCloneResponse.PSObject.Properties | select @{Name='Server'; Expression={$_.name}},@{Name='Progress'; Expression={($_.value + 1)}} #| sort server #add +1 as this source counts status from 0
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function
		}
	}
	elseif ($D2CloneTrackerSource -eq "D2runewizard.com"){
		$QLC = "zouaqcSTudL"
		$tokreg = ("QLC" + $qlc + 1 +"fnbttzP")
		$D2RWref = ""
		for ($i = $tokreg.Length - 1; $i -ge 0; $i--) {
			$D2RWref += $tokreg[$i]
		}
		$headers = @{
			"D2R-Contact" = "placeholderemail@email.com"
			"D2R-Platform" = "GitHub"
			"D2R-Repo" = "https://github.com/shupershuff/Diablo2RLoader"
		}
		try {
			$URI = "https://d2runewizard.com/api/diablo-clone-progress/all?token=$D2RWref"
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $D2CloneTrackerSource -Headers -$headers -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET -Headers $using:Headers
			} -TimeoutSeconds 3
			$CurrentStatus = $D2RDCloneResponse.servers | select @{Name='Server'; Expression={$_.server}},@{Name='Progress'; Expression={$_.progress}} #| sort server
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function	
		}
	}
	elseif ($D2CloneTrackerSource -eq "Diablo2.io"){
		$headers = @{
			"User-Agent" = "github.com/shupershuff/Diablo2RLoader"
		}
		$URI = "https://diablo2.io/dclone_api.php"
		try {
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $D2CloneTrackerSource -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET -Headers $using:Headers
			} -TimeoutSeconds 3
			$CurrentStatus = $D2RDCloneResponse | select @{Name='Server'; Expression={$_.region}},@{Name='Ladder'; Expression={$_.ladder}},@{Name='Core'; Expression={$_.hc}},@{Name='Progress'; Expression={$_.progress}}
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function
		}
	}
	Else {#if XML is invalid for D2CloneTrackerSource
		$DCloneErrorMessage = ("  Error: Couldn't check for DClone Status. ###  Check DCloneTrackerSource in config.xml is entered correctly.").Replace("###", "`n")
		Write-host
		Write-host $DCloneErrorMessage -Foregroundcolor red
		if ($DisableOutput -ne $True){
			Write-host
			$Script:AccountID = $null
			Presstheanykey
		}
		Return
	}
	if ($DCloneErrorMessage -ne $null){
		Write-host $DCloneErrorMessage -Foregroundcolor red
		if ($DisableOutput -ne $True){
			Write-host
			$Script:AccountID = $null
			Presstheanykey
		}
		Return
	}
	$DCloneLadderTable = New-Object -TypeName System.Collections.ArrayList
	$DCloneNonLadderTable = New-Object -TypeName System.Collections.ArrayList
	if ($DCloneChanges -eq "" -or $DCloneChanges -eq $null){
		$DCloneChangesArray = New-Object -TypeName System.Collections.ArrayList
	}
	Else {
		$DCloneChangesArray = $DCloneChanges | ConvertFrom-Csv -ErrorAction silentlycontinue #temporarily convert to array
	}
	foreach ($Status in $CurrentStatus){
		$DCloneLadderInfo = New-Object -TypeName psobject
		$DCloneNonLadderInfo = New-Object -TypeName psobject
		#Convert data from all sources into consistent names and tags to be sorted and filtered.
		if ($Status.server -like "*us*" -or $Status.server -like "*americas*" -or $Status.Server -eq "1"){$Tag = "-NA";$ServerName = "Americas"}
		elseif ($Status.server -like "*eu*" -or $Status.server -like "*europe*" -or $Status.Server -eq "2"){$Tag = "-EU";$ServerName = "Europe"}
		elseif ($Status.server -like "*kr*" -or $Status.server -like "*asia*" -or $Status.Server -eq "3"){$Tag = "-KR";$ServerName = "Asia"}
		if (($Status.server -notlike "*nonladder*" -and -not [int]::TryParse($Status.server,[ref]$null)) -or $Status.Ladder -eq "2"){
			if ($Status.server -match "hardcore" -or $Status.Core -eq "2"){$Tag = ("HCL" + $Tag);$ServerName = ("HCL - " + $ServerName)}
			else {$Tag = ("SCL" + $Tag);$ServerName = ("SCL - " + $ServerName)}
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name Tag -Value $Tag
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name LadderServer -Value $ServerName
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name LadderProgress -Value $Status.progress
			[VOID]$DCloneLadderTable.Add($DCloneLadderInfo)		
		}
		Else {
			if ($Status.server -match "hardcore" -or $Status.Core -eq 2){$Tag = ("HC" + $Tag);$ServerName = ("HC - " + $ServerName)}
			else {$Tag = ("SC" + $Tag);$ServerName = ("SC - " + $ServerName)}
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name Tag -Value $Tag
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderServer -Value $ServerName
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderProgress -Value $Status.progress
			[VOID]$DCloneNonLadderTable.Add($DCloneNonLadderInfo)
		}
		if ($True -eq $DisableOutput){	
			#$taglist = "SCL-NA, SCL-EU, HC-EU, HC-KR"#"SCL-NA, SCL-EU, SCL-KR, SC-NA, SC-EU, SC-KR, HCL-NA, HCL-EU, HCL-KR, HC-NA, HC-EU, HC-KR"
			if ($taglist -match $Tag ){#if D Dclone region and server matches what's in config, check for changes.
				#Write-Host " Tag $tag in taglist" #debug
				if ($DCloneChangesArray | where-object {$_.Tag -eq $Tag}){
					foreach ($Item in $DCloneChangesArray | where-object {$_.Tag -eq $Tag}) {#for each tag specified in config.xml...
						$item.VoiceAlarmStatus = $False 
						$item.TextAlarmStatus = $False 
						if ($item.OldStatus -ne $Status.progress){
							if ($Status.progress -ge 5 -and ($DCloneAlarmLevel -match 5)) {#if D Clone walk is imminent
								$item.VoiceAlarmStatus = $True
							}
							elseif ($Status.progress -eq 1 -or $Status.progress -eq 6){#if D Clone walk has just happened
								$item.VoiceAlarmStatus = $True
							}
							elseif ($DCloneAlarmLevel -match $Status.progress){#if User has configured alarms to happen for lower D Clone status levels (2,3,4)
								$item.VoiceAlarmStatus = $True
							}
							$item.TextAlarmStatus = $True 
							$item.OldStatus = $Status.progress
							$item.LastUpdate = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
						}
						elseif ($Status.progress -eq 5){#if status hasn't changed, but status is 5 (imminent), show alarm text on main menu
							$item.TextAlarmStatus = $True 
						}
						elseif ($Status.progress -lt 5 -and $item.LastUpdate -gt (get-date).addminutes(-5).ToString('yyyy.MM.dd HH:mm:ss')){#if status is less than 5 and has changed within the last 5 minutes, enable text alarm
							$item.TextAlarmStatus = $True
						}
						elseif ($item.LastUpdate -ne $null -and $item.LastUpdate -lt (get-date).addminutes(-5).ToString('yyyy.MM.dd HH:mm:ss')){#after 5 minutes remove the text alarm
							$item.LastUpdate = $null 
							$item.TextAlarmStatus = $False
						}
					}
				}
				Else {
					try {
						if ($Status.progress -ge 5){
							$VoiceAlarmStatus = $True
							$TextAlarmStatus = $True
						}
						Else {
							$VoiceAlarmStatus = $False
							$TextAlarmStatus = $False
						}
						$DCloneChangesArray += [PSCustomObject]@{Tag = $Tag; OldStatus = $Status.progress; VoiceAlarmStatus = $VoiceAlarmStatus; TextAlarmStatus = $TextAlarmStatus; LastUpdate = $null}
					}
					Catch {#if script is refreshed too quick and $DCloneChangesArray gets corrupted, reset it.
						$DCloneChangesArray = $null
						Return
					}
				}
				
			}
		}
	}
	if ($True -ne $DisableOutput){
		$DCloneLadderTable = $DCloneLadderTable | sort LadderServer
		$DCloneNonLadderTable = $DCloneNonLadderTable | sort NonLadderServer
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
			do { # formatting nonsense.
				$LadderServer = ($LadderServer + " ")
			}
			until ($LadderServer.length -ge 26)
			$NonLadderServer = ($DCloneNonLadderTable.NonLadderServer[$Count]).tostring()
			do { # formatting nonsense.
				$NonLadderServer = ($NonLadderServer + " ")
			}
			until ($NonLadderServer.length -ge 29)
			Write-Host (" #  " + $LadderServer + " " + $DCloneLadderTable.LadderProgress[$Count] + "    |  " + $NonLadderServer + " " + $DCloneNonLadderTable.NonLadderProgress[$Count]+ "    #")
			$Count = $Count + 1
			if ($Count -eq 6){
				Write-Host " #                                  |                                     #"
				Write-Host " ##########################################################################"
				Write-Host
				Write-Host "   DClone Status provided $D2CloneTrackerSource"
				Write-Host
			}
		}
		Until ($Count -eq 6)
		PressTheAnyKey
	}
	Elseif ($Taglist -ne "" -and $DCloneChangesArray -ne $null){#Else if Output is disabled and taglist has been specified, output dclone changes for alarm
		$DCloneChangesArray | ConvertTo-Csv -NoTypeInformation
	}
}

Function DCloneVoiceAlarm {
	$voice = New-Object -ComObject Sapi.spvoice
	$voice.rate = -2 #How quickly the voice message should be
	Write-Host
	if ($Script:Config.DCloneAlarmVoice -eq "Bloke" -or $Script:Config.DCloneAlarmVoice -eq "Man" -or $Script:Config.DCloneAlarmVoice -eq "Paladin"){$voice.voice = $voice.getvoices() | where {$_.id -like "*David*"}}
	elseif ($Script:Config.DCloneAlarmVoice -eq "Wench" -or $Script:Config.DCloneAlarmVoice -eq "Woman" -or $Script:Config.DCloneAlarmVoice -eq "Amazon"){$voice.voice = $voice.getvoices() | where {$_.id -like "*ZIRA*"}}
	else {break}# If specified voice doesn't exist
	foreach ($Item in ($Script:DCloneChangesCSV | ConvertFrom-Csv) | where-object {$_.VoiceAlarmStatus -Match "True" -or $_.TextAlarmStatus -Match "True"}) {
		if ($item.tag -match "l"){#if mode contains "L"
			$LadderText = "Ladder"
		}
		else {#else it's non ladder
			$LadderText = "NonLadder"
		}
		if ($item.tag -match "HC"){
			$CoreText = "Hardcore"
		}
		Else {#if mode contains "SC"
			$CoreText = "Softcore"
		}
		if ($item.tag -match "NA"){
			$DCloneRegion = "Americas"
		}
		elseif ($item.tag -match "KR"){
			$DCloneRegion = "Asia"
		}
		else {
			$DCloneRegion = "Europe"
		}
		if ($Item.OldStatus -eq 5){#change to 5
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone is about to walk in $DCloneRegion on $CoreText $LadderText ($($item.tag))!$X[0m"
			$Message = ("D Clone Imminent! DClone is about to walk in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		Elseif ($Item.OldStatus -eq 1 -or $Item.OldStatus -eq 6){
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone has just walked in $DCloneRegion on $CoreText $LadderText ($($item.tag)).$X[0m"
			$Message = ("D Clone Has just walked in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		Elseif ($Script:DCloneAlarmLevel -match $Item.OldStatus) {
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone Update! DClone is now $($Item.OldStatus)/6 in $DCloneRegion on $CoreText $LadderText ($($item.tag))$X[0m"
			$Message = ("D Clone is now " + $Item.OldStatus + " out of 6 in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		if ($item.VoiceAlarmStatus -eq $True){
			$voice.speak("$Message") | out-null
		}
	}
	if ($Message -ne $null){
		Write-Host "  $X[38;2;065;105;225;48;2;1;1;1;4mD Clone status provided by $D2CloneTrackerSource$X[0m"
	}
}
Function WebRequestWithTimeOut {#Used to timeout OCR requests that take too long so that another attempt can be made with a different OCR engine.
	param (
		[ScriptBlock] $ScriptBlock,
		[int] $TimeoutSeconds,
		[String] $InitiatingFunction,
		[String] $DCloneSource
	)
	$Script:DCloneErrorMessage = $null
	if ($InitiatingFunction -eq "DClone"){
		$ArgumentList = $URI	
	}
	$TimedJob = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList #Sets the parameter $EngineNumber to $Engine
	$timer = [Diagnostics.Stopwatch]::StartNew()
	while ($TimedJob.State -eq "Running" -and $timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
		Start-Sleep -Milliseconds 10
	}
	if ($TimedJob.State -eq "Running") {
		Stop-Job -Job $TimedJob 
		if ($InitiatingFunction -eq "DClone"){
			Write-Host
			$Script:DCloneErrorMessage = "   Error: Couldn't connect to $D2CloneTrackerSource to check for DClone Status."
			Write-Host ("   Timed out connecting to DClone Data Source.") -foregroundcolor red
			Throw "Timed Out :(" #force an exception to break out of the try statement.
			Write-Host
		}
	}
	elseif ($TimedJob.State -eq "Completed") {
		$result = Receive-Job -Job $TimedJob 
		#Write-Host "  Command completed successfully."
		foreach ($Object in $result){#Remove Properties from Result Array inserted by the Start-Job command. This prevents skewed data for DClone status checks.
			$Object.PSObject.Properties.Remove("RunspaceId")
			$Object.PSObject.Properties.Remove("PSComputerName")
			$Object.PSObject.Properties.Remove("PSShowComputerName")
		}
		$result
	}
	else {
		if ($InitiatingFunction -eq "DClone"){
			Write-Host " Couldn't connect to DClone Source." -foregroundcolor red	
		}
	}
	Remove-Job -Job $TimedJob 
}
Function TerrorZone {
	# Get the current time data was pulled
	$TimeDataObtained = (Get-Date -Format 'h:mmtt')	
	$TZProvider = "D2Emu.com"
	$TZURI = "https://www.d2emu.com/api/v1/tz"  
	$D2TZResponse = Invoke-RestMethod -Uri $TZURI
	foreach ($Level in $D2TZResponse.current){
		#Write-host "level ID is:" $Level
		foreach ($LevelID in $D2rLevels){
			if ($LevelID[0] -eq $Level){
				$CurrentTZ += $LevelID[1] + ", "
			}
		}
	}
	$CurrentTZ = $CurrentTZ -replace '..$', ''
	foreach ($Level in $D2TZResponse.next){
		#Write-host "level ID is:" $Level
		foreach ($LevelID in $D2rLevels){
			if ($LevelID[0] -eq $Level){
				$NextTZ += $LevelID[1] + ", "
			}
		}
	}
	$NextTZ = $NextTZ -replace '..$', ''
	if ($CheckForNextTZ -eq $False){# Messages if things are unsuccessful pulling Next TZ
		Write-Host
		Write-Host "   Current TZ is: "  -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta #print next tz. The regex code helps format the string into an indented new line for longer TZ's.
		if ($CheckForNextTZ -ne $False){
			$FailMessage += "`n                             Attempts made: $Script:OCRAttempts`n                             Token used:      $tokreg"
			Write-Host "   Next TZ info unavailable: $FailMessage" -ForegroundColor Red
		}
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  TZ info courtesy of:    $TZProvider"
		Write-Host
		PressTheAnyKey
	}
	else {# Messages if things are succesful pulling Next TZ
		Write-Host
		Write-Host "   Current TZ is:  " -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host "   Next TZ is:     " -nonewline;Write-Host ($NextTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  TZ info courtesy of:       $TZProvider"
		Write-Host
		PressTheAnyKey
	}
}
Function TerrorZoneOld {
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
		$script:D2ROCRref = $Null
		for ($i = $tokreg.Length - 1; $i -ge 0; $i--) {
			$script:D2ROCRref += $tokreg[$i]
		}
		$Script:OCRAttempts = 0
		do { ## Try each OCR engine until success. I've noticed that OCR.Space engines sometimes go down causing the script not being able to detect current TZ. This gives it a bit more reliability, albeit at the cost of time for each additional attempt.
			$Script:OCRAttempts ++
			$Engine = (2,1,3)[$Script:OCRAttempts-1] #try engines in order of the best quality: Engine 2, then 1 then 3.
			$NextTZOCR = WebRequestWithTimeOut -InitiatingFunction "TZ_OCR" -ScriptBlock {
				param ($EngineNumber) #Run the scriptblock with $EngineNumber as a parameter. This way we can insert the variable $Engine into the scriptblock. It's not possible to insert this variable into the script block any other way.
				(((Invoke-WebRequest -Uri ("https://api.ocr.space/parse/imageurl?apikey=" + $Using:D2ROCRref + "&filetype=png&isCreateSearchablePdf=false&OCREngine=" + $EngineNumber + "&scale=true&url=https://thegodofpumpkin.com/terrorzones/terrorzone.png") -ErrorAction Stop).Content | ConvertFrom-Json).ParsedResults.ParsedText)
			} -TimeoutSeconds 13 #allow up to 13 seconds per request. Standard requests are around 3 seconds but as it's a free service it can sometimes be slow. Most of the time when it takes longer than 10 seconds to retrieve OCR it fails. Default timeout from OCR.Space (when an engine is down) is actually 30 seconds before an error is thrown.
			if ($SuccessfulEngine -ne $null){
					$NextTZOCR = $NextTZOCR.replace(": :","::") # Prevents issues if OCR mistakenly reads "::" as ": :" eg: Next TerrorZone(s) : : Barracks / Jail 1 / Jail 2 / Jail 3
				if ($SuccessfulEngine -eq 2){#For Engine 2
					$NextTZ = [regex]::Match($NextTZOCR.Replace("`n", ""), "(?<=Next TerrorZone(?:\s)?(?:\(s\))?(?:\s)?(?:\.|::) ).*").Value #Regex for engine 2
				} Elseif ($SuccessfulEngine -eq 1){#For Engine 1
					$NextTZ = ($NextTZOCR.trim() -split "`n")[-1].replace("Next TerrorZone(s)","").replace(".","").replace(":","").trim()  #Regex for engine 1
				} Elseif ($SuccessfulEngine -eq 3) {#For Engine 3
					$NextTZ = $nexttzocr -replace "\r?\n", " " -replace "(?s).*Next TerrorZone\Ss\S\s","" #Regex for engine 3
				}
			}
			$Script:SuccessfulEngine = $Null
			if ($OCRSuccess -eq $False){
				$FailMessage = "Was unable to pull next TZ details :("
			}
			if ($NextTZ -eq ""){#if next tz variable is empty due to OCR not working.
				$FailMessage = "OCR failure, couldn't read next TZ details :("
				$Script:OCRSuccess = $False
				if ($Script:OCRAttempts -ne 3){
					Write-Host "   First Crack failed, having attempt number" ($Script:OCRAttempts + 1)
					$TryAgain = $True
				}
				Else {
					$TryAgain = $False
				}
			}
		} until ($true -eq $Script:OCRSuccess -or ($Script:OCRAttempts -eq 3 -and $OCRSuccess -eq $false -and $TryAgain -ne $true))
	}
	if ($Script:OCRSuccess -eq $False -or $CheckForNextTZ -eq $False){# Messages if things are unsuccessful pulling Next TZ
		Write-Host
		Write-Host "   Current TZ is: "  -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta #print next tz. The regex code helps format the string into an indented new line for longer TZ's.
		if ($CheckForNextTZ -ne $False){
			$FailMessage += "`n                             Attempts made: $Script:OCRAttempts`n                             Token used:      $tokreg"
			Write-Host "   Next TZ info unavailable: $FailMessage" -ForegroundColor Red
		}
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  Current TZ pulled from:    $CurrentTZProvider"
		Write-Host
		PressTheAnyKey
	}
	else {# Messages if things are succesful pulling Next TZ
		Write-Host
		Write-Host "   Current TZ is:  " -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host "   Next TZ is:     " -nonewline;Write-Host ($NextTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
		Write-Host
		Write-Host "  Information Retrieved at: " $TimeDataObtained
		Write-Host "  Current TZ pulled from:    $CurrentTZProvider"
		Write-Host "  Next TZ pulled from:       $NextTZProvider"
		Write-Host
		PressTheAnyKey
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
		$script:handle_id = $Line | Select-String -Pattern '^(?<g2>.+): Event.*DiabloII Check For Other Instances' | %{$_.Matches.Groups[1].value}
		if ($handle_id){
			$handle_id_populated = $handle_id
		}
		if ($handle_id){
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
		$Script:ActiveIDs = (Get-Process | Where {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "- Diablo II: Resurrected"} | Select-Object MainWindowTitle).mainwindowtitle.substring(0,1) #find all diablo 2 game windows and pull the account ID from the title
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
				$AcctPlayTime = (" " + (($time =([TimeSpan]::Parse($AccountOption.TimeActive))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes) + "   ")  # Add hours + (days*24) to show total hours, then show ":" followed by minutes
			}
			catch {#if account hasn't been opened yet.
				$AcctPlayTime = "   0   "
			}
			if ($AcctPlayTime.length -lt 15){#formatting. Depending on the amount of characters for this variable push it out until it's 15 chars long.
				while ($AcctPlayTime.length -lt 15){
					$AcctPlayTime = " " + $AcctPlayTime
				}
			}
		}
		if ($AccountOption.id -in $Script:ActiveAccountsList.id){#if account is currently active
			$Windowname = (Get-Process | Where {$_.processname -eq "D2r" -and $_.MainWindowTitle -match ($AccountOption.id + "*Diablo II: Resurrected")} | Select-Object MainWindowTitle).mainwindowtitle #Check active game instances to see which accounts are active. As this is based on checking window titles, this will only work for accounts opened from the script
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
	#cls
	if ($Script:ScriptHasBeenRun -eq $true){
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
		Write-Host ("  You have quite a treasure there in that Horadric multibox script v" + $Currentversion)
	}
	Notifications -check $True
	BannerLogo
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
	if ($Batch -ne $Null -or $Script:OpenBatches -eq $true){#if batch has been passed through parameter or if batch wass been selected from the menu.
		$Script:AcceptableBatchIDs = $Null #reset value
		foreach ($ID in $Script:AccountOptionsCSV){
			if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
				$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',') #collate acceptable options of batch ID's
				$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',') #collate acceptable options of account ID's
			}
		}
		$AcceptableBatchValues = @($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort) #Unique list of available batches that can be opened. @ converts this from a PSObject into an array which fixes the issue of -notin not working on PSobjects with only 1 item.
		do {
			if ($Batch -ne $Null -and $Batch -notin $AcceptableBatchValues){#if batch specified in the parameter isn't valid
				$Script:BatchToOpen = $Batch
				$Batch = $Null
				DisplayActiveAccounts
				Write-Host
				Write-Host " Batch specified in Parameter is either incorrect or all accounts in that" -foregroundcolor Yellow
				Write-Host " batch are already open. Adjust your parameter or manually specify below." -foregroundcolor Yellow
				Write-Host
				start-sleep -milliseconds 5000
				exit
			}
			if ($Batch -ne $Null -and $Batch -in $AcceptableBatchValues){#if batch is valid, set variable so that loop can be exited.
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
						if ($AcceptableBatchValues.count -gt 1){
							Write-Host " or " -nonewline
						}
						Write-Host "$X[38;2;255;165;000;22m$Value$X[0m"
					}
				}
				if ($Batch -eq $Null){
					Write-Host " Or Press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				}
				$Script:BatchToOpen = readkey
				Write-Host;	Write-Host
			}
			if ($Script:BatchToOpen -notin $AcceptableBatchValues + "c"){
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
	if ($Script:ParamsUsed -eq $false -and ($Script:RegionOption.length -ne 0 -or $Script:Region.length -ne 0)){
		if ($Script:AskForRegionOnceOnly -ne $true){#reset region
			$Script:Region = ""
			$Script:RegionOption = ""
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
				DClone -DisableOutput $False -D2CloneTrackerSource $D2CloneTrackerSource -TagList $Script:Config.DCloneAlarmList
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "j"){
				JokeMaster
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "i"){
				Inventory #show stats
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
				$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
				if ($Script:GemActivated -ne $True){
					$GibberingGemstone = get-random -minimum 0 -maximum  4095
					if ($GibberingGemstone -eq 69){#nice
						Write-Host "  Perfect Gem Activated" -ForegroundColor magenta
						Write-Host
						Write-Host "     OMG!" -foregroundcolor green
						$Script:PGemActivated = $True
						([int]$CurrentStats.PerfectGems) ++
						SetQualityRolls
						Start-Sleep -milliseconds 3750
					}
					else {
						if ($GibberingGemstone -in 16..32){
							Write-Host "  You Killed the Cow King!" -foregroundcolor green
							Write-Host "  $X[38;2;165;146;99;22mMoooooooo!$X[0m"
							$Script:CowKingActivated = $True
							([int]$CurrentStats.CowKingKilled) ++
							SetQualityRolls
							Start-Sleep -milliseconds 850
						}
						else {
							Write-Host "  Gem Activated" -ForegroundColor magenta
							([int]$CurrentStats.Gems) ++
						}
					}
					$Script:GemActivated = $True
					SetQualityRolls
					try {
						$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
					}
					Catch {
						Write-host "  Couldn't update stats.csv" -foregroundcolor yellow
					}
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
				#cls
				if ($Script:ScriptHasBeenRun -eq $true){
					Write-Host "Account previously opened was:"  -foregroundcolor yellow -backgroundcolor darkgreen
					$Lastopened = @(
						[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$Script:region}#Americas
					)
					Write-Host " " -NoNewLine
					Write-Host ("Account:  " + $Lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen 
					Write-Host " " -NoNewLine
					Write-Host "Region:  " $Lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
				}
				Notifications -check $True
				BannerLogo
				QuoteRoll	
			}
			CheckActiveAccounts
			DisplayActiveAccounts
			if ($Script:Config.TrackAccountUseTime -eq $True){
				$OpenD2LoaderInstances = Get-WmiObject -Class Win32_Process | Where-Object { $_.name -eq "powershell.exe" -and $_.commandline -match $Script:ScriptFileName} | select name,processid,creationdate | sort creationdate -descending
				if ($OpenD2LoaderInstances.length -gt 1){#If there's more than 1 D2loader.ps1 script open, close until there's only 1 open to prevent the time played accumulating too quickly.
					foreach ($Process in $OpenD2LoaderInstances[1..($OpenD2LoaderInstances.count -1)]){
						Stop-Process -id $Process.processid -force #Closes oldest running d2loader script
					}
				}
				if($Script:ActiveAccountsList.id.length -ne 0){#if there are active accounts open add to total script time
					#Add time for each account that's open
					$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv"
					$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date) #work out elapsed time to add to accounts.csv	
					foreach ($AccountID in $Script:ActiveAccountsList.id |sort){ #$Script:ActiveAccountsList.id				
						$AccountToUpdate = $Script:AccountOptionsCSV | Where-Object {$_.ID -eq $accountID}
						if ($AccountToUpdate) {
							try {#get current time from csv and add to it
								$AccountToUpdate.TimeActive = [TimeSpan]::Parse($AccountToUpdate.TimeActive) + $AdditionalTimeSpan
							}
							Catch {#if CSV hasn't been populated with a time yet.
								$AccountToUpdate.TimeActive = $AdditionalTimeSpan
							}
						}
						try {
							$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update accounts.csv with the new time played.
						}
						Catch {
							$WriteAcctCSVError = $True
						}
					}
					$Script:SessionTimer = $Script:SessionTimer + $AdditionalTimeSpan #track current session time but only if a game is running
					if ($WriteAcctCSVError -eq $true){
						Write-host
						Write-host "  Couldn't update accounts.csv with playtime info." -ForegroundColor Red
						Write-host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
						start-sleep -milliseconds 1500
						$WriteAcctCSVError = $False
					}
					#Add Time to Total Script Time only if there's an open game.
					$CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
					try {
						$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
						try {#get current time from csv and add to it
							$CurrentStats.TotalGameTime = [TimeSpan]::Parse($CurrentStats.TotalGameTime) + $AdditionalTimeSpan
						}
						Catch {#if CSV hasn't been populated with a time yet.
							$CurrentStats.TotalGameTime = $AdditionalTimeSpan
						}
						$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
					}
					Catch {
						Write-host
						Write-host "  Couldn't update Stats.csv with playtime info." -ForegroundColor Red
						Write-host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
						start-sleep -milliseconds 1500
					}
				}
				$Script:StartTime = Get-Date #restart timer for session time and account time.
			}
			Else {
				$Script:StartTime = Get-Date #restart timer for session time only
			}
			$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
			foreach ($AccountOption in $Script:AccountOptionsCSV){
				if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
					$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
				}
			}
			$accountoptions = ($Script:AcceptableValues -join  ", ").trim()
			#DClone Alarm check
			$GetDCloneFunc = $(Get-Command DClone).Definition
			$GetWebRequestFunc = $(Get-Command WebRequestWithTimeOut).Definition
			if ($Script:Config.DCloneAlarmList -ne ""){#if DClone alarms should be checked on refresh	
				try {
					if ($Script:DCloneChangesCSV -ne $null){
						$Script:DCloneChangesCSV = Receive-Job $Script:DCloneJob
						#$Script:DCloneChangesCSV #debugging
						if($Script:DCloneChangesCSV -match "true"){#if any of the text contains True
							DCloneVoiceAlarm #Create Voice Alarm
						}
					}
					Else {
						$Script:DCloneChangesCSV = ""#if menu is refreshed too quick
					}
					remove-job * -force
					$Script:DCloneJob = Start-Job -ScriptBlock {
						Invoke-Expression "function Dclone {$using:GetDCloneFunc}"
						Invoke-Expression "function WebRequestWithTimeOut {$Using:GetWebRequestFunc}"
						Dclone -DisableOutput $True -D2CloneTrackerSource $Using:D2CloneTrackerSource -TagList $Using:Config.DCloneAlarmList -DCloneChanges $using:DCloneChangesCSV -DCloneAlarmLevel $Using:DCloneAlarmLevel
					} #check for dclone status
				}
				catch {
					Write-Host
					Write-Host " Unable to check for DClone status via $D2CloneTrackerSource." -Foregroundcolor Red
					Write-Host " Try restarting script or changing the source in config.xml." -Foregroundcolor Red
				} 
			}
			do {
				Write-Host
				$Script:OpenAllAccounts = $False
				if ($accountoptions.length -gt 0){#if there are unopened account options available
					if ($Script:Config.ManualSettingSwitcherEnabled -eq $true){
						$ManualSettingSwitcherOption = "s"
						$ManualSettingSwitcherMenuText = "'$X[38;2;255;165;000;22ms$X[0m' to toggle the Manual Setting Switcher, "
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
						Write-Host ("  Select which account to sign into: " + "$X[38;2;255;165;000;22m$accountoptions$X[0m" + $AllAccountMenuTextNoBatch)
						Write-Host "  Alternatively choose from the following menu options:"
						$BatchMenuText = ""
					}
					else {
						$Script:BatchToOpen = $Null
						$BatchMenuText = "'$X[38;2;255;165;000;22mb$X[0m' to open a Batch of accounts,"
						$Script:AcceptableBatchIDs = $Null #reset value
						$AcceptableBatchValues = $Null
						foreach ($ID in $Script:AccountOptionsCSV){
							if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
								$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',')
								$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',')
							}
						}
						$AcceptableBatchValues = ($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort) #Unique list of available batches that can be opened
						if ($AcceptableBatchValues -eq $Null){
							$BatchOption = ""
							$BatchMenuText = ""
						}
						Else {
							$BatchOption = "b"
						}
						Write-Host "  Enter the ID# of the account you want to sign into."
						Write-Host "  Alternatively choose from the following menu options:"
						Write-Host ("  " + $AllAccountMenuText + $BatchMenuText )#+ "'$X[38;2;255;165;000;22mc$X[0m' to close All accounts,"	 )
					}
				}
				else {#if there aren't any available options, IE all accounts are open
					$AllOption = $Null
					$BatchOption = $Null
					Write-Host " All Accounts are currently open!" -foregroundcolor yellow
				}
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, '$X[38;2;255;165;000;22mt$X[0m' for TZ info, '$X[38;2;255;165;000;22md$X[0m' for DClone status, '$X[38;2;255;165;000;22mj$X[0m' for jokes,"
				Write-Host "  $ManualSettingSwitcherMenuText'$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: " -nonewline
				$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
				if ($Script:AccountID -notin ($Script:AcceptableValues + "x" + "r" + "t" + "d" + "g" + "j" + "i" + $ManualSettingSwitcherOption + $AllOption + $BatchOption) -and $Null -ne $Script:AccountID){
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
				if ($Script:AccountID -eq "x"){
					Write-Host
					Write-Host " Good day to you partner :)" -foregroundcolor yellow
					Start-Sleep -milliseconds 486
					Exit
				}
				$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
			}
		} until ($Script:AccountID -ne "r" -and $Script:AccountID -ne "t" -and $Script:AccountID -ne "d" -and $Script:AccountID -ne "g" -and $Script:AccountID -ne "j" -and $Script:AccountID -ne "s" -and $Script:AccountID -ne "i")
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
			Write-Host (" Alternatively select '$X[38;2;255;165;000;22mc$X[0m' to cancel or press enter for the default (" + $Config.DefaultRegion + "-" + ($Script:ServerOptions | Where-Object {$_.option -eq $Config.DefaultRegion}).region + "): ") -nonewline
			$Script:RegionOption = ReadKey
			Write-Host
			if ("" -eq $Script:RegionOption){
				$Script:RegionOption = $Config.DefaultRegion #default to NA
			}
			else {
				$Script:RegionOption = $Script:RegionOption.tostring()
			}
			if ($Script:RegionOption -notin $Script:ServerOptions.option + "c"){
				Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
				write-Host
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
		if (($Script:ConvertPlainTextPasswords -ne $false -and $Script:ParamsUsed -ne $true) -or ($Script:ParamsUsed -eq $true -and ($Script:OpenBatches -eq $True -or $Script:OpenAllAccounts -eq $True))){#Convert password if it's enabled in config and script is being run normally *OR* Convert password if script is being run from paramters using either -all batch or -all (but not if -username is used instead)
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
		if ($Script:Config.AuthenticationMethod -eq "Token"){
			$Script:Token = $Script:AccountChoice.Token.tostring()
			$EncryptedToken = $Script:Token | ConvertTo-SecureString
			$Tokenobject = New-Object System.Management.Automation.PsCredential("N/A", $EncryptedToken)
			$Token = $Tokenobject.GetNetworkCredential().Password
			$Entropy = @(0xc8, 0x76, 0xf4, 0xae, 0x4c, 0x95, 0x2e, 0xfe, 0xf2, 0xfa, 0x0f, 0x54, 0x19, 0xc0, 0x9c, 0x43)
			# Convert the token and entropy to byte arrays
			$TokenBytes = [System.Text.Encoding]::UTF8.GetBytes($Token)
			$EntropyBytes = [byte[]] $Entropy

			# Encrypt the token
			[void][System.Reflection.Assembly]::LoadWithPartialName("System.Security")
			$ProtectedData = [System.Security.Cryptography.ProtectedData]::Protect($TokenBytes, $EntropyBytes, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)

			$Path = "HKCU:\SOFTWARE\Blizzard Entertainment\Battle.net\Launch Options\OSI"
			Set-ItemProperty -Path $Path -Name "REGION" -Value $Script:Region.Substring(0, 2).ToUpper()
			Set-ItemProperty -Path $Path -Name "WEB_TOKEN" -Value $ProtectedData -Type Binary
		}
		try {
			$Script:AccountFriendlyName = $Script:AccountChoice.accountlabel.tostring()
		}
		Catch {
			$Script:AccountFriendlyName = $Script:AccountUsername
		}	
		#Open diablo with parameters
			# IE, this is essentially just opening D2r like you would with a shortcut target of "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected\D2R.exe" -username <yourusername -password <yourPW> -address <SERVERaddress>
		if ($Script:Config.AuthenticationMethod -eq "Parameter"){
			$arguments = (" -username " + $Script:acct + " -password " + $Script:PW +" -address " + $Script:Region + " " + $Script:AccountChoice.CustomLaunchArguments).tostring()
		}
		else {
			$arguments = (" -uid osi " + $Script:AccountChoice.CustomLaunchArguments).tostring()
		}
		if ($Config.ForceWindowedMode -eq $true){#starting with forced window mode sucks, but someone asked for it.
			$arguments = $arguments + " -w"
		}
		$Script:PW = $Null
		$Script:Token = $Null
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
						PressTheAnyKeyToExit
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
				PressTheAnyKey
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
					PressTheAnyKeyToExit
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
				 $Name = $file.Name -replace '^settings\.|\.json$' #remove 'settings.' and '.json'. The text in between the two periods is the name.
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "Name" -Value $Name
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
				 [VOID]$SettingsFileOptions.Add($SettingsFileOption)
			}
			if ($SettingsFileOptions -ne $Null){# If settings files are found, IE the end user has set them up prior to running script.
				$SettingsFileOptions = $SettingsDefaultOptionArray + $SettingsFileOptions
				Write-Host
				Write-Host "  Settings options you can choose from are:"
				foreach ($Option in $SettingsFileOptions){
					Write-Host ("   " + $Option.ID + ". " + $Option.name) -foregroundcolor green
				}
				do {
					Write-Host "  Choose the settings file you like to load from: " -nonewline
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
					if ($SettingsChoice.tostring() -notin $SettingsFileOptions.id + $SettingsCancelOption){
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
						PressTheAnyKey
					}
				}
			}
			Else {# if no custom settings files are found, IE user hasn't set them up yet.
				Write-Host
				Write-Host "  No Custom Settings files have been saved yet. Loading default settings." -foregroundcolor Yellow
				Write-Host "  See README for setup instructions." -foregroundcolor Yellow
				Write-Host
				PressTheAnyKey
			}
		}
		if ($SettingsChoice -ne "c"){
			#Start Game
			killhandle | out-null
			$process = Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments" -PassThru
			Start-Sleep -milliseconds 1500 #give D2r a bit of a chance to start up before trying to kill handle
			#Close the 'Check for other instances' handle
			Write-Host " Attempting to close `"Check for other instances`" handle..."
			$Output = killhandle | out-string #run killhandle function.
			if (($Output.contains("DiabloII Check For Other Instances")) -eq $true){
				$handlekilled = $true
				Write-Host " `"Check for Other Instances`" Handle closed." -foregroundcolor green
			}
			else {
				Write-Host " `"Check for Other Instances`" Handle was NOT closed." -foregroundcolor red
				Write-Host " Who even knows what happened. I sure don't." -foregroundcolor red
				Write-Host " You may need to kill this manually via procexp. Good luck hero." -foregroundcolor red
				Write-Host
				PressTheAnyKey
			}

			if ($handlekilled -ne $True){
				Write-Host " Couldn't find any handles to kill." -foregroundcolor red
				Write-Host " Game may not have launched as expected." -foregroundcolor red
				PressTheAnyKey
			}
			#Rename the Diablo Game window for easier identification of which account and region the game is.
			$rename = ($Script:AccountID + " - " + $Script:AccountFriendlyName + " (" + $Script:Region + ")" +" - Diablo II: Resurrected")
			#$Command = ('"'+ $WorkingDirectory + '\SetText\SetText.exe" "Diablo II: Resurrected" "' + $rename +'"')
			#$Command = ('"'+ $WorkingDirectory + '\SetText\SetTextv2.exe" /WindowToRename "Diablo II: Resurrected" "' + $rename +'"')
			$Command = ('"'+ $WorkingDirectory + '\SetText\SetTextv2.exe" /PID ' + $process.id + ' "' + $rename + '"')
			try {
				cmd.exe /c $Command
				#write-output $Command #debug
				Write-Host " Window Renamed." -foregroundcolor green
				Start-Sleep -milliseconds 250
			}
			catch {
				Write-Host " Couldn't rename window :(" -foregroundcolor red
				PressTheAnyKey
			}
			if ($Script:Config.AuthenticationMethod -eq "Token"){#wait for web_token to change twice (once for launch, once for char select screen, before being able to launch additional accounts. Token will have already changed once by the time script reaches this stage
				$CurrentTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
				write-host " Waiting for you to get to Character select screen..." -foregroundcolor yellow
				write-host " $X[38;2;255;255;0;4mDO NOT open or close another game instance until this is done.$X[0m"
				do {
					$NewTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
					$CompareCheck = Compare-Object $CurrentTokenRegValue $NewTokenRegValue
					if ($CompareCheck -ne $Null){#if comparecheck has some value, this means it found differences, IE the reg value changed.
						$CurrentTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
						$CurrentTokenRegValue = $NewTokenRegValue
						$WebTokenChangeCounter++
						write-host " Game Launch Successful for $Script:AccountFriendlyName!" -foregroundcolor green
					}
					else {
						Start-Sleep -milliseconds 451
					}
				} until ($WebTokenChangeCounter -eq 1)
			}
			if ($Script:LastAccount -eq $True -or ($Script:OpenAllAccounts -ne $True -and $Script:OpenBatches -ne $True)){
				Write-Host
				Write-Host "Good luck hero..." -foregroundcolor magenta
			}
			Start-Sleep -milliseconds 1000
			$Script:ScriptHasBeenRun = $true
		}
	}
}

ImportCSV
#cls
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
