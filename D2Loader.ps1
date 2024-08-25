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

Changes since 1.12.0 (next version edits):
New feature! You can now enable 'RememberWindowLocations' so that the script moves the game windows to your preferred locations at launch. To use, go to the options menu and choose to save coordinates (once enabled). Big thanks to Sir-Wilhelm for providing code to repurpose.
Added Options menu to be able to edit some of the common config from within script.
Fixed a typo when launching with Authtokens.
Fixed region display not working properly for account labels with brackets.
Minor change to notifications (it won't announce if todays date is less than publishdate).
Made window size slightly taller.
Script now checks accounts.csv to see if batches are used. EnableBatchFeature in config.xml is now redundant and will be removed.
Removed CheckForNextTZ from config.xml as it's redundant.
Removed AskForRegionOnceOnly from config.xml as it's not that useful.
Change 'ConvertPlainTextPasswords' to 'ConvertPlainTextSecrets'. This now aligns to both Passwords and tokens for those that would prefer to store in plain text. Will not convert already secured secrets to plain text.
Fixed up some error handling with the Joke screen.
Fixed non-numeric account ID's not displaying (Thanks loodakrawa)
Script now removes any empty rows accidentally left in accounts.csv to prevent issues.
Script now specifies friendly name if none is entered into accounts.csv.
Improvements to batch screen, now autopicks batch to open if there's only one available.
Improved Formatting function.
Fixed display issues for users with lots of accounts.
Error handling improvements.
Other minor tidy ups.

1.13.0+ to do list
Look at adding SinglePlayer autobackup feature
Add Capability for D2Emu Websocket connection as the current TZ/DClone API might be getting deprecated. 
In line with the above, if possible investigate the possibility of realtime DClone Alarms.
In line with the above, perhaps investigate putting TZ details on main menu and using the TZ screen for recent TZ's only.
To reduce lines, Tidy up all the import/export csv bits for stat updates into a function rather than copy paste the same commands throughout the script. Can't really be bothered though :)
Unlikely - ISboxer has CTRL + Alt + number as a shortcut to switch between windows. Investigate how this could be done. Would need an agent to detect key combos, Possibly via AutoIT or Autohotkey. Likely not possible within powershell and requires a separate project.
Fix whatever I broke or poorly implemented in the last update :)
#>

param($AccountUsername,$PW,$Region,$All,$Batch,$ManualSettingSwitcher) #used to capture parameters sent to the script, if anyone even wants to do that.
$CurrentVersion = "1.13.0"
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
if ($Null -ne $ScriptArguments){
	$Script:ParamsUsed = $true
}
Else {
	$Script:ParamsUsed = $false
}
#run script as admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){ Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $ScriptArguments"  -Verb RunAs;exit }
#DebugMode
#$DebugMode = $True # Uncomment to enable
if ($DebugMode -eq $True){
	$DebugPreference = "Continue"
	$VerbosePreference = "Continue"
}
#set window size
[console]::WindowWidth=77; #script has been designed around this width. Adjust at your own peril.
[console]::WindowHeight=50; #Can be adjusted to preference, but not less than 42
[console]::BufferWidth=[console]::WindowWidth
#set misc vars
$Script:X = [char]0x1b #escape character for ANSI text colors
$ProgressPreference = "SilentlyContinue"
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\')) #Set Current Directory path.
$Script:StartTime = Get-Date #Used for elapsed time. Is reset when script refreshes.
$Script:MOO = "%%%"
$Script:JobIDs = @()
$MenuRefreshRate = 30 #How often the script refreshes in seconds. This should be set to 30, don't change this please.
$Script:ScriptFileName = Split-Path $MyInvocation.MyCommand.Path -Leaf #find the filename of the script in case a user renames it.
$Script:SessionTimer = 0 #set initial session timer to avoid errors in info menu.
$Script:NotificationHasBeenChecked = $False
#Baseline of acceptable characters for ReadKey functions. Used to prevents receiving inputs from folk who are alt tabbing etc.
$Script:AllowedKeyList = @(48,49,50,51,52,53,54,55,56,57) #0 to 9
$Script:AllowedKeyList += @(96,97,98,99,100,101,102,103,104,105) #0 to 9 on numpad
$Script:AllowedKeyList += @(65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90) # A to Z
$Script:MenuOptions = @(65,66,67,68,71,73,74,79,82,83,84,88) #a, b, c, d, g, i, j, o, r, s, t and x. Used to detect singular valid entries where script can have two characters entered.
$EnterKey = 13
Function ReadKey([string]$message=$Null,[bool]$NoOutput,[bool]$AllowAllKeys){#used to receive user input
	$key = $Null
	$Host.UI.RawUI.FlushInputBuffer()
	if (![string]::IsNullOrEmpty($message)){
		Write-Host -NoNewLine $message
	}
	$AllowedKeyList = $Script:AllowedKeyList + @(13,27) #Add Enter & Escape to the allowedkeylist as acceptable inputs.
	while ($Null -eq $key){
	if ($Host.UI.RawUI.KeyAvailable){
			$key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
			if ($True -ne $AllowAllKeys){
				if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList){
					$key = $key_
				}
			}
			else {
				if ($key_.KeyDown){
					$key = $key_
				}
			}
		}
		else {
			Start-Sleep -m 200  # Milliseconds
		}
	}
	if ($key_.VirtualKeyCode -ne $EnterKey -and -not ($Null -eq $key) -and [bool]$NoOutput -ne $true){
		Write-Host ("$X[38;2;255;165;000;22m" + "$($key.Character)" + "$X[0m") -NoNewLine
	}
	if (![string]::IsNullOrEmpty($message)){
		Write-Host "" # newline
	}
	return $(
		if ($Null -eq $key -or $key.VirtualKeyCode -eq $EnterKey){
			""
		}
		ElseIf ($key.VirtualKeyCode -eq 27){ #if key pressed was escape
			"Esc"
		}
		else {
			$key.Character
		}
	)
}
Function ReadKeyTimeout([string]$message=$Null, [int]$timeOutSeconds=0, [string]$Default=$Null, [object[]]$AdditionalAllowedKeys = $null, [bool]$TwoDigitAcctSelection = $False){
	$key = $Null
	$inputString = ""
	$Host.UI.RawUI.FlushInputBuffer()
	if (![string]::IsNullOrEmpty($message)){
		Write-Host -NoNewLine $message
	}
	$Counter = $timeOutSeconds * 1000 / 250
	$AllowedKeyList = $Script:AllowedKeyList + $AdditionalAllowedKeys #Add any other specified allowed key inputs (eg Enter).
	while ($Null -eq $key -and ($timeOutSeconds -eq 0 -or $Counter-- -gt 0)){
		if ($TwoDigitAcctSelection -eq $True -and $inputString.length -ge 1){
			$AllowedKeyList = $AllowedKeyList + 13 + 8 # Allow enter and backspace to be used if 1 character has been typed.
		}
		if (($timeOutSeconds -eq 0) -or $Host.UI.RawUI.KeyAvailable){
			$key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
			if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList){
				if ($key_.VirtualKeyCode -eq [System.ConsoleKey]::Backspace){
					$Counter = $timeOutSeconds * 1000 / 250 #reset counter
					if ($inputString.Length -gt 0){
						$inputString = $inputString.Substring(0, $inputString.Length - 1) #remove last added character/number from variable
						# Clear the last character from the console
						$Host.UI.RawUI.CursorPosition = @{
							X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - 1, 0)
							Y = $Host.UI.RawUI.CursorPosition.Y
						}
						Write-Host -NoNewLine " " #-ForegroundColor Black
						$Host.UI.RawUI.CursorPosition = @{
							X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - 1, 0)
							Y = $Host.UI.RawUI.CursorPosition.Y
						}
					}
				}
				ElseIf ($TwoDigitAcctSelection -eq $True -and $key_.VirtualKeyCode -notin $Script:MenuOptions + 27){
					$Counter = $timeOutSeconds * 1000 / 250 #reset counter
					if ($key_.VirtualKeyCode -eq $EnterKey -or $key_.VirtualKeyCode -eq 27){
						break
					}
					$inputString += $key_.Character
					Write-Host ("$X[38;2;255;165;000;22m" + $key_.Character + "$X[0m") -nonewline
					if ($inputString.length -eq 2){#if 2 characters have been entered
						break
					}
				}
				Else {
					$key = $key_
					$inputString = $key_.Character
				}
			}
		}
		else {
			Start-Sleep -m 250 # Milliseconds
		}
	}
	if ($Counter -le 0){
		if ($InputString.Length -gt 0){# if it timed out, revert to no input if one character was entered.
			$InputString = "" #remove last added character/number from variable
		}
	}
	if ($TwoDigitAcctSelection -eq $False -or ($TwoDigitAcctSelection -eq $True -and $key_.VirtualKeyCode -in $Script:MenuOptions)){
		Write-Host ("$X[38;2;255;165;000;22m" + "$inputString" + "$X[0m")
	}
	if (![string]::IsNullOrEmpty($message) -or $TwoDigitAcctSelection -eq $True){
		Write-Host "" # newline
	}
	Write-Host #prevent follow up text from ending up on the same line.
	return $(
		If ($key.VirtualKeyCode -eq $EnterKey -and $EnterKey -in $AllowedKeyList){
			""
		}
		ElseIf ($key.VirtualKeyCode -eq 27){ #if key pressed was escape
			"Esc"
		}
		ElseIf ($inputString.Length -eq 0){
			$Default
		}
		else {
			$inputString
		}
	)
}
Function PressTheAnyKey {#Used instead of Pause so folk can hit any key to continue
	Write-Host "  Press any key to continue..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	Write-Host
}
Function PressTheAnyKeyToExit {#Used instead of Pause so folk can hit any key to exit
	Write-Host "  Press Any key to exit..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	remove-job * -force
	Exit
}
Function Red {
	process { Write-Host $_ -ForegroundColor Red }
}
Function Yellow {
	process { Write-Host $_ -ForegroundColor Yellow }
}
Function Green {
	process { Write-Host $_ -ForegroundColor Green }
}
Function NormalText {
	process { Write-Host $_ }
}
Function FormatFunction { # Used to get long lines formatted nicely within the CLI. Possibly the most difficult thing I've created in this script. Hooray for Regex!
	param (
		[string] $Text,
		[int] $Indents,
		[int] $SubsequentLineIndents,
		[switch] $IsError,
		[switch] $IsWarning,
		[switch] $IsSuccess
	)
	if ($IsError -eq $True){
		$Colour = "Red"
	}
	ElseIf ($IsWarning -eq $True){
		$Colour = "Yellow"
	}
	ElseIf ($IsSuccess -eq $True){
		$Colour = "Green"
	}
	Else {
		$Colour = "NormalText"
	}
	$MaxLineLength = 76
	If ($Indents -ge 1){
		while ($Indents -gt 0){
			$Indent += " "
			$Indents --
		}
	}
	If ($SubsequentLineIndents -ge 1){
		while ($SubsequentLineIndents -gt 0){
			$SubsequentLineIndent += " "
			$SubsequentLineIndents --
		}
	}
	$Text -split "`n" | ForEach-Object {
		$Line = " " + $Indent + $_	
		$SecondLineDeltaIndent = ""
		if ($Line -match '^[\s]*-'){ #For any line starting with any preceding spaces and a dash.
			$SecondLineDeltaIndent = "  "
		}
		if ($Line -match '^[\s]*\d+\.\s'){ #For any line starting with any preceding spaces, a number, a '.' and a space. Eg "1. blah".
			$SecondLineDeltaIndent = "   "
		}
		Function Formatter ([string]$line){
			$pattern = "[\e]?[\[]?[`"-,`.!']?\b[\w\-,'`"]+(\S*)" # Regular expression pattern to find the last word including any trailing non-space characters. Also looks to include any preceding special characters or ANSI escape character.
			$LastWordMatches = [regex]::Matches($Line, $pattern) # Find all matches of the pattern in the string
			# Initialize variables to track the match with the highest index
			$highestIndex = -1
			$SelectedMatch = $Null
			$PatternLengthCount = 0
			$ANSIPatterns = "\x1b\[38;\d{1,3};\d{1,3};\d{1,3};\d{1,3};\d{1,3}m","\x1b\[0m","\x1b\[4m"
			ForEach ($match in $LastWordMatches){# Iterate through each match (match being a block of characters, ie each word).
				ForEach ($ANSIPattern in $ANSIPatterns){ #iterate through each possible ANSI pattern to find any text that might have ANSI formatting.
					$ANSIMatches = $match.value | Select-String -Pattern $ANSIPattern -AllMatches
					ForEach ($ANSIMatch in $ANSIMatches){
						$Script:ANSIUsed = $True
						$PatternLengthCount = $PatternLengthCount + (($ANSIMatch.matches | ForEach-Object {$_.Value}) -join "").length #Calculate how many characters in the text are ANSI formatting characters and thus won't be displayed on screen, to prevent skewing word count.
					}
				}
				$matchIndex = $match.Index
				$matchLength = $match.Length
				$matchEndIndex = $matchIndex + $matchLength - 1
				if ($matchEndIndex -lt ($MaxLineLength + $PatternLengthCount)){# Check if the match ends within the first $MaxLineLength characters
					if ($matchIndex -gt $highestIndex){# Check if this match has a higher index than the current highest
						$highestIndex = $matchIndex # This word has a higher index and is the winner thus far.
						$SelectedMatch = $Match
						$lastspaceindex = $SelectedMatch.Index + $SelectedMatch.Length - 1 #Find the index (the place in the string) where the last word can be used without overflowing the screen.
					}
				}
			}
			try {
				$script:chunk = $Line.Substring(0, $lastSpaceIndex + 1) #Chunk of text to print to screen. Uses all words from the start of $line up until $lastspaceindex so that only text that fits on a single line is printed. Prevents words being cut in half and prevents loss of indenting.
			}
			catch {
				$script:chunk = $Line.Substring(0, [Math]::Min(($MaxLineLength), ($Line.Length))) #If the above fails for whatever reason. Can't exactly remember why I put this in here but leaving it in to be safe LOL.
			}	
		}
		Formatter $Line
		if ($Script:ANSIUsed -eq $True){ #if fancy pants coloured text (ANSI) is used, write out the first line. Check if ANSI was used in any overflow lines.
			do {
				$Script:ANSIUsed = $False
				Write-Output $Chunk | out-host #have to use out-host due to pipeline shenanigans and at this point was too lazy to do things properly :)
				$Line = " " + $SubsequentLineIndent + $Indent + $Line.Substring($chunk.Length).trimstart() #$Line is equal to $Line but without the text that's already been outputted.
				Formatter $Line
			} until ($Script:ANSIUsed -eq $False)
			if ($Chunk -ne " " -and $Chunk.lenth -ne 0){#print any remaining text.
				Write-Output $Chunk | out-host
			}
		}
		Else { #if line has no ANSI formatting.
			Write-Output $Chunk | &$Colour
		}
		$Line = $Line.Substring($chunk.Length).trimstart() #remove the string that's been printed on screen from variable.
		if ($Line.length -gt 0){ # I see you're reading my comment. How thorough of you! This whole function was an absolute mindf#$! to come up with and took probably 30 hours of trial, error and rage (in ascending order of frequency). Odd how the most boring of functions can take up the most time :)
				Write-Output ($Line -replace "(.{1,$($MaxLineLength - $($Indent.length) - $($SubsequentLineIndent.length) -1 - $($SecondLineDeltaIndent.length))})(\s+|$)", " $SubsequentLineIndent$SecondLineDeltaIndent$Indent`$1`n").trimend() | &$Colour
		}
	}
}
Function CommaSeparatedList {
	param (
		[object] $Values,
		[switch] $NoOr,
		[switch] $AndText
	)
	ForEach ($Value in $Values){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
		if ($Value -ne $Values[-1]){
			Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
			if ($Value -ne $Values[-2]){Write-Host ", " -nonewline}
		}
		else {
			if ($Values.count -gt 1){
				$AndOr = "or"
				if ($AndText -eq $True){
					$AndOr = "and"
				}
				if ($NoOr -eq $False){
					Write-Host " $AndOr " -nonewline
				}
				Else {
					Write-Host ", " -nonewline
				}
			}
			Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
		}
	}
}
Function DisplayPreviousAccountOpened {
		Write-Host "Account previously opened was:" -foregroundcolor yellow -backgroundcolor darkgreen
		$Lastopened = @(
			[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$Script:LastRegion}
		)
		Write-Host " " -NoNewLine
		Write-Host ("Account:  " + $Lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen
		Write-Host " " -NoNewLine
		Write-Host "Region:  " $Lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
}
Function InitialiseCurrentStats {
	if ((Test-Path -Path "$Script:WorkingDirectory\Stats.csv") -ne $true){#Create Stats CSV if it doesn't exist
		$Null = {} | Select-Object "TotalGameTime","TimesLaunched","LastUpdateCheck","HighRunesFound","UniquesFound","SetItemsFound","RaresFound","MagicItemsFound","NormalItemsFound","Gems","CowKingKilled","PerfectGems" | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
		Write-Host " Stats.csv created!"
	}
	do {
		Try {
			$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv" #Get current stats csv details
		}
		Catch {
			Write-Host " Unable to import stats.csv. File corrupt or missing." -foregroundcolor red
		}
		if ($null -ne $CurrentStats){
			#Todo: In the Future add CSV validation checks
			$StatsCSVImportSuccess = $True
		}
		else {#Error out and exit if there's a problem with the csv.
				if ($StatsCSVRecoveryAttempt -lt 1){
					try {
						Write-Host " Attempting Autorecovery of stats.csv from backup..." -foregroundcolor red
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
					Write-Host "`n Stats.csv is corrupted or empty." -foregroundcolor red
					Write-Host " Replace with data from stats.backup.csv or delete stats.csv`n" -foregroundcolor red
					PressTheAnyKeyToExit
				}
			}
	} until ($StatsCSVImportSuccess -eq $True)
	if (-not ($CurrentStats | Get-Member -Name "LastUpdateCheck" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.8.1+. If LastUpdateCheck column doesn't exist, add it to the CSV data
		$Script:CurrentStats | ForEach-Object {
			$_ | Add-Member -NotePropertyName "LastUpdateCheck" -NotePropertyValue "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
		}
	}
	ElseIf ($CurrentStats.LastUpdateCheck -eq "" -or $CurrentStats.LastUpdateCheck -like "*/*"){# If script has just been freshly downloaded or has the old Date format.
		$Script:CurrentStats.LastUpdateCheck = "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
		$CurrentStats | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
	}
}
Function CheckForUpdates {
	#Only Check for updates if updates haven't been checked in last 8 hours. Reduces API requests.
	if ($Script:CurrentStats.LastUpdateCheck -lt (Get-Date).addHours(-8).ToString('yyyy.MM.dd HH:mm:ss')){# Compare current date and time to LastUpdateCheck date & time.
		try {
			# Check for Updates
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
			if ($Script:LatestVersion -gt $Script:CurrentVersion){ #If a newer version exists, prompt user about update details and ask if they want to update.
				Write-Host "`n Update available! See Github for latest version and info" -foregroundcolor Yellow -nonewline
				if ([version]$CurrentVersion -in (($Releases.name.Trim('v') | ForEach-Object { [version]$_ } | Sort-Object -desc)[2..$releases.count])){
					Write-Host ".`n There have been several releases since your version." -foregroundcolor Yellow
					Write-Host " Checkout Github releases for fixes/features added. " -foregroundcolor Yellow
					Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/$X[0m`n"
				}
				Else {
					Write-Host ":`n $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m`n"
				}
				FormatFunction -Text $ReleaseInfo.body #Output the latest release notes in an easy to read format.
				Write-Host; Write-Host
				Do {
					Write-Host " Your Current Version is v$CurrentVersion."
					Write-Host (" Would you like to update to v"+ $Script:LatestVersion + "? $X[38;2;255;165;000;22mY$X[0m/$X[38;2;255;165;000;22mN$X[0m: ") -nonewline
					$ShouldUpdate = ReadKey
					if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes" -or $ShouldUpdate -eq "n" -or $ShouldUpdate -eq "no"){
						$UpdateResponseValid = $True
					}
					Else {
						Write-Host "`n Invalid response. Choose $X[38;2;255;165;000;22mY$X[0m $X[38;2;231;072;086;22mor$X[0m $X[38;2;255;165;000;22mN$X[0m.`n" -ForegroundColor red
					}
				} Until ($UpdateResponseValid -eq $True)
				if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes"){#if user wants to update script, download .zip of latest release, extract to temporary folder and replace old D2Loader.ps1 with new D2Loader.ps1
					Write-Host "`n Updating... :)" -foregroundcolor green
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
					if ($Null -ne $releaseinfo.assets.browser_download_url){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.
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
			$Script:CurrentStats.LastUpdateCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
			$Script:LatestVersionCheck = $CurrentStats.LastUpdateCheck
			$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update stats.csv with the new time played.
		}
		Catch {
			Write-Host "`n Couldn't check for updates. GitHub API limit may have been reached..." -foregroundcolor Yellow
			Start-Sleep -milliseconds 3500
		}
	}
	#Update (or replace missing) SetTextV2.bas file. This is an newer version of SetText (built by me and ChatGPT) that allows windows to be closed by process ID.
	if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.bas')) -ne $True){#if SetTextv2.bas doesn't exist, download it.
			try {
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
			}
			Catch {#if folder already exists for whatever reason.
				Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
			}
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$ZipURL = $ReleaseInfo.zipball_url #get zip download URL
			$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
			Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
			if ($Null -ne $releaseinfo.assets.browser_download_url){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.
				Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
			}
			$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
			Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
			$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
			Copy-Item -Path ($FolderPath + "\SetText\SetTextv2.bas") -Destination ($Script:WorkingDirectory + "\SetText\SetTextv2.bas")
			Write-Host "  SetTextV2.bas was missing and was downloaded."
			Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
	}
}
Function ImportXML { #Import Config XML
	try {
		$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig
		Write-Verbose "Config imported successfully."
	}
	Catch {
		Write-Host "`n Config.xml Was not able to be imported. This could be due to a typo or a special character such as `'&`' being incorrectly used." -foregroundcolor red
		Write-Host " The error message below will show which line in the config.xml is invalid:" -foregroundcolor red
		Write-Host (" " + $PSitem.exception.message + "`n") -foregroundcolor red
		PressTheAnyKeyToExit
	}
}
Function SetDCloneAlarmLevels {
	if ($Script:Config.DCloneAlarmLevel -eq "All"){
		$Script:DCloneAlarmLevel = "1,2,3,4,5,6"
	}
	ElseIf ($Script:Config.DCloneAlarmLevel -eq "Close"){
		$Script:DCloneAlarmLevel = "1,4,5,6"
	}
	ElseIf ($Script:Config.DCloneAlarmLevel -eq "Imminent"){
		$Script:DCloneAlarmLevel = "1,5,6"
	}
	else {#if user has typo'd the config file or left it blank.
		$DCloneErrorMessage = ("  Error: DClone Alarm Levels have been misconfigured in config.xml. ###  Check that the value for DCloneAlarmLevel is entered correctly.").Replace("###", "`n")
		Write-Host ("`n" + $DCloneErrorMessage + "`n") -Foregroundcolor red
		PressTheAnyKeyToExit
	}
}
Function ValidationAndSetup {
	#Perform some validation on config.xml. Helps avoid errors for people who may be on older versions of the script and are updating. Will look to remove all of this in a future update.
	if (Select-String -path $Script:WorkingDirectory\Config.xml -pattern "multiple game installs"){#Sort out an incorrect description text that will have been in folks config.xml for some time. This description was never valid and was from when the setting switcher feature was being developed and tested.
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = ";;`t`tNote, if using multiple game installs \(to keep client specific config persistent for each account\), ensure these are referenced in the CustomGamePath field in accounts.csv."
		$Pattern += ";;`t`tOtherwise you can use a single install instead by linking the path below.-->;;"
		$NewXML = [string]::join(";;",($XML.Split("`r`n")))
		$NewXML = $NewXML -replace $Pattern, "-->;;"
		$NewXML = $NewXML -replace ";;","`r`n"
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Write-Host "`n Corrected the description for GamePath in config.xml." -foregroundcolor Green
		Start-Sleep -milliseconds 1500
	}
	if ($Null -ne $Script:Config.CommandLineArguments){#remove this config option as arguments are now stored in accounts.csv so that different arguments can be set for each account
		Write-Host "`n Config option 'CommandLineArguments' is being moved to accounts.csv" -foregroundcolor Yellow
		Write-Host " This is to enable different CMD arguments per account." -foregroundcolor Yellow
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = ";;\t<!--Optionally add any command line arguments that you'd like the game to start with-->;;\t<CommandLineArguments>.*?</CommandLineArguments>;;"
		$NewXML = [string]::join(";;",($XML.Split("`r`n")))
		$NewXML = $NewXML -replace $Pattern, ""
		$NewXML = $NewXML -replace ";;","`r`n"
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		$Script:OriginalCommandLineArguments = $Script:Config.CommandLineArguments
		Write-Host " CommandLineArguments has been removed from config.xml" -foregroundcolor green
		Start-Sleep -milliseconds 1500
	}
	if ($Null -eq $Script:Config.ManualSettingSwitcherEnabled){#not to be confused with the AutoSettingSwitcher.
		Write-Host "`n Config option 'ManualSettingSwitcherEnabled' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is an optional config option to allow you to manually select which " -foregroundcolor Yellow
		Write-Host " config file you want to use for each account when launching." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
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
	if ($Null -eq $Script:Config.TrackAccountUseTime){
		Write-Host "`n Config option 'TrackAccountUseTime' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is an optional config option to track time played per account." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
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
	if ($Null -eq $Script:Config.ConvertPlainTextSecrets){
		Write-Host "`n Renaming ConvertPlainTextPasswords to ConvertPlainTextSecrets in config.xml" -foregroundcolor Yellow
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "<!--Whether script should convert plain text passwords in accounts.csv to a secure string, recommend leaving this set to True.-->"
		$Replacement = "<!--Whether script should convert plain text tokens and passwords in accounts.csv to a secure string. Recommend leaving this set to True.-->"
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$Pattern = "ConvertPlainTextPasswords"
		$Replacement = "ConvertPlainTextSecrets"
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Write-Host " Config.xml file updated :)`n" -foregroundcolor green
		Start-Sleep -milliseconds 1500
		ImportXML
		PressTheAnyKey
	}
	if ($Null -ne $Script:Config.EnableBatchFeature){ # remove from config.xml. Not needed anymore as script checks accounts.csv to see if batches are used.
		Write-Host "`n Config option 'EnableBatchFeature' is no longer needed." -foregroundcolor Yellow
		Write-Host " Removed this option from config.xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = ";;\t<!--Enable the ability to open a group of accounts.*?</EnableBatchFeature>;;"
		$NewXML = [string]::join(";;",($XML.Split("`r`n")))
		$NewXML = $NewXML -replace $Pattern, ""
		$NewXML = $NewXML -replace ";;","`r`n"
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -ne $Script:Config.CheckForNextTZ){
		Write-Host "`n Config option 'CheckForNextTZ' is no longer needed." -foregroundcolor Yellow
		Write-Host " Removed this option from config.xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = ";;\t<!--Choose whether or not TZ checker.*?</CheckForNextTZ>;;"
		$NewXML = [string]::join(";;",($XML.Split("`r`n")))
		$NewXML = $NewXML -replace $Pattern, ""
		$NewXML = $NewXML -replace ";;","`r`n"
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -ne $Script:Config.AskForRegionOnceOnly){
		Write-Host "`n Config option 'AskForRegionOnceOnly' is no longer needed." -foregroundcolor Yellow
		Write-Host " Removed this option from config.xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = ";;\t<!--Whether script should only prompt you once for region.*?</AskForRegionOnceOnly>;;"
		$NewXML = [string]::join(";;",($XML.Split("`r`n")))
		$NewXML = $NewXML -replace $Pattern, ""
		$NewXML = $NewXML -replace ";;","`r`n"
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -eq $Script:Config.DisableOpenAllAccountsOption){
		Write-Host "`n Config option 'DisableOpenAllAccountsOption' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is an optional config option to disable the functionality for opening all accounts." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</DefaultRegion>"
		$Replacement = "</DefaultRegion>`n`n`t<!--Disable the functionality of being able to open all accounts at once. This is for any crazy people who have a lot of accounts and want to prevent accidentally opening all at once.-->`n`t"
		$Replacement += "<DisableOpenAllAccountsOption>False</DisableOpenAllAccountsOption>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -eq $Script:Config.RememberWindowLocations){
		Write-Host "`n Config option 'RememberWindowLocations' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is an optional config option to remember game window locations/sizes." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</ConvertPlainTextSecrets>"
		$Replacement = "</ConvertPlainTextSecrets>`n`n`t<!--Make game launch each instance in the same screen location so you don't have to move your game windows around when starting the game.-->`n`t"
		$Replacement += "<RememberWindowLocations>False</RememberWindowLocations>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Script:Config.RememberWindowLocations -eq $True){
		if ((Test-Path -Path ($workingdirectory + 'WindowMover.ps1')) -ne $True){
			$WindowMoverUrl = "https://raw.githubusercontent.com/shupershuff/Diablo2RLoader/main/WindowMover.ps1"
			Invoke-WebRequest -Uri $WindowMoverUrl -OutFile "$WorkingDirectory\WindowMover.ps1"
		}
	}
	if ($Null -eq $Script:Config.DCloneTrackerSource){
		Write-Host "`n Config option 'DCloneTrackerSource' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is a required config option to determine which source should be used" -foregroundcolor Yellow
		Write-Host " for obtaining current DClone status." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</TrackAccountUseTime>"
		$Replacement = "</TrackAccountUseTime>`n`n`t<!--Options are d2emu.com, D2runewizard.com and diablo2.io.`n`t"
		$Replacement += "Default and recommended option is d2emu.com as this pulls live data from the game as opposed to crowdsourced data.-->`n`t"
		$Replacement += "<DCloneTrackerSource>d2emu.com</DCloneTrackerSource>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		ImportXML
		PressTheAnyKey
	}
	if ($Null -eq $Script:Config.DCloneAlarmLevel){
		Write-Host "`n Config option 'DCloneAlarmLevel' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This field determines if alarms should activate for all DClone status " -foregroundcolor Yellow
		Write-Host " changes or just when DClone is about to walk." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
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
	if ($Null -eq $Script:Config.DCloneAlarmList){
		Write-Host "`n Config option 'DCloneAlarmList' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This is an optional config option to enable both audible and text based" -foregroundcolor Yellow
		Write-Host " alarms for DClone Status changes." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</DCloneTrackerSource>"
		$Replacement = "</DCloneTrackerSource>`n`n`t<!--Allow you to have the script audibly warn you of upcoming dclone walks.`n`t"
		$Replacement +=	"Specify as many of the following options as you like: SCL-NA, SCL-EU, SCL-KR, SC-NA, SC-EU, SC-KR, HCL-NA, HCL-EU, HCL-KR, HC-NA, HC-EU, HC-KR`n`t"
		$Replacement +=	"EG if you want to be notified for all Softcore ladder walks on all regions, enter <DCloneAlarmList>SCL-NA, SCL-EU, SCL-KR</DCloneAlarmList>`n`t"
		$Replacement +=	"If left blank, this feature is disabled. Default is blank as this may be annoying for some people.-->`n`t"
		$Replacement +=	"<DCloneAlarmList></DCloneAlarmList>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		ImportXML
		PressTheAnyKey
	}
	if ($Null -ne $Script:Config.DCloneAlarmList -and $Script:Config.DCloneAlarmList -ne ""){#validate data to prevent errors from typos
		$pattern = "^(HC|SC)(L?)-(NA|EU|KR)$" #set pattern: must start with HC or SC, optionally has L after it, must end in -NA -EU or -KR
		ForEach ($Alarm in $Script:Config.DCloneAlarmList.split(",").trim()){
			if ($Alarm -notmatch $pattern){
				Write-Host "`n $Alarm is not a valid Alarm entry."  -foregroundcolor Red
				Write-Host " See valid options in Config.xml`n"  -foregroundcolor Red
				PressTheAnyKeyToExit
			}
		}
	}
	if ($Null -eq $Script:Config.DCloneAlarmVoice){
		Write-Host "`n Config option 'DCloneAlarmVoice' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This config allows you to choose between a Woman or Man's robot voice." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</DCloneAlarmLevel>"
		$Replacement = "</DCloneAlarmLevel>`n`n`t<!--Specify what voice you want. Choose 'Paladin' for David (Man) or 'Amazon' for Zira (Woman).-->`n`t"
		$Replacement +=	"<DCloneAlarmVoice>Paladin</DCloneAlarmVoice>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -eq $Script:Config.DCloneAlarmVolume){
		Write-Host "`n Config option 'DCloneAlarmVolume' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</DCloneAlarmVoice>"
		$Replacement = "</DCloneAlarmVoice>`n`t<!--Specify how loud notifications can be. Range from 1 to 100.-->`n`t"
		$Replacement +=	"<DCloneAlarmVolume>69</DCloneAlarmVolume>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	if ($Null -eq $Script:Config.ForceAuthTokenForRegion){
		Write-Host "`n Config option 'ForceAuthTokenForRegion' missing from config.xml" -foregroundcolor Yellow
		Write-Host " This is due to the config.xml recently being updated." -foregroundcolor Yellow
		Write-Host " This config allows you to force AuthToken based authentication for any." -foregroundcolor Yellow
		Write-Host " regions specified in config. Useful for when Blizzard have bricked." -foregroundcolor Yellow
		Write-Host " their authentication servers on a particular region." -foregroundcolor Yellow
		Write-Host " Added this missing option into .xml file :)`n" -foregroundcolor green
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
		$Pattern = "</DCloneAlarmVoice>"
		$Replacement = "</DCloneAlarmVoice>`n`n`t<!--Select regions which should be forced to use Tokens over parameters (overrides config in accounts.csv).`n`t"
		$Replacement += "Only use this if connecting via Parameters is down for a particular region and you don't want to have to manually toggle.`n`t"
		$Replacement +=	"Valid options are NA, EU and KR. Default is blank.-->`n`t"
		$Replacement +=	"<ForceAuthTokenForRegion></ForceAuthTokenForRegion>" #add option to config file if it doesn't exist.
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
	}
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml" -raw
	$Pattern = "d2rapi.fly.dev"
	if ($Null -ne (Select-Xml -Content $xml -XPath "//*[contains(.,'$pattern')]")){
		Write-Host "`n Replaced 'd2rapi.fly.dev' in config with 'd2emu.com'." -foregroundcolor Yellow
		$Replacement = "d2emu.com"
		$NewXML = $XML -replace [regex]::Escape($Pattern), $Replacement
		$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
		Start-Sleep -milliseconds 1500
		PressTheAnyKey
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
			Write-Host "`n Error: DCloneAlarmVoice in config has an invalid option set." -Foregroundcolor red
			Write-Host " Open config.xml and set this to either Paladin or Amazon.`n" -Foregroundcolor red
			PressTheAnyKeyToExit
		}
		SetDCloneAlarmLevels
	}
	$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig #import config.xml again for any updates made by the above.
	#check if there's any missing config.xml options, if so user has out of date config file.
	$AvailableConfigs = #add to this if adding features.
	"GamePath",
	"DefaultRegion",
	"ShortcutCustomIconPath"
	$BooleanConfigs =
	"ConvertPlainTextSecrets",
	"ManualSettingSwitcherEnabled",
	"RememberWindowLocations",
	"DisableOpenAllAccountsOption",
	"CreateDesktopShortcut",
	"ForceWindowedMode",
	"SettingSwitcherEnabled",
	"TrackAccountUseTime"
	$AvailableConfigs = $AvailableConfigs + $BooleanConfigs
	$ConfigXMLlist = ($Config | Get-Member | Where-Object {$_.membertype -eq "Property" -and $_.name -notlike "#comment"}).name
	Write-Host
	ForEach ($Option in $AvailableConfigs){#Config validation
		if ($Option -notin $ConfigXMLlist){
			Write-Host " Config.xml file is missing a config option for $Option." -foregroundcolor yellow
			Start-Sleep 1
			PressTheAnyKey
		}
	}
	if ($Option -notin $ConfigXMLlist){
		Write-Host "`n Make sure to grab the latest version of config.xml from GitHub" -foregroundcolor yellow
		Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m`n"
		PressTheAnyKey
	}
	if ($Config.GamePath -match "`""){#Remove any quotes from path in case someone ballses this up.
		$Script:GamePath = $Config.GamePath.replace("`"","")
	}
	else {
		$Script:GamePath = $Config.GamePath
	}
	ForEach ($ConfigCheck in $BooleanConfigs){#validate all configs that require "True" or "False" as the setting.
		if ($Null -ne $Config.$ConfigCheck -and ($Config.$ConfigCheck -ne $true -and $Config.$ConfigCheck -ne $false)){#if config is invalid
			Write-Host " Config option '$ConfigCheck' is invalid." -foregroundcolor Red
			Write-Host " Ensure this is set to either True or False.`n" -foregroundcolor Red
			PressTheAnyKeyToExit
		}
	}
	if ($Config.ShortcutCustomIconPath -match "`""){#Remove any quotes from path in case someone ballses this up.
		$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath.replace("`"","")
	}
	else {
		$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath
	}
	#Check Windows Game Path for D2r.exe is accurate.
	if ((Test-Path -Path "$GamePath\d2r.exe") -ne $True){
		Write-Host " Gamepath is incorrect. Looks like you have a custom D2r install location!" -foregroundcolor red
		Write-Host " Edit the GamePath variable in the config file.`n" -foregroundcolor red
		PressTheAnyKeyToExit
	}
	# Create Shortcut
	if ($Config.CreateDesktopShortcut -eq $True){
		$DesktopPath = [Environment]::GetFolderPath("Desktop")
		$Targetfile = "-ExecutionPolicy Bypass -File `"$WorkingDirectory\$ScriptFileName`""
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
		Write-Host "`n First Time run!`n" -foregroundcolor Yellow
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
		try {
			Write-Host "`n Handle64.exe not in .\Handle\ folder. Downloading now..." -foregroundcolor Yellow
			try {
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
			}
			Catch {#if folder already exists for whatever reason.
				Remove-Item -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -Recurse -Force
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") | Out-Null #create temporary folder to download zip to and extract
			}
			$ZipURL = "https://download.sysinternals.com/files/Handle.zip" #get zip download URL
			$ZipPath = ($WorkingDirectory + "\Handle\ExtractTemp\")
			Invoke-WebRequest -Uri $ZipURL -OutFile ($ZipPath + "\Handle.zip")
			Expand-Archive -Path ($ZipPath + "\Handle.zip") -DestinationPath $ZipPath -Force
			Copy-Item -Path ($ZipPath + "Handle64.exe") -Destination ($Script:WorkingDirectory + "\Handle\")
			Remove-Item -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -Recurse -Force #delete update temporary folder
			Write-Host " Successfully downloaded Handle64.exe :)" -ForeGroundcolor Green
			Start-Sleep -milliseconds 2024	
		}
		Catch {
			Write-Host " Handle.zip couldn't be downloaded." -foregroundcolor red
			FormatFunction -text "It's possible the download link changed. Try checking the Microsoft page or SysInternals.com site for a download link and ensure that handle64.exe is placed in the .\Handle\ folder." -IsError
			Write-Host "`n $X[38;2;69;155;245;4mhttps://learn.microsoft.com/sysinternals/downloads/handle$X[0m"
			Write-Host " $X[38;2;69;155;245;4mhttps://download.sysinternals.com/files/Handle.zip$X[0m`n"
			PressTheAnyKeyToExit
		}
	}
	#Set Region Array
	$Script:ServerOptions = @(
		[pscustomobject]@{Option='1';region='NA';region_server='us.actual.battle.net'} #Americas
		[pscustomobject]@{Option='2';region='EU';region_server='eu.actual.battle.net'} #Europe
		[pscustomobject]@{Option='3';region='Asia';region_server='kr.actual.battle.net'} #Asia
	)
}
Function ValidateTokenInput {
	param (
		[bool] $ManuallyEntered,
		[string] $TokenInput,
		[string] $AccountLabel
	)
	do {
		$extractedInfo = $null
		if ($ManuallyEntered){
			$TokenInput = Read-host (" Enter your token URL or enter the token for account '" + $AccountLabel + "'")			
		}
		$pattern = "(?<=\?ST=|&ST=|^|http://localhost:0/\?ST=)([^&]+)"
		if ($tokeninput -match $pattern){
			$extractedInfo = ($matches[1]).replace("http://localhost:0/?ST=","")
			return $extractedInfo
		}
		Else {
			Write-Host " Token details are incorrect." -foregroundcolor red
			if (!$ManuallyEntered){
				Write-Host " Please review the setup instructions.`n" -foregroundcolor red
				PressTheAnyKeyToExit
			}
			Else {
				Write-Host " Please enter the full URL." -foregroundcolor red
			}
		}
	}
	until ($null -ne $extractedInfo)
}
Function ImportCSV { #Import Account CSV
	do {
		if ($Null -eq $Script:AccountUsername){#If no parameters sent to script.
			try {
				$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv" #import all accounts from csv
			}
			Catch {
				FormatFunction -text "`nAccounts.csv does not exist. Make sure you create this and populate with accounts first." -IsError
				PressTheAnyKeyToExit
			}
		}
		if ($Null -ne $Script:AccountOptionsCSV){
			#check Accounts.csv has been updated and doesn't contain the example account.
			if ($Script:AccountOptionsCSV -match "yourbnetemailaddress"){
				Write-Host "`n You haven't setup accounts.csv with your accounts." -foregroundcolor red
				Write-Host " Add your account details to the CSV file and run the script again :)`n" -foregroundcolor red
				PressTheAnyKeyToExit
			}
			if ($Null -ne ($AccountOptionsCSV | Where-Object {$_.id -eq ""})){ 
				$Script:AccountOptionsCSV = $Script:AccountOptionsCSV | Where-Object {$_.id -ne ""} # To account for user error, remove any empty lines from accounts.csv
			}
			ForEach ($Account in $AccountOptionsCSV){
				if ($Account.accountlabel -eq ""){ # if user doesn't specify a friendly name, use id. Prevents display issues later on.
					$Account.accountlabel = ("Account " + $Account.id)
					$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
				}
			}
			$DuplicateIDs = $AccountOptionsCSV | Where-Object {$_.id -ne ""} | Group-Object -Property ID | Where-Object { $_.Count -gt 1 }
			if ($duplicateIDs.Count -gt 0){
				$duplicateIDs = ($DuplicateIDs.name | out-string).replace("`r`n",", ").trim(", ") #outputs more meaningful error.
				Write-Host "`n Accounts.csv has duplicate IDs: $duplicateIDs" -foregroundcolor red
				FormatFunction -Text "Please adjust Accounts.csv so that the ID numbers against each account are unique.`n" -IsError
				PressTheAnyKeyToExit
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Batches" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.7.0. If batch column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "Batches" -NotePropertyValue $Null
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
			}
			$BatchesInCSV = ($AccountOptionsCSV | group-object batches | where-object {$_.name -ne ""}).count
			if ($BatchesInCSV -ge 1 -or $Null -ne $Batch){
				$Script:EnableBatchFeature = $True
				$BatchOption = "b" #specified here as well as in the ChooseAccounts section so that this works when being passed as a parameter
			}
			else {
				$BatchOption = "$Null"
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Token" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "AuthenticationMethod" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.10.0. If token columns don't exist, add them. As of 1.12.0 onwards, TokenSecureString is no longer used.
				# Column does not exist, so add it to the CSV data
				if (-not ($Script:AccountOptionsCSV | Get-Member -Name "Token" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
					$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "Token" -NotePropertyValue $Null}
				}
				if (-not ($Script:AccountOptionsCSV | Get-Member -Name "AuthenticationMethod" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
					$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "AuthenticationMethod" -NotePropertyValue "Parameter"}
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
			}
			$OldColumnsToRemove = @("PWIsSecureString","TokenIsSecureString") #Old Account.csv columns that are now unused.
			$ExistingColumns = $Script:AccountOptionsCSV | Select-Object -First 1 | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
			$ColumnsRemoved = $ExistingColumns | Where-Object { $OldColumnsToRemove -contains $_ }
			$DesiredColumnOrder = @("ID","Acct","AccountLabel","Batches","TimeActive","CustomLaunchArguments","AuthenticationMethod","PW","Token") # Update with your desired column order
			if ($ColumnsRemoved.Count -gt 0){
				$Script:AccountOptionsCSV = $Script:AccountOptionsCSV | Select-Object -Property $DesiredColumnOrder
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #rewrite to accounts.csv to remove unused columns.
				FormatFunction -text ("Unused Columns were removed from Accounts.csv: " + ($columnsRemoved -join ", ") + ".`n") -IsWarning
				start-sleep -milliseconds 3000
			}
			else {
				Write-Verbose "No columns were removed as they do not exist in the CSV."
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "CustomLaunchArguments" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.8.0. If CustomLaunchArguments column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "CustomLaunchArguments" -NotePropertyValue $Script:OriginalCommandLineArguments
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
				Write-Host " Added CustomLaunchArguments column to accounts.csv.`n" -foregroundcolor green
				Start-Sleep -milliseconds 1200
				PressTheAnyKey
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "TimeActive" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.8.0. If TimeActive column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "TimeActive" -NotePropertyValue $Null
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
				Write-Host " Added TimeActive column to accounts.csv." -foregroundcolor Green
				PressTheAnyKey
			}
			#Secure any plain text tokens. Ask for tokens on accounts that don't have any if Config is configured to use Token Authentication.
			$NewCSV = ForEach ($Entry in $Script:AccountOptionsCSV){
				if 	($Entry.AuthenticationMethod -eq ""){
					$Entry.AuthenticationMethod = "Parameter"
					$UpdateAccountsCSV = $True
				}
				if ($Entry.AuthenticationMethod -ne "Token" -and $Entry.AuthenticationMethod -ne "Parameter"){
					Write-Host ("`n Error: AuthenticationMethod in accounts.csv for " + $Entry.AccountLabel + " is invalid.") -Foregroundcolor red
					Write-Host " Open accounts.csv and set this to either 'Parameter' or 'Token'.`n" -Foregroundcolor red
					PressTheAnyKeyToExit
				}
				if ($Entry.Token.length -ge 200){#if nothing needs converting, make sure existing entries still make it into the updated CSV
					$Entry
				}
				if ($Entry.Token.length -lt 200 -and $Entry.Token.length -ne 0 -and $Config.ConvertPlainTextSecrets -eq $True){#if accounts.csv has a Token and it's less than 200 chars long, it hasn't been secured yet.
					$ValidatedToken = ValidateTokenInput -TokenInput $Entry.Token
					$Entry.Token = ConvertTo-SecureString -String $ValidatedToken -AsPlainText -Force
					$Entry.Token = $Entry.Token | ConvertFrom-SecureString
					Write-Host (" Secured Token for " + $Entry.AccountLabel) -foregroundcolor green
					Start-Sleep -milliseconds 100
					$Entry
					$TokensUpdated = $true
				}
				if ($Entry.Token.length -eq 0 -and $Entry.AuthenticationMethod -eq "Token"){#if csv has account details but Token field has been left blank
					Write-Host ("`n The account " + $Entry.AccountLabel + " doesn't yet have a Token defined.") -foregroundcolor yellow
					Write-Host " See the readme on Github for how to obtain Auth token.`n" -foregroundcolor yellow
					Write-Host " https://github.com/shupershuff/Diablo2RLoader/#3-setup-your-accounts`n" -foregroundcolor cyan
					while ($Entry.Token.length -eq 0){#prevent empty entries as this will cause errors.
						$Entry.Token = ValidateTokenInput -ManuallyEntered $True -AccountLabel $Entry.AccountLabel
					}
					if ($Config.ConvertPlainTextSecrets -eq $True){
						$Entry.Token = ConvertTo-SecureString -String $Entry.Token -AsPlainText -Force
						$Entry.Token = $Entry.Token | ConvertFrom-SecureString
						Write-Host (" Secured Token for " + $Entry.AccountLabel) -foregroundcolor green
						Start-Sleep -milliseconds 100
					}
					$Entry
					$TokensUpdated = $true
				}
				if ($Entry.Token.length -ge 200 -or ($Config.ConvertPlainTextSecrets -eq $False -and $Entry.Token.length -gt 10)){
					$Script:TokensConfigured = $True
				}
			}
			#Check CSV for Plain text Passwords, convert to encryptedstrings and replace values in CSV
			$NewCSV = ForEach ($Entry in $Script:AccountOptionsCSV){
				if ($Entry.PW.length -ge 300 -and $Config.ConvertPlainTextSecrets -eq $True){#if nothing needs converting, make sure existing entries still make it into the updated CSV
					$Entry
				}
				ElseIf ($Entry.PW.length -ge 1 -and $Config.ConvertPlainTextSecrets -eq $False){#if nothing needs converting, make sure existing entries still make it into the updated CSV
					$Entry
				}
				if ($Entry.PW.length -lt 300 -and $Entry.PW.length -ne 0 -and $Config.ConvertPlainTextSecrets -ne $False){#if accounts.csv has a password and it isn't over 300characters, this means it's not converted yet. As such, convert PW to secure string and update CSV.
					$Entry.PW = ConvertTo-SecureString -String $Entry.PW -AsPlainText -Force
					$Entry.PW = $Entry.PW | ConvertFrom-SecureString
					Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
					Start-Sleep -milliseconds 100
					$Entry
					$PWsUpdated = $true
				}
				if ($Entry.PW.length -eq 0){#if csv has account details but password field has been left blank
					Write-Host ("`n The account " + $Entry.AccountLabel + " doesn't yet have a password defined.`n") -foregroundcolor yellow
					if ($Config.ConvertPlainTextSecrets -eq $True){
						while ($Entry.PW.length -eq 0){#prevent empty entries as this will cause errors.
							$Entry.PW = read-host -AsSecureString " Enter the Battle.net password for"$Entry.AccountLabel
						}
						$Entry.PW = $Entry.PW | ConvertFrom-SecureString
						Write-Host (" Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
					}
					Else { #if passwords aren't to be secured
						while ($Entry.PW.length -eq 0){#prevent empty entries as this will cause errors.
							$Entry.PW = read-host " Enter the Battle.net password for"$Entry.AccountLabel
						}
						Write-Host (" Saved Password for " + $Entry.AccountLabel) -foregroundcolor green
					}
					Start-Sleep -milliseconds 100
					$Entry
					$PWsUpdated = $true
				}
			}
			if ($PWsUpdated -eq $true -or $TokensUpdated -eq $True -or $UpdateAccountsCSV -eq $True){#if CSV needs to be updated
				Try {
					$NewCSV | Export-CSV "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update CSV file
					if ($Config.ConvertPlainTextSecrets -eq $True){$SavedOrSecured = "Secured"}Else{$SavedOrSecured = "Saved"}
					if ($TokensUpdated -eq $True){
						Write-Host " Accounts.csv updated: Tokens have been $SavedOrSecured." -foregroundcolor green
					}
					if ($PWsUpdated -eq $True){
						Write-Host " Accounts.csv updated: Passwords have been $SavedOrSecured." -foregroundcolor green
					}
					Start-Sleep -milliseconds 3000
				}
				Catch {
					Write-Host "`n Couldn't update Accounts.csv, probably because the file is open & locked." -foregroundcolor red
					Write-Host " Please close accounts.csv and run the script again!" -foregroundcolor red
					Write-Host "`n If Accounts.csv isn't open, then there must be a permission issue." -foregroundcolor red
					Write-Host " Try moving your install folder to a different location." -foregroundcolor red
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
				Write-Host "`n There's an issue with accounts.csv." -foregroundcolor red
				Write-Host " Please ensure that this is filled out correctly and rerun the script." -foregroundcolor red
				Write-Host " Alternatively, rebuild CSV from scratch or restore from accounts.backup.csv`n" -foregroundcolor red
				PressTheAnyKeyToExit
			}
		}
	} until ($AccountCSVImportSuccess -eq $True)
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	([int]$Script:CurrentStats.TimesLaunched) ++
	if ($CurrentStats.TotalGameTime -eq ""){
		$Script:CurrentStats.TotalGameTime = 0 #prevents errors from happening on first time run.
	} 
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
	}
	#Make Backup of CSV.
	 # Added this in as I had BSOD on my PC and noticed that this caused the files to get corrupted.
	Copy-Item -Path ($Script:WorkingDirectory + "\stats.csv") -Destination ($Script:WorkingDirectory + "\stats.backup.csv")
	Copy-Item -Path ($Script:WorkingDirectory + "\accounts.csv") -Destination ($Script:WorkingDirectory + "\accounts.backup.csv")
}
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
	ForEach ($Object in $Script:QualityArray | select-object type,probability){#convert PSOobjects to hashtable for enumerator
		$QualityHash.add($Object.type,$Object.probability) #add each PSObject to hash
	}
	$Script:ItemLookup = ForEach ($Entry in $QualityHash.GetEnumerator()){
		[System.Linq.Enumerable]::Repeat($Entry.Key, $Entry.Value) #This creates a hash table with 19036 normal items, 588 magic items, 200 rare items etc etc. Used later as a list to randomly pick from.
	}
}
Function CowKingKilled {
	Write-Host "`n                          You Killed the Cow King!" -foregroundcolor green
	Write-Host "                                $X[38;2;165;146;99;22mMoo.$X[0m"
	Write-Host "                                    $X[38;2;165;146;99;22mMoooooooo!$X[0m"
	$voice = New-Object -ComObject Sapi.spvoice
	$voice.rate = -4 #How quickly the voice message should be
	$voice.volume = $Config.DCloneAlarmVolume
	$voice.voice = $voice.getvoices() | Where-Object {$_.id -like "*David*"}
	$voice.speak("MooMoo-moo moo-moo, moo-moo, moo moo.") | out-null
	$CowLogo = @"
 -)[<-
   +[[[*
     *)}]>>=-----                   +:                   =-
      :=]#}[[[]))])>-               [>-     :------     =)*:
     :<]<<<][[]))][<*=              <]<>]]<>[%%%##}<><<<]]*
   :=*)))**)))[[]]>:                 -<)[}#%%%%%%%%%%%%[<
     -*><)]))))<)][[*:                    -[#%%%%%%%%%}-
       =*>><<<)<->[}}[*:                 *}#<>%%%)=)%%%}*
        :+>><)]])]]<[##[*:              *@@%%#[[[[}%%%%%}
          :-<<)))<)*  *#}[):            -}%@#[]]]]#%%%%%%>
              -=+><-   :>[#})+          :}%%#[[[]]###%%%%%>
                          -]#}}-:      -]%#}}[}}[[[[}#%%%%):
                            :>}})-    +#%%%#}}[][[)[}#%%%%#+
                               =##%%<=}}}}}}}}}}[[])}#%%%%%#+
                            :>[[##}}[[}[[[[[][]]]))))]]]]}#%%-
                            :]##}]<)][}[}#}}}}}[[[]])]][[}##%}:
                             =<}]>****+<[[}}}}[[[]]]]]][}}}###<
                                       :#%#[[[]]])))))][%%}[}}[-
                                       :#%###}[]]]))))][#%}]][]-
                                       :#%##}}}}[[[]]][[#}))<>-:
                                       :#####}}}}}[[[[]<>)}[:
                                         [#}}}}}}}}[}}##}%%+
                                         :#}}}}}}}}}[[[}}}}>-
                                         >##}}}}}}}}[[[[[}#[[<
                                        =}}}##}}}[[[[[[])]##<+
                                        >#}}#}}[[[[][]]]))#%*
                                       :]#}}###}[[[[[][[)]#[:
                                        *#}}#}}}}}}}[[[}[}#[:
          (__)                          *###}##}}}}}}}}}}#%):
          (oo)                          :[}}}}<]][[[}###}##+
   /-------\/                            :}}}}>    :)###}}#]-
  / |     ||                              +}}}>      +}%#}#%>
  * ||----||                              >}}[:        -=]}#<
    ||    ||                             -####*          )}]<:
                                        +#%%%}           -[]]>
                                        :-*+-:           -][])
                                                         =[]))+
                                                         =}##}):
"@
	Write-Host "  $X[38;2;255;165;0;22m$CowLogo$X[0m"
	$Script:CowKingActivated = $True
	([int]$Script:CurrentStats.CowKingKilled) ++
	SetQualityRolls
	start-sleep -milliseconds 4550
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played and cow stats.
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
	}
}
Function HighRune {
	process { Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4m$_$X[0m"}
}
Function Unique {
    process { Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4m$_$X[0m"}
}
Function SetItem {
    process { Write-Host "  $X[38;2;0;225;0;48;2;1;1;1;4m$_$X[0m"}
}
Function Rare {
    process { Write-Host "  $X[38;2;255;255;0;48;2;1;1;1;4m$_$X[0m"}
}
Function Magic {#ANSI text colour formatting for "magic" quotes. The variable $X (for the escape character) is defined earlier in the script.
    process { Write-Host "  $X[38;2;65;105;225;48;2;1;1;1;4m$_$X[0m" }
}
Function Normal {
    process { Write-Host "  $X[38;2;255;255;255;48;2;1;1;1;4m$_$X[0m"}
}
Function QuoteRoll {#stupid thing to draw a random quote but also draw a random quality.
	$Quality = get-random $Script:ItemLookup #pick a random entry from ItemLookup hashtable.
	Write-Host
	$LeQuote = (Get-Random -inputobject $Script:quotelist) #pick a random quote.
	$ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
	$DesiredIndent = 2  # indent spaces
	$ChunkSize = $ConsoleWidth - $DesiredIndent
	[RegEx]::Matches($LeQuote, ".{$ChunkSize}|.+").Groups.Value | ForEach-Object {
		Write-Output $_ | &$Quality #write the quote and write it in the quality colour
	}
	if ($LeQuote -match "Moo" -and $Quality -eq "Unique"){
		CowKingKilled
	}
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	if ($Quality -eq "HighRune"){([int]$Script:CurrentStats.HighRunesFound) ++}
	if ($Quality -eq "Unique"){([int]$Script:CurrentStats.UniquesFound) ++}
	if ($Quality -eq "SetItem"){([int]$Script:CurrentStats.SetItemsFound) ++}
	if ($Quality -eq "Rare"){([int]$Script:CurrentStats.RaresFound) ++}
	if ($Quality -eq "Magic"){([int]$Script:CurrentStats.MagicItemsFound) ++}
	if ($Quality -eq "Normal"){([int]$Script:CurrentStats.NormalItemsFound) ++}
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
		Start-Sleep -Milliseconds 256
	}
}
Function Inventory {#Info screen
	Clear-Host
	Write-Host "`n          Stay a while and listen! Here's your D2r Loader info.`n`n" -foregroundcolor yellow
	Write-Host "  $X[38;2;255;255;255;4mNote:$X[0m D2r Playtime is based on the time the script has been running"
	Write-Host "  whilst D2r is running. In other words, if you use this script when you're"
	Write-Host "  playing the game, it will give you a reasonable idea of the total time"
	Write-Host "  you've spent receiving disappointing drops from Mephisto :)`n"
	$QualityArraySum = 0
	$Script:QualityArray | ForEach-Object {
		$QualityArraySum += $_.Probability
	}
	$NormalProbability = ($QualityArray | where-object {$_.type -eq "Normal"} | Select-Object Probability).probability
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	$Line1 =   "                    ----------------------------------"
	$Line2 =  ("                   |  $X[38;2;255;255;255;22mD2r Playtime (Hours):$X[0m " +  ((($time =([TimeSpan]::Parse($CurrentStats.TotalGameTime))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line3 =  ("                   |  $X[38;2;255;255;255;22mCurrent Session (Hours):$X[0m" + ((($time =([TimeSpan]::Parse($Script:SessionTimer))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line4 =  ("                   |  $X[38;2;255;255;255;22mScript Launch Counter:$X[0m " + $CurrentStats.TimesLaunched)
	$Line5 =   "                    ----------------------------------"
	$Line6 =  ("                   |  $X[38;2;255;165;000;22mHigh Runes$X[0m Found: " + $(if ($CurrentStats.HighRunesFound -eq ""){"0"} else {$CurrentStats.HighRunesFound}))
	$Line7 =  ("                   |  $X[38;2;165;146;99;22mUnique$X[0m Quotes Found: " + $(if ($CurrentStats.UniquesFound -eq ""){"0"} else {$CurrentStats.UniquesFound}))
	$Line8 =  ("                   |  $X[38;2;0;225;0;22mSet$X[0m Quotes Found: " + $(if ($CurrentStats.SetItemsFound -eq ""){"0"} else {$CurrentStats.SetItemsFound}))
	$Line9 =  ("                   |  $X[38;2;255;255;0;22mRare$X[0m Quotes Found: " + $(if ($CurrentStats.RaresFound -eq ""){"0"} else {$CurrentStats.RaresFound}))
	$Line10 = ("                   |  $X[38;2;65;105;225;22mMagic$X[0m Quotes Found: " + $(if ($CurrentStats.MagicItemsFound -eq ""){"0"} else {$CurrentStats.MagicItemsFound}))
	$Line11 = ("                   |  $X[38;2;255;255;255;22mNormal$X[0m Quotes Found: " + $(if ($CurrentStats.NormalItemsFound -eq ""){"0"} else {$CurrentStats.NormalItemsFound}))
	$Line12 =  "                    ----------------------------------"
	$Line13 = ("                   |  $X[38;2;165;146;99;22mCow King Killed:$X[0m " + $(if ($CurrentStats.CowKingKilled -eq ""){"0"} else {$CurrentStats.CowKingKilled}))
	$Line14 = ("                   |  $X[38;2;255;0;255;22mGems Activated:$X[0m  " + $(if ($CurrentStats.Gems -eq ""){"0"} else {$CurrentStats.Gems}))
	$Line15 = ("                   |  $X[38;2;255;0;255;22mPerfect Gem Activated:$X[0m " + $(if ($CurrentStats.PerfectGems -eq ""){"0"} else {$CurrentStats.PerfectGems}))
	$Line16 =  "                    ----------------------------------"
	$Lines = @($Line1,$Line2,$Line3,$Line4,$Line5,$Line6,$Line7,$Line8,$Line9,$Line10,$Line11,$Line12,$Line13,$Line14,$Line15,$Line16)
	# Loop through each object in the array to find longest line (for formatting)
	ForEach ($Line in $Lines){
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -gt $LongestLine){
			$LongestLine = ($Line -replace '\[.*?22m', '' -replace '\[0m','').Length
		}
	}
	ForEach ($Line in $Lines){#Formatting nonsense to indent things nicely
		$Indent = ""
		$Dash = ""
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -lt $LongestLine + 2){
			if ($Line -notmatch "-"){
				while ((($Line -replace '\[.*?22m', '' -replace '\[0m','').Length + $Indent.length) -lt ($LongestLine + 2)){
					$Indent = $Indent + " "
				}
				Write-Host $Line.replace(":$X[0m",":$X[0m$Indent").replace(" Found:"," Found:$Indent") -nonewline
				Write-Host "  |" -nonewline
				Write-Host
			}
			else {
				while (($Line.Length + $Dash.length) -le ($LongestLine +1)){
					$Dash = $Dash + "-"
				}
				Write-Host $Line -nonewline
				Write-Host $Dash
			}
		}
		Else {
			Write-Host " |"
		}
	}
	Write-Host ("`n  Chance to find $X[38;2;65;105;225;22mMagic$X[0m quality quote or better: " + [math]::Round((($QualityArraySum - $NormalProbability + 1) * (1/$QualityArraySum) * 100),2) + "%" )
	Write-Host ("`n  $X[4mD2r Game Version:$X[0m    " + (Get-Command "$GamePath\D2R.exe").FileVersionInfo.FileVersion)
	Write-Host "  $X[4mScript Install Path:$X[0m " -nonewline
	Write-Host ("`"$Script:WorkingDirectory`"" -replace "((.{1,52})(?:\\|\s|$)|(.{1,53}))", "`n                        `$1").trim() #add two spaces before any line breaks for indenting. Add line break for paths that are longer than 53 characters.
	Write-Host "  $X[4mYour Script Version:$X[0m v$CurrentVersion"
	Write-Host "  $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/v$CurrentVersion$X[0m"
	if ($null -eq $Script:LatestVersionCheck -or $Script:LatestVersionCheck.tostring() -lt (Get-Date).addhours(-2).ToString('yyyy.MM.dd HH:mm:ss')){ #check for updates. Don't check if this has been checked in the couple of hours.
		try {
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/Diablo2RLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$Script:LatestVersionCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
			$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		}
		Catch {
			Write-Output "  Couldn't check for updates :(" | Red
		}
	}
	if ($Null -ne $Script:LatestVersion -and $Script:LatestVersion -gt $Script:CurrentVersion){
		Write-Host "`n  $X[4mLatest Script Version:$X[0m v$LatestVersion" -foregroundcolor yellow
		Write-Host "  $X[38;2;69;155;245;4mhttps://github.com/shupershuff/Diablo2RLoader/releases/latest$X[0m"
	}
	Write-Host "`n  $X[38;2;0;225;0;22mConsider donating as a way to say thanks via an option below:$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://www.buymeacoffee.com/shupershuff$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://paypal.me/Shupershuff$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://github.com/sponsors/shupershuff?frequency=one-time&amount=5$X[0m`n"
	if ($Script:NotificationsAvailable -eq $True){
		Write-Host "  -------------------------------------------------------------------------"
		Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4mNotification:$X[0m" -nonewline
		Notifications -check $False
		$Script:NotificationHasBeenChecked = $True
		Write-Host "  -------------------------------------------------------------------------"
	}
	Write-Host
	PressTheAnyKey
}
Function WindowMover { #Used to get window locations and place them in the same screen locations at launch. Code courtesy of Sir-Wilhelm and Microsoft.
    if ($Script:WindowMoverGoodToGo -ne $True){
        $Script:WindowMoverGoodToGo = $True
		. "$Script:WorkingDirectory\WindowMover.ps1"
	}
}
Function SaveWindowLocations {# Get Window Location coordinates and save to Accounts.csv
	WindowMover
	FormatFunction -indents 2 -text "Saving locations of each open account so that they the windows launch in the same place next time. Assumes you've configured the game to launch in windowed mode." 
	CheckActiveAccounts
	#If Feature is enabled, add 'WindowXCoordinates' and 'WindowYCoordinates' columns to accounts.csv with empty values.
	if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowYCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowWidth" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowHeight" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
		# Column does not exist, so add it to the CSV data
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowXCoordinates" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowYCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowYCoordinates" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowHeight" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowHeight" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowWidth" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowWidth" -NotePropertyValue $Null}
		}
		# Export the updated CSV data back to the file
		$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
	}
	if ($null -eq $Script:ActiveAccountsList){
		FormatFunction -text "`nThere are no open accounts to save coordinates from.`nTo save Window positions, you need to launch one or more instances first.`n" -indents 2 -IsError
		PressTheAnyKey
		Return $False
	}
	$NewCSV = ForEach ($Account in $Script:AccountOptionsCSV){
		if ($account.id -in $Script:ActiveAccountsList.id){
			$process = Get-Process -Id ($Script:ActiveAccountsList | where-object {$_.id -eq $account.id}).ProcessID
			$handle = $process.MainWindowHandle
			Write-Verbose "$($process.ProcessName) `(Id=$($process.Id), Handle=$handle`, Path=$($process.Path))"
			$rectangle = New-Object RECT
			[WindowAPI]::GetWindowRect($handle, [ref]$rectangle) | Out-Null
			FormatFunction -indents 2 -text "`nSaved Coordinates for account $($account.id) ($($account.AccountLabel))" -IsSuccess
			Write-Host "     X Position = $($rectangle.Left)" -Foregroundcolor Green
			Write-Host "     Y Position = $($rectangle.Top)" -Foregroundcolor Green
			write-Host "     Width = $($rectangle.Right - $rectangle.Left)" -Foregroundcolor Green
			write-Host "     Height = $($rectangle.Bottom - $rectangle.Top)" -Foregroundcolor Green
			$Account.WindowXCoordinates = $rectangle.Left
			$Account.WindowYCoordinates = $rectangle.Top
			$Account.WindowWidth = $rectangle.Right - $rectangle.Left
			$Account.WindowHeight = $rectangle.Bottom - $rectangle.Top	
			$Account
		}
		Else {#Leave as is.
			$Account
			write-Verbose "Account $($account.id) ($($account.AccountLabel)) isn't running."
		}
	}
	$NewCSV | Export-CSV "$Script:WorkingDirectory\Accounts.csv" -NoTypeInformation
	Write-Host "`n   Updated CSV with window positions." -foregroundcolor green
	start-sleep -milliseconds 2500
}
Function SetWindowLocations {#
	param(
		[int]$Id,
		[int]$X,
		[int]$Y,
		[int]$Width,
		[int]$Height
	)
	WindowMover
	$handle = (Get-Process -Id $Id).MainWindowHandle
	[WindowAPI]::MoveWindow($handle, $x, $y, $Width, $Height, $True)
	[WindowAPI]::SetForegroundWindow($handle)
}
Function Options {
	ImportXML
	Clear-Host
	Write-Host "`n This screen allows you to change script config options."
	Write-Host " Note that you can also change these settings (and more) in config.xml."
	Write-Host " Options you can change/toggle below:`n"
	$OptionList = "1","2","3","4","5","6","7","8"
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	if ($Script:Config.DefaultRegion -eq 1){
		$CurrentDefaultRegion = "NA"
	}
	ElseIf ($Script:Config.DefaultRegion -eq 2){
		$CurrentDefaultRegion = "EU"
	}
	ElseIf ($Script:Config.DefaultRegion -eq 3){
		$CurrentDefaultRegion = "Asia"
	}
	Write-Host "  $X[38;2;255;165;000;22m1$X[0m - $X[4mDefaultRegion$X[0m (Currently $X[38;2;255;165;000;22m$CurrentDefaultRegion$X[0m)"
	Write-Host "`n  $X[38;2;255;165;000;22m2$X[0m - $X[4mSettingSwitcherEnabled$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.SettingSwitcherEnabled -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)" 
	Write-Host "  $X[38;2;255;165;000;22m3$X[0m - $X[4mManualSettingSwitcherEnabled$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.ManualSettingSwitcherEnabled -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"
	Write-Host "  $X[38;2;255;165;000;22m4$X[0m - $X[4mRememberWindowLocations$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.RememberWindowLocations -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"
	Write-Host "`n  $X[38;2;255;165;000;22m5$X[0m - $X[4mDCloneTrackerSource$X[0m (Currently $X[38;2;255;165;000;22m$($Script:Config.DCloneTrackerSource)$X[0m)"
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text ("$X[38;2;255;166;000;22m6$X[0m - $X[4mDCloneAlarmList$X[0m (Currently $X[38;2;255;165;000;22m" + $(if ($Script:Config.DCloneAlarmList -eq ""){"Alarms are disabled"}Else{$Script:Config.DCloneAlarmList}) + "$X[0m)")
	if ($Script:Config.DCloneAlarmList -ne ""){
		Write-Host "  $X[38;2;255;165;000;22m7$X[0m - $X[4mDCloneAlarmLevel$X[0m (Currently $X[38;2;255;165;000;22m$($Script:Config.DCloneAlarmLevel)$X[0m)"
		Write-Host "  $X[38;2;255;165;000;22m8$X[0m - $X[4mDCloneAlarmVoice$X[0m (Currently $X[38;2;255;165;000;22m$($Script:Config.DCloneAlarmVoice)$X[0m)"
		Write-Host "  $X[38;2;255;165;000;22m9$X[0m - $X[4mDCloneAlarmVolume$X[0m (Currently $X[38;2;255;165;000;22m$($Script:Config.DCloneAlarmVolume)$X[0m)"
	}
	if ($Script:TokensConfigured -eq $True){
		foreach ($row in $Script:AccountOptionsCSV){
			if ($row.AuthenticationMethod -eq "Parameter"){$ParametersUsed = $True}
		}
		if ($ParametersUsed -eq $True){
			$OptionList += "t"
			Write-Host "`n  $X[38;2;255;165;000;22mt$X[0m - Temporarily force token authentication (for configured accounts ONLY)."
			if ($Script:ForceAuthToken -eq $True){
				Write-Host "      $X[4mForceAuthToken$X[0m (Currently $X[38;2;5;250;5;22mEnabled$X[0m)."
			}
			else {
				Write-Host "      $X[4mForceAuthToken$X[0m (Currently $X[38;2;255;165;000;22mDisabled$X[0m)."
			}
		}
	}
	Write-Host "`n Enter one of the above options to change the setting."
	Write-Host " Otherwise, press any other key to return to main menu... " -nonewline
	$Option = readkey
	Write-Host;Write-Host
	Function OptionSubMenu {
		param (
			[String]$Description,
			[hashtable]$OptionsList,
			[String]$OptionsText,
			[String]$Current,
			[String]$ConfigName,
			[switch]$OptionInteger
		)
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml" -Raw
		FormatFunction -indents 1 -text "Changing setting for $X[4m$($ConfigName)$X[0m (Currently $X[38;2;255;165;000;22m$($Current)$X[0m).`n" 
		FormatFunction -text $Description -indents 1
		Write-Host;Write-Host $OptionsText
		do {
			if ($OptionInteger -eq $True){
				Write-Host "   Enter a number between $X[38;2;255;165;000;22m1$X[0m and $X[38;2;255;165;000;22m99$X[0m or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = 1..99 # Allow user to enter 1 to 99
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True).tostring()
				$NewValue = $NewOptionValue
			}
			Else {
				Write-Host "   Enter " -nonewline;CommaSeparatedList -NoOr ($OptionsList.keys | sort-object); Write-Host " or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = $OptionsList.keys
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27).tostring()
				$NewValue = $($OptionsList[$NewOptionValue])
			}
			if ($NewOptionValue -notin $AcceptableOptions + "c" + "Esc"){
				Write-Host "   Invalid Input. Please enter one of the options above.`n" -foregroundcolor red
			}
		} until ($NewOptionValue -in $AcceptableOptions + "c" + "Esc")
		if ($NewOptionValue -in $AcceptableOptions){
			if ($NewOptionValue -ne "s" -and $NewOptionValue -ne "r"){
				try {
					$Pattern = "(<$ConfigName>)([^<]*)(</$ConfigName>)"
					$ReplaceString = '{0}{1}{2}' -f '${1}', $NewValue, '${3}'
					$NewXML = [regex]::Replace($Xml, $Pattern, $ReplaceString)
					$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
					return $True
				}
				Catch {
					Write-Host "`n  Was unable to update config :(" -foregroundcolor red
					start-sleep -milliseconds 2500
					return $False
				}
			}
			ElseIf ($NewOptionValue -eq "s"){
				$PositionsRecorded = SaveWindowLocations
				if ($PositionsRecorded -ne $False){#redundant?
					return $False
				}
			}
			ElseIf ($NewOptionValue -eq "r"){
				WindowMover
				CheckActiveAccounts
				if ($null -eq $Script:ActiveAccountsList){
					FormatFunction -text "`nThere are no open games.`nTo reset window positions, you need to launch one or more instances first.`n" -indents 2 -IsWarning
					PressTheAnyKey
					Return $False
				}
				ForEach ($Account in $Script:AccountOptionsCSV){
					if ($account.id -in $Script:ActiveAccountsList.id){
						if ($account.WindowXCoordinates -ne "" -and $account.WindowYCoordinates -ne ""){
							SetWindowLocations -X $Account.WindowXCoordinates -Y $Account.WindowYCoordinates -Width $Account.WindowWidth -height $Account.WindowHeight -Id ($Script:ActiveAccountsList | where-object {$_.id -eq $account.id}).ProcessID | out-null
						}
					}
				}
				FormatFunction -indents 2 -text "Moved game windows back to their saved screen coordinates and reset window sizes.`n" -IsSuccess
				start-sleep -milliseconds 1500
				Return $False
			}
		}
		else {
			Return $False
		}
	}
	if ($Option -eq "1"){ #DefaultRegion
		$Options = @{
			"1" = 1
			"2" = 2
			"3" = 3
		}
		$XMLChanged = OptionSubMenu -ConfigName "DefaultRegion" -OptionsList $Options -Current $CurrentDefaultRegion `
		-Description "This option is used so you can press enter instead of manually entering region on region select screen." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' for NA (Americas)`n    Choose '$X[38;2;255;165;000;22m2$X[0m' for EU (Europe)`n    Choose '$X[38;2;255;165;000;22m3$X[0m' for Asia (Also known as KR)`n"
	}
	ElseIf ($Option -eq "2"){ #SettingSwitcherEnabled
		If ($Script:Config.SettingSwitcherEnabled -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False"}
			$OptionsSubText = "disable"
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "SettingSwitcherEnabled" -OptionsList $Options -Current $CurrentState `
		-Description "This enables the script to automatically switch which settings file to use when launching the game based on the account you're launching.`nA very cool feature!`nPlease see GitHub for instructions on setting this up/editing settings." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "3"){ #ManualSettingSwitcherEnabled
		If ($Script:Config.ManualSettingSwitcherEnabled -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False"}
			$OptionsSubText = "disable"
			$Script:AskForSettings = $False
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "ManualSettingSwitcherEnabled" -OptionsList $Options -Current $CurrentState `
		-Description "This enables you to manually choose which settings file the game should use launching another game instance.`nFor example if you want to choose to launch with potato graphics or good graphics.`nPlease see GitHub for instructions on how to set this up and how to edit settings." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "4"){ #RememberWindowLocations
		If ($Script:Config.RememberWindowLocations -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$DescriptionSubText = "`nOnce enabled, return to this menu and choose the '$X[38;2;255;165;000;22ms$X[0m' option to save coordinates of any open game instances."
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False";"S" = "PlaceholderValue Only :)";"R" = "PlaceholderValue Only :)"} # SaveWindowLocations function used if user chooses "S"
			$OptionsSubText = "disable"
			$OptionsSubTextAgain = "    Choose '$X[38;2;255;165;000;22ms$X[0m' to save current window locations and sizes.`n"
			if ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue){
				$OptionsSubTextAgain += "    Choose '$X[38;2;255;165;000;22mr$X[0m' to reset window locations and sizes.`n"
			}
			$DescriptionSubText = "`nChoosing the '$X[38;2;255;165;000;22ms$X[0m' option will save coordinates (and window sizes) of any open game instances.`nChoosing the '$X[38;2;255;165;000;22mr$X[0m' option will move your windows back to their default placements."
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "RememberWindowLocations" -OptionsList $Options -Current $CurrentState `
		-Description "For those that have configured the game to launch in windowed mode, this setting is used to make the script move the window locations at launch, so that you never have to rearrange your windows when launching accounts.$DescriptionSubText" `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n$OptionsSubTextAgain"
	}
	ElseIf ($Option -eq "5"){ #DCloneTrackerSource
		$Options = @{
			"1" = "D2Emu.com"
			"2" = "D2runewizard.com"
			"3" = "diablo2.io"
		}
		$XMLChanged = OptionSubMenu -ConfigName "DCloneTrackerSource" -OptionsList $Options -Current $Script:Config.DCloneTrackerSource `
		-Description "Choose the API source for DClone Data.`nRecommend D2Emu.com as it pulls data directly from the game." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' for D2Emu.com (Recommended)`n    Choose '$X[38;2;255;165;000;22m2$X[0m' for D2runewizard.com`n    Choose '$X[38;2;255;165;000;22m3$X[0m' for diablo2.io`n"
	}
	ElseIf ($Option -eq "6"){ #DCloneAlarmList
		$Options = @{ # SCL-NA, SCL-EU, SCL-KR, SC-NA, SC-EU, SC-KR, HCL-NA, HCL-EU, HCL-KR, HC-NA, HC-EU, HC-KR
			"1" = "SCL-NA, SCL-EU, SCL-KR"
			"2" = "SC-NA, SC-EU, SC-KR"
			"3" = "HCL-NA, HCL-EU, HCL-KR"
			"4" = "HC-NA, HC-EU, HC-KR"
			"5" = "SCL-NA, SCL-EU, SCL-KR, SC-NA, SC-EU, SC-KR, HCL-NA, HCL-EU, HCL-KR, HC-NA, HC-EU, HC-KR"
			"6" = ""
		}
		$XMLChanged = OptionSubMenu -ConfigName "DCloneAlarmList" -OptionsList $Options -Current $(if ($Script:Config.DCloneAlarmList -eq ""){"Alarms are disabled"}Else{$Script:Config.DCloneAlarmList}) `
		-Description "Use this to change what game modes you'd like to have DClone alarms on.`nThis will select all regions for a given game mode. To make fine tuned adjustments (eg multiple game modes), edit config.xml directly." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to enable alarms for Softcore Ladder`n    Choose '$X[38;2;255;165;000;22m2$X[0m' to enable alarms for Softcore Non-Ladder`n    Choose '$X[38;2;255;165;000;22m3$X[0m' to enable alarms for Hardcore Ladder`n    Choose '$X[38;2;255;165;000;22m4$X[0m' to enable alarms for Hardcore Non-Ladder`n    Choose '$X[38;2;255;165;000;22m5$X[0m' to enable alarms for all game modes`n    Choose '$X[38;2;255;165;000;22m6$X[0m' to disable all Alarms`n"
		$Script:DCloneChangesCSV = $Null #Reset DClone tracking to remove old notifications appearing that may no longer be applicable.
		SetDCloneAlarmLevels
	}
	ElseIf ($Option -eq "7" -and $Script:Config.DCloneAlarmList -ne ""){ #DCloneAlarmLevel
		$Options = @{
			"1" = "All"
			"2" = "Close"
			"3" = "Imminent"
		}
		$XMLChanged = OptionSubMenu -ConfigName "DCloneAlarmLevel" -OptionsList $Options -Current $Script:Config.DCloneAlarmLevel `
		-Description "This allows you to customise what DClone status changes you want to be alarmed for if you only want specific alerts.`nBe aware that if you set it to immminent, you will likely miss DClone walks due to fast status changes (SOJ's are often sold in bulk all at once)." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to alarm on all status changes`n    Choose '$X[38;2;255;165;000;22m2$X[0m' to alarm on Close status changes (4/6, 5/6)`n    Choose '$X[38;2;255;165;000;22m3$X[0m' to alarm on imminent walks only (5/6)`n"
		$Script:DCloneChangesCSV = $Null #Reset DClone tracking to remove old notifications appearing that may no longer be applicable.
		SetDCloneAlarmLevels
	}
	ElseIf ($Option -eq "8" -and $Script:Config.DCloneAlarmList -ne ""){ #DCloneAlarmVoice
		$Options = @{
			"1" = "Amazon"
			"2" = "Paladin"
		}
		$XMLChanged = OptionSubMenu -ConfigName "DCloneAlarmVoice" -OptionsList $Options -Current $Script:Config.DCloneAlarmVoice `
		-Description "This option allows you to change the voice for the Text to Speech DClone Alarms." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' for Amazon (Female Voice)`n    Choose '$X[38;2;255;165;000;22m2$X[0m' for EU (Europe)`n    Choose '$X[38;2;255;165;000;22m3$X[0m' for Paladin (Male Voice)`n"
	}
	ElseIf ($Option -eq "9" -and $Script:Config.DCloneAlarmList -ne ""){ #DCloneAlarmVolume
		$XMLChanged = OptionSubMenu -ConfigName "DCloneAlarmVolume" -OptionInteger -Current $Script:Config.DCloneAlarmVolume `
		-Description "This options allows you to adjust the volume level for DClone alarms in case they're too loud or quiet."
	}
	ElseIf ($Option -eq "t"){
		if ($Script:ForceAuthToken -ne $True){
			$Script:ForceAuthToken = $True
			Write-Host "  ForceAuthToken now $X[38;2;5;250;5;22menabled$X[0m." -foregroundcolor yellow
			Write-Host "  For authentication, script will now force usage of AuthTokens (where " -foregroundcolor green
			Write-Host "  accounts have been configured)." -foregroundcolor green
		}
		else {
			$Script:ForceAuthToken = $False
			Write-Host "  ForceAuthToken now $X[38;2;255;165;000;22mdisabled$X[0m." -foregroundcolor yellow
			Write-Host "  For authentication, script will now use what's configured in the " -foregroundcolor green
			Write-Host "  AuthenticationMethod column in accounts.csv" -foregroundcolor green
		}
		start-sleep -milliseconds 3000
	}
	else {#go to main menu if no valid option was specified.
		return
	}
	if ($XMLChanged -eq $True){
		Write-Host "   Config Updated!" -foregroundcolor green
		ImportXML
		If ($Option -eq "4"){
			if ((Test-Path -Path ($workingdirectory + 'WindowMover.ps1')) -ne $True){
				$WindowMoverUrl = "https://raw.githubusercontent.com/shupershuff/Diablo2RLoader/main/WindowMover.ps1"
				Invoke-WebRequest -Uri $WindowMoverUrl -OutFile "$WorkingDirectory\WindowMover.ps1"
			}
			if ($Script:Config.RememberWindowLocations -eq $True -and -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#if this is the first time it's been enabled display a setup message
				Formatfunction -indents 2 -IsWarning -Text "`nYou've enabled RememberWindowsLocations but you still need to set it up. To set this up you need to perform the following steps:"
				FormatFunction -indents 3 -iswarning -SubsequentLineIndents 3 -text "`n1. Open all of your D2r account instances.`n2. Move the window for each game instance to your preferred layout and size."
				FormatFunction -indents 3 -iswarning -SubsequentLineIndents 3 -text "3. Come back to this options menu and go into the 'RememberWindowLocations' setting.`n4. Once in this menu, choose the option 's' to save coordinates of any open game instances."
				FormatFunction -indents 2 -iswarning -text  "`n`nNow when you open these accounts they will open in this screen location each time :)`n"
				PressTheAnyKey
			}
		}
		start-sleep -milliseconds 2500
	}
}
Function Notifications {
	param (
		[bool] $Check
	)
if ($Check -eq $True -and $Script:LastNotificationCheck -lt (Get-Date).addminutes(-30).ToString('yyyy.MM.dd HH:mm:ss')){#check for notifications once every 30mins
		try {
			$URI = "https://raw.githubusercontent.com/shupershuff/Diablo2RLoader/main/Notifications.txt"
			$Script:Notifications = Invoke-RestMethod -Uri $URI
			if ($Notifications.notification -ne ""){
				if ($Script:PrevNotification -ne $Notifications.notification){#if message has changed since last check
					$Script:PrevNotification = $Notifications.notification
					$Script:NotificationHasBeenChecked = $False
					if ((get-date).tostring('yyyy.MM.dd HH:mm:ss') -lt $Notifications.ExpiryDate -and (get-date).tostring('yyyy.MM.dd HH:mm:ss') -gt $Notifications.PublishDate){
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
			Write-Debug "  Couldn't check for notifications." # If this fails in production don't show any errors/warnings.
		}
	}
	ElseIf ($Check -eq $False){
		Write-Host
		formatfunction -text $Notifications.notification -indents 1
	}
	if ($Check -eq $True -and $Script:NotificationHasBeenChecked -eq $False -and $Script:NotificationsAvailable -eq $True){#only show message if user hasn't seen notification yet.
		Write-Host "     $X[38;2;255;165;000;48;2;1;1;1;4mNotification available. Press 'i' to go to info screen for details.$X[0m"
	}#%%%%%%%%%%%%%%%%%%%%
}
Function D2rLevels {
	$Script:D2rLevels = @(
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
}
Function QuoteList {
$Script:QuoteList =
"Stay a while and listen..",
"My brothers will not have died in vain!",
"My brothers have escaped you...",
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
"East... Always into the east...",
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
"'Live, Laugh, Love' - Andariel, 1264.",
"Oh no, snakes. I hate snakes.",
"Who would have thought that such primitive beings could cause so much `ntrouble.",
"Hail to you champion",
"Help us! LET US OUT!",
"'I cannot carry anymore' - Me, carrying my teammates.",
"You're an even greater warrior than I expected...Sorry for `nunderestimating you.",
"Cut them down, warrior. All of them!",
"How can one kill what is already dead?",
"...That which does not kill you makes you stronger."
}
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
  %%%%%$Script:MOO%%%%%%%*&@&&&%&&(,. @@@@,(%%%%%%#/,@@@& *#&&@&&%(%%%%%%%$Script:MOO%%%%%%
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
	if ($Script:PGemActivated -eq $True -or $Script:CowKingActivated -eq $True ){
		Write-Host "  $X[38;2;255;165;0;22m$BannerLogo$X[0m"
	}
	Else {
		Write-Host $BannerLogo -foregroundcolor yellow
	}
}
Function JokeMaster ([int]$JokeProviderRoll=""){
	#If you're going to leech and not provide any damage value in the Throne Room then at least provide entertainment value right?
	Write-Host "  Copy these mediocre jokes into the game while doing Baal Comedy Club runs`r`n  to mislead everyone into believing you have a personality:`n"
	do {
		$Joke = $Null
		$JokeType = "Joke"
		if ($JokeProviderRoll -eq "" -or $JokeProviderRoll -eq 69){
			$JokeProviderRoll = get-random -min 1 -max 3 #Randomly roll for first two joke providers
		}
		$attempt = 0
		do {
			$attempt ++
			if ($JokeProviderRoll -eq 1){
				if ((Get-Date).Month -eq 12 -or ((Get-Date).Month -eq 10 -and (Get-Date).Day -ge 24 -and (Get-Date).Day -le 31)){#If December or A week leading up to Halloween.
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
					$Joke = Invoke-RestMethod -uri https://official-joke-api.appspot.com/random_joke -Method GET -header $headers -ErrorAction Stop
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
			if ($JokeProviderRoll -eq 3){
				try {
					$headers = @{
						"User-Agent" = "github.com/shupershuff/Diablo2RLoader"
						"Accept" = "application/json"
					}
					$Joke = (Invoke-RestMethod -uri https://icanhazdadjoke.com -Method GET -header $headers -ErrorAction Stop).joke
					$JokeObtained = $true
					$SingleJoke =  ($joke -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
					Write-Host "   $X[38;2;255;165;000;22m$SingleJoke$X[0m"
					$JokeProvider = "icanhazdadjoke.com"
				}
				catch {
					$JokeProviderRoll = 1
					Write-Host "  Couldn't reach icanhazdadjoke.com :( API might be down." -foregroundcolor Red
				}
			}
			if ($JokeProviderRoll -eq 4){
				try {
					$Joke = (Invoke-RestMethod -uri https://api.chucknorris.io/jokes/random -ErrorAction Stop).value
					$JokeObtained = $true
					$SingleJoke =  ($joke -replace "(.{1,73})(?:\s|$)", "`n   `$1").trim() #add two spaces before any line breaks for indenting. Add line break for lines that are longer than 73 characters.
					Write-Host "   $X[38;2;255;165;000;22m$SingleJoke$X[0m"
					$JokeProvider = "api.chucknorris.io"
					$JokeType = "Fact"
				}
				catch {
					$JokeProviderRoll = 1
					Write-Host "  Couldn't reach api.chucknorris.io :( API might be down." -foregroundcolor Red
				}
			}
		} until ($JokeObtained -eq $true -or $attempt -eq 3)
		Write-Host
		if ($Null -ne $Joke){
			Write-Host "  $JokeType courtesy of $JokeProvider`n`n"
		}
		Write-Host "  Press '$X[38;2;255;165;000;22mj$X[0m' for more, '$X[38;2;255;165;000;22md$X[0m' for a Dad joke or '$X[38;2;255;165;000;22mc$X[0m' for a Chuck Norris fact."
		Write-Host "  Otherwise, press any other key to return to main menu... " -nonewline
		$JokeOption = readkey
		Write-Host;Write-Host
		if ($JokeOption -eq "j" -or $JokeOption -eq "d" -or $JokeOption -eq "c"){
			if ($JokeOption -eq "j"){
				$JokeProviderRoll = 69 #fish out another cringey joke.
			}
			if ($JokeOption -eq "d"){
				$JokeProviderRoll = 3
			}
			if ($JokeOption -eq "c"){
				$JokeProviderRoll = 4
			}
			# Clear faff from the console so more jokes show.
			$InputLength = 138
			$Host.UI.RawUI.CursorPosition = @{
				X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - $InputLength, 0)
				Y = $Host.UI.RawUI.CursorPosition.Y -3
			}
			Write-Host -NoNewLine (" " * ($InputLength)) #-ForegroundColor Black
			$Host.UI.RawUI.CursorPosition = @{
				X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - $InputLength, 0)
				Y = $Host.UI.RawUI.CursorPosition.Y -2
			}
		}
	} until ($JokeOption -ne "j" -and $JokeOption -ne "d" -and $JokeOption -ne "c")
}
Function DClone {# Display DClone Status.
	param (
		[bool] $DisableOutput,
		[String] $DCloneTrackerSource,
		[String] $Taglist,
		[object] $DCloneChanges,
		[String] $DCloneAlarmLevel
	)
	if ($DCloneTrackerSource -eq "d2emu.com"){
		$URI = "https://d2emu.com/api/v1/dclone"
		try {
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $DCloneTrackerSource -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET
			} -TimeoutSeconds 3
			$D2RDCloneResponse = $D2RDCloneResponse.PSObject.Properties | ForEach-Object {
				[PSCustomObject]@{
					Name = $_.Name
					Progress = $_.Value.status
					Updated_at = $_.Value.updated_at
				}
			}
			$CurrentStatus = $D2RDCloneResponse | Select-Object @{Name='Server'; Expression={$_.name}},@{Name='Progress'; Expression={($_.Progress + 1)}} #| sort server #add +1 as this source counts status from 0
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function
			Write-Debug "Problem connecting to $URI"
		}
	}
	ElseIf ($DCloneTrackerSource -eq "D2runewizard.com"){
		$QLC = "zouaqcSTudL"
		$tokreg = ("QLC" + $qlc + 1 +"fnbttzP")
		$D2RWref = ""
		for ($i = $tokreg.Length - 1; $i -ge 0; $i--){
			$D2RWref += $tokreg[$i]
		}
		$headers = @{
			"D2R-Contact" = "placeholderemail@email.com"
			"D2R-Platform" = "GitHub"
			"D2R-Repo" = "https://github.com/shupershuff/Diablo2RLoader"
		}
		$URI = "https://d2runewizard.com/api/diablo-clone-progress/all?token=$D2RWref"
		try {
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $DCloneTrackerSource -Headers -$headers -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET -Headers $using:Headers
			} -TimeoutSeconds 3
			$CurrentStatus = $D2RDCloneResponse.servers | Select-Object @{Name='Server'; Expression={$_.server}},@{Name='Progress'; Expression={$_.progress}} #| sort server
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function
			Write-Debug "Problem connecting to $URI"
		}
	}
	ElseIf ($DCloneTrackerSource -eq "Diablo2.io"){
		$headers = @{
			"User-Agent" = "github.com/shupershuff/Diablo2RLoader"
		}
		$URI = "https://diablo2.io/dclone_api.php"
		try {
			$D2RDCloneResponse = WebRequestWithTimeOut -InitiatingFunction "DClone" -DCloneSource $DCloneTrackerSource -ScriptBlock {
				Invoke-RestMethod -Uri $using:URI -Method GET -Headers $using:Headers
			} -TimeoutSeconds 3
			$CurrentStatus = $D2RDCloneResponse | Select-Object @{Name='Server'; Expression={$_.region}},@{Name='Ladder'; Expression={$_.ladder}},@{Name='Core'; Expression={$_.hc}},@{Name='Progress'; Expression={$_.progress}}
		}
		Catch {#catch commands captured in WebRequestWithTimeOut function
			Write-Debug "Problem connecting to $URI"
		}
	}
	Else {#if XML is invalid for DCloneTrackerSource
		$DCloneErrorMessage = ("  Error: Couldn't check for DClone Status. ###  Check DCloneTrackerSource in config.xml is entered correctly.").Replace("###", "`n")
		Write-Host
		Write-Host $DCloneErrorMessage -Foregroundcolor red
		if ($DisableOutput -ne $True){
			Write-Host
			$Script:AccountID = $null
			Presstheanykey
		}
		Return
	}
	if ($null -ne $DCloneErrorMessage){
		Write-Host $DCloneErrorMessage -Foregroundcolor red
		if ($DisableOutput -ne $True){
			Write-Host
			$Script:AccountID = $null
			Presstheanykey
		}
		Return
	}
	$DCloneLadderTable = New-Object -TypeName System.Collections.ArrayList
	$DCloneNonLadderTable = New-Object -TypeName System.Collections.ArrayList
	if ($DCloneChanges -eq "" -or $null -eq $DCloneChanges){
		$DCloneChangesArray = New-Object -TypeName System.Collections.ArrayList
	}
	Else {
		$DCloneChangesArray = $DCloneChanges | ConvertFrom-Csv -ErrorAction silentlycontinue #temporarily convert to array
	}
	ForEach ($Status in $CurrentStatus){
		$DCloneLadderInfo = New-Object -TypeName psobject
		$DCloneNonLadderInfo = New-Object -TypeName psobject
		#Convert data from all sources into consistent names and tags to be sorted and filtered.
		if ($Status.server -like "*us*" -or $Status.server -like "*americas*" -or $Status.Server -eq "1"){$Tag = "-NA";$ServerName = "Americas"}
		ElseIf ($Status.server -like "*eu*" -or $Status.server -like "*europe*" -or $Status.Server -eq "2"){$Tag = "-EU";$ServerName = "Europe"}
		ElseIf ($Status.server -like "*kr*" -or $Status.server -like "*asia*" -or $Status.Server -eq "3"){$Tag = "-KR";$ServerName = "Asia"}
		if (($Status.server -notlike "*nonladder*" -and -not [int]::TryParse($Status.server,[ref]$null)) -or $Status.Ladder -eq "1"){
			if ($Status.server -match "hardcore" -or $Status.Core -eq "1"){$Tag = ("HCL" + $Tag);$ServerName = ("HCL - " + $ServerName)}
			else {$Tag = ("SCL" + $Tag);$ServerName = ("SCL - " + $ServerName)}
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name Tag -Value $Tag
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name LadderServer -Value $ServerName
			$DCloneLadderInfo | Add-Member -MemberType NoteProperty -Name LadderProgress -Value $Status.progress
			[VOID]$DCloneLadderTable.Add($DCloneLadderInfo)
		}
		Else {
			if ($Status.server -match "hardcore" -or $Status.Core -eq "1"){$Tag = ("HC" + $Tag);$ServerName = ("HC - " + $ServerName)}
			else {$Tag = ("SC" + $Tag);$ServerName = ("SC - " + $ServerName)}
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name Tag -Value $Tag
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderServer -Value $ServerName
			$DCloneNonLadderInfo | Add-Member -MemberType NoteProperty -Name NonLadderProgress -Value $Status.progress
			[VOID]$DCloneNonLadderTable.Add($DCloneNonLadderInfo)
		}
		if ($True -eq $DisableOutput){
			if ($taglist -match $Tag ){#if D Dclone region and server matches what's in config, check for changes.
				Write-Debug " Tag $tag in taglist" #debug
				if ($DCloneChangesArray | where-object {$_.Tag -eq $Tag}){
					ForEach ($Item in $DCloneChangesArray | where-object {$_.Tag -eq $Tag}){#for each tag specified in config.xml...
						$item.VoiceAlarmStatus = $False
						$item.TextAlarmStatus = $False
						if ($item.Status -ne $Status.progress){
							if ($Status.progress -ge 5 -and ($DCloneAlarmLevel -match 5)){#if DClone walk is imminent
								$item.VoiceAlarmStatus = $True
							}
							ElseIf ($Status.progress -eq 1 -or $Status.progress -eq 6){#if DClone walk has just happened
								$item.VoiceAlarmStatus = $True
							}
							ElseIf ($DCloneAlarmLevel -match $Status.progress){#if User has configured alarms to happen for lower DClone status levels (2,3,4)
								$item.VoiceAlarmStatus = $True
							}
							$item.TextAlarmStatus = $True
							if ($item.Status -ne "" -and $null -ne $item.Status){
								$item.PreviousStatus = $item.Status
							}
							$item.Status = $Status.progress
							$item.LastUpdate = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
						}
						ElseIf ($Status.progress -eq 5){#if status hasn't changed, but status is 5 (imminent), show alarm text on main menu
							$item.TextAlarmStatus = $True
						}
						ElseIf ($Status.progress -lt 5 -and $item.LastUpdate -gt (get-date).addminutes(-5).ToString('yyyy.MM.dd HH:mm:ss')){#if status is less than 5 and has changed within the last 5 minutes, enable text alarm
							$item.TextAlarmStatus = $True
						}
						ElseIf ($null -ne $item.LastUpdate -and $item.LastUpdate -lt (get-date).addminutes(-5).ToString('yyyy.MM.dd HH:mm:ss')){#after 5 minutes remove the text alarm
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
						$DCloneChangesArray += [PSCustomObject]@{Tag = $Tag; PreviousStatus = $null; Status = $Status.progress; VoiceAlarmStatus = $VoiceAlarmStatus; TextAlarmStatus = $TextAlarmStatus; LastUpdate = $null}
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
		$DCloneLadderTable = $DCloneLadderTable | Sort-Object LadderServer
		$DCloneNonLadderTable = $DCloneNonLadderTable | Sort-Object NonLadderServer
		$Count = 0
		Do {
			if ($Count -eq 0){
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
				Write-Host "`n   DClone Status provided by $DCloneTrackerSource`n"
			}
		}
		Until ($Count -eq 6)
		PressTheAnyKey
	}
	ElseIf ($Taglist -ne "" -and $null -ne $DCloneChangesArray){#Else if Output is disabled and taglist has been specified, output dclone changes for alarm
		$DCloneChangesArray | ConvertTo-Csv -NoTypeInformation
	}
}
Function DCloneVoiceAlarm {
	$voice = New-Object -ComObject Sapi.spvoice
	$voice.rate = -2 #How quickly the voice message should be
	$voice.volume = $Config.DCloneAlarmVolume
	Write-Host
	if ($Script:Config.DCloneAlarmVoice -eq "Bloke" -or $Script:Config.DCloneAlarmVoice -eq "Man" -or $Script:Config.DCloneAlarmVoice -eq "Paladin"){$voice.voice = $voice.getvoices() | Where-Object {$_.id -like "*David*"}}
	ElseIf ($Script:Config.DCloneAlarmVoice -eq "Wench" -or $Script:Config.DCloneAlarmVoice -eq "Woman" -or $Script:Config.DCloneAlarmVoice -eq "Amazon"){$voice.voice = $voice.getvoices() | Where-Object {$_.id -like "*ZIRA*"}}
	else {break}# If specified voice doesn't exist
	ForEach ($Item in ($Script:DCloneChangesCSV | ConvertFrom-Csv) | where-object {$_.VoiceAlarmStatus -Match "True" -or $_.TextAlarmStatus -Match "True"}){
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
		ElseIf ($item.tag -match "KR"){
			$DCloneRegion = "Asia"
		}
		else {
			$DCloneRegion = "Europe"
		}
		if ($Item.Status -eq 5){
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone is about to walk in $DCloneRegion on $CoreText $LadderText ($($item.tag))!$X[0m"
			$Message = ("D Clone Imminent! DClone is about to walk in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		ElseIf (($Item.Status -eq 1 -and $Item.PreviousStatus -ne 6) -or $Item.Status -eq 6){#check if status has just changed to 6 or it has changed to 1 from any number other than 6 (to prevent duplicate alarms.
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone has just walked in $DCloneRegion on $CoreText $LadderText ($($item.tag)).$X[0m"
			$Message = ("D Clone has just walked in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		ElseIf ($Script:DCloneAlarmLevel -match $Item.Status){
			Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4mDClone Update! DClone is now $($Item.Status)/6 in $DCloneRegion on $CoreText $LadderText ($($item.tag))$X[0m"
			$Message = ("D Clone is now " + $Item.Status + " out of 6 in $DCloneRegion on " + $CoreText + " " + $LadderText)
		}
		if ($item.VoiceAlarmStatus -eq $True){
			$voice.speak("$Message") | out-null
		}
	}
	if ($null -ne $Message){
		Write-Host "  $X[38;2;065;105;225;48;2;1;1;1;4mDClone status provided by $($Script:Config.DCloneTrackerSource)$X[0m"
	}
}
Function WebRequestWithTimeOut {#Used to timeout web requests that take too long.
	param (
		[ScriptBlock] $ScriptBlock,
		[int] $TimeoutSeconds,
		[String] $InitiatingFunction,
		[String] $DCloneSource
	)
	$Script:DCloneErrorMessage = $null
	$TimedJob = Start-Job -ScriptBlock $ScriptBlock
	$timer = [Diagnostics.Stopwatch]::StartNew()
	while ($TimedJob.State -eq "Running" -and $timer.Elapsed.TotalSeconds -lt $TimeoutSeconds){
		Start-Sleep -Milliseconds 10
	}
	if ($TimedJob.State -eq "Running"){
		Stop-Job -Job $TimedJob
		if ($InitiatingFunction -eq "DClone"){
			$Script:DCloneErrorMessage = "   Error: Couldn't connect to $DCloneTrackerSource to check for DClone Status."
			Write-Host ("`n   Timed out connecting to DClone Data Source.") -foregroundcolor red
			Throw "Timed Out :(" #force an exception to break out of the try statement.
			Write-Host
		}
	}
	ElseIf ($TimedJob.State -eq "Completed"){
		$result = Receive-Job -Job $TimedJob
		Write-Verbose "Command completed successfully."
		ForEach ($Object in $result){#Remove Properties from Result Array inserted by the Start-Job command. This prevents skewed data for DClone status checks.
			$Object.PSObject.Properties.Remove("RunspaceId")
			$Object.PSObject.Properties.Remove("PSComputerName")
			$Object.PSObject.Properties.Remove("PSShowComputerName")
		}
		$result
	}
	else {
		if ($InitiatingFunction -eq "DClone"){
			Write-Host " Couldn't connect to $DCloneSource." -foregroundcolor red
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
	ForEach ($Level in $D2TZResponse.current){
		Write-Debug "Level ID is: $Level"
		ForEach ($LevelID in $D2rLevels){
			if ($LevelID[0] -eq $Level){
				$CurrentTZ += $LevelID[1] + ", "
			}
		}
	}
	$CurrentTZ = $CurrentTZ -replace '..$', ''
	ForEach ($Level in $D2TZResponse.next){
		Write-Debug "Level ID is: $Level"
		ForEach ($LevelID in $D2rLevels){
			if ($LevelID[0] -eq $Level){
				$NextTZ += $LevelID[1] + ", "
			}
		}
	}
	$NextTZ = $NextTZ -replace '..$', ''
	Write-Host "`n   Current TZ is:  " -nonewline;Write-Host ($CurrentTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
	Write-Host "   Next TZ is:     " -nonewline;Write-Host ($NextTZ -replace "(.{1,58})(\s+|$)", "`$1`n                   ").trim() -ForegroundColor magenta
	Write-Host "`n  Information Retrieved at: " $TimeDataObtained
	Write-Host "  TZ info courtesy of:       $TZProvider`n"
	PressTheAnyKey
}
Function KillHandle { #Thanks to sir-wilhelm for tidying this up.
	$handle64 = "$PSScriptRoot\handle\handle64.exe"
	$handle = & $handle64 -accepteula -a -p D2R.exe "Check For Other Instances" -nobanner | Out-String
	if ($handle -match "pid:\s+(?<d2pid>\d+)\s+type:\s+Event\s+(?<eventHandle>\w+):"){
		$d2pid = $matches["d2pid"]
		$eventHandle = $matches["eventHandle"]
		Write-Verbose "Closing handle: $eventHandle on pid: $d2pid"
		& $handle64 -c $eventHandle -p $d2pid -y #-nobanner
	}
}
Function CheckActiveAccounts {#Note: only works for accounts loaded by the script
	#check if there's any open instances and check the game title window for which account is being used.
	try {
		$Script:ActiveIDs = $Null
		$D2rRunning = $false
		$Script:ActiveIDs = New-Object -TypeName System.Collections.ArrayList
		$Script:ActiveIDs = (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "- Diablo II: Resurrected"} | Select-Object MainWindowTitle).mainwindowtitle.substring(0,2).trim() #find all diablo 2 game windows and pull the account ID from the title
		$Script:D2rRunning = $true
		Write-Verbose "Running Instances."
	}
	catch {#if the above fails then there are no running D2r instances.
		$Script:D2rRunning = $false
		Write-Verbose "No Running Instances."
		$Script:ActiveIDs = ""
	}
	if ($Script:D2rRunning -eq $True){
		$Script:ActiveAccountsList = New-Object -TypeName System.Collections.ArrayList
		ForEach ($ActiveID in $ActiveIDs){#Build list of active accounts that we can omit from being selected later
			$ActiveAccountDetails = $Script:AccountOptionsCSV | where-object {$_.id -eq $ActiveID}
			$ActiveAccount = New-Object -TypeName psobject
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ID -Value $ActiveAccountDetails.ID
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name AccountName -Value $ActiveAccountDetails.accountlabel
			$InstanceProcessID = (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "$($ActiveAccountDetails.ID) - $($ActiveAccountDetails.accountlabel)"} | Select-Object ID).id
			write-verbose "  ProcessID for $($ActiveAccountDetails.ID) - $($ActiveAccountDetails.accountlabel) is $InstanceProcessID"
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ProcessID -Value $InstanceProcessID
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
		if ($LongestAccountLabelLength -ge 28){$LongestAccountLabelLength = 28} #we will limit long display names on the main screen to prevent odd things happening.
		while ($AccountHeaderIndent.length -lt ($LongestAccountLabelLength -13)){#indent the header for batches based on how long the longest account name is
			$AccountHeaderIndent = $AccountHeaderIndent + " "
		}
	}
	if ($Script:ActiveAccountsList.id -ne ""){#if batch feature is enabled add a column to display batch number(s)
		if ($Script:Config.TrackAccountUseTime -eq $true){
			$PlayTimeHeader = "Hours Played   "
		}
		if ($Script:EnableBatchFeature -eq $true){
			$BatchesHeader = ("" + $AccountHeaderIndent + "Batch(es)")
		}
		Write-Host ("  ID   Region   " + $PlayTimeHeader + "Account Label   " + $BatchesHeader) #Header
	}
	else {
		Write-Host "  ID   Account Label"
	}
	$Pattern = "(?<=\()([a-z]+)(?=\.actual\.battle\.net\))" #Regex pattern to pull the region characters out of the window title.
	ForEach ($AccountOption in ($Script:AccountOptionsCSV | Sort-Object -Property @{ #Try sort by number first (needed for 2 digit ID's), then sort by character.
		Expression = {	
			$intValue = [int]::TryParse($_.ID, [ref]$null) # Try to convert the value to an integer
			if ($intValue){# If it's not null then it's a number, so return it as an integer for sorting.
				[int]$_.ID
			}
			else {# If it's not a number, return a character and sort that.
				[char]$_.ID
			}
		}
	}))
	{
		if ($AccountOption.accountlabel.length -gt 28){ #later we ensure that strings longer than 28 chars are cut short so they don't disrupt the display.
			$AccountOption.accountlabel = $AccountOption.accountlabel.Substring(0, 28)
		}
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
				Write-Debug "Account not opened yet."
			}
			if ($AcctPlayTime.length -lt 15){#formatting. Depending on the amount of characters for this variable push it out until it's 15 chars long.
				while ($AcctPlayTime.length -lt 15){
					$AcctPlayTime = " " + $AcctPlayTime
				}
			}
		}
		if ($AccountOption.id -in $Script:ActiveAccountsList.id){ #if account is currently active
			$Windowname = (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -like ($AccountOption.id + "*Diablo II: Resurrected")} | Select-Object MainWindowTitle).mainwindowtitle #Check active game instances to see which accounts are active. As this is based on checking window titles, this will only work for accounts opened from the script
			$CurrentRegion = [regex]::Match($WindowName, $Pattern).value #Check which region aka realm the active account is connected to.
			if ($CurrentRegion -eq "US"){$CurrentRegion = "NA"; $RegionDisplayPreIndent = " "; $RegionDisplayPostIndent = " "}
			if ($CurrentRegion -eq "KR"){$CurrentRegion = "Asia"}
			if ($CurrentRegion -eq "EU"){$CurrentRegion = "EU"; $RegionDisplayPreIndent = " "; $RegionDisplayPostIndent = " "}
			Write-Host ("  " + $IDIndent + $AccountOption.ID + "    "  + $RegionDisplayPreIndent + $CurrentRegion + $RegionDisplayPostIndent + "    " + $AcctPlayTime  + $AccountOption.accountlabel + " - Account Active.") -foregroundcolor yellow
		}
		else {#if account isn't currently active
			Write-Host ("  " + $IDIndent + $AccountOption.ID + "      -     " + $AcctPlayTime + $AccountOption.accountlabel + "  " + $AccountIndent + $Batches) -foregroundcolor green
		}
	}
}
Function Menu {
	Clear-Host
	if ($Script:ScriptHasBeenRun -eq $true){
		$Script:AccountUsername = $Null
		if ($DebugMode -eq $true){
			DisplayPreviousAccountOpened
		}
	}
	Else {
		Write-Host ("  You have quite a treasure there in that Horadric multibox script v" + $Currentversion)
	}
	Notifications -check $True
	BannerLogo
	QuoteRoll
	if ($Null -eq $Batch -and $Script:OpenAllAccounts -ne $True){#go through normal account selection screen if script hasn't been launched with parameters that already determine this.
		ChooseAccount
	}
	Else {
		CheckActiveAccounts
		$Script:PWmanualset = $False
		$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
		$Script:TwoDigitIDsUsed = $False
		$Script:TwoDigitBatchesUsed = $False
		ForEach ($AccountOption in $Script:AccountOptionsCSV){
			if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
				$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
				if ($AccountOption.id.length -eq 2){
					$Script:TwoDigitIDsUsed = $True
				}
			}
		}
	}
	if ($Null -ne $Batch -or $Script:OpenBatches -eq $true){#if batch has been passed through parameter or if batch was been selected from the menu.
		$Script:AcceptableBatchIDs = $Null #reset value
		$AcceptableBatchValues = $Null #reset value
		ForEach ($ID in $Script:AccountOptionsCSV){
			if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
				$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',') #collate acceptable options of batch ID's
				$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',') #collate acceptable options of account ID's
			}
		}
		$AcceptableBatchValues = @($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort-Object) #Unique list of available batches that can be opened. @ converts this from a PSObject into an array which fixes the issue of -notin not working on PSobjects with only 1 item.
		ForEach ($BatchValue in $AcceptableBatchValues){
			if ($BatchValue.length -eq 2){
				$Script:TwoDigitBatchesUsed = $True
			}
		}
		do {
			if ($Null -ne $Batch -and $Batch -notin $AcceptableBatchValues){#if batch specified in the parameter isn't valid
				$Script:BatchToOpen = $Batch
				$Batch = $Null
				DisplayActiveAccounts
				Write-Host "`n Batch specified in Parameter is either incorrect or all accounts in that" -foregroundcolor Yellow
				Write-Host " batch are already open. Adjust your parameter or manually specify below.`n" -foregroundcolor Yellow
				start-sleep -milliseconds 5000
				exit
			}
			if ($Null -ne $Batch -and $Batch -in $AcceptableBatchValues){#if batch is valid, set variable so that loop can be exited.
				$Script:BatchToOpen = $Batch
			}
			ElseIf ($AcceptableBatchValues.count -eq 1){
				Write-Host " Opening Batch $X[38;2;255;165;000;22m$AcceptableBatchValues$X[0m.`n"
				$Script:BatchToOpen = $AcceptableBatchValues[0]
			}
			Else {
				Write-Host " Which Batch of accounts would you like to open (" -nonewline
				CommaSeparatedList $AcceptableBatchValues #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
				Write-Host ")?"
				if ($Null -eq $Batch){
					Write-Host " Alternatively, press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				}
				if ($Script:TwoDigitBatchesUsed -eq $True){
					$Script:BatchToOpen = ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). If no button is pressed, send "c" for cancel.
				}
				else {
					$Script:BatchToOpen = ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). If no button is pressed, send "c" for cancel.
				}
				Write-Host
			}
			if ($Script:BatchToOpen -notin $AcceptableBatchValues + "c" + "Esc"){
				Write-Host " Invalid Input. Please enter one of the options above.`n" -foregroundcolor red
				$Script:BatchToOpen = ""
			}
		} until ($Script:BatchToOpen -in $AcceptableBatchValues + "c" + "Esc")
		if ($BatchToOpen -ne "c" -and $BatchToOpen -ne "Esc"){
			$Script:BatchedAccountIDsToOpen = New-Object -TypeName System.Collections.ArrayList
			ForEach ($ID in $Script:AccountOptionsCSV){
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
		$Script:Region = ""
		$Script:RegionOption = ""
	}
	if ($Script:BatchToOpen -ne "c" -and $Script:BatchToOpen -ne "Esc"){#get next region unless the cancel option has been specified.
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
	$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv" #Import accounts.csv again in case someone has updated Auth Method without closing the script.
	if ($Script:OpenAllAccounts -eq $True){
		Write-Host "`n Opening all accounts..."
		ForEach ($ID in $Script:AcceptableValues){
			$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $ID}
			$Script:AccountID = $ID
			if ($id -eq $Script:AcceptableValues[-1]){
				$Script:LastAccount = $true
			}
			Write-Host "`n Opening Account: $ID"
			Processing
		}
		$Script:LastAccount = $False
		$Script:OpenAllAccounts = $False
		if ($Script:ParamsUsed -ne $True){
			Menu
		}
	}
	Else {
		if ($Script:OpenBatches -eq $True -and $Script:RegionOption -ne "c" -and $Script:RegionOption -ne "Esc"){
			if ($Script:BatchedAccountIDsToOpen.count -gt 1){
				$BatchPlural = "s"
			}
			Write-Host "`n Opening account$BatchPlural " -nonewline 
			CommaSeparatedList $Script:BatchedAccountIDsToOpen -AndText
			Write-Host " from batch $BatchToOpen..."
			ForEach ($ID in $Script:BatchedAccountIDsToOpen){
				$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $ID}
				$Script:AccountID = $ID
				if ($id -eq $Script:BatchedAccountIDsToOpen[-1]){
					$Script:LastAccount = $true
				}
				Write-Host "`n Opening Account: $ID..."
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
			if ($Script:BatchToOpen -ne "c" -and $Script:BatchToOpen -ne "Esc"){
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
	Else {#if no account parameters have been set already
		do {
			if ($Script:AccountID -eq "t"){
				TerrorZone
				$Script:AccountID = "r"
			}
			if ($Script:AccountID -eq "d"){
				DClone -DisableOutput $False -DCloneTrackerSource $Script:Config.DCloneTrackerSource -TagList $Script:Config.DCloneAlarmList
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
			if ($Script:AccountID -eq "o"){ #options menu
				Options
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
				$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
				if ($Script:GemActivated -ne $True){
					$GibberingGemstone = get-random -minimum 0 -maximum  4095
					if ($GibberingGemstone -eq 69 -or $GibberingGemstone -eq 420){#nice
						Write-Host "  Perfect Gem Activated" -ForegroundColor magenta
						Write-Host "`n     OMG!" -foregroundcolor green
						$Script:PGemActivated = $True
						([int]$Script:CurrentStats.PerfectGems) ++
						SetQualityRolls
						Start-Sleep -milliseconds 4567
					}
					else {
						if ($GibberingGemstone -in 16..32){
							CowKingKilled
							$SkipCSVExport = $True
						}
						else {
							Write-Host "  Gem Activated" -ForegroundColor magenta
							([int]$Script:CurrentStats.Gems) ++
						}
					}
					$Script:GemActivated = $True
					SetQualityRolls
					if ($SkipCSVExport -ne $True){
						try {
							$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
						}
						Catch {
							Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
						}
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
				Clear-Host
				if ($Script:ScriptHasBeenRun -eq $true){
					if ($DebugMode -eq $true){
						DisplayPreviousAccountOpened
					}
				}
				Notifications -check $True
				BannerLogo
				QuoteRoll
			}
			CheckActiveAccounts
			DisplayActiveAccounts
			if ($Script:Config.TrackAccountUseTime -eq $True){
				$OpenD2LoaderInstances = Get-WmiObject -Class Win32_Process | Where-Object { $_.name -eq "powershell.exe" -and $_.commandline -match $Script:ScriptFileName} | Select-Object name,processid,creationdate | Sort-Object creationdate -descending
				if ($OpenD2LoaderInstances.length -gt 1){#If there's more than 1 D2loader.ps1 script open, close until there's only 1 open to prevent the time played accumulating too quickly.
					ForEach ($Process in $OpenD2LoaderInstances[1..($OpenD2LoaderInstances.count -1)]){
						Stop-Process -id $Process.processid -force #Closes oldest running d2loader script
					}
				}
				if ($Script:ActiveAccountsList.id.length -ne 0){#if there are active accounts open add to total script time
					#Add time for each account that's open
					$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\Accounts.csv"
					$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date) #work out elapsed time to add to accounts.csv
					ForEach ($AccountID in $Script:ActiveAccountsList.id |Sort-Object){ #$Script:ActiveAccountsList.id
						$AccountToUpdate = $Script:AccountOptionsCSV | Where-Object {$_.ID -eq $accountID}
						if ($AccountToUpdate){
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
						Write-Host "`n  Couldn't update accounts.csv with playtime info." -ForegroundColor Red
						Write-Host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
						start-sleep -milliseconds 1500
						$WriteAcctCSVError = $False
					}
					#Add Time to Total Script Time only if there's an open game.
					$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
					try {
						$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
						try {#get current time from csv and add to it
							$Script:CurrentStats.TotalGameTime = [TimeSpan]::Parse($CurrentStats.TotalGameTime) + $AdditionalTimeSpan
						}
						Catch {#if CSV hasn't been populated with a time yet.
							$Script:CurrentStats.TotalGameTime = $AdditionalTimeSpan
						}
						$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
					}
					Catch {
						Write-Host "`n  Couldn't update Stats.csv with playtime info." -ForegroundColor Red
						Write-Host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
						start-sleep -milliseconds 1500
					}
				}
				$Script:StartTime = Get-Date #restart timer for session time and account time.
			}
			Else {
				$Script:StartTime = Get-Date #restart timer for session time only
			}
			$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
			$Script:TwoDigitIDsUsed = $False
			ForEach ($AccountOption in $Script:AccountOptionsCSV){
				if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
					$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
					if ($AccountOption.id.length -eq 2){
						$Script:TwoDigitIDsUsed = $True
					}
				}
			}
			$accountoptions = ($Script:AcceptableValues -join  ", ").trim()
			#DClone Alarm check
			$GetDCloneFunc = $(Get-Command DClone).Definition
			$GetWebRequestFunc = $(Get-Command WebRequestWithTimeOut).Definition
			if ($Script:Config.DCloneAlarmList -ne ""){ # If DClone alarms should be checked on refresh
				try {
					if ($null -ne $Script:DCloneChangesCSV){
						$Script:DCloneChangesCSV = Receive-Job $Script:DCloneJob
						if ($DebugMode -eq $True){
							formatfunction -indents 1 -iswarning -text $Script:DCloneChangesCSV #debugging
						}
						if ($Script:DCloneChangesCSV -match "true"){#if any of the text contains True
							DCloneVoiceAlarm #Create Voice Alarm
						}
					}
					Else {
						$Script:DCloneChangesCSV = "" # If menu is refreshed too quick
					}
					if ($null -ne $RunOnce){ # Don't try remove job on startup for faster launch.
						Get-Job | Where-Object { $Script:JobIDs -notcontains $_.Id } | Remove-Job -Force
					}
					$Script:DCloneJob = Start-Job -ScriptBlock {
						Invoke-Expression "function Dclone {$using:GetDCloneFunc}"
						Invoke-Expression "function WebRequestWithTimeOut {$Using:GetWebRequestFunc}"
						Dclone -DisableOutput $True -DCloneTrackerSource $Using:Config.DCloneTrackerSource -TagList $Using:Config.DCloneAlarmList -DCloneChanges $using:DCloneChangesCSV -DCloneAlarmLevel $Using:DCloneAlarmLevel
					} #check for dclone status
				}
				catch {
					Write-Host "`n Unable to check for DClone status via $($Script:Config.DCloneTrackerSource)." -Foregroundcolor Red
					Write-Host " Try restarting script or changing the source in config.xml." -Foregroundcolor Red
				}
			}
			if ($Script:MovedWindowLocations -ge 1){ # Tidy up any old jobs created for moving windows.
				Get-Job | Where-Object { $Script:JobIDs -contains $_.Id -and $_.state -ne "Running"} | Remove-Job -Force
				$Script:JobIDs = @()
				$Script:MovedWindowLocations = 0
			}
			do {
				Write-Host
				$Script:OpenAllAccounts = $False
				if ($accountoptions.length -gt 0){#if there are unopened account options available
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
						$BatchMenuText = ""
						if ($accountoptions.length -le 24){ #if so many accounts are available to be used that it's too long and impractical to display all the individual options.
							Write-Host ("  Select which account to sign into: " + "$X[38;2;255;165;000;22m$accountoptions$X[0m" + $AllAccountMenuTextNoBatch)
							Write-Host "  Alternatively choose from the following menu options:"
						}
						Else {
							Write-Host " Enter the ID# of the account you want to sign into."
							Write-Host " Alternatively choose from the following menu options:"
							Write-Host ("  " + $AllAccountMenuText)
						}
					}
					else {
						$Script:BatchToOpen = $Null
						$BatchMenuText = "'$X[38;2;255;165;000;22mb$X[0m' to open a Batch of accounts,"
						$Script:AcceptableBatchIDs = $Null #reset value
						$AcceptableBatchValues = $Null
						ForEach ($ID in $Script:AccountOptionsCSV){
							if ($ID.id -in $Script:AcceptableValues){#Find batch values to choose from based on accounts that aren't already open.
								$AcceptableBatchValues = $AcceptableBatchValues + ($ID.batches).split(',')
								$Script:AcceptableBatchIDs = $Script:AcceptableBatchIDs + ($ID.id).split(',')
							}
						}
						$AcceptableBatchValues = @($AcceptableBatchValues | where-object {$_ -ne ""} | Select-Object -Unique | Sort-Object) #Unique list of available batches that can be opened
						if ($Null -eq $AcceptableBatchValues[0]){
							$BatchOption = ""
							$BatchMenuText = ""
							if ($accountoptions.length -le 24){ #if so many accounts are available to be used that it's too long and impractical to display all the individual options.
								Write-Host ("  Select which account to sign into: " + "$X[38;2;255;165;000;22m$accountoptions$X[0m" + $AllAccountMenuTextNoBatch)
								Write-Host "  Alternatively choose from the following menu options:"
							}
							Else {
								Write-Host " Enter the ID# of the account you want to sign into."
								Write-Host " Alternatively choose from the following menu options:"
								Write-Host ("  " + $AllAccountMenuText + $BatchMenuText)
							}
						}
						Else {
							$BatchOption = "b"
							Write-Host " Enter the ID# of the account you want to sign into."
							Write-Host " Alternatively choose from the following menu options:"
							Write-Host ("  " + $AllAccountMenuText + $BatchMenuText)
						}
					}
				}
				else {#if there aren't any available options, IE all accounts are open
					$AllOption = $Null
					$BatchOption = $Null
					Write-Host " All Accounts are currently open!" -foregroundcolor yellow
				}
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, '$X[38;2;255;165;000;22mt$X[0m' for TZ info, '$X[38;2;255;165;000;22md$X[0m' for DClone status, '$X[38;2;255;165;000;22mj$X[0m' for jokes,"
				if ($Script:Config.ManualSettingSwitcherEnabled -eq $true){
					$ManualSettingSwitcherOption = "s"
					Write-Host "  '$X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22ms$X[0m' to toggle the Manual Setting Switcher, "
					Write-Host "  '$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: "-nonewline
				}
				Else {
					$ManualSettingSwitcherOption = $null
					Write-Host "  '$X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: " -nonewline
				}
				if ($Script:TwoDigitIDsUsed -eq $True){
					$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" -TwoDigitAcctSelection $True #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
				}
				else {
					$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
				}
				if ($Script:AccountID -notin ($Script:AcceptableValues + "x" + "r" + "t" + "d" + "g" + "j" + "i" + "o" + $ManualSettingSwitcherOption + $AllOption + $BatchOption) -and $Null -ne $Script:AccountID){
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
					Write-Host "`n Good day to you partner :)" -foregroundcolor yellow
					Start-Sleep -milliseconds 486
					Exit
				}
				$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
			}
			$Script:RunOnce = $True
		} until ($Script:AccountID -ne "r" -and $Script:AccountID -ne "t" -and $Script:AccountID -ne "d" -and $Script:AccountID -ne "g" -and $Script:AccountID -ne "j" -and $Script:AccountID -ne "s" -and $Script:AccountID -ne "i" -and $Script:AccountID -ne "o")
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
	Write-Host " Choose a region to connect to. Available regions are:"
	Write-Host "  Option  Region  Server Address"
	Write-Host "  ------  ------  --------------"
	ForEach ($Server in $ServerOptions){
		if ($Server.region.length -eq 2){$Regiontablespacing = " "}
		if ($Server.region.length -eq 4){$Regiontablespacing = ""}
		Write-Host ("    " + $Server.option + "      " + $Regiontablespacing + $Server.region + $Regiontablespacing + "   " + $Server.region_server) -foregroundcolor green
	}
	do {
		Write-Host "`n Please select a region: $X[38;2;255;165;000;22m1$X[0m, $X[38;2;255;165;000;22m2$X[0m or $X[38;2;255;165;000;22m3$X[0m"
		Write-Host (" Alternatively select '$X[38;2;255;165;000;22mc$X[0m' to cancel or press enter for the default (" + $Config.DefaultRegion + "-" + ($Script:ServerOptions | Where-Object {$_.option -eq $Config.DefaultRegion}).region + "): ") -nonewline
		$Script:RegionOption = ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 13,27 #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). If no button is pressed, send "c" for cancel.
		if ("" -eq $Script:RegionOption){
			$Script:RegionOption = $Config.DefaultRegion #default to NA
		}
		else {
			$Script:RegionOption = $Script:RegionOption.tostring()
		}
		if ($Script:RegionOption -notin $Script:ServerOptions.option + "c" + "Esc"){
			Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
			$Script:RegionOption = ""
		}
	} until ("" -ne $Script:RegionOption)
	if ($Script:RegionOption -in 1..3 ){# if value is 1,2 or 3 set the region string.
		$Script:region = ($ServerOptions | where-object {$_.option -eq $Script:RegionOption}).region_server
		$Script:RegionLabel = $Script:Region.substring(0,2)
		if ($Script:RegionLabel -eq "US"){$Script:RegionLabel = "NA"}
		$Script:LastRegion = $Script:Region
	}
}
Function Processing {
	if ($Script:RegionOption -ne "c" -and $Script:RegionOption -ne "Esc"){
		if (($Script:PW -eq "" -or $Null -eq $Script:PW) -and $Script:PWmanualset -eq 0){
			$Script:PW = $Script:AccountChoice.PW.tostring()
		}
		if ($Script:AccountChoice.AuthenticationMethod -ne "Token"){
			try {
				if ($Script:ParamsUsed -ne $true -or ($Script:ParamsUsed -eq $true -and ($Script:OpenBatches -eq $True -or $Script:OpenAllAccounts -eq $True))){ # If Params aren't used or if Params are used with either batches or open all accounts.
					$Script:acct = $Script:AccountChoice.acct.tostring()
					if ($Config.ConvertPlainTextSecrets -eq $true -or $PW.Length -gt 200){ # if PW should be converted, update $Script:PW to the converted PW, otherwise leave the variable alone.
						$EncryptedPassword = $PW | ConvertTo-SecureString -ErrorAction Stop -Errorvariable ErrorVar #Try converting password.
						$PWobject = New-Object System.Management.Automation.PsCredential("N/A", $EncryptedPassword)
						$Script:PW = $PWobject.GetNetworkCredential().Password
					}
				}
				else {
					if ($Null -eq $Script:AccountID){
						$Script:acct = $Script:AccountUsername
						$Script:AccountID = "1"
					}
				}
			}
			Catch {
				if ($ErrorVar -match "is not a valid encrypted string"){
					Write-Host " There was an issue with decrypting this password." -foregroundcolor red
					FormatFunction -Text ("Please re-enter your password for " + $Script:AccountChoice.acct + " into accounts.csv and re-run the script.`n") -IsError
					PressTheAnyKey
				}
				else {
					Write-Host " Password for this account is in plain text in accounts.csv." -foregroundcolor red
					Write-Host " Run the script again and your password will be secured." -foregroundcolor red
					Write-Host " If errors persist, re-enter your password in accounts.csv`n" -foregroundcolor red
					PressTheAnyKey
				}
			}
		}
		if ($ParamsUsed -eq $True){
			$Script:RegionLabel = $Script:Region.substring(0,2)
			if ($Script:RegionLabel -eq "US"){$Script:RegionLabel = "NA"}
		}
		if ($Script:AccountChoice.AuthenticationMethod -eq "Token" -or ($Script:ForceAuthToken -eq $True -or $Script:Config.ForceAuthTokenForRegion -match $RegionLabel)){
			if ($Script:AccountChoice.Token.length -gt 200){
				$Script:Token = $Script:AccountChoice.Token.tostring()
				$EncryptedToken = $Script:Token | ConvertTo-SecureString
				$Tokenobject = New-Object System.Management.Automation.PsCredential("N/A", $EncryptedToken)
				$Token = $Tokenobject.GetNetworkCredential().Password
			}
			Else {# if token isn't stored as a secure string due to ConvertPlainTextSecrets being set to false.
				$Token = $Script:AccountChoice.Token
			}
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
		$CustomLaunchArguments = ($Script:AccountChoice.CustomLaunchArguments).replace("`"","").replace("'","") #clean up arguments in case they contain quotes (for folks that have used excel to edit accounts.csv).
		if ($Script:AccountChoice.AuthenticationMethod -eq "Parameter" -and $Script:ForceAuthToken -ne $True -and $Script:Config.ForceAuthTokenForRegion -notmatch $RegionLabel){
			$arguments = (" -username " + $Script:acct + " -password " + $Script:PW + " -address " + $Script:Region + " " + $CustomLaunchArguments).tostring()
		}
		else {
			$arguments = (" -uid osi " + $CustomLaunchArguments).tostring()
		}
		if ($Config.ForceWindowedMode -eq $true){#starting with forced window mode sucks, but someone asked for it.
			$arguments = $arguments + " -w"
		}
		$Script:PW = $Null
		$Script:Token = $Null
		#Switch Settings file to load D2r from.
		if ($Config.SettingSwitcherEnabled -eq $True -and $Script:AskForSettings -ne $True){#if user has enabled the auto settings switcher.
			$SettingsProfilePath = ("C:\Users\" + $Env:UserName + "\Saved Games\Diablo II Resurrected\")
			if ($Script:AccountChoice.CustomLaunchArguments -match "-mod"){
				$pattern = "-mod\s+(\S+)" #pattern to find the first word after -mod
				if ($Script:AccountChoice.CustomLaunchArguments -match $pattern){
					$ModName = $matches[1]
					try {
						Write-Verbose "Trying to get Mod Content..."
						try {
							$Modinfo = ((Get-Content "$($Config.GamePath)\Mods\$ModName\$ModName.mpq\Modinfo.json" -ErrorAction silentlycontinue | ConvertFrom-Json).savepath).Trim("/") 
						}
						catch {
							try {
								$Modinfo = ((Get-Content "$($Config.GamePath)\Mods\$ModName\Modinfo.json" -ErrorAction stop -ErrorVariable ModReadError | ConvertFrom-Json).savepath).Trim("/")
							}
							catch {
								FormatFunction -Text "Using standard settings path. Couldn't find Modinfo.json in '$($Config.GamePath)\Mods\$ModName\$ModName.mpq'" -IsWarning
								start-sleep -milliseconds 1500
							}
						}
						If ($Null -eq $Modinfo){
							Write-Verbose " No Custom Save Path Specified for this mod."
						}
						ElseIf ($Modinfo -ne "../"){
							$SettingsProfilePath += "mods\$Modinfo\"
							if (-not (Test-Path $SettingsProfilePath)){
								Write-Host " Mod Save Folder doesn't exist yet. Creating folder..."
								New-Item -ItemType Directory -Path $SettingsProfilePath -ErrorAction stop | Out-Null
								Write-Host " Created folder: $SettingsProfilePath" -ForegroundColor Green
							}
							Write-Host " Mod: $ModName detected. Using custom path for settings.json." -ForegroundColor Green
							Write-Verbose " $SettingsProfilePath"
						}
						Else {
							Write-Verbose " Mod used but save path is standard."
						}
					}
					Catch {
						Write-Verbose " Mod used but custom save path not specified."
					}
				}
				else {
					Write-Host " Couldn't detect Mod name. Standard path to be used for settings.json." -ForegroundColor Red
				}
			}
			$SettingsJSON = ($SettingsProfilePath + "Settings.json")
			if ((Test-Path -Path ($SettingsProfilePath + "Settings.json")) -eq $true){ #check if settings.json does exist in the savegame path (if it doesn't, this indicates first time launch or use of a new single player mod).
				ForEach ($id in $Script:AccountOptionsCSV){#create a copy of settings.json file per account so user doesn't have to do it themselves
					if ((Test-Path -Path ($SettingsProfilePath + "Settings" + $id.id +".json")) -ne $true){#if somehow settings<ID>.json doesn't exist yet make one from the current settings.json file.
						try {
							Copy-Item $SettingsJSON ($SettingsProfilePath + "Settings"+ $id.id + ".json") -ErrorAction Stop
						}
						catch {
							FormatFunction -Text "`nCouldn't find settings.json in $SettingsProfilePath" -IsError
							if ($Script:AccountChoice.CustomLaunchArguments -match "-mod"){
								Break
							}
							Else {
								Write-Host " Start the game normally (via Bnet client) & this file will be rebuilt." -foregroundcolor red
							}
							Write-Host
							PressTheAnyKeyToExit
						}
					}
				}
				try {
					Copy-item ($SettingsProfilePath + "settings"+ $Script:AccountID + ".json") $SettingsJSON -ErrorAction Stop #overwrite settings.json with settings<ID>.json (<ID> being the account ID). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
					$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).accountlabel
					formatfunction -text ("Custom game settings (settings" + $Script:AccountID + ".json) being used for " + $CurrentLabel) -success
					Start-Sleep -milliseconds 133
				}
				catch {
					FormatFunction -Text "Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -IsError
					PressTheAnyKey
				}
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
					Write-Host "`n Couldn't find settings.json in $SettingsProfilePath" -foregroundcolor red
					Write-Host " Please start the game normally (via Bnet client) & this file will be rebuilt." -foregroundcolor red
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
			ForEach ($file in $files){
				 $SettingsFileOption = New-Object -TypeName psobject
				 $Counter = $Counter + 1
				 $Name = $file.Name -replace '^settings\.|\.json$' #remove 'settings.' and '.json'. The text in between the two periods is the name.
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "Name" -Value $Name
				 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
				 [VOID]$SettingsFileOptions.Add($SettingsFileOption)
			}
			if ($Null -ne $SettingsFileOptions){# If settings files are found, IE the end user has set them up prior to running script.
				$SettingsFileOptions = $SettingsDefaultOptionArray + $SettingsFileOptions
				Write-Host "  Settings options you can choose from are:"
				ForEach ($Option in $SettingsFileOptions){
					Write-Host ("   " + $Option.ID + ". " + $Option.name) -foregroundcolor green
				}
				do {
					Write-Host "  Choose the settings file you like to load from: " -nonewline
					ForEach ($Value in $SettingsFileOptions.ID){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
						if ($Value -ne $SettingsFileOptions.ID[-1]){
							Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
							if ($Value -ne $SettingsFileOptions.ID[-2]){Write-Host ", " -nonewline}
						}
						else {
							Write-Host " or $X[38;2;255;165;000;22m$Value$X[0m"
						}
					}
					if ($Null -eq $ManualSettingSwitcher){#if not launched from parameters
						Write-Host "  Or Press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
						$SettingsCancelOption = "c"
					}
					$SettingsChoice = ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). If no button is pressed, send "c" for cancel.
					if ($SettingsChoice -eq ""){
						$SettingsChoice = 1
					}
					Write-Host
					if ($SettingsChoice.tostring() -notin $SettingsFileOptions.id + $SettingsCancelOption + "Esc"){
						Write-Host "  Invalid Input. Please enter one of the options above." -foregroundcolor red
						$SettingsChoice = ""
					}
				} until ($SettingsChoice.tostring() -in $SettingsFileOptions.id + $SettingsCancelOption + "Esc")
				if ($SettingsChoice -ne "c" -and $SettingsChoice -ne "Esc"){
					$SettingsToLoadFrom = $SettingsFileOptions | where-object {$_.id -eq $SettingsChoice.tostring()}
					try {
						Copy-item ($SettingsProfilePath + $SettingsToLoadFrom.FileName) -Destination $SettingsJSON #-ErrorAction Stop #overwrite settings.json with settings<Name>.json (<Name> being the name of the config user selects). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
						$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).accountlabel
						Write-Host (" Custom game settings (" + $SettingsToLoadFrom.Name + ") being used for " + $CurrentLabel + "`n") -foregroundcolor green
						Start-Sleep -milliseconds 100
					}
					catch {
						FormatFunction -Text "Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -IsError
						PressTheAnyKey
					}
				}
			}
			Else {# if no custom settings files are found, IE user hasn't set them up yet.
				Write-Host "`n  No Custom Settings files have been saved yet. Loading default settings." -foregroundcolor Yellow
				Write-Host "  See README for setup instructions.`n" -foregroundcolor Yellow
				PressTheAnyKey
			}
		}
		if ($SettingsChoice -ne "c" -and $SettingsChoice -ne "Esc"){
			#Start Game
			KillHandle | out-null
			$process = Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments" -PassThru
			Start-Sleep -milliseconds 1500 #give D2r a bit of a chance to start up before trying to kill handle
			#Close the 'Check for other instances' handle
			Write-Host " Attempting to close `"Check for other instances`" handle..."
			$Output = KillHandle | out-string #run KillHandle function.
			if (($Output.contains("DiabloII Check For Other Instances")) -eq $true){
				$handlekilled = $true
				Write-Host " `"Check for Other Instances`" Handle closed." -foregroundcolor green
			}
			else {
				Write-Host " `"Check for Other Instances`" Handle was NOT closed." -foregroundcolor red
				Write-Host " Who even knows what happened. I sure don't." -foregroundcolor red
				FormatFunction -text " If you are seeing this error and are running the script for the first time`n" -IsError
				PressTheAnyKey
			}
			if ($handlekilled -ne $True){
				Write-Host " Couldn't find any handles to kill." -foregroundcolor red
				Write-Host " Game may not have launched as expected." -foregroundcolor red
				PressTheAnyKey
			}
			#Rename the Diablo Game window for easier identification of which account and region the game is.
			$rename = ($Script:AccountID + " - " + $Script:AccountFriendlyName + " (" + $Script:Region + ")" +" - Diablo II: Resurrected")
			$Command = ('"'+ $WorkingDirectory + '\SetText\SetTextv2.exe" /PID ' + $process.id + ' "' + $rename + '"')
			try {
				cmd.exe /c $Command
				write-debug $Command #debug
				Write-debug " Window Renamed." #debug
				Start-Sleep -milliseconds 250
			}
			catch {
				Write-Host " Couldn't rename window :(" -foregroundcolor red
				PressTheAnyKey
			}
			If ($Script:Config.RememberWindowLocations -eq $True){ #If user has enabled the feature to automatically move game Windows to preferred screen locations.
				if ($Script:AccountChoice.WindowXCoordinates -ne "" -and $Script:AccountChoice.WindowYCoordinates -ne "" -and $Null -ne $Script:AccountChoice.WindowXCoordinates -and $Null -ne $Script:AccountChoice.WindowYCoordinates -and $Script:AccountChoice.WindowWidth -ne "" -and $Script:AccountChoice.WindowHeight -ne "" -and $Null -ne $Script:AccountChoice.WindowWidth -and $Null -ne $Script:AccountChoice.WindowHeight){ #Check if the account has had coordinates saved yet.
					#GetAddWindowTypeFunc = $(Get-Command WindowMover).Definition
					$GetSetWindowLocationsFunc = $(Get-Command SetWindowLocations).Definition
					$JobID = (Start-Job -ScriptBlock { # Run this in a background job so we don't have to wait for it to complete
						start-sleep -milliseconds 2024 # We need to wait for about 2 seconds for game to load as if we move it too early, the game itself will reposition the window. Absolute minimum is 420 milliseconds (funnily enough). Delay may need to be a bit higher for people with wooden computers.
						#Invoke-Expression "function WindowMover {$using:GetAddWindowTypeFunc}"
						. "$Using:WorkingDirectory\WindowMover.ps1"
						Invoke-Expression "function SetWindowLocations {$using:GetSetWindowLocationsFunc}"
						SetWindowLocations -x $Using:AccountChoice.WindowXCoordinates -y $Using:AccountChoice.WindowYCoordinates -Width $Using:AccountChoice.WindowWidth -Height $Using:AccountChoice.WindowHeight -Id $Using:process.id
					})
					$Script:MovedWindowLocations ++
					$Script:JobIDs += $JobID.id
					start-sleep -milliseconds 2024
				}
				Else { #Show a warning if user has RememberWindowLocations but hasn't configured it for this account yet.
					FormatFunction -iswarning -text "`n'RememberWindowLocations' config is enabled but can't move game window to preferred location as coordinates need to be defined for the account first.`n`nTo setup follow the quick steps below:"
					FormatFunction -iswarning -indents 1 -SubsequentLineIndents 3 -text "1. Open all of your D2r account instances.`n2. Move the window for each game instance to your preferred layout."
					FormatFunction -iswarning -indents 1 -SubsequentLineIndents 3 -text "3. Go to the options menu in the script and go into the 'RememberWindowLocations' setting.`n4. Once in this menu, choose the option 's' to save coordinates of any open game instances."
					FormatFunction -iswarning -text  "`nNow when you open these accounts they will open in this screen location each time :)`n"
					PressTheAnyKey
				}
			}
			if ($Script:AccountChoice.AuthenticationMethod -eq "Token" -or (($Script:ForceAuthToken -eq $True -or $Script:Config.ForceAuthTokenForRegion -match $RegionLabel) -and $Script:AccountChoice.Token.length -ge 200)){#wait for web_token to change twice (once for launch, once for char select screen, before being able to launch additional accounts. Token will have already changed once by the time script reaches this stage
				$CurrentTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
				Write-Host " Launched Game using an $X[38;2;165;146;99;4mAuthentication Token$X[0m."
				Write-Host " Waiting for you to get to character select screen..." -foregroundcolor yellow
				Write-Host " $X[38;2;255;255;0;4mDO NOT OPEN OR CLOSE ANOTHER GAME INSTANCE UNTIL YOU'VE DONE THIS.$X[0m"
				do {
					$NewTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
					$CompareCheck = Compare-Object $CurrentTokenRegValue $NewTokenRegValue
					if ($Null -ne $CompareCheck){#if CompareCheck has some value, this means it found differences, IE the reg value changed.
						$CurrentTokenRegValue = (Get-ItemProperty -Path $Path -Name WEB_TOKEN).WEB_TOKEN
						$CurrentTokenRegValue = $NewTokenRegValue
						$WebTokenChangeCounter++
						Write-Host " Game Launch Successful for $Script:AccountFriendlyName!" -foregroundcolor green
					}
					else {
						Start-Sleep -milliseconds 451
					}
				} until ($WebTokenChangeCounter -eq 1)
			}
			if ($Script:LastAccount -eq $True -or ($Script:OpenAllAccounts -ne $True -and $Script:OpenBatches -ne $True)){
				if ($Script:MovedWindowLocations -ge 1){
					if ($Script:MovedWindowLocations -gt 1){
						$plural = "s"
					}
					FormatFunction -IsSuccess -text "Moved game window$plural to preferred location$plural."
					Start-Sleep -milliseconds 750
				}
				Write-Host "`nGood luck hero..." -foregroundcolor magenta
			}
			Start-Sleep -milliseconds 1000
			$Script:ScriptHasBeenRun = $true
		}
	}
}
InitialiseCurrentStats
CheckForUpdates
ImportXML
ValidationAndSetup
ImportCSV
Clear-Host
D2rLevels
QuoteList
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
