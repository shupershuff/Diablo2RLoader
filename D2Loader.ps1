<# 
Author: Shupershuff
Version: See Github https://github.com/shupershuff/Diablo2RLoader
Last Edited: See Github https://github.com/shupershuff/Diablo2RLoader
Usage: Go nuts.
Purpose:
	Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle."
	Script will import account details from CSV. Alternatively you can run script with account, region and password parameters.
Pre-Requisites:
	1. Download the handle tool https://learn.microsoft.com/en-gb/sysinternals/downloads/handle. 
	2. Place the executables in a folder called "handle" in the same place this script lives. (ie folder path should be "./handle/")
#################
# Instructions: #
#################
See 
 1. Populate csv file "accounts.csv" with your account details. Headings for this CSV are ID,acct,pw,accountlabel
	- ID is simply a number, number these 1 to 5 if you have 5 accounts.
	- acct is your bnet email sign in address
	- pw is your bnet account password.
	- accountlabel is simply a friendly name you can add so you can tell which window is a particular account. I prefer to enter in my bnet usernames (without the #1234 at the end).
 2. Set any Script Options if required (Check Game Path is accurate, other options are...optional).
 3. ????
 4. ????
 5. Profit.

#########
# Notes #
#########
Multiple failed attempts (eg wrong Password) to sign onto a particular Realm via this method may temporarily lock you out. You should still be able to get in via the battlenet client if this occurs.
Handle script stolen from https://forums.d2jsp.org/topic.php?t=90563264&f=87
Servers:
 NA - us.actual.battle.net
 EU - eu.actual.battle.net
 Asia - kr.actual.battle.net
#>
param($AccountUsername,$PW,$region) #used to capture paramters sent to the script, if anyone even wants to do that.

###########################################################################################################################################
# Script Options
###########################################################################################################################################
#Adjust the below to suit your setup and preferences :)
$GamePath = "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected"
$DefaultRegion = 1 #default region, 1 for NA, 2 for EU, 3 for Asia.
$AskForRegionOnceOnly = $false
$CreateDesktopShortcut = $True #Script will recreate desktop shortcut each time it's run (updates the shortcut if you move the script location). If you don't want this, disable here by setting to $false.


###########################################################################################################################################
# Script itself
###########################################################################################################################################
$host.ui.RawUI.WindowTitle = "Diablo 2 Resurrected Loader"
if ($null -ne $AccountUsername){
	$scriptarguments = "-accountusername $AccountUsername"  
}
if ($null -ne $pw){
	$scriptarguments += " -pw $pw"  
}
if ($null -ne $region){
	$scriptarguments += " -region $region"
}

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $ScriptArguments "  -Verb RunAs;exit } #run script as admin
$script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))

#set window size
[console]::WindowWidth=77;
[console]::WindowHeight=52;
[console]::BufferWidth=[console]::WindowWidth

#Check SetText.exe setup
Add-MpPreference -ExclusionPath "$script:WorkingDirectory\SetText" #Add Defender Exclusion for SetText.exe directory, at least until Microsoft reviews the file (submitted 24.4.2023) and stops flagging as "Trojan:Win32/Wacatac.B!ml" as per https://github.com/hankhank10/false-positive-malware-reporting
if((Test-Path -Path ($workingdirectory + '\SetText\SetText.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	Write-Host
	Write-Host "First Time run!" -foregroundcolor Yellow
	Write-Host
	write-host "SetText.exe not in .\SetText\ folder and needs to be built."
	if((Test-Path -Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe") -ne $True){#check that .net4.0 is actually installed or compile will fail.
		write-host ".Net v4.0 not installed. This is required to compile the Window Renamer for Diablo." -foregroundcolor red
		write-host "Download and install it from Microsoft here:" -foregroundcolor red
		write-host "https://dotnet.microsoft.com/en-us/download/dotnet-framework/net40" #actual download link https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net40-web-installer
		pause
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
	start-sleep -milliseconds 4000 #a small delay so the first time run outputs can briefly be seen
}

#Check Handle64.exe downloaded and placed into correct folder
$script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))
if((Test-Path -Path ($workingdirectory + '\Handle\Handle64.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
	write-host "Handle64.exe is in the .\Handle\ folder. See instructions for more details on setting this up." -foregroundcolor red
	pause
	exit
}
#Check Windows Game Path for D2r.exe is accurate.
if((Test-Path -Path "$GamePath\d2r.exe") -ne $True){ 
	write-host "Gamepath is incorrect. Looks like you have a custom D2r install location! Edit the $GamePath variable in the script" -foregroundcolor red
	pause
	exit
}

# Create Shortcut
if ($CreateDesktopShortcut -eq $True){
	$DesktopPath = [Environment]::GetFolderPath("Desktop")
	$Targetfile = "-File `"$WorkingDirectory\D2Loader.ps1`""
	$shortcutFile = "$DesktopPath\D2R Loader.lnk"
	$WScriptShell = New-Object -ComObject WScript.Shell
	$shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
	$shortcut.TargetPath = "powershell.exe" 
	$Shortcut.Arguments = $TargetFile
	$shortcut.IconLocation = "$Script:GamePath\D2R.exe" 
	$shortcut.Save()
}
#check if username was passed through via parameter
if ($AccountUsername -ne $null){
	$script:ParamsUsed = $true
}
Else {
	$script:ParamsUsed = $false
}

#Import CSV
if ($script:AccountUsername -eq $null){#If no parameters sent to script.
	try {
		$Script:AccountOptionsCSV = import-csv "$script:WorkingDirectory\Accounts.csv" #import all accounts from csv
	}
	Catch {
		write-host
		write-host " Accounts.csv does not exist. Make sure you create this and populate with accounts first." -foregroundcolor red
		write-host " Script exiting..." -foregroundcolor red
		start-sleep 5
		Exit
	}
}
#Set Region Array
$Script:ServerOptions = @(
	[pscustomobject]@{Option='1';region='NA';region_server='us.actual.battle.net'}#Americas
	[pscustomobject]@{Option='2';region='EU';region_server='eu.actual.battle.net'}#Europe
	[pscustomobject]@{Option='3';region='Asia';region_server='kr.actual.battle.net'}
)
#Set item quality array for randomizing quote colours. A stupid addition to script but meh.
$QualityArray = @(#quality and chances for things to drop based on 0MF values in D2r (I think?)
	[pscustomobject]@{Type='Unique';Probability=25}
	[pscustomobject]@{Type='SetItem';Probability=62}
	[pscustomobject]@{Type='Rare';Probability=100}
	[pscustomobject]@{Type='Magic';Probability=294}
	[pscustomobject]@{Type='Normal';Probability=9518}
)

$QualityHash = @{}; 
foreach ($object in $QualityArray | select-object type,probability){#convert PSOobjects to hashtable for enumerator
	$QualityHash.add($object.type,$object.probability) #add each PSObject to hash
}
$script:itemLookup = foreach ($entry in $QualityHash.GetEnumerator()){
	[System.Linq.Enumerable]::Repeat($entry.Key, $entry.Value)
}
$x = [char]0x1b #escape character for ANSI text colors
function Magic {#text colour formatting for "magic" quotes
    process { write-host "  $x[38;2;65;105;225;48;2;1;1;1;4m$_$x[0m" }
}
function SetItem {
    process { write-host "  $x[38;2;0;225;0;48;2;1;1;1;4m$_$x[0m"}
}
function Unique {
    process { write-host "  $x[38;2;165;146;99;48;2;1;1;1;4m$_$x[0m"}
}
function Rare {
    process { write-host "  $x[38;2;255;255;0;48;2;1;1;1;4m$_$x[0m"}
}
function Normal {
    process { write-host "  $x[38;2;255;255;255;48;2;1;1;1;4m$_$x[0m"}
}

function quoteroll {#stupid thing to draw a random quote but also draw a random quality.
	$quality = get-random $itemlookup
	Write-Host
	#Write-output (Get-Random -inputobject $script:quotelist)  | &$quality
	$LeQuote = (Get-Random -inputobject $script:quotelist)  #| &$quality
	$consoleWidth = $Host.UI.RawUI.BufferSize.Width
	$desiredIndent = 2  # indent spaces
	$chunkSize = $consoleWidth - $desiredIndent
	[RegEx]::Matches($LeQuote, ".{$chunkSize}|.+").Groups.Value | ForEach-Object {
		write-output $_ | &$quality
	}
	Write-Host
}

$script:quotelist =
"Stay a while and listen..",
"Destruction rains upon you.",
"My brothers will not have died in vain!",
"Do you fear death, nephalem? The power of the lord of terror is mine.",
"Not even death can save you from me.",
"Good Day!",
"You have quite a treasure there in that Horadric Cube.",
"There's nothing the right potion can't cure.",
"Well, what the hell do you want? Oh, it's you. Uh, hi there.",
"What do you need?",
"Your presence honors me.",
"I'll put that to good use.",
"Good to see you!",
"Looking for Baal?",
"All who oppose me, beware",
"Greetings",
"Ner. Ner! Nur. Roah. Hork, Hork.",
"I shall make weapons from your bones",
"I am overburdened",
"This magic ring does me no good.",
"Beware, foul demons and beasts.",
"They'll never see me coming.",
"I will cleanse this wilderness.",
"I shall purge this land of the shadow.",
"I hear foul creatures about.",
"Ahh yes, ruins, the fate of all cities.",
"I have no grief for him. Oblivion is his reward.",
"The catapults have been silenced.",
"The staff of kings, you astound me",
"When - or if - I get to Lut Gholein, I'm going to find the largest bowl of Narlant weed and smoke 'til all earthly sense has left my body.",
"I've just about had my fill of the walking dead.",
"Oh I hate staining my hands with the blood of foul Sorcerers!",
"Damn it, I wish you people would just leave me alone!",
"Beware! Beyond lies mortal danger for the likes of you!",
"Only the darkest Magics can turn the sun black.",
"You are too late! HAA HAA HAA",
"You now speak to Ormus. He was once a great mage, but now lives like a rat in a sinking vessel",
"I knew there was great potential in you, my friend. You've done a fantastic job.",
"Hi there. I'm Charsi, the Blacksmith here in camp. It's good to see some strong adventurers around here.",
"Whatcha need?",
"Good day to you partner!",
"Moomoo, moo, moo. Moo, Moo Moo Moo Mooo.",
"Moo.",
"Moooooooooooooo",
"Gem Activated",
"Gem Deactivated"

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

Function Killhandle {#kudos the info in this post to save me from figuring it out: https://forums.d2jsp.org/topic.php?t=90563264&f=87
	& "$PSScriptRoot\handle\handle64.exe" -accepteula -a -p D2R.exe > $PSScriptRoot\d2r_handles.txt
	$proc_id_populated = ""
	$handle_id_populated = ""

	foreach($line in Get-Content $PSScriptRoot\d2r_handles.txt) {
		$proc_id = $line | Select-String -Pattern '^D2R.exe pid\: (?<g1>.+) ' | %{$_.Matches.Groups[1].value}
		if ($proc_id){
			$proc_id_populated = $proc_id
		}
		$handle_id = $line | Select-String -Pattern '^(?<g2>.+): Event.*DiabloII Check For Other Instances' | %{$_.Matches.Groups[1].value}
		if ($handle_id){
			$handle_id_populated = $handle_id
		}

		if($handle_id){
			Write-Host "Closing" $proc_id_populated $handle_id_populated
			& "$PSScriptRoot\handle\handle64.exe" -p $proc_id_populated -c $handle_id_populated -y
		}
	}
}

Function CheckActiveAccounts {#Note: only works for accounts loaded by the script and only if SetText has been setup.
	#check if there's any open instances and check the game title window for which account is being used.
	try {	
		$script:ActiveIDs = $null
		$D2rRunning = $false
		$script:ActiveIDs = New-Object -TypeName System.Collections.ArrayList
		$script:ActiveIDs = (Get-Process | Where {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "- Diablo II: Resurrected -"} | Select-Object MainWindowTitle).mainwindowtitle.substring(0,1) #find all diablo 2 game windows and pull the account ID from the title
		$script:D2rRunning = $true
		#write-host "Running Instances."
	}
	catch {#if the above fails then there are no running D2r instances.
		$script:D2rRunning = $false
		#write-host "No Running Instances."
		$Script:ActiveIDs = ""
	}
	if ($script:D2rRunning -eq $True){
		$Script:ActiveAccountsList = New-Object -TypeName System.Collections.ArrayList
		foreach ($ActiveID in $ActiveIDs){
			$ActiveAccountDetails = $Script:AccountOptionsCSV | where-object {$_.id -eq $ActiveID}
			$ActiveAccount = New-Object -TypeName psobject
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ID -Value $ActiveAccountDetails.ID
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name AccountName -Value $ActiveAccountDetails.accountlabel
			#$ActiveAccount | Add-Member -MemberType NoteProperty -Name SignIn -Value $ActiveAccountDetails.acct
			[VOID]$Script:ActiveAccountsList.Add($ActiveAccount)
		}
	}
	else {
		$Script:ActiveAccountsList = $null
	}
}
Function DisplayActiveAccounts {
	write-host 
	write-host "ID  Account Label"
	foreach ($AccountOption in $Script:AccountOptionsCSV){
		if ($AccountOption.id -in $Script:ActiveAccountsList.id){
			write-host ($AccountOption.ID + "   " + $AccountOption.accountlabel + " - Account Currently In Use.") -foregroundcolor yellow
		}
		else {
			write-host ($AccountOption.ID + "   " + $AccountOption.accountlabel) -foregroundcolor green
		}
	}
}
Function Menu {
	cls
	if($Script:ScriptHasBeenRun -eq $true){
		$script:AccountUsername = $null
		Write-Host "Account previously opened was:"  -foregroundcolor yellow -backgroundcolor darkgreen
		$lastopened = @(
			[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$script:region}#Americas
		)
		write-host " " -NoNewLine
		Write-Host ("Account:  " + $lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen 
		write-host " " -NoNewLine
		Write-Host "Region:  " $lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
	}
	Else {
		Write-Host "You have quite a treasure there in that Horadric multibox script"
		Write-Host
	}

	Write-Host $BannerLogo -foregroundcolor yellow
	QuoteRoll
	ChooseAccount
	if($script:ParamsUsed -eq $false -and ($Script:RegionOption.length -ne 0 -or $Script:Region.length -ne 0)){
		if ($Script:AskForRegionOnceOnly -ne $true){
			$Script:RegionOption = ""
			$Script:Region = ""
			#write-host "region reset" -foregroundcolor yellow #debug
		}
	}
	if ($script:region.length -eq 0){#if no region parameter has been set already.
		ChooseRegion
	}
	Else {#if region parameter has been set already.
		if ($script:region -ne "us.actual.battle.net" -and $script:region -ne "eu.actual.battle.net" -and $script:region -ne "kr.actual.battle.net"){
			Write-host "Region not valid. Please choose region" -foregroundcolor red
			ChooseRegion
		}
	}
	Processing
}
Function ChooseAccount {
	if ($null -ne $script:AccountUsername){ #if parameters have already been set.
		$Script:AccountOptionsCSV = @(
			[pscustomobject]@{pw=$script:PW;acct=$script:AccountUsername}
		)
	}
	else {#if no account parameters have been set already		
		do {
			if ($Script:AccountID -eq "r"){#refresh
				cls
				if($Script:ScriptHasBeenRun -eq $true){
					Write-Host "Account previously opened was:"  -foregroundcolor yellow -backgroundcolor darkgreen
					$lastopened = @(
						[pscustomobject]@{Account=$Script:AccountFriendlyName;region=$script:region}#Americas
					)
					Write-Host " " -NoNewLine
					Write-Host ("Account:  " + $lastopened.Account) -foregroundcolor yellow -backgroundcolor darkgreen 
					write-host " " -NoNewLine
					Write-Host "Region:  " $lastopened.Region -foregroundcolor yellow -backgroundcolor darkgreen
				}
				Write-Host $BannerLogo -foregroundcolor yellow
				QuoteRoll	
			}
			CheckActiveAccounts
			DisplayActiveAccounts
			$AcceptableValues = New-Object -TypeName System.Collections.ArrayList
			foreach ($AccountOption in $Script:AccountOptionsCSV){
				if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
					$AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"	
				}
			}
			$accountoptions = ($AcceptableValues -join  ", ").trim()
			do {
				Write-Host
				if($accountoptions.length -gt 0){
					Write-Host "Please select which account to sign into."
					$Script:AccountID = Read-host ("Your Options are: " + $accountoptions + ", r to refresh, or x to quit.")
				}
				else {#if there aren't any available options, IE all accounts are open
					Write-Host "All Accounts are currently open!" -foregroundcolor yellow
					$Script:AccountID = Read-host "Press r to refresh, or x to quit."
				}
				if($Script:AccountID -notin ($AcceptableValues + "x" + "r") -and $null -ne $Script:AccountID){
					Write-host "Invalid Input. Please enter one of the options above." -foregroundcolor red
					$Script:AccountID = $Null
				}
			} until ($Null -ne $Script:AccountID)

			if ($Null -ne $Script:AccountID){
				if($Script:AccountID -eq "x"){
					Exit
				}
				$Script:AccountChoice = $Script:AccountOptionsCSV |where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
			}
		} until ($Script:AccountID -ne "r")
	}
	if (($null -ne $script:AccountUsername -and ($null -eq $script:PW -or "" -eq $script:PW) -or ($Script:AccountChoice.id.length -gt 0 -and $Script:AccountChoice.pw.length -eq 0))){
		$script:PW = read-host "Enter the Battle.net password for $script:AccountUsername"
		$script:pwmanualset = $true
	}
	else {
		$script:pwmanualset = $false
	}
}
Function ChooseRegion {#AKA Realm. Not to be confused with the actual Diablo servers that host your games, which are all over the world :)
param	([string] $script:region)
	write-host
	write-host "Available regions are:"
	write-host "Option Region Server Address"
	write-host "------ ------ --------------"
	foreach ($server in $ServerOptions){
		if ($server.region.length -eq 2){$regiontablespacing = "  "}
		if ($server.region.length -eq 4){$regiontablespacing = ""}
		write-host ($server.option + "      " + $server.region + $regiontablespacing + "   " + $server.region_server)
	}
	write-host	
		do {
			$Script:RegionOption = Read-host ("Please select a region (1, 2 or 3) or press enter for the default (" + $script:defaultregion + ": " +($Script:ServerOptions | Where-Object {$_.option -eq $script:defaultregion}).region + ")")
			if ("" -eq $Script:RegionOption){
				$Script:RegionOption = $Script:DefaultRegion #default to NA
			}
			if($Script:RegionOption -notin $Script:ServerOptions.option){
				Write-host "Invalid Input. Please enter one of the options above." -foregroundcolor red
				$Script:RegionOption = ""
			}
		} until ("" -ne $Script:RegionOption)
	if ($Script:RegionOption -eq 1 ){$script:region = ($ServerOptions | where-object {$_.option -eq $Script:RegionOption}).region_server}
	if ($Script:RegionOption -eq 2 ){$script:region = ($ServerOptions | where-object {$_.option -eq $Script:RegionOption}).region_server}
	if ($Script:RegionOption -eq 3 ){$script:region = ($ServerOptions | where-object {$_.option -eq $Script:RegionOption}).region_server}
}

Function Processing {
	if (($script:pw -eq "" -or $script:pw -eq $null) -and $script:pwmanualset -eq 0 ){
		$script:pw = $Script:AccountChoice.pw.tostring()
	}
	if ($script:ParamsUsed -ne $true){
		$script:acct = $Script:AccountChoice.acct.tostring()
	}
	else {
		$script:acct = $script:AccountUsername
		$Script:AccountID = "1"
	}
	$script:region = $script:region.tostring()
	try {
		$Script:AccountFriendlyName = $Script:AccountChoice.accountlabel.tostring()
	}
	Catch {
		$Script:AccountFriendlyName = $script:AccountUsername
	}	

	#Open diablo with parameters
		# IE, this is essentially just opening D2r like you would with a shortcut target of "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected\D2R.exe" -username <yourusername -password <yourPW> -address <SERVERaddress>
	$arguments = (" -username " + $script:acct + " -password " + $script:PW +" -address " + $Script:Region).tostring()
	#write-output $arguments #debug
	Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments"
	start-sleep -milliseconds 1500

	#Close the 'Check for other instances' handle
	Write-host "Attempting to close `"Check for other instances`" handle..."
	$maxattempts = 5 #try closing 20 times.
	do {#wait for d2r process to start
		$output = killhandle | out-string
		start-sleep -milliseconds 1000
		if(($output.contains("DiabloII Check For Other Instances")) -eq $true){
			$handlekilled = $true
			$attempt = $attempt + 1
			write-host ("Attempt " + $attempt + "....")
			write-host "Check for Other Instances Handle closed." -foregroundcolor green
		}	
	} until ($handlekilled -eq $True -or $attempt -eq $maxattempts)
	if ($attempt -eq $maxattempts){
		Write-Host " Couldn't find any handles to kill." -foregroundcolor red
	}
	#Rename the Diablo Game window for easier identification of which account and region the game is.
	$rename = ($Script:AccountID + " - Diablo II: Resurrected - " + $Script:AccountFriendlyName + " (" + $Script:Region + ")")
	$command = ('"'+ $WorkingDirectory + '\SetText\SetText.exe" "Diablo II: Resurrected" "' + $rename +'"')
	try {
		cmd.exe /c $command
		#write-output $command  #testing
		write-host "Window Renamed." -foregroundcolor green
	}
	catch {
		write-host "Couldn't rename window :(" -foregroundcolor red
	}
	start-sleep -milliseconds 200
	$Script:ScriptHasBeenRun = $true
	if ($script:ParamsUsed -ne $true){
		Menu
	}
	else {
		write-host "I'm quitting LOL"
		exit
	}
}
cls
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
