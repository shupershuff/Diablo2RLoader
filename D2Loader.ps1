<# 
Author: Shupershuff
Usage: Go nuts.
Purpose:
	Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle."
	Script will import account details from CSV. Alternatively you can run script with account, region and password parameters.
Instructions: See GitHub readme https://github.com/shupershuff/Diablo2RLoader

#########
# Notes #
#########
Multiple failed attempts (eg wrong Password) to sign onto a particular Realm via this method may temporarily lock you out. You should still be able to get in via the battlenet client if this occurs.

Servers:
 NA - us.actual.battle.net
 EU - eu.actual.battle.net
 Asia - kr.actual.battle.net
#>
param($AccountUsername,$PW,$region) #used to capture paramters sent to the script, if anyone even wants to do that.
$CurrentVersion = "1.1.4"

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
#Check for updates
$tagList = Invoke-RestMethod https://api.github.com/repos/Shupershuff/Diablo2RLoader/tags
if ([version[]]$taglist.Name.Trim('v') -gt $Script:CurrentVersion) {
	Write-Host
	Write-Host " Update available, See Github for latest version:" -foregroundcolor Yellow
	write-host " https://github.com/shupershuff/Diablo2RLoader/releases/latest"
	Write-Host
	start-sleep 3
}

$script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))

#set window size
[console]::WindowWidth=77;
[console]::WindowHeight=48;
[console]::BufferWidth=[console]::WindowWidth

#Import Config XML
try {
	$Script:Config = ([xml](Get-Content "$script:WorkingDirectory\Config.xml")).D2loaderconfig
	#Write-host "Config imported successfully." -foregroundcolor green
}
Catch {
	write-host ""
	write-host "Config.xml Was not able to be imported. This could be due to a typo or a special character such as `'&`' being incorrectly used." -foregroundcolor red
	write-host "The error message below will show which line in the clientconfig.xml is invalid:" -foregroundcolor red
	write-host $PSitem.exception.message -foregroundcolor red
	write-host ""
	pause
	exit
}

if ($config.GamePath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$script:GamePath = $config.GamePath.replace("`"","")
}
else {
	$script:GamePath = $config.GamePath
}
if ($config.ShortcutCustomIconPath -match "`""){#Remove any quotes from path in case someone ballses this up.
	$ShortcutCustomIconPath = $config.ShortcutCustomIconPath.replace("`"","")
}
else {
	$ShortcutCustomIconPath = $config.ShortcutCustomIconPath
}
$DefaultRegion = $config.DefaultRegion
$AskForRegionOnceOnly = $config.AskForRegionOnceOnly
$CreateDesktopShortcut = $config.CreateDesktopShortcut
$ConvertPlainTextPasswords = $config.ConvertPlainTextPasswords

#Check Windows Game Path for D2r.exe is accurate.
if((Test-Path -Path "$GamePath\d2r.exe") -ne $True){ 
	write-host "Gamepath is incorrect. Looks like you have a custom D2r install location! Edit the $GamePath variable in the script" -foregroundcolor red
	pause
	exit
}

# Create Shortcut
if ($CreateDesktopShortcut -eq $True){
	$DesktopPath = [Environment]::GetFolderPath("Desktop")
	$Targetfile = "-ExecutionPolicy Bypass -File `"$WorkingDirectory\D2Loader.ps1`""
	$shortcutFile = "$DesktopPath\D2R Loader.lnk"
	$WScriptShell = New-Object -ComObject WScript.Shell
	$shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
	$shortcut.TargetPath = "powershell.exe" 
	$Shortcut.Arguments = $TargetFile
	if ($ShortcutCustomIconPath.length -eq 0){
		$shortcut.IconLocation = "$Script:GamePath\D2R.exe"
	}
	Else {
		$shortcut.IconLocation = $ShortcutCustomIconPath
	}
	$shortcut.Save()
}

#Check SetText.exe setup
#Add-MpPreference -ExclusionPath "$script:WorkingDirectory\SetText" #Removed on 24.4.23 as MS cleared this PUP from their end. #Add Defender Exclusion for SetText.exe directory, at least until Microsoft reviews the file (submitted 24.4.2023) and stops flagging as "Trojan:Win32/Wacatac.B!ml" as per https://github.com/hankhank10/false-positive-malware-reporting
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

if ($Script:AccountOptionsCSV -ne $null){
	#check Accounts.csv has been updated and doesn't contain the example account.
	if ($Script:AccountOptionsCSV -match "yourbnetemailaddress"){
		write-host
		write-host "You haven't setup accounts.csv with your accounts." -foregroundcolor red
		write-host "Add your account details to the CSV file and run the script again :)" -foregroundcolor red
		write-host
		pause
		exit
	}
	
	if ($ConvertPlainTextPasswords -ne $false){
		#Check CSV for Plain text Passwords, convert to encryptedstrings and replace values in CSV
		$NewCSV = Foreach ($Entry in $AccountOptionsCSV) {
			if ($Entry.PWisSecureString.length -gt 0 -and $Entry.PWisSecureString -ne $False){#if nothing needs converting, make sure existing entries still make it into the updated CSV
				$Entry
			}
			if (($Entry.PWisSecureString.length -eq 0 -or $Entry.PWisSecureString -eq "no" -or $Entry.PWisSecureString -eq $false) -and $Entry.PW.length -ne 0){
				$Entry.pw = ConvertTo-SecureString -String $Entry.pw -AsPlainText -Force
				$Entry.pw = $Entry.pw | ConvertFrom-SecureString
				$Entry.PWisSecureString = "Yes"
				write-host ("Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
				start-sleep -milliseconds 100
				$Entry
				$CSVupdated = $true
			}
			if ($Entry.PW.length -eq 0){#if csv has account details but password field has been left blank
				write-host
				write-host ("The account " + $Entry.AccountLabel + " doesn't yet have a password defined.") -foregroundcolor yellow
				write-host
				$Entry.pw = read-host -AsSecureString "Enter the Battle.net password for"$Entry.AccountLabel
				$Entry.pw = $Entry.pw | ConvertFrom-SecureString
				$Entry.PWisSecureString = "Yes"
				write-host ("Secured Password for " + $Entry.AccountLabel) -foregroundcolor green
				start-sleep -milliseconds 100
				$Entry
				$CSVupdated = $true
			}
		}
		if ($CSVupdated -eq $true){
			Try {
				$NewCSV | Export-CSV "$script:WorkingDirectory\Accounts.csv" -NoTypeInformation #update CSV file
				Write-host "Accounts.csv updated: Passwords have been secured." -foregroundcolor green
				start-sleep -milliseconds 4000
			}
			Catch {
				Write-host
				Write-host "Couldn't update Accounts.csv, probably because the file is open and locked." -foregroundcolor red
				write-host "Please close accounts.csv and run the script again!" -foregroundcolor red
				Write-host
				pause
				exit
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
#Set item quality array for randomizing quote colours. A stupid addition to script but meh.
$QualityArray = @(#quality and chances for things to drop based on 0MF values in D2r (I think?)
	[pscustomobject]@{Type='HighRune';Probability=1}
	[pscustomobject]@{Type='Unique';Probability=50}
	[pscustomobject]@{Type='SetItem';Probability=124}
	[pscustomobject]@{Type='Rare';Probability=200}
	[pscustomobject]@{Type='Magic';Probability=588}
	[pscustomobject]@{Type='Normal';Probability=19036}
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
function HighRune {
	process { write-host "  $x[38;2;255;165;000;48;2;1;1;1;4m$_$x[0m"}
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
"Greetings, stranger. I'm not surprised to see your kind here.",
"There is a place of great evil in the wilderness.",
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
			#Write-Host "Closing" $proc_id_populated $handle_id_populated
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
		if ($script:region -eq "NA" -or $script:region -eq "us"){$script:region = "us.actual.battle.net"}
		if ($script:region -eq "EU"){$script:region = "eu.actual.battle.net"}
		if ($script:region -eq "Asia" -or $script:region -eq "As"){$script:region = "kr.actual.battle.net"}
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
					Write-host
					Write-host "Good day to you partner :)"
					start-sleep -milliseconds 233
					Exit
				}
				$Script:AccountChoice = $Script:AccountOptionsCSV |where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
			}
		} until ($Script:AccountID -ne "r")
	}
	if (($null -ne $script:AccountUsername -and ($null -eq $script:PW -or "" -eq $script:PW) -or ($Script:AccountChoice.id.length -gt 0 -and $Script:AccountChoice.pw.length -eq 0))){#This is called when params are used but the password wasn't entered.
		$securedPW = read-host -AsSecureString "Enter the Battle.net password for $script:AccountUsername"
		$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedPW)
		$script:PW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
		$script:pwmanualset = $true
	}
	else {
		$script:pwmanualset = $false
	}
}
Function ChooseRegion {#AKA Realm. Not to be confused with the actual Diablo servers that host your games, which are all over the world :)
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
	if ($script:ParamsUsed -ne $true -and $ConvertPlainTextPasswords -ne $false){
		$script:acct = $Script:AccountChoice.acct.tostring()
		$encryptedPassword = $pw | ConvertTo-SecureString
		$pwobject = New-Object System.Management.Automation.PsCredential("N/A", $encryptedPassword)
		$script:pw = $pwobject.GetNetworkCredential().Password
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
	if ($config.ForceWindowedMode -eq $true){
		$arguments = $arguments + " -w"
	}
	$script:pw = $null

	#write-output $arguments #debug
	Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments"
	start-sleep -milliseconds 1500 #give D2r a bit of a chance to start up before trying to kill handle
	#Close the 'Check for other instances' handle
	Write-host "Attempting to close `"Check for other instances`" handle..."
	#$handlekilled = $true #debug
	$output = killhandle | out-string
	if(($output.contains("DiabloII Check For Other Instances")) -eq $true){
		$handlekilled = $true
		write-host "`"Check for Other Instances`" Handle closed." -foregroundcolor green
	}
	else {
		write-host "`"Check for Other Instances`" Handle was NOT closed." -foregroundcolor red
		write-host "Who even knows what happened. I sure don't." -foregroundcolor red
		write-host "You may need to kill this manually via procexp. Good luck hero." -foregroundcolor red
		write-host
		pause
	}

	if ($handlekilled -ne $True){
		Write-Host " Couldn't find any handles to kill." -foregroundcolor red
		Write-Host " Game may not have launched as expected." -foregroundcolor red
		pause
	}
	#Rename the Diablo Game window for easier identification of which account and region the game is.
	$rename = ($Script:AccountID + " - Diablo II: Resurrected - " + $Script:AccountFriendlyName + " (" + $Script:Region + ")")
	$command = ('"'+ $WorkingDirectory + '\SetText\SetText.exe" "Diablo II: Resurrected" "' + $rename +'"')
	try {
		cmd.exe /c $command
		#write-output $command  #debug
		#write-host "Window Renamed to $rename" -foregroundcolor green
		write-host "Window Renamed" -foregroundcolor green
		start-sleep -milliseconds 200
		write-host "Good luck hero..." -foregroundcolor magenta
	}
	catch {
		write-host "Couldn't rename window :(" -foregroundcolor red
		pause
	}
	start-sleep -milliseconds 900
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
