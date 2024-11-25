# Overview
Greetings Stranger! I'm not surprised to see your kind here.<br>
<br>
This is a free script I made for loading multiple Diablo 2 Resurrected instances (AKA Multiboxing/MultiLaunching etc), but can also be used for a single account, ideal if you want to play single player and switch between mods easily. 
This will help you load up your account(s) quickly from one place without having multiple install directories of the game eating up excessive drive space.<br>
This will also enable you easily switch realms for trades, DClones, rushes etc for one or more accounts from one simple menu.<br>

Oh yeah and you can check DClone status, the Current TZ AND the next TZ from this launcher. Cool aye?<br>
<br>
Using this script means you DON'T have to do any of this awful stuff to multibox:<br>
- Use the worst (but somehow most popular) method of opening the game from the battlenet client each time with multiple game installs taking up heaps of disk space.
- Set up shortcuts on your desktop for each account with parameters (including storing your password in plain text).
- Manually use ProcExp to kill the "Check for other instances" handle, or manually running a script for handle.exe to do the same thing.
- Use Virtual Machines, multiple computers or multiple user accounts (Windows account switching).
- Run any dodgy executables where you don't know what's actually running.
 	- Note that this script uses handle64.exe which is a [Microsoft](https://learn.microsoft.com/en-us/sysinternals/downloads/handle) recommended tool.
 	- Note that this script also builds an executable called SetTextv2.exe for window renaming, details of which can be seen on [StackOverflow](https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389). This .exe has been whitelisted by Microsoft.
	- Never take the authors word for it for anything you download. That's why this script is full open source if you want to have a skim through to see what it's doing.

Note that this script DOES NOT alter the game, automate key presses, game joining or add any efficiencies with RAM/VRAM usage. It's simply used to launch the game.
If you want to use additional scripts, programs or mods to achieve any of the above (which I would strongly discourage), then that's at your own risk.<br>

This is a labour of love, for a game I love, for what I feel is a pretty good gaming community :)<br>
This Readme is a bit wordy sorry, I've tried to capture all the information that anyone might ever ask for.<br>

I've put several 100's of hours into making this, so if you like it, consider buying me a beer at https://www.buymeacoffee.com/shupershuff or here [https://github.com/sponsors/shupershuff](https://github.com/sponsors/shupershuff?frequency=one-time&sponsor=shupershuff&amount=5) or here https://paypal.me/Shupershuff.<br>
Cheers!

Script Screenshot:<br>
![ScriptMainMenu](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/50dfcb19-8ef1-4e6f-8cde-35c1f92cbec6)<br>
What your windows will look like:<br>
![GameWindows](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/62129f82-bde4-4744-83f2-cc69d873988a)
## But I don't want to use your script you dodgy internet human
Not everyone wants to use a random script or an app and that's understandable.<br>
See guides for alternative multiboxing methods here: [https://github.com/shupershuff/D2r-Multiboxing-Without-A-Script](https://github.com/shupershuff/D2r-Multiboxing-Without-A-Script)

## Exactly what does the script do?
Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle.<br>
It will achieve this by importing account details from a CSV that you populate and essentially launches the game the same way you would with a shortcut: by passing account, password and region parameters to D2r.exe.<br>
	Note: Plain text passwords entered into the CSV will be converted into a secure string after running. If you don't want to enter plain text passwords EVER then you can leave the PW field in the CSV blank and manually type in when running the script.<br>
Once the game has initialised, the window will be renamed so it's easier to tell which account and region each game is using.<br>
This also helps the script know which games are open to prevent you accidentally opening a game with the same account twice.<br>
Optionally you can also have the game launch using custom settings.json for each account in case you want different graphics/audio/game settings for each account you have.<br>
<br>
**`* This script in no way enhances or changes gameplay. *`**

## Other Features
**Open All accounts at once**<br>
Time is precious so work smarter not harder by opening all your accounts at once to maximise your free time to actually play the game instead of clicking through menus.<br>
Note that if you've configured your account(s) or a region(s) to launch the game using an Authentication Token instead of parameters, you will need to wait for each game to reach the character selection screen before the next instance can launch.

**Batch Open Accounts**<br>
Rather than open all accounts, you can open a group of accounts. This feature is designed for you creatures that have several accounts but only want to launch a subset of these, for example only launch the 3 accounts you primarily play from.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/e2705d60-525e-4f10-a70c-8d78675c2529)
<br>

**Launch Each account with specific game settings**<br>
These features were made in mind for multiboxing where you may have different screen sizes and want your secondary accounts to have lower graphics settings:<br>
_Auto Settings Switcher_: If enabled you can essentially have it so all accounts have their own game settings to load from. Game settings are loaded from settings<_ID_>.json instead of settings.json.<br>
_Manual Settings Switcher_: Alternatively, if you want to specify which game settings you want to load from, you can choose the settings file each account should use when launching. Once enabled in config, this can be toggled on and off using 's' in the menu.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/3533250d-8558-41a9-911f-5adcb5b6360d)<br>
You can enable both of these features at the same time. See [Setup Steps](#setup-steps) below.<br>

**Statistics - Track your playtime**<br>
It was too technically difficult for Blizzard to track time played for D2r within their Battlenet Client so you can use my janky one instead.<br>
Time per account can be seen from the main menu. Total time the script has ran for with D2r running can be seen by going into the info screen ('i').
Now you can look back on your D2r playtime and think back on all of the productive things you should've done, but didn't.<br>

Other misc stats and info can be seen on the info screen. Statistics are recorded locally to stats.csv and accounts.csv in your script folder.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/ad4a386a-e1f7-4955-aefe-eeed632a9e95)

**Terror Zone Details**<br>
You can also check the current and next Terror Zone by pressing 't'.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/7f612465-be54-4883-ba37-5febfa39c58d)<br>
Data courtesy of [D2Emu.com](https://d2emu.com)<br>

**Check DClone Status**<br>
You can also manually check the current DClone status by pressing 'd'.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/f6f2c934-7fce-47db-a052-97e42874d9be)<br>
Data courtesy of [D2Emu.com](https://d2emu.com), [d2runewizard.com](https://d2runewizard.com) and [diablo2.io](https://diablo2.io). That's right, you can choose your source!<br>

**Alarms for DClone Walk status changes**<br>
If configured, you can select which regions and modes to monitor for D Clone (Über Diablo) walk status changes.<br>
If there's a change in status whilst the script is running, it will activate the alarm function.<br>
The alarm will a text warning (as seen in example below) as well as a text to speech alarm notifying you where the walk is happening.<br>
The voice alarm activates only once but the text warnings will remain in place for 5 minutes. <br>
You will also be notified after the script has launched if there's any imminent walks about to happen (ie status is 5/6).<br>
See the [DClone Status Alarms](#7-dclone-status-alarms-optional) and [config](#3-script-config-mostly-optional) sections for more information and how to configure this.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/929f62d0-a952-449d-bbdc-ad3a9126d008)<br>
Voice Examples (make sure to unmute, GitHub mutes by default):<br>

https://github.com/shupershuff/Diablo2RLoader/assets/63577525/50e9a49d-a01c-40e4-8654-a8da9fe40c05

https://github.com/shupershuff/Diablo2RLoader/assets/63577525/56bd87d5-157f-4119-b99b-bd1d26f06052

**Be an Entertainer in Baals Comedy Club**<br>
If you're an A grade leecher like me and typically stand around in Baals Throne room sapping up XP, why not at least pretend you have a sense of humour by using the built in joke generator to copy & paste mediocre jokes.<br>
That way instead of providing any real value in terms of damage, you can provide entertainment value instead.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/55f47b81-c86e-424b-9954-9a1796ae5f9b)<br>
Jokes courtesy of v2.jokeapi.dev, official-joke-api.appspot.com, icanhazdadjoke.com and api.chucknorris.io<br>

**Remember Windows Layout**<br>
If configured, the script can launch each game instance to the preferred screen location and window size so that you don't have to rearrange your game windows at launch.<br>
![image](https://github.com/user-attachments/assets/a416e4d2-c337-49d3-868e-0828b23dd66e)

**Launch Parameters**<br>
You can run the script using launch parameters instead.<br>
This is ideal if you want to create a desktop shortcut to open a set of accounts, or if you're a super nerd and you want to launch the accounts from a scheduled task or from Home Assistant so that your game is ready to go when you get home from work :)<br>
Available launch parameters and values to use are as per the table below:<br>
| &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Parameter&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | value example(s) | Purpose | Notes |
| ---------------------- | ------------------------------------------- | ------------------------------------- | ------------- |
| -account               | username@emailaddress.com                   | Specify Signin Address to pass through to the script | -AccountUsername also works as a parameter. Can't be used with -all or -batch. If you don't use the -pw parameter, you can simply specify the ID of the account you want to launch from accounts.csv. |
| -pw                    | YourBNetAccountPassword                     | Specify Password to pass through to the script | Can't be used with -all or -batch. If not specified, script will check if there's a matching account in accounts.csv to use password or token for. |
| -region                | na.actual.battle.net <br> 1/2/3<br>NA/EU/AS/KR | Used to specify the connection region | Specify either the full server name, use the realm initials (NA/EU/AS/KR) or use 1, 2 or 3 as values to select NA, EU or KR |
| -all                   | True                                        | Opens all accounts                    | Recommend using -region with this parameter. |
| -batch                 | 1                                           | Opens a batch of accounts at once     | Recommend using -region with this parameter |
| -manualsettingswitcher | True                                        | Use this if you want to manually choose which settings file to load with each account. | Recommend not using this but instead enabling SettingSwitcherEnabled in your config file so that it automatically loads from settings<_ID_>.json |

To make a shortcut to open a set of accounts, copy the D2RLoader Shortcut, rename it to whatever suits, open the properties and add parameters to the target eg -batch 1 -region na<br>

**Magic Find in the script**<br>
You might also notice the quotes sometimes change colour, each time you refresh the script you have a chance to roll for Normal, Magic, Rare, Set, Unique quality quotes.
There's also a 1 in 19,999 chance to land a High Rune but you'll never see this :)<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/710a8709-dc13-4f7f-bdd6-28d9552e6373)<br>
There are ways to improve your script MaGic Find...

# Setup Steps
Please see the detailed setup steps below, it's not as scary as it looks. I've included details for all of the features you may or may not want to use.
A basic setup only takes 3-5 minutes.<br>
For those allergic to reading, have a gander at [this youtube video tutorial here](https://www.youtube.com/watch?v=JYMs-soQr_c) that a great bloke made. I promise it's not a rick roll video :)<br>

TL;DR steps are:
1. Download the latest release, extract and unblock it within file properties. See [1. Download](#1-download).
2. Add account info to accounts.csv. See [2. Setup Your Accounts](#2-setup-your-accounts).
3. Edit config.xml to set game path and enable/configure desired features. See [3. Script Config](#3-script-config-mostly-optional).
4. Run script for first time. See [4. Run the script manually for the first time](#4-run-the-script-manually-for-the-first-time).
5. Optionally perform steps to configure game settings that each account should load from. See [5. Auto Settings Switcher](#5-auto-settings-switcher-optional-but-recommended) and [6. Manual Settings Switcher](#6-manual-settings-switcher-optional).
6. Optionally enable DClone alarms for your preferred regions and game mode. See [7. DClone Status Alarms](#7-dclone-status-alarms-optional).

&#x1F534; If you have any issues, come back and fully read these instructions. I can almost guarantee any issue you see is covered in the detailed setup steps below and/or FAQ section.

## 1. Download
1. Download the latest [release here](https://github.com/shupershuff/Diablo2RLoader/releases). Click on Diablo2RLoader-<version>.zip (eg Diablo2RLoader-1.13.2.zip) to download.
2. One downloaded, extract the zip file to a folder of your choosing.
3. Right click on D2loader.ps1 and open properties.
4. Check the "Unblock" box and click apply.<br>
![image](https://user-images.githubusercontent.com/63577525/234503557-22b7b8d4-0389-48fa-8ff4-f8a7870ccd82.png)

## 2. Setup Your Accounts
**NOTE: If you have MFA configured on one account it will not work with Parameter based authentication. This to work due to Blizzard not implementing MFA capability with this authentication method. If you want to keep MFA enabled, you can utilise the AuthToken method outlinned below.**<br>
&#x1F536;&#x1F536; **Special Note - PLEASE READ THIS** &#x1F536;&#x1F536; Please pay attention to these instructions and particularly to step 8 around auth token setup. Almost all setup issues are due to missing or incorrectly performing one or more of these steps.  
1. Open Accounts.csv in a text editor (eg notepad), excel or your preferred editor. Can recommend [moderncsv](https://www.moderncsv.com/) as a csv editor.
2. Add number for each account starting from 1.
	- Note, if you want you, can also use letters as an ID but take note that characters a, b, c, g, r, t, d, j, s, i and x are all reserved.
3. Add your battlenet account sign in address (eg bogan@askjeeves.com).
4. Add your battlenet account password. This will be converted to an encrypted string after first run. If left empty, you will be prompted to enter it when running script and the encrypted password string will be added to the csv.
	- If you're using a text editor to edit the CSV AND your password has a comma in it, ensure your password is surrounded by quotes eg "fjl3Ng2<,03h%mn"
5. Add a 'friendly' name for each account, this will appear in your diablo window. You can put anything here, I just added my Bnet usernames. You could simply set these as "Barb", "Sorc" etc.
6. [OPTIONAL] If you have several accounts and want to use the batch feature, ensure you add the number(s) into the batch column.
	- Note if editing the CSV using a text editor, ensure that if you're adding multiple batch options for an account that these are surrounded by quotes eg "1,2,4".
 	- Don't forget to enable the Batch feature in the config file.
7. [OPTIONAL] If you have any custom launch (AKA Command Line) arguments you want to set, add these under the 'CustomLaunchArguments' column for each account you want these to apply too.
	- EG If you're one of the people who have [Extracted game files with cascviewer to 'improve' game performance](https://www.reddit.com/r/Diablo/comments/qey05y/d2r_single_player_tips_to_improve_your_load_times/) and want to use the "-direct -txt" launch flags, this is where you put them.
8. [OPTIONAL BUT RECOMMENDED] If you want or need to use Token based authentication (eg if you have MFA enabled on your account / Blizzards Auth servers are down / You get errors about account being "locked for suspicious activity" on certain regions), you will need to populate the 'Token' column.<br>
&#x1F536;&#x1F536; **Special Note - PLEASE READ THIS** &#x1F536;&#x1F536;Due to some kind of Blizzard issue, it's becoming more common for username and password authentication (via Parameters) to not work for *some* accounts. There is no apparent pattern or reason for this. Whilst launching with Parameters offers the best experience, if you have issues you should try using the Auth Token method instead for any impacted accounts.<br>
	a. Open your preferred internet browser in private mode (Ctrl + Shift + N) and browse to this website https://us.battle.net/login/en/?externalChallenge=login&app=OSI<br>
	b. Log in with your credentials and approve MFA request (if enabled).<br>
	c. You will be brought to an error page (this is expected). Copy the URL from the error page into the 'token' column of accounts.csv. <br>
 		**DO NOT SHARE THIS TOKEN INFORMATION ONLINE.<br>**
		![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/dfde17f7-a068-4060-9304-f92bee4bd067)<br>
	d. If you want the account to launch with token based authentication by default, change 'Parameter' to 'token' in the 'AuthenticationMethod' column. You can alternatively leave as 'Parameter' and toggle in the script (from the info menu) if you want to temporarily force Token based auth. This is good for when you generally want to use parameters for authentication but need to temporarily use AuthTokens to switch to another server such as Asia when the Blizzard haven't fixed authentication issues. Generally speaking you will have the best experience launching with parameters. You can toggle the script to temporarily use auth tokens instead of parameters from the info menu.<br>
	e. Close the browser, reopen in private mode, log into each of your other accounts and repeat the steps A to C above.<br>
	f. The token will be converted to an encrypted string when script is next run.<br>
	g. You will need to redo this step if you add/remove MFA to your account.<br>

Make sure to save it and close the file :)

**Account CSV BEFORE running script:**<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/4b53981e-4915-46b1-afaf-54d59f77f041)<br>
**What it will look like AFTER running the script (in a later step):**<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/7ca6ed47-d5b7-486d-8bf2-a1fcfba2a612)<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/982aa776-492b-491d-81da-ee34e4151dca)

## 3. Script Config (Mostly Optional)
Default settings within config.xml *should* be ok but can be optionally changed. Recommend checking out the features here.
Open the .xml file in a text editor such as notepad, Powershell ISE, Notepad++ etc.
- **Most importantly**, if you have a game path that's not the default ("C:\Program Files (x86)\Diablo II Resurrected"), then you'll need to edit this to wherever you chose to install the game.<br>

All other config options below this are strictly optional:<br>
- Set 'DefaultRegion' to your preferred default region if you just want to mash enter instead of choosing the region. Default is 1 for NA.
- Set 'DisableOpenAllAccountsOption' to True if you want to disable the ability of opening all accounts at once. Recommend leaving this to False. Disabled by default.
- Set 'CreateDesktopShortcut' to False if you don't want a handy dandy shortcut on your desktop. Enabled by default.
- Set 'ShortcutCustomIconPath' to the location of a custom icon file if you want the desktop icon to be something else (eg the old D2LOD logo). Uses D2r logo by default.
- Set 'ConvertPlainTextSecrets' to False if you want your passwords and tokens to be ~~stolen~~ stored in plain text. This will not convert already encrypted passwords & tokens back to plain text if disabled.
- Set 'RememberWindowLocations' to True if you want to each game instance to launch to a preferred window layout.
- Set 'ForceWindowedMode' to True if you want to force windowed mode each time. This causes issues with Diablo remembering resolution settings, so I recommend leaving this as False and manually setting your game to windowed in your game settings. Disabled by default.
- Set 'SettingSwitcherEnabled' to True if you want your Diablo accounts to load different settings. This essentially changes settings.json each time you launch a game. See the [Auto Setting Switcher](#5-auto-settings-switcher-optional-but-recommended) section below for more info. Disabled by default.
- Set 'ManualSettingSwitcherEnabled' to True if you want the ability to be able to choose a settings profile to load from. Once enabled, this is toggleable from the script using 's'. See the [Manual Setting Switcher](#6-manual-settings-switcher-optional) section below for more info. Disabled by default.
- Set 'TrackAccountUseTime' to False if you don't want accounts.csv or stats.csv to be autoupdated with playtime. Other Stats are still tracked in stats.csv. Mainly added this option in the unlikely case there are any issues with accounts.csv getting corrupted. Enabled by default.
- Set 'DCloneTrackerSource' to one of the options noted in the config file. Default (and recommended) is d2emu.com as it provides realtime data (not crowdsourced).
- Set 'DCloneAlarmList' to any number of the options noted in the config file to enable DClone Status alarms for your preferred GameModes and regions. See the [DClone Status Alarms](#7-dclone-status-alarms-optional) section below for more info. Blank by default.
- Set 'DCloneAlarmLevel' depending on the DClone statuses changes you want to be alarmed on (if alarms are enabled). 'Imminent' notifies only on 1,5,6. 'Close' notifies on status changes to 1,4,5,6. 'All' notifies for...all status changes. All by default.
- Set 'DCloneAlarmVoice' to the preferred Text to Speech Robot voice. Choices are 'Amazon' or 'Paladin'. Amazon by Default.
- Set 'DCloneAlarmVolume' to a preferred volume (1-100) to prevent frights and save your ear drums. Default is 69. Nice.
- Set 'ForceAuthTokenForRegion' to enforce AuthToken based authentication for one or more regions. Useful if an Auth server goes down preventing parameter based connections (remember when Asia stopped working for several weeks?) Valid options are NA, EU and KR. Multiple values should be comma separated. Recommend leaving blank unless there are auth issues.

Done editing? What are your thoughts on saving the file? I've heard it helps. CTRL + S for the win :)

## 4. Run the script manually for the first time
1. Browse to the folder, right click on D2Loader.ps1 and choose run.
2. If you get prompted to change the execution policy so you can run the script, type y and press enter.
   ![image](https://user-images.githubusercontent.com/63577525/234580880-e78df284-edea-4a5e-b4c6-4825f6031b4e.png)   
   a) If the script opens up and immediately closes or you instead get a message about "D2Loader.ps1 cannot be loaded because running scripts is disabled on this system" then you will need to perform the following steps:   
   b) Open the start menu and type in powershell. Right click on PowerShell and click "Run as administrator".<br>
   c) Enter the following command: **Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser**<br>
   d) Type in "y" and press enter to confirm.<br>
   e) Run the D2Loader.ps1 script again.<br>
3. If the script prompts to trust it and add it to the unblock list, type in y and press enter to confirm.
4. This will perform the first time setup for compiling settext.exe, encrypting your passwords and will create a shortcut on your desktop.

If you've skipped ahead, the script will error out and tell you which of the previous setup steps you've skipped. 

## 5. Auto Settings Switcher (Optional but recommended)
Do you want your primary account to launch with decent graphics settings with your other accounts to be set to poo tier settings? This is for you! This feature is disabled by default, as it will cause confusing behaviour for users who haven't read and understood this first.<br>

What this feature does is create copies Settings.json (found in the "C:\Users\\\<yourusername>\Saved Games\Diablo II Resurrected" folder) for each account you have.<br>
E.G if you have 3 accounts, you will have Settings1.json, Settings2.json & Settings3.json. When the script runs and you choose account 1, it copies Settings1.json to Settings.json causing the game to essentially load off those settings. This essentially turns Settings.json into a temporary file that's just used at load time.<br>
NOTE: Any changes you make to non-character options in game (eg graphics, audio, game options) will be saved to Settings.json, which will be overwritten the next time you launch a game via this loader. Therefore if you want to edit your game settings for say your 2nd account, you would need to open diablo (on any account), make the options changes you'd like to see that account have each time, close the game and then copy Settings.json to Settings2.json.

**Quick Guide to updating Game settings** for a particular account with this auto switcher.<br>
*For these instructions, let's assume we're trying to edit the config for account 1*
1. Set 'SettingSwitcherEnabled' to True in your config file.
2. Launch the Game (via the Loader or via Bnet client, doesn't matter, the account you log into doesn't matter either).
3. Make the required graphics/audio/game changes via the menu.
4. Close the game.
5. Browse to "C:\Users\\\<yourusername>\Saved Games\Diablo II Resurrected"
6. If there's already a file called Settings1.json, delete it (1 being the account ID in accounts.csv).
7. Copy the Settings.json file and paste into the same folder.
8. Rename to Settings1.json
9. Launch the game and proceed find all of the high runes. All of them.

## 6. Manual Settings Switcher (Optional)
Do you want to manually choose which settings to use when launching the game? This is for you! This feature is disabled by default, as this needs to be setup first and understood this first.<br>
<br>
Setup is exactly the same as the Auto Settings Switcher, except for step 8 as you need to name the settings file settings._name_.json where name is whatever you want it called (eg settings.1440pHigh.json or settings.PotatoGraphics.json)<br>
- Note: If you name the file settings_name_.json it will not work. The name should be inside two fullstops "."<br>

Don't forget to enable this feature in the [config](#3-script-config-mostly-optional) file by setting 'ManualSettingSwitcherEnabled' to True.

## 7. DClone Status Alarms (Optional)
You can optionally configure the script to advise when DClone walk status changes.<br>
This is very handy for when you're playing on another region/game mode or playing another game entirely so you can be audibly warned on status changes without having to manually check a website.<br>
This will display text warnings on screen for 5 minutes since last status change and will also perform a one off voice alarm advising you what region and game mode the status changed on.
Checks are performed each time the script refreshes on the main menu (IE, every 30 seconds).<br>
Note that this check is run in a background job that takes 1-2 seconds to complete. As such manually refreshing the menu constantly (r) will prevent checks from occuring.<br>

This feature is disabled by default.<br>
If you want to enable it, simply add in the option(s) you want into DcloneAlarmList in config.xml.<br>
You will only be alarmed for the game options you add into config. EG if you only want Hardcore Ladder notifications you would enter: <DCloneAlarmList>HCL-NA, HCL-EU, HCL-KR</DCloneAlarmList><br>
You can also optionally change the other configs for DClone Tracking and DClone Alarms, however I recommend leaving the default values here.<br>
You can also adjust the volume of the alarms if they are too quiet or if it's so loud it's making you jump off your chair.
See the [Script Config](#3-script-config-mostly-optional) section for more info on each config.<br>

## 7. Setup your preferred window layout (Optional)
Perform these steps if you've enabled the feature for remembering window layout and size.
1. Enable the feature if not already enabled (you can do this from the options menu or by setting 'RememberWindowLocations' to 'true' config.xml).
2. Open all of your D2r account instances.
3. Move/adjust the window for each game instance to your preferred layout and size.
4. Go to options menu and go into the 'RememberWindowLocations' setting.
5. Once in this menu, choose the option 's' to save coordinates of any open game instances.

Now whenever you launch the accounts they will open in the same positions with the same window sizes. Use the 'r' option within that menu to reset them if need be.

# Notes #
## Graphics Performance Recommendations ##
If you don't want your Diablo games to run like a slideshow, here are some tips. You'll of course need to adjust based on your hardware and setup.
1. IMPORTANT. Set an Frame Rate (FPS) cap in the graphics options. Recommend 60 but adjust depending on your GPU power and instances you're running. This will prevent each instance from trying to fully utilise your GPU compute.
2. Reduce graphics settings in game to performance over quality.
3. Run games in Windowed mode. Especially if you have more than one monitor. Not only will it save you performance, it's waaaay easier than alt tabbing.
4. In addition to changing your settings to launch the game in windowed mode, you should also make your secondary account windows smaller (ie lower resolution) to save on GPU utilisation. You can adjust resolution from the in game menu or simply by shrinking the window. If need you can of course do the opposite and maximise one of the smaller windows if you are playing another character for a bit.
5. If you are wanting different graphics settings for different accounts (eg different FPS cap, different audio settings, different resolution, nicer graphics settings etc etc), then I would highly recommend making use of the QOL features of this script. The automatic settings switcher can be used so that each account loads with it's own settings at launch. There is also the manual setting switcher feature if you want to define what settings file to load for a given account. See the relevant sections above for more info.
6. Obviously if you run other graphic demanding things like wallpaper engine in the background, this will hurt your overall FPS. Some wallpapers will have neglible impact but some are  noticably impactful on performance.
7. Make use of hardware monitoring software to see how much you are utilising RAM, CPU, GPU - Processor and GPU - VRAM. If you are under or overutilised you can adjust your settings accordingly for the best experience.

**My setup**<br>
Note that on my specs (5950x, RTX2080s (which has 8GB VRAM), 32GB RAM), I run my instances on the lowest graphics options possible. <br>
DLSS is set to ultra performance. Framerate (FPS) caps for secondary accounts are around 50fps. For my primary account I set the FPS cap to 60fps.<br>
<br>
With 3 instances (2 windows at approx 1280x780 resolution, 1 window (primary account) at 2556x1373 resolution, The CPU is barely used, Memory is about 85% utilised, VRAM is 80% utilised and GPU is about 85-90% (6800MB).<br>
With 4 instances (an additional window at approx 1280x780), CPU is still fine (15%), Memory is maxed out, VRAM is 92%+ (7600MB+) and GPU is 95-100%. Different parts of the game can run pretty poorly and as such sometimes I reduce the secondary accounts FPS Cap to around 44 instead.<br>
<br>
I've noted that with my hardware running 4 instances is generally not fun graphically due to performance stutters, FPS drops and increased chance of crashing. I'd argue that running 4 or more accounts is confusing and isn't fun logistically either but I digress.

## FAQ / Common Issues

**Q:** I would like to say "Thankyou". How do I do that?<br>
**A:** Please pay my entire mortgage. Thanks in advance. Buy me a beer here https://www.buymeacoffee.com/shupershuff.<br>
Or here [https://github.com/sponsors/shupershuff](https://github.com/sponsors/shupershuff?frequency=one-time&sponsor=shupershuff&amount=5).<br>
Or [D2JSP funny money](https://forums.d2jsp.org/gold.php?i=1328510).<br>
Or your [local animal charity](https://www.youtube.com/watch?v=dQw4w9WgXcQ).<br>
Or even just a [message](https://github.com/shupershuff/Diablo2RLoader/discussions) to say thanks :)<br>

**Q:** The script won't let me run it as it gives me security prompts about scripts being disabled or the current script being untrusted :(<br>
**A:** See instructions [above](#4-run-the-script-manually-for-the-first-time). The default script policy for Windows 10/11 devices is restricted. We can change this to remote signed. A full write up of the policies can be seen [here](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.3).

**Q:** When Diablo opens it won't allow me to select an online character<br>
**A1:** This will be due to your password or username being entered in incorrectly. Please note that multiple failed authentication requests can cause a temporary lockout from that realm (seems to be around 15-30mins).<br>
**A2:** In some circumstances, Battlenet can also require a capcha code to be entered to verificaton. If in doubt, try logging in via the battlenet client and see if it prompts for captcha. It might take several hours for this to resolve itself (6 hours observed in [issue #17](https://github.com/shupershuff/Diablo2RLoader/issues/17)).<br>
**A3:** In some circumstances, for no real reason, you might randomly have issues with one or more of your accounts connecting with warnings about "Account locked for suspicious activity" after temporarily switching regions. This only happens when using parameter based authentication. Numerous people have seen this issue across different loaders and for people who launch the game from desktop shortcuts. I've discovered no pattern as to why this happens or how to resolve it. Workaround is to launch problematic accounts using Auth Tokens.

**Q:** I have reset one of my Battle.net account passwords, how do I update accounts.csv?<br>
**A:** Open accounts.csv and clearout the password field. Either enter your new password into the csv file or leave it blank and the script will ask you next time you run it.

**Q:** When I run this I'm unable to use Push to Talk (PTT) in Discord.<br>
**A:** As the script needs to run as administrator, this also means it starts the game with elevated rights. As such this can mess with Discord PTT. This is a known issue with Discord. To resolve this you can run [Discord in Administrator mode](https://support.discord.com/hc/en-us/articles/205082178-How-do-I-enable-push-to-talk-if-I-am-running-my-game-in-administrator-mode).<br>
I recommend that you find the Discord Shortcut or app, go into properties > Compatibility and check the box for "Run this program as administrator" so that it always runs as admin. :)

**Q:** A UAC prompt opens each time asking me to run as Admin. This is annoying. Can I disable this?<br>
**A:** Yes, there are a couple ways to do this, see here: https://silicophilic.com/add-program-to-uac-exception/#Method_2_Run_Programs_With_Admin_Privileges_Without_UAC_Prompt

**Q:** Will this work if I've extracted game files with Casc viewer in an attempt to make the game load faster?<br>
**A:** Yes, if you're one of the people that have done this, you can still run the game using this script. Make sure to specify "-direct -txt" in the CustomLaunchArguments column in Accounts.csv.

**Q:** Why does the script need to run as admin?<br>
**A:** The script needs to run as admin in order to kill the "Check for Other instances" process handle and to be able to rename your D2r windows once launched. The script uses the names of these Windows to detect which accounts are currently active.

**Q:** I get 2FA/MFA Battlenet prompts on my screen but even though I approve, when the game loads it won't show online characters.<br>
**A1:** Bad news here sorry, Diablo does not work with MFA enabled when launching the game from a shortcut with parameters. Blame Blizzard, their MFA solution overall isn't great either.<br>
That said, it is possible to connect if you utilise the Auth Token method instead of Parameters. To setup AuthToken authentication, (see Auth Token steps in [Setup Your Accounts](#2-setup-your-accounts)).

**Q:** I have suggestions and/or issues with this, where do I post these?<br>
**A:** Please use GitHub issues for any feedback. Thanks!

**Q:** If I've read this far through this long ass readme does that mean I can read good?<br>
**A:** Yes, you read at least up to a Dr Seuss reading level. Well done you!

**Q:** I get prompts to enter captcha's on my accounts randomly. Why is this?<br>
**A:** I'm not sure exactly why this happens but it does appear to happen more with newer accounts. Ensure you have logged into the battlenet client AND into your webbrowser with the account and approve any verifications/captchas needed. Do not share the account with other people. Do NOT use a VPN when signing into the account (whether it be via this script, bnet client or website).

**Q:** Will you implement a feature to skip intro videos when launching?<br>
**A:** Could I? Yes. Will I? No as this requires modifying game files (video files really). Whilst the script can cater for mods you want to use, I don't want this script directly modifying any game files, even something as innocent as intro videos. If skipping intro is of interest to you, you can look at the [introskip mod](https://www.nexusmods.com/diablo2resurrected/mods/194).

**Q:** I'm getting an error says "You have not been online in the last 30 days. Please start the game while online to check for any login agreements".
**A:** As the error suggests, try logging in from the battle.net client. If issues persist, use Google as there are dozens of threads where people have had this issue due to other reasons (eg firewall).

**Q:** Is this Script against ToS?<br>
**A:** Multiboxing itself is not against Blizzard TOS as per this [Blizzard Rep](https://us.forums.blizzard.com/en/d2r/t/blizzard-please-give-us-an-official-statement-on-multiboxing/21958/5) and this [Blizzard Article](https://eu.battle.net/support/en/article/24258). However the only way of achieving this without additional physical computers or Virtual Machines is by killing the "Check for Other instances" handle.

Outside of killing this handle and changing the window title, there are absolutely no modifications to the game made by this script, it's simply an improved, alternative way to start the application.
To be clear, this script in no way enhances or assists with actual game play and I would strongly advise against seeking/using any tools for automation.
The script is essentially launching the game the same way you would if you had setup shortcuts to the game with parameters (account, region and password), launching that way and then killing the "Check for other instances" handle (as suggested in several guides). This script is a QoL tool to help consolidate your accounts to one simple launcher to simply open the game with the account(s) and region(s) you want.

So the real question is, regardless of this script, is using procexp or another method to kill the "check for other instances" handle against ToS? Stricly speaking yes and this topic has been broached in Blizzard forums many times without an official response either for or against.
If you're reading this your real question is actually "Will I get banned for multiboxing by killing the 'check for other instances' handle?" which to that I'm confident the answer is no. If I wasn't confident, this script wouldn't exist and people wouldn't be using procexp/handle.exe to multibox through traditional methods. Given the widespread use of procexp/handle being used to multibox, this method (and therefore also this script), is considered safe.

This script can also be used for folk who have one account and want to track playtime details or simply want to be able to use an interface to launch different single player mods.

## Notes about the Window title rename (SetTextv2.exe)

The script will generate a file called SetTextv2.exe if it doesn't exist.
This is used to rename the Game Windows so that you and the script can tell each instance apart.
To compile the .exe this requires DotNet4.0. If you don't have it the script will prompt you to download this from Microsoft.
~~A Windows Defender exception will also be automatically added for the directory this sits in, as at the time of writing (24.4.2023), Windows Defender considers it to be dodgy.~~ <br>
~~A submission has since been sent to Microsoft and submission has been cleared :)~~ <br>
I have since made a new version of SetText so that it can rename Windows based on process ID (instead of looking for any window matching Diablo II) to prevent any issues with folk who launch an instance via Battlenet as well as the script.

If you have a 3rd Party Anti-Virus product installed and it kicks up a fuss, you may need to manually add an exception to the .\SetText\ folder location.

Optional: If you don't trust me and want to build the .exe yourself you can do the following.
1. Browse to the SetText Folder.
2. In the Address bar (the part that shows the path of the folder you're in), click in this, clear out anything that's in it, type in cmd and press enter.
3. This will open the command prompt with it set in the path of where you've saved the script.
4. Next copy and paste the following into CMD:
	SET var=%cd%
	"C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /target:winexe /out:"%var%\SetText.exe" "%var%\SetTextv2.bas" /verbose
5. This should compile SetText.exe. This is used to give the Diablo windows a unique name once it's initialized.

See this site for more information on what this does, this contains the original version and the updated version that I've made: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389

# What's Next #
You tell me. If there's something you want to see added or improved then let me know. Future updates may include:<br>
* ~~Possibly add the ability to mute minimised windows (as long as it can be done within windows without additional software)~~ - Would require an external script or software.
* ~~Investigate Ability to skip intro~~ - Not possible without mods.
* Fixing anything I broke in the last release.
* Adding whatever features you fools ask for (InB4 questions for autojoining etc. I will never add any in game automation activities to this script).
* Perhaps make a GUI *if* there's enough interest. Probably not though as there would be a lot of brain activity involved. Pay my mortgage and we'll perhaps maybe talk... probably.

# Usage and Limitations #
Happy for you to make any modifications this script for your own needs providing:
 - Any variants of this script are never sold.
 - Any variants of this script published online should always be open source.
 - Any variants of this script are never modifed to enable or assist in any game altering or malicious behaviour including (but not limited to): Bannable Mods, Cheats, Exploits, Phishing, Botting, Game Automation.

# Credit for things I stole #
- Handle64 tool (replaces process explorer aka procexp) - https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
- Set Text method: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389
  - After using the above for several months, I built my own to work with Process ID instead of process name for improved accuracy. Posted to the thread above.
- Thanks to MoonUnit for contributing thoughts around converting plain text passwords to encrypted strings.
- Thanks to never147 for contributing improvements for menu refresh and inputs. Huge QOL feature and allowed for more features to be implemented.
- Thanks to Mysterio from [D2Emu.com](https://d2emu.com/tz) for providing TZ source API. Consider buying Mysterio a coffee [here](https://www.buymeacoffee.com/d2emu).
- Thanks to Mysterio ([D2Emu.com](https://D2Emu.com)), Prowner ([d2runewizard.com](https://d2runewizard.com)) and Teebling "Teebs" ([diablo2.io](https://diablo2.io)) for providing their awesome respective API's for DClone status for you to choose from.
- Thanks to dschu012 for [discovering the AuthToken method](https://github.com/Farmith/D2RMIM/pull/11/files#diff-5408bbaf05738fe52729de093b38981abecffeb304b1cd388713cbe6a0461d21) and thanks to Sunblood for pointing me towards this discovery.
- Thanks to v2.jokeapi.dev, official-joke-api.appspot.com, icanhazdadjoke.com and api.chucknorris.io for API's providing top notch cringe for us to smirk at.
- Thanks to Sir-Wilhelm for tidying up Handle killer and providing code for resizing and relocating windows.
- ChatGPT for helping with regex patterns.
- Google.com for everything else.
- Live, Laugh, Love.
- Lastly big thanks to everyone that's messaged me with issues, ideas, performed testing, sent FG donations, or simply messaged to say thanks. It's awesome to be able to help out and bring some QOL to folks across the world :)

Proudly made in NZ<br>
![nzmade](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/39e332b8-71cb-4149-8afd-2fcfdac14abf)
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/746968fa-6be2-4846-bc62-850717d84daa)<br>
<br>
Latest Version Downloads:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ![GitHub Downloads (all assets, latest release)](https://img.shields.io/github/downloads/shupershuff/diablo2rloader/latest/total?label=Downloads&link=https%3A%2F%2Fgithub.com%2Fshupershuff%2FDiablo2RLoader%2Freleases%2Flatest)<br>
Release with the most downloads:&nbsp;&nbsp;![Downloads](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/shupershuff/Diablo2RLoader/main/.github/max-release-download-count.json?Label=Downloads)<br>
All Time Downloads:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ![Github All Releases](https://img.shields.io/github/downloads/shupershuff/Diablo2RLoader/total.svg?label=Downloads)<br>
Page views as of 2nd Oct 2024:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fshupershuff%2FDiablo2RLoader&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://www.youtube.com/watch?v=dQw4w9WgXcQ)<br>

Tags for Google SEO (maybe): Multiboxing, Multiboxes, multibox, multi-box, multi-boxing, multi-launcher, boxer, launcher, Shuper, d2loader, d2rloader, script, diabloloader, loader, D2r, Diablo 2: Resurrected, Diablo II: Resurrected, uber, dclone, youtube, DiabloII, powershell, process explorer, procexp, windows, battle.net, warriv, d2r Multi, d2r launcher, d2r loader, alt, d2r multibox, chat Gem workinG as intended
