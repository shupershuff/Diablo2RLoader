**Notice about Update Issues for v1.8.1 through to v1.8.3**<br>
In v1.8.1 and v1.8.2 there was an issue that didn't present itself until the script was run on a day of the month that's a single digit.<br>
This meant if you ran the script any time between 1st through to the 9th of August 2023, then the update checker will have broken :(<br>
There was also a separate issue where this didn't work properly for folk in different regions where your computers use the different date/time formatting settings.<br>
If you're impacted, then you will notice an error message (in red) briefly show when launching the script and/or notice your script version is stuck on a version below v1.8.4.<br>
There are two easy fixes for this, perform the following:<br>
1. Open stats.csv and delete the value for LastUpdateCheck (eg change "3/08/2023 8:32:06 pm" to "").
2. Download the latest release and copy the D2Loader.ps1 file from the latest release to where your script sits to overwrite it.

---

# Overview
Greetings Stranger! I'm not surprised to see your kind here.<br>
<br>
This is a free script I made for loading multiple Diablo 2 Resurrected instances (AKA Multiboxing/MultiLaunching etc), but can also be used for a single account. 
This will help you load up your account(s) quickly from one place without having multiple install directories of the game eating up excessive drive space.<br>
This will also enable you easily switch realms for trades, DClones, rushes etc for one or more accounts from one simple menu.<br>

Oh yeah and you can check DClone status, the Current TZ AND the next TZ from this launcher. Cool aye?<br>
<br>
Using this script means you DON'T have to do any of this awful stuff to multibox:<br>
- Use the worst (but somehow most popular) method of opening the game from the battlenet client each time with multiple game installs taking up heaps of disk space.
- Set up shortcuts on your desktop for each account with parameters (including storing your password in plain text).
- Manually use ProcExp to kill the "Check for other instances" handle, or manually running a script for handle.exe to do the same thing.
- Use Virtual Machines's, multiple computers or multiple user accounts (Windows account switching).
- Run any dodgy executables where you don't know what's actually running.
 	- Note that this script uses handle64.exe which is a [Microsoft](https://learn.microsoft.com/en-us/sysinternals/downloads/handle) recommended tool.
 	- Note that this script also builds an executable called SetText.exe for window renaming, details of which can be seen on [StackOverflow](https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389). This .exe has been whitelisted by Microsoft.
	- Never take the authors word for it for anything you download. That's why this script is full open source if you want to have a skim through to see what it's doing.

Note that this script DOES NOT alter the game, automate key presses, game joining or add any efficiencies with RAM/VRAM usage. It's simply used to launch the game.
If you want to use additional scripts, programs or mods to achieve any of the above, then that's at your own risk.<br>

This is a labour of love, for a game I love, for what I feel is a pretty good gaming community :)<br>
This Readme is a bit wordy sorry, I've tried to capture all the information that anyone might ever ask for.<br>

Script Screenshot:<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/50dfcb19-8ef1-4e6f-8cde-35c1f92cbec6)<br>
What your windows will look like:<br>
![GameWindows](https://user-images.githubusercontent.com/63577525/233829532-f81afad2-4806-4d6a-bb9e-817c25758346.png)

## Exactly what does it do?
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
**Batch Open Accounts**<br>
Rather than open all accounts, you can open a group of accounts. This feature is designed for you creatures that have several accounts but only want to launch a subset of these, for example only launch the 3 accounts you primarily play from.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/0f98dc8e-319a-43b8-b860-c84cb9d6dd60)<br>
**Launch Each account with specific game settings**<br>
These features were made in mind for multiboxing where you may have different screen sizes and want your secondary accounts to have lower graphics settings:<br>
_Auto Settings Switcher_: If enabled you can essentially have it so all accounts have their own game settings to load from. Game settings are loaded from settings<_ID_>.json instead of settings.json.<br>
_Manual Settings Switcher_: Alternatively, if you want to specify which game settings you want to load from, you can choose the settings file each account should use when launching. Once enabled in config, this can be toggled on and off using 's' in the menu.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/3533250d-8558-41a9-911f-5adcb5b6360d)<br>
You can enable both of these features at the same time. See [Setup Steps](#setup-steps) below.<br>
**Statistics - Track your playtime**<br>
It was too technically difficult for Blizzard to track time played for D2r within their Battlenet Client so you can use my janky one instead.<br>
Time per account can be seen from the main menu. Total time the script has ran for can be seen by going into the info screen ('i').
Now you can look back on your D2r playtime and think back on all of the productive things you should've done, but didn't.<br>

Other misc stats and info can be seen on the info screen. Statistics are recorded locally to stats.csv and accounts.csv in your script folder.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/ad4a386a-e1f7-4955-aefe-eeed632a9e95)

**Terror Zone Details**<br>
You can also check the current and next Terror Zone by pressing 't'.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/2bb22b1e-3ea7-4d47-bac4-25c9d6ceda61)<br>
**Check DClone Status**<br>
You can also check the current DClone status by pressing 'd'.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/f6f2c934-7fce-47db-a052-97e42874d9be)<br>
**Be an Entertainer in Baals Comedy Club**<br>
If you're an A grade leecher like me and typically stand around in Baals Throne room sapping up XP, why not at least pretend you have a sense of humour by using the built in joke generator to copy & paste mediocre jokes.<br>
That way instead of providing any real value in terms of damage, you can provide entertainment value instead.<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/7ba83409-eeba-4ae6-ba1e-b68d03e850e3)<br>
**Launch Parameters**<br>
You can run the script using launch parameters instead.<br>
This is ideal if you want to create a desktop shortcut to open a set of accounts, or if you're a super nerd and you want to launch the accounts from a scheduled task or from Home Assistant so that your game is ready to go when you get home from work :)<br>
Available launch parameters and values to use are as per the table below:<br>
| &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Parameter&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | value example(s) | Purpose | Notes |
| ---------------------- | ------------------------------------------- | ------------------------------------- | ------------- |
| -account               | username@emailaddress.com                   | Specify Signin Address to pass through to the script | -AccountUsername also works as a parameter. Can't be used with -all or -batch |
| -pw                    | YourBNetAccountPassword                     | Specify Password to pass through to the script | Can't be used with -all or -batch |
| -region                | na.actual.battle.net <br> 1/2/3<br>NA/EU/AS/KR | Used to specify the connection region | Specify either the full server name, use the realm initials (NA/EU/AS/KR) or use 1, 2 or 3 as values to select NA, EU or KR |
| -all                   | True                                        | Opens all accounts                    | Recommend using -region with this parameter. |
| -batch                 | 1                                           | Opens a batch of accounts at once     | Recommend using -region with this parameter |
| -manualsettingswitcher | True                                        | Use this if you want to manually choose which settings file to load with each account. | Recommend not using this but instead enabling SettingSwitcherEnabled in your config file so that it automatically loads from settings<_ID_>.json |

To make a shortcut to open a set of accounts, copy the D2RLoader Shortcut, rename it to whatever suits, open the properties and add parameters to the target eg -batch 1 -region na

**Magic Find in the script**<br>
You might also notice the quotes sometimes change colour, each time you refresh the script you have a chance to roll for Normal, Magic, Rare, Set, Unique quality quotes.
There's also a 1 in 19,999 chance to land a High Rune but you'll never see this :)<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/710a8709-dc13-4f7f-bdd6-28d9552e6373)<br>
There are ways to improve your script MaGic Find...

# Setup Steps
## 1. Download
1. Download the latest [release](https://github.com/shupershuff/Diablo2RLoader/releases) this and extract the zip file to a folder of your choosing.
2. Right click on D2loader.ps1 and open properties.
3. Check the "Unblock" box and click apply.<br>
![image](https://user-images.githubusercontent.com/63577525/234503557-22b7b8d4-0389-48fa-8ff4-f8a7870ccd82.png)

## 2. Setup Handle viewer
1. Download handle viewer from https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
2. Extract the executable files (specifically handle64.exe) to the .\Handle\ folder

## 3. Setup Your Accounts
1. Open Accounts.csv in a text editor (eg notepad), excel or your preferred editor.
2. Add number for each account starting from 1.
3. Add your account sign in address.
4. Add your account password. This will be converted to an encrypted string after first run. If left empty, you will be prompted to enter it when running script and the encrypted password string will be added to the csv.
	- If you're using a text editor to edit the CSV AND your password has a comma in it, ensure your password is surrounded by quotes eg "fjl3Ng2<,03h%mn"
5. Add a 'friendly' name for each account, this will appear in your diablo window. You can put anything here, I just added my Bnet usernames.
6. If you have several accounts and want to use the batch feature, ensure you add the number(s) into the batch column.
	- Note if editing the CSV using a text editor, ensure that if you're adding multiple batch options for an account that these are surrounded by quotes eg "1,2,4".
 	- Don't forget to enable the Batch feature in the config file.
7. If you have any custom launch (AKA Command Line) arguments you want to set, add these under the 'CustomLaunchArguments' column for each account you want these to apply too.
8. Leave the PWIsSecureString and TimeActive columns blank. These will be auto filled in.
9. Yeah and you should probably save it. That's CTRL + S. It helps :)

**Account CSV BEFORE running script:**<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/0e3c26c5-8c71-431c-8bf7-06bde4c36c91)<br>
**What it will look like AFTER running the script (in a later step):**<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/7ca6ed47-d5b7-486d-8bf2-a1fcfba2a612)<br>
![image](https://github.com/shupershuff/Diablo2RLoader/assets/63577525/d5453bf8-0af4-43ac-af5b-b16d31590c26)


## 4. Script Config (Mostly Optional)

Default settings within config.xml *should* be ok but can be optionally changed. Recommend checking out the features here.
Open the .xml file in a text editor such as notepad, Powershell ISE, Notepad++ etc.
- **Most importantly**, if you have a game path that's not the default ("C:\Program Files (x86)\Diablo II Resurrected"), then you'll need to edit this to wherever you chose to install the game.<br>

All other config options below this are strictly optional:<br>
- Set your default region if you just want to mash enter instead of choosing the region. Default is 1 for NA.
- Set 'EnableBatchFeature' to True if you want the ability to launch accounts in batches. You must also define the batches in your accounts.csv file. Disabled by default.
- Set 'DisableOpenAllAccountsOption' to True if you want to disable the ability of opening all accounts at once. Recommend leaving this to False. Disabled by default.
- Set 'CheckForNextTZ' to True if you want to enable the web request to find NextTZ details. Enabled by default.
- Set 'AskForRegionOnceOnly' to True if you only want to choose the region once. Disabled by default.
- Set 'CreateDesktopShortcut' to False if you don't want a handy dandy shortcut on your desktop. Enabled by default.
- Set 'ShortcutCustomIconPath' to the location of a custom icon file if you want the desktop icon to be something else (eg the old D2LOD logo). Uses D2r logo by default.
- Set 'ConvertPlainTextPasswords' to False if you want your passwords to be ~~stolen~~ in plain text. This will not convert already encrypted passwords back to plain text.
- Set 'ForceWindowedMode' to True if you want to force windowed mode each time. This causes issues with Diablo remembering resolution settings, so I recommend leaving this as False and manually setting your game to windowed in your game settings. Disabled by default.
- Set 'SettingSwitcherEnabled' to True if you want your Diablo accounts to load different settings. This essentially changes settings.json each time you launch a game. See the [Auto Setting Switcher](#6-auto-settings-switcher-optional-but-recommended) section above for more info. Disabled by default.
- Set 'ManualSettingSwitcherEnabled' to True if you want the ability to be able to choose a settings profile to load from. Once enabled, this is toggleable from the script using 's'. See the [Manual Setting Switcher](#7-manual-settings-switcher-optional) section above for more info. Disabled by default.
- Set 'TrackAccountUseTime' to False if you don't want accounts.csv or stats.csv to be autoupdated with playtime. Other Stats are still tracked in stats.csv. Mainly added this option in the unlikely case there are any issues with accounts.csv getting corrupted. Enabled by default.

Done editing? What are your thoughts on saving the file? I've heard it helps. CTRL + S for the win :)

## 5. Run the script manually for the first time
1. Browse to the folder, right click on D2Loader.ps1 and choose run.
2. If you get prompted to change the execution policy so you can run the script, type y and press enter.
   ![image](https://user-images.githubusercontent.com/63577525/234580880-e78df284-edea-4a5e-b4c6-4825f6031b4e.png)   
   a) If you instead get a message about "D2Loader.ps1 cannot be loaded because running scripts is disabled on this system" then you will need to perform the following steps:   
   b) Open the start menu and type in powershell. Open PowerShell.<br>
   c) Enter the following command: **Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser**<br>
   d) Type in y and press enter to confirm.<br>
   e) Run the D2Loader.ps1 script again.<br>
3. If the script prompts to trust it and add it to the unblock list, type in y and press enter to confirm.
4. This will perform the first time setup for compiling settext.exe, encrypting your passwords and will create a shortcut on your desktop.

If you've skipped ahead, the script will error out and tell you which of the previous setup steps you've skipped. 

## 6. Auto Settings Switcher (Optional but recommended)
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

## 7. Manual Settings Switcher (Optional)
Do you want to manually choose which settings to use when launching the game? This is for you! This feature is disabled by default, as this needs to be setup first and understood this first.<br>
<br>
Setup is exactly the same as the Auto Settings Switcher, except for step 8 as you need to name the settings file settings._name_.json where name is whatever you want it called (eg settings.1440pHigh.json or settings.PotatoGraphics.json)<br>
- Note: If you name the file settings_name_.json it will not work. The name should be inside two fullstops "."<br>

Don't forget to enable this feature in the [config](#4-script-config-mostly-optional) file by setting 'ManualSettingSwitcherEnabled' to True.

# Notes #
## FAQ / Common Issues

**Q:** The script won't let me run it as it gives me security prompts about scripts being disabled or the current script being untrusted :(<br>
**A:** See Instructions above. The default script policy for Windows 10/11 devices is restricted. We can change this to remote signed. A full write up of the policies can be seen [here](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.3).

**Q:** How do I update the script?<br>
**A:** As of 1.5.0, the script has the ability to update itself. I will also include steps in the script to update config.xml, accounts.csv and stats.csv if new fields are added. For users on older versions, to update manually, all you need to do is download the latest release, extract the .zip and copy the new D2Loader.ps1 over to where the old D2Loader.ps1 file is and overwrite it. If this script shows a warning "Couldn't check for updates. GitHub API limit may have been reached...", it's most likely due to me testing things and exceeding GitHubs API limit, ignore and try again later.

**Q:** When Diablo opens it won't allow me to select an online character<br>
**A1:** This will be due to your password or username being entered in incorrectly. Please note that multiple failed authentication requests can cause a temporary lockout from that realm (seems to be around 15-30mins).<br>
**A2:** In some circumstances, Battlenet can also require a capcha code to be entered to verificaton. If in doubt, try logging in via the battlenet client and see if it prompts for captcha. It might take several hours for this to resolve itself (6 hours observed in [issue #17](https://github.com/shupershuff/Diablo2RLoader/issues/17)).

**Q:** I have reset one of my Bnet account passwords, how do I update accounts.csv<br>
**A:** Open accounts.csv and clearout the password field and the PWIsSecureString field. Leave the PWIsSecureString field blank. Either enter your password into the csv file or leave it blank and the script will ask you next time you run it. 

**Q:** A UAC prompt opens each time asking me to run as Admin. This is annoying. Can I disable this?<br>
**A:** Yes, there are a couple ways to do this, see here: https://silicophilic.com/add-program-to-uac-exception/#Method_2_Run_Programs_With_Admin_Privileges_Without_UAC_Prompt

**Q:** Why does the script need to run as admin?<br>
**A:** The script needs to run as admin in order to kill the "Check for Other instances" process handle and to be able to rename your D2r windows once launched. The script uses the names of these Windows to detect which accounts are currently active.

**Q:** I get 2FA/MFA Battlenet prompts on my screen but even though I approve, when the game loads it won't show online characters.<br>
**A:** Bad news here sorry, Diablo does not work with MFA enabled when launching the game from a shortcut with parameters. Blame Blizzard, their MFA solution overall isn't great either.

**Q:** I would like to say "Thankyou". How do I do that?<br>
**A:** Please pay my entire mortgage. Thanks in advance. Or [D2JSP funny money](https://forums.d2jsp.org/gold.php?i=1328510). Or your [local animal charity](https://www.youtube.com/watch?v=dQw4w9WgXcQ). Or just a message to say thanks :)<br>

**Q:** I have suggestions and/or issues with this, where do I post these?<br>
**A:** Please use GitHub issues for any feedback. Thanks!

**Q:** Is this Script against ToS?<br>
**A:** Multiboxing itself is not against Blizzard TOS as per this [Blizzard Rep](https://us.forums.blizzard.com/en/d2r/t/blizzard-please-give-us-an-official-statement-on-multiboxing/21958/5) and this [Blizzard Article](https://eu.battle.net/support/en/article/24258). However the only way of achieving this without additional physical computers or Virtual Machines is by killing the "Check for Other instances" handle.

Outside of killing this handle and changing the window title, there are absolutely no modifications to the game made by this script, it's simply an improved alternative way to start the application.
To be clear, this script in no way enhances or assists with actual game play and I would strongly advise against seeking/using any tools.
The script is essentially launching the game the same way you would if you had setup shortcuts to the game with account,region,pw parameters, launching that way and then killing the "Check for other instances" handle (as suggested in several guides). This script is a QoL tool to help consolidate your accounts to one simple launcher to simply open the game with account(s) and region(s) you want.

So the real question is, regardless of this script, is using procexp or another method to kill the "check for other instances" handle against ToS? Stricly speaking yes and this topic has been broached in Blizzard forums many times without an official response either for or against.
If you're reading this your real question is actually "Will I get banned for multiboxing by killing the 'check for other instances' handle?" which to that I'm confident the answer is no. If I wasn't confident, this script wouldn't exist and people wouldn't be using procexp/handle.exe to multibox through traditional methods. Given the widespread use of procexp/handle being used to multibox, this method (and therefore also this script), is considered safe.

## Notes about the Window title rename (SetText.exe)

The script will generate a file called SetText.exe if it doesn't exist.
This is used to rename the Game Windows so that you and the script can tell each instance apart.
To compile the .exe this requires DotNet4.0. If you don't have it the script will prompt you to download this from Microsoft.
~~A Windows Defender exception will also be automatically added for the directory this sits in, as at the time of writing (24.4.2023), Windows Defender considers it to be dodgy.~~ A submission has since been sent to Microsoft and submission has been cleared :)

If you have a 3rd Party Anti-Virus product installed and it kicks up a fuss, you may need to manually add an exception to the .\SetText\ folder location.

Optional: If you don't trust me and want to build the .exe yourself you can do the following.
1. Browse to the SetText Folder.
2. In the Address bar (the part that shows the path of the folder you're in), click in this, clear out anything that's in it, type in cmd and press enter.
3. This will open the command prompt with it set in the path of where you've saved the script.
4. Next copy and paste the following into CMD:
	SET var=%cd%
	"C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /target:winexe /out:"%var%\SetText.exe" "%var%\SetText.bas" /verbose
5. This should compile SetText.exe. This is used to give the Diablo windows a unique name once it's initialized.

See this site for more information on what this does: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389

# What's Next #
At this point I've implemented everything that I wanted to, so I don't have any more planned features to add to the tool.<br>
If there's something you want to see added or improved then let me know. Future updates may include:
* Fixing anything I broke in the last release.
* Adding whatever features you fools ask for.
* Investigate use of battlenet login tokens (stored in registry) instead of passwords as other loaders have done. This is unlikely as I probably can't figure it out.
* Perhaps make a GUI *if* there's enough interest. Probably not though as there would be a lot of brain activity involved. Pay my mortgage and we'll perhaps maybe talk... probably.

# Usage and Limitations #
Happy for you to make any modifications this script for your own needs providing:
 - Any variants of this script are never sold.
 - Any variants of this script published online should always be open source.
 - Any variants of this script are never modifed to enable or assist in any game altering or malicious behaviour including (but not limited to): Bannable Mods, Cheats, Exploits, Phishing

# Credit for things I stole: #
- Handle killer script: https://forums.d2jsp.org/topic.php?t=90563264&f=87
- Set Text method: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389
- Handle tool (replaces procexp) - https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
- MoonUnit for thoughts around converting plain text passwords to encrypted strings.
- never147 for contributing improvements for menu refresh and inputs.
- TheGodOfPumpkin for Next TZ source
- https://ocr.space/OCRAPI for their free OCR API
- ChatGPT for helping with regex patterns.
- Google.com for everything else.
<br>
<br>
Tags for Google SEO (maybe): Multiboxing, Multiboxes, multibox, multi-box, multi-boxing, multi-launcher, boxer, launcher, Shuper, d2loader, d2rloader, diabloloader, loader, D2r, Diablo 2: Resurrected, Diablo II: Resurrected, powershell, process explorer, procexp, windows, battle.net, warriv, d2r Multi, d2r launcher, d2r loader, d2r multibox, chat Gem workinG as intended
