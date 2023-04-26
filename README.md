# Overview  #
This is a free tool I made for loading multiple Diablo 2 Resurrected instances (AKA Multiboxing/Multiboxes). 
Instead of setting up shortcuts on your desktop for each account (or worse, opening the game from the battlenet client with multiple game installs) and then manually using ProcExp to kill the "Check for other instances" handle, run this one script instead and keep it open to easily switch realms for trades/dclones etc.
Cool aye?

![Launcher](https://user-images.githubusercontent.com/63577525/233829526-2b28f2b9-761b-4d95-af0f-6561bda8ddf3.png)
![GameWindows](https://user-images.githubusercontent.com/63577525/233829532-f81afad2-4806-4d6a-bb9e-817c25758346.png)

# But what does it do? #
Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle.
It will achieve this by importing account details from a CSV that you populate.
	Note: If you don't want your password details stored here, you can leave blank and type in manually when running script.
Once the Game has initialised, the window will be renamed so it's easier to tell which game is which.
This also helps the script know which games are open to prevent you accidentally opening a game with the same account twice.

Note: If for some unknown reason you prefer to, you can call the script with account, password & region parameters: -account username@emailaddress.com -pw MyBNetAccountPassword -region na.actual.battle.net


# Setup Steps #
**Download**
1. Download the latest [release](https://github.com/shupershuff/Diablo2RLoader/releases) this and save it to a folder of your choosing :)

**Setup Handle viewer**
1. Download handle viewer from https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
2. Extract the executable files (specifically handle64.exe) to the .\Handle\ folder

**Setup Your Accounts**
1. Open Accounts.csv in a text editor or excel.
2. Add number for each account starting from 1.
3. Add your account sign in address
4. Add your account password. This will be converted to an encrypted string after first run. If left empty, you will be prompted to enter it when running script and the encrypted password string will be added to the csv.
5. Add a 'friendly' name for each account, this will appear in your diablo window. You can put anything here, I just added my Bnet usernames.

If opening in Excel you add this info in each column. Otherwise open and edit in notepad.

**Run the script manually for the first time**
1. Browse to the folder, right click on D2Loader.ps1 and choose run.
2. This will perform the first time setup for compiling settext.exe, encrypting your passwords and will create a shortcut on your desktop.

**Optional Script Config**

Some settings under the Script Options section can be changed by editing D2Loader.ps1 in a text editor.
- Most importantly, if you have a game path that's not the default ("C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected"), then you'll need to edit this.
- Set 'AskForRegionOnceOnly' to True if you only want to choose the region once.
- Set your default region if you just want to mash enter instead of choosing the region.

# Notes #
**FAQ / Common Issues**

Q: The script won't let me run it as it gives me security prompts about scripts being disabled or the current script being untrusted :(<br>
A: This issue is currently under investigation under [issue #1](https://github.com/shupershuff/Diablo2RLoader/issues/1). In the interim, the workaround is to open Powershell as administrator and run the following command: Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUse

Q: I can't login despite my password being entered in correctly.<br>
A: Try logging into the Battlenet client and approve any 2FA/verification requests that might pop up.

Q: I have reset one of my Bnet account passwords, how do I update accounts.csv<br>
A: Open accounts.csv and clearout the password field and the PWIsSecureString field. Leave the PWIsSecureString field blank. Either enter your password into the csv file or leave it blank and the script will ask you next time you run it. 

Q: Is this Script against ToS?<br>
A: Multiboxing itself is not against Blizzard TOS as per this [Blizzard Rep](https://us.forums.blizzard.com/en/d2r/t/blizzard-please-give-us-an-official-statement-on-multiboxing/21958/5). However the only way of achieving this without additional physical computers or Virtual Machines is by killing the "check for Other instances handle".

Outside of killing the this handle and renaming the window, there are absolutely no modifications to the game made by this script.
The script is essentially launching the game the same way you would if you had setup shortcuts to the game with account,region,pw parameters, launching that way and then killing the "Check for other instances" handle. This script just helps save you mouse clicks to get you to the character select screen faster with a reduced chance of RSI :).

So the real question is, is using procexp or another method to kill the "check for other instances" handle against ToS? I don't think so, but there are discussions on various forums regarding this matter.

**Notes about the Window title rename (SetText.exe)**

The script will generate the .exe file if it doesn't exist. This is to prevent your browser considering it a dodgy file and stopping the download.
To compile the .exe this requires DotNet4.0. If you don't have it the script will prompt you to download this from Microsoft.
~~A Windows Defender exception will also be automatically added for the directory this sits in, as at the time of writing (24.4.2023), Windows Defender considers it to be dodgy.~~ A submission has since been sent to Microsoft and submission has been cleared :)

If you have a Anti-Virus product installed and it kicks up a fuss you may need to manually add an exception to the .\SetText\ folder.

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
* Ability to launch accounts from separate game installs to retain game config for each account.
* Perhaps make a GUI if there's enough interest.
* Maybe add a counter for how many times you've launched each account.

# Credit for things I stole: #
- Handle killer script: https://forums.d2jsp.org/topic.php?t=90563264&f=87
- Set Text executable: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389
- https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
- MoonUnit for thoughts around converting plain text passwords to encrypted strings.
- Google.com
