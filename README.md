# Overview  #
This is a free tool I made for loading multiple Diablo 2 Resurrected instances (AKA Multiboxing/Multiboxes). 
Instead of setting up shortcuts on your desktop for each account (or worse, opening the game from battlenet client) and then manually using ProcExp to kill the handle, run this one script instead.
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

# Is this Script against ToS? #
This script is the same as opening multiple instaces from the battlenet client OR from setting up shortcuts to the game with account,region,pw parameters and launching that way and then killing the "Check for other instances" handle.
No modifications to the game are being made, this script just helps save you mouse clicks to launch the game because RSI is real with this game.

So the real question is, is Multiboxing against ToS?
I don't think so, but there are discussions on various forums regarding this matter.

# Setup Steps #
**Download**
1. Download the latest release this and save it to a folder of your choosing :)

**Setup Handle viewer**
1. Download handle viewer from https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
2. Extract the executable files (specifically handle64.exe) to the .\Handle folder

**Setup Your Accounts**
1. Open Accounts.csv in a text editor or excel.
2. Add number for each account starting from 1.
3. Add your account sign in address
4. Add your account password (or not, if empty in the csv you can type this into the window each time)
5. Add a 'friendly' name for each account, this will appear in your diablo window. I just added my Bnet usernames.

If opening in Excel you add this info in each column. Otherwise open and edit in notepad.

**Optional Script Config**
Some settings under the Script Options section can be changed by editing D2Loader.ps1 in a text editor.
- Most importantly, if you have a game path that's not the default ("C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected"), then you'll need to edit this.
- Set 'AskForRegionOnceOnly' to True if you only want to choose the region once.
- Set your default region if you just want to mash enter instead of choosing the region.

**Optional - Recompile SetText**
Optional, only do this if you don't trust me and want to build the .exe yourself. I've left my compile of SetText.exe in here for those who can't be bothered :)

1. Browse to the SetText Folder.
2. In the Address bar (the part that shows the path of the folder you're in), click in this, clear out anything that's in it, type in cmd and press enter.
3. This will open the command prompt with it set in the path of where you've saved the script.
4. Next copy and paste the following into CMD:
	SET var=%cd%
	"C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /target:winexe /out:"%var%\SetText.exe" "%var%\SetText.bas" /verbose
5. This should compile SetText.exe. This is used to give the Diablo windows a unique name once it's initialized.


# Credit for things I stole: #
- Handle killer script: https://forums.d2jsp.org/topic.php?t=90563264&f=87
- Set Text executable: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389
- https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
- Google.com
