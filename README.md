This is a tool I made. Cool Huh.
########################
# But what does it do? #
########################
Script will allow opening multiple Diablo 2 resurrected instances and will automatically close the 'DiabloII Check For Other Instances' handle.
Script will import account details from CSV. Alternatively you can run script with account, region and password parameters ( "username@emailaddress.com"  "MyBNetAccountPassword" "na.actual.battle.net").


Instructions for setup below.
*Quick note on paths, ".\" refers to the folder where your script sits. You can save this in any folder you like, but the script does look for subfolders within this folder (eg .\handle). 

##############################
# Download this thing I made #
##############################
1. Download this and save it to a folder of your choosing :)

##########
# Handle #
##########
1. Download handle viewer from https://learn.microsoft.com/en-gb/sysinternals/downloads/handle
2. Extract the executable files (specifically handle64.exe) to .\handle (in other words a subfolder)

############
# Set Text #
############
1. Browse to the SetText Folder.
2. In the Address bar (the part that shows the path of the folder you're in), click in this, clear out anything that's in it, type in cmd and press enter.
3. This will open the command prompt with it set in the path of where you've saved the script.
4. Next copy and paste the following into CMD:
	SET var=%cd%
	"C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /target:winexe /out:"%var%\SetText.exe" "%var%\SetText.bas" /verbose
5. This should compile SetText.exe. This is used to give the Diablo windows a unique name once they've initialized.

#################
# Your Accounts #
#################
1. Open Accounts.csv in a text editor or excel
2. Add number for each account starting from 1.
3. Add your account sign in name
4. Add your account password
5. Add a 'friendly' name for each account, this will appear in your diablo window. I just added my Bnet usernames.

If opening in Excel you add this info in each column. Otherwise


##########################
# Optional Script Config #
##########################
- If you have a game path that's not the default ("C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected"), then you'll need to edit this.

- Set your default region if you just want to mash enter instead

- Set 'AskForRegionOnceOnly' to True if you only want to choose the region once.

- If you don't want passwords stored in plain text files and would instead rather type these in, set EnterPasswordsManually to $True.

##############################
# Credit for things I stole: #
##############################
- Handle killer script: https://forums.d2jsp.org/topic.php?t=90563264&f=87
- Set Text executable: https://stackoverflow.com/questions/39021975/changing-title-of-an-application-when-launching-from-command-prompt/39033389#39033389
- https://learn.microsoft.com/en-gb/sysinternals/downloads/handle