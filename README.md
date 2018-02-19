# ModuleSyncXLSM
Syncs Modules between VBA Projects

## Purpose
This macro is used to 'sync' certain modules within VBA projects. The idea is that if you create a class module in one .xlsm project, and then import it into other projects and add improvments in the second projects, you can use this tool to make all your projects contain the latest version.

## File Structure
This macro should be kept in a folder alongside another .xlsm file called "BestModules" which holds copies of each tracked module.
The directory should also contain a .txt file called "ModSyncList.txt" that holds the paths of all macros that are desired to be sync'ed

## Version Control
modules/classes that you wish to version control should have the text added at the top of the modules. this macro will export all modules as .bas text files, and then look in each one for this text to determine the version number.
~~~~
'$VERSIONCONTROL
'$*MINOR_VERSION*x.x
'$*DATE*13Feb18
'$*ID*versionID
~~~~


## UI - subject to change
Buttons:
* Pull in ModSyncList Data: Pulls filepaths from ModSyncList (This is likely the first button you press)
* Browse to Files: This allows you to browse to your own list of files (use as alternative to ModSyncList)
* Compare Versions: This will pull in the modules data and display it in the excel file. it does this by opening each excel file and exporting the .bas files
* Update Modules to latest: this will take any module that is outdated by a more recent version and replace the old version with a new copy. (This will open and save any workbook that had an old module)


## Warnings
Not tested with macros that have any code that executes upon open or close, or that do any extensive modification of excel functionality (dictator apps)
also not usable on locked VBA projects.
