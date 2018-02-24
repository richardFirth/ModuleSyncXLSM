# ModuleSyncXLSM
Syncs Modules between VBA Projects

## Purpose
This macro is used to 'sync' certain modules within VBA projects. The idea is that if you create a module or class in one .xlsm project, and then import it into other projects and add improvements in the later projects, you can use this tool to make all your projects contain the latest version of the class or module.

## File Structure
This macro should be kept in a folder alongside another .xlsm file called "BestModules" which holds a copy of each tracked module.
The directory should also contain a .txt file called "ModSyncList.txt" that holds the paths of all macros that are desired to be synced.

## Version Control
modules/classes that you wish to version control should have the text added at the top of the modules. This macro will export all modules as .bas, .cls or .frm text files, and then look in each one for this text to determine the version number.
~~~~
'$VERSIONCONTROL
'$*MINOR_VERSION*x.x
'$*DATE*13Feb18
'$*ID*versionID
'$*CharCount*1234*xxxx
'$*RowCount*123*xxxx
~~~~

## Change Detection
when a module is accepted, the CharCount and RowCount are automatically updated to be the counts of that module. if you update the code, these values will most likely change, and the next time you run the macro the change will be detected and displayed.


## UI - subject to change
Buttons:
* Pull in ModSyncList Data: Pulls filepaths from ModSyncList (This is likely the first button you press)
* Browse to Files: This allows you to browse to your own list of files (use as alternative to ModSyncList)
* Compare Versions: This will pull in the modules data and display it in the excel file. it does this by opening each excel file and exporting the .bas files
* Update Modules to latest: this will take any module that is outdated by a more recent version and replace the old version with a new copy. (This will open and save any workbook that had an old module)


## Warnings
Not tested with macros that have any code that executes upon open or close, or that do any extensive modification of excel functionality (dictator apps)
also not usable on locked VBA projects. Won't work with two files that share names but live in different folders. Use at own risk.
