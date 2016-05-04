Nightmare Redux is an editor for the game MajorMUD(r).  

v1.8 (??/??/2016)  
------------------------------------------  
-NEW: Monster attack simulator and Average/Max round calculator
-UP: Added spells resist type field to MME export for future MME enhancement  
-UP: Added attack names field to MME export for future MME enhancement  
-UP: Changed drop down spell editor's "Type of Resist" to what I believe they actually mean  
-UP: Added logging of the second universal modifier's only if option when chosen  
-FIX: Universal modifier reported that it would set the directive to "0" when setting something to "= ##" (only the message was wrong, not the action)  
-FIX: Fix for MME export erroring out when having a large number of database records  

v1.7.1 (4/2/2016)  
------------------------------------------  
-FIX: Fixed rooms opening up to wrong tab  
-UP: Added energy field to MME export for future MME enhancement  

v1.7 (4/1/2016) - (by Syntax)  
------------------------------------------  
-NEW: ability to mark control rooms on the map in the room editor  
-NEW: tool to view and export said list of control rooms and their references  
-NEW: complete overhaul of database exporter.  support for multiple ranges of records, ability to save config files, **import record numbers from existing database**  
-NEW: added cross-referencing of class restrictions on learned spells via scrolls and textblocks for MME.  requires new version of MME to support it (coming soon).  
-NEW: added cross-referencing for where items/keys are used in rooms and textblocks.  requires new version of MME to support it (coming soon).  
-NEW: added default rooms to exclude on MME exports (currently only the sysop support chamber).  see additional options button.  
-UP: excluded rooms on MME export will now be completely excluded and not exported at all.  references from those rooms will also be considered as not in the game. option added to "hide" user-specified excluded rooms instead to retain previous functionality.  
-UP: added Wrist(2), Eyes and Face to universal modifier's "only if" options  
-UP: added second set of "only if" options on the universal modifier  
-UP: added text on db importer error to help determine which table was missing fields upon import from old exports  
-UP: updated about pages and links  
-UP: installer repackaged and updated to newer version  
 

v1.6.2/n (06/01/2007) - (by Ghaleon)  
------------------------------------------  
-BUGFIX: fixed container coin drop fieldmap locations  
-BUGFIX: fixed item race restriction fieldmap locations  
-UP: added hidden coins to rooms  
-UP: added hidden coins to export tool  
-UP: added offset value to actions so they display correctly in mud  

v1.6.1/n (02/03/2006)  
------------------------------------------  
-UP: Added support for v1.11p dat files.  

v1.6/n (07/01/2005)  
------------------------------------------  
General:  
-BUGFIX: ‘per item’ ‘only if’ on uni mod for monsters and shops  
-NEW: COMPLETE RECORD NUMBER CHANGER  
-NEW: monster group/index changer  
-NEW: slew of filtering options for monsters, items, and spells!  
-NEW: filter by ability, class, record numbers, and fields!  
-NEW: tool to fix/verify number of uses on items  
-NEW: tool to fix number of item uses on monster drops  
-NEW: tool to combine like items on the ground into one slot  
-NEW: tool to combine monster experience and experience multiplier  
-UP: added a spell range display for monster attacks  
-UP: added button on users to calculate the user’s exp  
-UP: exp calc will copy to clipboard the selected levels experience  
-UP: can now enter -1 to pad/delete buffer rooms on all maps at once  
-UP: most windows can now be maximized and will re-open that way  
-UP: changed some of the textblock line editor to be more intuitive  
-UP: added difficulty column to spell editor record list  
-UP: progress bar will now show % completion for all jobs in caption  
-UP: saving settings wont reload data files unless necessary  
-UP: option on MME export to only export certain databases  
-FIX: couldn't enter single room exclusion on MME export  
-FIX: importer/exporter will now delete extra textblock parts  
-FIX: progress bar progression on some tools  
-FIX: pressing alt, shift, control, alt-tab in search boxes  
-FIX: canceling on building new monster group/index list  

v1.56/n (07/19/2004)  
------------------------------------------  
General:  
-NEW: option to add/remove EDITED flag on users  
-FIX: selection issue on min damage field on item editor  

v1.55/n (07/17/2004)  
------------------------------------------  
General:  
-UP: all fields with quick-display boxes will auto-update  
-UP: added select-all text functions for almost every textbox  
-UP: description fields on all editors will auto trim spaces and auto-advance  
-UP: clicking save will now update the text on the row in the record list  
-UP: new records will no longer change the name to "New Record ##"  
-UP: spanned the record number columns to fit more numbers by default  
-UP: quick clear buttons for abilities  
-UP: added copy/paste options for monster attacks  
-UP: added a clear all button on message editor  
-UP: added a clear button on room editor to clear the advanced section  
-UP: freed up some memory with the monster group/index array +++  
-FIX: ability lookup when typing the ability name on the editors ++  
-FIX: 2nd wrist, face, and eye items displaying wrong in record list  
-FIX: tab order on all windows ...... again ... uhghh  
-FIX: error when inserting buffer rooms with room numbers > 32767  
-FIX: crash on room editor when clicking index list when it's minimized  
-FIX: no error handling when clicking on certain cells on map/map editor  
-FIX: clicking on a selected monster index line would edit the line  

MMUD Explorer:  
-NEW: map/room exclusion  
-NEW: save different configurations  
-NEW: database update link  
-UP: better chest and monster greet text handling  
-UP: datafile completely changed +  

Textblock Editor:  
-UP: line editor: when clicking save it will auto-click "update line"  
-UP: line editor: will only prompt for save if something was changed  
-UP: line editor: added a splitter bar so you can resize the panes  

Spell Editor:  
-UP: moved around spell editor fields again  
-UP: spell calculations will now take energy into account  
-UP: added boxes to show spell damage vs. X MR for damage spells  

Database Import/Export:  
-UP: added display of record names on import log  
-UP: monsters in room will not import or export anymore  
-UP: importer now cancelable  
-FIX: sometimes the last export path wouldn't save  
-FIX: progress bars will now accurately count to the end  

Universal Mod:  
-UP: added a few options for rooms  
-UP: added 'only if markup' for shops  

User Editor:  
-UP: can now paste full character stats (str, hlth, lives, cp, etc)  
-UP: added boxes to show the room names on user editor for the room trail  
-UP: added recognition of "Two Handed" for pasting user items  
-UP: clipboard text will automaticly paste into paste window  

+ Now users won't be able to import anything more than the record number  
and name columns of each of the databases.  
++ The ability lookup on the editors now functions like the search boxes.  
You can type any part of the ability name and press the right arrow key  
to find the next match.  When you tab out or click elsewhere it will fill in  
the entire ability name.  
+++ Only 11 monsters per index number will be listed.  The map may also  
load a little slower with the "look up monster regen" turned on as I no  
longer store the monster name in the array.  These two elements were  
causing a long delay when erasing the array (closing NMR) and creating it.  

v1.5/n (05/17/2004)  
------------------------------------------  
-NEW: taskbar for open NMR windows (like the windows taskbar)  
-NEW: textblock line editor! --makes editing scripts much easier  
-NEW: enhanced the item find to now search users and shops too  
-NEW: Find/Find Next for message editor  
-NEW: Find/Find Next for bank editor  
-NEW: universal modifier now does extensive logging  
-NEW: map now marks room spells/shops/exit & death rooms  
-NEW: option to only load # & name columns for faster loading  
-NEW: gave the action editor a record list  
-NEW: tool to set all monster killed times to a specified date  
-NEW: extra wrist, face, and eyes slots supported (v1.11p-beta12+ only)  
-UP: added different color hi-lighting to the editor icons :)  
-UP: importer now has options to import only room items, or visa versa  
-UP: NPC list now loads faster, unless btrieve is bogged down ++  
-UP: monster NPC list rows can now be double clicked to jump  
-UP: added reset class and race restriction buttons on item editor  
-UP: added "copy to" buttons on importer/exporter for ranges  
-UP: changed the font to terminal on the textblock editor  
-UP: got rid of undocking windows since it never worked anyway  
-UP: updated the swing calculator  
-UP: database export is now cancelable  
-UP: mass room editor is now cancelable  
-UP: database deleter now cancelable  
-UP: most tool operations from the menu are now cancelable  
-UP: moved range room delete to database deleter  
-UP: added copy/paste to mass room editor and re-arranged options  
-UP: resized monster/room/shop editors to fit better on 800x600 resolutions  
-UP: updated Nahr's Castle, Feather, and Witchunters in Quest Organizer +  
-FIX: inserting record on editor with list would go to last record in list  
-FIX: "doesn't have ability" only if option on universal mod  
-FIX: closing a room editor copy would load the original room editor  
-FIX: room quickjump buttons would fail if room edtor was not loaded  
-FIX: shop editor was disabling item lists on some shop types  
-NOTE: the copy/paste options on the room editors are indepdendent  
from the windows clipboard.  

++ What I've noticed is btrieve will be real slow running through  
a database for the first time in awhile, but after that it's real fast.  
+ Thanks Demosthenes  

v1.47/n (03/23/2004)  
------------------------------------------  
-NEW: hotkey (5) on map editor to toggle between new/existing rooms  
-NEW: map editor will auto get the room number of existing adjacent room  
-NEW: copy/paste options on rooms for "Advanced settings"  
-NEW: shops now show bank account/cost info for gangshops  
-UP: MMUD Explorer export update  
-UP: Updated the tooltip shown for para1 of hidden exits  
-FIX: importer could cut off 30th character in name  

v1.461/n (03/06/2004)  
------------------------------------------  
-UP: Updated MMUD Explorer export  
-UP: MMUD Explorer export won't export buffer rooms  

v1.46/n (03/03/2004)  
------------------------------------------  
-NEW: gang editor  
-UP: new format for MMUD Explorer exports  
-UP: map wont draw rooms that have already been drawn  
-UP: quest organizer now supports 2 files; default & custom  
-FIX: windows focus issues on map  
-FIX: had to save and re-click to see markup change on shops  
-FIX: class editor displayed combat wrong in list  
-FIX: spell editor wouldn’t filter  
-FIX: update file compile wouldn’t quit on error  
-FIX: one of the action lines was selecting text incorrectly  

v1.45/n (01/13/2004)  
------------------------------------------  
-NOTE: v1.45n (note the 'n') works with v1.11i through v1.11n, v1.45 is for v1.11o+  
-NEW: function on user editor to paste items/stats/spells from screen cap  
-NEW: revamped map editor to give it more cells and the features of the map  
-NEW: all editors with lists now have sortable columns  
-NEW: exp calculator tool  
-NEW: swing calculator tool  
-NEW: tool to create a list of NPC assigned monsters and their rooms  
-NEW: tool to restock all shops  
-UP: filter option on item and spell editors  
-UP: shop editor now shows all items on one page along with the adjusted cost  
-UP: all windows have finally had the tab order fixed (fixing that sucked!)  
-UP: added on option on map to hide the tooltips  
-UP: updated Group/Index and Boss/NPC lists in monster help  
-FIX: failed exit creations to new/existing rooms on map won’t change starting room  
-FIX: class/race editors were reloading the record lists on close  
-FIX: 'no line colors' and 'no color' options on map work now  
-FIX: lots of misc. small fixes  

v1.44 (12/19/2003)  
------------------------------------------  
-UP: compatible /ONLY/ with v1.11o  

v1.43 (11/26/2003)  
------------------------------------------  
-BUGFIX: wasn't exporting/importing monster's death spells  
-NEW: class editor now has sortable columns (in the works for the others editors)  
-UP: some coding changes to the mmud explorer export, that coming very soon!  
-UP: new tool to delete the buffer rooms created by the room pad tool  
-UP: changed map display to have a black background (option to be white)  
-UP: added a reset button to the monster editor to reset the kill time  
-NOTE: to create mmud explorer data files, you will need this version  

v1.42 (09/12/2003)  
------------------------------------------  
-UP: updated the NMR-Quests.txt file for the 5th evil alignment quest (thx Demonic)  
-FIX: fixed rooms not connecting on map view for some people who use 120dpi  
-FIX: fixed problem with canceling the map builder  
-FIX: fixed small problem with the map legend lines showing the wrong color  

v1.4 (09/01/2003)  
------------------------------------------  
-NEW: mapping overhaul: more options, more sizes, follow up/down exits, faster  
-NEW: can now open multiple copies of the editors  
-NEW: create a blank update file (for testing, don’t gotta compile everytime)  
-NEW: build monster group/index list, used in map to show monster regen  
-NEW: new field discovered in room DB: gang house number  
-NEW: chest cash drops fields on item editor (thx Locke Cole)  
-NEW: "BS Defense" field for monsters added (thx Tsunami)  
-NEW: "Charm Resistance" field for monsters added (thx Tsunami)  
-NEW: unknown flag on item editor changed to "robable"  
-NEW: option to search for room name/description on room editor  
-NEW: utility to pad out rooms on a given map (for shop regen and item sweeps)  
-NEW: dat call letter changer for rooms -- changes all the "WCC" references  
-NEW: database import now supports range import  
-NEW: textblock editor shows how many characters you have left on that block  
-NEW: tool to give all users a retrain  
-UP: compatible with v1.11n  
-UP: can now add/strip/changes abilities in the universal modifier  
-UP: added some elements to the universal mod like more "only if's"  
-UP: added some quests to the quest organizer  
-UP: updated some fields in the ability database (thx Tsunami)  
-UP: textblock insert will now copy the textblock you’re on  
-UP: updated what the no limited item tool excludes, see general help  
-UP: cleaned up some textblock code  
-UP: fixed a few of the room parameters  
-UP: added the uses field in rooms for items on the ground and hidden  
-UP: preview job option on importer  
-UP: forms shouldn't reload on insert/delete  
-UP: older export files can now be imported, it just skips the missing fields  
-UP: cleaned up textfile exporting a little, mainly for users  
-FIX: exporter was disabling the wrong controls on load  
-FIX: ability names on user editor were locked  
-NOTE: older export files wont work completely with this version  
-NOTE: all of the added fields above are now exported/imported  
-NOTE: big thanks to locke cole recently for working on database stuff with me  

v1.3 (05/12/2003)  
------------------------------------------  
-BUGFIX: was reading wrong value in the room DB for a room’s spell  
-NEW: export to mmud explorer (still tweaking), MMud Explorer coming soon  
-NEW: limited item list creator - scans users, rooms, and shops  
-NEW: abilities have quick links to what they refer to (like CastSP X)  
-NEW: added a previously viewed list on the textblock editor  
-NEW: can now search record number in users/monsters/spells/shops/items  
-UP: added 5th good and neutral (& last part of evil) quests to the quest organizer  
-UP: data export file now compacts after exporting to decrease the filesize  
-UP: shop editor now has quick links for the items sold  
-UP: room item search can now be canceled  
-UP: main window saves form size and won't maximize all the time  
-UP: can undock windows (doesn't work too well tho, you'll see what I mean)  
-UP: rewrote pretty much all the code on the ability editor  
-UP: lots of code rewrites for things dealing with the ability database  
-UP: optimized some loading code in users/monsters/spells/shops/items  
-UP: can now enter any part of a user/monster/spell/shop/item in search  
-UP: textblock/message/room/action editors will now 'goto' when you press enter  
-UP: textblock editor window can now be resized  
-UP: fixed some tab orders and made search/goto fields have focus on load  
-UP: added "only if <money type>" to universal for monsters  
-UP: added more reported required rooms to the room help  
-UP: map now auto-hides the unused blocks  
-UP: switched the user name and bbs id columns in the user editor  
-UP: added a bunch of quicklinks to the items/spells/rooms in the user editor  
-UP: when editing an item on a user, position won't go to the first item anymore  
-UP: user editor won't auto-save anymore, you must click save (safety measure)  
-FIX: user editor was doubling class and race list on copy or delete  


v1.2 (03/03/2003)  
------------------------------------------  
-NEW: option to use all available cpu for import/export/delete/compile/massrm/map  
-NEW: "only perform action if" options on the universal modifier  
-NEW: a pseudo tool to multiply just boss experience  
-NEW: room exit page now updates the parameters according to exit type  
-NEW: added the "retain item after uses expire" flag to item editor  
-NEW: the user editor user list is now storable by username or bbs name  
-NEW: find/find next on textblock editor  
-UP: monster attack page now updates the labels according to attack type  
-UP: tweaked the update file compiling process to hopefully make it a little faster  
-UP: added an unknown flag, and 2 unknown values to the item editor  
-UP: room editor now remembers the last room you were on  
-UP: sped up the monster/item/shop/spell list creations  
-UP: added cost as an option for items in universal mod  
-UP: added some info on making backups to the general help  
-FIX: fixed some of the forms jumping around when they load  
-FIX: monster grp/index list was cut off at the end, split it up to two pages  
-FIX: couldn't export to textfiles (thx dxpac)  
-NOTE: old export files won’t work with this version  


v1.1 (02/15/2003)  
------------------------------------------  
-UP: rewrote some code in the import and export to make it more efficient  
-FIX: class import wasn’t importing abilities (thx PanterraX!)  


v1.0 (02/14/2003)  
------------------------------------------  
-NEW: i decided to 1.0 it, i think everyone is aware it’s really always in ‘beta’  
-NEW: the 5 previous settings are available to choose from in settings  
-NEW: tool to search for an item sitting in a room (under tools - rooms)  
-UP: compatibility with v1.11m  
-UP: wrote a small tutorial about the update file and general editing (noobs, read!)  
-UP: updated the group/index list for v1.11m in monster help  
-UP: added a monster boss/NPC room list to the monster help  
-UP: cleaned up and added some info to the rooms help  
-UP: added some quests to the quest organizer, more coming soon  
-UP: slightly new format for the quest organizer, can now have 3 tree levels  
-UP: user merger now marks changed records with *changed* in the log file  
-UP: textblock editor no longer auto-loads the preview window  
-NOTE: old NMR-Quests.txt files may not work correctly with this version  


v0.7.7 BETA (01/27/03)  
------------------------------------------  
-NEW: user merge --merge two user databases  
-UP: compatibility with v1.11L  
-FIX: rooms by range for universal modifier wasn’t working (thx durin)  


v0.7.6 BETA (01/15/03)  
------------------------------------------  
-NEW: tool to strip characters off the end of the textblocks (temporary workaround)  
-UP: you can now specify your own file name for exporting (and importing)  
-UP: less buttons are disabled on the room editor when using the map editor  
-UP: using the add/insert button on the room editor will now insert a "blank" room  
-UP: added option to exporter to export using only one field for monster experience  
-FIX: adding/inserting a room wasn't clearing current room monsters  
-FIX: copying a room description didn't update map editor (thx PanterraX)  


v0.7.56 BETA (01/06/03)  
------------------------------------------  
-FIX: importer had trouble importing items  


v0.7.55 BETA (01/06/03)  
------------------------------------------  
-NEW: compatibility with v1.11k  
-UP: updated monster group/index list for v1.11k (in monster help)  


v0.7.51 BETA (01/04/03)  
------------------------------------------  
-NEW: compatibility with v1.11j  
-NEW: option in settings to switch between dat file versions  
-NEW: exporter now creates an "Info" table, shows the version of NMR and dat files  
-NEW: option to import old monsters and "convert" to new exp format (during import)  
-NEW: status bar at the bottom -shows call letters, dat file version, and if writing disabled  
-UP: monster editor now has exp 'base' and 'multiplier' fields to reflect v1.11j changes  
-UP: upon startup, NMR will check to make sure the correct dat file version is selected  
-UP: exits in room editor now have clickable buttons to jump to those rooms (thx tsunami)  
-UP: 'exp multiplier' and 'exp total' options added to universal modifier for monsters  
-FIX: class editor wasn't allowing negative number input into the ability values (thx xetox)  
-FIX: running two mass room operations consecutively caused an error (thx dxpac)  


v0.7.32 BETA (12/24/02)  
------------------------------------------  
-FIX: hopefully fixed the rest of the crash-on-exit problems  
-FIX: the exporter was checking for existing exported records all the time  


v0.7.3 BETA (12/18/02)  
------------------------------------------  
-UP: you can now cancel the update compiling and map building processes  
-UP: update comp., mass room, and import/export/deleter all show accurate progress bars  
-UP: implemented some error prevention for races/classes that didn't exist in an installation  
-UP: LOTS of work in general on the importer and exporter  
-UP: option to add to/update an existing DataExport file instead of creating a new one  
-UP: added checks for the importer to strip trailing spaces on text fields  
-UP: moved "Charm LvL" to a different column in the DataExport file (it was lost in the mix)  
-UP: added option to close all windows to the windows menu  
-FIX: fixed some crashing problems on exit  
-FIX: fixed some problems with the log file in the mass room editor  
-NOTE: old NMR-DataExport.mdb files may not work with this version  


v0.7.21 BETA (11/24/02)  
------------------------------------------  
-NEW: option on file menu to disable database writing  
-FIX: keypad movement in the map editor was broken (thx Epro)  
-FIX: fixed some possible problems with a file being in the root directory of a drive  
-FIX: fixed some problems in the exporter with the status bar panel and its record counts  


v0.7.2 BETA (11/14/02)  
------------------------------------------  
-NEW: universal modifier --perform math functions on many of the fields within the editors  
-NEW: added 5 more negate slots to the item editor (thx locke cole)  
-UP: tool tips on the map editor will now show instantly (hover mouse over a room)  
-UP: tool tips on the map now show a lot of useful information  
-UP: in-game AC and DR is calculated and shown on the item editor  
-UP: min/max/duration values now calculate as you change them in the spell editor  
-UP: as you tab or click through the fields in the room editor it will auto-select it’s contents  
-UP: tab order in room editor should be better now  
-UP: added get first/last record number buttons to db exporter  
-UP: db importer won’t log updated records anymore unless log-all is selected  
-UP: lots of code rewrites in general  
-FIX: clicking on the starting room in the map crashed everything (thx WarHawk)  
-FIX: negate spells in the item’s database weren’t being exported/imported  


v0.7.1 BETA (11/4/02)  
------------------------------------------  
-NEW: map display from the room editor, displays 1680 rooms (thx scorpion for help on this)  
-UP: map editor now updates if you make changes in the room editor and hit save  
-UP: added a clear button to clear the current room description  
-UP: when typing room descriptions it will automatically go to the next line  
-UP: MiniMap renamed to "Map Editor"  
-FIX: database exporter/importer wasn’t exporting/importing a spell’s level restriction field  


v0.7.0 BETA (09/29/02)  
------------------------------------------  
-NEW: database import!  
-NEW: quest organizer -- (props to Demonic for digging through the alignment textblocks)  
-NEW: new window menu for selecting/tiling/cascading the open windows  
-UP: edit rooms while using MiniMap  
-UP: two versions of MiniMap now, a big one and small one  
-UP: revamped database exporter, allowing for range export  
-UP: added jump buttons next to a lot of the fields, enabling a jump to a record in another  
editor  
-UP: mass room editor now logs errors (and only the critical errors) to a log file  
-UP: you can now minimize all of the windows  
-UP: missing or locked ability.mdb will no longer cripple program  
-UP: rewrote some code in the mass room editor to make it faster and more efficient  
-UP: added a Go2 button for the LinkTo field in the textblock editor  
-UP: when loading a form, it will now be given focus  
-UP: added button to recreate the settings.ini in the settings  
-FIX: fixed small bug with the select all/none in the database exporter  


v0.6.6 BETA (08/20/02)  
------------------------------------------  
-NEW: added export to access database in the database exporter  
-NEW: added a help for monsters, which is currently just my group to index list  
-UP: added a clear all to all of the user editor lists (items, keys, and spells)  
-UP: enabled double clicking on an item, key, or spell in the user editor lists to edit it  
-UP: revamped the general page in the spell editor, adding many new editing fields (thx  
atma)  
-UP: added all/none selection to the database exporter  
-UP: added a 'Jump to Last' option to auto jump to the last item in a list after an insert or  
delete  
-UP: spiffied up the search fields --pressing the right arrow key now searches for the next  
item  
-UP: added ini entries to save minimap settings, minimap position, and database export path  
-FIX: fixed the CP field in user editor to allow negative number input (thx bad_hex)  


v0.6.5 BETA (07/23/02)  
------------------------------------------  
-NEW: added a tool to reset all of the monster's last killed time  
-NEW: added the time & date stamp of when a monster was last killed to the monster editor  
-UP: changed the font on the preview window so it displays the textblocks MUCH better  
-FIX: added the "Part#" (formerly StepTo) to the textblocks (big thx to kix on noticing this  
missing)  
-FIX: the missing part# created problems deleting a range of textblocks and misc other  
things  


v0.6.4 BETA (07/22/02)  
------------------------------------------  
-NEW: added a help for messages (thx MrBlack for some help)  
-UP: added range deletion for all dats (via the database deleter)  
-UP: reorganized the data of the exported text files  
-UP: added column headings to the first line of exported text files so you know what is what  
-UP: tweaked the minimap some more, disabled the keypad when typing numbers  
-UP: added more message and text block quick-displays to the editors  
-UP: added code to write missing .ini settings instead of crashing program  
-UP: added more items not to be effected by the no limited tool (see general help)  
-FIX: fixed crash when you tried to maximize two help screens (thx kix)  
-FIX: fixed problems with a lot of the description fields  
-FIX: fixed problems with some of the name fields  


v0.6.3 BETA (07/18/02)  
------------------------------------------  
-UP: spruced up the menu with a few more shortcuts  
-UP: added 'select all' and 'select none' to the mass room editor  
-UP: added option on the MiniMap to specify the next incremented room number  
-UP: added option to auto use that incremented room number  
-FIX: fixed problem with the room type and ansi map in the mass room editor (thx dxpac)  
-FIX: fixed the messages showing up under the monsters spells/attacks (thx MrBlack)  


v0.6.2 BETA (07/16/02)  
------------------------------------------  
-UP: reconfigured cap on monster HPs, new cap should be 65535  
-FIX: fixed bug with the monster divider (it set the exp to 1!)  
-FIX: fixed some bad grammar in the room editor ;-) (thx MrBlack)  


v0.6.1 BETA (07/16/02)  
------------------------------------------  
-NEW: Added functionality to walk around and create rooms via the map editor  
-NEW: Added the ability to remove all item level restrictions  
-UP: Renamed the 'room range editor' to the 'mass room editor'  
-UP: Added the ability to delete a range of rooms in the mass room editor  
-FIX: few bug fixes  


v0.6.0 BETA  
------------------------------------------  
-NEW: compatibility with (and only with) the mod9 dat files  
-NEW: compile an update file  
-NEW: mass room editor  
-NEW: user editor  
-NEW: bank book editor  
-NEW: action editor  
-NEW: database export to text files  
-NEW: database record deleter  
-NEW: divide monster experience  
-NEW: edit multiple named dat files (wbb*.dat)  
-NEW: help with editing  
-NEW: settings menu; auto-compile, dat file locations, etc  
-UP: improved error message handling  
