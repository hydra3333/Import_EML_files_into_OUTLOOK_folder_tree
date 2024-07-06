**PRELIMINARY NOTE:    
Outlook has no menu option to import .eml files, nor any library function to open a .eml file.    
Who knows why, answer could be Microsoft could not be bothered.    
That's why the roundabout method in this vba is used, there seems to be no other known way ... nothing on the web that I could find.    
ChatGTP AI was hopeless, on a few occasions it actually fibbed about what would work, then tried to correct itself, suggested faulty logic in code, and never arrived at any workable solution. Micrsoft's CoPilot AI was much less useful even than that !**

This VBA macro in Outlook successfully imports .eml files exported from Thunderbird into a folder tree on disk,
into an Outlook folder tree in a nominated .PST having folder names replicated from the folder tree on disk.

The .eml email files on disk are moved to a "DONE" folder tree on disk on the fly, making it restartable should something go astray.

_It's the only thing I have ever seen which does this with a folder tree of .eml files, replicating the folder tree names into a .pst file and saving the .eml emails to their correct places in that folder tree._


1. **ALWAYS ALWAYS ALWAYS** first backup your .eml files and **ALL** of your Outlook data; this should never go wrong but one never knows !
2. Edit the .vba to edit/add your folder names and disk/tree root folder names
3. In outlook, ensure your new "destinatioon" .PST is opened and is named as you newly specified in the vba macro
4. Go into Outlook developer window (google how, if you need to) and add a new module and paste in your updated .vba code and save it
5. Do menu File, Options, Trust Center, Trust Center Settings, Macro Settings, Enable All Macros, OK ... remember you'll need to revert this back to a secure setting when you're done importing
6. Re-check you've specified the correct folder names in the vba code and that they exist in Outlook and on DISK where you expect to see them
7. Start the Immediate window by pressing control G
8. Start the import by clicking into the Immediate Window by typing the word `menu` and pressing Enter
9. It may take a few hours to import 5,000 or so .eml files ... the Immediate window only records the last 200 print statements, so perhaps look at it every 15 mins or so
10. When finished, re-enable macro security ... do menu File, Options, Trust Center, Trust Center Settings, Macro Settings, select a higher setting such as only signed, OK

Good Luck
