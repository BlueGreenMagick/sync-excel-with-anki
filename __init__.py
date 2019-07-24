import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from menu import modify_menu

ADDON_NAME = "sync-excel-with-anki"
modify_menu()



"""

The reason for this addon instead of importing through csv:
When you want to edit a card, or add a card, to preserve existing cards in Anki, while updating excel.

The workflow using this addon:

Create an excel file with all the cards
Import them into Anki automatically
When you want to edit a card, delete a card, Should it launch Excel instead?
So there will be no sync from Anki to Excel, only from Excel to Anki


"""

"""
Later, I need to implement different note fields existing inside one excel file.
There would be several fields row for each note type, and a short name for the note type.
That name will be in the first column, and id will go into the last.
"""

"""

for sync

sync:
    check filetype is excel
    check file is valid
    if note_id is not empty:
        if note with note_id exist:
            check if they are same and modify if not
        else:
            create note and modift note_id field in excel
    if empty:
        create note in excel        
    check if there are notes in anki that are not in excel and delete them

"""

"""
How an excel file would look like:

Basic | Cloze |
Basic | Front | Back | Extra | Comment |b
Cloze | Front | Extra| Else  |         |c
10001 | this? | that | ha    | comments|b
10002 |b{c1:a}| that |       |         |c



"""


