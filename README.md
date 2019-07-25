# Anki Excel Sync

## Requirements:
Spreadsheet software that can edit .xlsx files (Microsoft Excel files)
Notes **cannot** have more than two tags with selected super-tag. (Super-tag is the top-most level of hierarchical tags. Super-tag of tag `science::physics::var` would be `science`)

## How To Use:
This add-on batch syncs all excel files in a directory, and its corresponding cards in anki.

After downloading the add-on, restart Anki. Then in addons menu, choose this add-on, click config and in there, set directory and deck. The excel files will be located in the directory, while the cards will go into the deck.

The excel files are created per tag in tag hierarchy. For example, a card with tag `math::trigonometry` will be located inside `DIR/math/trigonometry.xlsx` and card with tag `math` will be located inside `DIR/math.xlsx` where `DIR` is the directory you set in the config file.

In top menu, click tools, and you will see `Anki -> Excel` and `Excel -> Anki` buttons added.
Clicking `Anki -> Excel` will automatically create excel files of notes in Anki, but before you do, you need to tell the addon, notes with which super-tag you want to keep in Excel files.

If you want all the cards in `math::*` to be created in Anki, in the chosen directory, create a directory named `math`. That will let the add-on know which super-tag you want to sync notes.

After creating the directories, you can go back to Anki and click Anki -> Excel. After editing the excel files, you can click Excel -> Anki and notes will be created, deleted, or modified.

### Excel File Format
**1st row**: Names of note types.

**2nd row~**: note-type-rows. Each row represents one note type.
&nbsp; &nbsp; **1st column**: id of note type. Can be any sequence of letters.
&nbsp; &nbsp; **2nd column~**: name of fields in note type. 
&nbsp; &nbsp; &nbsp; &nbsp; *Must be the same as the name of fields of note type in Anki, or it will not work.

**n+2 row~**: note-rows. Each row represents one note.
&nbsp; &nbsp; **1st column**: id of this note's note type
&nbsp; &nbsp; **2nd column**: value for each fields. 
&nbsp; &nbsp; &nbsp; &nbsp; *The order must be the same as the order specified in the note-type-row. 
&nbsp; &nbsp; &nbsp; &nbsp; *You can leave it blank if you want.
&nbsp; &nbsp; **Last column**: note id in Anki. 
&nbsp; &nbsp; &nbsp; &nbsp; *The addon will automatically fill this cell in, so there is no need to put anything here.
