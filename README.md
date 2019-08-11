# Anki Excel Sync

An addon that makes it much easier to export and import (sync) notes from Excel. Does not affect existing note schedule.

## Note:
Supports only .xlsx files. (Microsoft Excel file)
Notes **cannot** have more than two tags with selected super-tag. (Super-tag is the top-most level of hierarchical tags. Super-tag of tag `science::physics::var` would be `science`)
If there are formulas in excel files, the formula itself will be seen as values, not computed values.
Images put directly into excel files will not be loaded into Anki. Use Anki to put images inside notes.

## Setup

After downloading the add-on, restart Anki. Edit config set `_directory`. (Tools > Addon > Anki Excel Sync > Config) The excel files will be written to the directory, and searched from the directory.

In the directory, make sub-directories with the names, the super-tags of notes you wish to sync.

## How To Use

After following the steps in Setup, clicking on `Tools > Anki -> Excel` will export anki notes to excel files, clicking on `Tools > Excel -> Anki` will import excel files into anki notes. A confirmation popup will pop up. 

If a `Excel -> Anki` sync was accidentally made, and needs to be reverted, backups are created before sync, so use backup from Anki's back up folder. `Anki -> Excel` sync cannot be reverted.

## What It Does
This add-on batch syncs all excel files in a directory, and its corresponding cards in anki.

If  `_directory` is set to `D:\\Anki\\Excel`, and the directory structure looks like this:

    Anki
    ├── Excel
    |   ├── french
    |   ├── chinese

And your tag structure looks like this:

    french : total 1330 notes
    ├── french::word : 30 notes, total 1230 notes (including sub-tags)
    |   ├── french::word::noun : 1000 notes
    |   ├── french::word::verb : 200 notes
    ├── french::grammar: 100 notes
    chinese : 400 notes

The excel files are created per tag in tag hierarchy. After running a `Anki -> Excel` sync, the directory structure will look like this:

    Anki
    ├── Excel
    |   ├── french
    |   |   ├── word
    |   |   |   ├── noun.xlsx
    |   |   |   ├── verb.xlsx
    |   |   ├── word.xlsx
    |   |   ├── grammar.xlsx
    |   ├── chinese
    |   ├── chinese.xlsx

After exporting anki notes to excel files, you can edit the excel files. You may create, modify, or delete rows. When editing is complete, click Excel -> Anki to import them in. The schedules will be preserved as long as the card id cell was not modified. 

### Excel File Format
    **1st row**: Names of note types.

    **2nd row~**: note-type-rows. Each row represents one note type.<br>
            **1st column**: designator of note type. Can be any sequence of letters.<br>
            **2nd column~**: name of fields in note type. <br>
                    *Must be the same as the name of fields of note type in Anki, or it will not work.<br>

    **n+2 row~**: note-rows. Each row represents one note.<br>
            **1st column**: designator of this note's note type specified in note-type-rows<br>
            **2nd column**: value for each fields. <br>
                    *The order must be the same as the order specified in the note-type-row. <br>
                    *You can leave it blank if you want.<br>
            **Last column**: note id in Anki. <br>
                    *The addon will automatically fill this cell in, so there is no need to put anything here.<br>
