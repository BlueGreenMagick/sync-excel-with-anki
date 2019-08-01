###Anki Excel Sync

- `_directory` [string]: EDIT REQUIRED. The full path of the folder that the folders for your super-tags are in.
- `autosync_on_close` [bool]: If `true`, on profile launch, 'Anki -> Excel' will happen automatically. Set it to `false` if you do not want to auto-sync on close.
- `autosync_on_launch` [bool]: If `true`, on profile launch, 'Excel -> Anki' will happen automatically.
- `col-width` [list - integer]: Width of the columns of excel files. First element becomes the width of first column, Second the width of second column, etc. If your excel file has more columns, they are set to default width. Set it to `[]` if you want to use the default width for all columns.
- `new-deck` [string]: Name of the deck that new notes in your excel files will go into.
