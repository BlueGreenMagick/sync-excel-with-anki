###Anki Excel Sync

- `_directory` [string]: EDIT REQUIRED. The full path of the folder that the folders for your super-tags are in. Note that `\` need to be written as `\\`.
- `autosync_on_close` [bool]: If `true`, on profile launch, 'Anki -> Excel' will happen automatically. Set it to `false` if you do not want to auto-sync on close. Recommended: `false`
- `autosync_on_launch` [bool]: If `true`, on profile launch, 'Excel -> Anki' will happen automatically. Recommended: `false`, especially if you edit notes on other devices.
- `col-width` [list - integer]: Width of columns of excel files. First element becomes the width of first column, Second the width of second column, etc. If your excel file has more columns, they are set to default width. Set it to `[]` if you want to use the default width for all columns.
- `new-deck` [string]: Name of the deck that new notes in your excel files will go into.
- `detailed-log` [bool]: If `true`, after sync, a detailed log appears.
- `log` [bool]: If `true`, information on each sync is logged to /user_files/sync.log in addon directory.
