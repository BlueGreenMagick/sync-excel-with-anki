###Anki Excel Sync

-   `_directory` [string]: The full path of the directory that your excel files are in. Note that `\` and `"` needs to be escaped and written as `\\`, `\"`.
-   `autosync_on_close` [bool]: If `true`, on closing Anki, 'Anki -> Excel' will happen automatically. Set it to `false` if you do not want to auto-sync on close. Recommended: `false`
-   `autosync_on_launch` [bool]: If `true`, on launching Anki, 'Excel -> Anki' will happen automatically. If a note was both modified on excel file and on another device, that modification will be overridden with Excel file. Recommended: `false`
-   `col-width` [list - integer]: Width of columns of excel files. First element becomes the width of first column, Second the width of second column, etc. If your excel file has more columns, they are set to default width. Set it to `[]` if you want to use the default width for all columns.
-   `new-deck` [string]: Name of the deck that new notes in your excel files will go into.