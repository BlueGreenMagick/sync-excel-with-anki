from re import L
import traceback
import html


class AnkiExcelError(Exception):
    def output_message(self):
        msg = "<br>".join(
            (
                "<b>ERROR!</b>",
                "An error occured during sync.",
                "Depending on the error, some notes may have been synced while others weren't.",
                "Please sync again after fixing the issue.",
                "",
                "",
                "<b>Error Log</b>",
                "-" * 50,
                str(self),
                "",
                "",
                "<b>Detailed Error Log</b>",
                "-" * 50,
                html.escape(traceback.format_exc()).replace("\n", "<br>"),
            )
        )
        return msg + str(self)


class InvalidModelDesignatorError(AnkiExcelError):
    def __init__(self, filepath, row, value):
        """
        (filepath, row, value)

        Model designator on notes row
         cannot be found on the headers row
        """
        super().__init__()
        self.filepath = html.escape(filepath)
        self.row = row
        self.value = html.escape(value)

    def __str__(self):
        msg = "<br>".join(
            (
                "ERROR: Invalid note type designator in",
                f"file: {self.filepath}",
                f"row: {self.row}",
                f"designator: {self.value}",
                "",
                "The sync was stopped mid-way.",
                "Please run it again after editing the problem file.",
            )
        )
        return msg


class CannotWriteValueError(AnkiExcelError):
    """
    (filepath, row, column, value)

    Error occured when writing value to excel file cell
    """

    def __init__(self, filepath, row, col, value):
        super().__init__()
        self.filepath = html.escape(filepath)
        self.row = row
        self.col = col
        self.value = html.escape(value)

    def __str__(self):
        msg = "<br>".join(
            (
                "ERROR: Could not write value to file",
                f"file: {self.filepath}",
                f"row: {self.row}",
                f"column: {self.col}",
                f"value: {self.value}",
                "",
                "The sync was stopped mid-way.",
                "Please run it again after fixing the problem.",
            )
        )
        return msg


class LongDirectoryHierarchyError(AnkiExcelError):
    def __init__(self, dir):
        """
        dir --- current directory

        """
        super().__init__()
        self.dir = html.escape(dir)

    def __str__(self):
        msg = "<br>".format(
            (
                "Either you have a really long hierarchical tag, or something went wrong.",
                "Maximum level of nested tag(directory) is 200.",
                "dir: {}".format(self.dir),
            )
        )
        return msg


class ModelNameDoesNotExistError(AnkiExcelError):
    def __init__(self, file_path, model_name):
        super().__init__()
        self.file_path = html.escape(file_path)
        self.model_name = html.escape(model_name)

    def __str__(self):
        msg = "<br>".join(
            (
                "ERROR: Could not find note type with given name.",
                "Note type: {}".format(self.model_name),
                "in file: {}".format(self.file_path),
            )
        )
        return msg


class FieldNameDoesNotExistError(AnkiExcelError):
    def __init__(self, filepath, row, field_name, note_type):
        super.__init__()
        self.filepath = html.escape(filepath)
        self.row = row
        self.field_name = html.escape(field_name)
        self.note_type = html.escape(note_type)

    def __str__(self) -> str:
        msg = "<br>".format(
            (
                "ERROR: Field name does not exist.",
                "Field name: {}".format(self.field_name),
                "Note type: {}".format(self.note_type),
                "in file: {}".format(self.filepath),
                "in row: {}".format(self.row),
                "Aborted while in sync. Some notes were synced while others weren't.",
                "Please sync again after fixing the issue.",
            )
        )
        return msg


class DeckNameDoesNotExistError(AnkiExcelError):
    def __init__(self, deck_name):
        super.__init__()
        self.deck_name = html.escape(deck_name)

    def __str__(self) -> str:
        msg = "<br>".format(
            (
                "ERROR: Deck does not exist.",
                "Deck name: {}".format(self.deck_name),
                "Aborted while in sync. Some notes were synced while others weren't.",
                "Please sync again after fixing the issue.",
            )
        )
        return msg


class MultipleSuperTagError(AnkiExcelError):
    def __init__(self, note):
        self.id = note.id
        self.tags = html.escape(self.tags)

    def __str__(self):
        msg = "<br>".join(
            (
                "Note has multiple super tags.",
                "nids: {}".format(self.id),
                "tags: {}".format(self.tags),
            )
        )


class DidNotConfigureDirectoryError(AnkiExcelError):
    def __str__(self):
        msg = "<br>".format(
            (
                "ERROR: You need to set the '_directory' in addon config.",
                "to the directory your excel files reside in.",
                "Sync aborted",
            )
        )
        return msg
