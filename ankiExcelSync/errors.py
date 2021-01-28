from re import L
import traceback

class AnkiExcelError(Exception):
    def output_message(self):
        msg = "\n".join(
            (
                "<b>ERROR:</b>"
                "An error occured during sync.",
                "Depending on the error, some notes may have been synced while others weren't."
                "Please sync again after fixing the issue.",
                "",
                str(self),
                "",
                "Detailed Error Log"
                "-"*30,
                traceback.format_exc()


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
        self.filepath = filepath
        self.row = row
        self.value = value

    def __str__(self):
        msg = "\n".join(
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
        self.filepath = filepath
        self.row = row
        self.col = col
        self.value = value

    def __str__(self):
        msg = "\n".join(
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
    def __init__(self, tag, dir):
        """
        dir --- current directory

        """
        super().__init__()
        self.dir = dir

    def __str__(self):
        msg = "\n".format(
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
        self.file_path = file_path
        self.model_name = model_name

    def __str__(self):
        msg = "\n".join(
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
        self.filepath = filepath
        self.row = row
        self.field_name = field_name
        self.note_type = note_type

    def __str__(self) -> str:
        msg = "\n".format(
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
        self.deck_name = deck_name

    def __str__(self) -> str:
        msg = "\n".format(
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
        self.note = note

    def __str__(self):
        msg = "\n".join(
            (
                "Note has multiple super tags.",
                "nids: {}".format(self.note.id),
                "tags: {}".format(self.note.tags),
            )
        )


class DidNotConfigureDirectoryError(AnkiExcelError):
    def __str__(self):
        msg = "\n".format(
            (
                "ERROR: You need to set the '_directory' in addon config.",
                "to the directory your excel files reside in.",
                "Sync aborted",
            )
        )
        return msg
