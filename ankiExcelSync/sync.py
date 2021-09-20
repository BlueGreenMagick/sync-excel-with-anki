import os
import unicodedata
import urllib.parse
import re
from operator import itemgetter

from anki import version as ankiversion
from aqt import mw
from aqt.editor import Editor
from aqt.utils import showText

from .excel import ExcelFile
from .errors import *
from .menu import confirm_win
from .template import EditorTemplate

ankiver_minor = int(ankiversion.split(".")[2])
ankiver_major = ankiversion[0:3]


class ExcelSync:
    def __init__(self):
        self.log = []
        self.config = mw.addonManager.getConfig(__name__)

    def show_log(self):
        showText(
            "\n".join(self.log), title="Excel Sync Done", minWidth=450, minHeight=300
        )

    def output_error(self, exception):
        diag, box = showText(
            exception.output_message(),
            run=False,
            type="html",
            copyBtn=True,
            title="Error",
            minWidth=450,
            minHeight=300,
        )
        diag.resize(450, 500)
        diag.show()

    def get_super_dirs(self, dirc):
        super_dirs = []
        for name in os.listdir(dirc):
            if os.path.isdir(os.path.join(dirc, name)):
                name = unicodedata.normalize("NFC", name)
                # TODO: what will happen if directory contains whitespace?
                name = " ".join(mw.col.tags.canonify([name])).strip()
                super_dirs.append(name)
        return super_dirs

    def excel_files_in_dir(self, directory):
        super_tags = self.get_super_dirs(directory)

        file_list = []
        for root, dirs, files in os.walk(directory):
            high = root
            tag = ""
            fol = None
            max_loop = 200  # fail-safe that shouldn't run
            while high != directory and max_loop > 0:
                high, fol = os.path.split(high)
                tag = fol + "::" + tag
                max_loop -= 1

            if max_loop == 0:
                raise LongDirectoryHierarchyError(root)

            for f in files:
                if f[-5:] == ".xlsx" or f[-5:] == ".xlsm" or f[-4:] == ".xls":
                    tf = f.split(".")
                    tf.pop()
                    tf = "".join(tf)
                    tag_name = tag + tf
                    file_list.append({"src": os.path.join(root, f), "tag": tag_name})
        return (file_list, super_tags)

    def prepare_field_val(self, txt):
        # from Editor.onBridgeCmd
        if ankiver_minor <= 19:
            txt = urllib.parse.unquote(txt)
        if ankiver_minor <= 27:
            # after v28, normalization optionally occurs when saving note data
            txt = unicodedata.normalize("NFC", txt)
        if ankiver_minor <= 29:
            # after v30, occurs upon calling mungeHtml
            txt = txt.replace("\x00", "")
            txt = mw.col.media.escapeImages(txt, unescape=True)

        editor_templ = EditorTemplate()  # esp for editor.mw reference
        txt = Editor.mungeHTML(editor_templ, txt)
        return txt

    # Check if note and note_data is the same (fields and tag)
    def same_note(self, note, note_data, otag, super_tags):
        fields = note_data["fields"]
        nflds = note.keys()
        for fieldnm in fields:
            if fieldnm not in nflds:
                raise FieldNameDoesNotExistError(
                    note_data["path"], note_data["row"], fieldnm, note.model()["name"]
                )
        for fieldnm in fields:
            val = fields[fieldnm]
            if val:
                val = self.prepare_field_val(val)
                if note[fieldnm] != val:
                    return False
        for tag in note.tags:
            tag = unicodedata.normalize("NFC", tag)
            tag = mw.col.tags.canonify([tag])[0]
            for super_tag in super_tags:
                if tag.lower().startswith(super_tag.lower() + "::"):
                    if tag != otag:
                        return False
        return True

    def sync_note(self, note, note_data, otag, super_tags):
        """
        note[aqt.notes.Note]: existing note data in Anki
        note_data[dictionary]: note data in Excel.
        otag[string]: note_data["tag"]
        super_tags[list]: names of top level directories(tags) to sync
        """
        fields = note_data["fields"]
        nflds = note.keys()
        for fieldnm in fields:
            if fieldnm not in nflds:
                raise FieldNameDoesNotExistError(
                    note_data["path"], note_data["row"], fieldnm, note.model()["name"]
                )
        for fieldnm in fields:
            val = fields[fieldnm]
            if val:
                val = self.prepare_field_val(val)
                if note[fieldnm] != val:
                    note[fieldnm] = val
            else:
                note[fieldnm] = ""

        # remove target tags and reapply them again
        for tag in note.tags:
            for super_tag in super_tags:
                if tag.lower().startswith(super_tag.lower() + "::"):
                    note.tags.remove(tag)
        otag = unicodedata.normalize("NFC", otag)
        note.tags += mw.col.tags.canonify([otag])
        note.flush()

    def create_note(self, note_data, tag, decknm):
        """
        note_data: {"row":int, "id":int, "fields":{"fieldName":str_val,}, "model": str_model_name}
        https://github.com/inevity/addon-movies2anki/blob/master/anki2.1mvaddon/movies2anki/movies2anki.py#L786
        """
        fpath = note_data["path"]
        row = note_data["row"]
        model_name = note_data["model"]

        model = mw.col.models.byName(model_name)  # Returns None when not exist
        if not model:  # check if model doesn't exist
            raise ModelNameDoesNotExistError(fpath, model_name)

        mw.col.models.setCurrent(model)
        note = mw.col.newNote(forDeck=False)
        # check if fldnm not exist in model
        nflds = note.keys()
        for fldnm in note_data[
            "fields"
        ]:  # in different for loop so note data is not partially updated
            if fldnm not in nflds:
                raise FieldNameDoesNotExistError(fpath, row, fldnm, model_name)
        for fldnm in note_data["fields"]:
            fldval = note_data["fields"][fldnm]
            if not fldval:  # convert NoneType to string
                fldval = ""
            note[fldnm] = fldval
        tag = unicodedata.normalize("NFC", tag)
        note.tags = mw.col.tags.canonify([tag])
        did = mw.col.decks.byName(decknm)[
            "id"
        ]  # decknm is already validated in a2e_sync
        note.model()["did"] = did

        # Check if note is valid, from method aqt.addCards.addNote
        ret = note.dupeOrEmpty()
        if ret == 1:
            self.log.extend(
                (
                    "Non-fatal: Note skipped because first field is empty."
                    "Please sync again after fixing this issue.",
                    "From row: {}, file: {}".join(note_data["row"], note_data["path"]),
                )
            )
            return None

        if "{{cloze:" in note.model()["tmpls"][0]["qfmt"]:
            if not mw.col.models._availClozeOrds(
                note.model(), note.joinedFields(), False
            ):
                self.log.extend(
                    (
                        "Non-fatal: No cloze exist in cloze note type.",
                        "Note was still added.",
                        "From row: {}, file: {}".format(
                            note_data["row"], note_data["path"]
                        ),
                    )
                )

        cards = mw.col.addNote(note)
        if not cards:

            self.log.extend(
                (
                    "NON-fatal: No cards are made from this note.",
                    "Please sync again after fixing this issue.",
                    "From row: {}, file: {}".format(
                        note_data["row"], note_data["path"]
                    ),
                )
            )

        return note.id

    def get_remove_cards_id(self, super_tags, note_ids):
        del_ids = []
        for tag in super_tags:
            card_ids = mw.col.findCards("tag:" + tag + "::*")
            card_ids.extend(mw.col.findCards("tag:" + tag))
            for card_id in card_ids:
                if mw.col.getCard(card_id).nid not in note_ids:
                    del_ids.append(card_id)
        return del_ids

    def model_data(self):
        models_all = mw.col.models.all()
        models = []
        for mdl in models_all:
            mdlcount = mw.col.models.useCount(mdl)
            fields = []
            for fld in mdl["flds"]:
                fields.append(fld["name"])
            models.append({"name": mdl["name"], "flds": fields, "count": mdlcount})
        models = sorted(models, key=itemgetter("count"), reverse=True)
        ids = []
        for model in models:
            id = model["name"][0]
            c = 1
            while id in ids:
                id = model["name"][0] + str(c)
                c += 1
            ids.append(id)
            model["id"] = id
        return models

    def backup_then_sync(self, syncfunc):
        def on_unload():
            mw.loadCollection()
            syncfunc()

        mw.unloadCollection(on_unload)

    def e2a_sync(self):
        self.backup_then_sync(self._e2a_sync)

    def a2e_sync(self):
        self.backup_then_sync(self._a2e_sync)

    def compare_notes(self, files, super_tags):
        # Open files and collect all notes
        dirc = self.dirc
        exist_note_ids = []
        modify_notes_data = []
        add_notes_data = []
        add_note_cnt = 0
        cnt = 0
        for file in files:
            mw.progress.update(label="%d / %d files opened" % (cnt, len(files)))
            cnt += 1
            add_notes_data.append([])
            tag = file["tag"]
            ef = ExcelFile(file["src"])
            ef.load_file()
            try:
                dt = ef.read_file()
                ef.close()
            except Exception as e:
                ef.close()
                raise

            for note_data in dt:
                note_data["tag"] = tag
                if note_data["id"]:
                    note_id = note_data["id"]
                    if mw.col.findNotes("nid:%s" % note_id):
                        note = mw.col.getNote(note_id)
                        note_data["exist"] = True
                        exist_note_ids.append(note_id)
                        if not self.same_note(note, note_data, tag, super_tags):
                            modify_notes_data.append(note_data)

                    # if note with given id doesn't exist
                    else:
                        note_data["exist"] = False
                        add_note_cnt += 1
                        add_notes_data[-1].append(note_data)
                # new note
                else:
                    note_data["exist"] = False
                    add_note_cnt += 1
                    path = note_data["path"]
                    add_notes_data[-1].append(note_data)

        mw.progress.update(label="Finding cards to delete")
        del_ids = self.get_remove_cards_id(super_tags, exist_note_ids)
        return (
            exist_note_ids,
            modify_notes_data,
            add_note_cnt,
            add_notes_data,
            del_ids,
        )

    def _e2a_sync(self):
        success = True
        try:
            mw.progress.start(immediate=True, label="Searching for files")
            self.log.append("Excel -> Anki")

            # Get value from config
            dirc = self.config["_directory"]
            self.dirc = dirc
            self.log.append("directory: %s" % dirc)
            decknm = self.config["new-deck"]

            # Check if valid
            if dirc == "Z:/Somedirectory you want to save excel files":
                raise DidNotConfigureDirectoryError()
            if not mw.col.decks.byName(decknm):
                raise DeckNameDoesNotExistError(decknm)

            # Get all excel file names and supertags
            files, super_tags = self.excel_files_in_dir(dirc)

            (
                exist_note_ids,
                modify_notes_data,
                add_note_cnt,
                add_notes_data,
                del_ids,
            ) = self.compare_notes(files, super_tags)

            # No need to sync if there are no notes to sync
            if len(modify_notes_data) == 0 and add_note_cnt == 0 and len(del_ids) == 0:
                mw.progress.finish()
                self.log.append("No note to sync")
                return

            # Get Confirmation
            cnfrmtxt = "\n".join(
                (
                    "{} notes total,".format(len(exist_note_ids)),
                    "{} notes to modify,".format(len(modify_notes_data)),
                    "{} notes to add,".format(add_note_cnt),
                    "{} cards to delete.".format(len(del_ids)),
                    "Proceed?",
                )
            )
            mw.progress.finish()
            cf = confirm_win(cnfrmtxt, default=0)
            if not cf:
                self.log.append("Cancelled e2a sync midway")
                return

            mw.progress.start(label="Excel -> Anki Sync")
            # Update existing notes
            cnt = 0
            for note_data in modify_notes_data:
                if cnt % 100 == 0:
                    mw.progress.update(
                        label="Updating existing notes %d / %d"
                        % (cnt, len(modify_notes_data))
                    )
                note_id = note_data["id"]
                tag = note_data["tag"]
                note = mw.col.getNote(note_id)
                self.sync_note(note, note_data, tag, super_tags)
                cnt += 1

            # Add new notes
            cnt = 0
            for note_datas in add_notes_data:
                if len(note_datas) == 0:
                    continue
                ef = ExcelFile(note_datas[0]["path"])
                ef.load_file()
                try:
                    for note_data in note_datas:
                        if cnt % 100 == 0:
                            mw.progress.update(
                                label="%d / %d cards updated" % (cnt, add_note_cnt)
                            )
                        tag = note_data["tag"]
                        note_id = self.create_note(note_data, tag, decknm)
                        if note_id != None:
                            ef.set_id(note_data["row"], note_data["fields"], note_id)
                        cnt += 1
                    ef.save()
                finally:
                    ef.close()

            # Delete cards
            mw.col.remCards(del_ids)

            self.log.extend(
                (
                    "{} note exist".format(len(exist_note_ids)),
                    "{} notes modified".format(len(modify_notes_data)),
                    "{} notes created".format(add_note_cnt),
                    "{} cards deleted".format(len(del_ids)),
                )
            )
            mw.reset()

        except AnkiExcelError as e:
            success = False
            self.output_error(e)

        except Exception as e:
            success = False
            raise e

        finally:
            if mw.progress.busy():
                mw.progress.finish()
            if success:
                self.show_log()

    def _a2e_sync(self):
        success = True
        try:
            mw.progress.start(immediate=True, label="Looking at directories")
            self.log.append("Anki -> Excel")

            # Get value from config
            col_width = self.config["col-width"]
            dirc = self.config["_directory"]
            self.dirc = dirc
            self.log.append("directory: %s" % dirc)

            # Get directories
            files, super_tags = self.excel_files_in_dir(dirc)
            totn = 0
            notes = {}  # notes by super tag name
            nids = []
            err_spetags = []

            models = self.model_data()

            # Iterate through each tag and sort notes per tag
            mw.progress.update(label="Going through all the cards")

            for tag in super_tags:
                card_ids = mw.col.findCards("tag:" + tag + "::*")
                card_ids.extend(mw.col.findCards("tag:" + tag))
                for card_id in card_ids:
                    card = mw.col.getCard(card_id)
                    note = card.note()
                    note_tag = None
                    for t in note.tags:
                        for tt in super_tags:
                            if t.startswith(tt + "::") or t == tt:
                                if note_tag is not None:
                                    raise MultipleSuperTagError(note)
                                note_tag = t

                    note_tag = str(note_tag)
                    note_tag = note_tag.split("::")

                    # warn user once when there is a tag with special characters
                    if note_tag not in err_spetags:
                        r = re.compile(r"[\s\S]*[<>:\"/|?*\\\\]")
                        for t in note_tag:
                            if r.match(t):
                                err_spetags.append(note_tag)
                                self.log.extend(
                                    (
                                        "WARNING: You should avoid use of special characters in tags,",
                                        "as your OS may not support such characters in file path.",
                                        "tag: {}".format("::".join(note_tag)),
                                    )
                                )
                                break

                    # tags such as tg::: should become just tg
                    # TODO: the below code does not seem to achieve above comment?
                    # and why above comment in the first place?
                    # filter(None, ...) is shorthand for filter(lambda x: x, ...)
                    note_tag = "::".join(filter(None, note_tag))
                    if note_tag in notes:
                        if note.id not in nids:
                            notes[note_tag].append(note)
                            nids.append(note.id)
                    else:
                        notes[note_tag] = [note]
                        nids.append(note.id)
            self.log.append("total %d tags / files" % len(notes))
            exist_file = []
            finf = 0

            # Write excel files
            for tag in notes:
                mw.progress.update(
                    label="Writing Spreadsheets %d / %d" % (finf, len(notes))
                )
                dir_tree = tag.split("::")
                dir = os.path.join(dirc, *dir_tree)
                dir += ".xlsx"
                exist_file.append(dir)
                # Write to file
                ef = ExcelFile(dir)
                ef.create_file()
                try:
                    ef.write(notes[tag], models, col_width)
                    ef.save()
                finally:
                    ef.close()

                # Logging
                totn += len(notes[tag])
                finf += 1
            self.log.append("total %d notes" % totn)
            mw.progress.update(label="Finding files to delete")

            # Delete excel files if no cards with such tag exist
            to_remove = []
            for f in files:
                f = f["src"]
                if f not in exist_file:
                    if f[-5:] == ".xlsx" or f[-5:] == ".xlsm" or f[-4:] == ".xls":
                        to_remove.append(f)

            if to_remove:
                mw.progress.finish()
                cnfrmtxt = "\n".join(
                    (
                        "{} excel files to delete.".format(len(to_remove)),
                        "Proceed with deletion?",
                    )
                )
                cf = confirm_win(cnfrmtxt, default=0)
                if cf:
                    mw.progress.start(label="Deleting redundant files")
                    for f in to_remove:
                        os.remove(f)
                        relpath = f.replace(dirc, "")
                        self.log.append("deleted file: %s" % relpath)
                else:
                    self.log.append("File(s) not deleted")

            # Finish
            mw.reset()

        except AnkiExcelError as e:
            success = False
            self.output_error(e)

        except Exception as e:
            success = False
            raise e

        finally:
            if mw.progress.busy():
                mw.progress.finish()
            if success:
                self.show_log()
