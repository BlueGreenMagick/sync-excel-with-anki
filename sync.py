import os
import sys
from operator import itemgetter
import unicodedata
import urllib.parse

from aqt import mw
from aqt.editor import Editor
from aqt.utils import showInfo

from excel import ExcelFile, ExcelFileReadOnly

ADDON_NAME = "sync-excel-with-anki"


class ExcelSync:

    def __init__(self):
        self.log = ""
        self.config = mw.addonManager.getConfig(ADDON_NAME)

    def excel_files_in_dir(self, directory):
        high_tags = []
        file_list = []
        for root, dirs, files in os.walk(directory):
            high = root
            tag = ""
            fol = None
            while high != directory:
                high, fol = os.path.split(high)
                tag = fol + "::" + tag
            if fol and fol not in high_tags:
                high_tags.append(fol)

            for f in files:
                if (f[-5:] == ".xlsx" or ".xlsm") or f[-4:] == ".xls":
                    tf = f.split('.')
                    tf.pop()
                    tf = ''.join(tf)
                    tag_name = tag + tf
                    file_list.append(
                        {"src": os.path.join(root, f), "tag": tag_name})
        return (file_list, high_tags)


    def get_excel_file_names(self, dirc):
        files, high_tags = self.excel_files_in_dir(dirc)
        return (files, high_tags)


    def get_high_dirs(self):
        config = self.config
        dir = config["directory"]
        high_dirs = []
        for name in os.listdir(dir):
            if os.path.isdir(os.path.join(dir, name)):
                high_dirs.append(name)
        return high_dirs


    def prepare_field_val(self, val):
        txt = urllib.parse.unquote(val)
        txt = unicodedata.normalize("NFC", txt)
        txt = Editor.mungeHTML(None, txt)
        txt = txt.replace("\x00", "")
        txt = mw.col.media.escapeImages(txt, unescape=True)
        return txt


    def sync_note(self, note, note_data, otag, high_tags):
        #sys.stderr.write("\ntag:" + otag)
        fields = note_data["fields"]
        for field in fields:
            val = fields[field]
            if val:
                val = self.prepare_field_val(val)
                note[field] = val
            else:
                note[field] = ""
        for tag in note.tags:
            for high_tag in high_tags:
                if tag.lower().startswith(high_tag.lower() + "::"):
                    note.tags.remove(tag)
        note.tags.append(otag)
        note.flush()


    # note_data: {"row":int, "id":int, "fields":{"fieldName":str_val,}, "model": str_model_name}
    def create_note(self, note_data, tag, decknm):
        model_name = note_data["model"]
        model = mw.col.models.byName(model_name)
        #sys.stderr.write("\nmodel:" + model_name)

        if not model:
            raise "Model Not Found"
        mw.col.models.setCurrent(model)
        note = mw.col.newNote(forDeck=False)
        for fldnm in note_data["fields"]:
            fldval = note_data["fields"][fldnm]
            #sys.stderr.write("\nfield:" + str(fldnm) + "content:" +str(fldval))
            if not fldval:  # convert NoneType to string
                fldval = ""
            note[fldnm] = fldval
        note.tags = [tag]
        did = mw.col.decks.id(decknm)
        note.model()['did'] = did
        mw.col.addNote(note)
        return note.id
        # https://github.com/inevity/addon-movies2anki/blob/master/anki2.1mvaddon/movies2anki/movies2anki.py#L786


    def remove_notes(self, high_tags, note_ids):
        del_ids = []
        sys.stderr.write("\ntags:")
        for tag in high_tags:
            sys.stderr.write(tag + ',')
            card_ids = mw.col.findCards("tag:" + tag + "::*")
            for card_id in card_ids:
                if mw.col.getCard(card_id).nid not in note_ids:
                    del_ids.append(card_id)
        sys.stderr.write("\ndeleted cards:" + str(len(del_ids)))
        mw.col.remCards(del_ids)


    def model_data(self):
        models_all = mw.col.models.all()
        models = []
        for mdl in models_all:
            mdlcount = mw.col.models.useCount(mdl)
            fields = []
            for fld in mdl["flds"]:
                fields.append(fld["name"])
            models.append({"name": mdl["name"], "flds": fields, "count": mdlcount})
        models = sorted(models, key=itemgetter('count'), reverse=True)
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


    def e2a_sync(self):
        mw.progress.start(immediate=True, label="Searching for files")
        note_ids = []
        dirc = self.config["directory"]
        decknm = self.config["new-deck"]
        files, high_tags = self.get_excel_file_names(dirc)
        sys.stderr.write("\nnumber of files: " + str(len(files)))
        finf = 0
        for file in files:
            mw.progress.update(label="%d / %d files imported"%(finf, len(files)))
            sys.stderr.write("\n path: " + file["src"])
            tag = file["tag"]
            ef = ExcelFile(file["src"])
            ef.load_file()
            notes_data = ef.read_file()
            sys.stderr.write("\nnumber of notes: " + str(len(notes_data)))
            for note_data in notes_data:
                if note_data["id"]:
                    note_id = note_data["id"]
                    try:
                        note = mw.col.getNote(note_id)
                    except:
                        sys.stderr.write(
                            "\ninvalid id, create card(row): " + str(note_data["row"]))
                        note_id = self.create_note(note_data, tag, decknm)
                        ef.set_id(note_data["row"], note_data["fields"], note_id)
                        note = mw.col.getNote(note_id)
                    self.sync_note(note, note_data, tag, high_tags)
                else:
                    sys.stderr.write("\ncreate card(row): " + str(note_data["row"]))
                    note_id = self.create_note(note_data, tag, decknm)
                    ef.set_id(note_data["row"], note_data["fields"], note_id)
                note_ids.append(note_id)
            ef.save()
            ef.close()
            finf+=1
        sys.stderr.write("\ntotal number of notes: " + str(len(note_ids)))
        self.remove_notes(high_tags, note_ids)
        sys.stderr.write("\ndone")
        mw.progress.finish()
        mw.reset()


    def a2e_sync(self):
        mw.progress.start(immediate=True, label="Looking at directories")
        root_dir = self.config["directory"]
        col_width = self.config["col-width"]
        dirc = self.config["directory"]
        files, high_tags = self.get_excel_file_names(dirc)
        high_tags = self.get_high_dirs() #because high_tag from above do not detect folders without files in it.
        notes = {}
        nids = []
        models = self.model_data()
        sys.stderr.write("\nmodels done")
        mw.progress.update(label="Going through all the cards")
        for tag in high_tags:
            card_ids = mw.col.findCards("tag:" + tag + "::*")
            for card_id in card_ids:
                card = mw.col.getCard(card_id)
                note = card.note()
                if len(note.tags) == 1:
                    note_tag = note.tags[0]
                else:
                    tc = 0
                    for t in note.tags:
                        if t.startswith(tag + "::"):
                            note_tag = t
                            tc += 1
                    if tc > 1:
                        tstr = ','.join(note.tags)
                        raise Exception("""More than one selected super-tag: %s 
    Aborted sync. No excel files modified."""%tstr)
                if not note_tag:
                    continue

                note_tag = str(note_tag)
                if note_tag in notes:
                    if note.id not in nids:
                        notes[note_tag].append(note)
                        nids.append(note.id)
                else:
                    notes[note_tag] = [note]
        sys.stderr.write("\ntag get card done")
        exist_file = []
        finf = 0
        for tag in notes:
            mw.progress.update(label="Writing Spreadsheets %d / %d"%(finf, len(notes)))
            sys.stderr.write("\ndone tag" + tag)
            dir_tree = tag.split("::")
            dir_tree = list(filter(None, dir_tree))
            dir = os.path.join(root_dir, *dir_tree)
            dir += ".xlsx"
            exist_file.append(dir)
            ef = ExcelFile(dir)
            ef.create_file()
            ef.write(notes[tag], models, col_width)
            ef.save()
            ef.close()
            finf += 1
        sys.stderr.write("\nupdating excel done")
        mw.progress.update(label="Deleting redundant files")
        for f in files:
            f = f["src"]
            if f not in exist_file:
                os.remove(f)
                sys.stderr.write("\ndeleted file: " + f)
        sys.stderr.write("\ndone")
        mw.progress.finish()
        mw.reset()
