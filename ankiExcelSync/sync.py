import os
import sys
import unicodedata
import urllib.parse
import re
from operator import itemgetter
from datetime import datetime

from aqt import mw
from aqt.editor import Editor
from aqt.utils import showText

from .excel import ExcelFile, ExcelFileReadOnly


class ExcelSync:

    def __init__(self):
        self.log = ""
        self.simplelog = ""
        self.config = mw.addonManager.getConfig(__name__)
        self.create_ncount = 0

    def simplelog_output(self):
        showText(self.simplelog,title="Excel Sync Done",minWidth=450,minHeight=300)

    def log_output(self):
        self.log += "\n\n\n"
        dirc = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),"user_files", "sync.log")
        with open(dirc, 'a+', encoding='utf-8') as file:
            file.write(self.log)

    def excel_files_in_dir(self, directory):
        super_tags = []
        file_list = []
        for root, dirs, files in os.walk(directory):
            high = root
            tag = ""
            fol = None
            max_loop = 200 #fail-safe that shouldn't run, just in case to stop anki from being frozen
            while high != directory and max_loop > 0:
                high, fol = os.path.split(high)
                tag = fol + "::" + tag
                max_loop -= 1
            
            if max_loop == 0:
                raise Exception("""
Either you have a really long hierarchical tag, or something went wrong. Maximum level of nested tag is 200. 
dir:%s
current_high:%s
current_tag:%s
                    """%(dir,high,tag))

            if fol and fol not in super_tags:
                super_tags.append(fol)

            for f in files:
                if (f[-5:] == ".xlsx" or ".xlsm") or f[-4:] == ".xls":
                    tf = f.split('.')
                    tf.pop()
                    tf = ''.join(tf)
                    tag_name = tag + tf
                    file_list.append(
                        {"src": os.path.join(root, f), "tag": tag_name})
        return (file_list, super_tags)


    def get_super_dirs(self):
        config = self.config
        dir = config["directory"]
        super_dirs = []
        for name in os.listdir(dir):
            if os.path.isdir(os.path.join(dir, name)):
                super_dirs.append(name)
        return super_dirs


    def prepare_field_val(self, val):
        txt = urllib.parse.unquote(val)
        txt = unicodedata.normalize("NFC", txt)
        txt = Editor.mungeHTML(None, txt)
        txt = txt.replace("\x00", "")
        txt = mw.col.media.escapeImages(txt, unescape=True)
        return txt


    def sync_note(self, note, note_data, otag, super_tags):
        fields = note_data["fields"]
        nflds = note.keys()
        for fieldnm in fields:
            if fieldnm not in nflds:
                raise Exception("""
ERROR: Field name does not exist: %s
in file: %s
in row: %d
Aborted while in sync. Some notes were synced while others weren't.
Please sync again after fixing the issue.
"""%(fieldnm, note_data["path"], note_data["row"]))
        for fieldnm in fields:
            val = fields[fieldnm]
            if val:
                val = self.prepare_field_val(val)
                note[fieldnm] = val
            else:
                note[fieldnm] = ""
        for tag in note.tags:
            for super_tag in super_tags:
                if tag.lower().startswith(super_tag.lower() + "::"):
                    note.tags.remove(tag)
        note.tags.append(otag)
        note.flush()


    # note_data: {"row":int, "id":int, "fields":{"fieldName":str_val,}, "model": str_model_name}
    def create_note(self, note_data, tag, decknm):
        model_name = note_data["model"]
        model = mw.col.models.byName(model_name) # Returns None when not exist
        if not model: # check if model doesn't exist
            raise Exception("""
ERROR: Model with this name not found: '%s' 
in file:%s
in row: %d
Aborted while in sync. Please sync again after fixing the issue.
"""%(model_name, note_data["path"], note_data["row"]))
        mw.col.models.setCurrent(model)
        note = mw.col.newNote(forDeck=False)
        #check if fldnm not exist in model
        nflds = note.keys()
        for fldnm in note_data["fields"]: #in different for loop so note data is not partially updated
            if fldnm not in nflds:
                raise Exception("Field name does not exist: %s"%fldnm)
        for fldnm in note_data["fields"]:
            fldval = note_data["fields"][fldnm]
            if not fldval:  # convert NoneType to string
                fldval = ""
            note[fldnm] = fldval
        note.tags = [tag]
        did = mw.col.decks.byName(decknm)["id"] #decknm is already validated in a2e_sync
        note.model()['did'] = did
        mw.col.addNote(note)
        self.simplelog += "\ncreated note"
        self.log += "\ncreated note with id %d"%note.id
        self.create_ncount += 1
        return note.id
        # https://github.com/inevity/addon-movies2anki/blob/master/anki2.1mvaddon/movies2anki/movies2anki.py#L786


    def remove_notes(self, super_tags, note_ids):
        del_ids = []
        for tag in super_tags:
            card_ids = mw.col.findCards("tag:" + tag + "::*")
            for card_id in card_ids:
                if mw.col.getCard(card_id).nid not in note_ids:
                    del_ids.append(card_id)
        self.log += "\ndeleted cards: %d"%len(del_ids)
        self.simplelog += "\ndeleted %d card(s)"%len(del_ids)
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
        try:
            mw.progress.start(immediate=True, label="Searching for files")
            self.simplelog += "Excel -> Anki"
            self.log += "e2a sync started at %s"%datetime.now().isoformat()

            # Get value from config
            dirc = self.config["directory"]
            self.log += "\n%s"%dirc
            self.simplelog += "\ndirectory: %s"%dirc
            decknm = self.config["new-deck"]
            self.log += "\nto deck: %s"%decknm
            if not mw.col.decks.byName(decknm):
                raise Exception("FATAL: No deck exists with name %s"%decknm)

            #Get all excel file names and supertags
            files, super_tags = self.excel_files_in_dir(dirc)
            self.log += "\nnumber of files: %d"%len(files)
            self.log += "\nsuper tags: %s"%(','.join(super_tags))
            note_ids = []
            finf = 0

            for file in files:
                mw.progress.update(label="%d / %d files imported"%(finf, len(files)))
                tag = file["tag"]
                ef = ExcelFile(file["src"])
                ef.load_file()
                try:
                    notes_data = ef.read_file()

                    relpath = file["src"].replace(dirc,"")
                    self.log += "\n path: %s number of notes: %d"%(file["src"], len(notes_data))

                    for note_data in notes_data:
                        self.log += note_data["log"]
                        if note_data["id"]:
                            note_id = note_data["id"]
                            try:
                                note = mw.col.getNote(note_id)
                            except TypeError: # Note doesn't exist
                                self.log += "\ninvalid id, create card(row): %d"%note_data["row"]
                                note_id = self.create_note(note_data, tag, decknm)
                                ef.set_id(note_data["row"], note_data["fields"], note_id)
                                note = mw.col.getNote(note_id)
                            # not in try block because exception may come from this function instead
                            self.sync_note(note, note_data, tag, super_tags) 
                        else:
                            self.log += "\ncreate card(row): %d"%note_data["row"]
                            note_id = self.create_note(note_data, tag, decknm)
                            ef.set_id(note_data["row"], note_data["fields"], note_id)
                        note_ids.append(note_id)
                    ef.save()
                    ef.close()
                except Exception as e:
                    self.log += "\nError in file open: %s"%str(e)
                    ef.close()
                    raise Exception("Error occured while during sync. The problemitic excel file was not saved. \n%s"%str(e))
                finf+=1
            self.simplelog += "\ncreated %d notess"%self.create_ncount
            self.log += "\ntotal number of notes: %d"%len(note_ids)
            self.simplelog += "\ntotal %d notes"%len(note_ids)
            self.remove_notes(super_tags, note_ids)
            self.log += "\ne2a sync finished at: %s"%datetime.now().isoformat()
            mw.progress.finish()
            mw.reset()
            self.simplelog_output()
            self.log_output()
        except Exception as e:
            self.log += "ERROR! log:\n" + str(e)
            self.log_output()
            raise


    def a2e_sync(self):
        try:
            mw.progress.start(immediate=True, label="Looking at directories")
            self.simplelog += "Anki -> Excel"
            self.log += "a2e sync started at: %s"%datetime.now().isoformat()
            # Get value from config

            root_dir = self.config["directory"]
            col_width = self.config["col-width"]
            dirc = self.config["directory"]
            self.log += "\n%s"%dirc
            self.simplelog += "\ndirectory: %s"%dirc

            # Get directories
            files, super_tags = self.excel_files_in_dir(dirc)
            super_tags = self.get_super_dirs() #because super_tag from above do not detect folders without files in it.
            totn = 0
            notes = {}
            nids = []
            err_spetags = []
            models = self.model_data()
            self.log += "\nmodels done"
            mw.progress.update(label="Going through all the cards")

            # Iterate through each tag and sort notes per tag
            for tag in super_tags:
                card_ids = mw.col.findCards("tag:" + tag + "::*")
                self.log += "card count: %d"%len(card_ids)
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

                    note_tag = str(note_tag)
                    note_tag = note_tag.split("::")

                    #warn user once when there is a tag with special characters
                    if note_tag not in err_spetags:
                        r = re.compile(r"[\s\S]*[<>:\"/|?*\\\\]")
                        for t in note_tag:
                            if r.match(t):
                                err_spetags.append(note_tag)
                                self.simplelog += """\nWARNING: You should avoid use of special characters in tag, 
as your OS may not support such characters in file path.
tag: %s"""%(''.join(note_tag))
                                break

                    # tags such as tg::: should become just tg
                    note_tag = '::'.join(filter(None, note_tag))
                    if note_tag in notes:
                        if note.id not in nids:
                            notes[note_tag].append(note)
                            nids.append(note.id)
                    else:
                        notes[note_tag] = [note]
            self.log += "\ntag get card done, total tag count: %d"%len(notes)
            self.simplelog += "\ntotal %d tags / files"%len(notes)
            exist_file = []
            finf = 0

            # Write excel files
            for tag in notes:
                mw.progress.update(label="Writing Spreadsheets %d / %d"%(finf, len(notes)))
                dir_tree = tag.split("::")
                dir = os.path.join(root_dir, *dir_tree)
                dir += ".xlsx"
                exist_file.append(dir)

                #Write to file
                ef = ExcelFile(dir)
                ef.create_file()
                try:
                    ef.write(notes[tag], models, col_width)
                    ef.save()
                    ef.close()
                except Exception as e:
                    self.log += "\nError in file open: %s"%str(e)
                    ef.close()
                    raise Exception("Error occured while creating excel file. \n%s"%str(e))

                #Logging
                totn += len(notes[tag])
                self.log += "\ndone dir:%s note-count: %d"%(dir, len(notes[tag]))
                finf += 1
            self.log += "\nupdating excel done"
            self.log += "\ntotal notes: %d"%totn
            self.simplelog += "\ntotal %d notes"%totn
            mw.progress.update(label="Deleting redundant files")

            #Delete files if no cards with such tag exist
            for f in files:
                f = f["src"]
                if f not in exist_file:
                    os.remove(f)
                    self.log += "\ndeleted file: %s"%f
                    relpath = dir.replace(f,"")
                    self.simplelog += "\ndeleted file: %s"%relpath

            #Finish            
            self.log += "\na2e sync finished at: %s"%datetime.now().isoformat()
            mw.progress.finish()
            mw.reset()
            self.simplelog_output()
            self.log_output()

        except Exception as e:
            self.log += "\n" + str(e)
            self.log_output()
            raise
