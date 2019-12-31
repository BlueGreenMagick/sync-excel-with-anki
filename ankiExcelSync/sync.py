import os
import unicodedata
import urllib.parse
import re
from operator import itemgetter
from datetime import datetime

from aqt import mw
from aqt.editor import Editor
from aqt.utils import showText, tooltip

from .excel import ExcelFile, ExcelFileReadOnly
from .menu import confirm_win

class ExcelSync:

    def __init__(self):
        self.log = ""
        self.simplelog = ""
        self.config = mw.addonManager.getConfig(__name__)
        self.log_has_error = False
        self.tooltip_log = ""

    def simplelog_output(self):
        if self.config["detailed-log"] or self.log_has_error:
            showText(self.simplelog,title="Excel Sync Done",minWidth=450,minHeight=300)
        elif self.tooltip_log:
            tooltip(self.tooltip_log)
        else:
            tooltip("Sync succesful")

    def log_output(self):
        if self.config["log"]:
            self.log += "\n\n\n"
            dirc = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),"user_files")
            path = os.path.join(dirc, "sync.log")
            if not os.path.exists(dirc):
                os.makedirs(dirc)
            with open(path, 'a+', encoding='utf-8') as file:
                file.write(self.log)

    def get_super_dirs(self,dirc):
        super_dirs = []
        for name in os.listdir(dirc):
            if os.path.isdir(os.path.join(dirc, name)):
                name = unicodedata.normalize("NFC", name)
                name = mw.col.tags.canonify([name])[0]
                super_dirs.append(name)
        return super_dirs

    def excel_files_in_dir(self, directory):
        super_tags = self.get_super_dirs(directory)

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

            for f in files:
                if f[-5:] == ".xlsx" or f[-5:] == ".xlsm" or f[-4:] == ".xls":
                    tf = f.split('.')
                    tf.pop()
                    tf = ''.join(tf)
                    tag_name = tag + tf
                    file_list.append(
                        {"src": os.path.join(root, f), "tag": tag_name})
        return (file_list, super_tags)

    def prepare_field_val(self, val):
        txt = urllib.parse.unquote(val)
        txt = unicodedata.normalize("NFC", txt)
        txt = Editor.mungeHTML(None, txt)
        txt = txt.replace("\x00", "")
        txt = mw.col.media.escapeImages(txt, unescape=True)
        return txt

    #Check if note and note_data is the same (fields and tag)
    def same_note(self, note, note_data, otag, super_tags):
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
                if note[fieldnm] != val:
                    note[fieldnm] = val
            else:
                note[fieldnm] = ""
        for tag in note.tags:
            for super_tag in super_tags:
                if tag.lower().startswith(super_tag.lower() + "::"):
                    note.tags.remove(tag)
        tag = unicodedata.normalize("NFC", tag)
        note.tags += mw.col.tags.canonify([tag])
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
        tag = unicodedata.normalize("NFC", tag)
        note.tags = mw.col.tags.canonify([tag])
        did = mw.col.decks.byName(decknm)["id"] #decknm is already validated in a2e_sync
        note.model()['did'] = did

        #Check if note is valid, from method aqt.addCards.addNote
        ret = note.dupeOrEmpty()
        if ret == 1:
            msg = """Non-fatal: Note skipped because first field empty. Please sync again after fixing this issue.
            From row: %d, file: %s
            """%(note_data["row"], note_data["path"])
            self.log_has_error = True
            self.log += msg
            self.simplelog += msg
        
        if '{{cloze:' in note.model()['tmpls'][0]['qfmt']:
            if not mw.col.models._availClozeOrds(
                    note.model(), note.joinedFields(), False):
                msg = """Non-fatal: No cloze exist in cloze note type. Note was still added.
                From row: %d, file: %s
                """%(note_data["row"], note_data["path"])
                self.log_has_error = True
                self.log += msg
                self.simplelog += msg

        cards = mw.col.addNote(note)
        if not cards:
            msg = """NON-fatal: No cards are made from this note. Please sync again after fixing this issue.
            From row: %d, file: %s
            """%(note_data["row"], note_data["path"])
            self.log += msg
            self.simplelog += msg
            self.log_has_error = True        

        self.log += "\ncreated note with id %d"%note.id
        return note.id
        # https://github.com/inevity/addon-movies2anki/blob/master/anki2.1mvaddon/movies2anki/movies2anki.py#L786


    def get_remove_cards_id(self,super_tags,note_ids):
        del_ids = []
        for tag in super_tags:
            card_ids = mw.col.findCards("tag:" + tag + "::*")
            card_ids += mw.col.findCards("tag:" +tag)
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

    def backup_then_sync(self, syncfunc):
        def on_unload():
            mw.loadCollection()
            syncfunc()
        mw.unloadCollection(on_unload)

    def e2a_sync(self):
        self.backup_then_sync(self._e2a_sync)
    
    def a2e_sync(self):
        self.backup_then_sync(self._a2e_sync)


    def compare_notes(self, files, super_tags ):
        #Open files and collect all notes
            dirc = self.dirc
            exist_note_ids = []
            modify_notes_data = []
            add_notes_data = []
            add_note_cnt = 0
            cnt = 0
            for file in files:
                mw.progress.update(label="%d / %d files opened"%(cnt, len(files)))
                cnt+=1
                add_notes_data.append([])
                tag = file["tag"]
                ef = ExcelFile(file["src"])
                ef.load_file()
                try:
                    dt = ef.read_file()
                    ef.close()
                except Exception as e:
                    self.log += "ERROR: cannot read file.\n%s"%str(e)
                    ef.close()
                    raise Exception(str(e))

                relpath = file["src"].replace(dirc,"")
                self.log += "\n path: %s number of notes: %d"%(relpath, len(dt))

                for note_data in dt:
                    self.log += note_data["log"]
                    note_data["tag"] = tag
                    if note_data["id"]:
                        note_id = note_data["id"]
                        if mw.col.findNotes("nid:%s" % note_id):
                            note = mw.col.getNote(note_id)
                            note_data["exist"] = True
                            exist_note_ids.append(note_id)
                            if not self.same_note(note, note_data, tag, super_tags):
                                modify_notes_data.append(note_data)

                        #if note with given id doesn't exist
                        else:
                            self.log += "\ninvalid note id"
                            note_data["exist"] = False
                            add_note_cnt += 1
                            add_notes_data[-1].append(note_data)
                    #new note
                    else:
                        note_data["exist"] = False
                        add_note_cnt += 1
                        path = note_data["path"]
                        add_notes_data[-1].append(note_data)

            mw.progress.update(label="Finding cards to delete")
            del_ids = self.get_remove_cards_id(super_tags, exist_note_ids)
            return (exist_note_ids, modify_notes_data, add_note_cnt, add_notes_data, del_ids)

    def _e2a_sync(self):
        try:
            mw.progress.start(immediate=True, label="Searching for files")
            self.simplelog += "Excel -> Anki"
            self.log += "e2a sync started at %s"%datetime.now().isoformat()

            # Get value from config
            dirc = self.config["_directory"]
            self.dirc = dirc
            self.log += "\n%s"%dirc
            self.simplelog += "\ndirectory: %s"%dirc
            decknm = self.config["new-deck"]
            self.log += "\nto deck: %s"%decknm
            
            #Check if valid
            if dirc == "Z:/Somedirectory you want to save excel files":
                msg = "ERROR: You need to set the directory for your excel files, in addon config.\nSync aborted"
                raise Exception(msg)
            if not mw.col.decks.byName(decknm):
                raise Exception("ERROR: No deck exists with name %s"%decknm)

            #Get all excel file names and supertags
            files, super_tags = self.excel_files_in_dir(dirc)
            self.log += "\nnumber of files: %d"%len(files)
            self.log += "\nsuper tags: %s"%(','.join(super_tags))

            exist_note_ids, modify_notes_data, add_note_cnt, add_notes_data, del_ids = self.compare_notes(files, super_tags)

            #No need to sync if there are no notes to sync
            if len(modify_notes_data) == 0 and add_note_cnt == 0 and len(del_ids) == 0:
                mw.progress.finish()
                self.simplelog += "\nNo note to sync"
                self.tooltip_log = "No note to sync"
                self.simplelog_output()
                self.log += "\nNo note to sync, finish at %s"%datetime.now().isoformat()
                self.log_output()
                return

            #Get Confirmation
            cnfrmtxt = """%d notes total,
%d notes to modify,
%d notes to add,
%d cards to delete.
Proceed?
"""%(len(exist_note_ids), len(modify_notes_data),add_note_cnt,len(del_ids))
            self.log += ("\n" + cnfrmtxt)
            cf = confirm_win(cnfrmtxt,default=0)
            if not cf:
                self.simplelog += "\nCancelled e2a sync midway"
                self.log += "\nCancelled e2a sync midway"
                self.log_output()
                self.simplelog_output()
                mw.progress.finish()
                return
            
            #Update existing notes
            cnt = 0
            for note_data in modify_notes_data:
                if cnt % 100 == 0:
                    mw.progress.update(label="Updating existing notes %d / %d"%(cnt,len(modify_notes_data)))
                note_id = note_data["id"]
                tag = note_data["tag"]
                note = mw.col.getNote(note_id)
                self.sync_note(note, note_data, tag, super_tags)
                cnt += 1
            
            #Add new notes
            cnt = 0
            for note_datas in add_notes_data:
                if len(note_datas) == 0:
                    continue
                ef = ExcelFile(note_datas[0]["path"])
                ef.load_file()
                try:
                    for note_data in note_datas:
                        if cnt%100 == 0:
                            mw.progress.update(label="%d / %d cards updated"%(cnt, add_note_cnt))
                        tag = note_data["tag"]
                        note_id = self.create_note(note_data, tag, decknm)
                        ef.set_id(note_data["row"], note_data["fields"], note_id)
                        cnt += 1
                    ef.save()
                    ef.close()
                except Exception as e:
                    self.log += "\nError in updating note id.\n%s"%str(e)
                    ef.close()
                    raise Exception("Error occured while reading file. File was not saved. Please sync again after fixing the issue.\n%s"%str(e))

            #Delete cards
            mw.col.remCards(del_ids)

            #log
            logtxt = """
%d note exist
%d notes modified
%d notes created
%d cards deleted
"""%(len(exist_note_ids), len(modify_notes_data), add_note_cnt, len(del_ids))
            self.simplelog += logtxt
            self.log += logtxt
            self.log += "\ne2a sync finished at: %s"%datetime.now().isoformat()

            #Finish sync    
            mw.progress.finish()
            mw.reset()
            self.simplelog_output()
            self.log_output()
        except Exception as e:
            self.log_has_error = True
            self.log += "ERROR! log:\n" + str(e)
            self.log_output()
            raise e

    def _a2e_sync(self):
        try:
            mw.progress.start(immediate=True, label="Looking at directories")
            self.simplelog += "Anki -> Excel"
            self.log += "a2e sync started at: %s"%datetime.now().isoformat()
            
            # Get value from config
            col_width = self.config["col-width"]
            dirc = self.config["_directory"]
            self.dirc = dirc
            self.log += "\n%s"%dirc
            self.simplelog += "\ndirectory: %s"%dirc
            
            # Get directories
            files, super_tags = self.excel_files_in_dir(dirc)
            totn = 0
            notes = {}
            nids = []
            err_spetags = []
            models = self.model_data()
            self.log += "\nmodels done"
            
            # Iterate through each tag and sort notes per tag
            mw.progress.update(label="Going through all the cards")
            for tag in super_tags:
                card_ids = mw.col.findCards("tag:" + tag + "::*")
                card_ids += mw.col.findCards("tag:" + tag)
                self.log += "card count: %d"%len(card_ids)
                for card_id in card_ids:
                    card = mw.col.getCard(card_id)
                    note = card.note()
                    if len(note.tags) == 1:
                        note_tag = note.tags[0]
                    else:
                        tc = 0
                        for t in note.tags:
                            for tt in super_tags:
                                if t.startswith(tt + "::") or t == tt:
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
                                txt =  """\nWARNING: You should avoid use of special characters in tag, 
as your OS may not support such characters in file path.
tag: %s"""%(''.join(note_tag))
                                self.simplelog += txt
                                self.log += txt
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
                dir = os.path.join(dirc, *dir_tree)
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
            self.log += "\ntotal notes: %d"%totn
            self.simplelog += "\ntotal %d notes"%totn
            mw.progress.update(label="Deleting redundant files")

            #Delete excel files if no cards with such tag exist
            to_remove = []
            for f in files:
                f = f["src"]
                if f not in exist_file:
                    if f[-5:] == ".xlsx" or f[-5:] == ".xlsm" or f[-4:] == ".xls":
                        to_remote.append(f)

                cnfrmtxt = """%d excel files to delete.
Proceed with deletion?
"""%(len(to_remove))
            self.log += ("\n" + cnfrmtxt)
            cf = confirm_win(cnfrmtxt, default=0)
            if cf:
                for f in to_remove:
                    os.remove(f)
                    self.log += "\ndeleted file: %s"%f
                    relpath = f.replace(dirc,"")
                    self.simplelog += "\ndeleted file: %s"%relpath
            else:
                self.simplelog += "File(s) not deleted"
                self.log += "File(s) not deleted"

            #Finish            
            self.log += "\na2e sync finished at: %s"%datetime.now().isoformat()
            mw.progress.finish()
            mw.reset()
            self.simplelog_output()
            self.log_output()

        except Exception as e:
            self.log_has_error = True
            self.log += "\n" + str(e)
            self.log_output()
            raise e