import sys

from PyQt5.QtWidgets import QAction
from aqt import mw
from aqt.utils import askUserDialog

from sync import sync, sync_init

ADDON_NAME = "sync-excel-with-anki"

def create_action(name, handler):
    action = QAction(name, mw)
    action.triggered.connect(handler)
    return action

def confirm_sync():
    confirm_label = "Sync"
    cancel_label = "Cancel"
    diag = askUserDialog("""
<b> Excel -> Anki </b>
The Anki Cards with selected tags will be replaced by data from Excel.

Anki cards will be overwritten.
""", [confirm_label, cancel_label])
    diag.setDefault(1)
    ret = diag.run()
    if ret == confirm_label:
        sync()
    elif ret == cancel_label:
        return

def confirm_init_sync():
    confirm_label = "Create"
    cancel_label = "Cancel"
    diag = askUserDialog("""
<b>Anki -> Excel</b>
Excel files will be created from existing Anki Cards with selected tags.
""", [confirm_label, cancel_label])
    diag.setDefault(1)
    ret = diag.run()
    if ret == confirm_label:
        sync_init()
        cnfg = mw.addonManager.getConfig(ADDON_NAME)
        cnfg["need_init_sync"] = False
        mw.addonManager.writeConfig(ADDON_NAME, cnfg)


    elif ret == cancel_label:
        return

def modify_menu():
    config = mw.addonManager.getConfig(ADDON_NAME)
    is_first = config["need_init_sync"]
    label = "Anki -> Excel"
    action = create_action(label, confirm_init_sync)
    mw.form.menuTools.addAction(action)   
    label = "Excel -> Anki"
    action = create_action(label, confirm_sync)
    mw.form.menuTools.addAction(action)