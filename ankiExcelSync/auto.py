import sys

from anki.hooks import addHook
from aqt.main import AnkiQt
from aqt import mw
from .sync import ExcelSync

def sync_launch():
    ExcelSync().e2a_sync()
    
def sync_close():
    ExcelSync().a2e_sync()

def onlaunch():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_launch"]:    
        addHook("profileLoaded", sync_launch)

def setclose():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_close"]:   
        addHook("unloadProfile", sync_close)