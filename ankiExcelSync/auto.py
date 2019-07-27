import sys

from anki.hooks import addHook
from aqt.main import AnkiQt
from aqt import mw
from .sync import ExcelSync


#Auto launch on close
def new_unloadprofexit(self):
    onclose()
    self.orig_unloadProfileAndExit()

AnkiQt.orig_unloadProfileAndExit = AnkiQt.unloadProfileAndExit
AnkiQt.unloadProfileAndExit = new_unloadprofexit

#On Launch
def new_maybeAutoSync(self):
    return

def real_sync_launch():
    ExcelSync().e2a_sync()
    if (not mw.pm.profile['syncKey']
            or not mw.pm.profile['autoSync']
            or mw.safeMode
            or mw.restoringBackup):
            return
    mw.onSync()
    ExcelSync().a2e_sync()
    
def real_sync_close():
    ExcelSync().a2e_sync()
    if (not mw.pm.profile['syncKey']
        or not mw.pm.profile['autoSync']
        or mw.safeMode
        or mw.restoringBackup):
        return
    mw.onSync()

def onlaunch():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_launch"]:    
        AnkiQt.maybeAutoSync = new_maybeAutoSync
        addHook("profileLoaded",real_sync_launch)



def onclose():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_close"]:   
        addHook("unloadProfile",real_sync_close)