from anki.hooks import addHook
from aqt.main import AnkiQt
from aqt import mw
from .sync import ExcelSync


def sync_launch(self, onsuccess=None):
    self.loadCollection()
    ExcelSync()._e2a_sync()
    # from method aqt.AnkiQt.unloadCollection
    def callback():
        self.setEnabled(False)
        self._unloadCollection()
        self.aes_old_load_profile(onsuccess)

    self.closeAllWindows(callback)


def sync_close():
    ExcelSync().a2e_sync()


def onlaunch():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_launch"]:
        AnkiQt.aes_old_load_profile = AnkiQt.loadProfile
        AnkiQt.loadProfile = sync_launch


def setclose():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_close"]:
        addHook("unloadProfile", sync_close)
