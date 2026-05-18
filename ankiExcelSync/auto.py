from aqt import gui_hooks, mw
from aqt.utils import showText
from .sync import ExcelSync
from .menu import confirm_a2e_sync


def sync_close():
    confirm_a2e_sync()


def onlaunch():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_launch"]:
        def _run_e2a():
            try:
                ExcelSync()._e2a_sync()
            except Exception as e:
                showText(str(e))
        gui_hooks.profile_did_open.append(_run_e2a)


def setclose():
    config = mw.addonManager.getConfig(__name__)
    if config["autosync_on_close"]:
        gui_hooks.profile_will_close.append(sync_close)
