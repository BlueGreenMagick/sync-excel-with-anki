from PyQt5.QtWidgets import QAction
from aqt import mw
from aqt.utils import askUserDialog


def confirm_win(text="", conf="Yes", canc="Cancel", default=0):
    diag = askUserDialog(text, [conf, canc])
    diag.setDefault(default)
    ret = diag.run()
    if ret == conf:
        return True
    else:
        return False


from .sync import ExcelSync  # Prevent circular import


def create_action(name, handler):
    action = QAction(name, mw)
    action.triggered.connect(handler)
    return action


def confirm_e2a_sync():
    txt = """
<b> Excel -> Anki </b>
The Anki Cards with selected tags will be replaced by data from Excel.

Anki cards will be overwritten.
"""
    conf = confirm_win(txt, "Create", "Cancel")
    if conf:
        ExcelSync().e2a_sync()


def confirm_a2e_sync():
    txt = """
<b>Anki -> Excel</b>
Excel files will be created from existing Anki Cards with selected tags.
"""
    conf = confirm_win(txt, "Create", "Cancel")

    if conf:
        ExcelSync().a2e_sync()
        cnfg = mw.addonManager.getConfig(__name__)
        mw.addonManager.writeConfig(__name__, cnfg)


def modify_menu():
    label = "Anki -> Excel"
    action = create_action(label, confirm_a2e_sync)
    mw.form.menuTools.addAction(action)
    label = "Excel -> Anki"
    action = create_action(label, confirm_e2a_sync)
    mw.form.menuTools.addAction(action)
