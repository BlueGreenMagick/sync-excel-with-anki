#Don't use 'format document' on this file: dir needs to be appended to path before import
import sys
import os

sys.path.append(os.path.dirname(__file__))


from menu import modify_menu

ADDON_NAME = "sync-excel-with-anki"
modify_menu()
