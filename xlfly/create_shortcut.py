"""
This is not a general solution, re-requirements:
1. You have installed python
2. Installed xlfly
3. Added python path to windows environment variable
"""

import os
import sys
import win32com.client


# TODO: add this to post pip install
def create_shortcut():
    python_path = sys.prefix
    pythonw_path = os.path.join(python_path, "pythonw.exe")
    curr_path = os.path.dirname(__file__)
    icon_path = os.path.join(curr_path, "icon.ico")
    script_path = os.path.join(curr_path, "app.py")

    start_menu_folder = os.path.join(
        os.getenv("APPDATA"), r"Microsoft\Windows\Start Menu\Programs"
    )
    shortcut_name = "xlfly.lnk"
    shortcut_path = os.path.join(start_menu_folder, shortcut_name)

    # Create the Shell object
    _WSHELL = win32com.client.Dispatch("Wscript.Shell")
    wscript = _WSHELL.CreateShortCut(shortcut_path)
    wscript.TargetPath = pythonw_path
    wscript.Arguments = "-m xlfly"
    wscript.WorkingDirectory = os.path.dirname(pythonw_path)
    wscript.WindowStyle = 0
    wscript.Description = "Control Excel"
    wscript.IconLocation = icon_path
    wscript.save()


if __name__ == "__main__":
    create_shortcut()
