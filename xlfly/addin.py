"""
put addin xlam file to the directory
"""

import os
import xlwings as xw
import shutil

# Get the path to the AddIns directory

apps = xw.apps

if len(apps) == 0:
    app = xw.App(visible=True)


def move_addin(xlam_file: str):
    "move xlam file to addin directory"

    destination_dir = os.path.join(os.getenv("APPDATA"), "Microsoft", "AddIns")

    # Copy the file
    try:
        shutil.copy(xlam_file, destination_dir)
        print(f"Copied {xlam_file} to {destination_dir}")
    except FileNotFoundError:
        print(f"The file {xlam_file} does not exist.")
    except PermissionError:
        print(f"Permission denied to copy the file to {destination_dir}.")
    except Exception as e:
        print(f"An error occurred: {e}")


def rm_addin(addin_name: str):
    "remove addin from Excel Addin list, addin_name is usually *.xlam file name"

    app = xw.apps.active
    addins = app.api.AddIns

    for addin in addins:
        if addin.Name == addin_name:
            if addin.Installed:
                addin.Installed = False
                print(f"addin {addin_name} is removed")
            else:
                print(f"addin {addin_name} is already deactivated")

            return

    print(f"addin {addin_name} is not found")


def install_addin(xlam_file: str):
    app = xw.apps.active
    addin = app.api.AddIns.Add(xlam_file, False)
    addin.Installed = True
