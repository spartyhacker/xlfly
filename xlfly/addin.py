"""
put addin xlam file to the directory
"""

import os
import shutil

# Get the path to the AddIns directory


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
