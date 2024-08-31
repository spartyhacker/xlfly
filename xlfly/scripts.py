import xlfly.create_shortcut
import os
import xlfly.configs as configs
import argparse
import importlib


def init_default(default_init_file: str):
    print(f"init file:{default_init_file}")
    spec = importlib.util.spec_from_file_location("__init__", default_init_file)
    init = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(init)
    init.setup()


def main():
    parser = argparse.ArgumentParser(description="Run a Tkinter application.")
    parser.add_argument("--init", action="store_true", help="Create Start Menu")
    parser.add_argument("-t", "--tempfolder", type=str, help="Set Template Root Folder")
    args = parser.parse_args()

    if args.init:
        xlfly.create_shortcut.create_shortcut()
        print("Created Start Menu Item")

    if args.tempfolder:
        print(f"add {args.tempfolder} as template path")

        settings = configs.load_settings()
        folder_path = os.path.normpath(args.tempfolder)
        if folder_path:
            print(f"Selected folder: {folder_path}")
            settings["tempfolder"] = folder_path
            configs.save_settings(settings)
            print("Initilizing default init...")
            default_init_file = os.path.join(folder_path, "default", "__init__.py")
            init_default(default_init_file)
            print("Done")


if __name__ == "__main__":
    main()
