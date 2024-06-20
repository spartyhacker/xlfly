import xlfly.create_shortcut
import xlfly.configs as configs
import argparse


def main():
    parser = argparse.ArgumentParser(description="Run a Tkinter application.")
    parser.add_argument("--init", action="store_true", help="Create Start Menu")
    parser.add_argument("-t", "--tempfolder", type=str, help="Set Template Root Folder")
    args = parser.parse_args()

    if args.init:
        xlfly.create_shortcut.create_shortcut()

    if args.tempfolder:

        print(f"add {args.tempfolder} as template path")

        settings = configs.load_settings()
        folder_path = args.tempfolder
        if folder_path:
            print(f"Selected folder: {folder_path}")
            settings["tempfolder"] = folder_path
            configs.save_settings(settings)


if __name__ == "__main__":
    main()
