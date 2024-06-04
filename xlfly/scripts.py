import xlfly.create_shortcut
import argparse


def main():
    parser = argparse.ArgumentParser(description="Run a Tkinter application.")
    parser.add_argument("--init", action="store_true", help="Enable debug mode")
    args = parser.parse_args()

    if args.init:
        xlfly.create_shortcut.create_shortcut()


if __name__ == "__main__":
    main()
