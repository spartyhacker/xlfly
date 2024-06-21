import os
import json

CONFIG_FILE = "xlfly.json"
# Define the path to the settings file in the user's personal folder
settings_file_path = os.path.join(os.path.expanduser("~"), CONFIG_FILE)


# Function to save settings to the file
def save_settings(settings):
    with open(settings_file_path, "w") as settings_file:
        json.dump(settings, settings_file, indent=4)


# Function to load settings from the file
def load_settings():
    if os.path.exists(settings_file_path):
        with open(settings_file_path, "r") as settings_file:
            return json.load(settings_file)
    return {}


if __name__ == "__main__":
    # Load settings at the start
    settings = load_settings()
    last_selected_folder = settings.get("last_selected_folder", "No folder selected")
