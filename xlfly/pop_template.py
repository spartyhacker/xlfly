import tkinter as tk
from tkinter import ttk
import os
from xlfly.configs import load_settings


class TempWindow:
    def __init__(self, parent):
        self.parent = parent
        self.selected_temp = None
        settings = load_settings()
        tempfolder = settings["tempfolder"]

        subfolders = [
            name
            for name in os.listdir(tempfolder)
            if os.path.isdir(os.path.join(tempfolder, name))
        ]
        self.items = subfolders

        # Create the popup window
        self.popup_window = tk.Toplevel(parent)
        self.popup_window.title("Select Item")

        # Frame for filter entry, listbox, and buttons in popup
        self.frame = ttk.Frame(self.popup_window)
        self.frame.pack(padx=20, pady=20)

        # Filter Entry
        self.filter_entry = ttk.Entry(self.frame, width=30)
        self.filter_entry.pack(pady=10)
        self.filter_entry.bind("<KeyRelease>", self.update_listbox)

        # Listbox
        self.listbox = tk.Listbox(self.frame, selectmode=tk.SINGLE, width=30, height=5)
        self.listbox.pack(pady=10)

        # Populate listbox with initial items
        for item in self.items:
            self.listbox.insert(tk.END, item)

        # OK button
        self.ok_button = ttk.Button(self.frame, text="OK", command=self.on_ok)
        self.ok_button.pack(side=tk.LEFT, padx=10)

        # Cancel button
        self.cancel_button = ttk.Button(
            self.frame, text="Cancel", command=self.on_cancel
        )
        self.cancel_button.pack(side=tk.LEFT, padx=10)

        # Make the popup window modal
        self.popup_window.transient(parent)
        self.popup_window.grab_set()
        self.popup_window.wait_window()

    def update_listbox(self, event=None):
        filter_text = self.filter_entry.get().strip().lower()
        filtered_items = [item for item in self.items if filter_text in item.lower()]
        self.listbox.delete(0, tk.END)  # Clear previous items
        for item in filtered_items:
            self.listbox.insert(tk.END, item)

    def on_ok(self):
        selected_index = self.listbox.curselection()
        if selected_index:
            selected_item = self.listbox.get(selected_index[0])
        self.popup_window.destroy()
        self.selected_temp = selected_item

    def on_cancel(self):
        print("Selection canceled")
        self.selected_temp = None
        self.popup_window.destroy()


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Main Window")

        # Button to open popup window
        self.popup_button = ttk.Button(
            self.root, text="Open Popup", command=self.open_popup
        )
        self.popup_button.pack(padx=20, pady=20)

    def open_popup(self):
        popup = TempWindow(self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
