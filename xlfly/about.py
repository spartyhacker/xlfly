import tkinter as tk
import importlib.metadata
import os
from tkinter import ttk
from PIL import Image, ImageTk


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Main Window")

        self.popup_button = ttk.Button(
            self, text="Show Pop-up", command=self.show_popup
        )
        self.popup_button.pack(pady=20)

    def show_popup(self):
        popup = AboutWin(self)
        popup.show()


class AboutWin(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Pop-up Window")
        self.geometry("300x200")  # Set the size of the pop-up window

        self.create_widgets()

    def create_widgets(self):
        # Load and display an image
        currdir = os.path.dirname(__file__)
        image_path = os.path.join(currdir, "about.png")  # Replace with your image path
        image = Image.open(image_path)
        photo = ImageTk.PhotoImage(image)
        image_label = tk.Label(self, image=photo)
        image_label.image = photo
        image_label.pack()

        # Add text
        version = importlib.metadata.version("xlfly")
        texts_to_show = f"xlfly - {version} \n Time to upgrade your tools"
        text_label = tk.Label(self, text=texts_to_show)
        text_label.pack(pady=10)

        # Button to close the pop-up
        close_button = ttk.Button(self, text="Close", command=self.destroy)
        close_button.pack(pady=10)

    def show(self):
        self.grab_set()  # Make the pop-up window modal
        self.wait_window()  # Wait for the pop-up window to be closed


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
