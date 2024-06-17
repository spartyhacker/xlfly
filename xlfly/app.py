import tkinter as tk
import traceback
import xlwings as xw
from PIL import Image, ImageTk
from tkinter import ttk, messagebox, font, PhotoImage
import shutil
from tkinter.scrolledtext import ScrolledText
import sys
import os
import threading
import subprocess
import xlfly
import importlib.metadata

# to debug, you need this
import xlwings as xw
import pandas as pd
import xlfly.copyover
from xlfly.check_package import check_requirements, install_packages

CONFIG_PAGE_NAME = "xlfly"


# functions
def exec_func(func):
    try:
        func()
    except Exception as e:
        # Print the detailed traceback
        tb_str = traceback.format_exc()
        # print(f"Detailed traceback:\n{tb_str}")

        messagebox.showerror("Error", repr(e) + f"\n\n{tb_str}")


def create_config():
    wb = xw.books.active
    if CONFIG_PAGE_NAME in [s.name for s in wb.sheets]:
        messagebox.showerror("Error", "config page already exists")
        return None
    else:
        sht = wb.sheets.add(CONFIG_PAGE_NAME)
        sht.range("A1").value = "script_path"
        sht.range("A2").value = "pre_cmd"
        sht.range("B2").value = "sht = xw.books.active.sheets.active"
        sht.range("A3").value = "requirements"
        sht.range("A3").api.AddComment(
            "separate by blank, using requirements.txt syntax"
        )
        sht.range("A7").value = "Variable Definition"

        df_var = pd.DataFrame(columns=["Name", "Value"])
        sht.tables.add(sht["A8"], name="var").update(df_var, index=False)


def get_configs():
    wb = xw.books.active

    if CONFIG_PAGE_NAME not in [s.name for s in wb.sheets]:
        raise ValueError("Config page does not exist")
    else:
        df = (
            wb.sheets[CONFIG_PAGE_NAME]["A1:B1"]
            .options(pd.DataFrame, expand="vertical", index=False, header=False)
            .value
        )
        df.columns = ["param", "value"]
        df = df.set_index("param")
        return df


def run_install(pkgs):
    messagebox.showwarning(
        "Warning Message", "Missing depended python packages, will install now"
    )
    sub_install_packages(pkgs)


def sub_install_packages(pkgs):
    # open_terminal = ["cmd", "/c", "start", "cmd", "/k"]
    # command = open_terminal + [sys.executable, "-m", "pip", "install"] + pkgs
    command = [sys.executable, "-m", "pip", "install"] + pkgs
    subprocess.check_call(command)


def create_debug_file():
    wb_path = os.path.dirname(xw.books.active.fullname)
    dst_file = os.path.join(wb_path, "debug.py")
    src_file = os.path.join(os.path.dirname(__file__), "debug.py")
    shutil.copy2(src_file, dst_file)

    print("debug.py created")


# console output
class ConsoleOutput:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.tag_config("stdout", foreground="black")
        self.text_widget.tag_config("stderr", foreground="red")

    def write(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message, ("stdout",))
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

    def flush(self):
        pass


class ErrorOutput(ConsoleOutput):
    def write(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message, ("stderr",))
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)


class XlflyApp:
    def __init__(self, root):
        self.root = root
        # UI
        version = importlib.metadata.version("xlfly")
        self.root.title(f"xlfly-{version}")
        self.root.attributes("-topmost", 1)
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        self.root.iconbitmap(icon_path)
        self.console_visible = False

        # put widgets
        larger_font = font.Font(family="Helvetica", size=12)

        # Create a style and configure the custom style with the larger font
        style = ttk.Style()
        style.configure("Larger.TButton", font=larger_font)

        # add button image
        icon_path = os.path.join(os.path.dirname(__file__), "python.png")
        button_ht = 20
        icon_image = Image.open(icon_path).resize((button_ht, button_ht))
        icon = ImageTk.PhotoImage(icon_image)

        self.btn_run_selected = ttk.Button(
            self.root,
            text="Run Python",
            image=icon,
            command=lambda: exec_func(self.run_selected),
            compound=tk.LEFT,
            style="Larger.TButton",
        )
        self.btn_run_selected.pack(pady=5, padx=50)

        # Create a ScrolledText widget
        self.console_text = ScrolledText(self.root, wrap=tk.WORD, width=80, height=20)
        # self.console_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Redirect stdout and stderr
        sys.stdout = ConsoleOutput(self.console_text)
        sys.stderr = ErrorOutput(self.console_text)

        # Create the menu bar
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Tools", menu=self.file_menu)

        self.file_menu.add_command(
            label="Add Config Sheet", command=lambda: exec_func(create_config)
        )

        self.file_menu.add_command(
            label="Create Debug Script", command=lambda: exec_func(create_debug_file)
        )

        self.file_menu.add_command(
            label="Toggle Console", command=lambda: exec_func(self.toggle_console)
        )

        self.file_menu.add_command(
            label="Update xlfly", command=lambda: exec_func(self.update_xlfly)
        )

        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=root.quit)

    def toggle_console(self):
        if self.console_visible:
            self.console_text.pack_forget()
        else:
            self.console_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.console_visible = not self.console_visible

    def show_console(self):
        self.console_visible = True
        self.console_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    def hide_console(self):
        self.console_visible = False
        self.console_text.pack_forget()

    def restart_app(self):
        self.root.destroy()
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def update_xlfly(self):
        command = [sys.executable, "-m", "pip", "install", "xlfly", "-U"]
        subprocess.check_call(command)
        self.restart_app()

    def cmd_condition(self):

        # get configs
        df = get_configs()

        # append path
        curr_wb_path = xw.books.active.fullname
        if os.path.exists(curr_wb_path):
            sys.path.append(os.path.dirname(curr_wb_path))
        sys.path.append(df.loc["script_path"].value)

        # check packages
        rqm = df.loc["requirements"].value
        pkgs = check_requirements(rqm)
        if pkgs:
            print("\nInstalling missing packages...")
            run_install(pkgs)
            self.restart_app()
        else:
            pass

        # execute pre-command
        pre_cmd = df.loc["pre_cmd"].value

        # define variables
        wb = xw.books.active
        if CONFIG_PAGE_NAME not in [s.name for s in wb.sheets]:
            raise ValueError("Config page does not exist")

        config_sht = xw.books.active.sheets[CONFIG_PAGE_NAME]
        df_var: pd.DataFrame = (
            config_sht.tables["var"].range.options(pd.DataFrame, index=False).value
        )

        local_var = {}
        if len(df_var.dropna()) != 0:
            for id, r in df_var.iterrows():
                local_var[r.Name] = r.Value

        return pre_cmd, local_var

    def run_cell(self, selected: xw.Range):

        pre_cmd, local_var = self.cmd_condition()

        for key, val in local_var.items():
            locals()[key] = val

        exec(pre_cmd)

        # run the commands
        # support for non-continous range selection
        cmds = {}
        act_sht = xw.books.active.sheets.active
        rng_parts = selected.address.split(",")
        rng_lst = [act_sht.range(p) for p in rng_parts]

        for rng in rng_lst:
            for cell in rng:
                comment = cell.api.Comment
                val = cell.value
                if comment is not None:
                    cmds[cell.address] = comment.Text()
                elif (val is not None) and (not isinstance(val, float)):
                    cmds[cell.address] = val

        cmd = list(cmds.values())

        for c in cmd:
            exec(c, locals(), globals())

    def run_selected(self):

        app = xw.apps.active
        selected = app.selection
        self.run_cell(selected)


def _run_main():

    root = tk.Tk()
    app = XlflyApp(root)
    root.mainloop()


if __name__ == "__main__":
    _run_main()
