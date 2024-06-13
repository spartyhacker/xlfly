import tkinter as tk
import traceback
import xlwings as xw
from PIL import Image, ImageTk
from tkinter import ttk, messagebox, font, PhotoImage
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
root = tk.Tk()


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


def restart_app():
    root.destroy()
    python = sys.executable
    os.execl(python, python, *sys.argv)


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


def cmd_condition():

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
        restart_app()
    else:
        print("\nAll packages are already installed and meet the version requirements.")

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


def run_cell(selected: xw.Range):

    pre_cmd, local_var = cmd_condition()

    for key, val in local_var.items():
        locals()[key] = val

    exec(pre_cmd)

    # run the commands
    cmds = {}
    for cell in selected:
        comment = cell.api.Comment
        val = cell.value
        if comment is not None:
            cmds[cell.address] = comment.Text()
        elif (val is not None) and (not isinstance(val, float)):
            cmds[cell.address] = val

    cmd = list(cmds.values())

    for c in cmd:
        exec(c, locals(), globals())


def run_selected():

    app = xw.apps.active
    selected = app.selection
    run_cell(selected)


def create_debug_file():
    wb_path = os.path.dirname(xw.books.active.fullname)
    file_name = os.path.join(wb_path, "debug.py")
    text_to_write = """
import xlfly.app as app
import xlwings as xw
import pandas as pd

pre_cmd, local_var = app.cmd_condition()

for key, val in local_var.items():
    locals()[key] = val

exec(pre_cmd)

# put your debug command here, like
# sht["A1"].value = 1
    """

    with open(file_name, "w") as file:
        # Write the text to the file
        file.write(text_to_write)

    print("debug.py created")


def update_xlfly():
    command = [sys.executable, "-m", "pip", "install", "xlfly", "-U"]
    subprocess.check_call(command)
    restart_app()


def _run_main():

    # using the root instance from outside this function

    # UI
    version = importlib.metadata.version("xlfly")
    root.title(f"xlfly-{version}")
    # root.geometry("200x100")
    root.attributes("-topmost", 1)
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    root.iconbitmap(icon_path)

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
    style.configure("Larger.TButton", image=icon)

    btn_run_selected = ttk.Button(
        root,
        text="Run Python",
        command=lambda: exec_func(run_selected),
        compound=tk.LEFT,
        style="Larger.TButton",
    )
    btn_run_selected.pack(pady=5, padx=50)

    # Create the menu bar
    menu_bar = tk.Menu(root)
    root.config(menu=menu_bar)

    file_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Tools", menu=file_menu)

    file_menu.add_command(
        label="Add Config Sheet", command=lambda: exec_func(create_config)
    )

    file_menu.add_command(
        label="Create Debug Script", command=lambda: exec_func(create_debug_file)
    )

    file_menu.add_command(label="Update xlfly", command=lambda: exec_func(update_xlfly))

    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

    # run mainloop
    root.mainloop()


if __name__ == "__main__":
    _run_main()
