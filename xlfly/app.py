import tkinter as tk
from tkinter import ttk, messagebox, font
from tkinter.scrolledtext import ScrolledText
import sys
import os
import threading
import subprocess

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
        messagebox.showerror("Error", repr(e))


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


def run_cell():
    import xlwings as xw

    # get the command
    app = xw.apps.active
    selected = app.selection
    if isinstance(selected.value, str):
        df_construct = [selected.value]
    else:
        df_construct = selected.value
    cmd = pd.DataFrame(df_construct).stack().reset_index(drop=True)

    # comments commands
    comments = {}
    for cell in selected:
        comment = cell.api.Comment
        if comment is not None:
            comments[cell.address] = comment.Text()

    comments_cmd = list(comments.values())
    cmd = pd.concat([cmd, pd.Series(comments)])

    # get configs
    df = get_configs()

    # append path
    sys.path.append(df.loc["script_path"].value)
    curr_wb_path = xw.books.active.fullname
    if os.path.exists(curr_wb_path):
        sys.path.append(os.path.dirname(curr_wb_path))

    # execute pre-command
    pre_cmd = df.loc["pre_cmd"].value
    if pre_cmd:
        exec(pre_cmd)

    # check packages
    rqm = df.loc["requirements"].value
    pkgs = check_requirements(rqm)
    if pkgs:
        print("\nInstalling missing packages...")
        run_install(pkgs)
        restart_app()
    else:
        print("\nAll packages are already installed and meet the version requirements.")

    # define variables
    wb = app.books.active
    if CONFIG_PAGE_NAME not in [s.name for s in wb.sheets]:
        raise ValueError("Config page does not exist")

    config_sht = app.books.active.sheets[CONFIG_PAGE_NAME]
    df_var: pd.DataFrame = (
        config_sht.tables["var"].range.options(pd.DataFrame, index=False).value
    )
    if len(df_var.dropna()) != 0:
        for id, r in df_var.iterrows():
            # if a string, assign with quote
            if isinstance(r.Value, str):
                locals()[r.Name] = r.Value
            else:
                exec(f"{r.Name} = {r.Value}")

    # support multiple cell selection
    for c in cmd:
        exec(c, locals(), globals())


def main():

    # using the root instance from outside this function

    # UI
    root.title("xlfly")
    # root.geometry("200x100")
    root.attributes("-topmost", 1)
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    root.iconbitmap(icon_path)

    # put widgets
    larger_font = font.Font(family="Helvetica", size=12)

    # Create a style and configure the custom style with the larger font
    style = ttk.Style()
    style.configure("Larger.TButton", font=larger_font)

    btn_run_selected = ttk.Button(
        root,
        text="Run Python",
        command=lambda: exec_func(run_cell),
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
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

    # run mainloop
    root.mainloop()


if __name__ == "__main__":
    main()
