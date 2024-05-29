import tkinter as tk
from tkinter import ttk, messagebox
import xlwings as xw
import sys
import pandas as pd
import os

# import local module
from xlfly.copyover import *

CONFIG_PAGE_NAME = "xlfly"


# functions
def hello():
    # use the following syntax to catch error message for users
    try:
        rng = xw.books.active.sheets.active.range("A1")
        rng.value = "hello world"
    except Exception as e:
        messagebox.showerror("Error", repr(e))


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


def run_cell():
    app = xw.apps.active
    selected = app.selection
    cmd = selected.value
    df = get_configs()
    sys.path.append(df.loc["script_path"].value)
    pre_cmd = df.loc["pre_cmd"].value
    if pre_cmd:
        exec(pre_cmd)

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

    exec(cmd, locals(), globals())


def main():
    # UI
    root = tk.Tk()
    root.title("Excel智慧云V3")
    root.geometry("300x70")
    root.attributes("-topmost", 1)
    root.iconbitmap("icon.ico")

    # put widgets
    btn_run = ttk.Button(root, text="️Hello", command=hello)
    btn_run.pack(side=tk.LEFT)

    btn_create_config = ttk.Button(
        root, text="生成配置页", command=lambda: exec_func(create_config)
    )
    btn_create_config.pack(side=tk.LEFT)

    btn_run_selected = ttk.Button(
        root, text="运行单元格Python", command=lambda: exec_func(run_cell)
    )
    btn_run_selected.pack(side=tk.LEFT)

    root.mainloop()


if __name__ == "__main__":
    main()
