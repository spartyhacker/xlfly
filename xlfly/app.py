import tkinter as tk
from tkinter import ttk, messagebox, font
import xlwings as xw
import sys
import pandas as pd
import os

# import copyover to enable xw.Range.to_link()
import xlfly.copyover

CONFIG_PAGE_NAME = "xlfly"


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
    if isinstance(selected.value, str):
        df_construct = [selected.value]
    else:
        df_construct = selected.value
    cmd = pd.DataFrame(df_construct).stack().reset_index(drop=True)
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

    # support multiple cell selection
    for c in cmd:
        exec(c, locals(), globals())


def main():
    # UI
    root = tk.Tk()
    root.title("xlfly")
    # root.geometry("200x100")
    root.attributes("-topmost", 1)
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    root.iconbitmap(icon_path)

    # put widgets
    larger_font = font.Font(family="Helvetica", size=16, weight="bold")

    # Create a style and configure the custom style with the larger font
    style = ttk.Style()
    style.configure("Larger.TButton", font=larger_font)

    btn_run_selected = ttk.Button(
        root,
        text="Run Python",
        command=lambda: exec_func(run_cell),
        style="Larger.TButton",
    )
    btn_run_selected.pack(pady=10, padx=50)

    # btn_create_config = ttk.Button(
    #     root, text="Add Config Sheet", command=lambda: exec_func(create_config)
    # )
    # btn_create_config.pack(side=tk.LEFT)

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
