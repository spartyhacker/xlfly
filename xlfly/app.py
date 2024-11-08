import tkinter as tk
import inspect
from tkinter import filedialog
import importlib.util
import sys
import traceback
from PIL import Image, ImageTk
from tkinter import ttk, messagebox, font
import shutil
import os
import subprocess
import xlfly
import importlib.metadata

# to debug, you need this
import xlwings as xw
import pandas as pd
from xlfly.check_package import check_requirements
import xlfly.configs as configs
from xlfly.pop_template import TempWindow
from xlfly.about import AboutWin

CONFIG_PAGE_NAME = "xlfly"


# functions
def redirect_output():
    "without this, pyminitab would throw stderr is None error"
    log_directory = os.path.join(os.path.expanduser("~"), "xlfly_logs")
    os.makedirs(log_directory, exist_ok=True)


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
        return None
        # raise ValueError("Config page does not exist")
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
    # use python -m pip to conform to virtual environment
    command = [sys.executable, "-m", "pip", "install"] + pkgs
    subprocess.check_call(command)


def create_debug_file():
    wb_path = os.path.dirname(xw.books.active.fullname)
    dst_file = os.path.join(wb_path, "debug.py")
    src_file = os.path.join(os.path.dirname(__file__), "debug.py")
    shutil.copy2(src_file, dst_file)

    print("debug.py created")


def run_pic(pic: xw.Picture):
    """Event for Selecting Picture to Run

    This function is to be replaced with other defined functions
    """

    print(f"Selection is picture: {pic.name} in sheet {pic.parent.name}")


class XlflyApp:
    def __init__(self, root):
        # redirect output before starting the program
        # redirect_output()

        self.run_pic = run_pic
        self.root = root
        self.console_shown = False
        self.console_thread = None

        # load in settings
        self.settings = configs.load_settings()
        self.selected_temp = None

        # UI
        self.root.title("xlfly")
        self.root.attributes("-topmost", 1)
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        self.root.iconbitmap(icon_path)
        self.console_visible = False

        # put button
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
            command=lambda: self.exec_func(self.run_selected),
            compound=tk.LEFT,
            style="Larger.TButton",
        )
        self.btn_run_selected.pack(pady=5, padx=50)
        self.add_menu()

    def add_menu(self):
        # Create the menu bar
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # file_menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)

        self.file_menu.add_command(
            label="Add Config Sheet", command=lambda: self.exec_func(create_config)
        )

        self.file_menu.add_command(
            label="Update xlfly", command=lambda: self.exec_func(self.update_xlfly)
        )

        self.file_menu.add_separator()
        self.file_menu.add_command(
            label="About", command=lambda: self.exec_func(self.about)
        )
        self.file_menu.add_command(label="Exit", command=self.root.quit)

        # Debug menu
        self.debug_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Debug", menu=self.debug_menu)
        self.debug_menu.add_command(
            label="Create Debug Script",
            command=lambda: self.exec_func(create_debug_file),
        )

        self.debug_menu.add_command(
            label="Restart", command=lambda: self.exec_func(self.restart_app)
        )

        # Templates menu
        self.template_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Template", menu=self.template_menu)

        self.template_menu.add_command(
            label="Choose Template",
            command=lambda: self.exec_func(self.choose_temp),
        )

        self.template_menu.add_command(
            label="Set Template Folder",
            command=lambda: self.exec_func(self.set_tempfolder),
        )

    # show about window
    def about(self):
        abwin = AboutWin(self.root)
        abwin.show()

    # templates related functions
    def set_tempfolder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            print(f"Selected folder: {folder_path}")
            self.settings["tempfolder"] = folder_path
            configs.save_settings(self.settings)

    def choose_temp(self):
        "choose a template to open"

        # pop window to select the template to use
        popup = TempWindow(self.root)
        if popup.selected_temp is None:
            return
        temp_initpath = os.path.join(
            self.settings["tempfolder"], popup.selected_temp, "__init__.py"
        )
        if os.path.exists(os.path.join(temp_initpath)):
            spec = importlib.util.spec_from_file_location("__init__", temp_initpath)
            init = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(init)

            # if one parameter, only allow app object to be passed for manipulation
            signature = inspect.signature(init.main)
            num_params = len(signature.parameters)
            if num_params == 0:
                init.main()
            elif num_params == 1:
                init.main(self)
        else:
            raise FileNotFoundError(
                f"Template {popup.selected_temp} does not have a __init__.py file"
            )

    def exec_func(self, func):
        try:
            func()
        except Exception as e:
            # Print the detailed traceback
            tb_str = traceback.format_exc()

            messagebox.showerror("Error", repr(e) + f"\n\n{tb_str}")

    def restart_app(self):
        "restart the current program"
        self.root.destroy()
        python_exe = sys.executable
        # had to use subproess, otherwise it will fail due to c:\program files path
        # having blanks
        subprocess.run([python_exe] + sys.argv)

    def update_xlfly(self):
        "update xlfly app"

        command = [sys.executable, "-m", "pip", "install", "xlfly", "-U"]
        subprocess.check_call(command)

        # load default template to create functionality
        # get the default folder path
        from xlfly.scripts import init_default

        settings = configs.load_settings()
        if "tempfolder" not in settings.keys():
            return

        tempfolder = settings["tempfolder"]
        default_init_file = os.path.join(tempfolder, "default", "__init__.py")
        init_default(default_init_file)

        # restart at the end
        self.restart_app()

    def appendpath(self, path_str: str):
        if not isinstance(path_str, str):
            return

        if os.path.exists(path_str):
            sys.path.append(path_str)

    def cmd_condition(self):
        # get configs
        df_config = get_configs()

        # append path
        curr_wb_path = xw.books.active.fullname

        if os.path.exists(curr_wb_path):
            sys.path.append(os.path.dirname(curr_wb_path))

        if df_config is not None:
            self.appendpath(df_config.loc["script_path"].value)

        # append default folder, for case when there is no xlfly page to set up
        defaultfolder = os.path.join(self.settings["tempfolder"], "default")
        self.appendpath(defaultfolder)

        # check packages
        if df_config is not None:
            rqm = df_config.loc["requirements"].value

            # Need to find the default folder's requirements.txt to install
            # missing dependencies
            default_rqm_file = os.path.join(
                self.settings["tempfolder"], r"default/requirements.txt"
            )
            if os.path.exists(default_rqm_file):
                with open(default_rqm_file, "r") as file:
                    # Read the contents of the file into a string
                    default_rqm = file.read()
                    default_rqm = default_rqm.replace("\n", " ")
                    if rqm is not None:
                        rqm = " ".join([rqm, default_rqm])
                    else:
                        rqm = default_rqm
        else:
            rqm = None

        pkgs = check_requirements(rqm)
        if pkgs:
            run_install(pkgs)
            self.restart_app()
        else:
            pass

        # execute pre-command
        if df_config is not None:
            pre_cmd = df_config.loc["pre_cmd"].value
        else:
            pre_cmd = None

        # define variables
        wb = xw.books.active
        local_var = {}
        if CONFIG_PAGE_NAME in [s.name for s in wb.sheets]:
            # raise ValueError("Config page does not exist")
            config_sht = xw.books.active.sheets[CONFIG_PAGE_NAME]
            df_var: pd.DataFrame = (
                config_sht.tables["var"].range.options(pd.DataFrame, index=False).value
            )

            if len(df_var.dropna()) != 0:
                for id, r in df_var.iterrows():
                    local_var[r.Name] = r.Value

        return pre_cmd, local_var

    def run_cell(self, selected: xw.Range):
        "when clicking button, run this"

        pre_cmd, local_var = self.cmd_condition()

        for key, val in local_var.items():
            locals()[key] = val

        # to override the run_pic() function, use this:
        # self.run_pic =

        if pre_cmd:
            exec(pre_cmd, locals(), globals())

        # from default folder's default.py import all functions
        import default

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

        # this is for logging purpose
        for c in cmd:
            exec(c, locals(), globals())

            # if default module has run_callback() function, run it
            if hasattr(default, "run_callback"):
                default.run_callback(c)

    def run_pic_full_process(self, pic):
        pre_cmd, local_var = self.cmd_condition()

        for key, val in local_var.items():
            locals()[key] = val

        if pre_cmd:
            exec(pre_cmd, locals(), globals())

        # from default folder's default.py import all functions
        import default

        self.run_pic(pic)

    def run_selected(self):
        "callback when clicking button"

        app = xw.apps.active
        selected = app.selection
        selection_api = app.api.Selection
        if isinstance(selected, xw.Range):
            self.run_cell(selected)
        elif selection_api.ShapeRange.Type == 13:  # if it is a picture
            pic_name = selection_api.ShapeRange.Name
            print(f"pic name: {pic_name}")
            pic = xw.books.active.sheets.active.pictures[pic_name]
            self.run_pic_full_process(pic)
        else:
            print("not a range nor a picture")


def _run_main():
    root = tk.Tk()
    app = XlflyApp(root)
    root.mainloop()


if __name__ == "__main__":
    _run_main()
