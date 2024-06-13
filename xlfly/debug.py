import tkinter as tk
import xlfly.app as xlapp
import xlwings as xw
import pandas as pd


root = tk.Tk()
app = xlapp.XlflyApp(root)

pre_cmd, local_var = app.cmd_condition()

for key, val in local_var.items():
    locals()[key] = val

exec(pre_cmd)

# put your debug command here, like
sht["A1"].value = 1
