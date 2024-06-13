# xlfly - separate GUI based Excel tools

I love xlwings! However, I have noticed it very difficult to handle heavy calculations if using user defined functions. Hence xlfly: making using UDF easier. I chose separate GUI so that I can run parallel threads easier. 

## Usage

![usage demo](doc/usage_demo.png)

This package only works on Windows

1. Write python commands in Excel cells. Select it, click to run!
2. Write python in **commments** section in Excel. Select it, and click to run! Note: if there is comment, the cell value commands will NOT run.

Note: if there is numeric value and a comment on the cell, the cell will not **run** the numeric value. For example, say "A1" cell has value 1, and comment "sht["A1"].value = 1", you run it. It will not run the "1" in the cell

Create Windows Start Menu Item:

```bash
>>> xly --init
```

Excel selection to "Run Python"
1. Write python expressions in the cells
2. Select the cell with python scripts
3. Click "Run Python" button


I used the icon from https://www.iconfinder.com/search?q=wings&price=free drawn by Monsieur Steven Ankri. Thanks!

Run multiple cells with a command

`run_cell(xw.books.active.sheets.active["A1:A5"])`

The function `run_cell()` is always usage to put in any cell to run other batch cells

## PythonPath

When run python, both `current workbook` and `script_path` setting from setting page will be added to pythonpath. You can put your draft python script file in the same folder as current Excel file to debug.

## Debug

To debug the scripts in Excel, run the menu: tools > create debug script. There will be a debug.py file to the same folder as current Excel file. You can start from there to debug

script of the same name from Excel file folder will be imported first. 

- Thus you can debug by putting the scripts in the Excel folder
- Once done, move the file to the python path specified in the config page


Error messages will be thrown out:

![error message](doc/error_msg.png)

You can also toggle to show the console at Menu: Tools > Toggle Console