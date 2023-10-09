"""
This Graphical User Interface allows to calculate ratio of values from a packing list.

Example:
     {"Thickness":[0.12, 0.12, 0.13, 0.14, 0.14, 0.14],
      "Net Weight": [2.500, 3.500, 4.000, 3.250, 2.250, 2.000]}

Ratio Calculated:
0.12 = 6.000 (2 Occurrence of 0.12)
0.13 = 4.000 (1 Occurrence of 0.13)
0.14 = 7.500 (3 Occurrence of 0.14)
"""

# Front-End Imports:
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

# Back-End Imports:
import pandas as pd
import numpy as np
import math as m
import pyperclip
import openpyxl
import xlrd

# Setting up the root:
root = Tk()
root.title("Ratio Maker")
root.geometry("1920x1080")
root.state("zoomed")
root.config(bg="#FFFFFF")

# Create the canvas and the frame
canvas = Canvas(root)
frame = Frame(canvas)

# Create the vertical and horizontal scrollbars
vscroll = Scrollbar(root, orient="vertical", command=canvas.yview)
hscroll = Scrollbar(root, orient="horizontal", command=canvas.xview)

# Set the scrollbars to the canvas
canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

# Pack the scrollbars and the canvas
vscroll.pack(side="right", fill="y")
hscroll.pack(side="bottom", fill="x")
canvas.pack(side="left", fill="both", expand=True)

# Create a window to hold the frame and add it to the canvas
window = canvas.create_window((0, 0), window=frame, anchor="nw")

# Bind the frame to the canvas's scroll event
frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

canvas.config(bg="#FFFFFF")
frame.config(bg="#FFFFFF")


# ====== #
# Functions Corner:
# ====== #
def open_packing_list(prompt):
    global excel_path

    # To select our packing list file:
    excel_path = filedialog.askopenfilename(title="Select a file",
                                            filetypes=[("Microsoft Excel Spreadsheet", "*.xlsx"),
                                                       ("Microsoft Excel Spreadsheet", "*.xls")])

    if excel_path:

        # Using pd.ExcelFile to read sheet names:
        excel_file = pd.ExcelFile(excel_path)
        excel_file_sheets = excel_file.sheet_names

        display_excel_file_scrolled_text.delete(1.0, END)
        display_excel_file_scrolled_text.insert("insert", prompt + "\n\n")
        display_excel_file_scrolled_text.insert("insert", F"File Name: {excel_path}\n\n")
        display_excel_file_scrolled_text.insert("insert", F"Total Sheet: {len(excel_file_sheets)}\n")

        for index, sheet_name in enumerate(excel_file_sheets):
            display_excel_file_scrolled_text.insert("insert", F"{index + 1}). {sheet_name}\n")

        display_excel_file_scrolled_text.insert("insert", F"\nFirst 5 rows & 3 columns for: {excel_file_sheets[0]}\n\n")
        display_excel_file_scrolled_text.insert("insert",
                                                str(pd.read_excel(excel_file,
                                                                  sheet_name=excel_file_sheets[0]).iloc[:, :3].head()))

    else:
        display_excel_file_scrolled_text.delete(1.0, END)
        display_excel_file_scrolled_text.insert("insert", "No file was selected")


def get_col_names(excel_file, sheet_name, row_number):
    global pl
    global columns
    global option_menu_and_variables

    if excel_file:
        # To read Excel file using Pandas.
        pl = pd.read_excel(excel_file, sheet_name=sheet_name, header=(int(row_number) - 1))

        # Dropping any columns with in-complete information, done to remove total calculated in the end.
        pl.dropna(axis="index", how="all", inplace=True)

        # Insert names of columns into OptionMenu:
        columns = list(pl.columns.values)
        columns.insert(0, "None")
        columns.insert(1, "No, Single Item")

        for widgets_variables in option_menu_and_variables:
            for index, variable in enumerate(widgets_variables):
                variable.set_menu(*columns) if index == 0 else variable.set(columns[2])

        # delete anything that is currently present in ScrolledTextWidget (extracted_columns)
        display_extracted_columns.config(display_extracted_columns.delete(1.0, END))

        # inserting columns names on each line:
        for index, value in enumerate(list(pl.columns.values)):
            display_extracted_columns.insert("insert", F'{index + 1}. {value}\n')

    else:
        display_extracted_columns.delete(1.0, END)
        display_extracted_columns.insert("insert", "No file was selected")


def segregator(name_size_columns, name_delimiter):
    global pl
    global columns
    global option_menu_and_variables

    try:
        pl[["THICKNESS", "WIDTH"]] = pl[name_size_columns].str.split(name_delimiter, expand=True)

        try:
            pl["THICKNESS"] = pl["THICKNESS"].astype("float64").round(3)
            pl["WIDTH"] = pl["THICKNESS"].astype("int64")

        except ValueError:
            pass

        display_extracted_columns.config(display_extracted_columns.delete(1.0, END))

        # inserting columns names on each line:
        for index, value in enumerate(list(pl.columns.values)):
            display_extracted_columns.insert("insert", F'{index + 1}. {value}\n')

        # Insert names of columns into OptionMenu:
        columns = list(pl.columns.values)
        columns.insert(0, "None")
        columns.insert(1, "No, Single Item")

        for widgets_variables in option_menu_and_variables:
            for index, variable in enumerate(widgets_variables):
                variable.set_menu(*columns) if index == 0 else variable.set(columns[1])

        state_tool_status.delete(0, END)
        state_tool_status.insert("insert", "Segregation Successful!")

    except (ValueError, AttributeError):
        state_tool_status.delete(0, END)
        state_tool_status.insert("insert", "Segregation Failed!")


def use_segregator(check_button_state):
    if check_button_state == 1:

        get_size_column["state"] = ACTIVE
        get_delimiter.config(state=ACTIVE)
        segregate_columns.config(state=ACTIVE)

    elif check_button_state == 0:

        get_size_column["state"] = DISABLED

        get_delimiter.delete(0, END)
        get_delimiter.config(state=DISABLED)

        segregate_columns.config(state=DISABLED)

        state_tool_status.delete(0, END)


def calculate_ratios(packing_list):
    global ratio
    global tree_view_filled
    global unique_items

    if get_multiple_items_var.get() == "No, Single Item":

        ratio = packing_list.groupby(get_calculate_from_var.get(), as_index=False).agg(
            WEIGHT=(get_calculate_var.get(), "sum"),
            COILS=(get_calculate_var.get(), "count"))

        if tree_view_filled:
            for item in display_ratio_singular_tv.get_children():
                display_ratio_singular_tv.delete(item)

        else:
            tree_view_filled = True

    else:

        ratio = packing_list.groupby([get_multiple_items_var.get(), get_calculate_from_var.get()], as_index=False).agg(
            WEIGHT=(get_calculate_var.get(), "sum"),
            COILS=(get_calculate_var.get(), "count"))

        unique_items = pd.unique(ratio[get_multiple_items_var.get()])
        unique_items = np.insert(unique_items, 0, "Multiple Items")
        unique_items = np.insert(unique_items, 1, "Calculate All")
        option_menu.set_menu(*unique_items)

        if tree_view_filled:
            for item in display_ratio_singular_tv.get_children():
                display_ratio_singular_tv.delete(item)

        else:
            tree_view_filled = True

    ratio["WEIGHT"] = (ratio["WEIGHT"].astype("float64")).round(3)

    display_ratio_singular_tv["columns"] = list(ratio.columns)

    for cols in ratio.columns:
        display_ratio_singular_tv.heading(cols, text=cols)

    for cols_ in ratio.columns:
        display_ratio_singular_tv.column(str(cols_), minwidth=50, width=150, anchor=CENTER)

    for index, row in ratio.iterrows():
        display_ratio_singular_tv.insert("", "end", values=list(row))

    instructions.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky=N)

    display_ratio_singular_tv.grid_forget()
    display_ratio_singular_tv.grid(row=1, column=0, columnspan=2, pady=5, padx=10)

    import_selection.grid_forget()
    import_selection.grid(row=2, column=1, pady=10, padx=10)

    copy_data.grid_forget()
    copy_data.grid(row=2, column=0, pady=10, padx=10)


def copy_text(data_frame):
    data_frame.to_clipboard(index=False)


def add_widgets():
    global group_of_widgets_frame_3

    global serial_number
    global entry_from
    global label_seperator
    global entry_to

    group_of_widgets_frame_3 += 1

    serial_number.append("S.NO" + str(group_of_widgets_frame_3))
    entry_from.append("From" + str(group_of_widgets_frame_3))
    label_seperator.append(":" + str(group_of_widgets_frame_3))
    entry_to.append("To" + str(group_of_widgets_frame_3))

    serial_number[-1] = ttk.Label(frame_3, text=F"{group_of_widgets_frame_3}", borderwidth=2, relief=GROOVE,
                                  anchor=CENTER, style="Sub.TLabel")
    serial_number[-1].grid(row=4 + group_of_widgets_frame_3, column=0, padx=(80, 0), pady=5, ipadx=20, sticky=W + E)

    entry_from[-1] = ttk.Entry(frame_3, justify=CENTER, width=15)
    entry_from[-1].grid(row=4 + group_of_widgets_frame_3, column=1, padx=(5, 0), pady=5, sticky=E)

    label_seperator[-1] = ttk.Label(frame_3, text=":", style="Sub.TLabel")
    label_seperator[-1].grid(row=4 + group_of_widgets_frame_3, column=2, pady=5, padx=(40, 30), sticky=N)

    entry_to[-1] = ttk.Entry(frame_3, justify=CENTER, width=15)
    entry_to[-1].grid(row=4 + group_of_widgets_frame_3, column=3, padx=(0, 15), pady=5, sticky=W)

    delete_field.config(state=ACTIVE) if len(serial_number) > 0 else delete_field.config(state=DISABLED)
    add_field.config(state=DISABLED) if len(serial_number) == 20 else add_field.config(state=ACTIVE)


def delete_widgets():
    global group_of_widgets_frame_3

    global serial_number
    global entry_from
    global label_seperator
    global entry_to

    group_of_widgets_frame_3 -= 1

    serial_number[-1].destroy()
    entry_from[-1].destroy()
    label_seperator[-1].destroy()
    entry_to[-1].destroy()

    serial_number.pop()
    entry_from.pop()
    label_seperator.pop()
    entry_to.pop()

    delete_field.config(state=ACTIVE) if len(serial_number) > 0 else delete_field.config(state=DISABLED)
    add_field.config(state=DISABLED) if len(serial_number) == 20 else add_field.config(state=ACTIVE)


def display_range():
    global ratio

    ratio_ranges = []

    for ranges in range(len(serial_number)):

        _from = entry_from[ranges].get()
        _to = entry_to[ranges].get()

        if _from == "":
            _from = 0

        if _to == "":
            _to = 0

        ratio_ranges.append([float(_from), float(_to)])

    if (get_multiple_items_var.get() == "No, Single Item") or (option_menu_var.get() == "Calculate All"):

        for start, stop in ratio_ranges:

            if stop == 0:
                weight = (ratio.loc[ratio[get_calculate_from_var.get()] == float(start), 'WEIGHT']).sum()
                coils = (ratio.loc[ratio[get_calculate_from_var.get()] == float(start), 'COILS']).sum()

                display_ratio_ranges.insert("insert", F"{start:.2f} mm, {weight:.3f} MT, {coils} Coils\n")

            else:
                weight = (ratio.loc[ratio[get_calculate_from_var.get()] <= float(stop), 'WEIGHT']).sum() \
                         - (ratio.loc[ratio[get_calculate_from_var.get()] < float(start), 'WEIGHT'].sum())

                coils = (ratio.loc[ratio[get_calculate_from_var.get()] <= float(stop), 'COILS']).sum() \
                        - (ratio.loc[ratio[get_calculate_from_var.get()] < float(start), 'COILS'].sum())

                display_ratio_ranges.insert("insert",
                                            F"{start:.2f} mm - {stop:.2f} mm, {weight:.3f} MT, {coils} Coils\n")
    else:
        for start, stop in ratio_ranges:

            if stop == 0:
                expression_s = F"{get_multiple_items_var.get()} == '{option_menu_var.get()}' and ({get_calculate_from_var.get()} == {float(start)}) "

                weight = ratio.query(expression_s).WEIGHT.sum()

                coils = ratio.query(expression_s).WEIGHT.count()

                display_ratio_ranges.insert("insert", F"{option_menu_var.get()}, "
                                                      F"{start:.2f} mm,"
                                                      F" {weight:.3f} MT,"
                                                      F" {coils} Coils\n")

            else:
                expression_du = F"{get_multiple_items_var.get()} == '{option_menu_var.get()}' and " \
                                F"({get_calculate_from_var.get()} <= {float(stop)})"

                expression_dl = F"{get_multiple_items_var.get()} == '{option_menu_var.get()}' and " \
                                F"({get_calculate_from_var.get()} < {float(start)})"

                weight_u = ratio.query(expression_du).WEIGHT.sum()
                weight_l = ratio.query(expression_dl).WEIGHT.sum()

                weight = weight_u - weight_l

                coils_u = ratio.query(expression_du).WEIGHT.count()
                coils_l = ratio.query(expression_dl).WEIGHT.count()

                coils = coils_u - coils_l

                display_ratio_ranges.insert("insert",
                                            F"{option_menu_var.get()},"
                                            F" {start:.2f} mm - {stop:.2f} mm,"
                                            F" {weight:.3f} MT,"
                                            F" {coils} Coils\n")


def import_selection():
    selections = display_ratio_singular_tv.selection()

    # For Scrolled Text Widget:
    selections_list = [display_ratio_singular_tv.item(selection)["values"] for selection in selections]

    # Displaying Results:
    if get_multiple_items_var.get() == "No, Single Item":
        if len(selections_list) == 1:
            display_ratio_ranges.insert("insert", F"{selections_list[0][0]} mm,"
                                                  F" {selections_list[0][1]} MT,"
                                                  F" {m.trunc(float(selections_list[0][2]))} Coils"
                                                  F"\n")

        else:
            # Calculation for multiple:
            weight = 0.
            coils = 0

            valid_value = True
            min_value = selections_list[0][0]
            max_value = 0.

            for each_selection in selections_list:
                weight += float(each_selection[1])
                coils += m.trunc(float(each_selection[2]))

                try:
                    if float(each_selection[0]) <= float(min_value):
                        min_value = float(each_selection[0])

                    elif float(each_selection[0]) > max_value:
                        max_value = float(each_selection[0])

                except ValueError:
                    valid_value = False

            # Displaying Results:
            if valid_value:
                display_ratio_ranges.insert("insert", F"{min_value} mm"
                                                      F" - "
                                                      F"{max_value} mm, ")

            elif not valid_value:
                display_ratio_ranges.insert("insert", F"(Name your selection!), ")

            display_ratio_ranges.insert("insert", F"{weight:.3f} MT, "
                                                  F"{coils} Coils\n")

    else:
        if len(selections_list) == 1:
            display_ratio_ranges.insert("insert", F"{selections_list[0][0]},"
                                                  F" {selections_list[0][1]} mm,"
                                                  F" {selections_list[0][2]} MT,"
                                                  F" {m.trunc(float(selections_list[0][3]))} Coils"
                                                  F"\n")
        else:
            # Calculation for multiple:
            weight = 0.
            coils = 0

            valid_item = True
            valid_item_set = []

            valid_value = True

            min_value = selections_list[0][1]
            max_value = 0.

            for each_selection in selections_list:
                valid_item_set.append(each_selection[0])

                weight += float(each_selection[2])
                coils += m.trunc(float(each_selection[3]))

                try:
                    if float(each_selection[1]) <= float(min_value):
                        min_value = float(each_selection[1])

                    elif float(each_selection[1]) > max_value:
                        max_value = float(each_selection[1])

                except ValueError:
                    valid_value = False

                if len(set(valid_item_set)) > 1:
                    valid_item = False

            # Displaying Results:
            if valid_item:
                display_ratio_ranges.insert("insert", F"{valid_item_set[0]}, ")

            elif not valid_item:
                display_ratio_ranges.insert("insert", "(Multiple Item), ")

            if valid_value:
                display_ratio_ranges.insert("insert", F"{min_value} mm"
                                                      F" - "
                                                      F"{max_value} mm, ")

            elif not valid_value:
                display_ratio_ranges.insert("insert", "(Name your selection!), ")

            display_ratio_ranges.insert("insert", F"{weight:.3f} MT, "
                                                  F"{coils} Coils\n")


def copy_range():
    ranges = display_ratio_ranges.get(1.0, END)

    if get_multiple_items_var.get() == "No, Single Item":
        ranges = F"{get_calculate_from_var.get()}, WEIGHT, COILS\n" + ranges

    else:
        ranges = F"{get_multiple_items_var.get()}, {get_calculate_from_var.get()}, WEIGHT, COILS\n" + ranges

    x = ranges.split("\n")
    x = [y for y in x if y != ""]

    output = ""

    for rows in x:
        output += '\t'.join(rows.split(", ")) + '\n'

    pyperclip.copy(output)


# =====
# Global Variables:
# ===== #
excel_path = ""
pl = pd.DataFrame()
ratio = pd.DataFrame()
columns = []
option_menu_and_variables = []
unique_items = []

group_of_widgets_frame_3 = 0
serial_number = []
entry_from = []
label_seperator = []
entry_to = []

tree_view_filled = False

# ===== #
# Title:
# ===== #
title = ttk.Label(frame,
                  text="Ratio Maker",
                  font=("Cambria", 30),
                  background="#FFFFFF",
                  foreground="#425C5A")

title.grid(row=0, column=0,
           columnspan=2,
           pady=10,
           padx=10,
           sticky=NSEW)

# ===== #
# Styling:
# ===== #

style = ttk.Style()
style.theme_use("clam")

style.configure("TFrame",
                background="#425C5A")

style.configure("Main.TLabel",
                background="#425C5A",
                foreground="#FFC3A2")

style.configure("Sub.TLabel",
                background="#425C5A",
                foreground="#A2BFBD")

style.configure("TCheckbutton",
                background="#425C5A",
                foreground="#A2BFBD",
                indicatorbackground="#A2BFBD",
                indicatorforeground="#FFC3A2",
                font=("Cambria", 10))

style.configure("TEntry",
                fieldbackground="#A2BFBD",
                foreground="#425C5A")

style.map("TEntry",
          lightcolor=[('focus', '#425C5A')])

style.configure('TButton',
                focuscolor="#425C5A",
                background="#FFC3A2",
                foreground="#425C5A",
                font=("Cambria", 12))

style.map('TButton',
          background=[("pressed", "#425C5A")],
          foreground=[("pressed", "#FFC3A2")])

style.configure("TMenubutton",
                background="#A2BFBD",
                foreground="#425C5A",
                arrowcolor="#FFC3A2",
                font=("Cambria", 10),
                width=25)

style.map("TMenubutton",
          background=[("pressed", "#425C5A")],
          foreground=[("pressed", "#A2BFBD")])

style.configure("TCombobox",
                selectbackground="#A2BFBD",
                fieldbackground="#A2BFBD",
                background="#A2BFBD",
                selectforeground="#425C5A",
                foreground="#425C5A",
                arrowcolor="#FFC3A2",
                font=("Cambria", 10))

style.configure("Treeview",
                background="#A2BFBD",
                foreground="#425C5A",
                rowheight=20,
                fieldbackground="#A2BFBD")

style.configure("Treeview.Heading",
                background="#425C5A",
                foreground="#A2BFBD",
                fieldbackground="#425C5A",
                borderwidth=2,
                relief=GROOVE)

style.map("Treeview",
          background=[("selected", "#425C5A")],
          foreground=[("selected", "#A2BFBD")])

style.configure("TRadiobutton",
                background="#A2BFBD",
                foreground="#425C5A",
                font=("Cambria", 10))

# ===== #
# Packing List (frame_1):
# ===== #

# Calling Widgets (frame_1):
frame_1 = ttk.Frame(frame,
                    relief=SUNKEN,
                    borderwidth=2,
                    style="TFrame")

ask_excel_file = ttk.Label(frame_1,
                           text="Packing List",
                           anchor=CENTER,
                           font=("Cambria", 25),
                           style="Main.TLabel")

get_excel_file = ttk.Button(frame_1,
                            text="Open Microsoft Excel File",
                            command=lambda: open_packing_list("Note: In-order to open file on window, right click on "
                                                              "the file of your choice and then click on 'Open' to "
                                                              "view contents "))

horizontal_seperator_00 = ttk.Separator(frame_1,
                                        orient=HORIZONTAL)

display_excel_file_label = ttk.Label(frame_1,
                                     text="Details:",
                                     anchor=S,
                                     font=("Cambria", 12),
                                     style="Sub.TLabel")

display_excel_file_scrolled_text = ScrolledText(frame_1,
                                                font=("Cambria", 10, "bold"),
                                                background="#A2BFBD",
                                                foreground="#425C5A",
                                                height=9,
                                                width=50)

vertical_seperator_00 = ttk.Separator(frame_1,
                                      orient=VERTICAL)

specify_sheet_row = ttk.Label(frame_1,
                              text="Specify the following",
                              anchor=CENTER,
                              font=("Cambria", 25, "italic"),
                              style="Main.TLabel")

sheet_name_var = StringVar()
sheet_name_var.set("Sheet1")

row_number_var = IntVar()
row_number_var.set(1)

sheet_name_lb = ttk.Label(frame_1,
                          text="Sheet Name",
                          anchor=CENTER,
                          font=("Cambria", 15, "italic"),
                          style="Sub.TLabel")

sheet_name_et = ttk.Entry(frame_1,
                          textvariable=sheet_name_var,
                          justify=CENTER,
                          width=25)

vertical_seperator_01 = ttk.Separator(frame_1,
                                      orient=VERTICAL)

row_number_lb = ttk.Label(frame_1,
                          text="Row Number",
                          anchor=CENTER,
                          font=("Cambria", 15, "italic"),
                          style="Sub.TLabel")

row_number_et = ttk.Entry(frame_1,
                          textvariable=row_number_var,
                          justify=CENTER,
                          width=25)

get_columns = ttk.Button(frame_1,
                         text="Extract Column Names",
                         width=15,
                         command=lambda: get_col_names(excel_path, sheet_name_var.get(), row_number_var.get()))

vertical_seperator_02 = ttk.Separator(frame_1,
                                      orient=VERTICAL)

display_columns = ttk.Label(frame_1,
                            text="Extracted Columns",
                            anchor=W,
                            font=("Cambria", 12))

display_extracted_columns = ScrolledText(frame_1,
                                         background="#A2BFBD",
                                         foreground="#425C5A",
                                         font=("Cambria", 10, "bold"),
                                         height=5,
                                         width=10)

horizontal_seperator_01 = ttk.Separator(frame_1,
                                        orient=HORIZONTAL)

# Setting up, Tool to separate Size -> Thickness & Width:
tool_to_segregate = IntVar()
tool_to_segregate.set(0)

# Converting Size (0.11*1200) to Thickness (0.11) & Width (1200) - Temporarily
ask_size_to_thickness_width_tool = ttk.Checkbutton(frame_1,
                                                   text="Segregate",
                                                   variable=tool_to_segregate,
                                                   takefocus=0,
                                                   command=lambda: use_segregator(tool_to_segregate.get()))

ask_size_column = ttk.Label(frame_1,
                            text="Column Name:",
                            anchor=W,
                            font=("Cambria", 12),
                            style="Sub.TLabel")

size_tool = StringVar()
size_tool.set("Column Names")

get_size_column = ttk.OptionMenu(frame_1,
                                 size_tool,
                                 "Column Names",
                                 *columns)
get_size_column["state"] = DISABLED

ask_delimiter = ttk.Label(frame_1,
                          text="Delimiter:",
                          anchor=W,
                          font=("Cambria", 12),
                          style="Sub.TLabel")

delimiter = StringVar()
delimiter.set("Delimiter")

delimiter_options = ["*", ",", "x", "/t"]

get_delimiter = ttk.Combobox(frame_1,
                             justify=CENTER,
                             state=DISABLED,
                             textvariable=delimiter,
                             values=delimiter_options,
                             font=("Cambria", 10),
                             width=10)
get_delimiter.option_add('*TCombobox*Listbox.Justify', 'center')

segregate_columns = ttk.Button(frame_1,
                               text="Segregate Columns",
                               width=15,
                               state=DISABLED,
                               command=lambda: segregator(size_tool.get(), get_delimiter.get()))

status_tool = ttk.Label(frame_1,
                        text="Status:",
                        anchor=W,
                        font=("Cambria", 12),
                        style="Sub.TLabel")

state_tool_status = ttk.Entry(frame_1,
                              justify=CENTER,
                              width=25)

sep3 = ttk.Separator(frame_1)

vertical_seperator_03 = ttk.Separator(frame_1,
                                      orient=VERTICAL)

ask_multiple_items = ttk.Label(frame_1,
                               text="Multiple Items:",
                               anchor=SW,
                               font=("Cambria", 12),
                               style="Sub.TLabel")

get_multiple_items_var = StringVar()
get_multiple_items_var.set("Select the columns with items")

get_multiple_items = ttk.OptionMenu(frame_1,
                                    get_multiple_items_var,
                                    *columns)

ask_calculate = ttk.Label(frame_1,
                          text="Calculate:",
                          anchor=SW,
                          font=("Cambria", 12),
                          style="Sub.TLabel")

get_calculate_var = StringVar()
get_calculate_var.set("Select a value to calculate with")
get_calculate = ttk.OptionMenu(frame_1,
                               get_calculate_var,
                               *columns)

ask_calculate_from = ttk.Label(frame_1,
                               text="From:",
                               anchor=SW,
                               font=("Cambria", 12),
                               style="Sub.TLabel")

get_calculate_from_var = StringVar()
get_calculate_from_var.set("Select a value to calculate from")
get_calculate_from = ttk.OptionMenu(frame_1,
                                    get_calculate_from_var,
                                    *columns)

option_menu_and_variables = [[get_size_column, size_tool],
                             [get_multiple_items, get_multiple_items_var],
                             [get_calculate, get_calculate_var],
                             [get_calculate_from, get_calculate_from_var]]

vertical_seperator_04 = ttk.Separator(frame_1,
                                      orient=VERTICAL)

calculate_ratio = ttk.Button(frame_1,
                             text="Calculate Ratio",
                             width=15,
                             command=lambda: calculate_ratios(pl))

# Placing Widgets (frame_1)
frame_1.grid(row=1, column=0, columnspan=3, padx=15, pady=10, sticky=NW)

ask_excel_file.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky=NSEW)
get_excel_file.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky=NSEW)

display_excel_file_label.grid(row=2, column=0, columnspan=2, padx=10, ipadx=5, ipady=5, sticky=NSEW)
display_excel_file_scrolled_text.grid(row=3, rowspan=3, column=0, columnspan=2, padx=10, sticky=N)

vertical_seperator_00.grid(row=0, rowspan=6, column=2, sticky=N + S)

specify_sheet_row.grid(row=0, rowspan=2, column=3, columnspan=3, padx=5, pady=5, sticky=N + S)

sheet_name_lb.grid(row=2, column=3, padx=5, pady=5, sticky=NSEW)
sheet_name_et.grid(row=3, column=3, padx=5, pady=5, sticky=N)

vertical_seperator_01.grid(row=2, rowspan=2, column=4, sticky=N + S)

row_number_lb.grid(row=2, column=5, padx=5, pady=5, sticky=NSEW)
row_number_et.grid(row=3, column=5, padx=5, pady=5, sticky=N)

get_columns.grid(row=4, rowspan=2, column=3, columnspan=3, padx=5, pady=5, sticky=W + E)

vertical_seperator_02.grid(row=0, rowspan=6, column=6, sticky=N + S)

display_extracted_columns.grid(row=0, column=7, columnspan=2, padx=25, pady=5, sticky=NSEW)

horizontal_seperator_00.grid(row=1, column=7, columnspan=2, sticky=W + E)

segregate_columns.grid(row=2, column=7, padx=5, pady=5, sticky=W + E)
ask_size_to_thickness_width_tool.grid(row=2, column=8, padx=5, pady=5, sticky=N + S)

ask_size_column.grid(row=3, column=7, padx=5, pady=5, sticky=NSEW)
get_size_column.grid(row=3, column=8, padx=5, pady=5, sticky=W + E)

ask_delimiter.grid(row=4, column=7, padx=5, pady=5, sticky=NSEW)
get_delimiter.grid(row=4, column=8, padx=5, pady=5, sticky=W + E)

status_tool.grid(row=5, column=7, padx=5, pady=5, sticky=NSEW)
state_tool_status.grid(row=5, column=8, padx=5, pady=5, sticky=W + E)

vertical_seperator_03.grid(row=0, rowspan=6, column=9, sticky=N + S)

ask_multiple_items.grid(row=0, column=10, padx=5, pady=10, sticky=NSEW)
get_multiple_items.grid(row=1, column=10, padx=5, pady=10, sticky=W + E)

ask_calculate.grid(row=2, column=10, padx=5, pady=10, sticky=NSEW)
get_calculate.grid(row=3, column=10, padx=5, pady=10, sticky=W + E)

ask_calculate_from.grid(row=4, column=10, padx=5, pady=10, sticky=NSEW)
get_calculate_from.grid(row=5, column=10, padx=5, pady=10, sticky=W + E)

vertical_seperator_04.grid(row=0, rowspan=6, column=11, sticky=N + S)
calculate_ratio.grid(row=0, rowspan=6, column=12, padx=25, pady=10, sticky=W + E)

# ===== #
# Displaying Ratio (Single) - frame_2
# ===== #

# Calling Widgets:
frame_2 = ttk.Frame(frame,
                    relief=SUNKEN,
                    borderwidth=2)

instructions = ttk.Label(frame_2,
                         text="Ctrl + Click: to select specific rows | Shift + Click: to select multiple rows",
                         anchor=N,
                         font=("Cambria", 12, "bold"),
                         style="Sub.TLabel")

display_ratio_singular_tv = ttk.Treeview(frame_2,
                                         show="headings",
                                         selectmode="extended",
                                         height=20)

copy_data = ttk.Button(frame_2,
                       text="Copy Text",
                       command=lambda: copy_text(ratio))

import_selection = ttk.Button(frame_2,
                              text="Import Selection",
                              command=import_selection)

# Placing Widgets:
frame_2.grid(row=2, column=0, padx=(15, 5), pady=5, sticky=NSEW)

display_ratio_singular_tv.grid(row=1, column=0, columnspan=2, pady=5, padx=(50, 0), ipadx=175, sticky=NSEW)

copy_data.grid(row=2, column=0, pady=10, padx=(50, 0), sticky=NSEW)
import_selection.grid(row=2, column=1, pady=10, padx=(50, 0), sticky=NSEW)

# ===== #
# Taking inputs for range - frame_3
# ===== #

# Calling Widgets:
frame_3 = ttk.Frame(frame,
                    relief=SUNKEN,
                    borderwidth=2)

welcome_to_ranges = ttk.Label(frame_3,
                              text="Ratio Ranges",
                              anchor=N,
                              font=("Cambria", 20, "bold"),
                              style="Main.TLabel")

get_range = ttk.Button(frame_3,
                       text="Get Range",
                       command=display_range)

option_menu_var = StringVar()
option_menu = ttk.OptionMenu(frame_3,
                             option_menu_var,
                             "Multiple Items",
                             *unique_items)

sep1_frame_3 = ttk.Separator(frame_3)

add_field = ttk.Button(frame_3,
                       text="Add Range (+)",
                       command=lambda: add_widgets())

delete_field = ttk.Button(frame_3,
                          text="Delete Range (-)",
                          state=DISABLED,
                          command=lambda: delete_widgets())

# Placing Widgets:
frame_3.grid(row=2, column=1, padx=5, pady=5, sticky=NSEW)

welcome_to_ranges.grid(row=0, column=0, columnspan=4, padx=(55, 0), ipadx=150, pady=10, sticky=NSEW)

get_range.grid(row=1, column=0, columnspan=2, padx=(75, 5), pady=10, ipadx=50, sticky=NSEW)
option_menu.grid(row=1, column=2, columnspan=2, padx=(5, 55), pady=10, sticky=NSEW)

sep1_frame_3.grid(row=2, column=0, columnspan=4, sticky=W + E, padx=(75, 55))

add_field.grid(row=3, column=0, columnspan=2, padx=(75, 5), ipadx=50, pady=10, sticky=NSEW)
delete_field.grid(row=3, column=2, columnspan=2, padx=(5, 55), pady=10, sticky=NSEW)

# ===== #
# Displaying ratio ranges - frame_4
# ===== #

# Calling Widgets:
frame_4 = ttk.Frame(frame)

display_ratio_ranges = ScrolledText(frame_4,
                                    font=("Cambria", 10, "bold"),
                                    height=25,
                                    width=35,
                                    foreground="#425C5A",
                                    background="#A2BFBD")

clear_text_box = ttk.Button(frame_4,
                            text="Clear Text",
                            command=lambda: display_ratio_ranges.delete(1.0, END))

copy_ranges = ttk.Button(frame_4,
                         text="Copy Text",
                         command=copy_range)

# Placing Widgets:
frame_4.grid(row=2, column=2, padx=(5, 15), pady=5, sticky=NSEW)

display_ratio_ranges.grid(row=0, column=0, columnspan=2, padx=(15, 0), pady=10)

clear_text_box.grid(row=1, column=0, padx=(15, 0), pady=10, sticky=NSEW)
copy_ranges.grid(row=1, column=1, padx=(15, 0), pady=10, sticky=NSEW)

root.mainloop()
