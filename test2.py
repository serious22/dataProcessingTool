import tkinter as tk
from tkinter import filedialog

import pandas as pd


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)
    update_sheet_names(file_path)


def update_sheet_names(file_path):
    try:
        sheet_names = pd.read_excel(file_path, sheet_name=None).keys()
        sheet_name_var.set("")  # Clear the current selection
        sheet_name_menu['menu'].delete(0, 'end')  # Clear existing options
        for name in sheet_names:
            sheet_name_menu['menu'].add_command(label=name, command=tk._setit(sheet_name_var, name))
    except Exception as e:
        print(f"Error: {e}")


def select_columns():
    file_path = file_entry.get()
    sheet_name = sheet_name_var.get()

    if not file_path:
        print("Please select an Excel file.")
        return

    if not sheet_name:
        print("Please select a sheet name.")
        return

    try:
        main_df = pd.read_excel(file_path, sheet_name=sheet_name)
        column_headers = main_df.columns.tolist()
        available_columns.delete(0, tk.END)
        for col in column_headers:
            available_columns.insert(tk.END, col)
    except Exception as e:
        print(f"Error: {e}")


def process_selected_columns():
    selected_columns = []
    for index in available_columns.curselection():
        selected_columns.append(available_columns.get(index))
    file_path = file_entry.get()
    sheet_name = sheet_name_var.get()

    if not file_path:
        print("Please select an Excel file.")
        return

    if not sheet_name:
        print("Please select a sheet name.")
        return

    try:
        main_df = pd.read_excel(file_path, sheet_name=sheet_name)
        extracted_df = main_df[selected_columns].copy()

        print("Extracted Data:")
        print(extracted_df.head())
    except Exception as e:
        print(f"Error: {e}")


def export_file():
    file_path = file_entry.get()
    sheet_name = sheet_name_var.get()

    if not file_path:
        print("Please select an Excel file.")
        return

    if not sheet_name:
        print("Please select a sheet name.")
        return

    try:
        main_df = pd.read_excel(file_path, sheet_name=sheet_name)
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if export_path:
            main_df.to_excel(export_path, index=False)
            print(f"File exported successfully to: {export_path}")
    except Exception as e:
        print(f"Error: {e}")


# Create the main application window
app = tk.Tk()
app.title("Data Cleaning and Preprocessing Tool")

# File Selection
file_label = tk.Label(app, text="Select Excel File:")
file_label.pack()
file_entry = tk.Entry(app, width=50)
file_entry.pack()
browse_button = tk.Button(app, text="Browse", command=browse_file)
browse_button.pack()

# Sheet Name Selection
sheet_name_var = tk.StringVar()
sheet_name_var.set("")  # Set the initial selection to empty
sheet_name_label = tk.Label(app, text="Select Sheet Name:")
sheet_name_label.pack()
sheet_name_menu = tk.OptionMenu(app, sheet_name_var, "")
sheet_name_menu.pack()

# Update available columns when sheet name changes
def on_sheet_name_change(*args):
    select_columns()

sheet_name_var.trace("w", on_sheet_name_change)

# Available Columns Listbox
available_columns_label = tk.Label(app, text="Available Columns:")
available_columns_label.pack()
available_columns = tk.Listbox(app, selectmode=tk.MULTIPLE)
available_columns.pack()

# Process Columns
process_button = tk.Button(app, text="Process Selected Columns", command=process_selected_columns)
process_button.pack()

# Export File
export_button = tk.Button(app, text="Export File", command=export_file)
export_button.pack()

# Run the application
app.mainloop()
