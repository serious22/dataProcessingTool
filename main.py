import tkinter as tk
from tkinter import filedialog

import pandas as pd



def browse_file():
    def get_file():
        global file_path
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        update_sheet_names(file_path)
    def update_sheet_names(f):
        try:
            sheet_names = pd.read_excel(f, sheet_name=None).keys()
            for name in sheet_names:
                sheet_name_menu['menu'].add_command(label=name, command=tk._setit(sheet_name_var, name))
        except Exception as e:
            print(f"Error: {e}")
    
    browse_file_window = tk.Toplevel()
    browse_file_window.title("Select File")
    browse_file_window.geometry("250x150")
    file_label = tk.Label(browse_file_window, text="Select Excel File:")
    file_label.pack()
    file_entry = tk.Entry(browse_file_window, width=50)
    file_entry.pack()
    browse_button = tk.Button(browse_file_window, text="Browse", command=get_file)
    browse_button.pack()
    global sheet_name_var
    sheet_name_var = tk.StringVar()
    sheet_name_var.set("")
    sheet_name_label = tk.Label(browse_file_window, text="Select Sheet Name:")
    sheet_name_label.pack()
    sheet_name_menu = tk.OptionMenu(browse_file_window, sheet_name_var, "")
    sheet_name_menu.pack()  
    sheet_name_menu['menu'].delete(0, 'end')
    sheet_confirm_button = tk.Button(browse_file_window, text="Confirm file", command=browse_file_window.destroy)
    sheet_confirm_button.pack()

    def on_sheet_name_change(*args):
        select_columns()

    sheet_name_var.trace("w", on_sheet_name_change)

    
    


def update_column_list(file_path):
    sheet_name = sheet_name_entry.get()
    if not sheet_name:
        return

    try:
        main_df = pd.read_excel(file_path, sheet_name=sheet_name)
        column_headers = main_df.columns.tolist()
        available_columns.delete(0, tk.END)
        for col in column_headers:
            available_columns.insert(tk.END, col)
    except Exception as e:
        print(f"Error: {e}")


def select_columns():
    print("Hello here is the file........................." ,sheet_name_var)
    selected_columns = []
    for index in available_columns.curselection():
        selected_columns.append(available_columns.get(index))
    process_selected_columns(selected_columns)


def process_selected_columns(selected_columns):
    file_path = file_entry.get()
    
    if not file_path:
        return

    sheet_name = sheet_name_entry.get()
    if not sheet_name:
        print("Please enter a sheet name.")
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
    if not file_path:
        return

    sheet_name = sheet_name_entry.get()
    if not sheet_name:
        print("Please enter a sheet name.")
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
app.geometry("500x300")
app.title("Data Cleaning and Preprocessing Tool")


select_file_button = tk.Button(app, text="Select File", command=browse_file)
select_file_button.pack(pady=10)

filter_columns_button = tk.Button(app, text="Filter Columns", command=select_columns)
filter_columns_button.pack(pady=10)

export_file_button = tk.Button(app, text="Export File", command=export_file)
export_file_button.pack(pady=10)


# Run the application
app.mainloop()
