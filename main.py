import tkinter as tk
from tkinter import filedialog

import pandas as pd
import os

file_path = ""
sheet_name_var = ""


def browse_file():
    def get_file():
        global file_path
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        update_sheet_names(file_path)
        selected_file_label.config(text = f"Selected Excel File: {file_path}")

    def update_sheet_names(f):
        try:

            sheet_names = pd.read_excel(f, sheet_name=None).keys()
            for name in sheet_names:
                sheet_name_menu['menu'].add_command(label=name, command=tk._setit(sheet_name_svar, name))
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
    global sheet_name_svar
    sheet_name_svar.set("")
    sheet_name_label = tk.Label(browse_file_window, text="Select Sheet Name:")
    sheet_name_label.pack()
    sheet_name_menu = tk.OptionMenu(browse_file_window, sheet_name_svar, "")
    sheet_name_menu.pack()  
    sheet_name_menu['menu'].delete(0, 'end')
    confirm_button = tk.Button(browse_file_window, text="Confirm file", command=browse_file_window.destroy)
    confirm_button.pack()
    selected_file_label = tk.Label(app, text=f"")
    selected_file_label.pack()

    



def filter_columns():

    
    sheet_name_var = sheet_name_svar.get() #convert stringvar to string

    def update_columns(file_path):

        if not file_path:
            print("Please select an Excel file.")
            return

        if not sheet_name_var:
            print("Please select a sheet name.")
            return

        try:
            main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
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
        if not file_path:
            print("Please select file.")
            return
        if not sheet_name_var:
            print("Please select file.")
            return

        try:
            main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
            extracted_df = main_df[selected_columns].copy()
            
            preview.config(text=extracted_df.head().to_string(index=False))
            
            print(extracted_df.head())
        except Exception as e:
            print(f"Error: {e}")

    def on_sheet_name_change(*args):
        process_selected_columns()

    filter_columns_window = tk.Toplevel()
    filter_columns_window.geometry("500x400")
    filter_columns_window.title("Filter Columns")
    available_columns_label = tk.Label(filter_columns_window, text="Available Columns:")
    available_columns_label.pack()
    available_columns = tk.Listbox(filter_columns_window, selectmode=tk.MULTIPLE)
    available_columns.pack()
    sheet_name_svar.trace("w", on_sheet_name_change)
    
    process_columns_button = tk.Button(filter_columns_window,text="Process columns",command=process_selected_columns)
    process_columns_button.pack()
    preview = tk.Label(filter_columns_window,text = "", justify=tk.CENTER)
    preview.pack()
    confirm_button = tk.Button(filter_columns_window, text="Confirm", command=filter_columns_window.destroy)
    confirm_button.pack()

    update_columns(file_path)


def export_file():

    sheet_name_var = sheet_name_svar.get()
    
    if not file_path:
        print("Please select an Excel file.")
        return

    if not sheet_name_var:
        print("Please select a sheet name.")
        return

    try:
        main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if export_path:
            main_df.to_excel(export_path, index=False)
            exported_window = tk.Toplevel()
            exported_window.geometry("200x200")
            exported_window.title("Created Successfully")
            export_label=tk.Label(exported_window,text=f"Exported Successfully to {export_path}. Click button to open!!")
            export_label.pack()
            export_open_button = tk.Button(exported_window, text="Open File", command=os.startfile(export_file))
            export_open_button.pack()
    except Exception as e:
        print(f"Error: {e}")


# Create the main application window
app = tk.Tk()
app.geometry("500x300")
app.title("Data Cleaning and Preprocessing Tool")
sheet_name_svar = tk.StringVar()

select_file_button = tk.Button(app, text="Select/Change File", command=browse_file)
select_file_button.pack(pady=10)

filter_columns_button = tk.Button(app, text="Filter Columns", command=filter_columns)
filter_columns_button.pack(pady=10)

export_file_button = tk.Button(app, text="Export File", command=export_file)
export_file_button.pack(pady=10)


# Run the application
app.mainloop()
