import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox


import pandas as pd
import os

file_path = ""
sheet_name_var = ""
main_df = pd.DataFrame()
final_df = pd.DataFrame()

def browse_file():
    def get_file():
        global file_path
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        update_sheet_names()
        selected_file_label.config(text = f"Selected Excel File: {file_path}")

    def update_sheet_names():
        try:

            sheet_names = pd.read_excel(file_path, sheet_name=None).keys()
            for name in sheet_names:
                sheet_name_menu['menu'].add_command(label=name, command=tk._setit(sheet_name_svar, name))
            
        except Exception as e:
            messagebox.showerror("Error" ,f"Error: {e}")
    def update_main_df():
        global main_df
        sheet_name_var = sheet_name_svar.get()
        main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
        browse_file_window.destroy()
   
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
    confirm_button = tk.Button(browse_file_window, text="Confirm file", command=update_main_df)
    confirm_button.pack()
    selected_file_label = tk.Label(app, text=f"")
    selected_file_label.pack()



def filter_columns():
    if main_df.empty :
        messagebox.showwarning("File not found","Please select an Excel file.")
        return
    sheet_name_var = sheet_name_svar.get() #convert stringvar to string

    def update_columns(file_path):

        if not file_path:
            messagebox.showwarning("File not found","Please select an Excel file.")
            return

        if not sheet_name_var:
            messagebox.showwarning("No sheet selected","Please select a sheet name.")
            return

        try:
            main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
            column_headers = main_df.columns.tolist()
            available_columns.delete(0, tk.END)
            for col in column_headers:
                available_columns.insert(tk.END, col)
        except Exception as e:
            messagebox.showerror("Error" ,f"Error: {e}")

        
    
    def process_selected_columns():
        
        selected_columns = []
        for index in available_columns.curselection():
            selected_columns.append(available_columns.get(index))
        if not file_path:
            messagebox.showwarning("File not found","Please select an Excel file.")
            return
        if not sheet_name_var:
            messagebox.showwarning("No sheet selected","Please select a sheet name.")
            return

        try:
            global final_df
            main_df = pd.read_excel(file_path, sheet_name=sheet_name_var)
            final_df = main_df[selected_columns].copy()
            
            preview.config(text=final_df.head().to_string(index=False))
        except Exception as e:
            messagebox.showerror("Error" ,f"Error: {e}")

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

def check_missing_values():
    def new(data_frame):
        if main_df.empty:
            messagebox.showwarning("File not found","Please select an Excel file.")
            return
        
       


        missing_values_count = data_frame.isnull().sum()
        columns_with_missing_values = missing_values_count[missing_values_count > 0].index.tolist()
        total_missing_values_per_column = missing_values_count[missing_values_count > 0].tolist()
        missing_values_tree.delete(*missing_values_tree.get_children())
        for col, total_missing in zip(columns_with_missing_values, total_missing_values_per_column):
            missing_values_tree.insert("", "end", values=(col, total_missing))
        
    check_missing_values_window = tk.Toplevel()
    check_missing_values_window.geometry("500x400")
    check_missing_values_window.title('Missing Values')
    table_frame = ttk.Frame(check_missing_values_window)
    table_frame.pack(padx=10, pady=10)
    missing_values_tree = ttk.Treeview(table_frame, columns=("Column Name", "Total Missing Values"), show="headings", height=5)
    missing_values_tree.heading("Column Name", text="Column Name")
    missing_values_tree.heading("Total Missing Values", text="Total Missing Values")
    missing_values_tree.pack()
    missing_main = tk.Button(check_missing_values_window,text = "Original Data",command=new(main_df))
    missing_main.pack()
    missing_filter = tk.Button(check_missing_values_window,text = "Filtered  Data",command=new(final_df))
    missing_filter.pack()
def export_file():
    def open_export_file():
        os.startfile(export_path)
    try:
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if export_path:
            final_df.to_excel(export_path, index=False)
            exported_window = tk.Toplevel()
            exported_window.geometry("200x200")
            exported_window.title("Created Successfully")
            export_label=tk.Label(exported_window,text=f"Exported Successfully to {export_path}. Click button to open!!")
            export_label.pack()
            export_open_button = tk.Button(exported_window, text="Open File", command=open_export_file)
            export_open_button.pack()
    except Exception as e:
        messagebox.showerror("Error" ,f"Error: {e}")



# Create the main application window
app = tk.Tk()
app.geometry("500x300")
app.title("Data Cleaning and Preprocessing Tool")
sheet_name_svar = tk.StringVar()

select_file_button = tk.Button(app, text="Select/Change File", command=browse_file)
select_file_button.pack(pady=10)

filter_columns_button = tk.Button(app, text="Filter Columns", command=filter_columns)
filter_columns_button.pack(pady=10)

check_missing_values_button = tk.Button(app, text="Check Missing Values", command=check_missing_values)
check_missing_values_button.pack(pady=10)


export_file_button = tk.Button(app, text="Export File", command=export_file)
export_file_button.pack(pady=10)



# Run the application
app.mainloop()
