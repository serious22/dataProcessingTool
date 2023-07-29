import tkinter as tk
from tkinter import filedialog

import pandas as pd


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)
    update_column_list(file_path)


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
app.title("Data Cleaning and Preprocessing Tool")

select_file_button = tk.Button(app, text="Select File", command=browse_file)
select_file_button.pack(pady=10)

filter_columns_button = tk.Button(app, text="Filter Columns", command=select_columns)
filter_columns_button.pack(pady=10)

export_file_button = tk.Button(app, text="Export File", command=export_file)
export_file_button.pack(pady=10)

# Run the application
app.mainloop()
