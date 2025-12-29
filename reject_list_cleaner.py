"""
Reject List Cleaner

This script removes rejected email addresses from Excel mailing lists based on
a predefined reject list. It is designed to automate email list maintenance tasks.

Developed during internship work using anonymized data.
"""
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import os

# ========== Main processing function ==========
def process_cleaning():
    excel_file = file_textbox.get("1.0", tk.END).strip()
    if not excel_file:
        messagebox.showwarning("No File Path", "Please enter the Excel file path.")
        return
    
    # Get selected list sheets
    list_selected_indices = list_sheet_listbox.curselection()
    list_sheet_names = [list_sheet_listbox.get(i) for i in list_selected_indices]
    
    # Get selected reject sheet
    reject_sheet_name = reject_sheet_var.get()
    
    if not list_sheet_names or not reject_sheet_name:
        messagebox.showwarning("Sheet Selection", "Please select at least one list sheet and a reject sheet.")
        return
    
    try:
        # Load reject emails
        df_reject = pd.read_excel(excel_file, sheet_name=reject_sheet_name)
        email_col_reject = df_reject.columns[2]  # Email in 3rd column
        reject_emails = df_reject[email_col_reject].str.lower().dropna().unique()
        
        deleted_total = 0
        
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for sheet in list_sheet_names:
                df_list = pd.read_excel(excel_file, sheet_name=sheet)
                email_col_list = df_list.columns[2]  # Email in 3rd column
                df_list[email_col_list] = df_list[email_col_list].str.lower()
                
                initial_count = len(df_list)
                df_list_cleaned = df_list[~df_list[email_col_list].isin(reject_emails)]
                deleted_count = initial_count - len(df_list_cleaned)
                deleted_total += deleted_count
                
                # Save cleaned list sheet
                df_list_cleaned.to_excel(writer, sheet_name=sheet, index=False)
                
                # Set column width ~3.87 cm (25 units)
                worksheet = writer.sheets[sheet]
                for col in worksheet.columns:
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = 25
            
            # Also overwrite reject sheet to keep formatting
            df_reject.to_excel(writer, sheet_name=reject_sheet_name, index=False)
            worksheet_reject = writer.sheets[reject_sheet_name]
            for col in worksheet_reject.columns:
                col_letter = col[0].column_letter
                worksheet_reject.column_dimensions[col_letter].width = 25
        
        # Completion message
        messagebox.showinfo(
            "Completed",
            f"✅ Deleted total {deleted_total} rows from selected list sheets based on \"{reject_sheet_name}\" emails.\n"
        )
        
        # Ask if user wants to open Excel
        if messagebox.askyesno("Open Excel?", "Do you want to open the updated Excel file?"):
            os.startfile(excel_file)
        
        root.destroy()
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# ========== Load sheet names ==========
def load_sheets():
    excel_file = file_textbox.get("1.0", tk.END).strip()
    if not excel_file:
        messagebox.showwarning("No File Path", "Please enter the Excel file path first.")
        return
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        
        # Populate list sheet listbox
        list_sheet_listbox.delete(0, tk.END)
        for sheet in sheet_names:
            list_sheet_listbox.insert(tk.END, sheet)
        
        # Populate reject sheet combobox
        reject_sheet_dropdown['values'] = sheet_names
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read sheets:\n{e}")

# ========== Tkinter GUI ==========
root = tk.Tk()
root.title("Data Cleaning Tool")

tk.Label(root, text="Enter Excel file path:").pack(padx=10, pady=5)
file_textbox = tk.Text(root, width=80, height=2)
file_textbox.pack(padx=10, pady=5)

tk.Button(root, text="Load Sheets", command=load_sheets).pack(pady=5)

tk.Label(root, text="Select source (to) sheets:").pack(padx=10, pady=2)
list_sheet_listbox = tk.Listbox(root, selectmode='multiple', width=50, height=5)
list_sheet_listbox.pack(padx=10, pady=2)

reject_sheet_var = tk.StringVar()
tk.Label(root, text="Select reject (from) sheet:").pack(padx=10, pady=2)
reject_sheet_dropdown = ttk.Combobox(root, textvariable=reject_sheet_var, width=50)
reject_sheet_dropdown.pack(padx=10, pady=2)

tk.Button(root, text="Run Cleaning", command=process_cleaning).pack(pady=10)

root.mainloop()

