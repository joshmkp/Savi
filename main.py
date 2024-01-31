import os

import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def browse_file1():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file1_entry.delete(0, tk.END)
    file1_entry.insert(0, filename)

def browse_file2():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file2_entry.delete(0, tk.END)
    file2_entry.insert(0, filename)

def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:  # Check if a folder was selected
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_selected)

def process_files():
    file1 = file1_entry.get()
    file2 = file2_entry.get()
    output_folder = output_folder_entry.get()

    # Validation
    if not file1 or not file2 or not output_folder:
        messagebox.showerror("Error", "All fields are required")
        return

    # Load the data from Excel files
    try:
        file1_df = pd.read_excel(file1)
    except Exception as e:
        messagebox.showerror("Error","Current Price List could not be loaded")

    try:
        file2_df = pd.read_excel(file2)
    except Exception as e:
        messagebox.showerror("Error","Updated Price List could not be loaded")



    # Record original prices for comparisson
    original_prices = file1_df[['SKU', 'Product Cost']].copy()

    # Update prices in file1_df using a merge operation
    print("Merging files ...")
    updated_df = file1_df.merge(file2_df[['SKU', 'Product Cost']], on='SKU', how='left', suffixes=('', '_new'))
    print("Updating data ...")
    updated_df['Product Cost'] = updated_df['Product Cost_new'].combine_first(updated_df['Product Cost'])
    updated_df.drop('Product Cost_new', axis=1, inplace=True)

    # Save the updated DataFrame to a new Excel file
    output_file_path = os.path.join(output_folder, "updated_file.xlsx")
    updated_df.to_excel(output_file_path, index=False)

    # Identify which prices were updated
    updated_prices = updated_df[updated_df['Product Cost'] != original_prices['Product Cost']]

    if not updated_prices.empty:
        updated_count = len(updated_prices)
        results_text.insert(tk.END, f"{updated_count} Product Costs were updated for the following SKUs:\n")
        for index, row in updated_prices.iterrows():
            results_text.insert(tk.END, f"SKU: {row['SKU']}\n")
            results_text.insert(tk.END, f"Updated Product Cost: {row['Product Cost']}\n")
            results_text.insert(tk.END, 50 * "-" + "\n")
    else:
        results_text.insert(tk.END, "No Product Costs were updated.\n")

    messagebox.showinfo("Success", "Files processed successfully")

# Set up the main application window
root = tk.Tk()
root.title("Product Cost Updater")

# File 1
tk.Label(root, text="Current Price List:").grid(row=0, column=0, sticky=tk.W)
file1_entry = tk.Entry(root, width=50)
file1_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=browse_file1).grid(row=0, column=2)

# File 2
tk.Label(root, text="Updated Price List:").grid(row=1, column=0, sticky=tk.W)
file2_entry = tk.Entry(root, width=50)
file2_entry.grid(row=1, column=1)
tk.Button(root, text="Browse", command=browse_file2).grid(row=1, column=2)



#Output folder
tk.Label(root, text="Output Folder:").grid(row=6, column=0, sticky=tk.W)
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=6, column=1)
tk.Button(root, text="Browse", command=browse_output_folder).grid(row=6, column=2)


#  Text widget for displaying results
results_text = tk.Text(root, height=10, width=50)
results_text.grid(row=5, column=0, columnspan=3)


# Process Button
process_button = tk.Button(root, text="Update and Export", command=process_files)
process_button.grid(row=4, column=1, sticky=tk.W)

# Run the application
root.mainloop()





