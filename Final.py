import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import re
import logging



try:
    all_medicine_list_df = pd.read_excel('AllMedicineList.xlsx')
except FileNotFoundError:
    all_medicine_list_df = pd.DataFrame(columns=["Medicine Name", "Brand Name", "No. of Units", "Price"])
# Initialize dataframes from Excel files
try:
    inventory_homepage_df = pd.read_excel('inventory_homepage.xlsx')
except FileNotFoundError:
    inventory_homepage_df = pd.DataFrame(columns=[
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name",
        "Date of Purchase", "Date of Expiry", "Quantity Purchased",
        "Quantity Available", "No. of Units", "Price"
    ])

try:
    inventory_records_df = pd.read_excel('inventory_records.xlsx')
except FileNotFoundError:
    inventory_records_df = pd.DataFrame(columns=[
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name",
        "Date of Purchase", "Date of Expiry", "Quantity Purchased",
        "No. of Units", "Price"
    ])

try:
    sales_df = pd.read_excel('sales.xlsx')
except FileNotFoundError:
    sales_df = pd.DataFrame(columns=[
        "Medicine Name", "Brand Name", "Batch Number", "Date of Expiry",
        "Date of Sale", "Quantity Sold", "No. of Units", "Price"
    ])

try:
    returns_adjusted_df = pd.read_excel('returns_adjusted.xlsx')
except FileNotFoundError:
    returns_adjusted_df = pd.DataFrame(columns=[
        "Batch Number", "No. of Units", "Status"
    ])

try:
    permanent_returnsadjusted_df = pd.read_excel('permanent_returnsadjusted.xlsx')
except FileNotFoundError:
    permanent_returnsadjusted_df = pd.DataFrame(columns=[
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name",
        "Date of Expiry", "No. of Units", "Status"
    ])
def load_item_list():
    # Clear the existing items in the Treeview
    for item in inventory_tree.get_children():
        inventory_tree.delete(item)

    # Get the current date for comparison
    current_date = datetime.now()

    # Populate the Treeview with items from your inventory dataframe
    for index, row in inventory_homepage_df.iterrows():
        # Extract necessary data
        date_of_expiry = pd.to_datetime(row['Date of Expiry'])
        quantity_available = row['Quantity Available']

        # Determine the tag based on conditions
        tags = ()
        if (date_of_expiry - current_date).days <= 90:  # About to expire within 90 days
            tags = ('about_to_expire',)
        elif quantity_available < 4:  # Short quantity
            tags = ('short_quantity',)

        # Insert the item into the Treeview with the appropriate tag
        inventory_tree.insert("", "end", values=row.tolist(), tags=tags)

    # Apply conditional formatting
    inventory_tree.tag_configure('about_to_expire', background='red')
    inventory_tree.tag_configure('short_quantity', background='yellow')

def save_to_excel():
    inventory_homepage_df.to_excel('inventory_homepage.xlsx', index=False)
    inventory_records_df.to_excel('inventory_records.xlsx', index=False)
    sales_df.to_excel('sales.xlsx', index=False)
    returns_adjusted_df.to_excel('returns_adjusted.xlsx', index=False)
    permanent_returnsadjusted_df.to_excel('permanent_returnsadjusted.xlsx', index=False)
def load_or_create_temp_sales(filename="TemporarySales.xlsx"):
    if os.path.exists(filename):
        # Load the existing file
        return pd.read_excel(filename)
    else:
        # Create a new DataFrame with the necessary columns, including 'No. of Units'
        temp_sales_df = pd.DataFrame(columns=["Medicine Name", "Brand Name", "Batch Number",
                                              "Sell Quantity", "Sell Loose", "Price", "No. of Units"])
        temp_sales_df.to_excel(filename, index=False)
        return temp_sales_df

def update_temp_sales(medicine_name, brand_name, batch_number, sell_quantity, sell_loose, price, no_of_units):
    temp_sales_df = pd.read_excel('TemporarySales.xlsx')

    # Check if the entry already exists, if so, update it
    existing_entry = temp_sales_df[
        (temp_sales_df["Medicine Name"] == medicine_name) &
        (temp_sales_df["Brand Name"] == brand_name) &
        (temp_sales_df["Batch Number"] == batch_number)
    ]

    if not existing_entry.empty:
        index = existing_entry.index[0]
        temp_sales_df.at[index, 'Sell Quantity'] = sell_quantity
        temp_sales_df.at[index, 'Sell Loose'] = sell_loose
        temp_sales_df.at[index, 'Price'] = price
        temp_sales_df.at[index, 'No. of Units'] = no_of_units  # Ensure No. of Units is updated
    else:
        new_entry = pd.DataFrame([{
            "Medicine Name": medicine_name,
            "Brand Name": brand_name,
            "Batch Number": batch_number,
            "Sell Quantity": sell_quantity,
            "Sell Loose": sell_loose,
            "Price": price,
            "No. of Units": no_of_units  # Add No. of Units for the new entry
        }])
        temp_sales_df = pd.concat([temp_sales_df, new_entry], ignore_index=True)

    temp_sales_df.to_excel('TemporarySales.xlsx', index=False)

def clear_temp_sales():
    # Create an empty DataFrame with the required columns
    temp_sales_df = pd.DataFrame(columns=["Medicine Name", "Brand Name", "Batch Number",
                                          "Sell Quantity", "Sell Loose", "Price", "No. of Units"])
    # Overwrite the TemporarySales.xlsx file with the empty DataFrame
    temp_sales_df.to_excel('TemporarySales.xlsx', index=False)

def parse_date(date_str):
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError("Invalid date format. Please use DD-MM-YYYY or similar format.")

def validate_date(entry):
    date_text = entry.get()
    try:
        valid_date = datetime.strptime(date_text, '%d-%m-%Y')
        return valid_date
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter a valid date in DD-MM-YYYY format.")
        return None

def open_new_entry_window():
    global entries

    new_entry_window = tk.Toplevel(root)
    new_entry_window.title("Add New Entry")
    new_entry_window.geometry("530x380")  # Increased size to accommodate all fields

    # Create labels and entry fields
    labels = ["Medicine Name", "Brand Name", "Price", "No. of Units", "Batch Number",
              "Supplier Name", "Date of Purchase (DD-MM-YYYY)", "Date of Expiry (DD-MM-YYYY)",
              "Quantity Purchased", "Quantity Available"]
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(new_entry_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky='w')
        entry = tk.Entry(new_entry_window, width=50)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
        entries[label] = entry

    # Suggestion Box for Medicine Name
    suggestion_list = tk.Listbox(new_entry_window, height=5)

    # Set custom width (in pixels or character width)
    custom_suggestion_box_width = 50  # Example: 50 characters wide
    suggestion_list.configure(width=custom_suggestion_box_width)

    def update_suggestions(event):
        input_text = entries["Medicine Name"].get().lower()
        if input_text:
            # Match names that start with the input text
            matches = all_medicine_list_df[
                all_medicine_list_df['Medicine Name'].str.lower().str.startswith(input_text)
            ]

            suggestion_list.delete(0, tk.END)

            for match in matches['Medicine Name']:
                suggestion_list.insert(tk.END, match)

            # If no matches start with the input, fall back to contains
            if suggestion_list.size() == 0:
                matches = all_medicine_list_df[
                    all_medicine_list_df['Medicine Name'].str.lower().str.contains(input_text)
                ]
                for match in matches['Medicine Name']:
                    suggestion_list.insert(tk.END, match)

            suggestion_list.place(x=entries["Medicine Name"].winfo_x(),
                                  y=entries["Medicine Name"].winfo_y() + entries["Medicine Name"].winfo_height())
            suggestion_list.lift()
        else:
            suggestion_list.place_forget()

    def select_suggestion():
        if suggestion_list.curselection():
            selected_medicine = suggestion_list.get(suggestion_list.curselection())
            result = all_medicine_list_df[all_medicine_list_df['Medicine Name'] == selected_medicine]
            if not result.empty:
                # Auto-fill basic details from the medicine list
                entries["Medicine Name"].delete(0, tk.END)
                entries["Medicine Name"].insert(0, result.iloc[0]['Medicine Name'])
                entries["Brand Name"].delete(0, tk.END)
                entries["Brand Name"].insert(0, result.iloc[0]['Brand Name'])
                entries["Price"].delete(0, tk.END)
                entries["Price"].insert(0, result.iloc[0]['Price'])
                entries["No. of Units"].delete(0, tk.END)
                entries["No. of Units"].insert(0, result.iloc[0]['No. of Units'])

                # Auto-fill additional details from inventory records and homepage if available
                matching_row_records = inventory_records_df[
                    inventory_records_df['Medicine Name'].str.lower() == selected_medicine.lower()]

                matching_row_homepage = inventory_homepage_df[
                    inventory_homepage_df['Medicine Name'].str.lower() == selected_medicine.lower()]

                if not matching_row_records.empty:
                    # Autofill the most recent Batch Number and Supplier Name from records
                    entries["Batch Number"].delete(0, tk.END)
                    entries["Batch Number"].insert(0, matching_row_records.iloc[-1]['Batch Number'])

                    entries["Supplier Name"].delete(0, tk.END)
                    entries["Supplier Name"].insert(0, matching_row_records.iloc[-1]['Supplier Name'])

                if not matching_row_homepage.empty:
                    # Autofill the most recent Quantity Available from homepage
                    entries["Quantity Available"].delete(0, tk.END)
                    entries["Quantity Available"].insert(0, matching_row_homepage.iloc[-1]['Quantity Available'])

                # Auto-fill Date of Purchase as today's date
                today = datetime.now().strftime('%d-%m-%Y')
                entries["Date of Purchase (DD-MM-YYYY)"].delete(0, tk.END)
                entries["Date of Purchase (DD-MM-YYYY)"].insert(0, today)

                # Auto-fill Date of Expiry as one year later
                one_year_later = (datetime.now() + timedelta(days=365)).strftime('%d-%m-%Y')
                entries["Date of Expiry (DD-MM-YYYY)"].delete(0, tk.END)
                entries["Date of Expiry (DD-MM-YYYY)"].insert(0, one_year_later)

            suggestion_list.place_forget()

    # Bind the suggestion selection event to the select_suggestion function
    suggestion_list.bind("<<ListboxSelect>>", lambda event: select_suggestion())

    # Bind the update_suggestions function to the Medicine Name entry field
    entries["Medicine Name"].bind("<KeyRelease>", update_suggestions)
    new_entry_window.bind("<Button-1>", lambda event: suggestion_list.place_forget())

    button_add = tk.Button(new_entry_window, text="Add Medicine", command=add_medicine)
    button_add.grid(row=len(labels), column=0, columnspan=2, pady=10)

    new_entry_window.bind('<Return>', lambda event: add_medicine())


    def close_suggestion_box(event):
        suggestion_list.place_forget()

    # Bind the update_suggestions function to the Medicine Name entry field
    entries["Medicine Name"].bind("<KeyRelease>", update_suggestions)
    new_entry_window.bind("<Button-1>", close_suggestion_box)
    for entry in entries.values():
        entry.bind("<FocusIn>", close_suggestion_box)

    button_add = tk.Button(new_entry_window, text="Add Medicine", command=add_medicine)
    button_add.grid(row=len(labels), column=0, columnspan=2, pady=10)

    new_entry_window.bind('<Return>', lambda event: add_medicine())


def add_medicine():
    global inventory_homepage_df, inventory_records_df

    purchase_date = validate_date(entries["Date of Purchase (DD-MM-YYYY)"])
    expiry_date = validate_date(entries["Date of Expiry (DD-MM-YYYY)"])
    if purchase_date and expiry_date:
        # Proceed with adding the medicine to inventory
        print("Purchase Date:", purchase_date)
        print("Expiry Date:", expiry_date)

    # Retrieve all the input details from the entries dictionary
    medicine_name = entries["Medicine Name"].get()
    brand_name = entries["Brand Name"].get()
    batch_number = str(entries["Batch Number"].get())
    supplier_name = entries["Supplier Name"].get()
    purchase_date = entries["Date of Purchase (DD-MM-YYYY)"].get()
    expiry_date = entries["Date of Expiry (DD-MM-YYYY)"].get()
    quantity_available = float(entries["Quantity Available"].get()) if entries["Quantity Available"].get() else 0.0
    quantity_purchased = int(entries["Quantity Purchased"].get())
    no_of_units = int(entries["No. of Units"].get())
    price = float(entries["Price"].get())

    # Ensure all fields are properly validated
    if not all([medicine_name, brand_name, batch_number, supplier_name, purchase_date, expiry_date,
                quantity_purchased, no_of_units, price]):
        messagebox.showerror("Error", "All fields must be filled out.")
        return

    # Convert dates to a consistent format
    try:
        purchase_date = datetime.now()  # This captures both the current date and time
        expiry_date = pd.to_datetime(expiry_date, format='%d-%m-%Y')  # Date only for expiry
    except ValueError as e:
        messagebox.showerror("Error", str(e))
        return

    # Check if the medicine with the same batch number and supplier already exists in the inventory
    existing_row = inventory_homepage_df[
        (inventory_homepage_df['Medicine Name'] == medicine_name) &
        (inventory_homepage_df["Batch Number"] == batch_number) &
        (inventory_homepage_df["Supplier Name"] == supplier_name)
    ]

    if not existing_row.empty:
        # Update the existing entry by updating the available quantity and replacing the purchased quantity
        existing_index = existing_row.index[0]
        inventory_homepage_df.at[existing_index, 'Quantity Available'] += quantity_purchased  # Add only the purchased quantity to available stock
        inventory_homepage_df.at[existing_index, 'Quantity Purchased'] = quantity_purchased  # Replace with the new purchased quantity
        inventory_homepage_df.at[existing_index, 'No. of Units'] = no_of_units  # Update No. of Units if necessary

        # Also update the records DataFrame with a new entry for the purchased quantity
        new_entry_records = pd.DataFrame([{
            "Medicine Name": medicine_name,
            "Brand Name": brand_name,
            "Batch Number": batch_number,
            "Supplier Name": supplier_name,
            "Date of Purchase": purchase_date,
            "Date of Expiry": expiry_date,
            "Quantity Purchased": quantity_purchased,  # Reflect only the latest purchase quantity
            "No. of Units": no_of_units,
            "Price": price
        }])
        inventory_records_df = pd.concat([inventory_records_df, new_entry_records], ignore_index=True)
    else:
        # If the batch number or supplier is different, add a new entry to the inventory and to the records
        new_entry_homepage = pd.DataFrame([{
            "Medicine Name": medicine_name,
            "Brand Name": brand_name,
            "Batch Number": batch_number,
            "Supplier Name": supplier_name,
            "Date of Purchase": purchase_date,
            "Date of Expiry": expiry_date,
            "Quantity Purchased": quantity_purchased,
            "Quantity Available": quantity_available + quantity_purchased,  # Properly add the purchased quantity to available quantity
            "No. of Units": no_of_units,
            "Price": price
        }])
        inventory_homepage_df = pd.concat([inventory_homepage_df, new_entry_homepage], ignore_index=True)

        # Also record this new entry in the inventory records
        inventory_records_df = pd.concat([inventory_records_df, new_entry_homepage], ignore_index=True)

    # Sort the DataFrame by 'Medicine Name' to maintain alphabetical order
    inventory_homepage_df = inventory_homepage_df.sort_values(by='Medicine Name', ascending=True)
    inventory_records_df = inventory_records_df.sort_values(by='Medicine Name', ascending=True)

    # Save both DataFrames to Excel
    save_to_excel()

    update_treeview()
    clear_entries()
    messagebox.showinfo("Success", f"{medicine_name} has been added to the inventory.")


def clear_entries():
    for entry in entries.values():
        entry.delete(0, tk.END)
def update_treeview():
    global inventory_homepage_df

    # Clear the current contents of the treeview
    for item in inventory_tree.get_children():
        inventory_tree.delete(item)

    # Sort the dataframe by 'Medicine Name' column in ascending order
    sorted_df = inventory_homepage_df.sort_values(by='Medicine Name', ascending=True)

    # Get the current date for comparison
    current_date = datetime.now()

    # Populate the treeview with the data from the sorted dataframe
    for index, row in inventory_homepage_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Supplier Name"],
            row["Date of Purchase"],
            row["Date of Expiry"],
            row["Quantity Purchased"],
            row["Quantity Available"],
            row["No. of Units"],
            row["Price"]
        ]
        tags = ()

        # Convert the 'Date of Expiry' to datetime format within your function
        row['Date of Expiry'] = pd.to_datetime(row['Date of Expiry'], errors='coerce')

        # Check for items that are about to expire (within 90 days)
        if (row['Date of Expiry'] - current_date).days <= 90:
            tags = ('about_to_expire',)
        # Check for items with low stock (Quantity Available < 4)
        elif row['Quantity Available'] < 4:
            tags = ('short_quantity',)

        inventory_tree.insert("", "end", values=values, tags=tags)

    # Apply conditional formatting for tags
    inventory_tree.tag_configure('about_to_expire', background='red')
    inventory_tree.tag_configure('short_quantity', background='yellow')


def check_expiry():
    expiry_window = tk.Toplevel(root)
    expiry_window.title("Expiry Check")
    expiry_window.geometry("1100x450")  # Adjusted for additional UI elements

    frame_expiry = tk.Frame(expiry_window)
    frame_expiry.pack(pady=10, fill=tk.BOTH, expand=True)

    expiry_scroll = tk.Scrollbar(frame_expiry)
    expiry_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns = ("Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Purchase", "Date of Expiry",
               "Quantity Purchased", "Quantity Available", "No. of Units", "Price")

    expiry_tree = ttk.Treeview(frame_expiry, columns=columns, show="headings", yscrollcommand=expiry_scroll.set)
    expiry_tree.pack(fill=tk.BOTH, expand=True)

    expiry_scroll.config(command=expiry_tree.yview)

    column_widths = {
        "Medicine Name": 150,
        "Brand Name": 100,
        "Batch Number": 100,
        "Supplier Name": 120,
        "Date of Purchase": 100,
        "Date of Expiry": 100,
        "Quantity Purchased": 120,
        "Quantity Available": 120,
        "No. of Units": 100,
        "Price": 80
    }

    for col in columns:
        expiry_tree.heading(col, text=col)
        expiry_tree.column(col, width=column_widths[col], minwidth=50, stretch=True)

    current_date = datetime.now()

    # Convert 'Date of Expiry' column to datetime if it's not already
    inventory_homepage_df['Date of Expiry'] = pd.to_datetime(inventory_homepage_df['Date of Expiry'], errors='coerce')

    expiry_list = []  # List to store expiry rows

    for index, row in inventory_homepage_df.iterrows():
        # Skip rows where 'Date of Expiry' could not be converted
        if pd.isna(row['Date of Expiry']):
            continue

        days_to_expiry = (row['Date of Expiry'] - current_date).days
        if days_to_expiry <= 90:
            expiry_tree.insert("", "end", values=row.tolist())
            expiry_list.append(row.tolist())  # Add to expiry list

    # Function to filter by selected month, year and download expiry list
    def download_expiry_list():
        if not expiry_list:
            tk.messagebox.showinfo("No Data", "No expiry data to download.")
            return

        # Get selected month and year
        selected_month = month_combobox.get()
        selected_year = year_combobox.get()

        # Filter the list
        filtered_list = expiry_list
        if selected_month != "All":
            month_number = datetime.strptime(selected_month, "%B").month
            filtered_list = [
                row for row in filtered_list
                if pd.to_datetime(row[columns.index("Date of Expiry")]).month == month_number
            ]

        if selected_year != "All":
            filtered_list = [
                row for row in filtered_list
                if pd.to_datetime(row[columns.index("Date of Expiry")]).year == int(selected_year)
            ]

        if not filtered_list:
            tk.messagebox.showinfo("No Data", f"No expiry data for {selected_month} {selected_year}.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            title="Save Expiry List"
        )

        if file_path:
            expiry_df = pd.DataFrame(filtered_list, columns=columns)
            if file_path.endswith('.xlsx'):
                expiry_df.to_excel(file_path, index=False)
            elif file_path.endswith('.csv'):
                expiry_df.to_csv(file_path, index=False)
            tk.messagebox.showinfo("Success", f"Expiry list saved to {file_path}!")

    # Dropdown to select month
    month_label = tk.Label(expiry_window, text="Select Month:")
    month_label.pack(pady=5)

    months = ["All"] + [datetime(1900, i, 1).strftime('%B') for i in range(1, 13)]
    month_combobox = ttk.Combobox(expiry_window, values=months, state="readonly")
    month_combobox.current(0)  # Default to "All"
    month_combobox.pack(pady=5)

    # Dropdown to select year
    year_label = tk.Label(expiry_window, text="Select Year:")
    year_label.pack(pady=5)

    unique_years = sorted({d.year for d in inventory_homepage_df['Date of Expiry'].dropna()})
    years = ["All"] + [str(year) for year in unique_years]
    year_combobox = ttk.Combobox(expiry_window, values=years, state="readonly")
    year_combobox.current(0)  # Default to "All"
    year_combobox.pack(pady=5)

    # Add a "Download Expiry List" button
    download_button = tk.Button(expiry_window, text="Download Expiry List", command=download_expiry_list)
    download_button.pack(pady=10)

    print("Expiry check complete.")  # Optional debug message



def shortlist_items():
    shortlist_window = tk.Toplevel(root)
    shortlist_window.title("Short List Items")

    # Set custom window size (width x height)
    shortlist_window.geometry("1100x300")  # Adjust as needed

    frame_shortlist = tk.Frame(shortlist_window)
    frame_shortlist.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    shortlist_scroll = tk.Scrollbar(frame_shortlist)
    shortlist_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns = ("Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Purchase", "Date of Expiry",
               "Quantity Purchased", "Quantity Available", "No. of Units", "Price")

    shortlist_tree = ttk.Treeview(frame_shortlist, columns=columns, show="headings", yscrollcommand=shortlist_scroll.set)
    shortlist_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    shortlist_scroll.config(command=shortlist_tree.yview)

    # Set custom column widths
    column_widths = {
        "Medicine Name": 150,
        "Brand Name": 100,
        "Batch Number": 100,
        "Supplier Name": 120,
        "Date of Purchase": 100,
        "Date of Expiry": 100,
        "Quantity Purchased": 120,
        "Quantity Available": 120,
        "No. of Units": 100,
        "Price": 80
    }

    for col in columns:
        shortlist_tree.heading(col, text=col)
        shortlist_tree.column(col, width=column_widths[col], minwidth=50, stretch=True)

    shortlist_items_list = []  # List to store shortlisted rows

    # Add rows to the Treeview and shortlist list
    for index, row in inventory_homepage_df.iterrows():
        if row['Quantity Available'] < 4:
            shortlist_tree.insert("", "end", values=row.tolist())
            shortlist_items_list.append(row.tolist())

    # Function to download the shortlist items
    def download_shortlist_items():
        if not shortlist_items_list:
            tk.messagebox.showinfo("No Data", "No shortlist items to download.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            title="Save Shortlist Items"
        )

        if file_path:
            shortlist_df = pd.DataFrame(shortlist_items_list, columns=columns)
            if file_path.endswith('.xlsx'):
                shortlist_df.to_excel(file_path, index=False)
            elif file_path.endswith('.csv'):
                shortlist_df.to_csv(file_path, index=False)
            tk.messagebox.showinfo("Success", f"Shortlist items saved to {file_path}!")

    # Add a "Download Shortlist Items" button
    download_button = tk.Button(shortlist_window, text="Download Shortlist Items", command=download_shortlist_items)
    download_button.pack(pady=10)

    print("Shortlist items check complete.")  # Optional debug message

def search_item():
    search_term = entry_search.get().lower()

    for item in inventory_tree.get_children():
        inventory_tree.delete(item)

    current_date = datetime.now()  # Get the current date

    for index, row in inventory_homepage_df.iterrows():
        # Check if the search term is in any value of the row
        if any(search_term in str(value).lower() for value in row):
            values = row.tolist()
            tags = ()

            # Convert 'Date of Expiry' to datetime format, if it's not already
            try:
                expiry_date = pd.to_datetime(row['Date of Expiry'], errors='coerce')  # Convert to datetime
                if pd.isna(expiry_date):  # If conversion failed, skip the row
                    continue

                # Check if the item is about to expire (within 90 days)
                if (expiry_date - current_date).days <= 90:
                    tags = ('about_to_expire',)

            except Exception as e:
                print(f"Error processing row {index}: {e}")
                continue

            # Check for low stock (Quantity Available < 4)
            if row['Quantity Available'] < 4:
                tags = ('short_quantity',)

            # Insert the row into the inventory tree
            inventory_tree.insert("", "end", values=values, tags=tags)

    # Configure the tags for custom formatting
    inventory_tree.tag_configure('about_to_expire', background='red')
    inventory_tree.tag_configure('short_quantity', background='yellow')


def search_records(event=None):
    search_term = entry_search_records.get().lower()

    for item in records_tree.get_children():
        records_tree.delete(item)

    for index, row in inventory_records_df.iterrows():
        if any(search_term in str(value).lower() for value in row):
            values = row.tolist()
            records_tree.insert("", "end", values=values)


def open_sell_multiple_window():
    sell_multiple_window = tk.Toplevel(root)
    sell_multiple_window.title("Sell Multiple Items")

    temp_sales_df = load_or_create_temp_sales()

    frame_search = tk.Frame(sell_multiple_window)
    frame_search.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

    entry_search_sell_multiple = tk.Entry(frame_search, width=30)
    entry_search_sell_multiple.pack(side=tk.LEFT, padx=5)
    entry_search_sell_multiple.bind('<Return>', lambda event: search_sell_multiple(entry_search_sell_multiple.get()))

    button_search_sell_multiple = tk.Button(frame_search, text="Search",
                                            command=lambda: search_sell_multiple(entry_search_sell_multiple.get()))
    button_search_sell_multiple.pack(side=tk.LEFT, padx=5)

    # Main frame to hold both Treeview and the Listbox for selected items
    main_frame = tk.Frame(sell_multiple_window)
    main_frame.pack(pady=10, fill=tk.BOTH, expand=True)

    # Frame for Treeview (Left side)
    frame_items = tk.Frame(main_frame)
    frame_items.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Frame for selected items display (Right side)
    frame_selected_items = tk.Frame(main_frame)
    frame_selected_items.pack(side=tk.RIGHT, padx=10, fill=tk.Y)

    # Label for selected items
    label_selected_items = tk.Label(frame_selected_items, text="Selected Items")
    label_selected_items.pack(pady=5)

    # Listbox to display selected items
    listbox_selected_items = tk.Listbox(frame_selected_items, height=20, width=65)
    listbox_selected_items.pack(pady=5)

    # Function to clear the listbox and the TemporarySales.xlsx file
    def clear_listbox_and_temp_sales():
        # Clear the listbox
        listbox_selected_items.delete(0, tk.END)

        # Clear the TemporarySales.xlsx file
        clear_temp_sales()

        # Optionally, show a confirmation message
        messagebox.showinfo("Success", "Listbox and temporary sales file have been cleared.")


    # Create a vertical scrollbar for the Treeview
    sell_multiple_scroll = tk.Scrollbar(frame_items)
    sell_multiple_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns_sell_multiple = ["Medicine Name", "Brand Name", "Batch Number", "Date of Expiry", "No. of Units",
                             "Price", "Available Quantity", "Sell Quantity", "Sell Loose"]

    sell_multiple_tree = ttk.Treeview(frame_items, columns=columns_sell_multiple, show="headings",
                                      yscrollcommand=sell_multiple_scroll.set)
    sell_multiple_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    sell_multiple_scroll.config(command=sell_multiple_tree.yview)

    for col in columns_sell_multiple:
        sell_multiple_tree.heading(col, text=col)
        sell_multiple_tree.column(col, width=100, minwidth=50, stretch=True)

    for index, row in inventory_homepage_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Date of Expiry"],
            row["No. of Units"],
            row["Price"],
            row["Quantity Available"],
            0,  # Initial Sell Quantity
            0  # Initial Sell Loose
        ]
        sell_multiple_tree.insert("", "end", values=values)

    def search_sell_multiple(search_term):
        # Load the temporary sales data
        temp_sales_df = load_or_create_temp_sales()

        # Clear the Treeview as usual
        for item in sell_multiple_tree.get_children():
            sell_multiple_tree.delete(item)

        # Repopulate the Treeview with the search results
        for index, row in inventory_homepage_df.iterrows():
            if any(search_term.lower() in str(value).lower() for value in row):
                values = [
                    row["Medicine Name"],
                    row["Brand Name"],
                    row["Batch Number"],
                    row["Date of Expiry"],
                    row["No. of Units"],
                    row["Price"],
                    row["Quantity Available"],
                    0,  # Initial Sell Quantity
                    0  # Initial Sell Loose
                ]

                # Check if this item has saved quantities in the temp sales file
                temp_entry = temp_sales_df[
                    (temp_sales_df["Medicine Name"] == row["Medicine Name"]) &
                    (temp_sales_df["Brand Name"] == row["Brand Name"]) &
                    (temp_sales_df["Batch Number"] == row["Batch Number"])
                    ]

                if not temp_entry.empty:
                    values[7] = int(temp_entry["Sell Quantity"].values[0])  # Update Sell Quantity
                    values[8] = int(temp_entry["Sell Loose"].values[0])  # Update Sell Loose

                sell_multiple_tree.insert("", "end", values=values)

    def update_selected_items_listbox():
        """Update the listbox with selected items, including their name, quantity, and price."""
        # Clear the listbox
        listbox_selected_items.delete(0, tk.END)

        # Load the temporary sales data
        temp_sales_df = load_or_create_temp_sales()

        # Populate the listbox with the selected items from the temp sales data
        for index, row in temp_sales_df.iterrows():
            name = row["Medicine Name"]
            sell_quantity = row["Sell Quantity"]
            sell_loose = row["Sell Loose"]

            # Use .get() to safely access 'No. of Units' with a default value if the column is missing
            no_of_units = row.get("No. of Units", 1)  # Default to 1 if 'No. of Units' is not available

            # Get the price for the full item (ensure it's properly fetched)
            price = row.get("Price", 0)

            # Handle the case where no_of_units is zero to prevent ZeroDivisionError
            if no_of_units == 0:
                price_per_unit = 0  # Set price_per_unit to 0 to avoid division by zero
            else:
                price_per_unit = price / no_of_units

            # Calculate the loose price (based on sell_loose)
            loose_price = sell_loose * price_per_unit

            # Calculate the total price (full quantity price + loose quantity price)
            full_quantity_price = sell_quantity * price
            total_price = full_quantity_price + loose_price  # Sum of both full and loose quantities

            # Format the output string with padding to align columns
            formatted_text = f"{name:<20} | Full Qty: {sell_quantity:<5} | Loose Qty: {sell_loose:<5} | Total Price: {total_price:.2f}"
            listbox_selected_items.insert(tk.END, formatted_text)

            # Debugging print (optional, remove if unnecessary)
            print(
                f"Item: {name}, Full Qty: {sell_quantity}, Loose Qty: {sell_loose}, Price per Unit: {price_per_unit}, Total Price: {total_price}")

    def on_double_click(event):
        selected_item = sell_multiple_tree.selection()

        if selected_item:
            item = sell_multiple_tree.item(selected_item[0])  # Use the first selected item
            values = item['values']

            def save_quantity_and_loose():
                try:
                    new_quantity = int(entry_quantity.get())
                    new_loose = int(entry_loose.get())

                    # Update the Treeview with the selected values
                    sell_multiple_tree.set(selected_item[0], column='Sell Quantity', value=new_quantity)
                    sell_multiple_tree.set(selected_item[0], column='Sell Loose', value=new_loose)

                    # Extract the price and number of units from the current row values
                    price = float(values[5])  # Assuming 'values[5]' corresponds to the price column in your Treeview
                    no_of_units = int(values[8])  # Assuming 'values[8]' corresponds to the No. of Units

                    # Update the temporary sales file with the price and no_of_units as well
                    update_temp_sales(values[0], values[1], values[2], new_quantity, new_loose, price, no_of_units)

                    # Update the selected items listbox
                    update_selected_items_listbox()

                    popup.destroy()
                except ValueError:
                    messagebox.showerror("Error", "Please enter valid numbers for quantity and loose units.")

            # Popup for entering quantities
            popup = tk.Toplevel(sell_multiple_window)
            popup.title("Set Sell Quantity and Loose Units")

            tk.Label(popup, text="Enter quantity to sell:").pack(pady=5)
            entry_quantity = tk.Entry(popup)
            entry_quantity.pack(pady=5)
            entry_quantity.insert(0, values[7])  # Sell Quantity

            tk.Label(popup, text="Enter loose units to sell:").pack(pady=5)
            entry_loose = tk.Entry(popup)
            entry_loose.pack(pady=5)
            entry_loose.insert(0, values[8])  # Sell Loose

            button_save = tk.Button(popup, text="Save", command=save_quantity_and_loose)
            button_save.pack(pady=10)

    def clean_string(input_string):
        # Ensure that the input is a string before applying regex
        if not isinstance(input_string, str):
            input_string = str(input_string)
        return re.sub(r'\s+', ' ', input_string).strip()

    # Function to display the custom popup with "OK" and "Sell Anyways" options
    def ask_sell_anyways(item_name, batch_number):
        # Create a simple window for the custom dialog
        response = messagebox.askyesno(
            "Item Not Found",
            f"Price information for {item_name} with batch {batch_number} not found.\nDo you want to sell anyway?"
        )
        return response

    def sell_selected_items():
        """
        Handles the sale process for items selected in the temporary sales DataFrame.
        Ensures data integrity, updates inventory, and records sales.
        """
        global sales_df
        total_price = 0
        sold_items = []
        confirm_message = "Do you want to sell the following items?\n\n"

        # Load temporary sales data
        temp_sales_df = load_or_create_temp_sales()

        for _, row in temp_sales_df.iterrows():
            # Clean and prepare necessary fields
            name = clean_string(row["Medicine Name"])
            brand = clean_string(row["Brand Name"])
            batch = clean_string(str(row["Batch Number"]))
            sell_quantity = int(row["Sell Quantity"])
            sell_loose = int(row["Sell Loose"])

            # Debugging logs
            logging.debug(f"Processing item: {name}, {brand}, {batch}")

            # Retrieve matching inventory
            matching_inventory = inventory_homepage_df[
                (inventory_homepage_df["Medicine Name"].apply(clean_string).str.lower() == name.lower()) &
                (inventory_homepage_df["Brand Name"].apply(clean_string).str.lower() == brand.lower()) &
                (inventory_homepage_df["Batch Number"].astype(str).apply(clean_string).str.lower() == batch.lower())
                ]

            if matching_inventory.empty:
                # Handle missing inventory case
                logging.warning(f"Inventory not found for {name} with batch {batch}.")
                sell_anyways = ask_sell_anyways(name, batch)
                if not sell_anyways:
                    messagebox.showinfo("Sale Cancelled", f"Sale of {name} with batch {batch} cancelled.")
                    continue
                price, available_quantity, no_of_units = 0, sell_quantity, 1  # Defaults for selling without inventory
            else:
                # Extract inventory details
                inventory_row = matching_inventory.iloc[0]
                price = inventory_row.get("Price", 0)
                if pd.isna(price) or price <= 0:
                    logging.warning(f"Invalid price for {name} with batch {batch}. Setting price to 0.")
                    price = 0
                available_quantity = float(inventory_row["Quantity Available"])
                no_of_units = int(float(inventory_row["No. of Units"]))

            # Check stock availability
            if sell_quantity + sell_loose > available_quantity:
                messagebox.showerror("Error", f"Not enough stock for {name}.")
                continue

            # Calculate and deduct loose quantity
            qty_string, price_per_unit = "", 0
            if sell_loose > 0:
                if sell_loose > available_quantity * no_of_units:
                    messagebox.showerror("Error", f"Not enough loose units for {name}.")
                    continue
                units_sold_fraction = sell_loose / no_of_units
                inventory_homepage_df.loc[matching_inventory.index, "Quantity Available"] -= units_sold_fraction
                price_per_unit = price / no_of_units
                total_price += price_per_unit * sell_loose
                qty_string += f"{sell_loose} loose"

            # Deduct full quantity
            if sell_quantity > 0:
                inventory_homepage_df.loc[matching_inventory.index, "Quantity Available"] -= sell_quantity
                total_price += price * sell_quantity
                qty_string += f"{' and ' if qty_string else ''}{sell_quantity} full"

            # Prepare sale record
            sale_entry = pd.DataFrame([{
                "Medicine Name": name,
                "Brand Name": brand,
                "Batch Number": batch,
                "Date of Expiry": inventory_row.get("Date of Expiry"),
                "Date of Sale": datetime.now(),
                "Quantity Sold": sell_quantity,
                "Loose Quantity Sold": sell_loose,
                "No. of Units": no_of_units,
                "Price": price * sell_quantity + price_per_unit * sell_loose
            }])
            sales_df = pd.concat([sales_df, sale_entry], ignore_index=True)
            sold_items.append(f"{qty_string} units of {name} sold at {price if price > 0 else 'no price'}.")

            # Update confirmation message
            confirm_message += f"{name} ({qty_string}) @ {price:.2f} each\n"

        # Final confirmation and updates
        if sold_items:
            confirm_message += f"\nTotal Price: {total_price:.2f}"
            confirm = messagebox.askyesno("Confirm Sale", confirm_message)
            if confirm:
                update_treeview()
                save_to_excel()
                clear_temp_sales()
                messagebox.showinfo("Success", "Sale completed successfully.")
            else:
                messagebox.showinfo("Sale Cancelled", "Sale operation was cancelled.")
        else:
            messagebox.showinfo("No Items Sold", "No items were selected for sale.")

        if sold_items:
            sold_items_message = "\n".join(sold_items) + f"\nTotal Price = {total_price:.2f}"
            messagebox.showinfo("Sold", sold_items_message)

    # Frame to hold the listbox and buttons for the right side (selected items list and buttons)
    frame_right = tk.Frame(sell_multiple_window)
    frame_right.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH)


    # Frame to hold the buttons (Sell Selected Items and Clear Listbox) below the listboxes
    frame_buttons = tk.Frame(frame_right)
    frame_buttons.pack(pady=5)

    # Sell Selected Items button (colored green with white text)
    button_sell_selected = tk.Button(frame_buttons, text="Sell Selected Items", command=sell_selected_items, bg="green",
                                     fg="white")
    button_sell_selected.pack(side=tk.LEFT, padx=800)  # Adjust the spacing with padx

    # Clear Listbox button (colored green with white text)
    button_clear_listbox = tk.Button(frame_buttons, text="Clear Listbox", command=clear_listbox_and_temp_sales,
                                     bg="green", fg="white")
    button_clear_listbox.pack(side=tk.LEFT, padx=10)  # Adjust the spacing with padx

    sell_multiple_tree.bind('<Double-1>', on_double_click)
    sell_multiple_window.bind('<Control-f>', lambda event: entry_search_sell_multiple.focus_set())

    update_selected_items_listbox()  # Initialize the listbox with existing items

def clean_string(value):
    """
    Cleans and normalizes a string by trimming whitespace, converting to lowercase,
    and handling invalid (e.g., NaN) values.

    Args:
        value: The string or value to clean.

    Returns:
        str: The cleaned string, or an empty string if the input is invalid.
    """
    if pd.isnull(value):  # Handle NaN or None
        return ""
    return str(value).strip().lower()


def check_inventory_errors(inventory_file_path="inventory_homepage.xlsx"):
    import pandas as pd
    from datetime import datetime

    errors = []

    try:
        # Load the inventory file
        inventory_df = pd.read_excel(inventory_file_path)

        # Check for missing or invalid batch numbers
        if "Batch Number" not in inventory_df.columns or inventory_df["Batch Number"].isnull().any():
            errors.append("Some items are missing batch numbers.")

        # Check for duplicate batch numbers
        duplicate_batches = inventory_df["Batch Number"].duplicated(keep=False)
        if duplicate_batches.any():
            duplicate_batches_list = inventory_df[duplicate_batches]["Batch Number"].tolist()
            errors.append(f"Duplicate batch numbers found: {set(duplicate_batches_list)}")

        # Check for invalid or missing quantities
        if "Quantity Available" not in inventory_df.columns or inventory_df["Quantity Available"].isnull().any():
            errors.append("Some items are missing 'Quantity Available' values.")
        elif (inventory_df["Quantity Available"] < 0).any():
            errors.append("Negative quantities found in 'Quantity Available'.")

        # Check for invalid or missing prices
        if "Price" not in inventory_df.columns or inventory_df["Price"].isnull().any():
            errors.append("Some items are missing price information.")
        elif (inventory_df["Price"] < 0).any():
            errors.append("Negative prices found.")

        # Check for missing medicine or brand names
        if "Medicine Name" not in inventory_df.columns or inventory_df["Medicine Name"].isnull().any():
            errors.append("Some items are missing 'Medicine Name'.")
        if "Brand Name" not in inventory_df.columns or inventory_df["Brand Name"].isnull().any():
            errors.append("Some items are missing 'Brand Name'.")

        # Check for invalid or expired expiry dates
        if "Date of Expiry" in inventory_df.columns:
            expired_items = inventory_df[
                pd.to_datetime(inventory_df["Date of Expiry"], errors="coerce") < datetime.now()
            ]
            if not expired_items.empty:
                expired_batches = expired_items["Batch Number"].tolist()
                errors.append(f"Expired items found with batch numbers: {expired_batches}")
        else:
            errors.append("Column 'Date of Expiry' is missing in the inventory file.")

        # Check for mismatched entries (if any new data needs to be validated)
        # Example: Validate entries added from a temporary sales file
        temp_sales_df = load_or_create_temp_sales()
        for _, row in temp_sales_df.iterrows():
            name = clean_string(row["Medicine Name"])
            brand = clean_string(row["Brand Name"])
            batch = clean_string(str(row["Batch Number"]))

            matching_inventory = inventory_df[
                (inventory_df["Medicine Name"].apply(clean_string) == name) &
                (inventory_df["Brand Name"].apply(clean_string) == brand) &
                (inventory_df["Batch Number"].apply(clean_string) == batch)
            ]
            if matching_inventory.empty:
                errors.append(f"No matching item found for {name} (Batch: {batch}, Brand: {brand}).")

    except FileNotFoundError:
        errors.append(f"Inventory file '{inventory_file_path}' not found.")
    except Exception as e:
        errors.append(f"Unexpected error: {str(e)}")

    # Report the errors
    if errors:
        print("Inventory Check Report:")
        for error in errors:
            print(f"- {error}")
    else:
        print("No errors found in the inventory.")

    return errors

import pandas as pd

def scan_and_fix_inventory(file_path):
    """
    Scans the inventory file for issues and attempts to fix them. Reports unresolved issues.

    Args:
        file_path (str): Path to the inventory_homepage.xlsx file.

    Returns:
        str: Report of issues found and fixed.
    """
    try:
        # Load the inventory file
        inventory_df = pd.read_excel(file_path)

        # Initialize a report
        report = []

        # Check for missing or invalid prices
        if 'Price' in inventory_df.columns:
            invalid_prices = inventory_df[inventory_df['Price'].isna() | (inventory_df['Price'] <= 0)]
            if not invalid_prices.empty:
                report.append(f"Found {len(invalid_prices)} items with missing or invalid prices. Setting invalid prices to 1.0.")
                inventory_df.loc[invalid_prices.index, 'Price'] = 1.0  # Assign a default value of 1.0

        # Check for duplicate entries
        if all(col in inventory_df.columns for col in ['Medicine Name', 'Brand Name', 'Batch Number']):
            duplicates = inventory_df.duplicated(subset=['Medicine Name', 'Brand Name', 'Batch Number'], keep=False)
            if duplicates.any():
                report.append(f"Found {duplicates.sum()} duplicate entries. Keeping only the first occurrence.")
                inventory_df = inventory_df[~duplicates.duplicated(keep='first')]

        # Check for invalid quantities
        if 'Quantity Available' in inventory_df.columns:
            negative_quantities = inventory_df[inventory_df['Quantity Available'] < 0]
            if not negative_quantities.empty:
                report.append(f"Found {len(negative_quantities)} items with negative quantities. Setting them to 0.")
                inventory_df.loc[negative_quantities.index, 'Quantity Available'] = 0

        if 'No. of Units' in inventory_df.columns:
            invalid_units = inventory_df[inventory_df['No. of Units'] <= 0]
            if not invalid_units.empty:
                report.append(f"Found {len(invalid_units)} items with invalid 'No. of Units'. Setting them to 1.")
                inventory_df.loc[invalid_units.index, 'No. of Units'] = 1

        # Check for missing critical fields
        for field in ['Medicine Name', 'Brand Name', 'Batch Number']:
            if field in inventory_df.columns:
                missing_values = inventory_df[inventory_df[field].isna() | (inventory_df[field].astype(str).str.strip() == "")]
                if not missing_values.empty:
                    report.append(f"Found {len(missing_values)} items with missing {field}. These items were removed.")
                    inventory_df = inventory_df[~inventory_df.index.isin(missing_values.index)]

        # Normalize string fields
        for field in ['Medicine Name', 'Brand Name', 'Batch Number']:
            if field in inventory_df.columns:
                inventory_df[field] = inventory_df[field].astype(str).apply(clean_string)

        # Save the cleaned data back to the Excel file
        inventory_df.to_excel(file_path, index=False)
        report.append("Inventory file has been scanned and cleaned successfully.")

        return "\n".join(report)

    except Exception as e:
        return f"An error occurred while scanning the inventory: {e}"


def open_item_records_window():
    global records_tree, entry_search_records

    records_window = tk.Toplevel(root)
    records_window.title("Item Records")

    # Set custom window size (width x height)
    records_window.geometry("1100x250")  # Example size, adjust as needed

    frame_search_records = tk.Frame(records_window)
    frame_search_records.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

    entry_search_records = tk.Entry(frame_search_records, width=30)
    entry_search_records.pack(side=tk.LEFT, padx=5)
    entry_search_records.bind('<Return>', search_records)

    button_search_records = tk.Button(frame_search_records, text="Search", command=search_records)
    button_search_records.pack(side=tk.LEFT, padx=5)

    frame_records = tk.Frame(records_window)
    frame_records.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    records_scroll = tk.Scrollbar(frame_records)
    records_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    records_columns = [
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Purchase",
        "Date of Expiry", "Quantity Purchased", "No. of Units", "Price"
    ]

    records_tree = ttk.Treeview(frame_records, columns=records_columns, show="headings", yscrollcommand=records_scroll.set)
    records_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    records_scroll.config(command=records_tree.yview)

    # Set custom column widths
    column_widths = {
        "Medicine Name": 150,
        "Brand Name": 100,
        "Batch Number": 100,
        "Supplier Name": 120,
        "Date of Purchase": 120,
        "Date of Expiry": 120,
        "Quantity Purchased": 120,
        "No. of Units": 100,
        "Price": 80
    }

    for col in records_columns:
        records_tree.heading(col, text=col)
        records_tree.column(col, width=column_widths[col], minwidth=50, stretch=True)


    inventory_records_df['Date of Purchase'] = pd.to_datetime(inventory_records_df['Date of Purchase'], errors='coerce')
    sorted_records_df = inventory_records_df.sort_values(by="Date of Purchase", ascending=False)


    for index, row in sorted_records_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Supplier Name"],
            row["Date of Purchase"],
            row["Date of Expiry"],
            row["Quantity Purchased"],
            row["No. of Units"],
            row["Price"]
        ]
        records_tree.insert("", "end", values=values)


def open_sales_records_window():
    global sales_tree, entry_search_sales

    sales_window = tk.Toplevel(root)
    sales_window.title("Sales Records")

    frame_search_sales = tk.Frame(sales_window)
    frame_search_sales.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

    entry_search_sales = tk.Entry(frame_search_sales, width=30)
    entry_search_sales.pack(side=tk.LEFT, padx=5)
    entry_search_sales.bind('<Return>', search_sales_records)

    button_search_sales = tk.Button(frame_search_sales, text="Search", command=search_sales_records)
    button_search_sales.pack(side=tk.LEFT, padx=5)

    frame_sales = tk.Frame(sales_window)
    frame_sales.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    sales_scroll = tk.Scrollbar(frame_sales)
    sales_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    sales_columns = [
        "Medicine Name", "Brand Name", "Batch Number", "Date of Expiry",
        "Date of Sale", "Quantity Sold", "Loose Quantity Sold", "No. of Units", "Price"
    ]

    sales_tree = ttk.Treeview(frame_sales, columns=sales_columns, show="headings", yscrollcommand=sales_scroll.set)
    sales_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    sales_scroll.config(command=sales_tree.yview)

    for col in sales_columns:
        sales_tree.heading(col, text=col)
        sales_tree.column(col, width=100, minwidth=50, stretch=True)

    # Sort the DataFrame by "Date of Sale" in descending order
    sorted_sales_df = sales_df.sort_values(by="Date of Sale", ascending=False)

    for index, row in sorted_sales_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Date of Expiry"],
            row["Date of Sale"],
            row["Quantity Sold"],
            row.get("Loose Quantity Sold", 0),  # Ensure this is added
            row["No. of Units"],
            row["Price"]
        ]
        sales_tree.insert("", "end", values=values)


    def delete_selected_sale():
        selected_item = sales_tree.selection()
        if selected_item:
            confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this sale record?")
            if confirm:
                # Get the selected record's details
                item_values = sales_tree.item(selected_item)["values"]
                medicine_name = item_values[0]
                batch_number = item_values[2]
                quantity_sold = int(item_values[5])  # Convert to integer

                # Handle loose_quantity_sold properly, converting to int or defaulting to 0
                try:
                    loose_quantity_sold = int(float(item_values[6]))  # Convert to integer or float as needed
                except (ValueError, TypeError):
                    loose_quantity_sold = 0  # Default to 0 if conversion fails

                no_of_units = int(item_values[7])  # Convert to integer

                # Update inventory by adding the sold quantity and loose quantity back
                inventory_index = inventory_homepage_df[
                    (inventory_homepage_df["Medicine Name"] == medicine_name) &
                    (inventory_homepage_df["Batch Number"] == batch_number)
                    ].index

                if not inventory_index.empty:
                    inventory_index = inventory_index[0]
                    inventory_homepage_df.at[inventory_index, "Quantity Available"] += quantity_sold

                    # Add loose quantity back to the inventory
                    if loose_quantity_sold > 0:
                        # Convert loose quantity to a fractional value based on the units per full item
                        loose_quantity_fraction = loose_quantity_sold / no_of_units
                        inventory_homepage_df.at[
                            inventory_index, "Quantity Available"] += loose_quantity_fraction
                else:
                    messagebox.showerror("Error",
                                         f"No matching inventory record found for {medicine_name} with batch number {batch_number}.")
                    return

                # Remove the selected record from the sales DataFrame
                sales_df.drop(sales_tree.index(selected_item), inplace=True)
                sales_df.reset_index(drop=True, inplace=True)
                sales_tree.delete(selected_item)

                # Save the updated data
                save_to_excel()
                update_treeview()
                messagebox.showinfo("Deleted", "The selected sale record has been deleted and the inventory updated.")
            else:
                return
        else:
            messagebox.showerror("Error", "No record selected.")

    button_delete_sale = tk.Button(sales_window, text="Delete Record", command=delete_selected_sale)
    button_delete_sale.pack(pady=10)

    sales_window.bind('<Control-f>', lambda event: entry_search_sales.focus_set())

def search_sales_records(event=None):
    search_term = entry_search_sales.get().lower()

    for item in sales_tree.get_children():
        sales_tree.delete(item)

    for index, row in sales_df.iterrows():
        if any(search_term in str(value).lower() for value in row):
            values = row.tolist()
            sales_tree.insert("", "end", values=values)


def open_delete_edit_entry_window():
    global delete_edit_entry_window, delete_edit_tree, entry_search_delete_edit

    delete_edit_entry_window = tk.Toplevel(root)
    delete_edit_entry_window.title("Delete/Edit Entry")

    frame_search_delete_edit = tk.Frame(delete_edit_entry_window)
    frame_search_delete_edit.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

    entry_search_delete_edit = tk.Entry(frame_search_delete_edit, width=30)
    entry_search_delete_edit.pack(side=tk.LEFT, padx=5)
    entry_search_delete_edit.bind('<Return>', search_delete_edit_entry)

    button_search_delete_edit = tk.Button(frame_search_delete_edit, text="Search", command=search_delete_edit_entry)
    button_search_delete_edit.pack(side=tk.LEFT, padx=5)

    frame_delete_edit = tk.Frame(delete_edit_entry_window)
    frame_delete_edit.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    delete_edit_scroll = tk.Scrollbar(frame_delete_edit)
    delete_edit_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns = [
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Purchase",
        "Date of Expiry", "Quantity Purchased", "Quantity Available", "No. of Units", "Price"
    ]

    delete_edit_tree = ttk.Treeview(frame_delete_edit, columns=columns, show="headings", yscrollcommand=delete_edit_scroll.set)
    delete_edit_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    delete_edit_scroll.config(command=delete_edit_tree.yview)

    for col in columns:
        delete_edit_tree.heading(col, text=col)
        delete_edit_tree.column(col, width=100, minwidth=50, stretch=True)

    update_delete_edit_tree()

    button_delete_selected = tk.Button(delete_edit_entry_window, text="Delete Selected Entry",
                                       command=delete_selected_entry)
    button_delete_selected.pack(pady=10)

    delete_edit_tree.bind('<Double-1>', on_delete_edit_double_click)

    delete_edit_entry_window.bind('<Control-f>', lambda event: entry_search_delete_edit.focus_set())



def update_delete_edit_tree():
    for item in delete_edit_tree.get_children():
        delete_edit_tree.delete(item)

    for index, row in inventory_homepage_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Supplier Name"],
            row["Date of Purchase"],
            row["Date of Expiry"],
            row["Quantity Purchased"],
            row["Quantity Available"],
            row["No. of Units"],  # Make sure this is correctly placed
            row["Price"]  # Make sure this is correctly placed
        ]
        delete_edit_tree.insert("", "end", values=values)


def search_delete_edit_entry(event=None):
    search_term = entry_search_delete_edit.get().lower()

    for item in delete_edit_tree.get_children():
        delete_edit_tree.delete(item)

    for index, row in inventory_homepage_df.iterrows():
        if any(search_term in str(value).lower() for value in row):
            values = row.tolist()
            delete_edit_tree.insert("", "end", values=values)


def delete_selected_entry():
    selected_item = delete_edit_tree.selection()
    if selected_item:
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this entry?")
        if confirm:
            index = delete_edit_tree.index(selected_item)
            batch_number = delete_edit_tree.item(selected_item)["values"][2]
            inventory_homepage_df.drop(index, inplace=True)
            inventory_homepage_df.reset_index(drop=True, inplace=True)

            # Remove the item from returns_adjusted_df if exists
            if (returns_adjusted_df['Batch Number'] == batch_number).any():
                returns_adjusted_df.drop(returns_adjusted_df[returns_adjusted_df['Batch Number'] == batch_number].index,
                                         inplace=True)
                returns_adjusted_df.reset_index(drop=True, inplace=True)

            update_treeview()
            update_delete_edit_tree()
            save_to_excel()
            messagebox.showinfo("Deleted", "The selected entry has been deleted.")
        else:
            return
    else:
        messagebox.showerror("Error", "No entry selected.")


def on_delete_edit_double_click(event):
    selected_item = delete_edit_tree.selection()
    if selected_item:
        item = delete_edit_tree.item(selected_item)
        values = item['values']

        edit_window = tk.Toplevel(delete_edit_entry_window)
        edit_window.title("Edit Entry")

        frame_form = tk.Frame(edit_window)
        frame_form.pack(pady=10)

        # Create labels and entry fields
        labels = ["Medicine Name", "Brand Name", "Price", "No. of Units", "Batch Number",
                  "Supplier Name", "Date of Purchase (DD-MM-YYYY)", "Date of Expiry (DD-MM-YYYY)",
                  "Quantity Purchased", "Quantity Available"]
        entries = {}

        for i, label in enumerate(labels):
            tk.Label(frame_form, text=label).grid(row=i, column=0, padx=10, pady=5, sticky='w')
            entry = tk.Entry(frame_form, width=50)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            entries[label] = entry

        # Pre-fill the entries with the existing values
        entries["Medicine Name"].insert(0, values[0])
        entries["Brand Name"].insert(0, values[1])
        entries["Batch Number"].insert(0, values[2])
        entries["Supplier Name"].insert(0, values[3])
        entries["Date of Purchase (DD-MM-YYYY)"].insert(0, values[4])
        entries["Date of Expiry (DD-MM-YYYY)"].insert(0, values[5])
        entries["Quantity Purchased"].insert(0, values[6])
        entries["Quantity Available"].insert(0, values[7])
        entries["No. of Units"].insert(0, values[8])
        entries["Price"].insert(0, values[9])

        def save_edit():
            try:
                purchase_date = parse_date(entries["Date of Purchase (DD-MM-YYYY)"].get())
                expiry_date = parse_date(entries["Date of Expiry (DD-MM-YYYY)"].get())
                quantity_purchased = int(entries["Quantity Purchased"].get())
                quantity_available = float(entries["Quantity Available"].get())
                price = float(entries["Price"].get())
                no_of_units = int(entries["No. of Units"].get())  # Retrieve No. of Units
            except ValueError:
                messagebox.showerror("Error", "Invalid date format or numerical value")
                return

            index = delete_edit_tree.index(selected_item)
            inventory_homepage_df.loc[index, "Medicine Name"] = entries["Medicine Name"].get()
            inventory_homepage_df.loc[index, "Brand Name"] = entries["Brand Name"].get()
            inventory_homepage_df.loc[index, "Batch Number"] = entries["Batch Number"].get()
            inventory_homepage_df.loc[index, "Supplier Name"] = entries["Supplier Name"].get()
            inventory_homepage_df.loc[index, "Date of Purchase"] = purchase_date
            inventory_homepage_df.loc[index, "Date of Expiry"] = expiry_date
            inventory_homepage_df.loc[index, "Quantity Purchased"] = quantity_purchased
            inventory_homepage_df.loc[index, "Quantity Available"] = quantity_available
            inventory_homepage_df.loc[index, "No. of Units"] = no_of_units  # Update No. of Units
            inventory_homepage_df.loc[index, "Price"] = price

            update_treeview()
            update_delete_edit_tree()
            save_to_excel()
            edit_window.destroy()

        button_save = tk.Button(frame_form, text="Save Changes", command=save_edit)
        button_save.grid(row=len(labels), column=0, columnspan=2, pady=10)

        edit_window.bind('<Return>', lambda event: save_edit())

def focus_search_bar(event):
    entry_search.focus_set()

def open_returns_adjusted_window():
    global returns_adjusted_df, returns_adjusted_tree, entry_status, permanent_returnsadjusted_df

    returns_adjusted_window = tk.Toplevel(root)
    returns_adjusted_window.title("Returns/Adjusted")

    frame_items = tk.Frame(returns_adjusted_window)
    frame_items.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    returns_adjusted_scroll = tk.Scrollbar(frame_items)
    returns_adjusted_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns_returns_adjusted = ["Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Expiry",
                                "Status"]

    # Add yscrollcommand to connect the scrollbar with the Treeview
    returns_adjusted_tree = ttk.Treeview(frame_items, columns=columns_returns_adjusted, show="headings", yscrollcommand=returns_adjusted_scroll.set)
    returns_adjusted_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    returns_adjusted_scroll.config(command=returns_adjusted_tree.yview)

    for col in columns_returns_adjusted:
        returns_adjusted_tree.heading(col, text=col)
        returns_adjusted_tree.column(col, width=100, minwidth=50, stretch=True)

    def load_returns_adjusted_data():
        # Clear the current treeview items
        for item in returns_adjusted_tree.get_children():
            returns_adjusted_tree.delete(item)

        current_date = datetime.now()
        for index, row in inventory_homepage_df.iterrows():
            if (row['Date of Expiry'] - current_date).days <= 90:
                values = row.tolist()

                # Find the corresponding status in the returns_adjusted_df
                status_value = returns_adjusted_df[returns_adjusted_df['Batch Number'] == values[2]]

                # Check if there's a matching status
                if not status_value.empty:
                    status = status_value.iloc[-1]['Status']  # Use the most recent status if there are multiple matches
                else:
                    status = ""

                # Insert the values into the treeview
                returns_adjusted_tree.insert("", "end",
                                             values=values[:1] + values[1:3] + values[3:4] + values[5:6] + [status])

    def on_double_click(event):
        selected_item = returns_adjusted_tree.selection()
        if selected_item:
            item = returns_adjusted_tree.item(selected_item)
            values = item['values']

            def save_status():
                global returns_adjusted_df, permanent_returnsadjusted_df
                new_status = entry_status.get()
                returns_adjusted_tree.set(selected_item, column='Status', value=new_status)
                if (returns_adjusted_df['Batch Number'] == values[2]).any():
                    returns_adjusted_df.loc[returns_adjusted_df['Batch Number'] == values[2], 'Status'] = new_status
                else:
                    new_entry = pd.DataFrame([{
                        "Batch Number": values[2], "Status": new_status
                    }])
                    returns_adjusted_df = pd.concat([returns_adjusted_df, new_entry], ignore_index=True)

                # Ensure the entry is added to permanent_returnsadjusted_df
                permanent_entry = pd.DataFrame([{
                    "Medicine Name": values[0], "Brand Name": values[1], "Batch Number": values[2],
                    "Supplier Name": values[3], "Date of Expiry": values[4], "Status": new_status
                }])
                permanent_returnsadjusted_df = pd.concat([permanent_returnsadjusted_df, permanent_entry],
                                                         ignore_index=True)

                save_to_excel()
                load_returns_adjusted_data()
                popup.destroy()

            popup = tk.Toplevel(returns_adjusted_window)
            popup.title("Set Status")
            tk.Label(popup, text="Enter status (Returned/Adjusted):").pack(pady=10)
            entry_status = tk.Entry(popup)
            entry_status.pack(pady=10)
            entry_status.insert(0, values[5])
            button_save = tk.Button(popup, text="Save", command=save_status)
            button_save.pack(pady=10)

            popup.bind('<Return>', lambda event: save_status())

    returns_adjusted_tree.bind('<Double-1>', on_double_click)

    load_returns_adjusted_data()


def open_returns_adjusted_records_window():
    global permanent_returnsadjusted_df, records_returns_adjusted_tree, entry_search_records_returns_adjusted

    records_returns_adjusted_window = tk.Toplevel(root)
    records_returns_adjusted_window.title("Returns/Adjusted Records")

    frame_search_records_returns_adjusted = tk.Frame(records_returns_adjusted_window)
    frame_search_records_returns_adjusted.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

    entry_search_records_returns_adjusted = tk.Entry(frame_search_records_returns_adjusted, width=30)
    entry_search_records_returns_adjusted.pack(side=tk.LEFT, padx=5)
    entry_search_records_returns_adjusted.bind('<Return>', search_returns_adjusted_records)

    button_search_records_returns_adjusted = tk.Button(frame_search_records_returns_adjusted, text="Search",
                                                       command=search_returns_adjusted_records)
    button_search_records_returns_adjusted.pack(side=tk.LEFT, padx=5)

    frame_records_returns_adjusted = tk.Frame(records_returns_adjusted_window)
    frame_records_returns_adjusted.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a vertical scrollbar
    records_returns_adjusted_scroll = tk.Scrollbar(frame_records_returns_adjusted)
    records_returns_adjusted_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    columns_returns_adjusted_records = [
        "Medicine Name", "Brand Name", "Batch Number", "Supplier Name",
        "Date of Expiry", "Status"
    ]

    # Add yscrollcommand to connect the scrollbar with the treeview
    records_returns_adjusted_tree = ttk.Treeview(frame_records_returns_adjusted,
                                                 columns=columns_returns_adjusted_records, show="headings", yscrollcommand=records_returns_adjusted_scroll.set)
    records_returns_adjusted_tree.pack(fill=tk.BOTH, expand=True)

    # Configure the scrollbar to work with the Treeview
    records_returns_adjusted_scroll.config(command=records_returns_adjusted_tree.yview)

    for col in columns_returns_adjusted_records:
        records_returns_adjusted_tree.heading(col, text=col)
        records_returns_adjusted_tree.column(col, width=100, minwidth=50, stretch=True)

    for index, row in permanent_returnsadjusted_df.iterrows():
        values = [
            row["Medicine Name"],
            row["Brand Name"],
            row["Batch Number"],
            row["Supplier Name"],
            row["Date of Expiry"],
            row["Status"]          # Ensure this is in the correct position
        ]
        records_returns_adjusted_tree.insert("", "end", values=values)

    records_returns_adjusted_window.bind('<Control-f>', lambda event: entry_search_records_returns_adjusted.focus_set())


def search_returns_adjusted_records(event=None):
    search_term = entry_search_records_returns_adjusted.get().lower()

    for item in records_returns_adjusted_tree.get_children():
        records_returns_adjusted_tree.delete(item)

    for index, row in permanent_returnsadjusted_df.iterrows():
        if any(search_term in str(value).lower() for value in row):
            values = row.tolist()
            records_returns_adjusted_tree.insert("", "end", values=values)


def refresh_data():
    global inventory_homepage_df, sales_df, returns_adjusted_df, permanent_returnsadjusted_df
    inventory_homepage_df = pd.read_excel('inventory_homepage.xlsx')
    sales_df = pd.read_excel('sales.xlsx')
    returns_adjusted_df = pd.read_excel('returns_adjusted.xlsx')
    permanent_returnsadjusted_df = pd.read_excel('permanent_returnsadjusted.xlsx')
    update_treeview()


def generate_reports():
    report_window = tk.Toplevel(root)
    report_window.title("Generate Reports")

    frame_report_options = tk.Frame(report_window)
    frame_report_options.pack(pady=10, fill=tk.BOTH, expand=True)

    tk.Label(frame_report_options, text="Select Report Type:").grid(row=0, column=0, padx=5, pady=5, sticky='w')

    report_type = tk.StringVar()
    report_type.set("Sales")

    # Aligning radio buttons vertically
    tk.Radiobutton(frame_report_options, text="Sales", variable=report_type, value="Sales").grid(row=1, column=0, padx=5, pady=5, sticky='w')
    tk.Radiobutton(frame_report_options, text="Inventory", variable=report_type, value="Inventory").grid(row=2, column=0, padx=5, pady=5, sticky='w')

    def generate():
        selected_report = report_type.get()
        if selected_report == "Sales":
            df = sales_df

            # Ensure 'Quantity Sold' is a numeric type before proceeding
            df['Quantity Sold'] = pd.to_numeric(df['Quantity Sold'], errors='coerce')

            # Group by 'Medicine Name' and sum the 'Quantity Sold'
            item_sales = df.groupby('Medicine Name')['Quantity Sold'].sum()

            # Create a pie chart showing which item sells the most
            plt.figure(figsize=(10, 6))
            item_sales.plot.pie(autopct='%1.1f%%', startangle=90, counterclock=False)
            plt.title('Sales Distribution by Medicine')
            plt.ylabel('')  # Hide the y-label
            plt.tight_layout()

            # Ensure the directory exists for saving the plot
            directory = "Reports"
            if not os.path.exists(directory):
                os.makedirs(directory)

            # Save the pie chart as an image
            pie_chart_filename = os.path.join(directory,
                                              f"{selected_report}_Pie_Chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            plt.savefig(pie_chart_filename)
            plt.show()  # Optionally show the plot window

        elif selected_report == "Inventory":
            df = inventory_homepage_df
        else:
            messagebox.showerror("Error", "Invalid report type selected.")
            return

        # Ensure the directory exists for saving the report
        directory = "Reports"
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Save the report to Excel in the specified directory
        report_filename = os.path.join(directory,
                                       f"{selected_report}_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df.to_excel(report_filename, index=False)
        messagebox.showinfo("Report Generated",
                            f"{selected_report} report has been generated and saved as {report_filename}.\n"
                            f"A pie chart has been saved as {pie_chart_filename} if applicable.")

    button_generate = tk.Button(frame_report_options, text="Generate Report", command=generate)
    button_generate.grid(row=4, column=0, padx=5, pady=10, sticky='w')

# (Same functions for handling other parts of the application)
root = tk.Tk()
root.title("Inventory Management System by Siddhartha Das")

root.bind('<Control-f>', focus_search_bar)
root.bind('<Return>', lambda event: search_item())

frame_buttons = tk.Frame(root)
frame_buttons.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

# Define the button style
button_style = {
    'font': ('Helvetica', 10),
    'bg': '#4CAF50',
    'fg': 'white',
    'relief': 'raised',
    'bd': 2
}

# Place buttons using grid layout
buttons = [
    ("Add New Entry", open_new_entry_window),
    ("Check Expiry", check_expiry),
    ("Short List Item", shortlist_items),
    ("Item Records", open_item_records_window),
    ("Sell Multiple Items", open_sell_multiple_window),
    ("Sales Records", open_sales_records_window),
    ("Delete/Edit Entry", open_delete_edit_entry_window),
    ("Returns/Adjusted", open_returns_adjusted_window),
    ("Returns/Adjusted Records", open_returns_adjusted_records_window),
    ("Refresh", refresh_data),
    ("Generate Reports", generate_reports)
]

for i, (text, command) in enumerate(buttons):
    tk.Button(frame_buttons, text=text, command=command, **button_style).grid(row=0, column=i, padx=5, pady=5, sticky='ew')

entry_search = tk.Entry(frame_buttons, width=30)
entry_search.grid(row=0, column=len(buttons), padx=5, pady=5)

button_search = tk.Button(frame_buttons, text="Search", command=search_item, **button_style)
button_search.grid(row=0, column=len(buttons) + 1, padx=5, pady=5, sticky='ew')

# Frame for the Treeview and Scrollbar
frame_tree = tk.Frame(root)
frame_tree.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Create a vertical scrollbar for the Treeview
tree_scroll = tk.Scrollbar(frame_tree)
tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

columns = ("Medicine Name", "Brand Name", "Batch Number", "Supplier Name", "Date of Purchase", "Date of Expiry",
           "Quantity Purchased", "Quantity Available", "No. of Units", "Price")

# Add the Treeview widget
inventory_tree = ttk.Treeview(frame_tree, columns=columns, show="headings", yscrollcommand=tree_scroll.set)
inventory_tree.pack(fill=tk.BOTH, expand=True)

# Configure the scrollbar
tree_scroll.config(command=inventory_tree.yview)

for col in columns:
    inventory_tree.heading(col, text=col)
    inventory_tree.column(col, width=100, minwidth=50, stretch=True)

def on_tree_select(event):
    # Removed the sell item button functionality
    pass

inventory_tree.bind('<<TreeviewSelect>>', on_tree_select)

load_item_list()
root.mainloop()
