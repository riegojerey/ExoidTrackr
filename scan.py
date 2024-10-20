import tkinter as tk
from tkinter import filedialog, simpledialog
from tkinter import ttk, messagebox
import pandas as pd

class BarcodeScannerApp:
    def __init__(self, master):
        self.master = master
        master.title("ToolsNet - Inventory Management System")
        
        # Set window size to 900x900
        master.geometry("900x900")
        
        # Set window background color to white
        master.configure(bg="white")

        # Branding label for ToolsNet at the top
        self.branding_label = tk.Label(master, text="ToolsNet", 
                                       font=("Arial", 24, "bold"), bg="white", fg="#ff9800")
        self.branding_label.pack(pady=10)

        # Subtitle or description for the app
        self.description_label = tk.Label(master, text="Inventory Management System for Tools", 
                                          font=("Arial", 14), bg="white", fg="black")
        self.description_label.pack(pady=5)

        # Label with instructions
        self.status_label = tk.Label(master, text="Scan a barcode or enter item code", 
                                     font=("Arial", 16, "bold"), bg="white", fg="black")
        self.status_label.pack(pady=20)

        # Add a toggle button for switching between Check In and Check Out modes
        self.mode = tk.StringVar(value="Check In")
        self.toggle_button = tk.Button(master, text="Switch to Check Out", command=self.toggle_mode, 
                                       bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.toggle_button.pack(pady=5)

        # Export to Excel button
        self.export_button = tk.Button(master, text="Export to Excel", command=self.export_to_excel, 
                                       bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.export_button.pack(pady=5)

        # Entry field for item code with default styling
        self.item_code_entry = tk.Entry(master, width=30, font=("Arial", 12))
        self.item_code_entry.pack(pady=10)
        self.item_code_entry.focus()  # Focus on the entry to allow barcode scanning

        # Bind the Enter key to the process function (either Check In or Check Out based on mode)
        self.item_code_entry.bind("<Return>", self.process_item)

        # Data structure to hold items and their statuses
        self.items = {}

        # Create a Treeview widget for the overview
        self.tree = ttk.Treeview(master, columns=("Item Code", "Description", "Status", "Quantity"), show="headings")
        self.tree.heading("Item Code", text="Item Code")
        self.tree.heading("Description", text="Description")  # New column for description
        self.tree.heading("Status", text="Status")
        self.tree.heading("Quantity", text="Quantity")  # New column for quantity

        # Configure column widths
        self.tree.column("Item Code", width=250)
        self.tree.column("Description", width=350)  # Set width for the description
        self.tree.column("Status", width=100)
        self.tree.column("Quantity", width=100)  # Set width for quantity

        self.tree.pack(pady=5, fill=tk.BOTH, expand=True)

        # Add a scrollbar to the Treeview
        self.scrollbar = ttk.Scrollbar(master, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=self.scrollbar.set)

        # Style the Treeview
        self.tree.tag_configure('checked_in', background="#e8f5e9")  # Light green background
        self.tree.tag_configure('checked_out', background="#ffebee")  # Light red background

        # Create context menu for removing items
        self.context_menu = tk.Menu(master, tearoff=0)
        self.context_menu.add_command(label="Remove", command=self.remove_item)

        # Bind right-click event to the Treeview
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Load Excel file into pandas DataFrame (initialize with None)
        self.inventory_df = None
        self.load_excel_data()

        # Create plus and minus buttons for quantity adjustment
        self.quantity_frame = tk.Frame(master)
        self.quantity_frame.pack(pady=10)

        self.minus_button = tk.Button(self.quantity_frame, text="-", command=self.decrease_quantity, 
                                       bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.minus_button.pack(side=tk.LEFT)

        self.plus_button = tk.Button(self.quantity_frame, text="+", command=self.increase_quantity, 
                                      bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.plus_button.pack(side=tk.LEFT)

        # Enable buttons by default
        self.minus_button.config(state=tk.NORMAL)
        self.plus_button.config(state=tk.NORMAL)

    def load_excel_data(self):
        # Prompt the user to select an Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            messagebox.showerror("Error", "No Excel file selected. Please load a valid file.")
            return
        
        try:
            # Load the Excel file into a pandas DataFrame
            self.inventory_df = pd.read_excel(file_path)
            if 'Item Code' not in self.inventory_df.columns or 'Description' not in self.inventory_df.columns:
                raise ValueError("Excel file must have 'Item Code' and 'Description' columns.")
            self.inventory_file_path = file_path  # Store the path for saving later
            messagebox.showinfo("Success", "Excel file loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def toggle_mode(self):
        if self.mode.get() == "Check In":
            self.mode.set("Check Out")
            self.toggle_button.config(text="Switch to Check In")
            self.status_label.config(text="Scan a barcode or enter item code (Check Out mode)")
        else:
            self.mode.set("Check In")
            self.toggle_button.config(text="Switch to Check Out")
            self.status_label.config(text="Scan a barcode or enter item code (Check In mode)")

    def show_context_menu(self, event):
        try:
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            print(e)

    def remove_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_code = self.tree.item(selected_item, "values")[0]
            del self.items[item_code]
            self.tree.delete(selected_item)
        else:
            print("No item selected to remove.")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return
        df = pd.DataFrame([(code, details['Description'], details['Status'], details['Quantity']) 
                            for code, details in self.items.items()], 
                          columns=["Item Code", "Description", "Status", "Quantity"])
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "Data exported to Excel successfully.")

    def process_item(self, event=None):
        item_code = self.item_code_entry.get().strip().lower()

        if not item_code:
            print("No item code entered.")
            return
        
        # Lookup the item code in the loaded Excel file
        if self.inventory_df is not None:
            match = self.inventory_df[self.inventory_df['Item Code'].str.lower() == item_code]
            if not match.empty:
                description = match['Description'].values[0]
                if self.mode.get() == "Check In":
                    if item_code in self.items:
                        # If already in items, check the current status
                        if self.items[item_code]['Status'] == "Checked Out":
                            self.items[item_code]['Status'] = "Checked In"  # Change status
                            # No incrementing quantity on status change
                        else:
                            # Increment the quantity for the existing checked-in item
                            self.items[item_code]['Quantity'] += 1
                    else:
                        # New item entry
                        self.items[item_code] = {'Status': "Checked In", 'Description': description, 'Quantity': 1}
                else:
                    self.check_out(item_code, description)
                
                # Update the overview for the Treeview
                self.update_overview(item_code)
            else:
                # Prompt for a description for unknown item and save it
                description = self.prompt_for_description(item_code)
                if description:
                    self.add_unknown_item_to_database(item_code, description)
                    self.check_in(item_code, description)  # Check in the unknown item
        else:
            print("No inventory data loaded.")

        self.item_code_entry.delete(0, tk.END)  # Clear the input field after processing

    def check_in(self, item_code, description):
        if item_code in self.items:
            # If the item is already checked out, switch it back to checked in without incrementing quantity
            if self.items[item_code]['Status'] == "Checked Out":
                self.items[item_code]['Status'] = "Checked In"
                # No increment on quantity here
            else:
                self.items[item_code]['Quantity'] += 1  # Increment quantity if already checked in
        else:
            # Add a new item if it doesn't exist
            self.items[item_code] = {'Status': "Checked In", 'Description': description, 'Quantity': 1}
        self.update_overview(item_code)

    def check_out(self, item_code, description):
        if item_code in self.items:
            self.items[item_code]['Status'] = "Checked Out"  # Update status to Checked Out
        else:
            self.items[item_code] = {'Status': "Checked Out", 'Description': description, 'Quantity': 1}
        self.update_overview(item_code)

    def update_overview(self, item_code):
        # Clear the Treeview before updating
        for item in self.tree.get_children():
            self.tree.delete(item)

        for code, details in self.items.items():
            if details['Status'] == "Checked In":
                self.tree.insert("", tk.END, values=(code, details['Description'], details['Status'], details['Quantity']), tags=('checked_in',))
            else:
                self.tree.insert("", tk.END, values=(code, details['Description'], details['Status'], details['Quantity']), tags=('checked_out',))

    def prompt_for_description(self, item_code):
        return simpledialog.askstring("Unknown Item", f"Enter a description for the item code: {item_code}")

    def add_unknown_item_to_database(self, item_code, description):
        new_row = pd.DataFrame([[item_code, description]], columns=["Item Code", "Description"])
        self.inventory_df = pd.concat([self.inventory_df, new_row], ignore_index=True)
        # Save the updated DataFrame back to Excel
        self.inventory_df.to_excel(self.inventory_file_path, index=False)
        messagebox.showinfo("Success", f"Item {item_code} added to the database.")

    def increase_quantity(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_code = self.tree.item(selected_item, "values")[0]
            if item_code in self.items:
                self.items[item_code]['Quantity'] += 1  # Increase quantity by 1
                self.update_overview(item_code)  # Update the Treeview

    def decrease_quantity(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_code = self.tree.item(selected_item, "values")[0]
            if item_code in self.items and self.items[item_code]['Quantity'] > 0:
                self.items[item_code]['Quantity'] -= 1  # Decrease quantity by 1
                self.update_overview(item_code)  # Update the Treeview

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeScannerApp(root)
    root.mainloop()
