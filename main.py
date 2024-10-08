import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import pandas as pd

class BarcodeScannerApp:
    def __init__(self, master):
        self.master = master
        master.title("Barcode Check In/Out")
        
        # Set window size to 600x400
        master.geometry("600x400")
        
        # Set window background color to white
        master.configure(bg="white")

        # Label with custom font and color
        self.status_label = tk.Label(master, text="Scan a barcode or enter item code", 
                                     font=("Arial", 16, "bold"), bg="white", fg="black")
        self.status_label.pack(pady=20)

        # Check In and Check Out buttons with orange color
        self.check_in_button = tk.Button(master, text="Check In", command=self.check_in, 
                                         bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.check_in_button.pack(pady=5)

        self.check_out_button = tk.Button(master, text="Check Out", command=self.check_out, 
                                          bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.check_out_button.pack(pady=5)

        # Export to Excel button
        self.export_button = tk.Button(master, text="Export to Excel", command=self.export_to_excel, 
                                       bg="#ff9800", fg="black", font=("Arial", 12, "bold"))
        self.export_button.pack(pady=5)

        # Entry field for item code with default styling
        self.item_code_entry = tk.Entry(master, width=30, font=("Arial", 12))
        self.item_code_entry.pack(pady=10)
        self.item_code_entry.focus()  # Focus on the entry to allow barcode scanning

        # Data structure to hold items and their statuses
        self.items = {}

        # Create a Treeview widget for the overview
        self.tree = ttk.Treeview(master, columns=("Item Code", "Status"), show="headings")
        self.tree.heading("Item Code", text="Item Code")
        self.tree.heading("Status", text="Status")

        # Configure column widths
        self.tree.column("Item Code", width=250)  # Wider column for item codes
        self.tree.column("Status", width=100)

        self.tree.pack(pady=5, fill=tk.BOTH, expand=True)  # Allow the Treeview to expand

        # Add a scrollbar to the Treeview
        self.scrollbar = ttk.Scrollbar(master, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=self.scrollbar.set)

        # Style the Treeview
        self.tree.tag_configure('checked_in', background="#e8f5e9")  # Light green for checked in
        self.tree.tag_configure('checked_out', background="#ffebee")  # Light red for checked out

        # Create context menu for removing items
        self.context_menu = tk.Menu(master, tearoff=0)
        self.context_menu.add_command(label="Remove", command=self.remove_item)

        # Bind right-click event to the Treeview
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        # Get the item clicked on
        try:
            item = self.tree.identify_row(event.y)
            if item:  # If an item was clicked
                self.tree.selection_set(item)  # Select the item
                self.context_menu.post(event.x_root, event.y_root)  # Show the context menu
        except Exception as e:
            print(e)

    def remove_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_code = self.tree.item(selected_item, "values")[0]  # Get the item code
            del self.items[item_code]  # Remove from internal dictionary
            self.tree.delete(selected_item)  # Remove from Treeview
            messagebox.showinfo("Info", f"Item {item_code} removed.")
        else:
            messagebox.showerror("Error", "No item selected to remove.")

    def export_to_excel(self):
        # Ask for a file path to save the Excel file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return  # If no file is selected, exit the function

        # Create a DataFrame from the items dictionary
        df = pd.DataFrame(self.items.items(), columns=["Item Code", "Status"])
        
        # Save the DataFrame to an Excel file
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Info", f"Data exported to {file_path} successfully.")

    def check_in(self):
        item_code = self.item_code_entry.get()
        if item_code:
            self.items[item_code] = "Checked In"  # Update item status
            messagebox.showinfo("Info", f"Item {item_code} checked in.")
            self.update_overview()  # Update the Treeview
            self.item_code_entry.delete(0, tk.END)  # Clear entry after check in
        else:
            messagebox.showerror("Error", "No item code entered.")

    def check_out(self):
        item_code = self.item_code_entry.get()
        if item_code:
            self.items[item_code] = "Checked Out"  # Update item status
            messagebox.showinfo("Info", f"Item {item_code} checked out.")
            self.update_overview()  # Update the Treeview
            self.item_code_entry.delete(0, tk.END)  # Clear entry after check out
        else:
            messagebox.showerror("Error", "No item code entered.")

    def update_overview(self):
        # Clear the Treeview before updating
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert updated items into the Treeview with styling
        for item_code, status in self.items.items():
            tag = 'checked_in' if status == "Checked In" else 'checked_out'
            self.tree.insert("", tk.END, values=(item_code, status), tags=(tag,))

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeScannerApp(root)
    root.mainloop()
