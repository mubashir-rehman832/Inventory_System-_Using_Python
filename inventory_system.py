
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook

# ===== OOP CLASSES ===== #
class Product:
    def __init__(self, name, price, quantity, supplier):
        self.name = name
        self.price = price
        self.quantity = quantity
        self.supplier = supplier

    def get_discounted_price(self):
        return self.price  # Default: no discount

class Electronics(Product):
    def get_discounted_price(self):
        return self.price * 0.9  # 10% discount

class Clothing(Product):
    def get_discounted_price(self):
        return self.price * 0.8  # 20% discount

# ===== INVENTORY CLASS ===== #
class Inventory:
    def __init__(self):
        self.products = []

    def add_product(self, product):
        self.products.append(product)

    def get_all_products(self):
        return self.products

    def export_to_excel(self, filename="inventory_report.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Type", "Price", "Discounted Price", "Quantity", "Supplier"])

        for p in self.products:
            ws.append([
                p.name,
                type(p).__name__,
                p.price,
                p.get_discounted_price(),
                p.quantity,
                p.supplier
            ])
        wb.save(filename)
        return filename

# ===== GUI APP ===== #
class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Management System")

        self.inventory = Inventory()

        # Entry Fields
        tk.Label(root, text="Product Name").grid(row=0, column=0)
        tk.Label(root, text="Price").grid(row=1, column=0)
        tk.Label(root, text="Quantity").grid(row=2, column=0)
        tk.Label(root, text="Supplier").grid(row=3, column=0)
        tk.Label(root, text="Type").grid(row=4, column=0)

        self.name_var = tk.StringVar()
        self.price_var = tk.DoubleVar()
        self.quantity_var = tk.IntVar()
        self.supplier_var = tk.StringVar()
        self.type_var = tk.StringVar()

        tk.Entry(root, textvariable=self.name_var).grid(row=0, column=1)
        tk.Entry(root, textvariable=self.price_var).grid(row=1, column=1)
        tk.Entry(root, textvariable=self.quantity_var).grid(row=2, column=1)
        tk.Entry(root, textvariable=self.supplier_var).grid(row=3, column=1)
        ttk.Combobox(root, textvariable=self.type_var, values=["Electronics", "Clothing"]).grid(row=4, column=1)

        # Buttons
        tk.Button(root, text="Add Product", command=self.add_product).grid(row=5, column=0, columnspan=2, pady=5)
        tk.Button(root, text="Export to Excel", command=self.export_data).grid(row=6, column=0, columnspan=2, pady=5)

        # TreeView to display products
        self.tree = ttk.Treeview(root, columns=("Name", "Type", "Price", "Discounted", "Quantity", "Supplier"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.grid(row=7, column=0, columnspan=2, pady=10)

    def add_product(self):
        name = self.name_var.get()
        price = self.price_var.get()
        qty = self.quantity_var.get()
        supplier = self.supplier_var.get()
        ptype = self.type_var.get()

        if ptype == "Electronics":
            product = Electronics(name, price, qty, supplier)
        elif ptype == "Clothing":
            product = Clothing(name, price, qty, supplier)
        else:
            messagebox.showerror("Invalid Type", "Please select a valid product type.")
            return

        self.inventory.add_product(product)
        self.tree.insert("", "end", values=(name, ptype, price, product.get_discounted_price(), qty, supplier))
        messagebox.showinfo("Success", "Product added successfully!")

    def export_data(self):
        file = self.inventory.export_to_excel()
        messagebox.showinfo("Exported", f"Data exported to {file}")

# ===== RUN APP ===== #
if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()
