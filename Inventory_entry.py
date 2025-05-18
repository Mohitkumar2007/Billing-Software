import sys
from PyQt5 import QtWidgets, QtGui
import pandas as pd

try:
    import qdarkstyle
    THEME = True
except ImportError:
    THEME = False

ITEMS_FILE = r"C:\Users\mohit\Documents\Coding\Project\BILLING_SOFTWARE\items.xlsx"
LOGO_IMAGE = r"C:\Users\mohit\Documents\Coding\Project\BILLING_SOFTWARE\logo.jpg"  # Your company logo

class InventoryApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inventory Manager")
        self.setWindowIcon(QtGui.QIcon(LOGO_IMAGE))
        self.resize(700, 400)
        self.init_ui()

    def init_ui(self):
        main_layout = QtWidgets.QVBoxLayout(self)
        form_layout = QtWidgets.QFormLayout()

        self.barcode_input = QtWidgets.QLineEdit()
        self.barcode_input.setPlaceholderText("Enter barcode")
        self.barcode_input.setToolTip("Unique item barcode")
        self.name_input = QtWidgets.QLineEdit()
        self.name_input.setPlaceholderText("Enter item name")
        self.name_input.setToolTip("Item name")
        self.qty_input = QtWidgets.QSpinBox()
        self.qty_input.setRange(1, 100000)
        self.qty_input.setToolTip("Quantity to add/update")
        self.price_input = QtWidgets.QDoubleSpinBox()
        self.price_input.setRange(0.01, 1000000)
        self.price_input.setPrefix("₹")
        self.price_input.setToolTip("Item price")
        self.add_btn = QtWidgets.QPushButton("Add/Update Item")
        self.add_btn.setToolTip("Add new item or update existing one")
        self.add_btn.clicked.connect(self.add_item)
        self.add_btn.setShortcut("Ctrl+S")

        form_layout.addRow("Barcode:", self.barcode_input)
        form_layout.addRow("Name:", self.name_input)
        form_layout.addRow("Quantity:", self.qty_input)
        form_layout.addRow("Price:", self.price_input)
        form_layout.addRow(self.add_btn)

        # Search bar
        search_layout = QtWidgets.QHBoxLayout()
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.setPlaceholderText("Search by barcode or name...")
        self.search_input.textChanged.connect(self.filter_table)
        search_layout.addWidget(QtWidgets.QLabel("Search:"))
        search_layout.addWidget(self.search_input)

        # Table for inventory
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Barcode", "Name", "Quantity", "Price"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.table.setToolTip("Inventory items list")
        self.table.itemSelectionChanged.connect(self.on_table_select)

        # Edit/Delete/Export/Backup buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.edit_btn = QtWidgets.QPushButton("Edit Selected")
        self.edit_btn.setToolTip("Edit selected item")
        self.edit_btn.clicked.connect(self.edit_item)
        self.edit_btn.setEnabled(False)
        self.delete_btn = QtWidgets.QPushButton("Delete Selected")
        self.delete_btn.setToolTip("Delete selected item")
        self.delete_btn.clicked.connect(self.delete_item)
        self.delete_btn.setEnabled(False)
        self.export_btn = QtWidgets.QPushButton("Export as CSV")
        self.export_btn.setToolTip("Export inventory to CSV file")
        self.export_btn.clicked.connect(self.export_csv)
        self.backup_btn = QtWidgets.QPushButton("Backup Inventory")
        self.backup_btn.setToolTip("Backup inventory Excel file")
        self.backup_btn.clicked.connect(self.backup_inventory)
        btn_layout.addWidget(self.edit_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addWidget(self.export_btn)
        btn_layout.addWidget(self.backup_btn)

        main_layout.addLayout(form_layout)
        main_layout.addLayout(search_layout)
        main_layout.addWidget(self.table)
        main_layout.addLayout(btn_layout)

        self.load_table()

    def load_table(self):
        try:
            df = pd.read_excel(ITEMS_FILE)
        except Exception as e:
            df = pd.DataFrame(columns=["Barcode", "Name", "Quantity", "Price"])
        self.df = df
        self.displayed_df = df.copy()
        self.refresh_table()

    def refresh_table(self):
        df = self.displayed_df
        self.table.setRowCount(0)
        for _, row in df.iterrows():
            row_pos = self.table.rowCount()
            self.table.insertRow(row_pos)
            self.table.setItem(row_pos, 0, QtWidgets.QTableWidgetItem(str(row["Barcode"])))
            self.table.setItem(row_pos, 1, QtWidgets.QTableWidgetItem(str(row["Name"])))
            self.table.setItem(row_pos, 2, QtWidgets.QTableWidgetItem(str(row["Quantity"])))
            self.table.setItem(row_pos, 3, QtWidgets.QTableWidgetItem(f"₹{row['Price']:.2f}"))
        self.edit_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)

    def filter_table(self):
        text = self.search_input.text().lower()
        if text:
            mask = self.df["Barcode"].astype(str).str.lower().str.contains(text) | self.df["Name"].astype(str).str.lower().str.contains(text)
            self.displayed_df = self.df[mask].copy()
        else:
            self.displayed_df = self.df.copy()
        self.refresh_table()

    def on_table_select(self):
        selected = self.table.selectedItems()
        self.edit_btn.setEnabled(bool(selected))
        self.delete_btn.setEnabled(bool(selected))

    def add_item(self):
        barcode = self.barcode_input.text().strip()
        name = self.name_input.text().strip()
        qty = self.qty_input.value()
        price = self.price_input.value()
        if not barcode or not name:
            QtWidgets.QMessageBox.warning(self, "Error", "Barcode and Name are required!")
            return
        try:
            df = pd.read_excel(ITEMS_FILE)
            required_columns = {"Barcode", "Name", "Quantity", "Price"}
            if not required_columns.issubset(df.columns):
                raise ValueError(f"The file {ITEMS_FILE} must contain the following columns: {required_columns}")
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Barcode", "Name", "Quantity", "Price"])
            QtWidgets.QMessageBox.information(self, "Info", f"{ITEMS_FILE} not found. A new file will be created.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error reading file: {e}")
            return
        row = df[df["Barcode"].astype(str) == barcode]
        if not row.empty:
            idx = row.index[0]
            new_qty = df.at[idx, "Quantity"] + qty
            if new_qty < 0:
                QtWidgets.QMessageBox.warning(self, "Error", "Quantity cannot be negative!")
                return
            df.at[idx, "Quantity"] = new_qty
            df.at[idx, "Price"] = price
        else:
            df = pd.concat([df, pd.DataFrame([{
                "Barcode": barcode,
                "Name": name,
                "Quantity": qty,
                "Price": price
            }])], ignore_index=True)
        try:
            df.to_excel(ITEMS_FILE, index=False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save inventory: {e}")
            return
        QtWidgets.QMessageBox.information(self, "Success", "Item Added/Updated in Inventory!")
        self.barcode_input.clear()
        self.name_input.clear()
        self.qty_input.setValue(1)
        self.price_input.setValue(0.01)
        self.load_table()

    def edit_item(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        row = self.table.currentRow()
        barcode = self.table.item(row, 0).text()
        item = self.df[self.df["Barcode"].astype(str) == barcode]
        if item.empty:
            return
        item = item.iloc[0]
        self.barcode_input.setText(str(item["Barcode"]))
        self.name_input.setText(str(item["Name"]))
        self.qty_input.setValue(int(item["Quantity"]))
        self.price_input.setValue(float(item["Price"]))
        self.barcode_input.setFocus()

    def delete_item(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        row = self.table.currentRow()
        barcode = self.table.item(row, 0).text()
        df = self.df[self.df["Barcode"].astype(str) != barcode].copy()
        try:
            df.to_excel(ITEMS_FILE, index=False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to delete item: {e}")
            return
        QtWidgets.QMessageBox.information(self, "Success", "Item deleted from inventory.")
        self.load_table()

    def export_csv(self):
        try:
            df = pd.read_excel(ITEMS_FILE)
            save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export Inventory as CSV", "inventory.csv", "CSV Files (*.csv)")
            if save_path:
                df.to_csv(save_path, index=False)
                QtWidgets.QMessageBox.information(self, "Success", f"Inventory exported to {save_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to export: {e}")

    def backup_inventory(self):
        import shutil
        try:
            backup_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Backup Inventory Excel File", "items_backup.xlsx", "Excel Files (*.xlsx)")
            if backup_path:
                shutil.copy2(ITEMS_FILE, backup_path)
                QtWidgets.QMessageBox.information(self, "Success", f"Backup created at {backup_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to backup: {e}")

def main():
    app = QtWidgets.QApplication(sys.argv)
    if THEME:
        app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = InventoryApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()