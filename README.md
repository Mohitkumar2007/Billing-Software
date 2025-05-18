# Billing & Inventory Management System

A simple desktop application for small businesses to manage inventory and generate bills, built with Python and PyQt5. 

## Features
- **Inventory Management**: Add, update, delete, and search items by barcode or name. View inventory in a searchable table. Export inventory to CSV and create Excel backups.
- **Billing System**: Search and add items to a cart, validate stock, and generate bills with optional GST (18%). Each bill is saved as a new sheet in an Excel file, and inventory is updated automatically. Export bills as CSV and backup all bills.
- **User-Friendly Interface**: Modern UI with tooltips, keyboard shortcuts, and error handling. Optional dark mode support.

## Requirements
- Python 3.x
- PyQt5
- pandas
- openpyxl
- qdarkstyle (optional, for dark theme)

## Usage
1. Run `Inventory_entry.py` to manage your inventory.
2. Run `billing.py` to generate bills and update stock.
3. Data is stored in `items.xlsx` (inventory) and `bills.xlsx` (bills). Logo image is `logo.jpg`.

## How to Run
Install dependencies (if not already):
```bash
pip install PyQt5 pandas openpyxl qdarkstyle
```
Then run either script:
```bash
python Inventory_entry.py
python billing.py
```

## Project Structure
```
billing.py              # Billing system GUI
Inventory_entry.py      # Inventory management GUI
items.xlsx              # Inventory data (Excel)
bills.xlsx              # Bills data (Excel)
logo.jpg                # Logo image
```

## License
This project is provided as-is for educational and small business use.

