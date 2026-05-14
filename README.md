# 👟 Adarsh Shoes Inventory Management System

A comprehensive desktop-based inventory and sales management application built with Python and CustomTkinter for footwear businesses. The system helps manage stock, process sales and returns, generate QR codes, visualize sales trends, and export reports to Excel.

---

## 🚀 Features

### 📦 Inventory Management
- Add new stock items using unique item codes.
- Update quantities for existing items.
- Search inventory by name, size, color, article, or unique ID.
- Edit or delete existing inventory records.
- Automatic low-stock alerts.

### 💰 Sales Management
- Sell items with quantity tracking.
- Record return/restock transactions.
- Maintain complete sales history.

### 📊 Dashboard & Analytics
- Real-time dashboard with:
  - Total stock
  - Today's sales
  - Today's returns
- Interactive sales charts for:
  - Today
  - This week
  - This month
  - Last 3, 6, and 9 months
  - This year
  - Last 2 and 5 years

### 🔍 Reporting
- Filter sales reports by date and floor.
- Export stock and sales reports to Excel.

### 🏷️ QR Code Generator
- Generate QR codes for products based on:
  - Brand
  - Article
  - Size
  - Color
- Save QR codes as PNG files.

### 🖼️ Branding Support
- Display a company logo in the application sidebar.

---

## 🛠️ Technologies Used

- Python 3.x
- CustomTkinter
- Pandas
- OpenPyXL
- Matplotlib
- QRCode
- Pillow (PIL)

---

## 📂 Project Structure

```text
Adarsh-Shoes-Inventory/
├── app.py
├── logo.jpg
├── stock.xlsx
├── sales_history.xlsx
├── requirements.txt
└── README.md