# PaymentRecordsApp 🧾💰

This is a simple **payment tracking application** built with Python, Tkinter, and OpenPyXL.

## 🔹 Features:
- ✅ Add, update, search, and list payments  
- 📊 Analyze payments (Paid vs. Pending)  
- 📈 Generate visual charts  
- 📁 Save data in an Excel file (`payment_records.xlsx`)  
- 🎨 GUI built using Tkinter  

## 🚀 Installation & Run:
### 1. Clone the repository:
```sh
git clone https://github.com/ozcotech/paymentRecords.git
cd paymentRecords
```
### 2. Install dependencies:
```sh
pip install -r requirements.txt
```
### 3. Run the GUI:
```sh
python gui.py
```

## 📦 Build as an App:
To create a standalone macOS app:
```sh
pyinstaller --onefile --windowed --paths=data --hidden-import=tkinter --hidden-import=matplotlib --add-data "data/excel_manager.py:data" --add-data "data/models.py:data" --name "PaymentRecordsApp" gui.py
```

## 🛠️ Technologies Used:
- **Python** 🐍
- **Tkinter** 🎨
- **OpenPyXL** 📊
- **Matplotlib** 📈

---
Made with ❤️ by [ozcotech](https://ozco.studio)

