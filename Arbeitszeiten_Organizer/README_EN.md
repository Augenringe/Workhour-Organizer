# Working Hours Processing System
## 🚀 Modular Version - Completely Self-Contained!

## 📋 What is this?

This **professional application** automatically converts your CSV files with working hours into beautiful Excel files with overviews and calculations. **This version is fully modular and works anywhere - simply copy the entire `Working_Hours_Organizer` folder!**

## 🎯 What do you get?

A professional Excel file with **multiple worksheets** for each month:

1. **📊 Raw Data** - All your original data, beautifully formatted
2. **👥 Employee Overview** - Totals per employee
3. **📅 Daily Overview** - Statistics per day
4. **📈 Monthly Overview** - Total key figures for the month
5. **👤 Individual Employee Sheets** - A separate worksheet for each employee

## 📁 Modular Folder Structure

The program is fully modular and automatically creates a clean folder structure:

```
Working_Hours_Organizer/       ← Main folder (can be copied anywhere!)
├── CSV_Input/                 ← **PLACE CSV FILES HERE**
│   ├── Company_July_2025.csv
│   ├── Working_Hours_August_2025.csv
│   └── Time_Tracking_September_2025.csv
├── Excel_Output/              ← **FIND YOUR EXCEL FILES HERE**
│   ├── Working_Hours_Analysis_July_2025.xlsx
│   ├── Working_Hours_Analysis_August_2025.xlsx
│   └── Working_Hours_Analysis_September_2025.xlsx
├── CSV_Archive/               ← **PROCESSED CSV FILES GO HERE**
│   ├── Company_July_2025.csv
│   └── Working_Hours_August_2025.csv
├── app/                       ← Application folder
│   ├── simple_excel_processor.py
│   ├── gui_app.py
│   └── requirements_excel.txt
├── start_excel.bat            ← **START FILE - CLICK HERE!**
└── README.md                  ← This guide
```

## 🚀 It's that simple!

### **Step 1: Place CSV files**
- Place your CSV files in the `CSV_Input` folder
- The program automatically recognizes the month from the filename

### **Step 2: Start the application**
- **Double-click on `start_excel.bat`**
- Choose between GUI version (recommended) or command line version

### **Step 3: Done!**
- Excel files appear in the `Excel_Output` folder
- Processed CSV files are automatically moved to `CSV_Archive`

## 🖥️ GUI Version (Recommended)

The **GUI version** offers a user-friendly interface:

1. **Start:** Double-click on `start_excel.bat` → Choose option "1"
2. **Overview:** All folders and CSV files are displayed
3. **Processing:** Simply click "Process"
4. **Status:** Progress is shown live
5. **Open folders:** Direct links to all folders

### **GUI Advantages:**
- ✅ **User-friendly** - No command line needed
- ✅ **Clear overview** - All files at a glance
- ✅ **Safe** - Automatic folder creation
- ✅ **Fast** - One click for everything

## 📋 Command Line Version

For experienced users:

1. **Start:** Double-click on `start_excel.bat` → Choose option "2"
2. **Select file:** Enter the number of the desired CSV file
3. **Processing:** Automatic Excel creation
4. **Done:** Excel file in the `Excel_Output` folder

## 🔧 Automatic Month Recognition

The program automatically recognizes the month from the filename:

- ✅ `Company_July_2025.csv` → **July 2025**
- ✅ `Working_Hours_08_2025.csv` → **August 2025**
- ✅ `Time_Tracking_September_2025.csv` → **September 2025**
- ✅ `Working_Hours_12_2024.csv` → **December 2024**

## 📊 Excel Output

Each Excel file contains:

### **📋 Raw Data Sheet**
- All original data, beautifully formatted
- Clear table with all details

### **👥 Employee Overview**
- Sum of working hours per employee
- Total amount per employee
- Average hourly rate

### **📅 Daily Overview**
- Number of employees per day
- Total hours per day
- Total amount per day
- Average hourly rate per day

### **📈 Monthly Overview**
- Total hours of the month
- Total amount of the month
- Average hours per day
- Average hourly rate

### **👤 Individual Employee Sheets**
- A separate worksheet for each employee
- Detailed breakdown of all working days
- Totals and averages per employee

## 🗂️ Automatic Archiving

- **Processed CSV files** are automatically moved to the `CSV_Archive` folder
- **No duplicates** - If same filenames exist, a timestamp is added
- **Clean separation** - Input and archive are separated

## 🛠️ What's in this folder?

- **`start_excel.bat`** - Main start file (double-click to start)
- **`app/simple_excel_processor.py`** - Core processing (Python script)
- **`app/gui_app.py`** - User interface (GUI version)
- **`app/requirements_excel.txt`** - Required Python libraries
- **`README.md`** - This guide

## 🚀 Easy Start

1. **Double-click on `start_excel.bat`**
2. **Choose "1" for GUI version** (recommended)
3. **Place CSV files in `CSV_Input`**
4. **Click "Process"**
5. **Done!** Find Excel files in `Excel_Output`

## 🔧 System Requirements

- **Windows** (tested on Windows 10/11)
- **Python 3.7+** (automatically installed)
- **Internet connection** (only for first installation)

## ❓ Frequently Asked Questions

### **"No CSV files found!"**
- Place CSV files in the `CSV_Input` folder
- Make sure the files have the `.csv` extension

### **"Error installing dependencies!"**
- Make sure you have an internet connection
- Run `start_excel.bat` as administrator

### **"Excel file is not created!"**
- Check if the CSV file has the correct format
- Look in the `CSV_Input` folder for error messages

### **"Can I process multiple months at once?"**
- Yes! Simply place all CSV files in the `CSV_Input` folder
- The program creates a separate Excel file for each month

## 📞 Support

For problems or questions:
1. Check this guide
2. Look at the application's log output
3. Make sure all folders are created correctly

---

**🎉 Good luck with your working hours processing!**
