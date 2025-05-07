# PDF-to-Excel-Bank-Statements-Converter
This Python script offers a user-friendly GUI tool to convert bank PDF statements into organized Excel spreadsheets. It accurately extracts transaction data while preserving the original order from the PDF, making it ideal for individuals and businesses needing to process large volumes of financial data efficiently. 

# 🧪 PDF to Excel Converter: The Ultimate Tool for Bank Statements

## Introduction

This project was born out of necessity. I was doing my taxes and, to my surprise, **Chase Bank only provided PDF statements** for years of transactions. I had thousands of entries to process, and I needed them in Excel format. I tried **Adobe Acrobat Pro**, **Nitro PDF**, and several other paid solutions, but none of them worked properly. Adobe Pro, for example, took forever to process and spit out messy, unusable data. 

So, I decided to create my own Python script. To my amazement, it worked **flawlessly**. It processed **12 months of Chase Bank statements with thousands of transactions** in under a minute, and the output was perfectly organized. This tool saved me hours of manual work and helped me meet my tax deadline. 

Now, I’m sharing this tool with you. Whether you’re an individual or a company dealing with hundreds of thousands of transactions, this script will make your life easier. And the best part? It’s **free**, **fast**, and **better** than any paid solution for this task.

---

## Why This Script Rocks for converting Bank PDF Statements to Excel

1. **Accurate Sorting**: The script organizes the data **exactly as it appears in the PDF**. It doesn’t mess up the order or randomly sort by dates. This is crucial because when you’re done converting, you’ll want to compare the PDF and Excel side by side to ensure everything lines up perfectly.

2. **Handles Large Volumes**: While I tested it with 12 months of statements, this script can handle much more. Whether you have a year’s worth of data or several years, it will process everything efficiently.

3. **Customizable**: If you’re not using Chase Bank, you can easily tweak the script to work with other banks. Just provide the script to your favorite AI (like ChatGPT) along with screenshots of your bank’s PDF format and a description of the columns you need.

4. **Saves Time**: This script processes thousands of transactions in under a minute. Compare that to Adobe Pro or other paid tools, which are slower and less accurate.

5. **Free and Open Source**: Unlike expensive software, this tool leverages the power of Python and open-source libraries like `pdfplumber` and `pandas`. It’s completely free to use.

---

## How It Works

This script is a **GUI-based tool** (Graphical User Interface), so you don’t need to be a programmer to use it. Here’s what it does:

1. **Select PDF Files**: A pop-up window allows you to select multiple PDF statements at once.
2. **Choose Output Location**: You can specify where the Excel file should be saved.
3. **Extract and Organize Data**: The script extracts transaction data from the PDFs and organizes it into a clean Excel spreadsheet with columns like:
   - Date
   - Description
   - Amount
   - Balance
4. **Save the Output**: The Excel file is saved in the location you choose.

---

Special Note: You must have Python installed on your OS along with the libraries pdfplumber, pandas, and openpyxl. To use the script, download PDF-2-Excel.py from this repository and run the script and the Gui will do the rest!

For Newbies: If double clicking on the file does not open the file, go to your terminal and make sure you are in the same location as the script and run the script with the following command, for example if it is in your downloads folder, enter this in the terminal: 

cd Downloads
python3 PDF-2-Excel.py

Even though the script maintains exact PDF order, manual review is essential! 

My suggestion for Double Entry Accounting with my script: 

1. Open the PDF and the generated Excel side-by-side
2. Review page-by-page (e.g., January 2024, then February 2024,etc). 
3. Use a physical notepad with a pen to o to mark each confirmed section you have double checked.
4. Once you are done, repeat steps 1-3 AGAIN! Do not be lazy. The program does most of the work for you but it still needs human intervention. 

Happy converting! 🚀
