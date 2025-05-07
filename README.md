PDF to Excel Bank Statement Converter: The Tool Big Finance Forgot to Build (works on Linux / Mac OS and Windows) 

Please give me a Star on Github if you used this script! ⭐ Support This Project

This project was born out of necessity. I was doing my taxes and, to my surprise Adobe Acrobat Pro, Nitro PDF and other paid solutions could not properly convert my Bank PDF Statements into excel files properly 

So, I decided to create my own Python script. To my amazement, it worked **flawlessly** and did what Big Tech has failed to do. 

![PDF to Excel Bank Statement Converter](https://github.com/cloudpotions/PDF-to-Excel-Bank-Statements-Converter/raw/main/PDF2Excel.jpg)


This Python script is a lightning-fast, precision-engineered GUI tool that transforms cluttered bank PDF statements into beautifully organized Excel spreadsheets. Unlike Adobe Acrobat Pro—which crawls along using error-prone OCR technology that mangles your financial data. These enterprise solutions choked on simple bank PDFs, taking 15+ minutes to process what this script handles in seconds. While Fortune 500 companies were busy building AI assistants and subscription models, they somehow missed building a tool that actually works for this essential financial task.

The script harnesses the power of pdfplumber and pandas—specialized libraries that run circles around commercial OCR solutions—to process thousands of transactions with 99% accuracy in under a minute. It's been battle-tested with Chase Bank statements but adapts easily to other financial institutions; simply share your bank's PDF format with Claude 3.7 Sonnet or another AI assistant, and be sure to tell your Ai what columns and format your Bank PDF has and it will be able to easily modify the python script for you (I tested on a couple other Bank PDF Statements using Claude 3.7 Sonnet and it was able to modify on the first try). 

So if you're like me, trapped in the PDF banking purgatory with years of statements that refuse to cooperate with modern financial tools, you've finally found your escape route...

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

1. **Select PDF Folder**: A pop-up window will ask you to select a folder with all the PDF documents you wish to process. It will process ALL PDF in that folder, so make sure you do not run it in Downloads folder - Save your PDF files to a new Folder (For example, Chase Bank Statements 2024 Folder)

2. **Output Location**: The output file will be saved in the same folder that the PDFs are located in Step 1

 The script extracts transaction data from the PDFs and organizes it into a clean Excel spreadsheet with columns like:

   - Date of Transaction
   - Statement Date 
   - Description
   - Amount
   - Balance
   - Original Line
   - Source File

When the script is complete, the pop-up box will give you some statistics like how many files were processed and lines were added to the XLSX excel file. It will save the file with a naming convention of FolderName/Date/Time.xlsx

---

Special Note: You must have Python installed on your OS along with the libraries pdfplumber, pandas, and openpyxl. To use the script, download PDF-2-Excel.py from this repository and run the script and the Gui will do the rest! I am also going to include very detailed Installation instruction at the bottom of this ReadME file for newbies that do not know how to run Python Files with dependencies... 


Even though the script maintains exact PDF order, manual review is essential! 

My suggestion for Double Entry Accounting with my script: 

1. Open the PDF and the generated Excel side-by-side
2. Review page-by-page (e.g., January 2024, then February 2024,etc). 
3. Use a physical notepad with a pen to o to mark each confirmed section you have double checked.
4. Once you are done, repeat steps 1-3 AGAIN! Do not be lazy. The program does most of the work for you but it still needs human intervention. 

Happy converting! 🚀


While most developers visiting this repository can simply download and run the pdf-2-excel.py script, Cloud Potions believes in making technology accessible to everyone. If you're new to Python or programming in general, I've created below a comprehensive guide to run this script on Linux, Mac OS or Windows ... 

## Advanced Setup Instructions

### Windows Setup

1. **Install Python**
   - Download Python from [python.org](https://www.python.org/downloads/)
   - Run the installer
   - ✓ CHECK "Add Python to PATH" during installation (very important!)
   - Click "Install Now"

2. **Install Required Packages**
   - Open Command Prompt (search for "cmd" in Start menu)
   - Copy and paste this command:
     ```
     pip install pdfplumber pandas openpyxl
     ```
   - Press Enter and wait for installation to complete

3. **Run the Script**
   - Download `pdf-2-excel.py` from this repository
   - Double-click the Python script file to run it
   - If double-clicking doesn't work, right-click and select "Open with" → "Python"

### macOS Setup

1. **Install Python**
   - Download Python from [python.org](https://www.python.org/downloads/)
   - Run the installer package
   - Follow the installation instructions

2. **Install Required Packages**
   - Open Terminal (from Applications → Utilities → Terminal)
   - Copy and paste this command:
     ```
     pip3 install pdfplumber pandas openpyxl
     ```
   - Press Enter and wait for installation to complete

3. **Run the Script**
   - Download `pdf-2-excel.py` from this repository
   - Open Terminal and navigate to where you saved the file:
     ```
     cd ~/Downloads  (or wherever you saved it)
     ```
   - Make the script executable:
     ```
     chmod +x pdf-2-excel.py
     ```
   - Run the script:
     ```
     python3 pdf-2-excel.py
     ```

### Linux Setup

#### Option 1: Quick Setup (Easiest but Option 2 is recommended below as it is safer) 

This method works on most Linux systems but uses the `--break-system-packages` flag, which bypasses some system protections.

1. **Install Python and Required Tools**
   - Open Terminal and run:
     ```
     sudo apt update
     sudo apt install python3 python3-pip python3-tk
     ```
   - For Fedora/RHEL:
     ```
     sudo dnf install python3 python3-pip python3-tkinter
     ```

2. **Install Required Packages**
   - In Terminal, run:
     ```
     pip3 install --break-system-packages pdfplumber pandas openpyxl
     ```
   - **Note:** This flag overrides system package protections but is the simplest approach.

3. **Run the Script**
   - Download `pdf-2-excel.py` from this repository
   - Make the script executable:
     ```
     chmod +x pdf-2-excel.py
     ```
   - Run the script:
     ```
     python3 pdf-2-excel.py
     ```

#### Option 2: Using a Virtual Environment (Recommended)
This method is safer but requires a few more commands:

1. **Install Python and Required Tools**
   ```
   sudo apt update
   sudo apt install python3 python3-pip python3-venv python3-tk
   ```

2. **Create and Activate Virtual Environment**
   ```
   python3 -m venv chase_env
   source chase_env/bin/activate
   ```
   After running these commands, your terminal prompt should change to show "(chase_env)" at the beginning, indicating the virtual environment is active.

3. **Install Required Packages in the Virtual Environment**
   ```
   pip install pdfplumber pandas openpyxl
   ```

4. **Run the Script** 
   Make sure you're still in the activated virtual environment (you should see "(chase_env)" in your prompt) and in the same directory as the script. If you need to navigate to your downloads folder:
   ```
   cd ~/Downloads  # or wherever you saved the script
   python pdf-2-excel.py
   ```

5. **When Finished**
   To exit the virtual environment when you're done:
   ```
   deactivate
   ```
```

```markdown
