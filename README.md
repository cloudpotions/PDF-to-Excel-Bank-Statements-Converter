# PDF to Excel Bank Statement Converter
### The tool big finance forgot to build — works on Windows, macOS, and Linux

Banks have mastered a special kind of digital irony: they'll happily hand you statements from 7+ years ago as PDFs, but ask for that same data in Excel and suddenly it's like you requested nuclear launch codes. This project was born from that exact frustration during tax season — when I found that even expensive "professional" tools like Adobe Acrobat Pro and Nitro PDF utterly failed at the seemingly simple task of turning bank statements into usable spreadsheets.

So, with no viable paid solution, I wrote my own Python script to do the job — and it works **flawlessly**.

<p align="center">
  <img src="https://github.com/cloudpotions/PDF-to-Excel-Bank-Statements-Converter/raw/main/PDF2Excel.jpg" alt="PDF to Excel Bank Statement Converter">
</p>

This script beats OCR-based tools because it uses targeted libraries — **pdfplumber, pandas, and openpyxl** — to read the PDF's actual text structure instead of "looking at" the page. Adobe-style OCR can misread characters and scramble the order of transactions; this script extracts data directly, preserving precision **and** the original sequence. It also batch-processes a whole folder at once, writes a verification sheet, and prints financial summaries.

**Universal compatibility:** the script targets **Chase** checking statements, but the framework adapts easily to other banks. Hand this script plus one sample statement (or a screenshot) to a capable AI and it can retarget the parser for you. I'm also available for consulting — https://www.linkedin.com/in/ellis-jesse/

⭐ If this script saved you time, please give the repo a Star! ⭐

---

## Contents
- [Why this script rocks](#why-this-script-rocks)
- [Quick start](#quick-start)
- [How it works](#how-it-works)
- [Output columns](#output-columns)
- [Verify your results](#verify-your-results-please-actually-do-this)
- [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak)
- [Detailed setup](#detailed-setup) — [Windows](#windows) · [macOS](#macos) · [Linux](#linux)
- [Troubleshooting](#troubleshooting)

---

## Why this script rocks

1. **Accurate ordering** — data comes out **exactly as it appears in the PDF**, so you can lay the PDF and Excel side by side and check them line for line.
2. **Handles large volumes** — tested on 12 months at a time; it'll happily chew through years of statements.
3. **Works with any bank or card** — built and tested for Chase, but it's designed to be the *foundation* for any statement. Hand it to any AI with a sample of your PDF and it adapts in about 5 minutes — see [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak) for a ready-made prompt.
4. **Free, open source, and 100% local** — pure Python, no cloud, no fees, no tokens, no data ever leaving your machine. Run it on a junky spare laptop if you want.

---

## Quick start

**1. Install Python 3** (if you don't have it) — see [Detailed setup](#detailed-setup) for your OS.

**2. Install the libraries:**
```
pip install pdfplumber pandas openpyxl
```

**3. Put your statements in their own folder** (e.g. `Chase Statements 2024`). The script processes **every PDF in the folder you choose**, so don't run it on a junk-drawer folder like `Downloads`.

**4. Run it** — two equally good ways:

```
# A) Graphical: opens a folder picker
python3 pdf-2-excel.py

# B) Direct: hand it the folder and skip the picker (fastest)
python3 pdf-2-excel.py "/path/to/Chase Statements 2024"
```

The Excel file lands in **the same folder as your PDFs**, named `FolderName_transactions_<date>_<time>.xlsx`.

> On Windows you can also just **double-click `pdf-2-excel.py`**.

---

## How it works

This is a **GUI tool** — no programming required.

1. **Pick the folder** holding your PDFs (a file picker pops up), or pass the path on the command line to skip the picker.
2. The script reads every PDF, extracts the transactions, and **saves the Excel right next to your PDFs.**
3. A summary box tells you how many statements and transactions were processed.

### Output columns
| Column | Meaning |
|---|---|
| `Statement_Date` | The statement the row came from |
| `Date` | Transaction date (year inferred correctly across Dec/Jan boundaries) |
| `Description` | Merchant / payee text |
| `Amount` | Positive = credit, negative = debit |
| `Balance` | Running balance printed on the statement |
| `Original_Line` | The raw line as extracted (for auditing) |
| `Source_File` | Which PDF the row came from |

A second **`Verification`** sheet keeps the original line next to each parsed amount so you can audit fast.

---

## Verify your results (please — actually do this)

The script preserves exact PDF order, but you should still review. My method for safe, double-entry checking:

1. Open the PDF and the generated Excel **side by side**.
2. Review **page by page** (Jan 2024, then Feb 2024, …).
3. Tick off each confirmed section with a pen on a physical notepad.
4. When you finish — **do it again.** Don't be lazy. The script does the heavy lifting (I've yet to see it miss in testing), but a second human pass is cheap insurance.

Happy converting! 🚀

---

## Use it with ANY bank or credit card (≈5-minute AI tweak)

This script was tested and works **perfectly on Chase checking-account statements.** If you feed it a **different bank**, or a **credit-card statement**, it most likely **won't work as-is** — the layout, column names, and date formats are different.

That's expected, and it's not a problem. **This script is the foundation for converting any statement, and adapting it takes about 5 minutes with any AI** (ChatGPT, Claude, Gemini, Copilot, etc.). You don't need to know how to code.

**Here's all you do:**

1. Grab **one sample PDF** of the statement you want to convert (one page is enough).
2. Open any AI chat that can read files/images.
3. **Paste in the full `pdf-2-excel.py` script**, attach (or paste a screenshot of) your sample statement, and then paste the prompt below.
4. Run the modified script it gives you. Done.

### 📋 Copy-paste this prompt

```text
I have a working Python script (pasted below) that converts CHASE checking-account
PDF statements into an Excel spreadsheet. How it works: it uses pdfplumber to read
the PDF text, finds the "TRANSACTION DETAIL" section, and parses every line that
starts with a date (MM/DD) into Date, Description, Amount, and a running Balance —
then writes them to Excel in the EXACT order they appear on the statement.

I want to use it on a DIFFERENT document instead: [describe yours — e.g. "a Bank of
America checking statement" or "an American Express credit-card statement"]. I've
attached a sample. Its layout, headers, column names, and date format are probably
different from Chase. Please modify the script to work with my statement.

These requirements are critical — read them carefully:

1. PRESERVE EXACT ORDER. The script intentionally keeps transactions in the order
   they appear on the statement (it tags each with a within-statement sequence
   number and sorts by date + that sequence). I verify results by laying the PDF and
   the Excel side by side, line for line, so the order must match the PDF exactly.
   Do not silently re-sort in a way that breaks the original sequence.

2. DON'T LOSE TRANSACTIONS AT PAGE BREAKS. The original Chase version had a bug where
   transactions at the bottom/top of a page were silently dropped: the PDF text layer
   glued a page marker ("*start*/*end*transaction detail") onto the front of a real
   transaction line, so the line no longer started with a date and the parser skipped
   it. We fixed it by stripping that marker (and restoring a digit it had eaten)
   before parsing. Look for the EQUIVALENT trap in my statement — repeated page
   headers/footers, "continued" markers, summary/subtotal rows, or multi-line
   descriptions that wrap onto a second line — and make sure NO real transaction is
   ever silently lost.

3. MAP MY COLUMNS. Tell me what columns my statement actually has and map them
   correctly. Banks differ: some call it "Posting Date" not "Date", some split money
   into separate "Debits" and "Credits" columns instead of one signed Amount, etc.

4. CREDIT CARDS WORK DIFFERENTLY. If this is a credit-card statement, there is usually
   NO running balance on each line — instead there's a Previous Balance, then
   Purchases/Charges, Payments, and Credits that net to a New Balance. Adapt the
   parsing and the output columns to fit that model.

5. ADD A SELF-CHECK. After parsing, reconcile against the statement's own printed
   totals — e.g. "Beginning Balance + sum of transactions = Ending Balance" for
   checking, or "Previous Balance + charges − payments/credits = New Balance" for a
   credit card — and print a clear warning if it doesn't reconcile. This is how we
   catch dropped or duplicated lines automatically.

Then give me the COMPLETE modified script, and a short plain-English summary of
exactly what you changed and why.

--- SCRIPT START ---
[paste the full contents of pdf-2-excel.py here]
--- SCRIPT END ---
```

> **Tip:** if the first result misses a few rows, tell the AI *"these specific
> transactions are missing: …"* and paste the offending lines. The reconciliation
> self-check from requirement #5 will usually flag them for you automatically.

---

## Detailed setup

You must have **Python 3** plus the libraries **pdfplumber, pandas, and openpyxl**. Below are step-by-step instructions for each OS for folks newer to Python.

### Windows

1. **Install Python**
   - Download from [python.org](https://www.python.org/downloads/) and run the installer.
   - ✅ **Check "Add Python to PATH"** during install (important!), then click **Install Now**.
2. **Install the libraries** — open **Command Prompt** (search "cmd") and run:
   ```
   pip install pdfplumber pandas openpyxl
   ```
3. **Run the script**
   - Download `pdf-2-excel.py` from this repo.
   - **Double-click it**, or from Command Prompt:
     ```
     python pdf-2-excel.py
     ```

### macOS

1. **Install Python** from [python.org](https://www.python.org/downloads/) (run the installer package).
2. **Install the libraries** — open **Terminal** (Applications → Utilities → Terminal):
   ```
   pip3 install pdfplumber pandas openpyxl
   ```
3. **Run the script**
   ```
   cd ~/Downloads        # or wherever you saved it
   python3 pdf-2-excel.py
   ```

### Linux

Linux needs Tk for the GUI, and (optionally) `zenity` for a nicer native folder picker.

**1. Install Python, Tk, and the picker:**
```
# Debian / Ubuntu
sudo apt update
sudo apt install python3 python3-pip python3-venv python3-tk zenity

# Fedora / RHEL
sudo dnf install python3 python3-pip python3-tkinter zenity
```
> `zenity` is optional but recommended — with it, the folder picker is the clean native GNOME/GTK one instead of Tk's clunky built-in dialog. Without it, the script falls back to the Tk picker automatically.

**2. Install the libraries — recommended: a virtual environment** (keeps your system Python clean):
```
python3 -m venv chase_env
source chase_env/bin/activate          # your prompt now shows (chase_env)
pip install pdfplumber pandas openpyxl
```

**3. Run it** (while the venv is active):
```
python3 pdf-2-excel.py                          # opens the folder picker
# or, skip the picker entirely:
python3 pdf-2-excel.py "/path/to/your/folder"
```

**4. When you're done**, leave the virtual environment:
```
deactivate
```

<details>
<summary><b>Prefer not to use a virtual environment? (quick one-liner)</b></summary>

You can install the libraries straight into the system Python with the `--break-system-packages` flag. It's quicker but bypasses a safety guard, so the venv route above is preferred:
```
pip install --break-system-packages pdfplumber pandas openpyxl
python3 pdf-2-excel.py
```
</details>

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'pdfplumber'` (or pandas/openpyxl) | Run the `pip install` step. On Linux, make sure your venv is **activated**, or use the `--break-system-packages` one-liner. |
| `No module named 'tkinter'` / no window appears (Linux) | Install Tk: `sudo apt install python3-tk` (Debian/Ubuntu) or `sudo dnf install python3-tkinter` (Fedora). |
| The Linux folder picker feels clunky | Install `zenity` for the native GNOME picker (see Linux step 1), or just pass the folder path on the command line. |
| Running over SSH / no display | Pass the folder path directly: `python3 pdf-2-excel.py "/path/to/folder"` — it runs fully in the terminal, no GUI needed. |
| `command not found: python3` (Windows) | Use `python` instead of `python3` on Windows, and make sure "Add Python to PATH" was checked at install. |
| No transactions found | Confirm the PDFs are genuine **Chase** statements with a "TRANSACTION DETAIL" section. Other banks need the parser retargeted (see "Customizable" above). |

---

*Tip: inside an activated virtual environment, `python` and `python3` point to the same Python 3, so either works there. Outside a venv, prefer `python3` on macOS/Linux and `python` on Windows.*
