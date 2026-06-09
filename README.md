# PDF to Excel Bank Statement Converter
### The tool big finance forgot to build — works on Windows, macOS, and Linux

Banks have mastered a special kind of digital irony: they'll happily hand you statements from 7+ years ago as PDFs, but ask for that same data in Excel and suddenly it's like you requested nuclear launch codes. This project was born from that exact frustration during tax season — when I found that even expensive "professional" tools like Adobe Acrobat Pro and Nitro PDF utterly failed at the seemingly simple task of turning bank statements into usable spreadsheets.

So, with no viable paid solution, I wrote my own Python script to do the job — and it works **flawlessly**.

<p align="center">
  <img src="https://github.com/cloudpotions/PDF-to-Excel-Bank-Statements-Converter/raw/main/PDF2Excel.jpg" alt="PDF to Excel Bank Statement Converter">
</p>

This script beats OCR-based tools because it uses targeted libraries — **pdfplumber, pandas, and openpyxl** — to read the PDF's actual text structure instead of "looking at" the page. Adobe-style OCR can misread characters and scramble the order of transactions; this script extracts data directly, preserving precision **and** the original sequence. It also batch-processes a whole folder at once, writes a verification sheet, and prints financial summaries.

**Universal compatibility:** the script targets **Chase** checking statements, but the framework adapts easily to other banks. Hand this script plus a sample statement to a capable AI and it can retarget the parser for you — see [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak). I'm also available for consulting — https://www.linkedin.com/in/ellis-jesse/

⭐ If this script saved you time, please give the repo a Star! ⭐

---

## Contents
- [Why this script rocks](#why-this-script-rocks)
- [Quick start](#quick-start)
- [How it works](#how-it-works)
- [Verify your results](#verify-your-results-please-actually-do-this)
- [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak)
- [Detailed setup](#detailed-setup)
- [Troubleshooting](#troubleshooting)

---

## Why this script rocks

1. **Accurate ordering** — data comes out **exactly as it appears in the PDF**, so you can lay the PDF and Excel side by side and check them line for line.
2. **Handles large volumes** — tested on 12 months at a time; it'll happily chew through years of statements.
3. **Works with any bank or card** — built and tested for Chase, but it's designed to be the *foundation* for any statement. Hand it to any AI with a sample of your PDF and it adapts in about 5 minutes — see [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak).
4. **Free, open source, and 100% local** — pure Python, no cloud, no fees, no tokens, no data ever leaving your machine. Run it on a junky spare laptop if you want.

---

## Quick start

**1. Install Python 3** (if you don't have it) — see [Detailed setup](#detailed-setup) for your OS.

**2. Install the libraries:**
```
pip install pdfplumber pandas openpyxl
```

**3. Make a folder and put your statement PDFs in it.** Create a new folder (e.g. `Chase Statements 2026`) and put **only** the PDF statements you want to convert inside it.

> ⚠️ **There is no picking individual files.** You point the tool at **one folder**, and it automatically converts **every single PDF inside that folder.** So keep that folder clean — put nothing in it except the statements you want converted. (Don't aim it at something like your whole `Downloads` folder.)

**4. Run the script.** First save `pdf-2-excel.py` somewhere easy to find — the steps below assume your **Downloads** folder. Then pick whichever way is easier:

<details>
<summary><b>▶ Easiest — just double-click it (and how to fix it if it won't run)</b></summary>

Most of the time you can just **double-click `pdf-2-excel.py`** and a welcome window + folder picker appear. If it instead opens in a text editor, or nothing happens, give the file permission to run:

- **Windows:** double-click usually works as long as Python was installed with **"Add Python to PATH"** checked. If it opens in an editor, right-click the file → **Open with → Python**.
- **macOS:** right-click the file → **Open With → Python Launcher**. The first time, macOS may warn it's from an unidentified developer — right-click → **Open**, then confirm **Open**. To mark it runnable, open **Terminal** and run:
  ```
  chmod +x ~/Downloads/pdf-2-excel.py
  ```
- **Linux:** make it executable first — either right-click → **Properties → Permissions → ✅ "Allow executing file as program"**, or in a terminal run:
  ```
  chmod +x ~/Downloads/pdf-2-excel.py
  ```
  Then double-click and choose **Run** if your file manager asks "Run" vs "Display".

</details>

<details>
<summary><b>▶ Reliable — run it from the Terminal / Command Prompt (always works)</b></summary>

If double-clicking is being fussy, this way is rock-solid. Two short steps:

*Step A — go to the folder where you saved the script:*
```
# Linux or macOS
cd ~/Downloads

# Windows (Command Prompt)
cd %USERPROFILE%\Downloads
```

*Step B — start the script:*
```
# Linux or macOS
python3 pdf-2-excel.py

# Windows
python pdf-2-excel.py
```

A folder picker then pops up — **select the folder you made in step 3** (the one holding your PDFs).

> 💡 **Optional shortcut:** skip the picker by typing the folder's path right after the command, e.g.
> `python3 pdf-2-excel.py "/home/you/Chase Statements 2026"`

</details>

When it finishes, the Excel file is saved **inside that same folder, next to your PDFs**, named `FolderName_transactions_<date>_<time>.xlsx`.

---

## How it works

This is a **GUI tool** — no programming required.

1. **You choose one folder** (a folder picker pops up — you select the *folder*, not individual files). Or pass the folder's path on the command line to skip the picker.
2. The script then **automatically converts every PDF inside that folder**, extracts the transactions, and **saves one combined Excel file right next to your PDFs.**
3. A summary box tells you how many statements and transactions were processed.

<details>
<summary><b>📊 What's in the output spreadsheet (columns)</b></summary>

| Column | Meaning |
|---|---|
| `Statement_Date` | The statement the row came from |
| `Date` | Transaction date (year inferred correctly across Dec/Jan boundaries) |
| `Description` | Merchant / payee text |
| `Amount` | Positive = credit, negative = debit |
| `Balance` | Running balance printed on the statement |
| `Original_Line` | The raw line as extracted (for auditing) |
| `Source_File` | Which PDF the row came from |
| `Source_File_Page_Number` | The page **inside that PDF** where the row appears (1-based, matches your PDF viewer) — so you can jump straight to it |

A second **`Verification`** sheet keeps the original line — plus the source file and page number — next to each parsed amount, so you can audit and locate any row fast.

</details>

---

## Verify your results (please — actually do this)

The script preserves exact PDF order, but you should still review. My method for safe, double-entry checking:

1. Open the PDF and the generated Excel **side by side**.
2. Review **page by page** (Jan 2026, then Feb 2026, …).
3. Tick off each confirmed section with a pen on a physical notepad.
4. When you finish — **do it again.** Don't be lazy. The script does the heavy lifting (I've yet to see it miss in testing), but a second human pass is cheap insurance.

Happy converting! 🚀

---

## Use it with ANY bank or credit card (≈5-minute AI tweak)

This script was tested and works **perfectly on Chase checking-account statements.** If you feed it a **different bank**, or a **credit-card statement**, it most likely **won't work as-is** — the layout, column names, and date formats are different.

That's expected, and it's not a problem. **This script is the foundation for converting any statement, and adapting it takes about 5 minutes with any AI** (ChatGPT, Claude, Gemini, Copilot, etc.). You don't need to know how to code.

**Here's all you do:**

1. **Share your statement PDF with the AI — the whole file is best.** Pages often differ (the first page may have a summary, the middle pages hold the transactions, the last pages can be legal fine print), so don't trim it to a single page. **You do not need to cut or edit the PDF.**
2. **If your PDF is huge (say 100+ pages), you still don't have to edit it.** The prompts below tell the AI to analyze the structure itself and focus only on the pages that matter. The *only* time to trim is if your AI rejects the upload for being too large — then upload roughly the **first 10–15 pages**, which almost always covers the full repeating pattern.
3. Copy the prompt that matches how you use AI (below), paste in the full `pdf-2-excel.py` script where shown, and give the AI your PDF.
4. Run the modified script it gives you. Done.

### 📋 Pick the prompt that matches your setup

Both prompts ask for the same safety guarantees. The difference: the **terminal** AI can run and test the code itself; with the **browser** one, you run it and report back.

<details>
<summary><b>🟢 Option 1 — Browser AI, no code access (ChatGPT, Claude, Gemini in a web browser) — start here if you're new</b></summary>

Use this if you're chatting with an AI in your web browser that can read your attached PDF but **cannot run code**. It writes the script; you run it and tell it the result. Click the copy icon in the top-right of the box.

```text
I have a working Python script (pasted below) that converts CHASE checking-account
PDF statements into an Excel spreadsheet. How it works: it reads the PDF text with
pdfplumber page by page, finds the "TRANSACTION DETAIL" section, parses every line
that starts with a date (MM/DD) into Date, Description, Amount and a running Balance,
records which page each row came from, and writes everything to Excel in the EXACT
order it appears on the statement.

I want to use it on a DIFFERENT document instead: [describe yours — e.g. "a Bank of
America checking statement" or "an American Express credit-card statement"]. I've
attached the PDF. You can't run code here, so READ the attached PDF carefully and
modify the script for me; I'll run it on my own computer and tell you the result.

Work through these requirements IN ORDER and don't skip any:

1. STUDY THE PDF STRUCTURE FIRST. Read the attached statement and tell me, before
   writing any code: which pages contain the actual transactions; what a transaction
   row looks like (date format, column order, where the amount and any balance sit);
   and whether some pages have a DIFFERENT layout — a cover/summary page, legal pages
   with no transactions, or a format that changes partway through. If the document is
   long and clearly repeats the same pattern, say so and base the parser on that
   pattern plus any one-off pages.

2. PRESERVE EXACT ORDER. Keep transactions in the order they appear on the statement
   (tag each with a within-statement sequence number and sort by date + that sequence).
   I verify by laying the PDF and Excel side by side, line for line, so the order must
   match the PDF exactly. Don't re-sort in a way that breaks the original sequence.

3. NEVER SILENTLY DROP A TRANSACTION. The original Chase version had a bug where rows
   at the bottom/top of a page were lost: the PDF text layer glued a page marker
   ("*start*/*end*transaction detail") onto the front of a real transaction line, so it
   no longer started with a date and got skipped. We fixed it by stripping that marker.
   Watch for the EQUIVALENT traps in my statement — repeated page headers/footers,
   "continued" markers, subtotal/summary rows, wrapped multi-line descriptions, or any
   oddly-formatted page — and make sure no real transaction is ever dropped.

4. MAP MY COLUMNS. Tell me what columns my statement has and map them correctly (some
   banks say "Posting Date" not "Date", some split money into separate "Debits" and
   "Credits" columns instead of one signed Amount). Keep the page-number column so I
   can still find each row back in the PDF.

5. CREDIT CARDS WORK DIFFERENTLY. If this is a credit card there's usually NO running
   balance per line — instead Previous Balance + Purchases − Payments/Credits = New
   Balance. Adapt the parsing and the output columns to that model.

6. BUILD IN SELF-CHECKS so the script tells ME if something's wrong when I run it:
   (a) reconcile against the statement's printed totals ("Beginning + transactions =
   Ending Balance", or "Previous Balance + charges − payments = New Balance") and print
   a clear warning if it doesn't match; and (b) print a warning for any page that
   should have held transactions but matched none.

Give me the COMPLETE modified script plus a short plain-English list of what you
changed. After I run it, I'll paste back the script's reconciliation output (and any
rows that look missing) so we can fix it together.

--- SCRIPT START ---
[paste the full contents of pdf-2-excel.py here]
--- SCRIPT END ---
```

</details>

<details>
<summary><b>🔵 Option 2 — Terminal / agent AI that can run code (Claude Code, Cursor, Copilot agent, etc.) — best results</b></summary>

Use this if your AI can actually run Python on your machine. It will analyze your PDF, build the parser, and **test it until it reconciles** before handing it back. Click the copy icon in the top-right of the box.

```text
I have a working Python script (pasted below) that converts CHASE checking-account
PDF statements into an Excel spreadsheet. How it works: it reads the PDF text with
pdfplumber page by page, finds the "TRANSACTION DETAIL" section, parses every line
that starts with a date (MM/DD) into Date, Description, Amount and a running Balance,
records which page each row came from, and writes everything to Excel in the EXACT
order it appears on the statement.

I want to use it on a DIFFERENT document instead: [describe yours — e.g. "a Bank of
America checking statement" or "an American Express credit-card statement"]. The PDF
is at: [path/to/your/statement.pdf]. You can run code on my machine, so analyze,
modify, AND TEST it before handing it back.

Work through these requirements IN ORDER and don't skip any:

1. ANALYZE THE PDF STRUCTURE FIRST — PROGRAMMATICALLY, NOT BY EYE. Before you change
   the parser, write and run a short pdfplumber diagnostic that, for EACH page, prints
   a quick fingerprint: page number, total line count, the top/header line, any
   section markers, and how many lines look like real transactions (start with a date
   and/or end in a money amount). Use that fingerprint to:
     (a) find which pages actually contain transactions;
     (b) flag pages whose layout DIFFERS — a cover/summary page, legal/disclosure
         pages with no transactions, or a format that changes partway through;
     (c) if the document is long, detect the REPEATING pattern (e.g. "every
         transaction page looks like this") so you do NOT have to read every page by
         hand — deeply analyze a representative sample plus any one-off pages.
   Show me a short summary of what you found before writing the parser. (This means I
   can hand you a 100-page PDF and you figure out the structure yourself.)

2. PRESERVE EXACT ORDER. The script intentionally keeps transactions in the order they
   appear on the statement (it tags each with a within-statement sequence number and
   sorts by date + that sequence). I verify results by laying the PDF and the Excel
   side by side, line for line, so the order must match the PDF exactly. Don't silently
   re-sort in a way that breaks the original sequence.

3. NEVER SILENTLY DROP A TRANSACTION. The original Chase version had a bug where rows
   at the bottom/top of a page were lost: the PDF text layer glued a page marker
   ("*start*/*end*transaction detail") onto the front of a real transaction line, so it
   no longer started with a date and the parser skipped it. We fixed it by stripping
   that marker (and restoring a digit it had eaten). Look for the EQUIVALENT traps in
   my statement — repeated page headers/footers, "continued" markers, subtotal/summary
   rows, multi-line descriptions that wrap, or any page whose layout differs from the
   rest — and make sure no real transaction is ever dropped.

4. MAP MY COLUMNS. Tell me what columns my statement actually has and map them
   correctly. Banks differ: some say "Posting Date" not "Date", some split money into
   separate "Debits" and "Credits" columns instead of one signed Amount, etc. Keep the
   page-number column so I can still find each row back in the original PDF.

5. CREDIT CARDS WORK DIFFERENTLY. If this is a credit-card statement there is usually
   NO running balance per line — instead there's a Previous Balance, then Purchases/
   Charges, Payments and Credits that net to a New Balance. Adapt the parsing and the
   output columns to fit that model.

6. ADD SELF-CHECKS (this is the safety net). After parsing: (a) reconcile against the
   statement's own printed totals — "Beginning Balance + sum of transactions = Ending
   Balance" for checking, or "Previous Balance + charges − payments/credits = New
   Balance" for a credit card — and print a clear warning if it doesn't reconcile; and
   (b) print a warning for any page that looked like it should hold transactions but
   matched none, so a layout change can never silently swallow rows.

7. TEST IT before handing it over: run the modified script on my actual PDF, confirm it
   reconciles to the printed totals, and fix anything that doesn't. Then show me, in
   this order: your structure summary from step 1, the COMPLETE final script, and a
   plain-English list of exactly what you changed and why.

--- SCRIPT START ---
[paste the full contents of pdf-2-excel.py here]
--- SCRIPT END ---
```

</details>

> **Tip:** if the first result misses a few rows, tell the AI *"these specific transactions are missing: …"* and paste the offending lines. The two self-checks (totals reconciliation + empty-page warning) will usually flag them for you automatically.

---

## Detailed setup

You need **Python 3** plus the libraries **pdfplumber, pandas, and openpyxl**. Open the section for your operating system:

<details>
<summary><b>🪟 Windows — setup</b></summary>

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

</details>

<details>
<summary><b>🍎 macOS — setup</b></summary>

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

</details>

<details>
<summary><b>🐧 Linux — setup</b></summary>

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

**Quick alternative (no virtual environment):** install straight into the system Python with `--break-system-packages`. It's faster but bypasses a safety guard, so the venv route above is preferred:
```
pip install --break-system-packages pdfplumber pandas openpyxl
python3 pdf-2-excel.py
```

</details>

---

## Troubleshooting

<details>
<summary><b>Common problems &amp; fixes</b></summary>

| Problem | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'pdfplumber'` (or pandas/openpyxl) | Run the `pip install` step. On Linux, make sure your venv is **activated**, or use the `--break-system-packages` one-liner. |
| `No module named 'tkinter'` / no window appears (Linux) | Install Tk: `sudo apt install python3-tk` (Debian/Ubuntu) or `sudo dnf install python3-tkinter` (Fedora). |
| The Linux folder picker feels clunky | Install `zenity` for the native GNOME picker (see Linux setup), or just pass the folder path on the command line. |
| Running over SSH / no display | Pass the folder path directly: `python3 pdf-2-excel.py "/path/to/folder"` — it runs fully in the terminal, no GUI needed. |
| `command not found: python3` (Windows) | Use `python` instead of `python3` on Windows, and make sure "Add Python to PATH" was checked at install. |
| No transactions found | Confirm the PDFs are genuine **Chase** statements with a "TRANSACTION DETAIL" section. Other banks need the parser retargeted — see [Use it with ANY bank or credit card](#use-it-with-any-bank-or-credit-card-5-minute-ai-tweak). |

</details>

---

*Tip: inside an activated virtual environment, `python` and `python3` point to the same Python 3, so either works there. Outside a venv, prefer `python3` on macOS/Linux and `python` on Windows.*
