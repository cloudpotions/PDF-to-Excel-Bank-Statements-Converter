#!/usr/bin/env python3
"""
Chase Bank Statement Processor
Extracts transactions from Chase PDF statements and organizes them in Excel.
"""
import os
import re
import sys
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def main():
    # Create GUI root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Show welcome message with clear instructions
    messagebox.showinfo(
        "Chase Statement Processor",
        "This tool extracts transactions from Chase PDF statements and organizes them in Excel.\n\n"
        "INSTRUCTIONS:\n"
        "Select the folder containing your Chase PDF statements.\n"
        "All PDF files in the folder will be processed."
    )
    
    # Try to import the required libraries
    try:
        import pandas as pd
        import pdfplumber
    except ImportError as e:
        messagebox.showerror(
            "Missing Dependencies",
            f"Required libraries not found: {str(e)}\n\n"
            "Please install the required libraries with:\n"
            "pip install pdfplumber pandas openpyxl\n\n"
            "On Linux, you may need to add: --break-system-packages"
        )
        return
    
    # Show folder selection dialog
    chase_pdf_folder = filedialog.askdirectory(
        title="SELECT FOLDER CONTAINING CHASE PDF STATEMENTS"
    )
    
    if not chase_pdf_folder:
        messagebox.showinfo("Cancelled", "No folder was selected. Exiting.")
        return
    
    # Display selected folder
    print(f"\nSelected folder: {chase_pdf_folder}")
    
    # Check if the folder contains PDF files directly
    pdf_files = [f for f in os.listdir(chase_pdf_folder) if f.lower().endswith('.pdf')]
    
    # If no PDFs found, check if there are subfolders with PDFs
    if not pdf_files:
        # Look in immediate subfolders
        for subfolder in os.listdir(chase_pdf_folder):
            subfolder_path = os.path.join(chase_pdf_folder, subfolder)
            if os.path.isdir(subfolder_path):
                subfolder_pdfs = [f for f in os.listdir(subfolder_path) if f.lower().endswith('.pdf')]
                if subfolder_pdfs:
                    # If we find PDFs in a subfolder, use that folder instead
                    chase_pdf_folder = subfolder_path
                    pdf_files = subfolder_pdfs
                    print(f"Found PDFs in subfolder: {subfolder}")
                    print(f"Using folder: {chase_pdf_folder}")
                    break
    
    # If still no PDFs found, show error
    if not pdf_files:
        messagebox.showerror(
            "No PDFs Found", 
            f"No PDF files were found in the selected folder or immediate subfolders:\n{chase_pdf_folder}\n\n"
            "Please try again with a folder containing Chase PDF statements."
        )
        return
    
    # Your existing functions for processing
    def extract_statement_date(filename):
        """Extract date from filename for sorting"""
        date_match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
        if date_match:
            year, month, day = date_match.groups()
            return f"{year}-{month}-{day}"
        return filename

    def determine_transaction_year(transaction_date, statement_date):
        """
        Determine correct year for transaction based on statement date and transaction date 
        transaction_date: MM/DD format
        statement_date: YYYY-MM-DD format
        """
        trans_month = int(transaction_date.split('/')[0]) 
        statement_year = int(statement_date.split('-')[0])
        statement_month = int(statement_date.split('-')[1])
        
        # If statement is from January and transaction is from December
        if statement_month == 1 and trans_month == 12:
            return str(statement_year - 1)
        # If statement is from December and transaction is from January 
        elif statement_month == 12 and trans_month == 1:
            return str(statement_year + 1)
        return str(statement_year)

    def extract_chase_transactions(text, filename, statement_date):
        """Extract transactions from Chase bank statement text"""
        transactions = []
        in_transaction_section = False
        lines = text.split('\n')
        for line in lines:
            if "TRANSACTION DETAIL" in line:
                in_transaction_section = True
                continue
            if not in_transaction_section:
                continue
            if "Beginning Balance" in line:
                continue
                
            date_match = re.match(r'^(\d{2}/\d{2})', line.strip())
            if date_match:
                current_date = date_match.group(1)
                amounts = re.findall(r'-?[\d,]+\.\d{2}', line)
                if len(amounts) >= 2:
                    amount_str = amounts[-2]
                    balance_str = amounts[-1]
                    amount = float(amount_str.replace(',', ''))
                    balance = float(balance_str.replace(',', ''))
                    desc_start = line.find(current_date) + len(current_date)
                    desc_end = line.rfind(amount_str)
                    description = line[desc_start:desc_end].strip()
                    
                    # Determine correct year for transaction
                    transaction_year = determine_transaction_year(current_date, statement_date)
                    full_date = f"{current_date}/{transaction_year}"
                    
                    transactions.append({
                        'Statement_Date': statement_date,
                        'Date': full_date, 
                        'Description': description,
                        'Amount': amount,
                        'Balance': balance,
                        'Original_Line': line.strip(),
                        'Source_File': filename
                    })
        return transactions

    # Process statements
    try:
        # Get list of PDF files and sort by statement date
        pdf_files = []
        for filename in os.listdir(chase_pdf_folder):
            if filename.lower().endswith('.pdf'):
                statement_date = extract_statement_date(filename)
                pdf_files.append((statement_date, filename))
        
        # Sort files by statement date        
        pdf_files.sort(key=lambda x: x[0])
        
        all_transactions = []
        print("\n1. PROCESSING STATEMENTS") 
        print("-" * 30)
        
        # Process each PDF in chronological order
        for statement_date, filename in pdf_files:
            pdf_path = os.path.join(chase_pdf_folder, filename)
            print(f"\nProcessing: {filename} (Statement Date: {statement_date})")
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    text = ""
                    for page in pdf.pages:
                        text += page.extract_text() + "\n"
                    transactions = extract_chase_transactions(text, filename, statement_date)
                    print(f"Found {len(transactions)} transactions")
                    
                    # Add sequence number within statement
                    for i, t in enumerate(transactions):
                        t['Statement_Sequence'] = i + 1
                    all_transactions.extend(transactions)
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                continue
                
        if not all_transactions:
            messagebox.showerror("No Transactions", "No transactions found in the PDF files.")
            return
            
        # Create DataFrame preserving order  
        df = pd.DataFrame(all_transactions)
        
        # Sort by actual transaction date and sequence
        df['Sort_Date'] = pd.to_datetime(df['Date']) 
        df = df.sort_values(['Sort_Date', 'Statement_Sequence'])
        
        # Create a unique filename with folder name and timestamp
        folder_name = os.path.basename(chase_pdf_folder)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"{folder_name}_chase_transactions_{timestamp}.xlsx"
        output_path = os.path.join(chase_pdf_folder, excel_filename)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Save main transactions
            save_df = df.drop(['Sort_Date', 'Statement_Sequence'], axis=1)
            save_df.to_excel(writer, sheet_name='Transactions', index=False)
            
            # Save verification sheet
            verification_df = df[['Statement_Date', 'Date', 'Description', 'Amount', 'Original_Line']].copy()
            verification_df.to_excel(writer, sheet_name='Verification', index=False)
            
        print("\n2. VERIFICATION SUMMARY")
        print("-" * 30) 
        print(f"Total statements processed: {len(pdf_files)}")
        print(f"Total transactions found: {len(df)}")
        
        # Financial summary
        print("\n3. FINANCIAL SUMMARY")
        print("-" * 30)
        credits = df[df['Amount'] > 0]['Amount'].sum() 
        debits = df[df['Amount'] < 0]['Amount'].sum()
        print(f"Total Credits: ${credits:,.2f}")
        print(f"Total Debits: ${debits:,.2f}") 
        print(f"Net Change: ${(credits + debits):,.2f}")
        
        # Monthly summary
        print("\n4. MONTHLY SUMMARY")
        print("-" * 30)
        df['Month'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m')
        monthly = df.groupby('Month').agg({
            'Amount': ['count', 'sum']
        }).round(2)
        print(monthly)
        
        print(f"\nResults saved to: {output_path}")
        
        # Show success message
        messagebox.showinfo(
            "Processing Complete", 
            f"Successfully processed {len(pdf_files)} statements with {len(df)} transactions.\n\n"
            f"Results saved to:\n{excel_filename}"
        )
        
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror(
            "Error", 
            f"An error occurred:\n{str(e)}\n\nSee console for details."
        )

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        input("\nPress Enter to exit...")
