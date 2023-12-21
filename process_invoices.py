'''
This program parses a PDF with one or more pages of invoices.  
For each page, it writes a single page invoice with a filename format: customer_name_invoice_number.pdf
if the page has an email address, then it emails the PDF as an attachment
'''
import sys

import tkinter as tk
from tkinter import filedialog

from PyPDF2 import PdfReader, PdfWriter
import re

TK_SILENCE_DEPRECATION=1

def parse_pdf(filename):
   reader = PdfReader(filename) 

   # printing number of pages in pdf file 
   # print(len(reader.pages)) 

   for pageno in range(len(reader.pages)):
 
      page = reader.pages[pageno] 
      text = page.extract_text() 
      #print(type(text))
      #print(text + '\n') 

      if 0:
         acct_pos = re.search(r'Account #', text)
         account_num = int((text[acct_pos.end()+1:(acct_pos.end()+4)]))

      lines = text.split("\n")
      for lineno, line in enumerate(lines):
         if line.endswith("Invoice #"):
            try:
               invoice_num = int(lines[lineno+1])                        
            except:
               invoice_num = None
         if line.endswith("Account #"):
            try:
               account_num = int(lines[lineno+1])
            except:
               account_num = None
         if line.endswith("Customer E-mail"):
            email_str = lines[lineno+1]
         if line.endswith("Bill To"):
            customer_name_str = lines[lineno+1]

      print('\nPage: {pageno}')
      # account number is currently not used.  In the future, it could be used to look up more info in a customer database.
      # print(f'\t{account_num = }')
      print(f'\t{customer_name_str = }')
      print(f'\t{invoice_num = }')
      print(f'\t{email_str = }')

      writer = PdfWriter() 
      writer.add_page(page)
      customer_name_for_filename = customer_name_str.replace(" ","_")
      customer_name_for_filename = customer_name_for_filename.replace("&","")
      out_filename = f'{customer_name_str}_invoice_{invoice_num}.pdf'
      out_file = open(out_filename,'wb') 
      writer.write(out_file) 
      out_file.close() 


def pick_file():
   root = tk.Tk()
   root.withdraw()  # Hide the main window
   file_path = filedialog.askopenfilename(title="Select a File")
    
   if file_path:
      # Extract the filename from the full path
      filename = file_path.split('/')[-1]  # For Unix/Linux
      # filename = file_path.split('\\')[-1]  # For Windows

      # print(f"Selected filename: {filename}")
      return filename
   else:
      print("No file selected.")
      return None

if __name__ == "__main__":
   pdf_filename = pick_file()
   if pdf_filename is None:
      sys.exit(1)
   up_pdf_filename = f'../{pdf_filename}'
   parse_pdf(up_pdf_filename)
