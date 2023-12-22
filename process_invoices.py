#!/usr/bin/env python3
'''
This program parses a PDF with one or more pages of invoices.  
For each page, it writes a single page invoice with a filename format: customer_name_invoice_number.pdf
if the page has an email address, then it emails the PDF as an attachment
'''

import sys
try:
   from PyPDF2 import PdfReader, PdfWriter
except:
   sys.exit('failed to load PyPDF2; try pip3 install PyPDF2 to resolve')

try:
   import yagmail
except:
   sys.exit('failed to load yagmail; try pip3 install yagmail to resolve')
   # print("failed to load yagmail", file=sys.stderr )

# rm oauth2_file="../client_secret_737781932207-sh8h40p4k1h9d01sabihs6kts9hk02no.apps.googleusercontent.com.json"
try:
   yag = yagmail.SMTP("roi.co.4444@gmail.com", oauth2_file="../client_secret_737781932207-sh8h40p4k1h9d01sabihs6kts9hk02no.apps.googleusercontent.com.json")
except:
   sys.exit('failed to load email credentials file: {oauth2_file}')

import tkinter as tk
from tkinter import filedialog

TK_SILENCE_DEPRECATION=1

def parse_pdf(filename):
   reader = PdfReader(filename) 

   for pageno in range(len(reader.pages)):
      page = reader.pages[pageno] 
      text = page.extract_text() 
      #print(text + '\n') 

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

      print_status_per_page = False
      if print_status_per_page:
         print('\nPage: {pageno}')
         # account number is currently not used.  In the future, it could be used to look up more info in a customer database.
         # print(f'\t{account_num = }')
         print(f'\t{customer_name_str = }')
         print(f'\t{invoice_num = }')
         print(f'\t{email_str = }')

      if customer_name_str is None or invoice_num is None:
         print(f'Error on {pageno =}: {customer_name_str = }, {invoice_num = }')
      else:
         writer = PdfWriter() 
         writer.add_page(page)
         customer_name_for_filename = customer_name_str.replace(" ","_")
         customer_name_for_filename = customer_name_for_filename.replace("&","")
         out_filename = f'{customer_name_str}_invoice_{invoice_num}.pdf'
         out_file = open(out_filename,'wb') 
         writer.write(out_file) 
         out_file.close() 

         if email_str is not None:
            receiver = "tom@manningetal.com"
            yag.send(to=receiver,
               subject="Plowing invoice attached",
               contents="Thank You-\nCharlie",
               attachments=out_filename,
            )


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
