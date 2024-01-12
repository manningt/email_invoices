#!/usr/bin/env python3
'''
This program parses a PDF with one or more pages of invoices.  
For each page, it writes a single page invoice with a filename format: customer_name_invoice_number.pdf
if the page has an email address, then it emails the PDF as an attachment
'''

import sys
import os
from enum import IntEnum
import array as arr
import logging
# logging.basicConfig()
# logging.getLogger().setLevel(logging.DEBUG)

import argparse
import tkinter as tk
from tkinter import filedialog
import json

try:
   from PyPDF2 import PdfReader, PdfWriter
except:
   sys.exit('failed to load PyPDF2; try pip3 install PyPDF2 to resolve')

try:
   import yagmail
except:
   sys.exit('failed to load yagmail; try pip3 install yagmail to resolve')

test_connection= False
if test_connection:
   try:
      with yagmail.SMTP("test", oauth2_file="../link_to_oauth2.json") as yag:
         yag.close()
   except:
      sys.exit('failed to load email credentials file: {oauth2_file}')
# yag.set_logging(yagmail.logging.DEBUG)
# yag.setLog(log_level = logging.DEBUG)
   

def parse_pdf(filename, parse_only = False):
   email_list = list() #contains tuples of (email, attachment_filename); used as return data

   class FieldType(IntEnum):
      account_num = 1
      invoice_num = 2
      email = 3
      name = 4
      terms = 5

   terms_dont_email_str = "EoS"

   missing_field_count = arr.array('i',[0,0,0,0])
   missing_field_list = ([],[],[]) #list of customer names missing a field

   reader = PdfReader(filename)
   for pageno in range(len(reader.pages)):
      page = reader.pages[pageno] 
      text = page.extract_text() 
      # print(text + '\n') 

      email_str = None
      customer_name_str = None
      terms_str = None
      out_filename = None

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
         if line.endswith("Terms"):
            terms_str = lines[lineno+1].partition('Account')[0] 
            # The line after Terms has Account # in it, e.g. Net 30Account #
            # The partition operator returns a tuple, which we only want the prefix, which is the zero of the tuple.
            # prefix, match, suffix = terms_str.partition('Account')

      print_status_per_page = False
      # print_status_per_page = True
      if print_status_per_page:
         print(f'\nPage: {pageno}')
         # account number is currently not used.  In the future, it could be used to look up more info in a customer database.
         # print(f'\t{account_num = }')
         print(f'\t{customer_name_str = }')
         print(f'\t{invoice_num = }')
         print(f'\t{email_str = }')
         print(f'\t{account_num = }')
         print(f'\t{terms_str = }')

      if customer_name_str is None:
         print(f'Missing customer_name {pageno =}: {account_num = }, {invoice_num = }')
         missing_field_count[FieldType.name] += 1
      elif invoice_num is None:
         missing_field_count[FieldType.invoice_num] += 1
         missing_field_list[FieldType.invoice_num].append(customer_name_str)
      elif parse_only:
         pass
      else:
         # write single page PDF
         writer = PdfWriter() 
         writer.add_page(page)
         customer_name_for_filename = customer_name_str.replace(" ","_")
         customer_name_for_filename = customer_name_for_filename.replace("&","")
         out_filename = f'{customer_name_str}_invoice_{invoice_num}.pdf'
         out_file = open(out_filename,'wb') 
         writer.write(out_file) 
         out_file.close()

      if email_str is not None:
         if terms_str is not None and terms_dont_email_str in terms_str:
            print(f'Not emailing "{customer_name_str}" because {terms_str=} contains {terms_dont_email_str}')
         elif "@" not in email_str:
            print(f'Not emailing "{customer_name_str}" because missing or invalid email address')
         # the folllowing was to handle a bug, can be deleted in the long run.
         # elif account_num == 130:
         #    print(f'Not emailing "{customer_name_str}" because account number is excluded')
         else:
            customer_tuple = (email_str, out_filename)
            email_list.append(customer_tuple)
      else:
         missing_field_count[FieldType.email] += 1
         missing_field_list[FieldType.email].append(customer_name_str)

      if account_num is None:
         missing_field_count[FieldType.account_num] += 1
         missing_field_list[FieldType.account_num].append(customer_name_str)

   missing_count = 0
   for i in missing_field_count:
      missing_count = missing_count + i
   print(f'Invoice Count={len(reader.pages)}; ', end = '')
   if missing_count == 0:
       print(f'No missing fields')
   else:
      print(f'')

   if missing_field_count[FieldType.invoice_num] > 0:
      print(f'{missing_field_count[FieldType.invoice_num]} missing invoice_number.  Customers are {missing_field_list[FieldType.invoice_num]}')

   if missing_field_count[FieldType.account_num] > 0:
      print(f'{missing_field_count[FieldType.account_num]} missing account_num.  Customers are {missing_field_list[FieldType.account_num]}')

   if missing_field_count[FieldType.email] > 0:
      print(f'{missing_field_count[FieldType.email]} missing email addr.  Customers are {missing_field_list[FieldType.email]}')

   return email_list


def email_invoices(email_list, auth_filepath):
   import time
   i = 0
   try:
      with open(auth_filepath, 'r') as auth_file:
         auth_data = json.load(auth_file)
         # print(f'{auth_data["email_address"]= } {auth_data["app_password"]= }')
   except:
      print(f'Failed to parse {auth_filepath}')
      return 0

   #email_list format: [(email1, attachment1), (email2, attachment2), etc]
   print() #to make a blank line before printing out statistics

   with yagmail.SMTP(auth_data["email_address"], auth_data["app_password"]) as yag:
      for customer in email_list:
         yag.send(to=customer[0], subject="Plowing invoice attached", contents="Thank You-\nCharlie",
                  attachments=customer[1])
         print(f'sent: {customer[0]}  attachment: {customer[1]}')
         i += 1
         time.sleep(2)
   return i
         

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
   default_auth_path = "email_credentials.json"
   argParser = argparse.ArgumentParser()
   argParser.add_argument("input", type=str, help="input PDF filename with path")
   argParser.add_argument("-c", "--auth_path", help="path to credentails file", nargs='?',
               const=default_auth_path, default=default_auth_path, type=str)
   # store_false will default to True when the command-line argument is not present
   argParser.add_argument("-p", "--parse_only", action='store_true', help="if false, dont output PDFs or email - parse only")
   argParser.add_argument("-d", "--dont_email", action='store_true', help="if false, dont send emails")

   args = argParser.parse_args()
   # print(f'\n\t{args.input= } {args.auth_path= }\n\t{args.dont_email= } {args.parse_only= }')

   if args.input is None:
      pdf_filename = pick_file()
      if pdf_filename is None:
         sys.exit("No PDF files selected to parse.")
      pdf_filename = f'../{pdf_filename}'
   else:
      pdf_filename = args.input

   email_list = parse_pdf(pdf_filename, args.parse_only)
   # print(f'{email_list=}')

   if len(email_list) < 1:
      print("no email addresses in the PDF file.")
   elif args.dont_email is False and args.parse_only is False:
      if not os.path.isfile(args.auth_path):
         auth_filename = pick_file()
         if auth_filename is None:
            sys.exit("No credentials file selected; unable to email.")
         auth_filepath = f'../{auth_filename}'
      else:
         auth_filepath = args.auth_path

      emails_sent= email_invoices(email_list, auth_filepath)
      print(f'Sent {emails_sent} emails')
