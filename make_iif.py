#!/usr/bin/env python3
'''
Script to convert payments CSV (customer_name,payment) to IIF output.
the CSV file should have the first row be: customer_id,amount
'''

import os
import sys, traceback
import datetime
# import re
import csv

PROJECT_ROOT = os.path.dirname(os.path.realpath(__file__))

# example file format; Undeposited Funds has to come before Accounts Recievable
'''
!TRNS   TRNSTYPE        DATE    ACCNT   NAME    AMOUNT
!SPL    TRNSTYPE        DATE    ACCNT   NAME    AMOUNT
!ENDTRNS
TRNS    PAYMENT 12/27/2023      Undeposited Funds       AAtest1, A1     35
SPL     PAYMENT 12/27/2023      Accounts Receivable     AAtest1, A1     -35
ENDTRNS
'''


def main(input_file_name):

   input_file_path = os.path.join(PROJECT_ROOT, input_file_name) + ".csv"

   if not os.path.isfile(input_file_path):
      sys.exit(f"{input_file_path} does not exist")

   output_file = open(os.path.join(PROJECT_ROOT, input_file_name + '.iif'), 'w')

   # write IIF template
   head = "!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\r\n"
   head += "!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\r\n"
   head += "!ENDTRNS\r\n"
   output_file.write(head)

   formatted_date = datetime.datetime.now().strftime('%m/%d/%Y')   

   template = "TRNS\tPAYMENT\t{payment_date}\tUndeposited Funds\t{customer_name}\t{amount}\r\n"
   template += "SPL\tPAYMENT\t{payment_date}\tAccounts Receivable\t{customer_name}\t-{amount}\r\n"
   template += "ENDTRNS\r\n"

   cust_key = "customer_id"
   amt_key = "amount"

   with open(input_file_path) as csvfile:
      csv_reader = csv.DictReader(csvfile)
      # print(f'{csv_reader.fieldnames= }')

      if cust_key not in csv_reader.fieldnames:
         sys.exit(f'{cust_key} column is not in {input_file_name}.csv')
      if amt_key not in csv_reader.fieldnames:
         sys.exit(f'{amt_key} column is not in {input_file_name}.csv')

      for row in csv_reader:
         try:
            amt_float = float(row[amt_key])
         except:
            amt_float = -1

         if (amt_float) < 0:
            print(f'"{row[cust_key]}" has a invalid payment of {row[amt_key]}')
         elif (amt_float) == 0:
            print(f'skipping "{row[cust_key]}" because payment is {row[amt_key]}')
         else:
            output_file.write(template.format(payment_date= formatted_date, customer_name = row[cust_key], amount = row[amt_key]))


if __name__ == '__main__':

   if len(sys.argv) != 2:
      print(f"usage:   make_iif.py input_filepath")

   main(sys.argv[1])