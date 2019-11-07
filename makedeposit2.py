import win32com.client
import re
# from datetime import datetime
# from decimal import Decimal
import pyodbc
import os
import pandas as pd
import numpy as np
import csv
import win32clipboard as cb


BANKACCOUNT = '10001 · A-Woodforest LLC 3221'
QUERY_BACK_DAYS = 5
batch_number = 204
deposit_date = '2019-11-07'

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# GetDefaultFolder with index = 6 is the Inbox
inbox = outlook.GetDeFaultFolder(6)

messages = inbox.Items.Restrict("[SenderEmailAddress] = 'notifications@paytrace.com' And [SenderName] = 'PayTrace'")



for message in messages:
    if f"Settlement Report for Batch {batch_number}" in message.subject:
        print (message.Subject)
        message_content = message.body
        Amounts = re.findall(r'(Purchase|Refund) \| Amount = \$(\d*,?\d*\.\d*)',message_content)

        amount_list = []

        for i in range(len(Amounts)):
            
            if Amounts[i][0] == "Refund":
                amount_list.append(-float(Amounts[i][1].replace(",","")))
            else:
                amount_list.append(float(Amounts[i][1].replace(",","")))
        
        amount_list = tuple(amount_list)        
        
        net_settlement_amount = re.search(r'settlement amount is \$(\d+,?\d+\.\d+)',message_content).group(1).replace(",","")
        net_settlement_amount = float(net_settlement_amount)
        total_transactions = round(float(sum(amount_list)),2)
        
if net_settlement_amount == total_transactions:
    print ("Total transactions match net settlement amount in batch report")
    print (f"net_settlement_amount: {net_settlement_amount:,.2f}")
    print (f"total_transactions: {total_transactions:,.2f}")
    print (*amount_list, sep='\n')
    deposit_count = len(amount_list)
    print (f"deposit_count :{deposit_count}")
else:
  print ("Total transactions don't match net settlement amount in batch report")
  print (f"net_settlement_amount: {net_settlement_amount:,.2f}")
  print (f"total_transactions: {total_transactions:,.2f}")
  print (*amount_list, sep='\n')
  exit()    


#####################################################

cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = f"SELECT TxnID, TxnDate, Amount FROM ReceivePaymentToDeposit WHERE TxnDate >= ({{d'{deposit_date}'}} - {QUERY_BACK_DAYS}) AND Amount IN {amount_list}"

df2 = pd.read_sql(sql,cn, parse_dates=['TxnDate'])
df2['Amount'] = df2['Amount'].astype(np.float64).round(2)

print (df2)

deposit_amount_qb = round(df2['Amount'].sum(),2)
deposit_count_qb = len(df2.index)

print (f"Total deposit amount from QB undeposited funds: {deposit_amount_qb:,.2f}")
print (f"Total deposit counts: {deposit_count_qb}")

TxnID = df2['TxnID'].tolist()
amount_qb = df2['Amount'].tolist()


def print_insert():
  insert_query = ""
  counter = len(TxnID) - 1
  for i in TxnID:
    if counter != 0:
      print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},1);")
      insert_query = insert_query + f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},1);" + "\n"
    else:
      print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},0)")
      insert_query = insert_query + f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},0)" + "\n"
    counter = counter - 1
  
  ## this code is to copy the insert query string to clipboard
  cb.OpenClipboard()
  cb.EmptyClipboard()
  cb.SetClipboardData(cb.CF_UNICODETEXT,insert_query)
  cb.CloseClipboard()
  ## this code is to copy the insert query string to clipboard    

def get_amount_not_in_qb(amount_list,amount_qb):
  print ("This amount is not in QBB")
  for each in amount_list:
    if each not in amount_qb:
      print ("="*5,each)      

if net_settlement_amount != deposit_amount_qb or deposit_count != deposit_count_qb :
  print ("Don't match")
  get_amount_not_in_qb(amount_list,amount_qb)
  response = input("Still want to process:? enter 'y' or 'n'")
  if response == 'y':
        print_insert()
  else:
    print("Stop")
else:
  print_insert()

cn.close()