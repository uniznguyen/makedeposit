import pyodbc
import os
import pandas as pd
import numpy as np
import csv
from decimal import Decimal

BANKACCOUNT = '10001 · A-Woodforest LLC 3221'
QUERY_BACK_DAYS = 5

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
depositcsv = os.path.join(BASE_DIR,'deposit.xlsx')


df = pd.read_excel(depositcsv)
df['Amount'] = df['Amount'].astype(np.float64).round(2)

deposit_date = df['When'].iloc[0].date()
deposit_amount = df['Amount'].sum().round(2)
deposit_count = len(df.index)
amount_list = tuple(df['Amount'].tolist())


print (f"Deposit date: {deposit_date}")
print (f"Total deposit amount from batch report: {deposit_amount}")
print (f"Total deposit counts: {deposit_count}")



cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = f"SELECT TxnID, TxnDate, Amount FROM ReceivePaymentToDeposit WHERE TxnDate >= ({{d'{deposit_date}'}} - {QUERY_BACK_DAYS}) AND Amount IN {amount_list}"

df2 = pd.read_sql(sql,cn, parse_dates=['TxnDate'])
df2['Amount'] = df2['Amount'].astype(np.float64).round(2)

print (df2)

deposit_amount_qb = df2['Amount'].round(2).sum()
deposit_count_qb = len(df2.index)

print (f"Total deposit amount from QB undeposited funds: {deposit_amount_qb}")
print (f"Total deposit counts: {deposit_count_qb}")

TxnID = df2['TxnID'].tolist()
# counter = len(TxnID) - 1

def print_insert():
  counter = len(TxnID) - 1
  for i in TxnID:
    if counter != 0:
      print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},1);")
    else:
      print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{BANKACCOUNT}',{{d'{deposit_date}'}},0)")
    counter = counter - 1  
    

if deposit_amount != deposit_amount_qb or deposit_count != deposit_count_qb :
  print ("Don't match")
  response = input("Still want to process:? enter 'y' or 'n'")
  if response == 'y':
        print_insert()
  else:
    print("Stop")
  
        

cn.close()

