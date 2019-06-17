TnxID = ['6F26E0-1560434603',
'6F273F-1560435049',
'6F274C-1560435298',
'6F27EB-1560436932',
'6F2829-1560437167',
'6F2C06-1560441099',
'6F2DBD-1560443901',
'6F304C-1560447768',
'6F3052-1560448528',
'6F305A-1560448738',
'6F3516-1560459954',
'6F3583-1560461304',
'6F3597-1560463238',
'6F3248-1560455033'
]

counter = len(TnxID) - 1
date = '2019-06-17'
bankaccount = '10001 · A-Woodforest LLC 3221'
for i in TnxID:
  
  if counter != 0:
    print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{bankaccount}',{{d'{date}'}},1);")
  else:
    print (f"INSERT INTO DepositLine (DepositLinePaymentTxnID, DepositToAccountRefFullName,TxnDate,FQSaveToCache) Values ('{i}','{bankaccount}',{{d'{date}'}},0);")
  counter = counter - 1