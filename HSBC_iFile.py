#!/usr/bin/env python
# coding: utf-8

# In[ ]:


__author__ = 'Arnaud Dahan'

#import openpyxl
import csv 
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
import os, shutil
#from zipfile import ZipFile
import pyminizip


# In[ ]:


#datetime.today().strftime("%Y%m%d")


# In[ ]:


HSBCConnectID_default='ABC22448001'
HSBCID_default='SGHBAPGSG052007432'
Account_default='741208219001'
FPName_default = 'APRIL Hong Kong Limited'
FPAddress_default = '1-13 Hollywood Road 9/F Central'

#columns
xlsxCol = ['Name', 'BankID', 'BranchNo', 'AccountNo', 'Amount', 'Reference']
xlsxSet = set(xlsxCol)


# In[ ]:


#defaultValues
NumExec_default = '1'
Currency_default = 'HKD'
ReferenceExec_default = 'Payroll'
DateExec_default = (datetime.today()+timedelta(days=1)).strftime("%Y%m%d")


# In[ ]:


#def getExcelWithDateNum(date,num):
#    xlsx_file = 'HSBC_'+date+'_'+str(num)+'.xlsx'#Path('SimData', 'play_data.xlsx')
#    if os.path.isfile(xlsx_file):
#        wb_obj = openpyxl.load_workbook(xlsx_file)
#        sheet = wb_obj.active
#        return sheet
#    else:
#        messagebox.showinfo("Excel File Missing", "Didnt found: "+str(xlsx_file))
#        return None
#
#def makeDataFromSheet(sheet):
#    data=[]
#    totals={'line':2,'amount':0}
#    xlsxLocation = {}
#    for i, row in enumerate(sheet.iter_rows(values_only=True)):
#        if i ==0:
#            cols = set(row)
#            diff = xlsxSet-cols
#            if len(diff)!=0:
#                messagebox.showinfo("Excel Column Format Error", "Missing : "+str(diff))
#                return []
#            for y, col in enumerate(row):
#                xlsxLocation[y]=col
#        else:
#            dictio = {}
#            for y, val in enumerate(row):
#                col=xlsxLocation[y]
#                if col=='Amount':
#                    totals['amount']+=val
#                dictio[col]=val
#            dictio['ID']='X'+str(i)
#            totals['line']+=1
#            
#            data.append(dictio)
#    return [data, totals]    
##sheet = getExcelWithDateNum(dateExec,numExec)


# In[ ]:


def getExcelWithDateNum(date,num):
    xlsx_file = 'HSBC_'+date+'_'+str(num)+'.csv'#Path('SimData', 'play_data.xlsx')
    if os.path.isfile(xlsx_file):
        with open(xlsx_file, mode="r", encoding="utf-8-sig") as csvfile: 
            sheet = csv.DictReader(csvfile)
            array = []
            for row in sheet:
                array.append(dict(row))
            return array
    else:
        messagebox.showinfo("CSV File Missing", "Didnt found: "+str(xlsx_file))
        return None

def makeDataFromSheet(sheet):
    #data=[]
    totals={'line':2,'bLine':0,'amount':0}
    xlsxLocation = {}
    
    cols=set(sheet[0].keys())
    diff = xlsxSet-cols
    if len(diff)!=0:
        messagebox.showinfo("Excel Column Format Error", "Missing : "+str(diff))
        return []
    
    #dictio = {}
    for i, row in enumerate(sheet):
        totals['amount']+=float(row['Amount'])
        totals['line']+=1
        totals['bLine']+=1
        row['ID']='X'+str(i)
        row['BankID'] = row['BankID'].zfill(3)
        row['BranchNo'] = row['BranchNo'].zfill(3)
    
    return [sheet, totals]    

#sheet = getExcelWithDateNum('20200417','1')
#makeDataFromSheet(sheet)


# In[ ]:





# In[ ]:


def makeString(data, totals, ref, date, num):
    iString = ""
    iString+=genHeader(totals['line'], date, num)
    iString+=genBathLine(ref, totals['amount'], totals['bLine'], date, num)
    for row in data:
        iString+=genSecLine(row, ref)
    return iString
    
#data, totals = makeDataFromSheet(sheet)
#data
#totals

#iString = makeString(data, totals, referenceExec, dateExec, numExec)
#iString


# In[ ]:


def genHeader(totalLine, date, num):
    now = datetime.now()
    dateT = now.strftime("%Y/%m/%d")
    timeT =now.strftime("%H:%M:%S")
    reference=date+'X'+str(num)
    return 'IFH,IFILE,CSV,'+HSBCConnectID+','+HSBCID+','+reference+','+dateT+','+timeT+',P,1.0,'+str(totalLine)+'\n'

#genHeader(4, dateExec, numExec)

def genBathLine(ref, totalValue, bLines, date, num):
    #reference=date+'X'+str(num)
    return 'BATHDR,ACH-CR,'+str(bLines)+',,,,,,,@1ST@,'+date+','+Account+','+Currency+','+str(totalValue)+',,,,,,,'+FPName+','+FPAddress+',,,,O0'+str(num)+','+ref+'\n'

#genBathLine(referenceExec, 2000, dateExec, numExec)

def genSecLine(row, refDefault):
    ref = refDefault if row['Reference']=='' else row['Reference']
    return 'SECPTY,'+row['AccountNo']+','+row['Name']+','+row['ID']+','+row['BankID']+','+row['BranchNo']+',,'+str(row['Amount'])+',,'+ref+',,,,,N,N'+'\n'

#genSecLine(data[0], referenceExec)

# In[ ]:


def makeIFile(date,num, iString):
    iFileName = "HSBC_iFile_"+date+"_"+str(num)+".txt"
    if os.path.isfile(iFileName): os.remove(iFileName)
    with open(iFileName,"w") as f:
        f.write(iString)
        
    if os.path.isfile(iFileName):
        return iFileName
    else: return False
        
#makeIFile(dateExec,numExec, iString)
        


# In[ ]:





# In[ ]:


def main_screen():
    global mainScreen
    mainScreen = tk.Tk()
    mainScreen.title("HSBC iFile Generator")
    mainScreen.geometry("740x350")
    mainScreen.lift()
    
    tk.Label(text="").pack()
    tk.Label(text="File needs to HSBC_YYYYMMDD_X.csv", font='Helvetica 10 bold').pack()
    tk.Label(text="YYYYMMDD is execution date, X is the operation number - 2 max by date").pack()
    colString = 'Columns: '+' | '.join(xlsxCol)
    tk.Label(text=colString).pack()
    tk.Label(text="").pack()
    
    global HSBCConnectID_value
    global HSBCID_value
    global Account_value
    global FPName_value
    global FPAddress_value
    global NumExec_value
    global Currency_value
    global ReferenceExec_value
    global DateExec_value
    
    HSBCConnectID_value = tk.StringVar()
    HSBCID_value = tk.StringVar()
    Account_value = tk.StringVar()
    FPName_value = tk.StringVar()
    FPAddress_value = tk.StringVar()
    NumExec_value = tk.StringVar()
    Currency_value = tk.StringVar()
    ReferenceExec_value = tk.StringVar()
    DateExec_value = tk.StringVar()
    
    HSBCConnectID_value.set(HSBCConnectID_default)
    HSBCID_value.set(HSBCID_default)
    Account_value.set(Account_default)
    FPName_value.set(FPName_default)
    FPAddress_value.set(FPAddress_default)
    NumExec_value.set(NumExec_default)
    Currency_value.set(Currency_default)
    ReferenceExec_value.set(ReferenceExec_default)
    DateExec_value.set(DateExec_default)
    
    
    tk.Label(text="Payer Particular", font='Helvetica 10 bold').pack()
    
    pane1 = tk.Frame(mainScreen) 
    pane1.pack(fill = tk.BOTH, expand = True)
    
    connectID_lbl = tk.Label(pane1, text="HSBCConnectID")
    connectID_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    connectID_field = tk.Entry(pane1, textvariable=HSBCConnectID_value)
    connectID_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w1_lbl = tk.Label(pane1, text=" ")
    w1_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    HSBCID_lbl = tk.Label(pane1, text="HSBCID")
    HSBCID_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    HSBCID_field = tk.Entry(pane1, textvariable=HSBCID_value)
    HSBCID_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w2_lbl = tk.Label(pane1, text=" ")
    w2_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    account_lbl = tk.Label(pane1, text="Account")
    account_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    account_field = tk.Entry(pane1, textvariable=Account_value)
    account_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w3_lbl = tk.Label(pane1, text=" ")
    w3_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    tk.Label(text="").pack()
    pane2 = tk.Frame(mainScreen) 
    pane2.pack(fill = tk.BOTH, expand = True)
    
    name_lbl = tk.Label(pane2, text="Name")
    name_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    name_field = tk.Entry(pane2, textvariable=FPName_value)
    name_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w4_lbl = tk.Label(pane2, text=" ")
    w4_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    address_lbl = tk.Label(pane2, text="Address")
    address_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    address_field = tk.Entry(pane2, textvariable=FPAddress_value)
    address_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w5_lbl = tk.Label(pane2, text=" ")
    w5_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    tk.Label(text="").pack()
    tk.Label(text="Payment Details", font='Helvetica 10 bold').pack()
    
    pane3 = tk.Frame(mainScreen) 
    pane3.pack(fill = tk.BOTH, expand = True)
    
    date_lbl = tk.Label(pane3, text="Exec Date")
    date_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    date_field = tk.Entry(pane3, textvariable=DateExec_value)
    date_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w6_lbl = tk.Label(pane3, text=" ")
    w6_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    num_lbl = tk.Label(pane3, text="Op Num")
    num_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    num_field = tk.Entry(pane3, textvariable=NumExec_value)
    num_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w7_lbl = tk.Label(pane3, text=" ")
    w7_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    ref_lbl = tk.Label(pane3, text="Int Ref")
    ref_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    ref_field = tk.Entry(pane3, textvariable=ReferenceExec_value)
    ref_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w8_lbl = tk.Label(pane3, text=" ")
    w8_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    cur_lbl = tk.Label(pane3, text="Currency")
    cur_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    cur_field = tk.Entry(pane3, textvariable=Currency_value)
    cur_field.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)
    w9_lbl = tk.Label(pane3, text=" ")
    w9_lbl.pack(side = tk.LEFT, expand = False, fill = tk.BOTH)
    
    tk.Label(text="").pack()
    tk.Button(mainScreen, text="Build iFile", width=10, height=1, command = genButtonPressed).pack()
    #genButton.pack(side = LEFT, expand = True, fill = BOTH)
    tk.Label(text="").pack()
    
    mainScreen.mainloop()
    


# In[ ]:


def genButtonPressed():
    global HSBCConnectID
    global HSBCID
    global Account
    global FPName
    global FPAddress
    global NumExec
    global Currency
    global ReferenceExec
    global DateExec
    
    HSBCConnectID = HSBCConnectID_value.get()
    HSBCID = HSBCID_value.get()
    Account = Account_value.get()
    FPName = FPName_value.get()
    FPAddress = FPAddress_value.get()
    NumExec = NumExec_value.get()
    Currency = Currency_value.get()
    ReferenceExec = ReferenceExec_value.get()
    DateExec = DateExec_value.get()
    
    sheet = getExcelWithDateNum(DateExec,NumExec)
    if sheet==None:
        return
    
    dArray = makeDataFromSheet(sheet)
    if len(dArray)!=2:
        return
    data, totals = dArray
    
    iString = makeString(data, totals, ReferenceExec, DateExec, NumExec)
    status = makeIFile(DateExec,NumExec, iString)
    if status!=False:
        messagebox.showinfo("iFile Success", "Created: "+str(status))
        
    else:
        messagebox.showinfo("Failed", "Unknown error please contact me")
    
    


# In[ ]:


def main():
    main_screen()
    
if __name__ == '__main__':
    main()
    


# In[ ]:





# In[ ]:





# In[ ]:




