#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
RakutenInvoices = pd.read_csv(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\RakutenInvoices_20191119.csv',sep=';')


# In[2]:


RakutenInvoices.head()


# In[3]:


RakutenCreditNotes = pd.read_csv(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\RakutenCreditNotes_20191119.csv',sep=';')


# In[4]:


RakutenCreditNotes.head()


# In[5]:


pivotInvoices = RakutenInvoices.pivot_table(index=['DistributorId'], values=['InvoiceNumberFull','GrossValue',"AgioAbsolute","RisikoAbsolute"], aggfunc={"InvoiceNumberFull":'count',"GrossValue":"sum","AgioAbsolute":'sum', "RisikoAbsolute":'sum'},fill_value=0)
print(pivotInvoices)
#print(type(pivotInvoices))


# In[6]:


pivotCreditNotes = RakutenCreditNotes.pivot_table(index=['DistributorId'], values=['InvoiceNumberFull','GrossValue',"AgioAbsolute","RisikoAbsolute"], aggfunc={"InvoiceNumberFull":'count',"GrossValue":"sum","AgioAbsolute":'sum', "RisikoAbsolute":'sum'},fill_value=0)
print(pivotCreditNotes)


# In[7]:


pivotInvoices.loc[[42],["AgioAbsolute"]]


# In[8]:


pivotInvoices.get_value(42, "AgioAbsolute") 


# In[9]:


Summary = pd.read_excel(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\Summary.xlsx')


# In[10]:


with pd.ExcelWriter(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\WeeklyFactoringReport.xlsx') as writer:  
    Summary.to_excel(writer, sheet_name='Summary', index=False)
    RakutenInvoices.to_excel(writer, sheet_name='Invoices', index=False)
    RakutenCreditNotes.to_excel(writer, sheet_name='Credit notes', index=False)
    pivotInvoices.to_excel(writer, sheet_name='pivotTableInvoices')
    pivotCreditNotes.to_excel(writer, sheet_name='pivotTableCreditNotes')


# In[11]:


from openpyxl import load_workbook


# In[12]:


#create workbook object
wb = load_workbook(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\WeeklyFactoringReport.xlsx')

#create reference for sheet on which to write
worksheet= wb.get_sheet_by_name("Summary") 

#use sheet reference and write the cell address
#Change week
from datetime import date
import datetime


worksheet["F4"].value= date.today().isocalendar()[1]
#MS GSA: Invoices
worksheet["B13"].value= '=IFERROR(-D13/(D12/B12+D11),0)'
worksheet["B14"].value= '=IFERROR(-(D14/C14)/D11,0)'
worksheet["B15"].value= '=IFERROR(-(D15/C15)/D12,0)'

worksheet["D11"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$B:$B,42)'
worksheet["D12"].value= '=SUMIFS(Invoices!$R:$R,Invoices!$C:$C,1,Invoices!$Q:$Q,"CHF")'
worksheet["D13"].value= '=-SUMIFS(Invoices!$W:$W,Invoices!$B:$B,42)'
worksheet["D14"].value= '=-(SUMIFS(Invoices!$AA:$AA,Invoices!$B:$B,42))'  #change formula
worksheet["D15"].value= '=-SUMIFS(Invoices!V:V,Invoices!$C:$C,1,Invoices!$Q:$Q,"CHF")*C15'
#MS GSA: CreditNotes
worksheet["H12"].value= '=B12'
worksheet["H13"].value= '=IFERROR(-J13/(J12/H12+J11),0)'
worksheet["H14"].value= '=IFERROR(-(J14/I14)/J11,0)'
worksheet["H15"].value= '=IFERROR(-(J15/I15)/J12,0)'

worksheet["J11"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!B:B,42)"
worksheet["J12"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!$C:$C,1,'Credit notes'!R:R,\"CHF\")"
worksheet["J13"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$B:$B,42)" #change formula
worksheet["J14"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$B:$B,42)" #change formula
worksheet["J15"].value= "=SUMIFS('Credit notes'!Y:Y,'Credit notes'!$C:$C,1,'Credit notes'!R:R,\"CHF\")*I15"

#Net value and net no. Invoices
worksheet["L11"].value= '=D11+J11'
worksheet["M11"].value= """=IFERROR(VLOOKUP(42,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(42,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[13]:


#Media Shop BENELUX: Invoices

worksheet["B19"].value= '=IFERROR((-D19/D18),0)'
worksheet["B20"].value= '=IFERROR(-(D20/C20)/D18,0)'

worksheet["D18"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$B:$B,"3631328")+SUMIFS(Invoices!$T:$T,Invoices!$B:$B,"3631332")'
worksheet["D19"].value= '=-SUMIFS(Invoices!$W:$W,Invoices!$B:$B,"3631328")-SUMIFS(Invoices!$W:$W,Invoices!$B:$B,"3631332")'
worksheet["D20"].value= '=-((SUMIFS(Invoices!$AA:$AA,Invoices!$B:$B,3631328)+(SUMIFS(Invoices!$AA:$AA,Invoices!$B:$B,3631332))))' #change formula

#Media Shop BENELUX: CreditNotes
worksheet["H19"].value= '=IFERROR((-J19/J18),0)'
worksheet["H20"].value= '=IFERROR(-(J20/I20)/J18,0)'


worksheet["J18"].value= "=-(SUMIFS('Credit notes'!S:S,'Credit notes'!$B:$B,\"3631328\") + SUMIFS('Credit notes'!S:S,'Credit notes'!$B:$B, \"3631332\" ))"
worksheet["J19"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$B:$B,\"3631328\") + SUMIFS('Credit notes'!X:X,'Credit notes'!$B:$B,\"3631332\")" #change formula
worksheet["J20"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$B:$B,\"3631328\")*I20  + SUMIFS('Credit notes'!Y:Y,'Credit notes'!$B:$B, \"3631332\" )" #change formula



#Net value and net no. Invoices
worksheet["L18"].value= '=D18+J18'
worksheet["M18"].value = """=IFERROR(VLOOKUP(3631328,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(3631328,pivotTableCreditNotes!A:E,4,FALSE),0)+IFERROR(VLOOKUP(3631332,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(3631332,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[14]:


#Chronext: Invoices

worksheet["B24"].value= '=IFERROR(-D24/(D23),0)'
worksheet["B25"].value= '=IFERROR(-(D25/C25)/D23,0)'

worksheet["D23"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$B:$B,1642346)'
worksheet["D24"].value= '=-(SUMIFS(Invoices!$W:$W,Invoices!$B:$B,1642346))'
worksheet["D25"].value= '=-(SUMIFS(Invoices!$AA:$AA,Invoices!$B:$B,1642346))' #change formula

#Chronext: CreditNotes
worksheet["H24"].value= '=IFERROR(-J24/(J23),0)'
worksheet["H25"].value= '=IFERROR(-(J25/I25)/J23,0)'


worksheet["J23"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!$B:$B,1642346,'Credit notes'!R:R,\"EUR\")"
worksheet["J24"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$B:$B,1642346)" #change formula
worksheet["J25"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$B:$B,1642346)" #change formula



#Net value and net no. Invoices
worksheet["L23"].value= '=D23+J23'
worksheet["M23"].value = """=IFERROR(VLOOKUP(1642346,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(1642346,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[15]:


#Agility: Invoices

worksheet["B29"].value= '=IFERROR((-D29/D28),0)'
worksheet["B30"].value= '=IFERROR(-(D30/C30)/D28,0)'

worksheet["D28"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$C:$C,4)'
worksheet["D29"].value= '=-SUMIFS(Invoices!$W:$W,Invoices!$C:$C,4)'
worksheet["D30"].value= '=-(SUMIFS(Invoices!$AA:$AA,Invoices!$C:$C,4))' #change formula

#Agility: CreditNotes
worksheet["H29"].value= '=IFERROR((-J29/J28),0)'
worksheet["H30"].value= '=IFERROR(-(J30/I30)/J28,0)'


worksheet["J28"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!$C:$C,4,'Credit notes'!R:R,\"EUR\")"
worksheet["J29"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$C:$C,4)" #change formula
worksheet["J30"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$C:$C,4)" #change formula



#Net value and net no. Invoices
worksheet["L28"].value= '=D28+J28'
worksheet["M28"].value = """=IFERROR(VLOOKUP(581029,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(581029,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[16]:


#MAPA : Invoices

worksheet["B34"].value= '=IFERROR((-D34/D33),0)'
worksheet["B35"].value= '=IFERROR(-(D35/C35)/D33,0)'

worksheet["D33"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$C:$C,12)'
worksheet["D34"].value= '=-SUMIFS(Invoices!$W:$W,Invoices!$C:$C,12)'
worksheet["D35"].value= '=-(SUMIFS(Invoices!$AA:$AA,Invoices!$C:$C,12))' #change formula

#MAPA : CreditNotes
worksheet["H34"].value= '=IFERROR((-J34/J33),0)'
worksheet["H35"].value= '=IFERROR(-(J35/I35)/J33,0)'


worksheet["J33"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!$C:$C,12,'Credit notes'!R:R,\"EUR\")"
worksheet["J34"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$C:$C,12)" #change formula
worksheet["J35"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$C:$C,12)" #change formula



#Net value and net no. Invoices
worksheet["L33"].value= '=D33+J33'
worksheet["M33"].value = """=IFERROR(VLOOKUP(1597975,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(1597975,pivotTableCreditNotes!A:E,4,FALSE),0)+IFERROR(VLOOKUP(1597997,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(1597997,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[17]:


#Wachtmaster: Invoices

worksheet["B39"].value= '=IFERROR((-D39/D38),0)'
worksheet["B40"].value= '=IFERROR(-(D40/C40)/D38,0)'

worksheet["D38"].value= '=SUMIFS(Invoices!$T:$T,Invoices!$B:$B,1926195)'
worksheet["D39"].value= '=-SUMIFS(Invoices!$W:$W,Invoices!$B:$B,1926195)'
worksheet["D40"].value= '=-(SUMIFS(Invoices!$AA:$AA,Invoices!$B:$B,1926195))'  #change formula

#Wachtmaster: CreditNotes
worksheet["H39"].value= '=IFERROR((-J39/J38),0)'
worksheet["H40"].value= '=IFERROR(-(J40/I40)/J38,0)'


worksheet["J38"].value= "=-SUMIFS('Credit notes'!S:S,'Credit notes'!$B:$B,1926195,'Credit notes'!R:R,\"EUR\")"
worksheet["J39"].value= "=SUMIFS('Credit notes'!W:W,'Credit notes'!$B:$B,1926195)" #change formula
worksheet["J40"].value= "=SUMIFS('Credit notes'!X:X,'Credit notes'!$B:$B,1926195)" #change formula



#Net value and net no. Invoices
worksheet["L38"].value= '=D38+J38'
#worksheet["M38"].value= pivotInvoices.get_value(1926195, "InvoiceNumberFull") - pivotCreditNotes.get_value(1926195, "InvoiceNumberFull")
worksheet["M38"].value = """=IFERROR(VLOOKUP(1926195,pivotTableInvoices!A:E,4,FALSE),0)-IFERROR(VLOOKUP(1926195,pivotTableCreditNotes!A:E,4,FALSE),0)"""


# In[18]:


#Total EUR Invoices
worksheet["D51"].value= "=D11+D18+D23+D28+D33+D46+D38"
worksheet["D52"].value= "=D12"
worksheet["D53"].value= "=D46"
worksheet["D54"].value= "=D13+D19+D24+D29+D34+D39+D47"
worksheet["D55"].value= "=+D14+D20+D25+D30+D35+D40+D48"
worksheet["D56"].value= "=+D15"
worksheet["D57"].value= "=+D48"

worksheet["D60"].value= "=D51+D55"
worksheet["D62"].value= "=D52+D56"
worksheet["D64"].value= "=D53+D57"

#Total EUR Credit Notes
worksheet["J51"].value= "=J11+J18+J23+J28+J33+J38+J46"
worksheet["J52"].value= "=J12"
worksheet["J53"].value= "=J46"
worksheet["J54"].value= "=J13+J19+J24+J29+J34+J39+J47"
worksheet["J55"].value= "=+J14+J20+J25+J30+J35+J40"
worksheet["J56"].value= "=+J15"
worksheet["J57"].value= "=+J48"

worksheet["J60"].value= "=J51+J55"
worksheet["J62"].value= "=J52+J56"
worksheet["J64"].value= "=J53+J57"


# In[19]:


#CASH FLOW 
worksheet["C72"].value= "=+D51"
worksheet["C73"].value= "=+J51"
worksheet["C74"].value= "=D55+J55"
worksheet["C75"].value= "=J54"
worksheet["C76"].value= "=SUM(C72:C75)"

worksheet["C79"].value= "=+D54"
worksheet["C80"].value= "=+D54"

worksheet["D72"].value= "=+D52"
worksheet["D73"].value= "=+J53"
worksheet["D74"].value= "=D56+J56"
worksheet["D76"].value= "=SUM(D72:D74)"

worksheet["E72"].value= "=+D53"
worksheet["E73"].value= "=+J52"
worksheet["E74"].value= "=D57+J57"
worksheet["E76"].value= "=SUM(E72:E74)"


# In[20]:


#create reference for sheet on which to write
worksheetpivotInvoices= wb.get_sheet_by_name("pivotTableInvoices") 

#use sheet reference and write the cell address
#Change week
worksheet["F4"].value= 43
#MS GSA: Invoices
worksheet["B13"].value= '=IFERROR(-D13/(D12/B12+D11),0)'
worksheet["B14"].value= '=IFERROR(-(D14/C14)/D11,0)'
worksheet["B15"].value= '=IFERROR(-(D15/C15)/D12,0)'


# In[21]:


#save workbook
wb.save(r'C:\NhanLeDocomoDigital\Automation\FactoringReport\CW latest\template.xlsx')


# In[ ]:





# In[ ]:





# In[ ]:




