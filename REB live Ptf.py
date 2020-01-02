#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pyodbc
import pandas as pd
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};.............')


# In[2]:


query = """

declare @Ultimo as datetime = '20191201 00:01:00'

select 
	cast(o.TimeStamp as date)		'OrderDate',
	i.InvoiceNumberFull, i.Id		'InvoiceId',
	cast(i.SellDate as date)		'Sold to the bank',
	a.Name							'DistributorName',
	g.Description					'DistributorGroup',
	--o.CustomerId					'CustomerId',
	cast(i.InvoiceDate as date)		'InvoiceDate',
	concat(year(i.InvoiceDate)%100 , FORMAT(i.InvoiceDate,'MM')) InvPeriods,
	--'DUMMY'							'PaymentCompleted',
	p.Amount								'TotalPaid',
	--'DUMMY'							'Payment Date',
	coalesce(r.MaxReminder,0)		'ReminderLevel',
	e.Code							'WHG',
	i.OrderTotal					'Original Amount',
	case when cast(i.OrderTotal-coalesce(c.Amount,0) as money) < 0 then 0 else cast(i.OrderTotal-coalesce(c.Amount,0) as money)  end as  'Inv_netCN',																								
	case when cast(i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0)  as money) <0 then 0 else																								
		 cast(i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0)  as money) end as'OpenAmount_WOReminder',	
	--i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0) 'AmountTotal',
	--'DUMMY'							'AmountTotalLastMonth',
	coalesce(r2.ReminderAmount,0)	'RemindersAmount',
	coalesce(pa.PaidInk,0) 'RecoveredAmount',
	--'DUMMY'							'RemindersAmountLastMonth',
	--cast(h.AgreementDate as date)	'RemindersAgreementAmountDatum',
	i.TermOfPaymentDays				'Due Days',
	coalesce(r2.ReminderAmount,0)	'RemindersAgreementAmountNOW',
	coalesce(p.Amount,0)			'PaymentsAmount',
	co.taxrate 'Tax rate', --'DUMMY'	'PaymentsAmountLastMonth',
	coalesce(c.Amount,0)			'CreditsAmount',
	--case when ie.grundlagennummer is not null then 'YES' else 'NO' end ' InkassoEinstellung',
	--case when r.AgreementAmount is null then 'NO' else 'YES' end 'AgreementsYN',
	case when coalesce(Inkasso,0) = 0 then 'NO' else 'YES' end 'InkassoYN',
	cast(r.InkassoDate as date)		'InkassoTimestamp',
	concat(year(r.InkassoDate )%100 , FORMAT(r.InkassoDate ,'MM')) InkPeriods,																																											
	case when r.InkassoDate is null or  cast(i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0)  as money) < 0 																								
		then 0 else cast(i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0) as money) end as  'InkassoDueAmount_WOReminder',	
																																																
	case when r.InkassoDate is null or  cast(i.OrderTotal-coalesce(p.Amount,0)+coalesce(pa.PaidInk,0)-coalesce(c.Amount,0)	 as money) < 0 																								
		                    then 0 else cast(i.OrderTotal-coalesce(p.Amount,0)+coalesce(pa.PaidInk,0)-coalesce(c.Amount,0) as money) end as 'InkassoHandover_WOReminder'	,	

	y.Name							'PaymentMethods',
	--coalesce(p.CashBackFee,0)		'CashBackFee',
	coalesce(l.MaxInstallment,0)	'InstallmentsCount',
	n.ScoreValue					'ScoreValue',
	v.Percentage					'RatingPercentage',
	--b.Name							'CustomerName',
	b.Birthday						'Birthday',
	--u.Street						'Street',
	--u.City							'City',
	--u.Zip							'ZIP',
	w.Name							'Country',
    case when n.ScoreValue < '450' then '< 450' when n.ScoreValue between  '450' and '475' then '450-475' when n.ScoreValue between  '475' and '500' then '475-500'																								
	when n.ScoreValue between  '500' and '525' then '500-525' when n.ScoreValue between  '525' and '550' then '525-550' when n.ScoreValue between  '550' and '575' then '550-575'																								
	when n.ScoreValue between  '575' and '600' then '575-600' when n.ScoreValue between  '600' and '625' then '600-625' when n.ScoreValue between '625' and '650' then '625-650' 																								
	 when n.ScoreValue > '650' then '> 650' else 'NA' end as  ScoreGroup	
	 , Duedate = cast(i.InvoiceDate + i.TermOfPaymentDays as date)
	 , case when cast(i.InvoiceDate + i.TermOfPaymentDays as date) <= @Ultimo then 'Yes' else 'No' end as 'FlagDue'
	 , case when coalesce(Inkasso,0) = 0  and cast(i.InvoiceDate + i.TermOfPaymentDays + 60 as date) <= @Ultimo and cast(i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0) as money) > 0 then 1 else 0 end as Overaged
/* + INVOCES				*/ from Invoices i

/* - PAYMENTS				*/ left join (select InvoiceId, sum(Amount) Amount, sum(PaymentFee) CashBackFee from Payments where PaymentType = 0 and PaymentDate <= @Ultimo group by InvoiceId) p on p.InvoiceId = i.Id
/* - CREDIT NOTES			*/ left join (select InvoiceId, sum(Amount) Amount from Payments where PaymentType = 1 and PaymentDate <= @Ultimo group by InvoiceId) c on c.InvoiceId = i.Id
/* + REMINDERS - AGREEMENTS	*/ 
					left join (select InvoiceId, max(Id) ReminderId, sum(case when AgreementAmount is null then ReminderFee else AgreementAmount end) Amount, sum(ReminderFee) ReminderAmount, 
					sum(AgreementAmount) AgreementAmount, max(ReminderCount) MaxReminder, max(ExportedTimeStamp) InkassoDate, max(cast(Exported as varchar)) Inkasso 
					from Reminders group by InvoiceId) r on r.InvoiceId = i.Id
/* + REMINDERS - AGREEMENTS	*/
					left join (select InvoiceId, max(Id) ReminderId, sum(ReminderFee) ReminderAmount 					 
					from Reminders where timestamp < @Ultimo group by InvoiceId) r2 on r2.InvoiceId = i.Id 					 
		
					left join (select ReminderId, max(TimeStamp) AgreementDate from ReminderAgreementHistories 
					where TimeStamp < @Ultimo group by ReminderId) h on h.ReminderId = r.ReminderId
/* --- OrderDetails			*/ join CustomerOrders o on o.id = i.CustomerOrderId join Currencies e on e.Id = o.CurrencyId join PaymentOptions y on y.Id = o.PaymentOptionId 
left join (select InvoiceId, max(InstallmentNumber) MaxInstallment from Installments group by InvoiceId) l on l.InvoiceId = i.Id
/* --- ParnerDetails		*/ join Partners a on a.Id = o.DistributorId join Partners b on b.Id = o.CustomerId join Partners_Distributor d on d.Id = a.Id 
left join DistributorGroups g on d.DistributorGroupId = g.Id 
join (select PartnerId, Street, City, Zip, CountryId from Addresses where AddressTypeId = 1) u on u.PartnerId = b.Id 
join Countries w on w.Id = u.CountryId
/* --- ClientScores			*/ left join ClientScores n on n.ClientScoreId = o.ClientScoreId 
							left join RiskConversionRates v on v.RiskFinalValue = n.ScoreValue
/*----InkassoEinstellungen  */ --left join (select distinct(grundlagennummer) as grundlagennummer, max(Einstellungsdatum) as max_data from   [FineTrade.Reporting].dbo.InkassoEinstellungen group by Grundlagennummer)  ie on ie.Grundlagennummer=i.InvoiceNumberFull
/*-----CustomerOrders       */ left join Customerorders co on co.id=i.customerorderid 
                               left join ( select i.id, sum(p.Amount) as PaidInk from Invoices i join Payments p on p.InvoiceId = i.Id 	join Reminders r on r.InvoiceId = i.Id																								
				where  p.PaymentDate between r.ExportedTimeStamp and @Ultimo and PaymentType = 0 and r.ExportedTimeStamp is not null	group by i.Id) pa on pa.Id = i.Id	


where   i.FactoryBankId = 2 and i.InvoiceDate <= @Ultimo
  
  ---and (i.PaymentCompleted > @Ultimo or i.PaymentCompleted is null)  --i.OrderTotal-coalesce(p.Amount,0)-coalesce(c.Amount,0)+coalesce(r.Amount,0) >0 
--and y.Id in ( 4,5,9,12,14,16,17,19)
--i.SellDate <= @Ultimo and i.FactoryBankId = 2 and
--and  w.Name in ('Switzerland', 'Austria', 'Germany', 'Spain')  i.Sold = 0  and 
--is null --= 1 -- (i.PaymentCompleted > @Ultimo or i.PaymentCompleted is null) and
order by i.InvoiceDate asc

"""


# In[3]:


data = pd.read_sql(query,conn)
print(data.head())


# In[4]:


data.to_excel(r"C:\NhanLeDocomoDigital\Automation\SQL\REB_Live_Pft_Data201911.xlsx") 


# In[6]:


data.describe()


# In[7]:


data.info()


# In[9]:


pivotOverview = data.pivot_table(index=['InkassoYN', 'Overaged'], values=['InvoiceNumberFull','Inv_netCN',"OpenAmount_WOReminder","InkassoHandover_WOReminder", "InkassoDueAmount_WOReminder", "PaymentsAmount"], aggfunc={"InvoiceNumberFull":'count',"Inv_netCN":"sum","OpenAmount_WOReminder":'sum', "InkassoHandover_WOReminder":'sum', "InkassoDueAmount_WOReminder":'sum', "PaymentsAmount":'sum'},fill_value=0)
print(pivotOverview)


# In[10]:


pivotOverview.index


# In[11]:


idx = pd.IndexSlice
pivotOverview.loc[idx['NO', 1], idx['InvoiceNumberFull']]


# In[12]:


col_names1 =  ['Invoices', 'AmountNet_CN', 'Outstanding', 'Collection Rate']
index1 = ['Performing', "Collection", "Overaged" ]


# In[13]:


df_overview  = pd.DataFrame(columns = col_names1, index = index1)


# In[14]:


df_overview


# In[15]:


df_overview.index


# In[16]:


df_overview.loc['Performing',"Invoices"] = pivotOverview.loc[idx['NO', 0], idx['InvoiceNumberFull']]


# In[17]:


df_overview


# In[18]:


df_overview.loc['Performing',"AmountNet_CN"] = pivotOverview.loc[idx['NO', 0], idx['Inv_netCN']]


# In[19]:


df_overview.loc['Performing',"Outstanding"] = pivotOverview.loc[idx['NO', 0], idx['OpenAmount_WOReminder']]


# In[20]:


df_overview.loc['Performing',"Collection Rate"] = (df_overview.loc['Performing','AmountNet_CN']-df_overview.loc['Performing','Outstanding']) / df_overview.loc['Performing','AmountNet_CN'] 


# In[21]:


df_overview


# In[22]:


df_overview.loc['Collection',"Invoices"] = pivotOverview.loc[idx['YES', 0], idx['InvoiceNumberFull']]
df_overview.loc['Collection',"AmountNet_CN"] = pivotOverview.loc[idx['YES', 0], idx['Inv_netCN']]
df_overview.loc['Collection',"Outstanding"] = pivotOverview.loc[idx['YES', 0], idx['OpenAmount_WOReminder']]
df_overview.loc['Collection',"Collection Rate"] = (df_overview.loc['Collection','AmountNet_CN']-df_overview.loc['Collection','Outstanding']) / df_overview.loc['Collection','AmountNet_CN'] 


# In[23]:


df_overview


# In[24]:


df_overview.loc['Overaged',"Invoices"] = pivotOverview.loc[idx['NO', 1], idx['InvoiceNumberFull']]
df_overview.loc['Overaged',"AmountNet_CN"] = pivotOverview.loc[idx['NO', 1], idx['Inv_netCN']]
df_overview.loc['Overaged',"Outstanding"] = pivotOverview.loc[idx['NO', 1], idx['OpenAmount_WOReminder']]
df_overview.loc['Overaged',"Collection Rate"] = (df_overview.loc['Overaged','AmountNet_CN']-df_overview.loc['Overaged','Outstanding']) / df_overview.loc['Overaged','AmountNet_CN'] 


# In[25]:


df_overview


# In[26]:


pivotMerchants = data.pivot_table(index=['DistributorGroup','InkassoYN', 'Overaged'], values=['InvoiceNumberFull','Inv_netCN',"OpenAmount_WOReminder","InkassoHandover_WOReminder", "InkassoDueAmount_WOReminder", "PaymentsAmount"], aggfunc={"InvoiceNumberFull":'count',"Inv_netCN":"sum","OpenAmount_WOReminder":'sum', "InkassoHandover_WOReminder":'sum', "InkassoDueAmount_WOReminder":'sum', "PaymentsAmount":'sum'},fill_value=0)
print(pivotMerchants)


# In[27]:


pivotMerchants.loc[idx['Agility','NO', 0], idx['InvoiceNumberFull']]


# In[28]:


pivotMerchants.loc[idx['Agility','NO'], idx['InvoiceNumberFull']]


# In[29]:


pivotMerchants.loc[idx['Agility','NO'], idx['InvoiceNumberFull']].sum()


# In[30]:


pivotMerchants.loc[idx['Agility'], idx['InvoiceNumberFull']].sum()


# In[32]:


col_names2 =  ['Invoices', 'AmountNet_CN', 'Performing', 'Inkasso Handover',"Outst. in Debt Collection", "Overaged", "Handover Rate", "Collection Rate"]
index2 = ['Agility', "Chronext", "Watchmaster", "Mapa", "Mediashop", "Total Live Portfolio" ]


# In[33]:


df_merchants  = pd.DataFrame(columns = col_names2, index = index2)


# In[34]:


df_merchants.loc['Agility',"Invoices"] = pivotMerchants.loc[idx['Agility'], idx['InvoiceNumberFull']].sum()
df_merchants.loc['Agility',"AmountNet_CN"] = pivotMerchants.loc[idx['Agility'], idx['Inv_netCN']].sum()
df_merchants.loc['Agility',"Performing"] = pivotMerchants.loc[idx['Agility','NO', 0], idx['OpenAmount_WOReminder']]
df_merchants.loc['Agility',"Inkasso Handover"] = pivotMerchants.loc[idx['Agility','YES', 0], idx['InkassoHandover_WOReminder']]
df_merchants.loc['Agility',"Outst. in Debt Collection"] = pivotMerchants.loc[idx['Agility','YES', 0], idx['InkassoDueAmount_WOReminder']]
df_merchants.loc['Agility',"Overaged"] = pivotMerchants.loc[idx['Agility','NO', 1], idx['OpenAmount_WOReminder']]
df_merchants.loc['Agility',"Handover Rate"] = df_merchants.loc['Agility','Inkasso Handover']/ df_merchants.loc['Agility','AmountNet_CN'] 
df_merchants.loc['Agility',"Collection Rate"] = (df_merchants.loc['Agility','Inkasso Handover']-df_merchants.loc['Agility','Outst. in Debt Collection']) / df_merchants.loc['Agility','Inkasso Handover'] 
df_merchants


# In[35]:


df_merchants.loc['Mapa',"Invoices"] = pivotMerchants.loc[idx['MAPA Germany'], idx['InvoiceNumberFull']].sum()
df_merchants.loc['Mapa',"AmountNet_CN"] = pivotMerchants.loc[idx['MAPA Germany'], idx['Inv_netCN']].sum()
df_merchants.loc['Mapa',"Performing"] = pivotMerchants.loc[idx['MAPA Germany','NO', 0], idx['OpenAmount_WOReminder']]
df_merchants.loc['Mapa',"Inkasso Handover"] = pivotMerchants.loc[idx['MAPA Germany','YES', 0], idx['InkassoHandover_WOReminder']]
df_merchants.loc['Mapa',"Outst. in Debt Collection"] = pivotMerchants.loc[idx['MAPA Germany','YES', 0], idx['InkassoDueAmount_WOReminder']]
df_merchants.loc['Mapa',"Overaged"] = pivotMerchants.loc[idx['MAPA Germany','NO', 1], idx['OpenAmount_WOReminder']]
df_merchants.loc['Mapa',"Handover Rate"] = df_merchants.loc['Mapa','Inkasso Handover']/ df_merchants.loc['Mapa','AmountNet_CN'] 
df_merchants.loc['Mapa',"Collection Rate"] = (df_merchants.loc['Mapa','Inkasso Handover']-df_merchants.loc['Mapa','Outst. in Debt Collection']) / df_merchants.loc['Mapa','Inkasso Handover'] 
df_merchants


# In[36]:


df_merchants.loc['Mediashop',"Invoices"] = pivotMerchants.loc[idx['Mediashop Group'], idx['InvoiceNumberFull']].sum()
df_merchants.loc['Mediashop',"AmountNet_CN"] = pivotMerchants.loc[idx['Mediashop Group'], idx['Inv_netCN']].sum()
df_merchants.loc['Mediashop',"Performing"] = pivotMerchants.loc[idx['Mediashop Group','NO', 0], idx['OpenAmount_WOReminder']]
df_merchants.loc['Mediashop',"Inkasso Handover"] = pivotMerchants.loc[idx['Mediashop Group','YES', 0], idx['InkassoHandover_WOReminder']]
df_merchants.loc['Mediashop',"Outst. in Debt Collection"] = pivotMerchants.loc[idx['Mediashop Group','YES', 0], idx['InkassoDueAmount_WOReminder']]
df_merchants.loc['Mediashop',"Overaged"] = pivotMerchants.loc[idx['Mediashop Group','NO', 1], idx['OpenAmount_WOReminder']]
df_merchants.loc['Mediashop',"Handover Rate"] = df_merchants.loc['Mediashop','Inkasso Handover']/ df_merchants.loc['Mediashop','AmountNet_CN'] 
df_merchants.loc['Mediashop',"Collection Rate"] = (df_merchants.loc['Mediashop','Inkasso Handover']-df_merchants.loc['Mediashop','Outst. in Debt Collection']) / df_merchants.loc['Mediashop','Inkasso Handover'] 
df_merchants


# In[37]:


pivotMerchants2 = data.pivot_table(index=['DistributorName','InkassoYN', 'Overaged'], values=['InvoiceNumberFull','Inv_netCN',"OpenAmount_WOReminder","InkassoHandover_WOReminder", "InkassoDueAmount_WOReminder", "PaymentsAmount"], aggfunc={"InvoiceNumberFull":'count',"Inv_netCN":"sum","OpenAmount_WOReminder":'sum', "InkassoHandover_WOReminder":'sum', "InkassoDueAmount_WOReminder":'sum', "PaymentsAmount":'sum'},fill_value=0)
print(pivotMerchants2)


# In[38]:


df_merchants.loc['Chronext',"Invoices"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH'], idx['InvoiceNumberFull']].sum()
df_merchants.loc['Chronext',"AmountNet_CN"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH'], idx['Inv_netCN']].sum()
df_merchants.loc['Chronext',"Performing"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH','NO', 0], idx['OpenAmount_WOReminder']]
df_merchants.loc['Chronext',"Inkasso Handover"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH','YES', 0], idx['InkassoHandover_WOReminder']]
df_merchants.loc['Chronext',"Outst. in Debt Collection"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH','YES', 0], idx['InkassoDueAmount_WOReminder']]
df_merchants.loc['Chronext',"Overaged"] = pivotMerchants2.loc[idx['Chronext Service Germany GmbH','NO', 1], idx['OpenAmount_WOReminder']]
df_merchants.loc['Chronext',"Handover Rate"] = df_merchants.loc['Chronext','Inkasso Handover']/ df_merchants.loc['Chronext','AmountNet_CN'] 
df_merchants.loc['Chronext',"Collection Rate"] = (df_merchants.loc['Chronext','Inkasso Handover']-df_merchants.loc['Chronext','Outst. in Debt Collection']) / df_merchants.loc['Chronext','Inkasso Handover'] 
df_merchants


# In[39]:


df_merchants.loc['Watchmaster',"Invoices"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH'], idx['InvoiceNumberFull']].sum()
df_merchants.loc['Watchmaster',"AmountNet_CN"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH'], idx['Inv_netCN']].sum()
df_merchants.loc['Watchmaster',"Performing"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH','NO', 0], idx['OpenAmount_WOReminder']]
df_merchants.loc['Watchmaster',"Inkasso Handover"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH','YES', 0], idx['InkassoHandover_WOReminder']]
df_merchants.loc['Watchmaster',"Outst. in Debt Collection"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH','YES', 0], idx['InkassoDueAmount_WOReminder']]
df_merchants.loc['Watchmaster',"Overaged"] = pivotMerchants2.loc[idx['Watchmaster ICP GmbH','NO', 1], idx['OpenAmount_WOReminder']]
df_merchants.loc['Watchmaster',"Handover Rate"] = df_merchants.loc['Watchmaster','Inkasso Handover']/ df_merchants.loc['Watchmaster','AmountNet_CN'] 
df_merchants.loc['Watchmaster',"Collection Rate"] = (df_merchants.loc['Watchmaster','Inkasso Handover']-df_merchants.loc['Watchmaster','Outst. in Debt Collection']) / df_merchants.loc['Watchmaster','Inkasso Handover'] 
df_merchants


# In[41]:


df_merchants


# In[42]:


df_merchants.loc['Total Live Portfolio',"AmountNet_CN"] = df_merchants["AmountNet_CN"].sum()
df_merchants.loc['Total Live Portfolio',"Performing"] = df_merchants["Performing"].sum()
df_merchants.loc['Total Live Portfolio',"Inkasso Handover"] = df_merchants["Inkasso Handover"].sum()
df_merchants.loc['Total Live Portfolio',"Outst. in Debt Collection"] = df_merchants["Outst. in Debt Collection"].sum()
df_merchants.loc['Total Live Portfolio',"Overaged"] = df_merchants["Overaged"].sum()
df_merchants.loc['Total Live Portfolio',"Handover Rate"] = df_merchants.loc['Total Live Portfolio','Inkasso Handover']/ df_merchants.loc['Total Live Portfolio','AmountNet_CN'] 
df_merchants.loc['Total Live Portfolio',"Collection Rate"] = (df_merchants.loc['Total Live Portfolio','Inkasso Handover']-df_merchants.loc['Total Live Portfolio','Outst. in Debt Collection']) / df_merchants.loc['Total Live Portfolio','Inkasso Handover']
df_merchants


# In[44]:


data=pd.read_excel(r"C:\NhanLeDocomoDigital\Automation\REB report\Capital Employed 201911.xlsx", index_col="Date")


# In[45]:


data


# In[47]:


import matplotlib.pyplot as plt
xs = data.index.values
ys = data["Captial Employed"].values
lines = data[["Open Invoices", "Installments", "Captial Employed"]].plot.line(figsize = (20,10), title = "Capital Employed", legend = True )


for x,y in zip(xs,ys):

    label = "{:.2f}".format(y/1000000)

    lines.annotate(label, # this is the text
                 (x,y), # this is the point to label
                 textcoords="offset points", # how to position the text
                 xytext=(0,10), # distance from text to points (x,y)
                 ha='center') # horizontal alignment can be left, right or center

lines


# In[49]:


lines.figure.savefig(r"C:\NhanLeDocomoDigital\Automation\REB report\capitalEmployed_Graph.png")


# In[50]:


with pd.ExcelWriter(r"C:\NhanLeDocomoDigital\Automation\REB report\REB live Pft 201911.xlsx") as writer:  # doctest: +SKIP
    df_overview.to_excel(writer, sheet_name='overview')
    df_merchants.to_excel(writer, sheet_name='merchants')


# In[51]:


from openpyxl import load_workbook
import openpyxl
#create workbook object
wb = load_workbook(r"C:\NhanLeDocomoDigital\Automation\REB report\REB live Pft 201911.xlsx")
ws = wb.active
img = openpyxl.drawing.image.Image(r"C:\NhanLeDocomoDigital\Automation\REB report\capitalEmployed_Graph.png")
ws.add_image(img, 'B9')
wb.save(r"C:\NhanLeDocomoDigital\Automation\REB report\REB live Pft 201911.xlsx")                  


# In[ ]:




  


# In[ ]:





# In[ ]:




