import pandas as pd

#Open Csv to get encode so file can be opened
data_encode = open('E-Commerce.csv',mode='r')
data = pd.read_csv('E-Commerce.csv',encoding = data_encode.encoding)

#Lower the columns name
data.columns = [x.lower() for x in data.columns]

#Remove the data that has potential as fraud data
data = data.dropna(subset=['description','customerid'])
data = data.drop_duplicates(keep='first')

#There are some data that has the same description but has one separator that makes it different
data['description'] = data['description'].replace(',','',regex=True)

#Change type of some data
data['invoicedate'] = data['invoicedate'].astype('Datetime64')
data['customerid'] = data['customerid'].astype('int64')

#Make some additional data to analyze data easier
data['totalprice'] = (data['unitprice']*data['quantity']).round(2)
data['month'] = data['invoicedate'].dt.month_name()
data['day'] = data['invoicedate'].dt.day_name()
data['hour'] = data['invoicedate'].dt.hour
def hour_group (x):
    if x < 6:
        return '00.00 - 05.59'
    elif x >= 6 and x < 12:
        return '06.00 - 11.59'
    elif x >= 12 and x < 18:
        return '12.00 - 17.59'
    else:
        return '18.00 - 23.59'

data['hour'] = data['hour'].apply(hour_group)

#Filter data that customer has canceled the purchases and customer got discount data_cancel (cancel transaction)
data_cancel = data[(data['invoiceno'].str.contains('C')) & (data['stockcode']!= 'D')]
 
#Group the data_cancel to look sum of quantity and mean of unit price based on description and country
data_trans_cancel = data_cancel.groupby(['description','customerid']).agg({
    'quantity' : 'sum',
    'totalprice' : 'mean'
}).abs().round(2).reset_index()

#In this data, there are some discounts that has given so divide the data to look the given discount
data_discount = data[(data['invoiceno'].str.contains('C')) & (data['stockcode']== 'D')]

#Group the data_buy to look sum of quantity and mean of unit price based on customerid and month
data_get_discount = data_discount.groupby(['customerid','month']).agg({
    'quantity' : 'sum',
    'totalprice' : 'mean'
}).abs().round(2).reset_index()

#Finally store all of data to excel file
writer = pd.ExcelWriter('E-Commerce.xlsx', engine = 'xlsxwriter')
data.to_excel(writer, sheet_name = 'dataecommerce', index=False)
data_trans_cancel.to_excel(writer, sheet_name = 'datatranscancel', index=False)
data_get_discount.to_excel(writer, sheet_name = 'datagetdiscount', index=False)
writer.save()
