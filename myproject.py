from functions import list_to_excel
import pandas as pd

my_data = pd.read_excel('data1.xlsx')
my_list = my_data.values.tolist()
my_header = my_data.columns.tolist()
headers =[]
for item in my_header:
    headers.append({'header': item})
temp = []
typ = []
for row in my_list:
    typ.append(row[1])

unique_type = set(typ)

for item in unique_type:
    temp=[]
    for row in my_list:
        if row[1] == item:
            temp.append(row)

    list_to_excel(temp,headers,item)
