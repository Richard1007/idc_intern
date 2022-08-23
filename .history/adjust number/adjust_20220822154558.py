import pandas as pd

# Read file
data = pd.read_excel('adjust number/adjust.xlsx')
print(data)
columns = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI','InDirect - eTailer','InDirect - LFR','InDirect - Retail','InDirect - Telco']
all_province = []
prov_sum = []

for index, row in data.iterrows():
    province = []
    for column in columns:
        province.append(row[column])
    all_province.append(province)
    prov_sum.append(row['sum_province'])

print(len(all_province))
print(prov_sum)

for i in range(len(all_province)-1):
    for element in all_province[i]:
        element = element*prov_sum[i]
print(all_province)