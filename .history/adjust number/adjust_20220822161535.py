import pandas as pd

# Read file
data = pd.read_excel('adjust number/adjust.xlsx')
columns = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI','InDirect - eTailer','InDirect - LFR','InDirect - Retail','InDirect - Telco']
all_prov = []
sum_prov = []

for index, row in data.iterrows():
    province = []
    for column in columns:
        province.append(row[column])
    all_prov.append(province)
    sum_prov.append(row['sum_province'])
print(all_prov)
print(sum_prov)

for i in range(len(all_prov)-1):
    for n in range (len(all_prov[i])):
        all_prov[i][n] = all_prov[i][n] * sum_prov[i]
        
print(all_prov)

