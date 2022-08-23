import pandas as pd
data = pd.read_excel('adjust number/adjust.xlsx')
print(data)
columns = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI','InDirect - eTailer','InDirect - LFR','InDirect - Retail','InDirect - Telco']
all_province = []

for index, row in data.iterrows():
    province = []
    for column in columns:
        province.append(row[column])
    all_province.append(province)

print(len(all_province))