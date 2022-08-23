import pandas as pd
import numpy


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


# 按比例分配给每个channel*province
for i in range(len(all_prov)-1):
    for n in range (len(all_prov[i])):
        all_prov[i][n] = all_prov[i][n] * sum_prov[i]


# 每个channel目标的和
print(all_prov[-1])



# 现在每个channel的和
sum_channel = []
for i in range(len(columns)):  
    channel =  0
    for n in range(len(all_prov)-1):
        channel += all_prov[n][i]
    sum_channel.append(channel)
print(sum_channel)

# 从最少的channel开始更新，更新顺序为order中的0-6
order=[]
min_index = []
sort = sorted(sum_channel)
for i in range(len(sum_channel)):
    order.append(sort.index(sum_channel[i]))

for i in range(len(sum_channel)-1):
    min_index.append(order.index(i))


for index in min_index:
    target = all_prov[-1][index]
    # print('target',target)
    coefficient = target/sum_channel[index]
    # print(coefficient)
    for n in range(len(all_prov)-1):
        all_prov[n][index] = all_prov[n][index] * coefficient

