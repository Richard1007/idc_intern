import pandas as pd
import numpy
import xlsxwriter

# Input: channelmix, the target sum of each province and channel (topline) 
# 由于直接读取excel，columns需要按照input的先后顺序写
data = pd.read_excel('adjust number/adjust.xlsx')
columns = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI','InDirect - eTailer','InDirect - LFR','InDirect - Retail','InDirect - Telco']
all_prov = []
sum_prov = []

# all_prov is a list of lists, consisting of the value of every channel of a province
# len(all_prov) = # of province +1, since the last line is the target sum of each channel (topline)
# sum_prov is a list consisting of the target sum of each province (another topline)，len(sum_prov) = # of province 
for index, row in data.iterrows():
    province = []
    for column in columns:
        province.append(row[column])
    all_prov.append(province)
    sum_prov.append(row['sum_province'])


# Distribute based on channelmix and target sum of every province
for i in range(len(all_prov)-1):
    for n in range (len(all_prov[i])):
        all_prov[i][n] = all_prov[i][n] * sum_prov[i]

print('Target sum of every province:',sum_prov)

# target sum of every channel (topline)
target_channel = all_prov[-1]
print('Target sum of every channel :',target_channel)


# 此时每个channel的和
sum_channel = []
for i in range(len(columns)):  
    channel =  0
    for n in range(len(all_prov)-1):
        channel += all_prov[n][i]
    sum_channel.append(channel)
print('此时每个channel的和:',sum_channel)


# 更新顺序为从最小的channel到大的channel依次更新，按现在channel sum和target sum的算出来需要的coefficient，再将channel的每个数都*coefficient
# order是现在channel sum的大小顺序，min_index是需要被更新的column顺序
order=[]
min_index = []
sort = sorted(sum_channel)
for i in range(len(sum_channel)):
    order.append(sort.index(sum_channel[i]))
for i in range(len(sum_channel)):
    min_index.append(order.index(i))
print('此时channel sum的大小:',order,'更新顺序:', min_index)


# 按照channel总和算出需要的coefficient，相乘之后得到满足条件的sum of a channel 
for index in min_index[:-1]:
    target = all_prov[-1][index]
    # print('target',target)
    coefficient = target/sum_channel[index]
    print('channel需要*的系数:', coefficient)
    
    for n in range(len(all_prov)-1):
        all_prov[n][index] = all_prov[n][index] * coefficient


# The largest channel is not yet distributed, last是最大的channel的index
# 最后一个channel的更新逻不同：从每个province的topline减去已有的七个channel
last = min_index[-1]

for i in range(len(all_prov)-1):
    other_channel_sum = 0
    for n in range(len(all_prov[-1])):
        if n != last:
            other_channel_sum += all_prov[i][n]
    all_prov[i][last] = sum_prov[i] - other_channel_sum


# Export data， output和input province，channel顺序一致
workbook = xlsxwriter.Workbook('adjust number/output.xlsx')
worksheet = workbook.add_worksheet()

for i in range(len(all_prov)-1):
    for n in range(len(all_prov[-1])):
        worksheet.write(i, n, all_prov[i][n])
workbook.close()

# Potential Problems: negaive numbers, decimal places 