import pandas as pd
import numpy

# Input: channelmix, the target sum of each province and channel (topline) 
# Read file, columns consists of all channels
data = pd.read_excel('adjust number/adjust.xlsx')
columns = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI','InDirect - eTailer','InDirect - LFR','InDirect - Retail','InDirect - Telco']
all_prov = []
sum_prov = []

# all_prov is a list of lists, each sub-list consists of the value of all channels of a province
# len(all_prov) = # of province +1, because the last line is the target sum of each channel (topline)
# sum_prov is a list consisting of the sum of each province (topline)，len(sum_prov) = # of province 
for index, row in data.iterrows():
    province = []
    for column in columns:
        province.append(row[column])
    all_prov.append(province)
    sum_prov.append(row['sum_province'])


# Distribute based on channelmix and province target sum 
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


# 从最少的channel开始更新，更新顺序为order中的0-6, 最大的channel剩着，用每个省的差值计算（每个省其余七个channel的数值都已经确定）
# 更新顺序为从最小的channel到大的channel依次更新，min_index是更新的column顺序
order=[]
min_index = []
sort = sorted(sum_channel)
for i in range(len(sum_channel)):
    order.append(sort.index(sum_channel[i]))
for i in range(len(sum_channel)):
    min_index.append(order.index(i))

# 按照channel总和算出需要的coefficient，相乘之后得到满足条件的sum of a channel 
for index in min_index[:-1]:
    target = all_prov[-1][index]
    # print('target',target)
    coefficient = target/sum_channel[index]
    # print(coefficient)
    for n in range(len(all_prov)-1):
        all_prov[n][index] = all_prov[n][index] * coefficient


# The biggest channel is not yet distributed, last是最大的channel的index
# 最后一个channel的更新逻辑是：从每个province的topline减去已有的七个channel
last = min_index[-1]

for i in range(len(all_prov)-1):
    other_channel_sum = 0
    for n in range(len(all_prov[-1])):
        if n != last:
            other_channel_sum += all_prov[i][n]
    all_prov[i][last] = sum_prov[i] - other_channel_sum


# Export data
import xlsxwriter

workbook = xlsxwriter.Workbook('adjust number/output.xlsx')
worksheet = workbook.add_worksheet()

for i in range(len(all_prov)-1):
    for n in range(len(all_prov[-1])):
        worksheet.write(i, n, all_prov[i][n])

workbook.close()

# Potential Problems: negaive numbers, decimal places 