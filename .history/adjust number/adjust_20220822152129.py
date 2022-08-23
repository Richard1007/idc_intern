import pandas as pd
data = pd.read_excel('adjust number/adjust.xlsx')
print(data)

for index, row in data.iterrows():
    print(row['Direct - Inbound/Outbound'], row['Direct - Internet'])