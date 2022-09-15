from ferpy.ferpy import FerPy
import pandas as pd
import os
import sys

os.chdir(r'C:\Users\mjia\Desktop\Vincy\IB')

def get_flat_file(ferda, logger, selected_attributes, export_filter, figures):
    if 'Country' not in selected_attributes:
        selected_attributes.append('Country')
    flat = ferda.flatfile_export(selected_attributes=selected_attributes,
                                 user_filter=export_filter,
                                 figures=figures)
    if flat.get("state") != "DONE":
        logger.error("Flatfile export failed due to: %s.", flat.get("error"))
        ferda.fail("Flatfile export failed due to: {}".format(flat.get("error")))
    data = flat.get("files")
    return data

# ferda = FerPy("PRODUCTION", "user_properties.ini", "mobile", log_to_console = True)

ferda = FerPy("./conn.json") # connect to Mobile Phone Hist

attributes = ["Country", "Product Category", "OS", "Brand", "Company", "Quarter", "Year"]
user_filter = {"Country": "PRC", "Product Category": "Smartphone", "Year": [2014,2015,2016,2017,2018,2019,2020,2021,2022]}
figures = ["Units"]
data = get_flat_file(ferda, ferda.logger, attributes, user_filter, figures)
print(data.Quarter.value_counts())

def to_list(obj):
    return obj if isinstance(obj, list) else [obj]

# def get_previous_quarter(quarter: str):
#     """ Returns previous quarter for given quarter as string """
#     if quarter[-1] == "1":
#         return str(int(quarter[0:4]) - 1) + "Q4"
#     return quarter[0:5] + str(int(quarter[-1]) - 1)

# def get_requested_quarters(n: int, quarter_list):
#     """Returns the previous n quarters of a given quarter or quarter list"""
#     first_quarter = to_list(quarter_list)[0]
#     result = [get_previous_quarter(first_quarter)]
#     for i in range(n-1):
#         result.append(get_previous_quarter(result[-1]))
#     return sorted(result)

def get_quarter_distance(ib_quarter, exact_quarter):
    return (int(ib_quarter[:4])-int(exact_quarter[:4])) * 4 + (int(ib_quarter[-1]) - int(exact_quarter[-1])) + 1

def calculate_IB(df):
    if df['quarter_distance'] <= 0:
        return 0
    res = df['life_cycle'] - df['quarter_distance']
    if res >= 0:
        return df['Units']
    elif res < 0 and res > -1:
        return df['Units'] * (-res)
    else:
        return 0

def generate_IB_for_quarter(quarter, full_data):
    life_cycle = pd.read_excel('IB factor - 20220831.xlsx', sheet_name = 'Year')
    life_cycle['life_cycle'] *= 4
    android_life = life_cycle.drop(life_cycle[life_cycle['Brand'] == 'Apple'].index)
    android_mean_life = android_life.groupby(['Year']).mean().reset_index()
    full_data['quarter_distance'] = full_data['Quarter'].apply(lambda x: get_quarter_distance(quarter, x))
    full_data['Year'] = full_data['Year'].astype(float)
    full_data = full_data.merge(life_cycle, on = ['Brand', 'Year'], how = 'left')
    others_data = full_data[full_data['life_cycle'].isna()]
    others_data = others_data.drop(columns = ['life_cycle']).merge(android_mean_life, on = 'Year', how ='left').dropna()
    full_data = full_data.dropna().append(others_data)
    full_data['IB Quarter'] = quarter
    full_data['Installed Base'] = full_data.apply(lambda x: calculate_IB(x), axis = 1)
    pivot = pd.pivot_table(full_data,
                           values = ['Installed Base'],
                           index = ["Country", "Product Category", "OS", "Brand", "Company","IB Quarter"],
                           aggfunc = 'sum').reset_index()
    return pivot

quarters = ['2017Q1', '2017Q2', '2017Q3', '2017Q4',
            '2018Q1', '2018Q2', '2018Q3', '2018Q4',
            '2019Q1', '2019Q2', '2019Q3', '2019Q4',
            '2020Q1', '2020Q2', '2020Q3', '2020Q4',
            '2021Q1', '2021Q2', '2021Q3', '2021Q4',
            '2022Q1','2022Q2']
           # , '2022Q2', '2022Q3', '2022Q4',
           #'2023Q1', '2023Q2', '2023Q3', '2023Q4',
           #'2024Q1', '2024Q2', '2024Q3', '2024Q4',
           #'2025Q1', '2025Q2', '2025Q3', '2025Q4',
           #'2026Q1', '2026Q2', '2026Q3', '2026Q4',
           #'2027Q1', '2027Q2', '2027Q3', '2027Q4',
           #'2028Q1', '2028Q2', '2028Q3', '2028Q4',
           #'2029Q1', '2029Q2', '2029Q3', '2029Q4',
           #'2030Q1', '2030Q2', '2030Q3', '2030Q4']


res = pd.DataFrame()

for quar in quarters:
    temp = data.copy()
    ib_data = generate_IB_for_quarter(quar, temp)
    res = res.append(ib_data)

print("Execution Completed")
res.to_excel("result_IB - 20220831.xlsx", index=False)

#"For Testing Purpose"
# df = pd.DataFrame()

# temp = data.copy()

# life_cycle = pd.read_excel('IB factor - 1123.xlsx', sheet_name = 'Year') # User-defined EXCEL file
# life_cycle['life_cycle'] *= 4
# life_cycle.to_excel('life_cycle_1123.xlsx')

# android_life = life_cycle.drop(life_cycle[life_cycle['Brand'] == 'Apple'].index)
# android_mean_life = android_life.groupby(['Year']).mean().reset_index() # calculate the mean value for Android brands

# temp['quarter_distance'] = temp['Quarter'].apply(lambda x: get_quarter_distance('2020Q3', x))
# temp['Year'] = temp['Year'].astype(int)
# temp = temp.merge(life_cycle, on = ['Brand', 'Year'], how = 'left')

# others_data = temp[temp['life_cycle'].isna()]

# others_data = others_data.drop(columns = ['life_cycle'])
# others_data = others_data.merge(android_mean_life, on = 'Year', how ='left').dropna()

# temp = temp.dropna()
# temp = temp.append(others_data)
# temp['IB Quarter'] = '2020Q3'

# temp['res'] = temp['life_cycle'] - temp['quarter_distance']

# temp.to_excel('temp20Q3before calcu-1123.xlsx')
# temp['Installed Base'] = temp.apply(lambda x: calculate_IB(x), axis = 1)
# temp.to_excel('temp20Q3after calcu-1123.xlsx')


# pivot = pd.pivot_table(temp,
#                        values = ['Installed Base'],
#                        index = ["Country", "Product Category", "OS", "Brand", "Company","IB Quarter"],
#                        aggfunc = 'sum').reset_index()
