from thinkcellbuilder import Presentation, Template
import pandas as pd
from datetime import datetime
from builder import Builder
from thinkcell import Thinkcell
from write_excel import Write_Excel
from Office365_API import SharePoint
import re
import sys,os
import xlwings as xw
from pathlib import PurePath

#Downloading Files from Sharepoint

FOLDER_NAME = sys.argv[1]

FOLDER_DEST = sys.argv[2]

FILE_NAME = sys.argv[3]

FILE_NAME_PATTERN = sys.argv[4]

def save_file(file_n, file_obj):
    file_dir_path = PurePath(FOLDER_DEST,file_n)
    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)

def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n,folder)
    save_file(file_n,file_obj)

def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)

def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name,folder)

if __name__ == "__main__":
    if FILE_NAME != 'None':
        get_file(FILE_NAME,FOLDER_NAME)
    elif FILE_NAME_PATTERN != 'None':
        get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
    else:
        get_files(FOLDER_NAME)

#Updating Charts


file_path = r"C:\Users\imran.s\Desktop\POC Disney\Thinkcell_Automation\storage\20240528_Weekly_Leads_Summary_0525_v3.xlsb"
template_file_path = r"C:\Users\imran.s\Desktop\POC Disney\Thinkcell_Automation\storage\Weekly Leads Summary Templates (1).xlsb"

wb = xw.Book(file_path)
sht = wb.sheets['Sheet1']
new_wb = xw.Book(r'C:\new_path\to\file.xlsx')
new_wb.sheets['name'] = sht

file_name_1 = r"storage\20240528_Weekly_Leads_Summary_0525_v3.xlsb" 
sheet_name_1 = 'By Marketing Channel (TEMPLATE)'
sheet_name_2 = 'National Monthly'
Write_Excel().close_all_excel_instances()
Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS')

df1 = Builder().read_excel(file_name_1, sheet_name_1)
df2 =Builder().read_excel(file_name_1,sheet_name_2)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

df1.columns = custom_column_names_df1
df2.columns = custom_column_names_df2

data_for_chart1 = Builder().extract_data(df1, 'C', 'P', 20, 26)

data_for_chart1 = Builder().add_row(df1,data_for_chart1,28,'C','P','D')

data_for_chart1 = Builder().add_row(df1, data_for_chart1, 52, 'C', 'P', 'D')

updated_column_names = Builder().dates(df1,18, 'D','P')

converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

data_for_chart1.columns = [data_for_chart1.columns[0]]+formated_updated_column_names




# For Chart 2

data_for_chart2 = Builder().extract_data(df1, 'K', 'P', 32, 38)

column_list = Builder().add_column(df1, 'D', 32,39 )

data_for_chart2.insert(loc=0,column = 'D', value = column_list )

column_names = Builder().add_column(df1, 'C', 32, 39)

data_for_chart2.insert(loc=0,column='C',value=column_names)

updated_column_names_chart2 = Builder().dates(df1,30, 'K','P')
updated_column_names1_chart2 = df1.loc[30, "D"]
updated_column_names1_chart2_list = [updated_column_names1_chart2]

converted_updated_column_names_chart2 = Builder().convert_to_date_time(updated_column_names_chart2)
converted_updated_column_names1_chart2 = Builder().convert_to_date_time(updated_column_names1_chart2_list)

formated_updated_column_names_chart2 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart2]
formated_updated_column_names1_chart2 = [Builder().format_date_time(d) for d in converted_updated_column_names1_chart2]

data_for_chart2.columns = [data_for_chart2.columns[0]]+formated_updated_column_names1_chart2+formated_updated_column_names_chart2



# For Chart3

data_for_chart3 = Builder().extract_data(df1, 'K', 'P', 60, 61)

data_for_chart3 = Builder().add_row(df1, data_for_chart3, 64, 'K','P',None)
data_for_chart3 = Builder().add_row(df1,data_for_chart3,65,'K','P',None)

column_list_chart3 = Builder().add_column(df1, 'D', 60,62 )

column_list_chart3.append(df1.loc[64,'D'])
column_list_chart3.append(df1.loc[65,'D'])


data_for_chart3.insert(loc=0,column = 'D', value = column_list_chart3 )

column_names_chart3 = Builder().add_column(df1, 'C', 60, 62)
column_names_chart3.append(df1.loc[64,'C'])
column_names_chart3.append(df1.loc[65,'C'])

data_for_chart3.insert(loc=0,column='C',value=column_names_chart3)

length = len(data_for_chart3.loc[60])
for i in range(1,length):
    data_for_chart3.iloc[:, i] = data_for_chart3.iloc[:, i].apply(lambda x: float(f"{x * 100:.1f}"))

updated_column_names_chart3 = Builder().dates(df1,58, 'K','P')
updated_column_names1_chart3 = df1.loc[58, "D"]
updated_column_names1_chart3_list = [updated_column_names1_chart3]

converted_updated_column_names_chart3 = Builder().convert_to_date_time(updated_column_names_chart3)
converted_updated_column_names1_chart3 = Builder().convert_to_date_time(updated_column_names1_chart3_list)

formated_updated_column_names_chart3 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart3]
formated_updated_column_names1_chart3 = [Builder().format_date_time(d) for d in converted_updated_column_names1_chart3]

data_for_chart3.columns = [data_for_chart3.columns[0]]+formated_updated_column_names1_chart3+formated_updated_column_names_chart3


#For Chart4

data_for_chart4 = Builder().extract_data(df1, 'C', 'P', 81, 86)

data_for_chart4 = data_for_chart4.drop(index = 82)

updated_column_names_chart4 = Builder().dates(df1,79, 'D','P')

converted_updated_column_names_chart4 = Builder().convert_to_date_time(updated_column_names_chart4)

formated_updated_column_names_chart4 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart4]

data_for_chart4.columns = [data_for_chart4.columns[0]]+formated_updated_column_names_chart4

#For Chart5

data_for_chart5 = Builder().extract_data(df2, 'AG', 'AS', 87, 88)

data_for_chart5 = data_for_chart5.drop(index = 88)

data_for_chart5 = Builder().add_row(df2, data_for_chart5, 59, 'AG','AS',None)

row1 = df2.loc[59, 'AG':'AS']
row2 = df2.loc[54, 'AG':'AS']
resultant_row = row2-row1
data_for_chart5.loc[60] = resultant_row

names = ['Web Leads',
'IB Call - Online',
'IB Calls - DM/Demand'
]
data_for_chart5.insert(loc=0,column = 'C', value = names )
updated_column_names_chart5 = Builder().dates(df2,3, 'AG','AS')

converted_updated_column_names_chart5 = Builder().convert_to_date_time(updated_column_names_chart5)

formated_updated_column_names_chart5 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart5]

data_for_chart5.columns = [data_for_chart5.columns[0]]+formated_updated_column_names_chart5

#For Chart6

data_for_chart6 = Builder().extract_data(df1, 'C', 'P', 32, 38)

data_for_chart6 = Builder().add_row(df1,data_for_chart6,52,'C','P','D')

data_for_chart6 = Builder().add_row(df1, data_for_chart6, 28, 'C', 'P', 'D')

updated_column_names_chart6 = Builder().dates(df1,30, 'D','P')

converted_updated_column_names_chart6 = Builder().convert_to_date_time(updated_column_names_chart6)

formated_updated_column_names_chart6 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart6]

data_for_chart6.columns = [data_for_chart6.columns[0]]+formated_updated_column_names_chart6

#For Chart7

data_for_chart7 = Builder().extract_data(df1, 'C', 'P', 81, 86)

data_for_chart7 = data_for_chart7.drop(index = 82)

data_for_chart7 = Builder().add_row(df1,data_for_chart7,56,'C','P','D')

data_for_chart7 = Builder().add_row(df1, data_for_chart7, 59, 'C', 'P', 'D')

updated_column_names_chart7 = Builder().dates(df1,79, 'D','P')

converted_updated_column_names_chart7 = Builder().convert_to_date_time(updated_column_names_chart7)

formated_updated_column_names_chart7 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart7]

data_for_chart7.columns = [data_for_chart7.columns[0]]+formated_updated_column_names_chart7

#For Chart8

data_for_chart8 = Builder().extract_data(df1, 'C', 'P', 157, 162)

data_for_chart8 = data_for_chart8.drop(index = 158)

data_for_chart8 = Builder().add_row(df1,data_for_chart8,131,'C','P','D')

data_for_chart8 = Builder().add_row(df1, data_for_chart8, 153, 'C', 'P', 'D')

updated_column_names_chart8 = Builder().dates(df1,155, 'D','P')

converted_updated_column_names_chart8 = Builder().convert_to_date_time(updated_column_names_chart8)

formated_updated_column_names_chart8 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart8]

data_for_chart8.columns = [data_for_chart8.columns[0]]+formated_updated_column_names_chart8


#For Chart9

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS')

df1 = Builder().read_excel(file_name_1, sheet_name_1)
df2 =Builder().read_excel(file_name_1,sheet_name_2)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

df1.columns = custom_column_names_df1
df2.columns = custom_column_names_df2

data_for_chart9 = Builder().extract_data(df1, 'C', 'P', 60, 65)
data_for_chart9 = data_for_chart9.drop(index = 62)
data_for_chart9 = data_for_chart9.drop(index = 63)
data_for_chart9.loc[60, 'D':] = data_for_chart9.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart9.loc[61, 'D':] = data_for_chart9.loc[61, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart9.loc[64, 'D':] = data_for_chart9.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart9.loc[65, 'D':] = data_for_chart9.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart9.iloc[:, 0] = None
columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
data_for_chart9.drop(columns=columns_to_drop, inplace=True)

updated_column_names_chart9 = Builder().dates(df1,58, 'M','P')

converted_updated_column_names_chart9 = Builder().convert_to_date_time(updated_column_names_chart9)

formated_updated_column_names_chart9 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart9]

data_for_chart9.columns = [data_for_chart9.columns[0]]+formated_updated_column_names_chart9

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA')

df3 = Builder().read_excel(file_name_1, sheet_name_1)
custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
df3.columns = custom_column_names_df3

weekly_data_for_chart9 = Builder().extract_data(df3, 'C', 'P', 60, 65)
weekly_data_for_chart9 = weekly_data_for_chart9.drop(index = 62)
weekly_data_for_chart9 = weekly_data_for_chart9.drop(index = 63)
weekly_data_for_chart9.loc[60, 'D':] = weekly_data_for_chart9.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart9.loc[61, 'D':] = weekly_data_for_chart9.loc[61, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart9.loc[64, 'D':] = weekly_data_for_chart9.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart9.loc[65, 'D':] = weekly_data_for_chart9.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
weekly_data_for_chart9.drop(columns=weekly_columns_to_drop, inplace=True)

weekly_updated_column_names_chart9 = Builder().dates(df3,58, 'M','P')

weekly_data_for_chart9.columns = [weekly_data_for_chart9.columns[0]]+weekly_updated_column_names_chart9

insert_position = 1

for i, col in enumerate(formated_updated_column_names_chart9):
    weekly_data_for_chart9.insert(insert_position+i,col,None)
#print(weekly_data_for_chart9)

for col in weekly_updated_column_names_chart9:
    data_for_chart9[col] = None

final_data_for_chart9 = pd.concat([weekly_data_for_chart9, data_for_chart9], axis=0)

final_data_for_chart9 = final_data_for_chart9.fillna("")


#For Chart10

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS')

df1 = Builder().read_excel(file_name_1, sheet_name_1)
df2 =Builder().read_excel(file_name_1,sheet_name_2)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

df1.columns = custom_column_names_df1
df2.columns = custom_column_names_df2

data_for_chart10 = Builder().extract_data(df1, 'C', 'P', 81, 86)
data_for_chart10 = data_for_chart10.drop(index = 82)
columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
data_for_chart10.drop(columns=columns_to_drop, inplace=True)

updated_column_names_chart10 = Builder().dates(df1,79, 'M','P')

converted_updated_column_names_chart10 = Builder().convert_to_date_time(updated_column_names_chart10)

formated_updated_column_names_chart10 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart10]

data_for_chart10.columns = [data_for_chart10.columns[0]]+formated_updated_column_names_chart10

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA')

df3 = Builder().read_excel(file_name_1, sheet_name_1)
custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
df3.columns = custom_column_names_df3

weekly_data_for_chart10 = Builder().extract_data(df3, 'C', 'P', 81, 86)
weekly_data_for_chart10 = weekly_data_for_chart10.drop(index = 82)

weekly_columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
weekly_data_for_chart10.drop(columns=weekly_columns_to_drop, inplace=True)

weekly_updated_column_names_chart10 = Builder().dates(df3,79, 'M','P')

weekly_data_for_chart10.columns = [weekly_data_for_chart10.columns[0]]+weekly_updated_column_names_chart10

final_data_for_chart10 = pd.concat([data_for_chart10, weekly_data_for_chart10], axis=1)


#For Chart11

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS')

df1 = Builder().read_excel(file_name_1, sheet_name_1)
df2 =Builder().read_excel(file_name_1,sheet_name_2)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

df1.columns = custom_column_names_df1
df2.columns = custom_column_names_df2

data_for_chart11 = Builder().extract_data(df1, 'C', 'P', 132, 137)
data_for_chart11 = data_for_chart11.drop(index = 134)
data_for_chart11 = data_for_chart11.drop(index = 135)
data_for_chart11.loc[132, 'D':] = data_for_chart11.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart11.loc[133, 'D':] = data_for_chart11.loc[133, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart11.loc[136, 'D':] = data_for_chart11.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart11.loc[137, 'D':] = data_for_chart11.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
data_for_chart11.iloc[:, 0] = None
columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
data_for_chart11.drop(columns=columns_to_drop, inplace=True)

updated_column_names_chart11 = Builder().dates(df1,130, 'M','P')

converted_updated_column_names_chart11 = Builder().convert_to_date_time(updated_column_names_chart11)

formated_updated_column_names_chart11 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart11]

data_for_chart11.columns = [data_for_chart11.columns[0]]+formated_updated_column_names_chart11

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA')

df3 = Builder().read_excel(file_name_1, sheet_name_1)
custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
df3.columns = custom_column_names_df3

weekly_data_for_chart11 = Builder().extract_data(df3, 'C', 'P', 132, 137)
weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 134)
weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
weekly_data_for_chart11.loc[132, 'D':] = weekly_data_for_chart11.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart11.loc[133, 'D':] = weekly_data_for_chart11.loc[133, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart11.loc[136, 'D':] = weekly_data_for_chart11.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_data_for_chart11.loc[137, 'D':] = weekly_data_for_chart11.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
weekly_columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
weekly_data_for_chart11.drop(columns=weekly_columns_to_drop, inplace=True)

weekly_updated_column_names_chart11 = Builder().dates(df3,130, 'M','P')

weekly_data_for_chart11.columns = [weekly_data_for_chart11.columns[0]]+weekly_updated_column_names_chart11

insert_position = 1

for i, col in enumerate(formated_updated_column_names_chart11):
    weekly_data_for_chart11.insert(insert_position+i,col,None)


for col in weekly_updated_column_names_chart11:
    data_for_chart11[col] = None

final_data_for_chart11 = pd.concat([weekly_data_for_chart11, data_for_chart11], axis=0)

final_data_for_chart11 = final_data_for_chart11.fillna("")


#For Chart12

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS')

df1 = Builder().read_excel(file_name_1, sheet_name_1)
df2 =Builder().read_excel(file_name_1,sheet_name_2)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

df1.columns = custom_column_names_df1
df2.columns = custom_column_names_df2

data_for_chart12 = Builder().extract_data(df1, 'C', 'P', 157, 162)
data_for_chart12 = data_for_chart12.drop(index = 158)
columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
data_for_chart12.drop(columns=columns_to_drop, inplace=True)

updated_column_names_chart12 = Builder().dates(df1,155, 'M','P')

converted_updated_column_names_chart12 = Builder().convert_to_date_time(updated_column_names_chart12)

formated_updated_column_names_chart12 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart12]

data_for_chart12.columns = [data_for_chart12.columns[0]]+formated_updated_column_names_chart12

Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA')

df3 = Builder().read_excel(file_name_1, sheet_name_1)
custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
df3.columns = custom_column_names_df3

weekly_data_for_chart12 = Builder().extract_data(df3, 'C', 'P', 157, 162)
weekly_data_for_chart12 = weekly_data_for_chart12.drop(index = 158)

weekly_columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
weekly_data_for_chart12.drop(columns=weekly_columns_to_drop, inplace=True)

weekly_updated_column_names_chart12 = Builder().dates(df3,155, 'M','P')

weekly_data_for_chart12.columns = [weekly_data_for_chart12.columns[0]]+weekly_updated_column_names_chart12

final_data_for_chart12 = pd.concat([data_for_chart12, weekly_data_for_chart12], axis=1)


#Updating chart1

chart_name = "Demand Pacing - Monthly and Weekly - 1"
dataframe = data_for_chart1
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name, dataframe, output_file_name)

#Updating Chart2

chart_name2 = "Demand Pacing - Monthly and Weekly - 2"
dataframe2 = data_for_chart2
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name2, dataframe2, output_file_name)

#Updating Chart3

chart_name3 = "Demand Pacing - Monthly and Weekly - 3"
dataframe3 = data_for_chart3
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name3, dataframe3, output_file_name)

#Updating Chart4

chart_name4 = "Demand Pacing - Monthly and Weekly - 4"
dataframe4 = data_for_chart4
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name4, dataframe4, output_file_name)

#Updating Chart5

chart_name5 = "Demand Pacing - Monthly and Weekly - 5"
dataframe5 = data_for_chart5
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name5, dataframe5, output_file_name)


#Updating Chart6
chart_name6 = "Demand Pacing - Monthly and Weekly - 6"
dataframe6 = data_for_chart6
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name6, dataframe6, output_file_name)

#Updating Chart7
chart_name7 = "Demand Pacing - Monthly and Weekly - 7"
dataframe7 = data_for_chart7
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name7, dataframe7, output_file_name)

#Updating Chart8
chart_name8 = "Demand Pacing - Monthly and Weekly - 8"
dataframe8 = data_for_chart8
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name8, dataframe8, output_file_name)

#Updating Chart9

chart_name9 = "Demand Pacing - Monthly and Weekly - 9"
dataframe9 = final_data_for_chart9
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name9, dataframe9, output_file_name)

#Updating Chart10

chart_name10 = "Demand Pacing - Monthly and Weekly - 10"
dataframe10 = final_data_for_chart10
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name10, dataframe10, output_file_name)


#Updating Chart11

chart_name11 = "Demand Pacing - Monthly and Weekly - 11"
dataframe11 = final_data_for_chart11
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name11, dataframe11, output_file_name)

#Updating Chart12

chart_name12 = "Demand Pacing - Monthly and Weekly - 12"
dataframe12 = final_data_for_chart12
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name12, dataframe12, output_file_name)

