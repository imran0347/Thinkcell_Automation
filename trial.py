import streamlit as st
from thinkcellbuilder import Presentation, Template
import pandas as pd
from datetime import datetime
from builder import Builder
import requests
from thinkcell import Thinkcell
from write_excel import Write_Excel
from Office365_API import SharePoint
import re
import sys,os
from pathlib import PurePath
from Office365_API import SharePoint
import re
import sys,os
from pathlib import PurePath
import win32com.client as win32
from excel_copy import Excel_Copy
import keyboard
import gdown

def main():
    st.title("THINKCELL AUTOMATION")

    # Define default values
    # FOLDER_NAME = "Comcast_Data"

    FOLDER_DEST = r"C:\Users\imran.s\Desktop\POC\Thinkcell_Automation\storage"

    # FILE_NAME = "None"

    # FILE_NAME_PATTERN = "None"

    # Create input fields with default values
    # folder_name = st.text_input("ENTER FOLDER NAME", FOLDER_NAME)
    # folder_dest = st.text_input("ENTER FOLDER DESTINATION", FOLDER_DEST)
    # file_name = st.text_input("ENTER FILE NAME (OPTIONAL)", FILE_NAME)
    # file_name_pattern = st.text_input("ENTER FILE NAME PATTERN (OPTIONAL)", FILE_NAME_PATTERN)

    file_id = "https://drive.google.com/uc?id=1bYPVibaEXoT-wOXuJAfcPhvUAT-v6X--"

    file_path = st.text_input("ENTER THE FILE PATH", file_id)

    # Button to execute script
    if st.button("START"):
        download_file_from_google_drive(file_id, FOLDER_DEST)
        update_charts()
        
    

# def download_files(folder_name, folder_dest, file_name, file_name_pattern):
#     def save_file(file_n, file_obj):
#         file_dir_path = PurePath(folder_dest,file_n)
#         with open(file_dir_path, 'wb') as f:
#             f.write(file_obj)

#     def get_file(file_n, folder):
#         file_obj = SharePoint().download_file(file_n,folder)
#         save_file(file_n,file_obj)

#     def get_files(folder):
#         files_list = SharePoint()._get_files_list(folder)
#         for file in files_list:
#             get_file(file.name, folder)

#     def get_files_by_pattern(keyword, folder):
#         files_list = SharePoint()._get_files_list(folder)
#         for file in files_list:
#             if re.search(keyword, file.name):
#                 get_file(file.name,folder)

#     if file_name != 'None':
#         get_file(file_name,folder_name)
#     elif file_name_pattern != 'None':
#         get_files_by_pattern(file_name_pattern, folder_name)
#     else:
#         get_files(folder_name)

def download_file_from_google_drive(file_id, destination_folder):
    # Construct the URL for the file
    url = f'{file_id}'
    
    # Construct the path to save the file
    destination_path = f'{destination_folder}/downloaded_file.xlsb'
    
    # Download the file
    gdown.download(url, destination_path, quiet=False)
    print(f"File downloaded and saved to {destination_path}")
Write_Excel().close_all_excel_instances()
def update_charts():


    Excel_Copy().copy()
    
   
    #Updating Charts
    file_path = r"C:\Users\imran.s\Desktop\POC\Thinkcell_Automation\storage\downloaded_file.xlsb"
    file_name_1 = r"storage\downloaded_file.xlsb" 
    sheet_name_1 = 'By Marketing Channel (TEMPLATE)'
    sheet_name_2 = 'National Monthly'
    Write_Excel().close_all_excel_instances()


    #chart2
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','National')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart2 = Builder().extract_data(df1, 'C', 'P', 32, 38)
    data_for_chart2 = Builder().add_row(df1, data_for_chart2, 59, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['Conversion Rate by Sales Tactic'] + [''] * (len(data_for_chart2.columns) - 1), ],
                        columns=data_for_chart2.columns)
    data_for_chart2 = pd.concat([data_for_chart2, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart2.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart2.columns = [data_for_chart2.columns[0]]+formated_updated_column_names

    data_for_chart2.set_index('C',inplace=True)
    print(data_for_chart2)
    data_for_chart2.to_csv("data_for_chart2.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','National')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart2 = Builder().extract_data(df3, 'C', 'P', 32, 38)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart2 = pd.concat([weekly_data_for_chart2, new_rows], ignore_index=True)

    weekly_data_for_chart2 = Builder().add_row(df1, weekly_data_for_chart2, 59, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart2.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart2 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart2.columns = [weekly_data_for_chart2.columns[0]]+weekly_updated_column_names_chart2

    weekly_data_for_chart2.set_index('C',inplace=True)
    weekly_data_for_chart2.to_csv("weekly_data_for_chart_2.csv")

    final_data_for_chart2 = pd.concat([data_for_chart2, weekly_data_for_chart2], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    final_data_for_chart2.reset_index(inplace=True)
    print(final_data_for_chart2)
    final_data_for_chart2.to_csv("final_data_for_chart2.csv")



    #Chart1

    data_for_chart1 = Builder().extract_data(df1, 'C', 'P', 20, 26)
    data_for_chart1 = Builder().add_row(df1, data_for_chart1, 52, 'C', 'P', 'D')

    data_for_chart1 = Builder().add_row(df1,data_for_chart1,28,'C','P','D')

    new_rows = pd.DataFrame([['Eng. Traffic Rate Excl Display'] + [''] * (len(data_for_chart1.columns) - 1), 
                         ['Engaged Visit Rate'] + [''] * (len(data_for_chart1.columns) - 1)],
                        columns=data_for_chart1.columns)
    data_for_chart1 = pd.concat([data_for_chart1, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart1.drop(columns=columns_to_drop, inplace=True)

    data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart1.columns = [data_for_chart1.columns[0]]+formated_updated_column_names

    data_for_chart1.set_index('C',inplace=True)
    print(data_for_chart1)
    data_for_chart1.to_csv("data_for_chart1.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','National')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart1 = Builder().extract_data(df3, 'C', 'P', 20, 26)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart1 = pd.concat([weekly_data_for_chart1, new_rows], ignore_index=True)

    weekly_data_for_chart1 = Builder().add_row(df1, weekly_data_for_chart1, 52, 'C', 'P', 'D')

    weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart1.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart1 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart1.columns = [weekly_data_for_chart1.columns[0]]+weekly_updated_column_names_chart1

    weekly_data_for_chart1.set_index('C',inplace=True)
    weekly_data_for_chart1.to_csv("weekly_data_for_chart_1.csv")

    final_data_for_chart1 = pd.concat([data_for_chart1, weekly_data_for_chart1], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')

    final_data_for_chart1.reset_index(inplace=True)
    print(final_data_for_chart1)
    final_data_for_chart1.to_csv("final_data_for_chart1.csv")

    #For Chart3

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','National')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart3 = Builder().extract_data(df1, 'C', 'P', 60, 65)
    data_for_chart3 = data_for_chart3.drop(index = 61)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart3.loc[60, 'D':] = data_for_chart3.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart3.loc[62, 'D':] = data_for_chart3.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart3.loc[63, 'D':] = data_for_chart3.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart3.loc[64, 'D':] = data_for_chart3.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart3.loc[65, 'D':] = data_for_chart3.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart3.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart3.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart3 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart3 = Builder().convert_to_date_time(updated_column_names_chart3)

    formated_updated_column_names_chart3 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart3]

    data_for_chart3.columns = [data_for_chart3.columns[0]]+formated_updated_column_names_chart3
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','National')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart3 = Builder().extract_data(df3, 'C', 'P', 60, 65)
    weekly_data_for_chart3 = weekly_data_for_chart3.drop(index = 61)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart3.loc[60, 'D':] = weekly_data_for_chart3.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart3.loc[62, 'D':] = weekly_data_for_chart3.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart3.loc[63, 'D':] = weekly_data_for_chart3.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart3.loc[64, 'D':] = weekly_data_for_chart3.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart3.loc[65, 'D':] = weekly_data_for_chart3.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart3.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart3 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart3.columns = [weekly_data_for_chart3.columns[0]]+weekly_updated_column_names_chart3

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart3):
        weekly_data_for_chart3.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart3:
        data_for_chart3[col] = None

    final_data_for_chart3 = pd.concat([weekly_data_for_chart3, data_for_chart3], axis=0)

    final_data_for_chart3 = final_data_for_chart3.fillna("")
    final_data_for_chart3.to_csv("final_data_for_chart3.csv")





    #For Chart4

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','National')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart4 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart4 = data_for_chart4.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart4.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart4 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart4 = Builder().convert_to_date_time(updated_column_names_chart4)

    formated_updated_column_names_chart4 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart4]

    data_for_chart4.columns = [data_for_chart4.columns[0]]+formated_updated_column_names_chart4

    data_for_chart4.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','National')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart4 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart4= weekly_data_for_chart4.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart4.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart4 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart4.columns = [weekly_data_for_chart4.columns[0]]+weekly_updated_column_names_chart4

    weekly_data_for_chart4.set_index('C',inplace=True)

    final_data_for_chart4 = pd.concat([data_for_chart4, weekly_data_for_chart4], axis=1)
    final_data_for_chart4.reset_index(inplace=True)

    final_data_for_chart4.to_csv("final_data_for_chart4.csv")




    #For Chart5

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','National')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart5 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart5 = data_for_chart5.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart5.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart5 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart5 = Builder().convert_to_date_time(updated_column_names_chart5)

    formated_updated_column_names_chart5 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart5]

    data_for_chart5.columns = [data_for_chart5.columns[0]]+formated_updated_column_names_chart5

    data_for_chart5.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','National')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart5 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart5= weekly_data_for_chart5.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart5.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart5 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart5.columns = [weekly_data_for_chart5.columns[0]]+weekly_updated_column_names_chart5

    weekly_data_for_chart5.set_index('C',inplace=True)

    final_data_for_chart5 = pd.concat([data_for_chart5, weekly_data_for_chart5], axis=1)

    final_data_for_chart5.reset_index(inplace=True)

    final_data_for_chart5.to_csv("final_data_for_chart5.csv")

    #For Chart6

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','National')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart6 = Builder().extract_data(df1, 'C', 'P', 132, 137)
    data_for_chart6 = data_for_chart6.drop(index = 133)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart6.loc[132, 'D':] = data_for_chart6.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart6.loc[134, 'D':] = data_for_chart6.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart6.loc[135, 'D':] = data_for_chart6.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart6.loc[136, 'D':] = data_for_chart6.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart6.loc[137, 'D':] = data_for_chart6.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart6.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart6.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart6 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart6 = Builder().convert_to_date_time(updated_column_names_chart6)

    formated_updated_column_names_chart6 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart6]

    data_for_chart6.columns = [data_for_chart6.columns[0]]+formated_updated_column_names_chart6
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','National')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart6 = Builder().extract_data(df3, 'C', 'P', 132, 137)
    weekly_data_for_chart6 = weekly_data_for_chart6.drop(index = 133)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart6.loc[132, 'D':] = weekly_data_for_chart6.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart6.loc[134, 'D':] = weekly_data_for_chart6.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart6.loc[135, 'D':] = weekly_data_for_chart6.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart6.loc[136, 'D':] = weekly_data_for_chart6.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart6.loc[137, 'D':] = weekly_data_for_chart6.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart6.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart6 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart6.columns = [weekly_data_for_chart6.columns[0]]+weekly_updated_column_names_chart6

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart6):
        weekly_data_for_chart6.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart6:
        data_for_chart6[col] = None

    final_data_for_chart6 = pd.concat([weekly_data_for_chart6, data_for_chart6], axis=0)

    final_data_for_chart6 = final_data_for_chart6.fillna("")
    final_data_for_chart6.to_csv("final_data_for_chart6.csv")

    #chart7
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','National')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart7 = Builder().extract_data(df1, 'C', 'P', 157, 162)
    data_for_chart7 = data_for_chart7.drop(index = 158)
    data_for_chart7 = Builder().add_row(df1, data_for_chart7, 131, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['CR by Sales Tactic'] + [''] * (len(data_for_chart7.columns) - 1), ],
                        columns=data_for_chart7.columns)
    data_for_chart7 = pd.concat([data_for_chart7, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart7.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart7.columns = [data_for_chart7.columns[0]]+formated_updated_column_names

    data_for_chart7.set_index('C',inplace=True)
    print(data_for_chart7)
    data_for_chart7.to_csv("data_for_chart7.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','National')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart7 = Builder().extract_data(df3, 'C', 'P', 157, 162)
    weekly_data_for_chart7 = weekly_data_for_chart7.drop(index = 158)
    
    weekly_data_for_chart7 = pd.concat([weekly_data_for_chart7, new_rows], ignore_index=True)

    weekly_data_for_chart7 = Builder().add_row(df1, weekly_data_for_chart7, 131, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart7.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart7 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart7.columns = [weekly_data_for_chart7.columns[0]]+weekly_updated_column_names_chart7

    weekly_data_for_chart7.set_index('C',inplace=True)
    weekly_data_for_chart7.to_csv("weekly_data_for_chart_7.csv")

    final_data_for_chart7 = pd.concat([data_for_chart7, weekly_data_for_chart7], axis=1)

    final_data_for_chart7.reset_index(inplace=True)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    # print(final_data_for_chart7)
    final_data_for_chart7.to_csv("final_data_for_chart7.csv")

##########################################################################################################
    #chart9
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','Northeast')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart9 = Builder().extract_data(df1, 'C', 'P', 32, 38)
    data_for_chart9 = Builder().add_row(df1, data_for_chart9, 59, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['Conversion Rate by Sales Tactic'] + [''] * (len(data_for_chart9.columns) - 1), ],
                        columns=data_for_chart9.columns)
    data_for_chart9 = pd.concat([data_for_chart9, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart9.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart9.columns = [data_for_chart9.columns[0]]+formated_updated_column_names

    data_for_chart9.set_index('C',inplace=True)
    print(data_for_chart9)
    data_for_chart9.to_csv("data_for_chart9.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Northeast')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart9 = Builder().extract_data(df3, 'C', 'P', 32, 38)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart9 = pd.concat([weekly_data_for_chart9, new_rows], ignore_index=True)

    weekly_data_for_chart9 = Builder().add_row(df1, weekly_data_for_chart9, 59, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart9.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart9 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart9.columns = [weekly_data_for_chart9.columns[0]]+weekly_updated_column_names_chart9

    weekly_data_for_chart9.set_index('C',inplace=True)
    weekly_data_for_chart9.to_csv("weekly_data_for_chart_9.csv")

    final_data_for_chart9 = pd.concat([data_for_chart9, weekly_data_for_chart9], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    final_data_for_chart9.reset_index(inplace=True)
    print(final_data_for_chart9)
    final_data_for_chart9.to_csv("final_data_for_chart9.csv")



    #Chart1

    data_for_chart8 = Builder().extract_data(df1, 'C', 'P', 20, 26)
    data_for_chart8 = Builder().add_row(df1, data_for_chart8, 52, 'C', 'P', 'D')

    data_for_chart8 = Builder().add_row(df1,data_for_chart8,28,'C','P','D')

    new_rows = pd.DataFrame([['Eng. Traffic Rate Excl Display'] + [''] * (len(data_for_chart8.columns) - 1), 
                         ['Engaged Visit Rate'] + [''] * (len(data_for_chart8.columns) - 1)],
                        columns=data_for_chart8.columns)
    data_for_chart8 = pd.concat([data_for_chart8, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart8.drop(columns=columns_to_drop, inplace=True)

    data_for_chart8.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart8.columns = [data_for_chart8.columns[0]]+formated_updated_column_names

    data_for_chart8.set_index('C',inplace=True)
    print(data_for_chart8)
    data_for_chart8.to_csv("data_for_chart8.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Northeast')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart8 = Builder().extract_data(df3, 'C', 'P', 20, 26)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart8 = pd.concat([weekly_data_for_chart8, new_rows], ignore_index=True)

    weekly_data_for_chart8 = Builder().add_row(df1, weekly_data_for_chart8, 52, 'C', 'P', 'D')

    weekly_data_for_chart8 = Builder().add_row(df1,weekly_data_for_chart8,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart8.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart8 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart8.columns = [weekly_data_for_chart8.columns[0]]+weekly_updated_column_names_chart8

    weekly_data_for_chart8.set_index('C',inplace=True)
    weekly_data_for_chart8.to_csv("weekly_data_for_chart_8.csv")

    final_data_for_chart8 = pd.concat([data_for_chart8, weekly_data_for_chart8], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')

    final_data_for_chart8.reset_index(inplace=True)
    print(final_data_for_chart8)
    final_data_for_chart8.to_csv("final_data_for_chart8.csv")

    #For Chart3

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Northeast')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart10 = Builder().extract_data(df1, 'C', 'P', 60, 65)
    data_for_chart10 = data_for_chart10.drop(index = 61)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart10.loc[60, 'D':] = data_for_chart10.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart10.loc[62, 'D':] = data_for_chart10.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart10.loc[63, 'D':] = data_for_chart10.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart10.loc[64, 'D':] = data_for_chart10.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart10.loc[65, 'D':] = data_for_chart10.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart10.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart10.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart10 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart10 = Builder().convert_to_date_time(updated_column_names_chart10)

    formated_updated_column_names_chart10 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart10]

    data_for_chart10.columns = [data_for_chart10.columns[0]]+formated_updated_column_names_chart10
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Northeast')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart10 = Builder().extract_data(df3, 'C', 'P', 60, 65)
    weekly_data_for_chart10 = weekly_data_for_chart10.drop(index = 61)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart10.loc[60, 'D':] = weekly_data_for_chart10.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart10.loc[62, 'D':] = weekly_data_for_chart10.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart10.loc[63, 'D':] = weekly_data_for_chart10.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart10.loc[64, 'D':] = weekly_data_for_chart10.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart10.loc[65, 'D':] = weekly_data_for_chart10.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart10.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart10 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart10.columns = [weekly_data_for_chart10.columns[0]]+weekly_updated_column_names_chart10

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart10):
        weekly_data_for_chart10.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart10:
        data_for_chart10[col] = None

    final_data_for_chart10 = pd.concat([weekly_data_for_chart10, data_for_chart10], axis=0)

    final_data_for_chart10 = final_data_for_chart10.fillna("")
    final_data_for_chart10.to_csv("final_data_for_chart10.csv")





    #For Chart11

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Northeast')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart11 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart11 = data_for_chart11.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart11.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart11 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart11 = Builder().convert_to_date_time(updated_column_names_chart11)

    formated_updated_column_names_chart11 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart11]

    data_for_chart11.columns = [data_for_chart11.columns[0]]+formated_updated_column_names_chart11

    data_for_chart11.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Northeast')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart11 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart11= weekly_data_for_chart11.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart11.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart11 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart11.columns = [weekly_data_for_chart11.columns[0]]+weekly_updated_column_names_chart11

    weekly_data_for_chart11.set_index('C',inplace=True)

    final_data_for_chart11 = pd.concat([data_for_chart11, weekly_data_for_chart11], axis=1)
    final_data_for_chart11.reset_index(inplace=True)

    final_data_for_chart11.to_csv("final_data_for_chart11.csv")




    #For Chart12

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Northeast')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart12 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart12 = data_for_chart12.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart12.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart12 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart12 = Builder().convert_to_date_time(updated_column_names_chart12)

    formated_updated_column_names_chart12 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart12]

    data_for_chart12.columns = [data_for_chart12.columns[0]]+formated_updated_column_names_chart12

    data_for_chart12.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Northeast')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart12 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart12= weekly_data_for_chart12.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart12.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart12 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart12.columns = [weekly_data_for_chart12.columns[0]]+weekly_updated_column_names_chart12

    weekly_data_for_chart12.set_index('C',inplace=True)

    final_data_for_chart12 = pd.concat([data_for_chart12, weekly_data_for_chart12], axis=1)

    final_data_for_chart12.reset_index(inplace=True)

    final_data_for_chart12.to_csv("final_data_for_chart12.csv")

    #For Chart13

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Northeast')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart13 = Builder().extract_data(df1, 'C', 'P', 132, 137)
    data_for_chart13 = data_for_chart13.drop(index = 133)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart13.loc[132, 'D':] = data_for_chart13.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart13.loc[134, 'D':] = data_for_chart13.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart13.loc[135, 'D':] = data_for_chart13.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart13.loc[136, 'D':] = data_for_chart13.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart13.loc[137, 'D':] = data_for_chart13.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart13.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart13.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart13 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart13 = Builder().convert_to_date_time(updated_column_names_chart13)

    formated_updated_column_names_chart13 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart13]

    data_for_chart13.columns = [data_for_chart13.columns[0]]+formated_updated_column_names_chart13
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Northeast')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart13 = Builder().extract_data(df3, 'C', 'P', 132, 137)
    weekly_data_for_chart13 = weekly_data_for_chart13.drop(index = 133)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart13.loc[132, 'D':] = weekly_data_for_chart13.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart13.loc[134, 'D':] = weekly_data_for_chart13.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart13.loc[135, 'D':] = weekly_data_for_chart13.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart13.loc[136, 'D':] = weekly_data_for_chart13.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart13.loc[137, 'D':] = weekly_data_for_chart13.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart13.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart13 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart13.columns = [weekly_data_for_chart13.columns[0]]+weekly_updated_column_names_chart13

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart13):
        weekly_data_for_chart13.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart13:
        data_for_chart13[col] = None

    final_data_for_chart13 = pd.concat([weekly_data_for_chart13, data_for_chart13], axis=0)

    final_data_for_chart13 = final_data_for_chart13.fillna("")
    final_data_for_chart13.to_csv("final_data_for_chart13.csv")

    #chart14
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','Northeast')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart14 = Builder().extract_data(df1, 'C', 'P', 157, 162)
    data_for_chart14 = data_for_chart14.drop(index = 158)
    data_for_chart14 = Builder().add_row(df1, data_for_chart14, 131, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['CR by Sales Tactic'] + [''] * (len(data_for_chart14.columns) - 1), ],
                        columns=data_for_chart14.columns)
    data_for_chart14 = pd.concat([data_for_chart14, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart14.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart14.columns = [data_for_chart14.columns[0]]+formated_updated_column_names

    data_for_chart14.set_index('C',inplace=True)
    print(data_for_chart14)
    data_for_chart14.to_csv("data_for_chart14.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Northeast')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart14 = Builder().extract_data(df3, 'C', 'P', 157, 162)
    weekly_data_for_chart14 = weekly_data_for_chart14.drop(index = 158)
    
    weekly_data_for_chart14 = pd.concat([weekly_data_for_chart14, new_rows], ignore_index=True)

    weekly_data_for_chart14 = Builder().add_row(df1, weekly_data_for_chart14, 131, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart14.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart14 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart14.columns = [weekly_data_for_chart14.columns[0]]+weekly_updated_column_names_chart14

    weekly_data_for_chart14.set_index('C',inplace=True)
    weekly_data_for_chart14.to_csv("weekly_data_for_chart_14.csv")

    final_data_for_chart14 = pd.concat([data_for_chart14, weekly_data_for_chart14], axis=1)

    final_data_for_chart14.reset_index(inplace=True)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    # print(final_data_for_chart7)
    final_data_for_chart14.to_csv("final_data_for_chart14.csv")



##########################################################################################################
    #chart16
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','Central')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart16 = Builder().extract_data(df1, 'C', 'P', 32, 38)
    data_for_chart16 = Builder().add_row(df1, data_for_chart16, 59, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['Conversion Rate by Sales Tactic'] + [''] * (len(data_for_chart16.columns) - 1), ],
                        columns=data_for_chart16.columns)
    data_for_chart16 = pd.concat([data_for_chart16, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart16.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart16.columns = [data_for_chart16.columns[0]]+formated_updated_column_names

    data_for_chart16.set_index('C',inplace=True)
    print(data_for_chart16)
    data_for_chart16.to_csv("data_for_chart16.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Central')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart16 = Builder().extract_data(df3, 'C', 'P', 32, 38)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart16 = pd.concat([weekly_data_for_chart16, new_rows], ignore_index=True)

    weekly_data_for_chart16 = Builder().add_row(df1, weekly_data_for_chart16, 59, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart16.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart16 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart16.columns = [weekly_data_for_chart16.columns[0]]+weekly_updated_column_names_chart16

    weekly_data_for_chart16.set_index('C',inplace=True)
    weekly_data_for_chart16.to_csv("weekly_data_for_chart_16.csv")

    final_data_for_chart16 = pd.concat([data_for_chart16, weekly_data_for_chart16], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    final_data_for_chart16.reset_index(inplace=True)
    print(final_data_for_chart16)
    final_data_for_chart16.to_csv("final_data_for_chart16.csv")



    #Chart15

    data_for_chart15 = Builder().extract_data(df1, 'C', 'P', 20, 26)
    data_for_chart15 = Builder().add_row(df1, data_for_chart15, 52, 'C', 'P', 'D')

    data_for_chart15 = Builder().add_row(df1,data_for_chart15,28,'C','P','D')

    new_rows = pd.DataFrame([['Eng. Traffic Rate Excl Display'] + [''] * (len(data_for_chart15.columns) - 1), 
                         ['Engaged Visit Rate'] + [''] * (len(data_for_chart15.columns) - 1)],
                        columns=data_for_chart15.columns)
    data_for_chart15 = pd.concat([data_for_chart15, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart15.drop(columns=columns_to_drop, inplace=True)

    data_for_chart15.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart15.columns = [data_for_chart15.columns[0]]+formated_updated_column_names

    data_for_chart15.set_index('C',inplace=True)
    print(data_for_chart15)
    data_for_chart15.to_csv("data_for_chart15.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Central')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart15 = Builder().extract_data(df3, 'C', 'P', 20, 26)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart15 = pd.concat([weekly_data_for_chart15, new_rows], ignore_index=True)

    weekly_data_for_chart15 = Builder().add_row(df1, weekly_data_for_chart15, 52, 'C', 'P', 'D')

    weekly_data_for_chart15 = Builder().add_row(df1,weekly_data_for_chart15,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart15.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart15 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart15.columns = [weekly_data_for_chart15.columns[0]]+weekly_updated_column_names_chart15

    weekly_data_for_chart15.set_index('C',inplace=True)
    weekly_data_for_chart15.to_csv("weekly_data_for_chart_15.csv")

    final_data_for_chart15 = pd.concat([data_for_chart15, weekly_data_for_chart15], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')

    final_data_for_chart15.reset_index(inplace=True)
    print(final_data_for_chart15)
    final_data_for_chart15.to_csv("final_data_for_chart15.csv")

    #For Chart17

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Central')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart17 = Builder().extract_data(df1, 'C', 'P', 60, 65)
    data_for_chart17 = data_for_chart17.drop(index = 61)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart17.loc[60, 'D':] = data_for_chart17.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart17.loc[62, 'D':] = data_for_chart17.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart17.loc[63, 'D':] = data_for_chart17.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart17.loc[64, 'D':] = data_for_chart17.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart17.loc[65, 'D':] = data_for_chart17.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart17.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart17.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart17 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart17 = Builder().convert_to_date_time(updated_column_names_chart17)

    formated_updated_column_names_chart17 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart17]

    data_for_chart17.columns = [data_for_chart17.columns[0]]+formated_updated_column_names_chart17
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Central')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart17 = Builder().extract_data(df3, 'C', 'P', 60, 65)
    weekly_data_for_chart17 = weekly_data_for_chart17.drop(index = 61)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart17.loc[60, 'D':] = weekly_data_for_chart17.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart17.loc[62, 'D':] = weekly_data_for_chart17.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart17.loc[63, 'D':] = weekly_data_for_chart17.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart17.loc[64, 'D':] = weekly_data_for_chart17.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart17.loc[65, 'D':] = weekly_data_for_chart17.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart17.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart17 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart17.columns = [weekly_data_for_chart17.columns[0]]+weekly_updated_column_names_chart17

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart17):
        weekly_data_for_chart17.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart17:
        data_for_chart17[col] = None

    final_data_for_chart17 = pd.concat([weekly_data_for_chart17, data_for_chart17], axis=0)

    final_data_for_chart17 = final_data_for_chart17.fillna("")
    final_data_for_chart17.to_csv("final_data_for_chart17.csv")





    #For Chart18

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Central')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart18 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart18 = data_for_chart18.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart18.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart18 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart18 = Builder().convert_to_date_time(updated_column_names_chart18)

    formated_updated_column_names_chart18 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart18]

    data_for_chart18.columns = [data_for_chart18.columns[0]]+formated_updated_column_names_chart18

    data_for_chart18.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Central')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart18 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart18= weekly_data_for_chart18.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart18.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart18 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart18.columns = [weekly_data_for_chart18.columns[0]]+weekly_updated_column_names_chart18

    weekly_data_for_chart18.set_index('C',inplace=True)

    final_data_for_chart18 = pd.concat([data_for_chart18, weekly_data_for_chart18], axis=1)
    final_data_for_chart18.reset_index(inplace=True)

    final_data_for_chart18.to_csv("final_data_for_chart18.csv")




    #For Chart19

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Central')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart19 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart19 = data_for_chart19.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart19.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart19 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart19 = Builder().convert_to_date_time(updated_column_names_chart19)

    formated_updated_column_names_chart19 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart19]

    data_for_chart19.columns = [data_for_chart19.columns[0]]+formated_updated_column_names_chart19

    data_for_chart19.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Central')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart19 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart19= weekly_data_for_chart19.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart19.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart19 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart19.columns = [weekly_data_for_chart19.columns[0]]+weekly_updated_column_names_chart19

    weekly_data_for_chart19.set_index('C',inplace=True)

    final_data_for_chart19 = pd.concat([data_for_chart19, weekly_data_for_chart19], axis=1)

    final_data_for_chart19.reset_index(inplace=True)

    final_data_for_chart19.to_csv("final_data_for_chart19.csv")

    #For Chart20

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','Central')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart20 = Builder().extract_data(df1, 'C', 'P', 132, 137)
    data_for_chart20 = data_for_chart20.drop(index = 133)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart20.loc[132, 'D':] = data_for_chart20.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart20.loc[134, 'D':] = data_for_chart20.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart20.loc[135, 'D':] = data_for_chart20.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart20.loc[136, 'D':] = data_for_chart20.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart20.loc[137, 'D':] = data_for_chart20.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart20.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart20.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart20 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart20 = Builder().convert_to_date_time(updated_column_names_chart20)

    formated_updated_column_names_chart20 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart20]

    data_for_chart20.columns = [data_for_chart20.columns[0]]+formated_updated_column_names_chart20
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','Central')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart20 = Builder().extract_data(df3, 'C', 'P', 132, 137)
    weekly_data_for_chart20 = weekly_data_for_chart20.drop(index = 133)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart20.loc[132, 'D':] = weekly_data_for_chart20.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart20.loc[134, 'D':] = weekly_data_for_chart20.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart20.loc[135, 'D':] = weekly_data_for_chart20.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart20.loc[136, 'D':] = weekly_data_for_chart20.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart20.loc[137, 'D':] = weekly_data_for_chart20.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart20.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart20 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart20.columns = [weekly_data_for_chart20.columns[0]]+weekly_updated_column_names_chart20

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart20):
        weekly_data_for_chart20.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart20:
        data_for_chart20[col] = None

    final_data_for_chart20 = pd.concat([weekly_data_for_chart20, data_for_chart20], axis=0)

    final_data_for_chart20 = final_data_for_chart20.fillna("")
    final_data_for_chart20.to_csv("final_data_for_chart20.csv")

    #chart7
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','Central')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart21 = Builder().extract_data(df1, 'C', 'P', 157, 162)
    data_for_chart21 = data_for_chart21.drop(index = 158)
    data_for_chart21 = Builder().add_row(df1, data_for_chart21, 131, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['CR by Sales Tactic'] + [''] * (len(data_for_chart21.columns) - 1), ],
                        columns=data_for_chart21.columns)
    data_for_chart21 = pd.concat([data_for_chart21, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart21.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart21.columns = [data_for_chart21.columns[0]]+formated_updated_column_names

    data_for_chart21.set_index('C',inplace=True)
    print(data_for_chart21)
    data_for_chart21.to_csv("data_for_chart21.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','Central')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart21 = Builder().extract_data(df3, 'C', 'P', 157, 162)
    weekly_data_for_chart21 = weekly_data_for_chart21.drop(index = 158)
    
    weekly_data_for_chart21 = pd.concat([weekly_data_for_chart21, new_rows], ignore_index=True)

    weekly_data_for_chart21 = Builder().add_row(df1, weekly_data_for_chart21, 131, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart21.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart21 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart21.columns = [weekly_data_for_chart21.columns[0]]+weekly_updated_column_names_chart7

    weekly_data_for_chart21.set_index('C',inplace=True)
    weekly_data_for_chart21.to_csv("weekly_data_for_chart_21.csv")

    final_data_for_chart21 = pd.concat([data_for_chart21, weekly_data_for_chart21], axis=1)

    final_data_for_chart21.reset_index(inplace=True)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    # print(final_data_for_chart7)
    final_data_for_chart21.to_csv("final_data_for_chart21.csv")


##########################################################################################################
    #chart23
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','West')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart23 = Builder().extract_data(df1, 'C', 'P', 32, 38)
    data_for_chart23 = Builder().add_row(df1, data_for_chart23, 59, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['Conversion Rate by Sales Tactic'] + [''] * (len(data_for_chart23.columns) - 1), ],
                        columns=data_for_chart23.columns)
    data_for_chart23 = pd.concat([data_for_chart23, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart23.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart23.columns = [data_for_chart23.columns[0]]+formated_updated_column_names

    data_for_chart23.set_index('C',inplace=True)
    print(data_for_chart23)
    data_for_chart23.to_csv("data_for_chart23.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','West')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart23 = Builder().extract_data(df3, 'C', 'P', 32, 38)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart23 = pd.concat([weekly_data_for_chart23, new_rows], ignore_index=True)

    weekly_data_for_chart23 = Builder().add_row(df1, weekly_data_for_chart23, 59, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart23.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart23 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart23.columns = [weekly_data_for_chart23.columns[0]]+weekly_updated_column_names_chart23

    weekly_data_for_chart23.set_index('C',inplace=True)
    weekly_data_for_chart23.to_csv("weekly_data_for_chart_23.csv")

    final_data_for_chart23 = pd.concat([data_for_chart23, weekly_data_for_chart23], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    final_data_for_chart23.reset_index(inplace=True)
    print(final_data_for_chart23)
    final_data_for_chart23.to_csv("final_data_for_chart23.csv")



    #Chart1

    data_for_chart22 = Builder().extract_data(df1, 'C', 'P', 20, 26)
    data_for_chart22 = Builder().add_row(df1, data_for_chart22, 52, 'C', 'P', 'D')

    data_for_chart22 = Builder().add_row(df1,data_for_chart22,28,'C','P','D')

    new_rows = pd.DataFrame([['Eng. Traffic Rate Excl Display'] + [''] * (len(data_for_chart22.columns) - 1), 
                         ['Engaged Visit Rate'] + [''] * (len(data_for_chart22.columns) - 1)],
                        columns=data_for_chart22.columns)
    data_for_chart22 = pd.concat([data_for_chart22, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart22.drop(columns=columns_to_drop, inplace=True)

    data_for_chart22.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart22.columns = [data_for_chart22.columns[0]]+formated_updated_column_names

    data_for_chart22.set_index('C',inplace=True)
    print(data_for_chart22)
    data_for_chart22.to_csv("data_for_chart22.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','West')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart22 = Builder().extract_data(df3, 'C', 'P', 20, 26)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart22 = pd.concat([weekly_data_for_chart22, new_rows], ignore_index=True)

    weekly_data_for_chart22 = Builder().add_row(df1, weekly_data_for_chart22, 52, 'C', 'P', 'D')

    weekly_data_for_chart22 = Builder().add_row(df1,weekly_data_for_chart22,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart22.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart22 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart22.columns = [weekly_data_for_chart22.columns[0]]+weekly_updated_column_names_chart22

    weekly_data_for_chart22.set_index('C',inplace=True)
    weekly_data_for_chart22.to_csv("weekly_data_for_chart_22.csv")

    final_data_for_chart22 = pd.concat([data_for_chart22, weekly_data_for_chart22], axis=1)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')

    final_data_for_chart22.reset_index(inplace=True)
    print(final_data_for_chart22)
    final_data_for_chart22.to_csv("final_data_for_chart22.csv")

    #For Chart24

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','West')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart24 = Builder().extract_data(df1, 'C', 'P', 60, 65)
    data_for_chart24 = data_for_chart24.drop(index = 61)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart24.loc[60, 'D':] = data_for_chart24.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart24.loc[62, 'D':] = data_for_chart24.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart24.loc[63, 'D':] = data_for_chart24.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart24.loc[64, 'D':] = data_for_chart24.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart24.loc[65, 'D':] = data_for_chart24.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart24.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart24.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart24 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart24 = Builder().convert_to_date_time(updated_column_names_chart24)

    formated_updated_column_names_chart24 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart24]

    data_for_chart24.columns = [data_for_chart24.columns[0]]+formated_updated_column_names_chart24
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','West')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart24 = Builder().extract_data(df3, 'C', 'P', 60, 65)
    weekly_data_for_chart24 = weekly_data_for_chart24.drop(index = 61)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart24.loc[60, 'D':] = weekly_data_for_chart24.loc[60, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart24.loc[62, 'D':] = weekly_data_for_chart24.loc[62, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart24.loc[63, 'D':] = weekly_data_for_chart24.loc[63, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart24.loc[64, 'D':] = weekly_data_for_chart24.loc[64, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart24.loc[65, 'D':] = weekly_data_for_chart24.loc[65, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart24.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart24 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart24.columns = [weekly_data_for_chart24.columns[0]]+weekly_updated_column_names_chart24

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart24):
        weekly_data_for_chart24.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart24:
        data_for_chart24[col] = None

    final_data_for_chart24 = pd.concat([weekly_data_for_chart24, data_for_chart24], axis=0)

    final_data_for_chart24 = final_data_for_chart24.fillna("")
    final_data_for_chart24.to_csv("final_data_for_chart24.csv")





    #For Chart25

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','West')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart25 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart25 = data_for_chart25.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart25.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart25 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart25 = Builder().convert_to_date_time(updated_column_names_chart25)

    formated_updated_column_names_chart25 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart25]

    data_for_chart25.columns = [data_for_chart25.columns[0]]+formated_updated_column_names_chart25

    data_for_chart25.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','West')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart25 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart25= weekly_data_for_chart25.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart25.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart25 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart25.columns = [weekly_data_for_chart25.columns[0]]+weekly_updated_column_names_chart25

    weekly_data_for_chart25.set_index('C',inplace=True)

    final_data_for_chart25 = pd.concat([data_for_chart25, weekly_data_for_chart25], axis=1)
    final_data_for_chart25.reset_index(inplace=True)

    final_data_for_chart25.to_csv("final_data_for_chart25.csv")




    #For Chart26

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','West')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart26 = Builder().extract_data(df1, 'C', 'P', 81, 86)
    data_for_chart26 = data_for_chart26.drop(index = 82)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart26.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart26 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart26 = Builder().convert_to_date_time(updated_column_names_chart26)

    formated_updated_column_names_chart26 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart26]

    data_for_chart26.columns = [data_for_chart26.columns[0]]+formated_updated_column_names_chart26

    data_for_chart26.set_index('C',inplace=True)

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','West')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart26 = Builder().extract_data(df3, 'C', 'P', 81, 86)
    weekly_data_for_chart26= weekly_data_for_chart26.drop(index = 82)

    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart26.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart26 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart26.columns = [weekly_data_for_chart26.columns[0]]+weekly_updated_column_names_chart26

    weekly_data_for_chart26.set_index('C',inplace=True)

    final_data_for_chart26 = pd.concat([data_for_chart26, weekly_data_for_chart26], axis=1)

    final_data_for_chart26.reset_index(inplace=True)

    final_data_for_chart26.to_csv("final_data_for_chart26.csv")

    #For Chart27

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','West')

    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart27 = Builder().extract_data(df1, 'C', 'P', 132, 137)
    data_for_chart27 = data_for_chart27.drop(index = 133)
    # data_for_chart3 = data_for_chart3.drop(index = 135)
    data_for_chart27.loc[132, 'D':] = data_for_chart27.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart27.loc[134, 'D':] = data_for_chart27.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart27.loc[135, 'D':] = data_for_chart27.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart27.loc[136, 'D':] = data_for_chart27.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart27.loc[137, 'D':] = data_for_chart27.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    data_for_chart27.iloc[:, 0] = None
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart27.drop(columns=columns_to_drop, inplace=True)

    updated_column_names_chart27 = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names_chart27 = Builder().convert_to_date_time(updated_column_names_chart27)

    formated_updated_column_names_chart27 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart27]

    data_for_chart27.columns = [data_for_chart6.columns[0]]+formated_updated_column_names_chart27
    

    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','West')

    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart27 = Builder().extract_data(df3, 'C', 'P', 132, 137)
    weekly_data_for_chart27 = weekly_data_for_chart27.drop(index = 133)
    # weekly_data_for_chart11 = weekly_data_for_chart11.drop(index = 135)
    weekly_data_for_chart27.loc[132, 'D':] = weekly_data_for_chart27.loc[132, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart27.loc[134, 'D':] = weekly_data_for_chart27.loc[134, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart27.loc[135, 'D':] = weekly_data_for_chart27.loc[135, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart27.loc[136, 'D':] = weekly_data_for_chart27.loc[136, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_data_for_chart27.loc[137, 'D':] = weekly_data_for_chart27.loc[137, 'D':].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart27.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart27 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart27.columns = [weekly_data_for_chart27.columns[0]]+weekly_updated_column_names_chart27

    insert_position = 1

    for i, col in enumerate(formated_updated_column_names_chart27):
        weekly_data_for_chart27.insert(insert_position+i,col,None)


    for col in weekly_updated_column_names_chart27:
        data_for_chart27[col] = None

    final_data_for_chart27 = pd.concat([weekly_data_for_chart27, data_for_chart27], axis=0)

    final_data_for_chart27 = final_data_for_chart27.fillna("")
    final_data_for_chart27.to_csv("final_data_for_chart6.csv")

    #chart28
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','West')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

    data_for_chart28 = Builder().extract_data(df1, 'C', 'P', 157, 162)
    data_for_chart28 = data_for_chart28.drop(index = 158)
    data_for_chart28 = Builder().add_row(df1, data_for_chart28, 131, 'C', 'P', 'D')


    new_rows = pd.DataFrame([['CR by Sales Tactic'] + [''] * (len(data_for_chart28.columns) - 1), ],
                        columns=data_for_chart28.columns)
    data_for_chart28 = pd.concat([data_for_chart28, new_rows], ignore_index=True)
    columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    data_for_chart28.drop(columns=columns_to_drop, inplace=True)

    # data_for_chart1.to_csv("sample_data.csv")

    updated_column_names = Builder().dates(df1,18, 'M','P')

    converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

    formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

    data_for_chart28.columns = [data_for_chart28.columns[0]]+formated_updated_column_names

    data_for_chart28.set_index('C',inplace=True)
    print(data_for_chart28)
    data_for_chart28.to_csv("data_for_chart28.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','West')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart28 = Builder().extract_data(df3, 'C', 'P', 157, 162)
    weekly_data_for_chart28 = weekly_data_for_chart28.drop(index = 158)
    
    weekly_data_for_chart28 = pd.concat([weekly_data_for_chart28, new_rows], ignore_index=True)

    weekly_data_for_chart28 = Builder().add_row(df1, weekly_data_for_chart28, 131, 'C', 'P', 'D')

    # weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['D', 'E', 'F', 'G']
    weekly_data_for_chart28.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart28 = Builder().dates(df3,18, 'H','P')

    weekly_data_for_chart28.columns = [weekly_data_for_chart28.columns[0]]+weekly_updated_column_names_chart28

    weekly_data_for_chart28.set_index('C',inplace=True)
    weekly_data_for_chart28.to_csv("weekly_data_for_chart_28.csv")

    final_data_for_chart28 = pd.concat([data_for_chart28, weekly_data_for_chart28], axis=1)

    final_data_for_chart28.reset_index(inplace=True)
    # final_data_for_chart1 = pd.merge(data_for_chart1, weekly_data_for_chart1, on = 'C', how = 'outer')
    # print(final_data_for_chart7)
    final_data_for_chart28.to_csv("final_data_for_chart28.csv")





#  #For Chart12

#     Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AS','D3','National')

#     df1 = Builder().read_excel(file_name_1, sheet_name_1)
#     df2 =Builder().read_excel(file_name_1,sheet_name_2)

#     custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
#     custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

#     df1.columns = custom_column_names_df1
#     df2.columns = custom_column_names_df2

#     data_for_chart12 = Builder().extract_data(df1, 'C', 'P', 157, 162)
#     data_for_chart12 = data_for_chart12.drop(index = 158)
#     columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
#     data_for_chart12.drop(columns=columns_to_drop, inplace=True)

#     updated_column_names_chart12 = Builder().dates(df1,155, 'M','P')

#     converted_updated_column_names_chart12 = Builder().convert_to_date_time(updated_column_names_chart12)

#     formated_updated_column_names_chart12 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart12]

#     data_for_chart12.columns = [data_for_chart12.columns[0]]+formated_updated_column_names_chart12

#     Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','GA','D3','National')

#     df3 = Builder().read_excel(file_name_1, sheet_name_1)
#     custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
#     df3.columns = custom_column_names_df3

#     weekly_data_for_chart12 = Builder().extract_data(df3, 'C', 'P', 157, 162)
#     weekly_data_for_chart12 = weekly_data_for_chart12.drop(index = 158)

#     weekly_columns_to_drop = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
#     weekly_data_for_chart12.drop(columns=weekly_columns_to_drop, inplace=True)

#     weekly_updated_column_names_chart12 = Builder().dates(df3,155, 'M','P')

#     weekly_data_for_chart12.columns = [weekly_data_for_chart12.columns[0]]+weekly_updated_column_names_chart12

#     final_data_for_chart12 = pd.concat([data_for_chart12, weekly_data_for_chart12], axis=1)
#     print(final_data_for_chart12)
#     final_data_for_chart12.to_csv("data_for_chart12.csv")


    #Updating chart1

    chart_name = "Demand Pacing - Monthly and Weekly - 1"
    dataframe = final_data_for_chart1
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name, dataframe, output_file_name)
    print("Chart-1 has been updated")
    print("")
    #Updating Chart2

    chart_name2 = "Demand Pacing - Monthly and Weekly - 2"
    dataframe2 = final_data_for_chart2
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name2, dataframe2, output_file_name)
    
    print("Chart-2 has been updated")
    print("")
    #Updating Chart3

    chart_name3 = "Demand Pacing - Monthly and Weekly - 3"
    dataframe3 = final_data_for_chart3
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name3, dataframe3, output_file_name)
    print("Chart-3 has been updated")
    print("")
    #Updating Chart4

    chart_name4 = "Demand Pacing - Monthly and Weekly - 4"
    dataframe4 = final_data_for_chart4
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name4, dataframe4, output_file_name)
    print("Chart-4 has been updated")
    print("")
    #Updating Chart5

    chart_name5 = "Demand Pacing - Monthly and Weekly - 5"
    dataframe5 = final_data_for_chart5
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name5, dataframe5, output_file_name)
    print("Chart-5 has been updated")
    print('')

    #Updating Chart6
    chart_name6 = "Demand Pacing - Monthly and Weekly - 6"
    dataframe6 = final_data_for_chart6
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name6, dataframe6, output_file_name)
    print("Chart-6 has been updated")
    print("")

    #Updating Chart7
    chart_name7 = "Demand Pacing - Monthly and Weekly - 7"
    dataframe7 = final_data_for_chart7
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name7, dataframe7, output_file_name)
    print("Chart-7 has been updated")
    print("")
    #Updating Chart8
    chart_name8 = "Demand Pacing - Monthly and Weekly - 8"
    dataframe8 = final_data_for_chart8
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name8, dataframe8, output_file_name)
    print("Chart-8 has been updated")
    print("")
    #Updating Chart9

    chart_name9 = "Demand Pacing - Monthly and Weekly - 9"
    dataframe9 = final_data_for_chart9
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name9, dataframe9, output_file_name)
    print("Chart-9 has been updated")
    print("")
    #Updating Chart10

    chart_name10 = "Demand Pacing - Monthly and Weekly - 10"
    dataframe10 = final_data_for_chart10
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name10, dataframe10, output_file_name)
    print("Chart-10 has been updated")
    print("")

    #Updating Chart11

    chart_name11 = "Demand Pacing - Monthly and Weekly - 11"
    dataframe11 = final_data_for_chart11
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name11, dataframe11, output_file_name)
    print("Chart-11 has been updated")
    print("")
    #Updating Chart12

    chart_name12 = "Demand Pacing - Monthly and Weekly - 12"
    dataframe12 = final_data_for_chart12
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name12, dataframe12, output_file_name)
    print("Chart-12 has been updated")
    print("")


    #Updating chart13

    chart_name13 = "Demand Pacing - Monthly and Weekly - 13"
    dataframe13 = final_data_for_chart13
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name13, dataframe13, output_file_name)
    print("Chart-13 has been updated")
    print("")
    #Updating Chart14

    chart_name14 = "Demand Pacing - Monthly and Weekly - 14"
    dataframe14 = final_data_for_chart14
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name14, dataframe14, output_file_name)
    
    print("Chart-14 has been updated")
    print("")
    #Updating Chart15

    chart_name15 = "Demand Pacing - Monthly and Weekly - 15"
    dataframe15 = final_data_for_chart15
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name15, dataframe15, output_file_name)
    print("Chart-15 has been updated")
    print("")
    #Updating Chart16

    chart_name16 = "Demand Pacing - Monthly and Weekly - 16"
    dataframe16 = final_data_for_chart16
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name16, dataframe16, output_file_name)
    print("Chart-16 has been updated")
    print("")
    #Updating Chart17

    chart_name17 = "Demand Pacing - Monthly and Weekly - 17"
    dataframe17 = final_data_for_chart17
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name17, dataframe17, output_file_name)
    print("Chart-17 has been updated")
    print('')

    #Updating Chart18
    chart_name18 = "Demand Pacing - Monthly and Weekly - 18"
    dataframe18 = final_data_for_chart18
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name18, dataframe18, output_file_name)
    print("Chart-18 has been updated")
    print("")

    #Updating Chart19
    chart_name19 = "Demand Pacing - Monthly and Weekly - 19"
    dataframe19 = final_data_for_chart19
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name19, dataframe19, output_file_name)
    print("Chart-19 has been updated")
    print("")
    #Updating Chart20
    chart_name20 = "Demand Pacing - Monthly and Weekly - 20"
    dataframe20 = final_data_for_chart20
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name20, dataframe20, output_file_name)
    print("Chart-20 has been updated")
    print("")
    #Updating Chart21

    chart_name21 = "Demand Pacing - Monthly and Weekly - 21"
    dataframe21 = final_data_for_chart21
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name21, dataframe21, output_file_name)
    print("Chart-21 has been updated")
    print("")
    #Updating Chart22

    chart_name22 = "Demand Pacing - Monthly and Weekly - 22"
    dataframe22 = final_data_for_chart22
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name22, dataframe22, output_file_name)
    print("Chart-22 has been updated")
    print("")

    #Updating Chart23

    chart_name23 = "Demand Pacing - Monthly and Weekly - 23"
    dataframe23 = final_data_for_chart23
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name23, dataframe23, output_file_name)
    print("Chart-23 has been updated")
    print("")
    #Updating Chart24

    chart_name24 = "Demand Pacing - Monthly and Weekly - 24"
    dataframe24 = final_data_for_chart24
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name24, dataframe24, output_file_name)
    print("Chart-24 has been updated")
    print("")


    #Updating Chart25

    chart_name25 = "Demand Pacing - Monthly and Weekly - 25"
    dataframe25 = final_data_for_chart25
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name25, dataframe25, output_file_name)
    print("Chart-25 has been updated")
    print("")
    #Updating Chart26

    chart_name26 = "Demand Pacing - Monthly and Weekly - 26"
    dataframe26 = final_data_for_chart26
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name26, dataframe26, output_file_name)
    print("Chart-26 has been updated")
    print("")

    #Updating Chart27

    chart_name27 = "Demand Pacing - Monthly and Weekly - 27"
    dataframe27 = final_data_for_chart27
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name27, dataframe27, output_file_name)
    print("Chart-27 has been updated")
    print("")
    #Updating Chart28

    chart_name28 = "Demand Pacing - Monthly and Weekly - 28"
    dataframe28 = final_data_for_chart28
    output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

    Thinkcell().update_chart(chart_name28, dataframe28, output_file_name)
    print("Chart-28 has been updated")
    print("")










if __name__ == "__main__":
    main()