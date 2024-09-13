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
    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Monthly", 'D2','AV','D3','National')
    
    df1 = Builder().read_excel(file_name_1, sheet_name_1)
    df2 =Builder().read_excel(file_name_1,sheet_name_2)

    custom_column_names_df1 = Builder().generate_columns(df1.shape[1])
    custom_column_names_df2 = Builder().generate_columns(df2.shape[1])

    df1.columns = custom_column_names_df1
    df2.columns = custom_column_names_df2

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
    print(data_for_chart1)
    data_for_chart1.to_csv("data_for_chart1.csv")


    Write_Excel().modify_excel(file_path, sheet_name_1, 'D4', "Weekly", 'D2','AV','D3','National')
    df3 = Builder().read_excel(file_name_1, sheet_name_1)
    custom_column_names_df3 = Builder().generate_columns(df3.shape[1])
    df3.columns = custom_column_names_df3

    weekly_data_for_chart1 = Builder().extract_data(df3, 'D', 'P', 20, 26)
    # weekly_data_for_chart1 = weekly_data_for_chart1.drop(index = 158)
    
    weekly_data_for_chart1 = pd.concat([weekly_data_for_chart1, new_rows], ignore_index=True)

    weekly_data_for_chart1 = Builder().add_row(df1, weekly_data_for_chart1, 52, 'C', 'P', 'D')

    weekly_data_for_chart1 = Builder().add_row(df1,weekly_data_for_chart1,28,'C','P','D')
    weekly_columns_to_drop = ['C','D', 'E', 'F', 'G']
    weekly_data_for_chart1.drop(columns=weekly_columns_to_drop, inplace=True)

    weekly_updated_column_names_chart1 = Builder().dates(df3,18, 'G','P')

    weekly_data_for_chart1.columns = [weekly_data_for_chart1.columns[0]]+weekly_updated_column_names_chart1

    final_data_for_chart1 = pd.concat([data_for_chart1, weekly_data_for_chart1], axis=1)
    print(final_data_for_chart1)
    final_data_for_chart1.to_csv("data_for_chart1.csv")


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

if __name__ == "__main__":
    main()