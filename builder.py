from thinkcellbuilder import Presentation, Template
import pandas as pd
from datetime import datetime
from string import ascii_uppercase
import numpy as np
class Builder:
    def read_excel(self, file_name, sheet_naming):
        df = pd.read_excel(file_name, sheet_name = sheet_naming, engine = 'pyxlsb')
        i  = 0
        while(i<2):
            null_rows = pd.DataFrame([np.nan]*len(df.columns)).transpose()
            null_rows.columns = df.columns
            df = pd.concat([null_rows, df]).reset_index(drop=True)
            i = i+1
        return df

    def generate_columns(self, n):
        column_names = []
        for i in range(n):
            name = ""
            while i>=0:
                name = ascii_uppercase[i % 26] + name
                i = i // 26 - 1
            column_names.append(name)
        return column_names
    
    def add_column(self, df, column, r1, r2):
        return df[column][r1:r2].tolist()

    def add_row(self, df_row_copy,df_row_paste, row1, c1, c2, c3):
        if c3 != None:
            df_row_paste.loc[row1] = df_row_copy.loc[row1, c1:c2]
            df_row_paste.loc[row1, c3:] = df_row_paste.loc[row1, c3:].apply(lambda x: x*100).apply(lambda x: float(f"{x:.1f}"))
            return df_row_paste
        else:
            df_row_paste.loc[row1] = df_row_copy.loc[row1, c1:c2]
            return df_row_paste
    
    def dates(self, df, row, c1, c2):
        return df.loc[row,c1:c2].tolist()
    
    def extract_data(self, df, c1, c2, r1, r2):
        return df.loc[r1:r2,c1:c2]
    
    def convert_to_date_time(self, column_list):
        column_names_1=[]
        column_names_1 = [pd.to_datetime(i, origin='1899-12-30', unit='D').strftime('%Y-%m-%d') for i in column_list]
        return column_names_1
    
    def format_date_time(self, column_value):
        date_obj = datetime.strptime(column_value, "%Y-%m-%d")
        updated_formated_date = date_obj.strftime("%b'%y")
        return updated_formated_date
