# import win32com.client as win32
# import psutil
# class Write_Excel:
#     def modify_excel(self,file_path, sheet_name, cell_address1, new_value1,cell_address2, new_value2):
    
#         excel = win32.Dispatch('Excel.Application')
#         excel.Visible = False
#         workbook = excel.Workbooks.Open(file_path)
#         sheet = workbook.Sheets(sheet_name)
#         sheet.Range(cell_address1).Value = new_value1
#         sheet.Range(cell_address2).Value = new_value2
#         workbook.Save()
#         workbook.Close()
#         excel.Quit()
#     def close_all_excel_instances(self):
#     # Iterate through all running processes
#         for process in psutil.process_iter(['pid', 'name']):
#             if 'EXCEL.EXE' in process.name():  # Check if process belongs to Excel
#                 try:
#                 # Terminate the Excel process
#                     process.terminate()
#                     print(f"Terminated Excel process with PID {process.pid}")
#                 except Exception as e:
#                     print(f"Failed to terminate Excel process with PID {process.pid}: {e}")
import win32com.client as win32
from win32com.client import constants as c
import psutil

class Write_Excel:
    def __init__(self):
        # Try to get an existing instance of Excel or create a new one
        try:
            self.excel = win32.GetActiveObject("Excel.Application")
            print("Found existing Excel instance.")
        except:
            self.excel = win32.Dispatch("Excel.Application")
            print("Created new Excel instance.")
    
    def modify_excel(self, file_path, sheet_name, cell_address1, new_value1, cell_address2, new_value2,cell_address3, new_value3):
        try:
            # Set Excel visibility to False
            self.excel.Visible = False
            
            # Open workbook and worksheet
            workbook = self.excel.Workbooks.Open(file_path)
            sheet = workbook.Sheets(sheet_name)
            
            # Modify cell values
            sheet.Range(cell_address1).Value = new_value1
            sheet.Range(cell_address2).Value = new_value2
            sheet.Range(cell_address3).Value = new_value3
            
            # Save workbook and close
            workbook.Save()
            workbook.Close()
            
        except Exception as e:
            print(f"An error occurred: {e}")
            
        finally:
            # Quit Excel application
            if 'excel' in locals():
                self.excel.Quit()
    def close_all_excel_instances(self):
    # Iterate through all running processes
        for process in psutil.process_iter(['pid', 'name']):
            if 'EXCEL.EXE' in process.name():  # Check if process belongs to Excel
                try:
                # Terminate the Excel process
                    process.terminate()
                    print(f"Terminated Excel process with PID {process.pid}")
                except Exception as e:
                    print(f"Failed to terminate Excel process with PID {process.pid}: {e}")


