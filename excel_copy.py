# import win32com.client as win32
# import psutil
# import keyboard
# import time

# class Excel_Copy:
#     def copy(self):
#         template_file = r'C:\Users\imran.s\Desktop\POC Disney\Thinkcell_Automation\storage\Weekly Leads Summary Templates (1).xlsb'
#         target_file = r'C:\Users\imran.s\Desktop\POC Disney\Thinkcell_Automation\storage\20240528_Weekly_Leads_Summary_0525_v3.xlsb'
#         sheet_name = 'By Marketing Channel (TEMPLATE)'

#         excel = win32.Dispatch('Excel.Application')
#         excel.Visible = False  

#         template_wb = excel.Workbooks.Open(template_file)

#         target_wb = excel.Workbooks.Open(target_file)

#         template_wb.Sheets(sheet_name).Copy(Before=target_wb.Sheets(1))

#         target_wb.Save()
#         target_wb.Close()
#         template_wb.Save()
#         template_wb.Close()
#         time.sleep(20)
#         keyboard.press_and_release('enter')

#         excel.Quit()
#         time.sleep(20)
#         keyboard.press_and_release('enter')
    
#         for process in psutil.process_iter(['pid', 'name']):
#             if 'EXCEL.EXE' in process.name():  # Check if process belongs to Excel
#                 try:
#                 # Terminate the Excel process
#                     process.terminate()
#                     print(f"Terminated Excel process with PID {process.pid}")
#                 except Exception as e:
#                     print(f"Failed to terminate Excel process with PID {process.pid}: {e}")

import win32com.client as win32
import time
import psutil
import keyboard
class Excel_Copy:
    def copy(self):
        template_file = r'C:\Users\imran.s\Desktop\POC\Thinkcell_Automation\storage\Weekly Leads Summary Templates (1).xlsb'
        target_file = r'C:\Users\imran.s\Desktop\POC\Thinkcell_Automation\storage\downloaded_file.xlsb'
        sheet_name = 'By Marketing Channel (TEMPLATE)'

        try:
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False  
            excel.DisplayAlerts = False  # Disable alerts

            template_wb = excel.Workbooks.Open(template_file)

            target_wb = excel.Workbooks.Open(target_file)

            template_wb.Sheets(sheet_name).Copy(Before=target_wb.Sheets(1))

            target_wb.Save()
            target_wb.Close()
            template_wb.Save()
            template_wb.Close()

            # Pause for 20 seconds
            time.sleep(10)

            # Simulate pressing Enter key
            
            keyboard.press_and_release('enter')

        except Exception as e:
            print(f"Error occurred: {e}")

        finally:
            excel.Quit()
            time.sleep(10)
            keyboard.press_and_release('enter')
    
            for process in psutil.process_iter(['pid', 'name']):
                if 'EXCEL.EXE' in process.name():  # Check if process belongs to Excel
                    try:
                        # Terminate the Excel process
                        process.terminate()
                        print(f"Terminated Excel process with PID {process.pid}")
                    except Exception as e:
                        print(f"Failed to terminate Excel process with PID {process.pid}: {e}")

# Example usage:
if __name__ == "__main__":
    excel_copy = Excel_Copy()
    excel_copy.copy()
