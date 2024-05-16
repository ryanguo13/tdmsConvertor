import pandas as pd
from nptdms import TdmsFile
# import xlwt
def tdms_to_xls(tdms_path, xls_path):
       # Load TDMS file
       tdms_file = TdmsFile.read(tdms_path)
       
       # Create a Pandas Excel writer using xlwt (for .xls)
       writer = pd.ExcelWriter(xls_path, engine='openpyxl')
       
       # Iterate through all groups
       for group in tdms_file.groups():
           # Convert each group to a dataframe
           group_df = group.as_dataframe()
           # Write each group to a separate sheet
           group_df.to_excel(writer, sheet_name=group.name)

       # Save the Excel file
       writer._save()
       print(f'TDMS file "{tdms_path}" has been converted to Excel file "{xls_path}".')
   
   # Replace with the path to your TDMS file and desired Excel output file path
my_tdms_path = 'example.tdms'
my_xls_path = 'output.xlsx'
tdms_to_xls(my_tdms_path, my_xls_path)