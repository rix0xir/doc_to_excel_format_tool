import pandas as pd

# Load the data from Excel
excel_path = 'C:/Users/Ayman/Documents/Abhijit_mail_attachments/Test_PW.xlsm'
granted_df = pd.read_excel(excel_path, sheet_name='Granted')

# Print column names
print(granted_df.columns)
