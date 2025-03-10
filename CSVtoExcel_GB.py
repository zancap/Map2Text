import pandas as pd

# Read the CSV file; update the path if needed
df = pd.read_csv('C:/Users/nikyt/Documents/IDL M2/Stage/Data/GBPN.csv', encoding='utf-8')

# Define the output Excel file path
output_excel = 'C:/Users/nikyt/Documents/IDL M2/Stage/Data/GBPN.xlsx'

# Create an Excel writer using the XlsxWriter engine
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    # Write the DataFrame to the Excel file; header row will be row 0
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # Access the XlsxWriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Determine the number of rows and columns.
    # Note: df.shape[0] is the number of data rows (header is written separately)
    max_row, max_col = df.shape
    
    # Apply an autofilter over the full range (header + data).
    # Since XlsxWriter uses 0-indexed rows and columns:
    # - Row 0 is the header.
    # - Data rows go from 1 to max_row.
    # So, we set the filter from row 0 to row max_row (inclusive) and from column 0 to max_col-1.
    worksheet.autofilter(0, 0, max_row, max_col - 1)

print("Excel file created successfully with proper columns and autofilters!")
