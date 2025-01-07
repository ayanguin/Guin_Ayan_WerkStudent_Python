import sys
import pdfplumber as pf
import pandas as pd
import re
from datetime import datetime as dt
import locale
import openpyxl
import os

# Set locale to German (Germany) for date parsing
locale.setlocale(locale.LC_TIME, 'de_DE')

# Get the directory where the execulable script is located
current_dir = os.path.dirname(os.path.abspath(sys.executable))

# Define the path to the first PDF invoice
url = os.path.join(current_dir, "sample_invoice_1.pdf")
pdf = pf.open(url) #Open the PDF


## Processing the 1st Invoice (sample_invoice_1.pdf) ##
# Extract the first table from the first page
table = pdf.pages[0].extract_tables()
df = pd.DataFrame(table[1][1:], columns=table[1][0])  # Create DataFrame from 1st table data

# Select the relevant rows (reversing the table and excluding the last 18 rows and drop the columns with missing values in any column)
extracted_vals = df[:18:-1].dropna(axis=1)

# Extract the date from the first table (the date is in the first row)
date = pd.DataFrame(table[0][1:], columns=table[0][0])
date = date['Date'].values[0]

# Parse the date string into a datetime object and format it as "dd/mm/yyyy"
date_obj = dt.strptime(date,"%d. %B %Y")
date = date_obj.strftime("%d/%m/%Y")

# Extract the value for column and the value from the last row of the extracted data
value_column = extracted_vals.values[len(extracted_vals)-1][0]
value = float(extracted_vals.values[len(extracted_vals)-1][1].replace('€', '').replace(" ","").replace(',', '.')) # Further trimming the data as per our need and changing the data type to float for arithmetic calculation.



## Processing the 2nd Invoice (sample_invoice_2.pdf) ##
# Define the path to the second PDF invoice
url2 = os.path.join(current_dir, "sample_invoice_2.pdf")
pdf2 = pf.open(url2)# Open the pdf

# Extract the first table from the first page
table2 = pdf2.pages[0].extract_tables() 
df2 = pd.DataFrame(table2[0][1:],columns=["Total"," "," ","Amount"]).dropna(axis=1) # Create DataFrame from table data and drop the columns with missing values.

# Extract the value column and value from the last row
value_column2 = df2.values[len(df2)-1][0]
value2 = float(df2.values[len(df2)-1][1].split("$")[1])

# Parcing through the rxtracted text lines from the pdf
for i in range (0,len(pdf2.pages[0].extract_text_lines())):
    # Search for the line containing "Invoice.*2016" to identify the date
    if re.search(pattern="^Invoice.*2016$", string=pdf2.pages[0].extract_text_lines()[i]['text']):
        # Storing the date and the column for further processing in two veriables
        date2_column = pdf2.pages[0].extract_text_lines()[i]['text'].split(":")[0] 
        date2 = pdf2.pages[0].extract_text_lines()[i]['text'].split(":")[1]

# Parse the date string into a datetime object and format it as "dd/mm/yyyy"
date_obj2 = dt.strptime(date2.strip(),"%b %d, %Y") 
date2 = date_obj2.strftime("%d/%m/%Y")

## Creating the Excel file ##
# Define the path to the output Excel file
excel_url = os.path.join(current_dir, "WerkStudent_Python_Excel.xlsx")

# Prepare data for the Excel sheet
# Create a list of dictionaries to store the extracted data into excel
excel_data = [
    {"File Name": "Sample_invoice_1", "Date": date, "Value": value},
    {"File Name": "Sample_invoice_2", "Date": date2, "Value": value2},
    ]

# Create a pandas DataFrame from the list of dictionaries
e_data = pd.DataFrame(excel_data)

# Check if the Excel file already exists and delete it if found
if os.path.exists(excel_url):
    os.remove(excel_url)
    
# Create an ExcelWriter object using the openpyxl engine
with pd.ExcelWriter(excel_url, engine='openpyxl') as writer:
    # Write the original DataFrame to the first sheet ('sheet1')
    e_data.to_excel(writer, sheet_name='sheet1', index=False)
    
    # Create a pivot table summarizing the data by Date and File Name
    e_data_pivot = e_data.pivot_table(
        index= ['Date','File Name'],
        values='Value',
        aggfunc='sum'
    )
    
    # Reset the index to create columns for 'Date' and 'File Name'
    e_data_pivot.reset_index(inplace=True)
    
    # Write the pivot table to the second sheet ('sheet2')
    e_data_pivot.to_excel(writer, sheet_name='sheet2',index=False)

# Define the path to the output CSV file
csv_url = os.path.join(current_dir, "WerkStudent_Python_CSV.csv")

# Check if the CSV file already exists and delete it if found
if os.path.exists(csv_url):
    os.remove(csv_url)
    
# Write the original DataFrame to a CSV file with semicolon as the delimiter
e_data.to_csv(csv_url, index=False, sep=';')

