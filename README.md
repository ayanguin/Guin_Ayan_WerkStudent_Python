# WerkStudent_Python
Submitted by Ayan Guin.

## Overview

This repository contains the interview task for the WerkStudent position in Python. The goal is to collect data from two sample invoices, create an Excel file with two sheets, and generate a CSV file. Additionally, an executable file should be provided to run the code.

## Task Details

1. **Data Extraction**:
    - Extract specific values from three sample invoices.
    - For Sample 1, extract the value shown in the provided image.
    - <img width="289" alt="image" src="https://github.com/user-attachments/assets/0cf000ff-c305-4ffe-beb4-1c02a04d06b6" />
    - For Samples 2, extract the value shown in the provided image.
    - <img width="497" alt="image" src="https://github.com/user-attachments/assets/ea6eb368-604d-4dd4-9235-fbc8ec36d275" />

2. **Excel File Creation**:
    - Create an Excel file with two sheets:
        - **Sheet 1**: Contains three columns - File Name, Date (scraped from the document), and Value.
        - **Sheet 2**: Contains a pivot table with the date and value sum, and also by document name.

3. **CSV File Creation**:
    - Create a CSV file with all the data, including headers, and use a semicolon (;) as the separator.

4. **Executable File**:
    - Provide an executable file (.exe) that can run the code if the files are in the same folder.

5. **Fork Creation**:
    - Create a fork of this repository named `LastName_FirstName_WerkStudent_Python` (e.g., `Shovon_Golam_WerkStudent_Python`).
    - Upload your code to this branch. No need to submit a pull request; the fork will be checked directly.

6. **Documentation**:
    - Include an explanation in the README file that a non-technical person can understand.
    - Ensure the code is documented so that a technical person can understand it.

7. **Problem Reporting**:
    - If you face any problems or find it impossible to complete a task, document the issue in the README file of your branch. Explain what the problem was and why you were unable to complete it.


## How It Works

1. **Data Extraction**:
    - The script reads the sample invoices and extracts the required values.
    - The extracted data is stored in variables for further processing.

2. **Excel File Creation**:
    - The script creates an Excel file with two sheets.
    - Sheet 1 contains the file name, extracted data, and value.
    - Sheet 2 contains a pivot table summarizing the data by date and document name.

3. **CSV File Creation**:
    - The script generates a CSV file with the extracted data, including headers, and uses a semicolon as the separator.

4. **Executable File**:
    - An executable file is provided to run the entire code. Ensure the sample invoices are in the same folder as the executable file.

5. **Requirements File**:
    -A requirements.txt file is included to create the environment needed to run the code

## Running the Code

1. Place the sample invoices in the same folder as the executable file.
2. If needed, for setting up the environment use the requirements.text file.
3. Run the executable file to execute the code and generate the Excel and CSV files.

## Explanation of the code
- This Python code processes two invoices in PDF format and creates an Excel file and a CSV file summarizing the information.

1. **Setting Up**: 
- The script first ensures the code uses German locale settings for proper date parsing.
- Then it  finds the location of the script itself on your computer.
- Then, defines the file paths for the two invoices (sample_invoice_1.pdf and sample_invoice_2.pdf) and creates an Excel file (WerkStudent_Python_Excel.xlsx) and a CSV file (WerkStudent_Python_CSV.csv)

2. **Processing Invoice 1 (sample_invoice_1.pdf)**:
- The program opens the first invoice PDF.
- Then it  extracts the first table from the first page of the PDF, assuming that's where the invoice details are located.
- Then it  selects the relevant rows from the table, excluding rows with missing information in any column.
- Then it  extracts the date from the pdf.
- Then it  converts the extracted date from text format (e.g., "15. März 2024") to a more standard format ("dd/mm/yyyy").
- Then it  will  extracts the  required value (Total) and cleans up any symbols or extra spaces in the value and converts it to a number for further operationas.

3. **Processing Invoice 2 (sample_invoice_2.pdf)**:
- The program will again follows similar steps as for invoice 1: opening the PDF, extracting the values, selecting relevant values, and so on.
- Meanwhile, it will searches for a line containing the date(assuming the date format is different for this invoice) to identify the date.
- Then it will extracts the date and the column containing the date from that line.
- Then it  converts the extracted date to a standard format ("dd/mm/yyyy").
- Then it  extracts the value(Total) from the selected data, similar to invoice 1.

4. **Creating the Excel File**:
- The script then creates a list containing information extracted from both invoices, including file name, date, and value.
- Then it converts this list into a data frame, which is a more organised way to store and managing the data.
- Then it  writes this data frame to the first sheet ("sheet1") of the Excel file.
- Then it  creates a pivot table to summarise the data by invoice date and file name (showing the total value for each combination).
- Then it  writes the pivot table to the second sheet ("sheet2") of the Excel file.

5. **Creating the CSV File**:
- Finally, it will writes the data frame (containing details for both invoices) to a CSV file, using a semicolon (';') as the separator between the values.

**Conclusion**:
- Overall, this code helps you efficiently extract key information (date and value) from two invoices, organise it in a clear format, and save it in both an Excel file and a CSV file for further analysis or record keeping.


## Problem Reporting

- As per the instruction, I have extract the specific values from the 2 PDFs and created the Excel and CSV accordingly. My program is designed to summarise the value based on the date and document name in CSV file, as the instructions were to fetch only one value from each documents, thus the resultant output is looks like nothing is happening there with the final output. If we have chances to extract more values from the pdfs it would have give more appropriate output.

- I am using mac os(Silicon), so the generated executable file is only compatible with mac os(silicon), I am not sure if it could be run using the mac os(Intel) or windows pc, as I have no access to those machines during the whole implementation.



