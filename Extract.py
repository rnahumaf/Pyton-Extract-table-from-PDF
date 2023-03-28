# import os
# os.chdir("") # Change the working directory, if needed

import PyPDF2
import pdfplumber
import openpyxl
from openpyxl import Workbook



# Set the PDF file name and range of pages containing the table
pdf_file = 'drugs.pdf'
start_page = 9
end_page = 60

# Create a new Excel workbook and worksheet
wb = Workbook()
ws = wb.active

# Open the PDF file with pdfplumber
with pdfplumber.open(pdf_file) as pdf:
    # Loop through the specified range of pages
    for page_number in range(start_page - 1, end_page):
        page = pdf.pages[page_number]

        # Extract the table from the page
        table = page.extract_table()

        # Loop through the rows and columns of the table
        for row in table:
            ws.append(row)

# Save the workbook as an XLSX file
wb.save('drugs.xlsx')
