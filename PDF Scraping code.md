# sid-PDF-scrape
import os
import pandas as pd
from tabula import read_pdf

# Define input folder and output Excel file
path = os.getcwd()
pdf_folder = path + '\data'  # Folder where PDFs are stored currently using 2009 data
output_excel = "Extracted_pdfs.xlsx"  # Output Excel file

os.makedirs(pdf_folder, exist_ok=True)
pdf_files = sorted([f for f in os.listdir(pdf_folder) if f.endswith(".pdf")])

# Create an Excel writer
df_dict = {}
df_num = 1
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)

        # Skip empty PDFs
        if os.path.getsize(pdf_path) == 0:
            print(f"Skipping empty file: {pdf_file}")
            continue

        try:
            # Extract tables from PDF
            tables = read_pdf(pdf_path, pages="all", multiple_tables=True)

            for i, df in enumerate(tables):
                sheet_name = f"{os.path.splitext(pdf_file)[0]}_Table{i+1}"
                df_dict[df_num] = df
                df_num = df_num +1  
                df_cleaned= df.dropna(how= 'any')
 
                if df_cleaned.empty:
                    continue
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Extracted tables from {pdf_file}")
        
        except Exception as e:
            print(f"Error reading {pdf_file}: {e}")

print(f"Extracted tables saved to: {output_excel}")

