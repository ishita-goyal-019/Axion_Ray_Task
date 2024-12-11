# Improved Python script for cleaning PDF data
import fitz
import pandas as pd
import numpy as np

def clean_pdf_data(pdf_path, output_excel):
    # Open and read PDF
    pdf_document = fitz.open(pdf_path)
    text_content = ""
    
    # Extract text from all pages
    for page in pdf_document:
        text_content += page.get_text()
    
    # Split into lines and clean empty lines
    lines = [line.strip() for line in text_content.split('\
') if line.strip()]
    
    # Identify the header row (looking for key columns like 'VIN')
    header_index = -1
    for i, line in enumerate(lines):
        if 'VIN' in line:
            header_index = i
            break
    
    if header_index != -1:
        # Extract headers
        headers = lines[header_index:header_index + 32]  # Based on the number of columns we saw
        
        # Process the remaining lines as data
        data_lines = lines[header_index + 32:]
        
        # Create structured data
        data_dict = {header: [] for header in headers}
        current_col = 0
        
        for line in data_lines:
            if current_col < len(headers):
                data_dict[headers[current_col]].append(line)
                current_col += 1
            if current_col == len(headers):
                current_col = 0
        
        # Create DataFrame
        df = pd.DataFrame(data_dict)
        
        # Clean up any obvious errors
        df = df.replace('', np.nan)
        df = df.dropna(how='all')
        
        # Save to Excel
        df.to_excel(output_excel, index=False)
        print("Data cleaned and saved successfully")
        return df
    else:
        print("Could not find header row")
        return None

# Clean the data
df_cleaned = clean_pdf_data("Copy of Task 2.pdf", "Task_2_Final_Cleaned.xlsx")

# Display sample of cleaned data
if df_cleaned is not None:
    print("\
First few rows of cleaned data:")
    print(df_cleaned.head())
    
    print("\
DataFrame shape:")
    print(df_cleaned.shape)
    
    print("\
Columns in cleaned dataset:")
    print(df_cleaned.columns.tolist())