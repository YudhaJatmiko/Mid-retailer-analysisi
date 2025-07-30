import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Set display options for better output
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Load the Excel file
file_path = "/Users/yudhajatmiko/CFI/Mid retailer analysisi/Mid Retailer Financial Analysis - Blank.xlsx"

def load_and_examine_sheets():
    """Load and examine all sheets in the Excel file"""
    excel_file = pd.ExcelFile(file_path)
    print("="*80)
    print("EXCEL SHEET ANALYSIS")
    print("="*80)
    
    sheets_data = {}
    
    for sheet_name in excel_file.sheet_names:
        print(f"\n--- SHEET: {sheet_name} ---")
        try:
            # Load sheet with different approaches to handle formatting
            df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Remove completely empty rows and columns
            df_cleaned = df_raw.dropna(how='all').dropna(axis=1, how='all')
            
            if not df_cleaned.empty:
                print(f"Shape: {df_cleaned.shape}")
                print(f"Non-empty data preview:")
                print(df_cleaned.head(10))
                
                # Store the cleaned data
                sheets_data[sheet_name] = df_cleaned
            else:
                print("Sheet appears to be empty")
                
        except Exception as e:
            print(f"Error reading sheet {sheet_name}: {e}")
    
    return sheets_data

def analyze_financial_statements(sheets_data):
    """Analyze the Financial Statements sheet"""
    if 'Financial Statements' not in sheets_data:
        print("Financial Statements sheet not found")
        return None
    
    print("\n" + "="*80)
    print("FINANCIAL STATEMENTS ANALYSIS")
    print("="*80)
    
    df = sheets_data['Financial Statements']
    print(f"Raw Financial Statements shape: {df.shape}")
    print(f"First 15 rows:")
    print(df.head(15))
    
    # Try to identify the structure - look for typical financial statement headers
    financial_keywords = ['revenue', 'sales', 'income', 'expense', 'assets', 'liabilities', 'equity', 'cash', 'year']
    
    for i, row in df.iterrows():
        row_str = ' '.join([str(cell).lower() for cell in row if pd.notna(cell)])
        if any(keyword in row_str for keyword in financial_keywords):
            print(f"Found potential financial data at row {i}: {row.tolist()}")
    
    return df

def analyze_ratio_template(sheets_data):
    """Analyze the Ratio Calculations sheet to understand the template"""
    if 'Ratio Calculations' not in sheets_data:
        print("Ratio Calculations sheet not found")
        return None
    
    print("\n" + "="*80)
    print("RATIO CALCULATIONS TEMPLATE ANALYSIS")
    print("="*80)
    
    df = sheets_data['Ratio Calculations']
    print(f"Ratio Calculations shape: {df.shape}")
    print(f"Content preview:")
    print(df.head(20))
    
    return df

def analyze_dupont_templates(sheets_data):
    """Analyze the DuPont analysis sheets"""
    dupont_sheets = ['3 Step DuPont Pyramid ', '5 Step DuPont Pyramid']
    
    for sheet_name in dupont_sheets:
        if sheet_name in sheets_data:
            print(f"\n" + "="*80)
            print(f"{sheet_name.upper()} TEMPLATE ANALYSIS")
            print("="*80)
            
            df = sheets_data[sheet_name]
            print(f"Shape: {df.shape}")
            print(f"Content preview:")
            print(df.head(20))

# Main execution
if __name__ == "__main__":
    # Load and examine all sheets
    sheets_data = load_and_examine_sheets()
    
    # Analyze each sheet
    financial_df = analyze_financial_statements(sheets_data)
    ratio_df = analyze_ratio_template(sheets_data)
    analyze_dupont_templates(sheets_data)
    
    print("\n" + "="*80)
    print("ANALYSIS COMPLETE")
    print("="*80)
    print("Next steps will depend on the actual data structure found in the Financial Statements sheet")