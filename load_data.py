try:
    import pandas as pd
except ImportError:
    print("pandas not installed. Install with: pip install pandas openpyxl")
    exit(1)

# Load the Excel file
file_path = "/Users/yudhajatmiko/CFI/Mid retailer analysisi/Mid Retailer Financial Analysis - Blank.xlsx"
try:
    # First, check all sheets in the workbook
    excel_file = pd.ExcelFile(file_path)
    print("Available sheets:", excel_file.sheet_names)
    print("-" * 50)
    
    # Load the first sheet and examine its structure
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    print("Raw data shape:", df_raw.shape)
    print("\nFirst 10 rows of raw data:")
    print(df_raw.head(10))
    print("-" * 50)
    
    # Try to find the actual data by skipping empty rows
    df_cleaned = df_raw.dropna(how='all').dropna(axis=1, how='all')
    
    if not df_cleaned.empty:
        print("Cleaned data shape:", df_cleaned.shape)
        print("\nCleaned data:")
        print(df_cleaned.head(10))
        
        # Try to identify header row (look for row with most non-null values)
        header_row = 0
        max_non_null = 0
        for i in range(min(10, len(df_cleaned))):
            non_null_count = df_cleaned.iloc[i].notna().sum()
            if non_null_count > max_non_null:
                max_non_null = non_null_count
                header_row = i
        
        print(f"\nPotential header row: {header_row}")
        print("Header row content:", df_cleaned.iloc[header_row].tolist())
        
        # Load with identified header
        if header_row > 0:
            df = pd.read_excel(file_path, sheet_name=0, header=header_row)
        else:
            df = pd.read_excel(file_path, sheet_name=0)
            
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        print("\nFinal processed data:")
        print("Shape:", df.shape)
        print("Columns:", df.columns.tolist())
        print("\nFirst few rows:")
        print(df.head())
        
    else:
        print("The Excel file appears to be empty or contains only blank cells.")
        
except FileNotFoundError:
    print(f"File '{file_path}' not found")
    exit(1)
except Exception as e:
    print(f"Error loading file: {e}")
    exit(1)