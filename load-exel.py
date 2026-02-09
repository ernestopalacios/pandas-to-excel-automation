import pandas as pd
import datetime
import os
import uuid  # For optional UUID conversion

def main():
    # Hardcoded paths (adjust as needed; assumes the file from main.py output)
    excel_path = 'output/20240209_template.xlsx'  # Replace with your actual edited file name/path
    
    # Generate output pickle filename with date prefix (e.g., 20240209_edited_data.pkl)
    today = datetime.date.today().strftime('%Y%m%d')
    output_pickle_filename = f"{today}_edited_data.pkl"
    output_pickle_path = os.path.join('pickles', output_pickle_filename)
    
    # Load the Excel file into a Pandas DataFrame
    # Read from the second sheet (0-based index 1; adjust if your sheet is named)
    # Skip no rows (assumes row 1 is headers)
    df = pd.read_excel(
        excel_path,
        sheet_name=1,  # Or use sheet_name='Sheet2' if named
        header=0,      # Row 1 as headers
        engine='openpyxl'  # Ensures compatibility with your formatted file
    )
    
    # Optional: Clean up or convert data types if needed
    # For example, strip any extra whitespace from strings
    df = df.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)
    
    # Optional: Convert 'uuid' column back from strings to uuid.UUID objects
    # (To match your original pickle format; uncomment if needed)
    if 'uuid' in df.columns:
        df['uuid'] = df['uuid'].apply(lambda x: uuid.UUID(x) if pd.notna(x) else None)
    
    # Print for verification
    print(f"Loaded DataFrame with {len(df)} rows and columns: {df.columns.tolist()}")
    print(df.head())  # Show first few rows
    
    # Save the DataFrame as a pickle file
    df.to_pickle(output_pickle_path)
    print(f"DataFrame saved as pickle to {output_pickle_path}")

if __name__ == "__main__":
    main()