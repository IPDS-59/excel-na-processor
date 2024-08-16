"""
Excel Data Processor for BPS Tables

This script processes Excel files containing BPS (Badan Pusat Statistik) data tables.
It filters data based on user input, creates new sheets, and applies specific
data processing rules.

@author: Fajrian Aidil Pratama
@email: fajrianaidilp@gmail.com
@created: 2024-08-16
@last_modified: 2024-08-16
@description: Processes BPS Excel tables and generates a new Excel file with processed data
"""

import pandas as pd
import os
import glob
import re

def format_animal_name(name):
    """
    Format animal name by capitalizing each word and replacing underscores with spaces.
    
    :param name: The animal name to format (e.g., 'puyuh_pedaging')
    :return: The formatted animal name (e.g., 'Puyuh Pedaging')
    """
    # Split the name by underscore
    words = name.split('_')
    # Capitalize each word and join with space
    return ' '.join(word.capitalize() for word in words)

def get_kabupaten_name(df, kab_code):
    """Get the kabupaten name from the dataframe and clean it."""
    kab_row = df[df['kab'] == kab_code].iloc[0]
    full_name = kab_row.get('id_kab', str(kab_code))  # Fallback to kab_code if nama_kab not found
    # Remove [kode kab] part
    clean_name = re.sub(r'\[.*?\]\s*', '', full_name).strip()
    # Replace spaces with underscores and remove any non-alphanumeric characters
    return re.sub(r'\W+', '_', clean_name)

def get_column_keywords():
    """Get column keywords from user input."""
    print("Enter keywords for columns in derived table (comma-separated):")
    print("1. rerata")
    print("2. populasi")
    print("3. Other (specify)")
    choice = get_input_with_default("Choose option (1/2/3)", "1", str)
    
    if choice == "1":
        return ["rerata"]
    elif choice == "2":
        return ["populasi"]
    elif choice == "3":
        keywords = input("Enter custom keywords (comma-separated): ")
        return [k.strip().lower() for k in keywords.split(",")]
    else:
        print("Invalid choice. Using default 'rerata'.")
        return ["rerata"]

def get_input_with_default(prompt: str, default: any, input_type: type = str) -> any:
    """
    Get user input with a default value and type checking.
    
    :param prompt: The prompt to display to the user
    :param default: The default value if user input is empty
    :param input_type: The expected type of the input (default is str)
    :return: The user input or default value, converted to the specified type
    """
    while True:
        user_input = input(f"{prompt} (default: {default}): ")
        if user_input == "":
            return default
        try:
            return input_type(user_input)
        except ValueError:
            print(f"Invalid input. Please enter a valid {input_type.__name__}.")

def get_user_input():
    """Get user input for kabupaten code and table codes."""
    kab_code = get_input_with_default("Masukkan kode kabupaten (contoh: 7205): ", 7205, int)
    ref_table = get_input_with_default("Masukkan kode tabel acuan (contoh: 6_06): ", "6_06", str)
    derived_table = get_input_with_default("Masukkan kode tabel turunan (contoh: 6_30): ", "6_30", str)
    column_keywords = get_column_keywords()
    return kab_code, ref_table, derived_table, column_keywords

def process_excel_data(ref_file, derived_file, output_file, kab_code, ref_table, derived_table, column_keywords):
    """Process Excel data based on user input and specified rules."""
    # Read the reference Excel file
    df_ref = pd.read_excel(ref_file, sheet_name=f"{ref_table}_kec")
    df_ref_filtered = df_ref[df_ref['kab'] == kab_code]
    
    # Get kabupaten name
    kab_name = get_kabupaten_name(df_ref, kab_code)
    
    # Update output file name with kabupaten name
    output_file = output_file.replace('.xlsx', f'_{kab_name.upper()}.xlsx')
    
    # Read the derived Excel file
    df_derived = pd.read_excel(derived_file, sheet_name=f"{derived_table}_kec")
    df_derived_filtered = df_derived[df_derived['kab'] == kab_code]
    
    # Create a new Excel writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write reference data to 'acuan' sheet
        df_ref_filtered.to_excel(writer, sheet_name='acuan', index=False)
        
        # Write derived data to 'riil' sheet
        df_derived_filtered.to_excel(writer, sheet_name='riil', index=False)
        
        # Process template sheet
        df_template = df_derived_filtered.copy()
        
        # Apply the rule: if n_rtup_usaha_ternak > 0 and < 3, set rerata to 'NA'
        for col in df_ref_filtered.columns:
            if col.startswith('n_rtup_ternak_usaha_'):
                animal = '_'.join(col.split('_')[4:])
                print(f"Processing: {format_animal_name(animal)}")
                # Find matching columns in the derived table
                matching_cols = [c for c in df_template.columns if animal in c and any(keyword in c.lower() for keyword in column_keywords)]
                
                for matching_col in matching_cols:
                    # Convert the column to string type
                    df_template[matching_col] = df_template[matching_col].astype(str)
                    
                    mask = (df_ref_filtered[col] > 0) & (df_ref_filtered[col] < 3)
                    matched_kec = df_ref_filtered.loc[mask, 'kec']
                    
                    # Set 'NA' as string
                    df_template.loc[df_template['kec'].isin(matched_kec), matching_col] = 'NA'
        
        # Write template data to 'template' sheet
        df_template.to_excel(writer, sheet_name='template', index=False)

    print(f"Processed data saved to {output_file}")

def main():
    kab_code, ref_table, derived_table, column_keywords = get_user_input()
    
    # Set base directory
    # Change this to your specific directory path
    directory = '/CheckNA/'
    input_directory = os.path.join(directory, 'data')
    output_directory = os.path.join(directory, 'output')
    

    # Find the reference and derived table files
    ref_files = glob.glob(os.path.join(input_directory, f'*{ref_table}*.xlsx'))
    derived_files = glob.glob(os.path.join(input_directory, f'*{derived_table}*.xlsx'))

    if not ref_files or not derived_files:
        print(f"Missing files in {directory}")
        print(f"Reference table files found: {ref_files}")
        print(f"Derived table files found: {derived_files}")
        return

    ref_file = ref_files[0]  # Assume the first matching file is the correct one
    derived_file = derived_files[0]  # Assume the first matching file is the correct one

    # Generate output file name
    output_file = f"PROCESSED_{os.path.basename(derived_file)}"
    output_path = os.path.join(output_directory, output_file)

    # Process the Excel files
    process_excel_data(ref_file, derived_file, output_path, kab_code, ref_table, derived_table, column_keywords)

if __name__ == "__main__":
    main()