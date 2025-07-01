import pandas as pd
import os
import re
import zipfile
import tempfile
import shutil

def clean_sheet_name(name):
    """
    Sanitize sheet name by removing or replacing invalid characters
    and ensuring it meets Excel's requirements.
    """
    # Convert to string and truncate to 31 characters
    name = str(name)[:31]
    # Remove invalid characters: / \ * ? : [ ]
    name = re.sub(r'[\/\\*?:\[\]]', '_', name)
    # If name is empty after cleaning, provide a default name
    return name if name.strip() else 'Sheet'

def split_and_group_excel(input_file, group_column_index=1):
    """
    Group data by specified column, create temporary Excel files for each group,
    and produce a single zip file containing all grouped Excel files.
    
    Parameters:
    input_file (str): Path to input Excel file
    group_column_index (int): Index of the column to group by (default is 1 for Column B)
    """
    # Get the directory of the input file
    input_dir = os.path.dirname(input_file) or '.'
    
    # Define output directory as SEPERATED_DATA in the same directory as input file
    output_dir = os.path.join(input_dir, "SEPERATED_DATA")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # List to store paths of temporary Excel files
    created_files = []
    
    try:
        # Step 1: Read and group data by specified column
        df = pd.read_excel(input_file)
        header = df.columns.tolist()
        
        # Handle blanks in the specified column
        df[df.columns[group_column_index]] = df[df.columns[group_column_index]].fillna("Blank")
        
        # Step 2: Create temporary Excel files for each group
        for group_value, group in df.groupby(df.columns[group_column_index]):
            # Use cleaned sheet name
            sheet_name = clean_sheet_name(group_value)
            # Create a temporary file for this group
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False, dir=output_dir) as temp_file:
                temp_file_path = temp_file.name
                # Export the group to a temporary Excel file
                group.to_excel(temp_file_path, index=False, sheet_name=sheet_name)
                created_files.append((temp_file_path, sheet_name))
                print(f"Created temporary file: {temp_file_path}")
        
        # Step 3: Create a zip file containing all temporary Excel files
        zip_file_path = os.path.join(input_dir, "Separated_Data_Archive.zip")
        with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path, sheet_name in created_files:
                # Sanitize sheet name for filename
                safe_sheet_name = clean_sheet_name(sheet_name)
                # Use the sheet name as the filename in the zip
                arcname = os.path.join("SEPERATED_DATA", f"{safe_sheet_name}.xlsx")
                zipf.write(file_path, arcname)
        print(f"Created zip archive: {zip_file_path}")
    
    finally:
        # Step 4: Clean up temporary files and SEPERATED_DATA folder
        for file_path, _ in created_files:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Cleaned up temporary file: {file_path}")
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
            print(f"Cleaned up temporary folder: {output_dir}")
    
    print("Successfully created Separated_Data_Archive.zip containing all grouped Excel files.")

# Example usage
if __name__ == "__main__":
    # Specify input Excel file
    input_excel = "tea_data.xlsx"  # Replace with your input Excel file path
    
    try:
        split_and_group_excel(input_excel, group_column_index=1)
    except Exception as e:
        print(f"An error occurred: {str(e)}")
