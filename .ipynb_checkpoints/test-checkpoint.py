import pandas as pd
import os
import re

def sanitize_filename(filename):
    """
    Convert a string to a safe filename by removing or replacing invalid characters.
    """
    # Remove or replace invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    # Replace spaces and colons with underscores
    filename = filename.replace(' ', '').replace(':', '')
    # Remove any multiple underscores
    filename = re.sub(r'+', '', filename)
    # Remove leading/trailing underscores
    filename = filename.strip('_')
    return filename.lower()

def convert_excel_to_formatted_text(excel_file_path, output_directory):
    """
    Converts Excel data to formatted text files with entries in the format:
    {title: 'Name - StallNumber', description: 'Description'}
    
    Parameters:
    excel_file_path (str): Path to the Excel file
    output_directory (str): Directory where category files will be saved
    """
    try:
        # Create output directory if it doesn't exist
        os.makedirs(output_directory, exist_ok=True)
        
        # Read the Excel file
        df = pd.read_excel(excel_file_path)
        
        # Group by category
        grouped_data = df.groupby('Category')
        
        # Process each category
        for category, group in grouped_data:
            # Create safe filename for this category
            safe_filename = sanitize_filename(category)
            file_path = os.path.join(output_directory, f"{safe_filename}.txt")
            
            # Open file for writing
            with open(file_path, 'w', encoding='utf-8') as f:
                entries = []
                
                # Format each entry
                for _, row in group.iterrows():
                    # Combine stall name and number in the title
                    combined_title = f"{row['Title']} - {row['Stall_No']}"
                    entry = f"{{title: '{combined_title}', description: '{row['Description']}'}}"
                    entries.append(entry)
                
                # Join all entries with comma and space
                formatted_text = ", ".join(entries)
                f.write(formatted_text)
        
        print(f"Files have been created in {output_directory}")
        return True
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return False

# Example usage
#if _name_ == "_main_":
    # Replace with your Excel file path
    excel_file = "Order.xlsx"
    
    # Specify output directory
    output_dir = "category_files"
    
    # Convert and save files
    convert_excel_to_formatted_text(excel_file, output_dir)