import pandas as pd
import json
from datetime import datetime, timedelta
import re

def clean_text(text):
    """Clean text by removing asterisks, extra whitespace, and normalizing formatting."""
    if pd.isna(text):
        return None
    
    text = str(text).strip()
    
    # Skip empty strings, asterisks, and day names
    if (not text or 
        '*' in text or 
        text.upper() in ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'] or
        text.upper() in ['BREAKFAST', 'LUNCH', 'DINNER']):
        return None
    
    # Clean up extra spaces and normalize formatting
    # Remove multiple spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Clean up parentheses formatting
    text = re.sub(r'\s*\(\s*', ' (', text)
    text = re.sub(r'\s*\)\s*', ')', text)
    
    # Remove any trailing brackets
    text = text.rstrip(')') + ')' if text.count('(') > text.count(')') else text
    
    # Clean up plus signs
    text = re.sub(r'\s*\+\s*', ' + ', text)
    
    # Clean up slashes
    text = re.sub(r'\s*\/\s*', ' / ', text)
    
    return text.strip()

def parse_excel_to_json(excel_path):
    """Parse excel file and convert to desired JSON format."""
    
    df = pd.read_excel(excel_path)
    
    # Initialize result dictionary
    result = {}
    
    # Finding indexes....
    section_ranges = {}
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]):
            value = str(row.iloc[0]).strip().upper()
            if 'BREAKFAST' in value:
                section_ranges['BREAKFAST'] = idx
            elif 'LUNCH' in value:
                section_ranges['LUNCH'] = idx
            elif 'DINNER' in value:
                section_ranges['DINNER'] = idx
    
    # creating dates list
    start_date = datetime(2025, 2, 1)  # Saturday
    dates = []
    for i in range(len(df.columns[1:])):  # Skip the first column
        dates.append((start_date + timedelta(days=i)).strftime('%Y-%m-%d'))
    
    # The "mega loop" for iterating through the excel
    for col_idx, col_name in enumerate(df.columns[1:], 0):  
        date = dates[col_idx]
        result[date] = {
            "BREAKFAST": [],
            "LUNCH": [],
            "DINNER": []
        }
        
        # Add breakfast items to the result 
        breakfast_start = section_ranges['BREAKFAST'] + 1 # to start with the items
        lunch_start = section_ranges['LUNCH']
        for idx in range(breakfast_start, lunch_start):
            item = clean_text(df.iloc[idx][col_name])
            if item:
                result[date]["BREAKFAST"].append(item)
        
        # Add lunch items to the result 
        lunch_start = section_ranges['LUNCH'] + 1
        dinner_start = section_ranges['DINNER']
        for idx in range(lunch_start, dinner_start):
            item = clean_text(df.iloc[idx][col_name])
            if item:
                result[date]["LUNCH"].append(item)
        
        # Add dinner items to the result 
        dinner_start = section_ranges['DINNER'] + 1
        for idx in range(dinner_start, len(df)):
            item = clean_text(df.iloc[idx][col_name])
            if item and item.upper() not in ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']:
                result[date]["DINNER"].append(item)
    
    return result

def save_to_json(data, output_path):
    """Save data to JSON file with proper formatting."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def main():
    try:
       
        excel_path = r"C:\Users\Karan Singh\Downloads\mess menu.xlsx"
        menu_data = parse_excel_to_json(excel_path)
        
       
        print("Number of days processed:", len(menu_data))
        first_date = list(menu_data.keys())[0]
        print("\nSample of first day's menu:", json.dumps(menu_data[first_date], indent=2))
        
       
        output_path = "mess_menu.json"
        save_to_json(menu_data, output_path)
        print(f"\nSuccessfully converted mess menu to JSON. Saved at: {output_path}")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
