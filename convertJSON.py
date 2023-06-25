import pandas as pd
import json
import os

def xlsx_to_json():
    # Prompt the user to enter the XLSX file
    xlsx_file = input("Enter the path to the XLSX file: ")
    
    # Read the XLSX file into a pandas DataFrame
    df = pd.read_excel(xlsx_file)
    
    # Convert each column of the DataFrame into an array, skipping empty cells
    columns = []
    for column in df.columns:
        column_data = df[column].tolist()
        filtered_data = [value for value in column_data if pd.notna(value)]
        columns.append(filtered_data)
    
    # Create the first object with file name and an empty description
    file_name = os.path.basename(xlsx_file)
    first_object = {
        "Name": file_name,
        "Description": ""
    }

    # Add the first object at the beginning of the columns list
    columns.insert(0, first_object)

    # Convert the column arrays into a JSON array with indentation
    json_data = json.dumps(columns, indent=4)

    # Save the JSON file next to the original XLSX file
    json_file = os.path.splitext(xlsx_file)[0] + '.json'
    with open(json_file, 'w') as file:
        file.write(json_data)

    print("Conversion complete. JSON file saved as:", json_file)

# Example usage
xlsx_to_json()
