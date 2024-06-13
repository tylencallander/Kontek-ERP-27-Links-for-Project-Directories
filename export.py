import json
import pandas as pd

# Creating Excel file & exporting projects.json file

def create_excel_from_json(json_file, output_file):
    try:
        with open(json_file, 'r') as file:
            data = json.load(file)
    except FileNotFoundError:
        print(f"\nError: JSON file, in path '{json_file}' not found.")
        return
    
    print("\nImporting projects.json file...")
    
    project_data = []
    for project_number, details in data.items():
        project_data.append([project_number, details['projectfullpath']])
    print("\nGenerating Excel sheet...")
    
    df = pd.DataFrame(project_data, columns=['Project Number', 'Project Path'])
    
    df.to_excel(output_file, index=False)
    print("\nExcel file created successfully.")

def main():
    json_file_path = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/ERP/3. ConRec Folder Search/V3_2024_06_12/projects.json"
    output_excel_file = "projectlinks.xlsx"
    create_excel_from_json(json_file_path, output_excel_file)

if __name__ == "__main__":
    main()