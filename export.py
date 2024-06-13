import pandas as pd
import json

# Creating the Excel file

def create_excel_from_json(json_file, output_file):
    try:
        with open(json_file, 'r') as file:
            data = json.load(file)
    except FileNotFoundError:
        print(f"\nError: JSON file, in path '{json_file}' not found.")
        return

    print("\nImporting projects.json file...")

# Making all of the folder paths into hyperlinks.

    project_data = [{'Project Number': key, 'Project Path': f"=HYPERLINK(\"{val['projectfullpath']}\", \"{val['projectfullpath']}\")"} for key, val in data.items()]
    df = pd.DataFrame(project_data)

    print("\nGenerating Excel sheet...")

# Making the folder paths blue so people can actually tell its a hyperlink

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Project Links')

        workbook = writer.book
        worksheet = writer.sheets['Project Links']

        link_format = workbook.add_format({'font_color': 'blue', 'underline': 1})

        worksheet.set_column('B:B', None, link_format)

    print("\nExcel file created successfully with styled hyperlinks.")

def main():
    json_file_path = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/ERP/3. ConRec Folder Search/V3_2024_06_12/projects.json"
    output_excel_file = "projectlinks.xlsx"
    create_excel_from_json(json_file_path, output_excel_file)

if __name__ == "__main__":
    main()
