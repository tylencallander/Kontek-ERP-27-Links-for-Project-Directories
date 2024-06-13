import pandas as pd
import json

# Creating the Excel file from the projects.json file

def create_excel_from_json(json_file, output_file):
    try:
        with open(json_file, 'r') as file:
            data = json.load(file)
    except FileNotFoundError:
        print(f"\nError: JSON file at '{json_file}' not found.")
        return

    print("\nImporting projects.json file...")

    project_data = []
    for key, val in data.items():
        project_data.append({
            'Project Number': key,
            'Project Path': val['projectfullpath']
        })

    df = pd.DataFrame(project_data)

    print("\nCreating Excel sheet...")

# Setting the paths as hyperlinks

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Project Links')

        workbook  = writer.book
        worksheet = writer.sheets['Project Links']

# Manually making the paths blue and underlined so the user can actually tell its a link

        link_format = workbook.add_format({'font_color': 'blue', 'underline': 1})

        for idx, val in enumerate(project_data, start=2): 
            url = f"=HYPERLINK(\"{val['Project Path']}\", \"{val['Project Path']}\")"
            worksheet.write_url(f'B{idx}', val['Project Path'], link_format, val['Project Path'])

    print("\nExcel file successfully created!")

def main():
    json_file_path = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/ERP/3. ConRec Folder Search/V3_2024_06_12/projects.json"
    output_excel_file = "projectlinks.xlsx"
    create_excel_from_json(json_file_path, output_excel_file)

if __name__ == "__main__":
    main()