import re
import shutil
import os
import zipfile

# Duplicate the input Excel file to create a backup zip file
def duplicate_excel_file(original_file, output_file):
    shutil.copyfile(original_file, output_file)

# Check if a file exists at the given path
def check_exits_file(file_path):
    return not os.path.isfile(file_path)

# Check if the input Excel file exists
def check_input_file(original_file):
    return check_exits_file(original_file)

# Check if the backup zip file exists and create it if it doesn't
def check_zip_file(original_file, zip_file_path):
    if check_exits_file(zip_file_path):
        return duplicate_excel_file(original_file, zip_file_path)

# Modify the worksheets in the input zip file
def modify_worksheets(output_file, zip_file_path, worksheet_path, workbook_path, regex_patterns):
    # Open the output file in write mode
    with zipfile.ZipFile(output_file, mode="w") as new_zip_file:
        # Open the input zip file in read mode
        with zipfile.ZipFile(zip_file_path, mode="r") as zip_file:
            # Loop over each file in the input zip file
            for file_info in zip_file.infolist():
                # If the file is a worksheet xml file or the workbook xml file, process it
                if (file_info.filename.startswith(worksheet_path) and file_info.filename.endswith(".xml")) or file_info.filename == workbook_path:
                    file_contents = zip_file.read(file_info)
                    try:
                        file_contents = file_contents.decode()
                        # If the file is a worksheet xml file, remove sheet protection elements
                        if file_info.filename.startswith(worksheet_path) and file_info.filename.endswith(".xml"):
                            new_contents = regex_patterns['sheet_protection'].sub(" ", file_contents)
                            new_zip_file.writestr(file_info, new_contents)
                        # If the file is the workbook xml file, modify it accordingly
                        elif file_info.filename == workbook_path:
                            new_contents = regex_patterns['workbook_protection'].sub("", file_contents)
                            new_contents = regex_patterns['workbook_properties'].sub("<workbookPr/>", new_contents)
                            new_zip_file.writestr(file_info, new_contents)
                    except UnicodeDecodeError:
                        # If decoding fails, just copy the file as-is
                        new_zip_file.writestr(file_info.filename, file_contents)
                # Otherwise, copy the file to the output zip file as-is
                else:
                    new_zip_file.writestr(file_info.filename, zip_file.read(file_info))

    # Return the output file path
    return output_file

# Remove a file if it exists
def remove_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        print("Remove zip file")

# Main function
def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    original_file = os.path.join(script_dir, "input.xlsx")  # Input Excel file path
    output_file = os.path.join(script_dir, "output.xlsx")  # Output Excel file path
    zip_file_path = os.path.join(script_dir, "input.zip")  # Backup zip file path
    worksheet_path = "xl/worksheets/"  # Path to worksheet xml files in the zip file
    workbook_path = "xl/workbook.xml"  # Path to workbook xml file in the zip file

    regex_patterns = {
        'sheet_protection': re.compile(r"<sheetProtection(.+?)/>"),  # Regular expression pattern to remove sheet protection
        'workbook_protection': re.compile(r"<workbookProtection(.+?)/>"),  # Regular expression pattern to remove workbook protection
        'workbook_properties': re.compile(r"<workbookPr(.+?)/>")  # Regular expression pattern to modify workbook properties
    }

    print("Checking...")
    # Check if the input file exists
    if check_input_file(original_file):
        return print("The file does not exist. Please rename to input.xlsx")

    # Check if the backup zip file exists and create it if it doesn't
    check_zip_file(original_file, zip_file_path)
    print("Create input.zip")

    # Modify the worksheets in the input zip file
    modify_worksheets(output_file, zip_file_path, worksheet_path, workbook_path, regex_patterns)
    print("Create output.xlsx")

    # Remove the backup zip file
    remove_file(zip_file_path)

    print("Done")

# Call the main function if this script is being run directly
if __name__ == '__main__':
    main()
