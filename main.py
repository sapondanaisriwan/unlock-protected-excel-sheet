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
def modify_worksheets(output_file, zip_file_path, file_path, regex_pattern):
    # Open the output file in write mode
    with zipfile.ZipFile(output_file, mode="w") as new_zip_file:
        # Open the input zip file in read mode
        with zipfile.ZipFile(zip_file_path, mode="r") as zip_file:
            # Loop over each file in the input zip file
            for file_info in zip_file.infolist():
                # If the file is a worksheet xml file, remove sheet protection elements
                if file_info.filename.startswith(file_path) and file_info.filename.endswith(".xml"):
                    file_contents = zip_file.read(file_info).decode()
                    new_contents = regex_pattern.sub(" ", file_contents)
                    new_zip_file.writestr(file_info, new_contents)
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
    original_file = "./input.xlsx"  # Input Excel file path
    output_file = "output.xlsx"  # Output Excel file path
    zip_file_path = "./input.zip"  # Backup zip file path
    file_path = "xl/worksheets/"  # Path to worksheet xml files in the zip file
    regex_pattern = re.compile(r"<sheetProtection(.+?)/>")  # Regular expression pattern to remove sheet protection

    print("Checking...")
    # Check if the input file exists
    if check_input_file(original_file):
        return print("The file does not exist. Please rename to input.xlsx")

    # Check if the backup zip file exists and create it if it doesn't
    check_zip_file(original_file, zip_file_path)
    print("Create input.zip")

    # Modify the worksheets in the input zip file
    modify_worksheets(output_file, zip_file_path, file_path, regex_pattern)
    print("Create output.xlsx")

    # Remove the backup zip file
    remove_file(zip_file_path)

    print("Done")


# Call the main function if this script is being run directly
if __name__ == '__main__':
    main()
