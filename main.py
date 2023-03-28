import re
import shutil
import os
import zipfile


def duplicate_excel_file(original_file, output_file):
    shutil.copyfile(original_file, output_file)


def check_zip_file(original_file, zip_file_path):
    if not os.path.exists(zip_file_path):
        print("File does not exist. Creating a new file.")
        duplicate_excel_file(original_file, zip_file_path)

    with zipfile.ZipFile(zip_file_path, 'r') as archive:
        if not archive.namelist():
            print("Zip file is empty. Creating a new file.")
            duplicate_excel_file(original_file, zip_file_path)


def remove_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)


def rename_file(file_path):
    if os.path.exists(file_path):
        file_name, file_extension = os.path.splitext(file_path)
        os.rename(file_path, file_name + ".xlsx")


def modify_worksheets(zip_file_path, file_path, regex_pattern):
    new_zip_file_path = "output.zip"
    with zipfile.ZipFile(new_zip_file_path, mode="w") as new_zip_file:
        with zipfile.ZipFile(zip_file_path, mode="r") as zip_file:
            for file_info in zip_file.infolist():
                if file_info.filename.startswith(file_path) and file_info.filename.endswith(".xml"):
                    file_contents = zip_file.read(file_info).decode()
                    new_contents = regex_pattern.sub(" ", file_contents)
                    new_zip_file.writestr(file_info, new_contents)
                else:
                    new_zip_file.writestr(
                        file_info.filename, zip_file.read(file_info))

    return new_zip_file_path


def main():
    original_file = "./input.xlsx"
    zip_file_path = "./input.zip"
    file_path = "xl/worksheets/"
    regex_pattern = re.compile(r"<sheetProtection(.+?)/>")

    check_zip_file(original_file, zip_file_path)
    output_file_path = modify_worksheets(
        zip_file_path, file_path, regex_pattern)
    rename_file(output_file_path)
    remove_file(zip_file_path)


if __name__ == '__main__':
    main()
