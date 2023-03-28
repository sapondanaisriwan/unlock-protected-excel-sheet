import re
import shutil
import os
import zipfile


def duplicate_file(original_file, output_file):
    shutil.copyfile(original_file, output_file)


def file_exists(file_path):
    return os.path.isfile(file_path)


def check_input_file(original_file):
    if not file_exists(original_file):
        return print("The file does not exist. Please rename to input.xlsx")


def check_zip_file(original_file, zip_file_path):
    if not file_exists(zip_file_path):
        duplicate_file(original_file, zip_file_path)
        return print("Create input.zip")


def modify_worksheets(output_file, zip_file_path, file_path, regex_pattern):
    with zipfile.ZipFile(output_file, mode="w") as new_zip_file:
        with zipfile.ZipFile(zip_file_path, mode="r") as zip_file:
            for file_info in zip_file.infolist():
                if file_info.filename.startswith(file_path) and file_info.filename.endswith(".xml"):
                    file_contents = zip_file.read(file_info).decode()
                    new_contents = regex_pattern.sub(" ", file_contents)
                    new_zip_file.writestr(file_info, new_contents)
                else:
                    new_zip_file.writestr(
                        file_info.filename, zip_file.read(file_info))

    return output_file


def remove_file(file_path):
    if file_exists(file_path):
        os.remove(file_path)
        return print("Remove zip file")


def main():
    original_file = "./input.xlsx"
    output_file = "output.xlsx"
    zip_file_path = "./input.zip"
    file_path = "xl/worksheets/"
    regex_pattern = re.compile(r"<sheetProtection(.+?)/>")

    print("Checking...")
    check_input_file(original_file)
    check_zip_file(original_file, zip_file_path)

    print("Create output.xlsx")
    modify_worksheets(output_file,
                      zip_file_path, file_path, regex_pattern)

    remove_file(zip_file_path)
    print("Done")


if __name__ == '__main__':
    main()
