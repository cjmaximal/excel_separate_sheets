import copy
import os
import openpyxl
from datetime import date

SOURCE_PATH = './source'
OUTPUT_PATH = './out'


def _get_files() -> list:
    if os.path.exists(SOURCE_PATH):
        all_files = os.listdir(SOURCE_PATH)
        return list(filter(lambda f: f.endswith('.xlsx'), all_files))
    return []


def _save_sheet_as_new_workbook(workbook: openpyxl.Workbook, sheet_name: str, path: str):
    new_workbook = copy.deepcopy(workbook)
    sheets = workbook.sheetnames

    for sheet in sheets:
        if sheet == sheet_name:
            continue
        del new_workbook[sheet]

    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)

    file_path = '{}/{}'.format(OUTPUT_PATH, path)
    if not os.path.exists(file_path):
        os.makedirs(file_path)

    today = date.today()
    new_workbook.save('{}/{}_{}.xlsx'.format(file_path, sheet_name, today.strftime("%d_%m_%Y")))


def _separate_file(file: str):
    file_path = '{}/{}'.format(SOURCE_PATH, file)
    workbook = openpyxl.load_workbook(file_path)
    sheets = workbook.sheetnames

    file_path = file.split('.')[0]
    for sheet in sheets:
        _save_sheet_as_new_workbook(workbook, sheet, file_path)


def run():
    source_files = _get_files()
    for file in source_files:
        _separate_file(file)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    run()
