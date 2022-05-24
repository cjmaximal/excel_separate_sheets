import argparse
import copy
import os
import openpyxl
from datetime import date
from rich.console import Console

SOURCE_PATH = './source'
OUTPUT_PATH = './out'


def _get_files() -> list:
    if os.path.exists(SOURCE_PATH):
        all_files = os.listdir(SOURCE_PATH)
        return list(filter(lambda f: f.endswith('.xlsx'), all_files))
    return []


def _save_sheet_as_new_workbook(workbook: openpyxl.Workbook, workbook_name: str, sheet_name: str,
                                is_separately: bool = False):
    new_workbook = copy.deepcopy(workbook)
    sheets = workbook.sheetnames

    for sheet in sheets:
        if sheet == sheet_name:
            continue
        del new_workbook[sheet]

    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)

    today = date.today()
    file_path = '{}/{}'.format(OUTPUT_PATH, workbook_name)
    if is_separately:
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        path_to_save = '{}/{}_{}.xlsx'.format(file_path, sheet_name, today.strftime("%d_%m_%Y"))
    else:
        path_to_save = '{}_{}_{}.xlsx'.format(file_path, sheet_name, today.strftime("%d_%m_%Y"))
    new_workbook.save(path_to_save)


def _separate_file(file: str, is_separately: bool = False):
    file_path = '{}/{}'.format(SOURCE_PATH, file)
    workbook = openpyxl.load_workbook(file_path)
    sheets = workbook.sheetnames

    workbook_name = file.split('.')[0]
    for sheet in sheets:
        _save_sheet_as_new_workbook(workbook, workbook_name, sheet, is_separately)


def run():
    parser = argparse.ArgumentParser(description='Separate workbook sheets as self workbooks.')
    parser.add_argument('--by-folders', dest='is_separately', action='store_const',
                        const=True, default=False,
                        help='Out workbooks separately by folders (default: place all results to out directory)')
    args = parser.parse_args()

    console = Console()

    source_files = _get_files()
    count_files = len(source_files)
    console.log(f"[bold yellow]Files to process: {count_files}")

    with console.status("[bold green]Fetching data...") as status:
        counter = 0
        for file in source_files:
            counter += 1
            console.log(f"[green]Processing file[/green] {counter}: {file}")
            _separate_file(file, args.is_separately)
        console.log(f'[bold][red]Done!')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    run()
