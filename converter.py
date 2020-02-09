#!/usr/bin/env python3
import sys
import translation_extractor
from text_converter import *
from folder_converter import *


def validate_folder(folder_path: str):
    result = True
    for file_name in os.listdir(folder_path):
        if not file_name.endswith('.txt'):
            continue

        print(f"validating {file_name}")
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            text_converter = TextConverter(f.read())
            result &= text_converter.validate_text()

    if not result:
        exit(-1)


# deprecated
def combine_xlsx(original_folder, translated_folder):
    for file_name in os.listdir(translated_folder):
        if not file_name.endswith('.xlsx'):
            continue
        file_name = file_name.replace('kor', '')

        original_path = os.path.join(original_folder, file_name)
        if not os.path.exists(original_path):
            continue

        original_wb = openpyxl.open(original_path)
        original_ws = original_wb.active

        translated_wb = openpyxl.open(os.path.join(translated_folder, file_name))
        translated_ws = translated_wb.active

        for index, row in enumerate(translated_ws.iter_rows(), 1):
            original_ws.cell(row=index, column=4).value = row[2].value

        original_wb.save(original_path)
        original_wb.close()


# deprecated
def insert_actor_column(old_folder, actor_folder):
    for file_name in os.listdir(old_folder):
        if not file_name.endswith('.xlsx'):
            continue

        old_path = os.path.join(old_folder, file_name)
        old_wb = openpyxl.open(old_path)
        old_ws = old_wb.active

        actor_wb = openpyxl.open(os.path.join(actor_folder, file_name))
        actor_ws = actor_wb.active

        for index, row in enumerate(actor_ws.iter_rows(), 1):
            if old_ws.cell(row=index, column=2).value != row[2].value:
                print(f"{file_name} has different row at {index} {old_ws.cell(row=index, column=2).value} != {row[2].value}")
                break

        old_ws.insert_cols(2)

        for index, row in enumerate(actor_ws.iter_rows(), 1):
            old_ws.cell(row=index, column=2).value = row[1].value

        old_wb.save(old_path)
        old_wb.close()


if __name__ == '__main__':
    if len(sys.argv) == 1 or sys.argv[1] == 'help':
        print(
"""
usage: converter.py [commands]
available commands:
    export_text <Update folder>
    - export text parameter to xlsx file from the script
    replace_text <Update folder> <translation folder>
    - Replace english text to translated text
    extract_text <file_path>
    - extract text line from the onscript file and export to xlsx
    combine_xlsx <original_folder> <translated_folder>
    insert_actor_column <old_folder> <actor_folder>
"""
        )
    elif sys.argv[1] == 'export_text':
        converter = FolderConverter(sys.argv[2])
        converter.export_text()
    elif sys.argv[1] == 'replace_text':
        converter = FolderConverter(sys.argv[2])
        converter.replace_text(sys.argv[3])
    elif sys.argv[1] == 'validate_folder':
        validate_folder(sys.argv[2])
    elif sys.argv[1] == 'extract_text':
        extractor = translation_extractor.TextExtractor()
        extractor.extract_text(sys.argv[2])
    elif sys.argv[1] == 'combine_xlsx':
        combine_xlsx(sys.argv[2], sys.argv[3])
    elif sys.argv[1] == 'insert_actor_column':
        insert_actor_column(sys.argv[2], sys.argv[3])
    else:
        print("invalid command", file=sys.stderr)
        exit(-1)
