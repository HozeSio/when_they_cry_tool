import os
import csv
import openpyxl
from text_converter import *

HEADER_ROW = ('actor', 'japanese', 'english', 'korean', 'papago', 'comments')


def has_header(worksheet: openpyxl.workbook.workbook.Worksheet):
    return worksheet.cell(1, 1).value == 'actor' and worksheet.cell(1, 2).value == 'japanese'


def ignore_row(first_cell):
    value = first_cell.value
    if value and (value.startswith(script_method) or
                  value.startswith('void') or
                  value == 'actor'):
        return True
    return False


class FolderConverter:
    def __init__(self, folder_path):
        self.folder_path = os.path.normpath(folder_path)
        (self.folder_directory, self.folder_name) = os.path.split(self.folder_path)

    def save_xlsx(self, sentences, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        for sentence in sentences:
            ws.append(sentence)
        wb.save(path)
        wb.close()

    def load_xlsx(self, path, key_col, value_col, prefix):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        translation = {}
        index = 0
        for row in ws.rows:
            if prefix:
                if ignore_row(row[0]):
                    continue

                key = f"{index}_{str(row[key_col].value)}"
            else:
                key = str(row[key_col].value)

            if key != 'None' and key in translation:
                print(f'key duplication {key}')
            translation[key] = str(row[value_col].value)

            index += 1
        wb.close()
        return translation

    def save_tsv(self, sentences, path):
        with open(path, 'w', encoding='utf-8', newline='') as f:
            wr = csv.writer(f, delimiter='\t')
            for sentence in sentences:
                wr.writerow(sentence)

    def load_tsv(self, path):
        with open('test.tsv', 'r', encoding='utf-8') as f:
            rdr = csv.reader(f, delimiter='\t')
            translation = {}
            for row in rdr[1:]:
                translation[row[1]] = row[3]
        return translation

    def export_text(self, format):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_output')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)

        for file_name in os.listdir(self.folder_path):
            if not file_name.endswith('.txt'):
                continue

            file_path = os.path.join(self.folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                print(f"start converting {file_name}....", end='')
                text_converter = TextConverter(f.read())
                if not text_converter.validate_text():
                    sys.exit(-1)
                sentences = text_converter.extract_text()

                file_name_only = os.path.splitext(file_name)[0]
                if format == 'xlsx':
                    self.save_xlsx(sentences, os.path.join(converted_folder, f'{file_name_only}.xlsx'))
                elif format == 'tsv':
                    self.save_tsv(sentences, os.path.join(converted_folder, f'{file_name_only}.tsv'))
                print(f"finished")

    def replace_text(self, translation_folder, actor_path):
        replaced_folder = os.path.join(self.folder_directory, self.folder_name + '_replaced')
        if not os.path.exists(replaced_folder):
            os.mkdir(replaced_folder)

        translation_base = self.load_xlsx(actor_path, 1, 2, False)
        translation_base[None] = ''

        translation_folder = os.path.normpath(translation_folder)
        for file_name in os.listdir(translation_folder):
            (file_name_only, ext) = os.path.splitext(file_name)
            script_file_name = f'{file_name_only}.txt'
            script_path = os.path.join(self.folder_path, script_file_name)
            if not os.path.exists(script_path):
                continue
            print(f'start replacing {script_file_name}....', end='')
            translation = dict(translation_base)

            file_path = os.path.join(translation_folder, file_name)
            # deprecated
            if ext == '.tsv':
                translation.update(self.load_tsv(file_path))
            elif ext == '.xlsx':
                translation.update(self.load_xlsx(file_path, 1, 3, True))
            else:
                raise ModuleNotFoundError

            with open(script_path, 'r', encoding='utf-8') as f:
                text_converter = TextConverter(f.read())
                replaced_text = text_converter.replace_text(translation)
                if not TextConverter(replaced_text).validate_text():
                    sys.exit(-1)
                with open(os.path.join(replaced_folder, script_file_name), 'w', encoding='utf-8') as o:
                    o.write(replaced_text)

            print('finished')
