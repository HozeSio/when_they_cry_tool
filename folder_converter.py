import os
import csv
import openpyxl
from text_converter import *


class FolderConverter:
    def __init__(self, folder_path):
        self.folder_path = os.path.normpath(folder_path)
        (self.folder_directory, self.folder_name) = os.path.split(self.folder_path)

    # deprecated
    def save_xlsx(self, sentences, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        for sentence in sentences:
            ws.append(sentence)
        wb.save(path)
        wb.close()

    # deprecated
    def load_xlsx(self, path):
        wb = openpyxl.open(path)
        ws = wb.active
        translation = {}
        for row in ws.rows:
            translation[row[2].value] = row[4].value
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

    def export_text(self):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_tsv')
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
                    exit(-1)
                sentences = text_converter.extract_text()

                file_name_only = os.path.splitext(file_name)[0]
                # write xlsx file
                # self.save_xlsx(sentences, os.path.join(converted_folder, file_name_only + '.xlsx'))
                # write tsv file
                self.save_tsv(sentences, os.path.join(converted_folder, f'{file_name_only}.tsv'))
                print(f"finished")

    def replace_text(self, translation_folder):
        replaced_folder = os.path.join(self.folder_directory, self.folder_name + '_replaced')
        if not os.path.exists(replaced_folder):
            os.mkdir(replaced_folder)

        translation_folder = os.path.normpath(translation_folder)
        for file_name in os.listdir(translation_folder):
            (file_name_only, ext) = os.path.splitext(file_name)
            script_file_name = f'{file_name_only}.txt'
            script_path = os.path.join(self.folder_path, script_file_name)
            if not os.path.exists(script_path):
                continue
            print(f'start replacing {script_file_name}....', end='')

            file_path = os.path.join(translation_folder, file_name)
            # new file format .tsv (actor, japanese, english, translation) and has header
            if ext == '.tsv':
                translation = self.load_tsv(file_path)
            # old file format .xlsx is (key, actor, japanese, english, translation)
            elif ext == '.xlsx':
                translation = self.load_xlsx(file_path)
            else:
                raise ModuleNotFoundError

            with open(script_path, 'r', encoding='utf-8') as f:
                text_converter = text_converter()
                replaced_text = text_converter.replace_text(f.read(), translation)
                with open(os.path.join(replaced_folder, script_file_name), 'w', encoding='utf-8') as o:
                    o.write(replaced_text)

            print('finished')