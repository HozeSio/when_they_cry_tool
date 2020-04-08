import re
import os
import openpyxl

text_pattern = re.compile(r"""
^.*[@\\/][ ]*$
""", re.VERBOSE | re.MULTILINE)


class OnscriptParser:
    HEADER = ('korean',)

    def __init__(self, text):
        self.text = text

    def parse_text(self):
        sentences_jp = []
        sentences_kr = []

        for match in text_pattern.finditer(self.text):
            if match.group().startswith(';'):
                sentences_jp.append(match.group().lstrip(';'))
            else:
                sentences_kr.append(match.group())
        return sentences_kr


class FolderParser:
    def __init__(self, folder_path):
        self.folder_path = os.path.normpath(folder_path)
        (self.folder_directory, self.folder_name) = os.path.split(self.folder_path)

    def export_text(self):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_output')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)

        for file_name in os.listdir(self.folder_path):
            if not file_name.endswith('.txt'):
                continue

            file_path = os.path.join(self.folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                print(f"start converting {file_name}....", end='')
                text_converter = OnscriptParser(f.read())
                sentences = text_converter.parse_text()

                file_name_only = os.path.splitext(file_name)[0]
                self.save_xlsx(sentences, os.path.join(converted_folder, f'{file_name_only}.xlsx'))
                print(f"finished")

    def save_xlsx(self, sentences, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(OnscriptParser.HEADER)
        for sentence in sentences:
            ws.append((sentence,))
        wb.save(path)
        wb.close()
