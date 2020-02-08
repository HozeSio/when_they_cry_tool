#!/usr/bin/env python
import sys
import os
import re
import openpyxl
import translation_extractor

method_pattern = re.compile(r"""
(                       # method front part [0]
\s*OutputLine
\s*\(
\s*
)
([^,]*)                    # first parameter (actor) [1]
(\s*,\s*)              # comma [2]
(.*)                  # second parameter (text) [3]
([ \t\x0B\f\r]*,\n?[ \t\x0B\f\r]*) # comma [4]
([^,\n]*)                    # third parameter (actor-alt) [5]
(\s*,\s*)               # comma [6]
(.*)                  # fourth parameter (text-alt) [7]
(\s*,\s*)               # comma [8]
(.*)                    # fifth parameter (show option) [9]
(                       # last part [10]
\s*\)
\s*;
)
""", re.VERBOSE | re.MULTILINE)


def validate_folder(folder_path: str):
    result = True
    for file_name in os.listdir(folder_path):
        if not file_name.endswith('.txt'):
            continue

        print(f"validating {file_name}")
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            result &= validate_text(f.read())

    if not result:
        exit(-1)


def validate_text(text: str):
    for method_match in method_pattern.finditer(text):
        groups = method_match.groups()
        # output text
        if groups[1] == 'NULL':
            if not groups[3].startswith('\"') or not groups[3].endswith('\"') or not groups[7].startswith('\"') or not groups[7].endswith('\"'):
                print(method_match.group(), file=sys.stderr)
                return False
        # set actor
        else:
            pass
    return True


class TextConverter:
    def __init__(self, file_name):
        self.match_count = 0
        self.file_name = file_name
        self.sentences = {}
        self.last_actor = None
        self.actor_pattern = re.compile(r'"<color=.*>(.*)</color>"')
        self.remove_quotation_mark_pattern = re.compile(r'"(.*)"')

    def strip_quotation_mark(self, text: str):
        m = self.remove_quotation_mark_pattern.match(text)
        if not m:
            print(text)
            raise Exception
        return m.groups()[0]

    def repl_text_to_key(self, match_obj) -> str:
        groups = list(match_obj.groups())

        count = str(self.match_count * 10).zfill(6)
        key = f"{self.file_name}_{count}"
        param1 = groups[1]
        param2 = groups[3]
        param3 = groups[5]
        param4 = groups[7]
        param5 = groups[9]

        if param1 != 'NULL':  # actor setting
            self.last_actor = self.actor_pattern.match(param3).groups()[0]
            return match_obj.group()

        if param2.startswith('\"<size=') or param4.startswith('\"<size='):
            return match_obj.group()

        # store sentence
        self.sentences[key] = (self.last_actor, self.strip_quotation_mark(param2), self.strip_quotation_mark(param4))
        self.last_actor = None

        # replace text to key
        groups[3] = key
        groups[7] = key
        self.match_count += 1
        return "".join(groups)

    def text_to_key(self, text):
        return method_pattern.sub(self.repl_text_to_key, text)


class FolderConverter:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        (self.folder_directory, self.folder_name) = os.path.split(folder_path)

    def text_to_key(self):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_converted')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)
        text_folder = os.path.join(converted_folder, 'txt')
        if not os.path.exists(text_folder):
            os.mkdir(text_folder)
        xlsx_folder = os.path.join(converted_folder, 'xlsx')
        if not os.path.exists(xlsx_folder):
            os.mkdir(xlsx_folder)

        for file_name in os.listdir(self.folder_path):
            if not file_name.endswith('.txt'):
                continue

            file_path = os.path.join(self.folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                print(f"start converting {file_name}....", end='')
                text = f.read()
                if not validate_text(text):
                    exit(-1)
                file_name_only = os.path.splitext(file_name)[0]
                text_converter = TextConverter(file_name_only)
                converted_text = text_converter.text_to_key(text)

                # write text to key converted text file
                new_path = os.path.join(text_folder, file_name)
                with open(new_path, 'w', encoding='utf-8') as w:
                    w.write(converted_text)

                # write xlsx content
                wb = openpyxl.Workbook()
                # wb.remove(wb.active)
                # ws = wb.create_sheet(file_name)
                ws = wb.active
                ws.append(['actor', 'japanese', 'english', 'translation'])
                for key, item in text_converter.sentences.items():
                    ws.append([item[0], item[1], item[2]])
                wb.save(os.path.join(xlsx_folder, file_name_only + '.xlsx'))
                wb.close()
                print(f"now converted to {file_name_only}.xlsx")


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
    -  export text parameter to xlsx file from the 
    extract_text <file_path>
    - extract text line from the onscript file and export to xlsx
    combine_xlsx <original_folder> <translated_folder>
    insert_actor_column <old_folder> <actor_folder>
"""
        )
    elif sys.argv[1] == 'export_text':
        converter = FolderConverter(sys.argv[2])
        converter.text_to_key()
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
