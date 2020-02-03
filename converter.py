#!/usr/bin/env python
import sys
import os
import re
import openpyxl


def validate_text(file_path: str):
    with open(file_path, 'r', encoding='utf-8') as f:
        method_pattern = re.compile(r"""
        ^
        \s*OutputLine       # function name
        \s*\(
        \s*(.*\n.*)         # parameters
        \s*\)
        \s*;
        """, re.VERBOSE | re.MULTILINE)

        text_parameter_validate_pattern = re.compile(r"""
        \s*NULL
        \s*,     
        \s*
        (".*")          # first text parameter (japanese)
        \s*,
        \s*NULL
        \s*,
        \s*
        (".*")          # second text parameter (english)
        \s*,
        \s*
        (.*)            # text output parameter
        """, re.VERBOSE | re.DOTALL)

        text = f.read()
        for index, method_match in enumerate(method_pattern.finditer(text)):
            method_parameters = method_match.groups()[0]

            validate_match = text_parameter_validate_pattern.sub(method_parameters)
            if validate_match is None:
                print(method_match.group(), file=sys.stderr)
            else:
                (first_text, second_text, option) = validate_match.groups


class TextConverter:
    def __init__(self, file_name):
        self.match_count = 0
        self.file_name = file_name
        self.sentences = {}

    def repl_text_to_key(self, match_obj) -> str:
        groups = list(match_obj.groups())
        count = str(self.match_count * 10).zfill(6)
        key = f"{self.file_name}_{count}"
        # store sentence
        self.sentences[key] = (groups[1], groups[3])

        # replace text to key
        groups[1] = key
        groups[3] = key
        self.match_count += 1
        return "".join(groups)

    def text_to_key(self, text):
        replace_pattern = re.compile(r"""
        ^
        (                       # method front part /1
        \s*OutputLine
        \s*\(
        \s*NULL\s*,\s*
        )
        (".*")                  # first text part /2
        (                       # part between first text and second text /3
        \s*,\s*NULL\s*,\s*
        )
        (".*")                  # second text part /4
        (                       # part between second text and show option /5
        \s*,\s*
        )
        (.*)                    # show option /6
        (                       # last part /7
        \s*\)
        \s*;
        )
        """, re.VERBOSE | re.MULTILINE)

        return replace_pattern.sub(self.repl_text_to_key, text)


class FolderConverter:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        (self.folder_directory, self.folder_name) = os.path.split(folder_path)

        self.remove_quotation_mark_pattern = re.compile(r'"(.*)"')

    def strip_quotation_mark(self, text: str):
        return self.remove_quotation_mark_pattern.match(text).groups()[0]

    def text_to_key(self):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_key')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)

        for file_name in os.listdir(self.folder_path):
            with open(os.path.join(self.folder_path, file_name), 'r', encoding='utf-8') as f:
                file_name_only = os.path.splitext(file_name)[0]
                text_converter = TextConverter(file_name_only)
                converted_text = text_converter.text_to_key(f.read())

                # write text to key converted text file
                new_path = os.path.join(converted_folder, file_name)
                with open(new_path, 'w', encoding='utf-16') as w:
                    w.write(converted_text)

                # write xlsx content
                wb = openpyxl.Workbook()
                # wb.remove(wb.active)
                # ws = wb.create_sheet(file_name)
                ws = wb.active
                for key, item in text_converter.sentences.items():
                    ws.append([key, self.strip_quotation_mark(item[0]), self.strip_quotation_mark(item[1])])
                wb.save(os.path.join(converted_folder, file_name_only + '.xlsx'))
                wb.close()
                print(f"converted {file_name} to {file_name_only}.xlsx")


if __name__ == '__main__':
    if len(sys.argv) == 1 or sys.argv[1] == 'help':
        print("""
        available commands :
            text_to_key <folder_path>
        """)
    elif sys.argv[1] == 'text_to_key':
        converter = FolderConverter(sys.argv[2])
        converter.text_to_key()
