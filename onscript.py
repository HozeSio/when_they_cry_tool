import re
import os
import openpyxl


class OnscriptParser:
    HEADER = ('korean',)
    parse_pattern = re.compile(r"""
    ^.*[@\\/][ ]*$
    """, re.VERBOSE | re.MULTILINE)

    def __init__(self, text):
        self.text = text

    def parse_text(self):
        sentences_jp = []
        sentences_kr = []

        for match in self.parse_pattern.finditer(self.text):
            if match.group().startswith(';'):
                sentences_jp.append(match.group().lstrip(';'))
            else:
                sentences_kr.append(match.group())

        sentences = []
        if len(sentences_jp) != len(sentences_kr):
            print(f"Sentence count missmatch!!")

        for i in range(0, max(len(sentences_kr), len(sentences_jp))):
            jp = None
            kr = None
            try:
                jp = sentences_jp[i]
                kr = sentences_kr[i]
            except IndexError:
                jp = None if not jp else jp
                kr = None if not kr else kr
            sentences.append((jp, kr))
        return sentences


class SteamParser:
    HEADER = ('japanese', 'english', 'korean')
    parse_pattern = re.compile(r"^lang.*$", re.VERBOSE | re.MULTILINE)
    text_pattern_jp = re.compile(r"(![ds]+\d*)?([^@/\\]*)[@/\\]", re.MULTILINE)

    def __init__(self, text):
        self.text = text

    def parse_text(self):
        sentences_jp = []
        sentences_en = []

        for match in self.parse_pattern.finditer(self.text):
            match_text = match.group()
            param = match_text[6:]
            params = param.split(':')
            sentences = (p.replace('', '') for p in params if p.endswith(('@', '\\', '/')))
            if match_text.startswith('langjp'):
                sentences_jp.extend(sentences)
            else:
                sentences_en.extend(sentences)

        sentences = []
        if len(sentences_jp) != len(sentences_en):
            print(f"Sentence count missmatch!!")

        for i in range(0, max(len(sentences_en), len(sentences_jp))):
            jp = None
            en = None
            try:
                jp = sentences_jp[i]
                en = sentences_en[i]
            except IndexError:
                jp = None if not jp else jp
                en = None if not en else en
            sentences.append((jp, en))
        return sentences


class FolderParser:
    def __init__(self, folder_path):
        self.folder_path = os.path.normpath(folder_path)
        (self.folder_directory, self.folder_name) = os.path.split(self.folder_path)

    def export_text(self, mode='steam'):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_output')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)

        for file_name in os.listdir(self.folder_path):
            if not file_name.endswith('.txt'):
                continue

            file_path = os.path.join(self.folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                print(f"start converting {file_name}....", end='')
                text_converter = SteamParser(f.read()) if mode == 'steam' else OnscriptParser(f.read())
                sentences = text_converter.parse_text()

                file_name_only = os.path.splitext(file_name)[0]
                self.save_xlsx(text_converter.HEADER, sentences, os.path.join(converted_folder, f'{file_name_only}.xlsx'))
                print(f"finished")

    def save_xlsx(self, header, sentences, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(header)
        for sentence in sentences:
            ws.append(sentence)
        wb.save(path)
        wb.close()
