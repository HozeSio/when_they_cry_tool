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
    text_pattern_split = re.compile(r"(![ds]+\d*)?([^@/\\]*)[@/\\]")
    text_pattern_en = re.compile(r"[^^]*\^([^^]*)\^?")

    def __init__(self, text):
        self.text = text

    def save_text_block(self, sentences_jp, sentences_en):
        rows = []
        for i in range(0, max(len(sentences_en), len(sentences_jp))):
            jp = None
            en = None
            try:
                jp = sentences_jp[i]
            except IndexError:
                jp = None if not jp else jp
            try:
                en = sentences_en[i]
            except IndexError:
                en = None if not en else en
            rows.append((jp, en))
        return rows

    def parse_text(self):
        sentences_jp = []
        sentences_en = []
        rows = []
        current_lang = 'jp'

        for match in self.parse_pattern.finditer(self.text):
            match_text = match.group()
            param = match_text[6:]
            params = param.split(':') if param[0] == ':' or param.find('dwave_') != -1 else (param,)
            sentences = list(p.replace('', '') for p in params if p.endswith(('@', '\\', '/')))
            if match_text.startswith('langjp'):
                if current_lang != 'jp':
                    current_lang = 'jp'
                    # save previous text block
                    rows.extend(self.save_text_block(sentences_jp, sentences_en))
                    sentences_jp.clear()
                    sentences_en.clear()

                for sentence in sentences:
                    for sub_match in self.text_pattern_split.finditer(sentence):
                        if sub_match and sub_match.group(2):
                            sentences_jp.append(sub_match.group(2))
            else:
                current_lang = 'en'
                for sentence in sentences:
                    for sub_match in self.text_pattern_split.finditer(sentence):
                        if sub_match and sub_match.group(2):
                            for sub_sub_match in self.text_pattern_en.finditer(sub_match.group(2)):
                                if sub_sub_match:
                                    sentences_en.append(sub_sub_match.group(1))

        rows.extend(self.save_text_block(sentences_jp, sentences_en))
        return rows


class FolderParser:
    def __init__(self, folder_path):
        self.folder_path = os.path.normpath(folder_path)
        (self.folder_directory, self.folder_name) = os.path.split(self.folder_path)

    def export_text(self, mode='steam'):
        converted_folder = os.path.join(self.folder_directory, self.folder_name + '_output')
        if not os.path.exists(converted_folder):
            os.mkdir(converted_folder)
        print(f"Exporting to {converted_folder}")

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
