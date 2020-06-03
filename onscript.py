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
    text_pattern_split = re.compile(r"(\s*![ds]+\d*)?([^@/\\]*)[@/\\]?")
    text_pattern_en = re.compile(r"[^^]*\^([^^]*)\^?")

    def __init__(self, text):
        self.text = text
        self.translation = {}
        self.index = 0

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

    def get_sentences(self, lang_text):
        param = lang_text[6:]
        if not param:
            return []
        result = []
        if lang_text.startswith('langjp'):
            params = param.split(':') if param[0] == ':' or param.find('dwave_') != -1 else (param,)
            sentences = list(p.replace('', '') for p in params if not p.startswith('dwave'))
            for sentence in sentences:
                for sub_match in self.text_pattern_split.finditer(sentence):
                    if sub_match.group(2):
                        result.append(sub_match.group(2))
        elif lang_text.startswith('langen'):
            param = param.replace('', '')
            for sub_match in self.text_pattern_en.finditer(param):
                result.append(sub_match.group(1))
        else:
            raise NotImplementedError()
        return result

    def parse_text(self):
        sentences_jp = []
        sentences_en = []
        rows = []
        current_lang = 'jp'

        for match in self.parse_pattern.finditer(self.text):
            match_text = match.group()
            param = match_text[6:]
            if not param:
                continue
            if match_text.startswith('langjp'):
                if current_lang != 'jp':
                    current_lang = 'jp'
                    # save previous text block
                    rows.extend(self.save_text_block(sentences_jp, sentences_en))
                    sentences_jp.clear()
                    sentences_en.clear()

                sentences_jp.extend(self.get_sentences(match_text))
            else:
                current_lang = 'en'
                sentences_en.extend(self.get_sentences(match_text))

        rows.extend(self.save_text_block(sentences_jp, sentences_en))
        return rows

    def replace_text_int(self, match) -> str:
        match_text = match.group()
        param = match_text[6:]
        if not param or not match_text.startswith('langen'):
            return match_text

        sentences = self.get_sentences(match_text)
        for sentence in sentences:
            found = False
            for i in range(self.index, len(self.translation)):
                if sentence == self.translation[i][1]:
                    found = True
                    self.index = i + 1
                    match_text = match_text.replace(sentence, self.translation[i][2] or '')
                    break
            if not found:
                print(f"\nExpected translation at row {self.index + 1} but couldn't find translation for {sentence}")
                raise Exception('Translation not found Exception')
        return match_text

    def replace_text(self, translation):
        self.translation = translation
        self.index = 0
        return self.parse_pattern.sub(self.replace_text_int, self.text)


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

    def load_xlsx(self, path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        translation = []
        for row in ws.rows:
            translation.append((row[0].value, row[1].value, row[2].value))
        wb.close()
        return translation

    def replace_text(self, translation_folder):
        replaced_folder = os.path.join(self.folder_directory, self.folder_name + '_translated')
        if not os.path.exists(replaced_folder):
            os.mkdir(replaced_folder)

        translation_folder = os.path.normpath(translation_folder)
        for file_name in os.listdir(translation_folder):
            (file_name_only, ext) = os.path.splitext(file_name)
            script_file_name = f'{file_name_only}.txt'
            script_path = os.path.join(self.folder_path, script_file_name)
            if not os.path.exists(script_path):
                continue

            print(f'start translating {script_file_name}....', end='')
            file_path = os.path.join(translation_folder, file_name)
            translation = self.load_xlsx(file_path)

            with open(script_path, 'r', encoding='utf-8') as f:
                text_converter = SteamParser(f.read())
                replaced_text = text_converter.replace_text(translation)
                with open(os.path.join(replaced_folder, script_file_name), 'w', encoding='utf-8') as o:
                    o.write(replaced_text)

            print(f'{script_file_name} finished')
