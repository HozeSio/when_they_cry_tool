import sys
import re

output_pattern = re.compile(r"""
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

script_method = 'ModCallScriptSection'
script_pattern = re.compile(rf"""
{script_method}
\(
"(.*)"
,
"(.*)"
\);
""", re.VERBOSE | re.MULTILINE)

dialog_pattern = re.compile(r'void\ dialog\d+\s*\(\)', re.VERBOSE | re.MULTILINE)

play_bgm_method = 'PlayBGM'
play_bgm_pattern = re.compile(rf"""
(
{play_bgm_method}
\(\s*
)               # [11]
([^,]*)         # BGM channel [12]
(\s*,\s*)
([^,]*)         # BGM name [14]
(\s*,\s*)
([^,]*)
(\s*,\s*)
([^,]*)
(\s*\);)        # [19]
""", re.VERBOSE | re.MULTILINE)

fade_bgm_method = 'FadeOutBGM'
fade_bgm_pattern = re.compile(rf"""
{fade_bgm_method}
\(\s*
([^,]*)
\s*,\s*
([^,]*)
\s*,\s*
([^,]*)
\s*\);
""", re.VERBOSE | re.MULTILINE)

parse_pattern = re.compile(rf"""(?:
{output_pattern.pattern}|
{script_pattern.pattern}|
{dialog_pattern.pattern}|
{play_bgm_pattern.pattern}|
{fade_bgm_pattern.pattern})
""", re.VERBOSE | re.MULTILINE)

repl_pattern = re.compile(rf"""(?:
{output_pattern.pattern}|
{play_bgm_pattern.pattern})
""", re.VERBOSE | re.MULTILINE)

actor_pattern = re.compile(r'<color=[^>]*>([^<]*)</color>')
remove_quotation_mark_pattern = re.compile(r'"(.*)"')


full_to_half_ascii = dict((i + 0xFEE0, i) for i in range(0x21, 0x7F))
custom_map = {
    ord('～'): '~',
    ord('―'): '-',
    ord('ー'): '-',
    ord('…'): '...',
    ord('궃'): '궂',
    ord('줫'): '줬',
    ord('졋'): '졌',
    ord('됬'): '됐',
}


def strip_quotation_mark(line: str):
    m = remove_quotation_mark_pattern.match(line)
    if not m:
        print(line)
        raise ValueError
    return m.groups()[0]


class OutputLine:
    def __init__(self, match_obj):
        self.match_obj = match_obj
        self.groups = list(match_obj.groups())

    @property
    def param1(self):
        return self.groups[1]

    @param1.setter
    def param1(self, param):
        self.groups[1] = param

    @property
    def param2(self):
        return self.groups[3]

    @param2.setter
    def param2(self, param):
        self.groups[3] = param

    @property
    def param3(self):
        return self.groups[5]

    @param3.setter
    def param3(self, param):
        self.groups[5] = param

    @property
    def param4(self):
        return self.groups[7]

    @param4.setter
    def param4(self, param):
        self.groups[7] = param

    @property
    def param5(self):
        return self.groups[9]

    @param5.setter
    def param5(self, param):
        self.groups[9] = param

    @property
    def text(self):
        return ''.join(self.groups[:11])

    def is_ignore_line(self):
        return self.param2.startswith('\"<size=') or self.param4.startswith('\"<size=')

    def is_actor_line(self):
        return self.param2 == 'NULL' and self.param4 == 'NULL'

    def get_actor_text(self, param):
        return strip_quotation_mark(actor_pattern.sub(r'\1', param)) if param != 'NULL' else None

    def get_actor1(self):
        return self.get_actor_text(self.param1)

    def get_actor2(self):
        return self.get_actor_text(self.param3)

    def get_actors(self, param):
        actors = []
        for match in actor_pattern.finditer(param):
            actors.append(match.group(1))
        return actors


class TextConverter:
    def __init__(self, text):
        self.text = text
        self.translation = {}
        self.index = 0
        self.last_translated_text:str = None

    def extract_text(self):
        sentences = []
        last_actor = None
        for match in parse_pattern.finditer(self.text):
            if match.group().startswith(script_method):
                sentences.append((script_method, match.group(12), match.group(13)))
            elif match.group().startswith('void'):
                sentences.append((match.group(),))
            elif match.group().startswith(play_bgm_method):
                sentences.append((play_bgm_method, strip_quotation_mark(match.group(17))))
            elif match.group().startswith(fade_bgm_method):
                pass
            else:
                line = OutputLine(match)

                if line.is_actor_line():
                    last_actor = line.get_actor2()
                    continue
                if line.is_ignore_line():
                    continue

                sentences.append((last_actor, strip_quotation_mark(line.param2), strip_quotation_mark(line.param4)))
                last_actor = None
        return sentences

    def repl_replace_text(self, match_obj) -> str:
        if match_obj.group().startswith(play_bgm_method):
            groups = list(match_obj.groups())
            key = f"{self.index}_{play_bgm_method}"
            try:
                groups[14] = f"\"{self.translation[key]}\""
            except KeyError:
                print("BGM not found...(ignore if xlsx is old version)")
                return match_obj.group()
            self.index += 1
            return ''.join(groups[11:20])

        line = OutputLine(match_obj)

        if line.is_ignore_line():
            return line.text

        if line.is_actor_line():
            keys = line.get_actors(line.param1)
            line.param3 = line.param1
            for key in keys:
                line.param3 = line.param3.replace(key, self.translation[key])
                line.param3 = line.param3.translate(full_to_half_ascii)
            return line.text

        # replace english text to translation text based on japanese text
        try:
            clean_param = strip_quotation_mark(line.param2)
            if not clean_param:
                clean_param = None
            key = f"{self.index}_{clean_param}"
            self.index += 1
            # empty text handling
            if not key:
                key = None
            translated_text = self.translation[key]
            if translated_text:
                if len(translated_text.strip()) > 1:
                    translated_text = f"{translated_text.strip()} "
                translated_text = translated_text.translate(full_to_half_ascii)
                translated_text = translated_text.translate(custom_map)
                translated_text = translated_text.replace('&', ' & ')
        except KeyError:
            print('Found text inconsistency between txt and xlsx at ')
            print(list(self.translation.keys())[self.index])
            print(clean_param)
            raise
        line.param4 = f'\"{translated_text}\"'
        return line.text

    def replace_text(self, translation: {}):
        self.translation = translation
        self.index = 0
        self.last_translated_text = None
        return repl_pattern.sub(self.repl_replace_text, self.text)

    def validate_text(self):
        result = True
        for method_match in output_pattern.finditer(self.text):
            line = OutputLine(method_match)
            if not line.is_actor_line():
                try:
                    param2 = strip_quotation_mark(line.param2)
                    param4 = strip_quotation_mark(line.param4)

                    index = -1
                    while True:
                        index = param4.find('\\', index + 1)
                        if index == -1:
                            break

                        if param4[index + 1] != '"' and param4[index + 1] != 'n':
                            raise Exception

                    index = -1
                    while True:
                        index = param4.find('"', index + 1)
                        if index == -1:
                            break

                        if param4[index - 1] != '\\':
                            raise Exception
                except Exception:
                    print(f"\nValidation error!!\n{line.text}")
                    result = False
            else:
                pass
        return result

    def extract_actor(self):
        actors = set()
        for method_match in output_pattern.finditer(self.text):
            line = OutputLine(method_match)
            if not line.is_actor_line():
                continue

            jp_actors = line.get_actors(line.param1)
            en_actors = line.get_actors(line.param3)
            longest = max(len(jp_actors), len(en_actors))
            for index in range(0, longest):
                actors.add((jp_actors[index] if len(jp_actors) > index else '',
                            en_actors[index] if len(en_actors) > index else ''))
        return actors
