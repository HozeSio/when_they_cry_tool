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

parse_pattern = re.compile(rf"""(?:
{output_pattern.pattern}|
{script_pattern.pattern}|
{dialog_pattern.pattern})
""", re.VERBOSE | re.MULTILINE)

actor_pattern = re.compile(r'<color=[^>]*>([^<]*)</color>')
remove_quotation_mark_pattern = re.compile(r'"(.*)"')


full_to_half = [
    ('～', '~'),
    ('―', '-'),
    ('ー', '-'),
    ('！', '!'),
    ('？', '?'),
    ('Ⅹ', '?'),
    ('…', '...'),
    ('０', '0'),
    ('１', '1'),
    ('２', '2'),
    ('３', '3'),
    ('４', '4'),
    ('５', '5'),
    ('６', '6'),
    ('７', '7'),
    ('８', '8'),
    ('９', '9'),
]


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
        return ''.join(self.groups)

    def is_ignore_line(self):
        return self.param2.startswith('\"<size=') or self.param4.startswith('\"<size=')

    def is_actor_line(self):
        return self.param2 == 'NULL' and self.param4 == 'NULL'

    def get_actor_text(self, param):
        return strip_quotation_mark(actor_pattern.sub(r'\1', param))

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

    def extract_text(self):
        sentences = []
        last_actor = None
        for match in parse_pattern.finditer(self.text):
            if match.group().startswith(script_method):
                sentences.append((script_method, match.group(12), match.group(13)))
            elif match.group().startswith('void'):
                sentences.append((match.group(),))
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
        line = OutputLine(match_obj)

        if line.is_ignore_line():
            return line.text

        if line.is_actor_line():
            if line.param3 == "NULL":
                return line.text
            keys = line.get_actors(line.param3)
            for key in keys:
                line.param3 = line.param3.replace(key, self.translation[key])
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
                for full_half in full_to_half:
                    translated_text = translated_text.replace(full_half[0], full_half[1])
                translated_text.replace('궃', '궂')
                translated_text = translated_text.strip()
        except KeyError:
            print(line.text)
            raise
        line.param4 = f'\"{translated_text} \"'
        return line.text

    def replace_text(self, translation: {}):
        self.translation = translation
        self.index = 0
        return output_pattern.sub(self.repl_replace_text, self.text)

    def validate_text(self):
        for method_match in output_pattern.finditer(self.text):
            line = OutputLine(method_match)
            if not line.is_actor_line():
                try:
                    param2 = strip_quotation_mark(line.param2)
                    param4 = strip_quotation_mark(line.param4)
                except ValueError:
                    print(f"\nValidation error!!\n{line.text}")
                    return False

                index = -1
                while True:
                    index = param4.find('\\', index + 1)
                    if index == -1:
                        break

                    if param4[index + 1] != '"' and param4[index + 1] != 'n':
                        print(f"\nValidation error!!\n{line.text}")
                        return False

                index = -1
                while True:
                    index = param4.find('"', index + 1)
                    if index == -1:
                        break

                    if param4[index - 1] != '\\':
                        print(f"\nValidation error!!\n{line.text}")
                        return False

            else:
                pass
        return True

    def extract_actor(self):
        actors = set()
        for method_match in output_pattern.finditer(self.text):
            line = OutputLine(method_match)
            if not line.is_actor_line():
                continue

            jp_actors = line.get_actors(line.param1)
            en_actors = line.get_actors(line.param3)
            for index in range(0, len(en_actors)):
                try:
                    actors.add((jp_actors[index], en_actors[index]))
                except IndexError:
                    pass
        return actors
