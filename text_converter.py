import sys
import re

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

actor_pattern = re.compile(r'"<color=.*>(.*)</color>"')


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
        return self.param1 != 'NULL'

    def get_actor(self):
        return actor_pattern.match(self.param3).groups()[0]


class TextConverter:
    def __init__(self, text):
        self.text = text
        self.remove_quotation_mark_pattern = re.compile(r'"(.*)"')
        self.translation = {}

    def strip_quotation_mark(self, line: str):
        m = self.remove_quotation_mark_pattern.match(line)
        if not m:
            print(line)
            raise ValueError
        return m.groups()[0]

    def extract_text(self):
        sentences = []
        last_actor = None
        for match in method_pattern.finditer(self.text):
            line = OutputLine(match)

            if line.is_actor_line():
                last_actor = line.get_actor()
                continue
            if line.is_ignore_line():
                continue

            sentences.append((last_actor, self.strip_quotation_mark(line.param2), self.strip_quotation_mark(line.param4)))
            last_actor = None
        return sentences

    def repl_replace_text(self, match_obj) -> str:
        line = OutputLine(match_obj)

        if line.is_actor_line() or line.is_ignore_line():
            return line.text

        # replace english text to translation text based on japanese text
        try:
            key = self.strip_quotation_mark(line.param2)
            # empty text handling
            if not key:
                key = None
            translated_text = self.translation[key]
        except KeyError:
            print(line.text)
            raise
        line.param4 = f'\"{translated_text}\"'
        return line.text

    def replace_text(self, translation: {}):
        self.translation = translation
        return method_pattern.sub(self.repl_replace_text, self.text)

    def validate_text(self):
        for method_match in method_pattern.finditer(self.text):
            line = OutputLine(method_match)
            if not line.is_actor_line():
                try:
                    self.strip_quotation_mark(line.param2)
                    self.strip_quotation_mark(line.param4)
                except ValueError:
                    print(line.text, file=sys.stderr)
                    return False
            else:
                pass
        return True
