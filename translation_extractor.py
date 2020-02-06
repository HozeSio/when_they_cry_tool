import os
import re
import openpyxl


class TextExtractor:
    pattern = re.compile(r'![sdw]\d+')

    def extract_text(self, file_path: str):
        (file_directory, file_name) = os.path.split(file_path)
        (file_name_only, ext) = os.path.splitext(file_name)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        with open(file_path, 'r', encoding='utf-8') as f:
            ws = wb.create_sheet(file_name_only)
            while True:
                line = f.readline()
                if not line:
                    break

                line = line.strip()
                if not line:
                    continue

                last_char = line[len(line) - 1]
                if last_char != '@' and last_char != '\\':
                    continue

                line = line.rstrip("@\\")
                line = line.replace('「', '\\\"').replace('」', '\\\"')
                line = line.replace('!sd', '')
                line = self.pattern.sub('', line)
                for part in line.split('@'):
                    ws.append([part])

        wb.save(os.path.join(file_directory, file_name_only + '.xlsx'))
        wb.close()
