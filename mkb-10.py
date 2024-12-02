import os
import re
from typing import Dict, List, Tuple
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment


class PDFProcessor:
    def __init__(self, pdf_path: str):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"Файл {pdf_path} не найден.")
        self.pdf_path = pdf_path

    def extract_data(self) -> Dict[str, Tuple[str, List[Tuple[str, str]]]]:
        grouped_data = {}
        code_pattern = re.compile(r'^([A-Z][0-9]{2}\.[0-9]+)\s+(.+)$')
        group_pattern = re.compile(r'^([A-Z][0-9]{2})\s+(.+)$')

        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.splitlines()
                for line in lines:
                    group_match = group_pattern.match(line)
                    if group_match:
                        group_code, group_name = group_match.groups()
                        grouped_data[group_code] = (group_name, [])
                    else:
                        match = code_pattern.match(line)
                        if match:
                            code, name = match.groups()
                            group_code = code.split(".")[0]
                            if group_code in grouped_data:
                                grouped_data[group_code][1].append((code, name))

        return grouped_data


class ExcelWriter:
    def __init__(self, output_path: str):
        self.output_path = output_path

    def write_data(self, grouped_data: Dict[str, Tuple[str, List[Tuple[str, str]]]]):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "МКБ-10"

        sheet.append(["Группа", "Код", "Наименование заболевания"])
        header_font = Font(bold=True, size=12)
        for cell in sheet[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")


        for group_code, (group_name, entries) in grouped_data.items():
            group_row = [group_code, group_name, ""]
            sheet.append(group_row)
            group_font = Font(bold=True, size=11)
            for cell in sheet[sheet.max_row]:
                cell.font = group_font


            for code, name in entries:
                sheet.append([group_code, code, name])
            sheet.append([])


        for col in sheet.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            sheet.column_dimensions[col[0].column_letter].width = max_length + 2

        workbook.save(self.output_path)
        print(f"Данные сохранены в {self.output_path}")


class MKBProcessor:
    def __init__(self, pdf_path: str, excel_path: str):
        self.pdf_processor = PDFProcessor(pdf_path)
        self.excel_writer = ExcelWriter(excel_path)

    def process(self):
        grouped_data = self.pdf_processor.extract_data()
        self.excel_writer.write_data(grouped_data)


if __name__ == "__main__":
    pdf_file = "2018-mkb10.pdf"
    excel_file = "МКБ10-группы.xlsx"

    try:
        processor = MKBProcessor(pdf_file, excel_file)
        processor.process()
    except Exception as e:
        print(f"Ошибка: {e}")
