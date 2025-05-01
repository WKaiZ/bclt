"""
Filename: spreadsheet_script.py
Author: Wesley Zheng
Date: 2025-4-29
Description: ​This script automates the process of eliminating duplicate summaries for the same citizen petitions and responses.
"""

import re
from openpyxl import load_workbook

def clean_excel_cells(file_path, regex_pattern):
    wb = load_workbook(file_path)
    pattern = re.compile(regex_pattern)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row = 2, min_col = 2):
            for cell in row:
                new_value = pattern.sub('', cell.value)
                if new_value != cell.value:
                    cell.value = new_value
    wb.save(file_path)

def collect_bad_summaries(file_path, regex_pattern):
    wb = load_workbook(file_path)
    pattern = re.compile(regex_pattern)

    files = []

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            if sum(pattern.search(cell) is not None for cell in row) >= 1:
                files.append(row[0])

    return files

def write_list_to_txt(data_list, file_path):
    with open(file_path, 'w', encoding='utf-8') as f:
        for item in data_list:
            f.write(f"{item}\n")


def remove_rows_with_duplicate_numbers(file_path):
    wb = load_workbook(file_path)
    first_or_last = 1 if re.findall(r'citizens|responses', file_path)[0] == 'citizens' else -1
    for ws in wb.worksheets:
        seen_numbers = set()
        rows_to_delete = []
        if first_or_last == 1:
            for idx, row in enumerate(ws.iter_rows(min_row = 2), start = 2):
                row_has_duplicate = False
                row_numbers = set()

                number = re.findall(r'\d{4}', row[0].value)[1]
                if number in seen_numbers:
                    row_has_duplicate = True
                else:
                    row_numbers.add(number)

                if row_has_duplicate:
                    rows_to_delete.append(idx)
                else:
                    seen_numbers.update(row_numbers)
        else:
            for idx in range(ws.max_row, 1, -1):
                row = ws[idx]
                row_has_duplicate = False
                row_numbers = set()

                number = re.findall(r'\d{4}', row[0].value)[1]
                if number in seen_numbers:
                    row_has_duplicate = True
                else:
                    row_numbers.add(number)

                if row_has_duplicate:
                    rows_to_delete.append(idx)
                else:
                    seen_numbers.update(row_numbers)
                    
        for row_idx in reversed(rows_to_delete):
            ws.delete_rows(row_idx)

    wb.save(file_path)

clean_excel_cells("outputs_citizens.xlsx", r'^.*:\s|^\d\.\s')
clean_excel_cells("outputs_responses.xlsx", r'^.*:\s|^\d\.\s')
remove_rows_with_duplicate_numbers("outputs_responses.xlsx")
remove_rows_with_duplicate_numbers("outputs_citizens.xlsx")
bad_summaries = collect_bad_summaries("outputs_citizens.xlsx", r'Not Mentioned')
write_list_to_txt(bad_summaries, "bad_summaries_petitions.txt")
bad_summaries = collect_bad_summaries("outputs_responses.xlsx", r'Not Mentioned')
write_list_to_txt(bad_summaries, "bad_summaries_responses.txt")

