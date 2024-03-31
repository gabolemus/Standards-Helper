"""Module with the functions to work with the worksheets."""

import re
from typing import Union

from openpyxl import load_workbook

import standards.config as cfg


def get_standards(file_path: str, worksheet_name: str, min_row: int, min_col: int, max_col: int) -> list[dict[str, Union[str, None]]]:
    """Gets the list of standards from a given file and worksheet."""
    workbook = load_workbook(file_path)
    worksheet = workbook[worksheet_name]
    rows = worksheet.iter_rows(
        min_row=min_row, min_col=min_col, max_col=max_col, values_only=True)

    standards: list[dict[str, Union[str, None]]] = []
    for row in rows:
        if row[0] and row[1]:
            if row[0] != "No." and row[0] != "Criterion #":
                standard = {
                    "id": row[0],
                    "text": row[1],
                    "level": row[2] if len(row) > 2 and row[2] else None
                }
                standards.append(standard)

    return standards


def get_original_standards() -> list[dict[str, Union[str, None]]]:
    """Gets the list with the original standards."""
    return get_standards(cfg.original_standards_file, cfg.original_standards_ws[0], 4, 2, 4)


def get_new_standards() -> list[dict[str, Union[str, None]]]:
    """Gets the list with the new standards."""
    return get_standards(cfg.new_standards_file, "Unified Standard", 4, 1, 3)

# def get_original_standards() -> list[dict[str, Union[str, None]]]:
#     """Gets the list with the original standards."""
#     original_workbook = load_workbook(cfg.original_standards_file)
#     original_worksheet = original_workbook[cfg.original_standards_ws[0]]
#     original_rows = original_worksheet.iter_rows(
#         min_row=4, min_col=2, max_col=4, values_only=True)
#
#     original_standards: list[dict[str, Union[str, None]]] = []
#     for row in original_rows:
#         if row[0] and row[1] and row[0] != "No." and \
#             (re.match(r"[A-Z]{1}\d.+", row[0]) or
#                 re.match(r"(?<!\S)(?=[ivxlcdm]+\.$)[ivxlcdm]+\.", row[0])):
#             standard = {
#                 "id": row[0],
#                 "text": row[1],
#                 "level": row[2] if row[2] else None
#             }
#             original_standards.append(standard)
#
#     return original_standards
#
#
# def get_new_standards() -> list[dict[str, Union[str, None]]]:
#     """Gets the list with the new standards."""
#     new_workbook = load_workbook(cfg.new_standards_file)
#     new_worksheet = new_workbook["Unified Standard"]
#     new_rows = new_worksheet.iter_rows(
#         min_row=4, min_col=1, max_col=3, values_only=True)
#
#     new_standards: list[dict[str, Union[str, None]]] = []
#     for row in new_rows:
#         if row[0] and row[1] and row[0] != "Criterion #":
#             standard = {
#                 "id": row[0],
#                 "text": row[1],
#                 "level": row[2] if row[2] else None
#             }
#             new_standards.append(standard)
#
#     return new_standards
