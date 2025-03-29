import shutil
import os
import zipfile
from fastapi import UploadFile
from typing import BinaryIO
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.worksheet.worksheet import Worksheet
from python_calamine import CalamineWorkbook
from io import BytesIO
from datetime import date


def excel_cell_to_index(cell_reference):
    """
    Converts an Excel cell reference to row and column indices for a 2D array.

    Args:
    cell_reference (str): The cell reference in the format "C1", "D2", etc.

    Returns:
    tuple: A tuple containing the row and column indices.
    """
    column_letters = ""
    row_number = ""

    for char in cell_reference:
        if char.isalpha():
            column_letters += char
        elif char.isdigit():
            row_number += char
        else:
            raise ValueError("Invalid cell reference format")

    column_index = 0
    for i, char in enumerate(reversed(column_letters)):
        column_index += (ord(char) - ord("A") + 1) * (26**i)

    # Convert to 0 based
    row_index = int(row_number) - 1
    column_index -= 1

    return (row_index, column_index)


def column_letter_to_idx(column_letter):
    """
    Converts an Excel column letter to its corresponding index.

    Args:
    column_letter (str): The column letter(s) to convert.

    Returns:
    int: The corresponding column index.
    """
    index = 0
    for i, char in enumerate(reversed(column_letter)):
        index += (ord(char) - ord("A") + 1) * (26**i)

    return index


def zip_xlsx_files(path):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(".xlsx"):
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, path)
                    with open(file_path, "rb") as f:
                        zipf.writestr(arcname, f.read())

    zip_buffer.seek(0)
    return zip_buffer


def get_po_sheet(po_file: str | BinaryIO, sheet_idx: int = 0):
    """Returns python object of the first sheet of the given PO file
    Using calamine library to read the sheet
    """

    if isinstance(po_file, str):
        wb = CalamineWorkbook.from_path(po_file)
    else:
        wb = CalamineWorkbook.from_filelike(po_file)

    ws = wb.get_sheet_by_index(sheet_idx).to_python(skip_empty_area=True)

    return ws


# checks if the current row tuple is completely of empty string
def isEnd(tuple):
    if all(element == "" for element in tuple):
        return True

    return False
