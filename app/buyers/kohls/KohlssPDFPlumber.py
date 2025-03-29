import pdfplumber
import re
from typing import List, Dict


def remove_duplicates(data: List[List[str | None]]):
    unique_data = []
    for row in data:
        if row not in unique_data:
            unique_data.append(row)
    return unique_data


def remove_none(data: List[List[str | None]]):
    """Remove None items in a list of lists even remove inside the list"""
    new_data = []
    for row in data:
        new_row = []
        for item in row:
            if item is not None:
                new_row.append(item)
        if new_row:
            new_data.append(new_row)
    return new_data


def clean_data(data: List[List[str | None]]):
    clean_data = []
    line_number = 1
    line_item = []
    for row in data:
        match = re.search(r"UPC/EAN \(GTIN\) (\d+)", row[0])
        if match:
            line_item.append(int(match.group(1)))
        try:
            line_num = int(row[0])
            line_item = row
        except ValueError:
            continue

        clean_data.append(row)

    for row in clean_data:
        row[3] = row[3].split(" ")[0]
        row[3] = int(row[3].replace(",", ""))

    return clean_data


def extract_metadata_pdfplumber(pdf_path: str, page_range: tuple = None):
    search_patterns = {
        "po": r"Order Number.*\n\s*(\d+)",
        "port of shipment": r"FOB -\s*(\w+)",
        "channel type": r"Order Indicator",
        "ship window": r"Shipment Window.*\n.*(\d{4}-\d{2}-\d{2} / \d{4}-\d{2}-\d{2})",
    }
    meta_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[slice(*page_range) if page_range else slice(len(pdf.pages))]
        for page in pages[0:1]:
            text = page.extract_text(layout=True)
            if not text:
                continue

            for column, pattern in search_patterns.items():
                match = re.search(pattern, text, re.MULTILINE)
                if match:
                    if column == "ship window":
                        meta_data["ship start date"] = match.group(1).split(" / ")[0]
                        meta_data["ship cancel date"] = match.group(1).split(" / ")[1]
                    elif column == "port of shipment":
                        incoterm = match.group(1)
                        if "MUNDRA" in incoterm:
                            meta_data[column] = "MUNDRA"
                        else:
                            meta_data[column] = "JNPT"
                    elif column == "channel type":
                        channel = match.group()
                        if channel is not None:
                            meta_data[column] = "ECOM"
                    elif column == "po":
                        meta_data[column] = int(match.group(1))
                else:
                    if column == "channel type":
                        meta_data[column] = "RETAIL"

    return meta_data


def extract_table_rows_pdfplumber(pdf_path: str, page_range: tuple = None):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[slice(*page_range) if page_range else slice(len(pdf.pages))]
        for page in pages:
            table = page.extract_table(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 10,
                    "min_words_vertical": 1,
                    "min_words_horizontal": 1,
                }
            )
            for row in table:
                data.append(row)
    data = remove_duplicates(data)
    data = remove_none(data)
    data = clean_data(data)
    return data
