# import pymupdf
# from typing import List


# def extract_metadata_pymupdf(pdf_path: str):
#     doc = pymupdf.open(pdf_path)
#     text = doc[0].get_text("blocks", sort=True)
#     metadata = {}
#     for block in text:
#         line = block[-3].split("\n")
#         for i, word in enumerate(line):
#             if word.lower() == "order number":
#                 metadata["po"] = int(line[i + 3])
#             elif word.lower() == "incoterm":
#                 port = line[i + 3].split(" ")
#                 metadata["port of shipment"] = port[2]

#                 if port[2] == "MUMBAI":
#                     metadata["port of shipment"] = "JNPT"
#                 else:
#                     metadata["port of shipment"] = "MUNDRA"

#             elif word.lower() == "shipment window":
#                 shipment_window = line[i + 3].split(" / ")
#                 metadata["ship start date"] = shipment_window[0]
#                 metadata["ship cancel date"] = shipment_window[1]
#             elif word.lower() == "order indicator":
#                 metadata["channel type"] = "ECOM"

#     if "channel type" not in metadata:
#         metadata["channel type"] = "RETAIL"

#     return metadata


# def remove_none(data: List[str | None]):
#     """Remove None items from a list"""
#     new_data = []
#     for item in data:
#         if item is not None:
#             new_data.append(item)
#     return new_data


# def remove_duplicates(data: List[List[str | None]]):
#     unique_data = []
#     for row in data:
#         if row not in unique_data:
#             unique_data.append(row)
#     return unique_data


# def clean_data(data: List[List[str | None]]):
#     clean_data = []
#     line_item = []
#     for row in data:
#         if "UPC/EAN" in row[0]:
#             upc = row[0].split("(GTIN)")[1]
#             line_item.append(int(upc.strip()))
#         try:
#             line_num = int(row[0])
#             line_item = row
#         except ValueError:
#             continue

#         clean_data.append(row)

#     for row in clean_data:
#         row[3] = row[3].split(" ")[0]
#         row[3] = int(row[3].replace(",", ""))

#     return clean_data


# def extract_table_rows_pymupdf(pdf_path: str):
#     doc = pymupdf.open(pdf_path)
#     data = []
#     start = False
#     for page in doc:
#         tables = page.find_tables(
#             vertical_strategy="lines",
#             horizontal_strategy="lines",
#             intersection_tolerance=10,
#             min_words_vertical=1,
#             min_words_horizontal=1,
#         ).tables
#         if not tables:
#             continue

#         table = tables[0]
#         rows = table.extract()
#         for row in rows:
#             if start:
#                 data.append(remove_none(row))

#             if row[0]:
#                 if row[0].startswith("Line #"):
#                     start = True
#                     data.append(remove_none(row))

#     data = remove_duplicates(data)
#     data = clean_data(data)
#     return data
