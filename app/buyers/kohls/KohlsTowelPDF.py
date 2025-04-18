import os
import pandas as pd
from .KohlsPDF import *
from glob import glob
from typing import BinaryIO
from app.common import Mastersheet
from app.utils.timer import timer_decorator
from app.utils.format import (
    apply_borders_and_alignment,
    format_date,
    format_number,
    get_date,
)


class KohlsTowelPDF(Mastersheet):
    MACRO_FILE_PATH = "macro/kohls-towel.xlsm"
    NOTIFY_ADDRESS = (
        "Li & Fung (Trading) Limited\n7/F, HK SPINNERS INDUSTRIAL BUILDING\nPhase I & II,\n800 CHEUNG SHA WAN ROAD,\nKOWLOON, HONGKONG\nAir8 Pte Ltd,\n3 Kallang Junction\n#05-02 Singapore 339266\n 2% commission to WUSA",
    )

    MASTER_SHEET_COLS = {
        "program name": None,
        "upc": None,
        "plant": None,
        "material number": None,
        "sort number": None,
        "shade name": None,
        "set type": None,
        "yarn dyed matching": None,
        "pis": None,
        "sales unit": None,
    }

    def __init__(
        self,
        mastersheet: BinaryIO | str,
        po_folder: str = "uploads",
        backend: str = "pdfplumber",
    ):
        super().__init__(self.MASTER_SHEET_COLS, self.MACRO_FILE_PATH, mastersheet)
        self.po_folder = po_folder
        self.backend = backend
        self.pis = pd.read_excel(io=mastersheet, sheet_name="PACKING & PIS-1")

    def process(self):
        for po_file in glob(os.path.join(self.po_folder, "*.pdf")):
            self._create_macro(po_file=po_file)

        apply_borders_and_alignment(self.macro_ws)
        format_number(self.macro_ws, 32)  # 32 is upc col idx in macro sheet
        format_date(self.macro_ws, 13, 14)
        self._add_wb_to_zip(self.macro_wb, f"KOHLS_MACRO_SHEET_TOWEL_{get_date()}.xlsm")

    @timer_decorator
    def _create_macro(self, po_file: str):
        metadata = extract_metadata(pdf_path=po_file, backend=self.backend)
        po_sheet = extract_table_rows(pdf_path=po_file, backend=self.backend)
        (
            po,
            port_of_shipment,
            channel_type,
            ship_start_date,
            ship_cancel_date,
        ) = (
            metadata["po"],
            metadata["port of shipment"],
            metadata["channel type"],
            metadata["ship start date"],
            metadata["ship cancel date"],
        )
        sub_channel_type = "PURE PLAY ECOM" if channel_type == "ECOM" else "NIL"
        row_data = {}

        notify = (
            self.NOTIFY_ADDRESS[0] if "notify" in metadata else "2% commission to WUSA"
        )
        for row in po_sheet:
            _, _, _, qty, _, _, upc = row
            mastersheet_row = self.mastersheet_dict.get(
                upc, ["#N/A"] * len(self.mastersheet_header)
            )
            plant = int(mastersheet_row[self.MASTER_SHEET_COLS["plant"] - 1])

            program_name = mastersheet_row[
                self.MASTER_SHEET_COLS["program name"] - 1
            ].strip()

            sales_unit = mastersheet_row[
                self.MASTER_SHEET_COLS["sales unit"] - 1
            ].strip()

            if sales_unit not in row_data:
                row_data[sales_unit] = []

            packing_type = "BULK" if channel_type == "RETAIL" else "ECOM"
            # filter self.pis pandas dataframe so that it matches program_name, sales_unit, and packing_type
            filtered_pis = self.pis[
                (self.pis["Program Name"] == program_name)
                & (self.pis["Sales Unit"] == sales_unit)
                & (self.pis["Packing Type"] == packing_type)
            ]

            pis_value = filtered_pis.iloc[0]["PIS"] if not filtered_pis.empty else "N/A"
            f_part = filtered_pis.iloc[0]["F PART"] if not filtered_pis.empty else "N/A"
            date_month = int(ship_start_date.split("-")[1])
            s_part = ""
            if plant == 2100:
                if packing_type == "BULK" and sales_unit == "PC":
                    s_part = f"ALT{date_month}"
                elif packing_type == "ECOM" and sales_unit == "PC":
                    s_part = f"ALT{date_month+12}"
                elif packing_type == "ECOM" and (
                    sales_unit == "12 PC SET" or sales_unit == "6 PC SET"
                ):
                    s_part = f"ALT{date_month+24}"

            macro_row = [
                po,  # PO
                "",  # SO
                102083,  # SOLD TO PARTY
                102083,  # SHIP TO PARTY
                "W137",  # PAYMENT TERM
                "",  # INCO TERMS
                "",  # INCO TERM 2
                "",  # order reason
                100023,  # end customer
                channel_type,
                sub_channel_type,
                "REPLENISHMENT",  # order type
                ship_start_date,  # ship start date
                ship_cancel_date,  # ship cancel date
                port_of_shipment,  # port of shipment
                "NEW YORK",  # destinatoin
                "USA",  # country
                "NEW YORK",  # port of loading
                mastersheet_row[
                    self.MASTER_SHEET_COLS["material number"] - 1
                ],  # mat code
                qty,  # order qty
                "",  # amount
                mastersheet_row[
                    self.MASTER_SHEET_COLS["sort number"] - 1
                ],  # TT sort no
                mastersheet_row[self.MASTER_SHEET_COLS["shade name"] - 1],  # TT shade
                mastersheet_row[self.MASTER_SHEET_COLS["set type"] - 1],  # TT set
                "",  # Embroidery code L
                "",  # sublistatic code
                s_part,  # TT packing type for S part
                f_part,  # TT packing type for F part
                "NA",  # Destination
                mastersheet_row[
                    self.MASTER_SHEET_COLS["yarn dyed matching"] - 1
                ],  # yarn dyed
                mastersheet_row[self.MASTER_SHEET_COLS["plant"] - 1],  # plant
                upc,  # customer material
                pis_value,  # PIS
                "",  # PO AVAIL DATE
                "Saluja Tirkey",
                notify,
            ]
            row_data[sales_unit].append(macro_row)
        # loop through the row_data dictionary and append the rows to the macro worksheet
        for sales_unit, macro_rows in row_data.items():
            for macro_row in macro_rows:
                self.macro_ws.append(macro_row)
            # add a blank row after each sales unit
            self.macro_ws.append([""])
