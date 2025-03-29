import os
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


class KohlsRugsPDF(Mastersheet):
    MACRO_FILE_PATH = "macro/kohls-rugs.xlsm"
    MASTER_SHEET_COLS = {
        "upc": None,
        "plant": None,
        "material number": None,
        "sort number": None,
        "shade name": None,
        "set type": None,
        "yarn dyed matching": None,
        "pis": None,
        "sales unit": None,
        "product packing type": None,
        "printing shade no": None,
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

    def process(self):
        for po_file in glob(os.path.join(self.po_folder, "*.pdf")):
            self._create_macro(po_file=po_file)

        apply_borders_and_alignment(self.macro_ws)
        format_number(self.macro_ws, 32)  # 32 is upc col idx in macro sheet
        format_date(self.macro_ws, 13, 14)
        self._add_wb_to_zip(self.macro_wb, f"KOHLS_MACRO_SHEET_RUGS_{get_date()}.xlsm")

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
        for row in po_sheet:
            _, _, _, qty, _, _, upc = row
            mastersheet_row = self.mastersheet_dict.get(
                upc, ["#N/A"] * len(self.mastersheet_header)
            )
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
                mastersheet_row[self.MASTER_SHEET_COLS["shade name"] - 1],
                mastersheet_row[self.MASTER_SHEET_COLS["printing shade no"] - 1],
                "",  # product packing style
                mastersheet_row[self.MASTER_SHEET_COLS["product packing type"] - 1],
                "",  # production type
                mastersheet_row[self.MASTER_SHEET_COLS["set type"] - 1],  # Set No.
                mastersheet_row[self.MASTER_SHEET_COLS["yarn dyed matching"] - 1],
                "NA",  # destination
                mastersheet_row[self.MASTER_SHEET_COLS["plant"] - 1],
                upc,  # cust material
                mastersheet_row[self.MASTER_SHEET_COLS["pis"] - 1],  # PIS
                "",  # PO AVAIL DATE
                "SALUJA TIRKEY",  # AC HOLDER
                "",  # NOTIFY
            ]
            self.macro_ws.append(macro_row)
        self.macro_ws.append([""])
