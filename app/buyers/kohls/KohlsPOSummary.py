import os
from glob import glob
from typing import BinaryIO, List
from app.utils.timer import timer_decorator
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from io import BytesIO
import pandas as pd


class KohlsPOSummary:
    BASE_COLUMNS = [
        "PO",
        "SO",
        "PROGRAM",
        "TOTAL",
        "VALUE",
        "SW START",
        "SW END",
        "PACKING",
    ]

    def __init__(
        self,
        mastersheet: BinaryIO | str,
        poNumbers: List[int],
        sheet_name: str = "Sheet1",
    ):

        self.master_df = pd.read_excel(io=mastersheet, sheet_name=sheet_name)
        self.poNumbers = poNumbers
        filtered_df = self.master_df[self.master_df["PO Number"].isin(self.poNumbers)]
        self.size_desc = set(filtered_df["Size Desc"])
        columns = self.BASE_COLUMNS[:3] + list(self.size_desc) + self.BASE_COLUMNS[3:]
        self.df = pd.DataFrame(columns=columns)

    def process(self):
        for po in self.poNumbers:
            filtered_df = self.master_df[self.master_df["PO Number"] == po]
            # group by size desc
            grouped_df = filtered_df.groupby("Size Desc").agg(
                {
                    "Start X Factory Date": "first",
                    "Last X Factory Date": "first",
                    "PO Type Code": "first",
                    "Style Desc": "first",
                    "Ordered Units": "sum",
                    "Ordered First Cost $": "sum",
                }
            )
            # create a new row to add into the dataframe as per the columns header
            new_row = {
                "PO": po,
                "SO": "",
                "PROGRAM": filtered_df["Style Desc"].values[0],
                "TOTAL": grouped_df["Ordered Units"].sum(),
                "VALUE": grouped_df["Ordered First Cost $"].sum(),
                "SW START": grouped_df["Start X Factory Date"].values[0],
                "SW END": grouped_df["Last X Factory Date"].values[0],
                "PACKING": (
                    "ECOM" if filtered_df["PO Type Code"].values[0] == "IE" else "BULK"
                ),
            }
            for size in self.size_desc:
                new_row[size] = (
                    grouped_df.loc[size, "Ordered Units"]
                    if size in grouped_df.index
                    else 0
                )
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)

    def get_bytes_buffer(self):
        output = BytesIO()
        self.df["SW START"] = self.df["SW START"].dt.strftime("%d-%m-%Y")
        self.df["SW END"] = self.df["SW END"].dt.strftime("%d-%m-%Y")
        with pd.ExcelWriter(
            output,
            engine="openpyxl",
            date_format="dd-mm-yy",
            datetime_format="dd-mm-yy",
        ) as writer:
            self.df.to_excel(writer, index=False, sheet_name="Summary")
        output.seek(0)  # Reset the pointer to the beginning of the stream
        return output
