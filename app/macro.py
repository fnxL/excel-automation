from typing import BinaryIO
from fastapi import UploadFile
from app.utils.filesystem import save_files, delete_folders
from app.buyers.kohls import KohlsTowelPDF, KohlsBedsheet, KohlsRugsPDF, KohlsPOMismatch
from app.buyers.walmart import WalmartBedsheet
from app.utils.format import get_date


async def create_macro(
    customer: str,
    mastersheet: BinaryIO,
    po_files: list[UploadFile],
    backend: str,
    filename: str = None,
):
    match customer:
        case "kohls-towel-pdf":
            upload_folder = await save_files(customer, po_files)
            kohls = KohlsTowelPDF(mastersheet, upload_folder, backend)
            kohls.process()
            delete_folders([upload_folder])
            fileName = f"KOHLS_MACRO_SHEET_TOWEL_{get_date()}.zip"
            return fileName, kohls.get_zip_buffer()
        case "kohls-bedsheet":
            upload_folder = await save_files(customer, po_files)
            kohls = KohlsBedsheet(mastersheet, upload_folder)
            kohls.process()
            delete_folders([upload_folder])
            fileName = f"KOHLS_MACRO_SHEET_BEDSHEET_{get_date()}.zip"
            return fileName, kohls.get_zip_buffer()
        case "kohls-rugs-pdf":
            upload_folder = await save_files(customer, po_files)
            kohls = KohlsRugsPDF(mastersheet, upload_folder, backend)
            kohls.process()
            delete_folders([upload_folder])
            fileName = f"KOHLS_MACRO_SHEET_RUGS_{get_date()}.zip"
            return fileName, kohls.get_zip_buffer()
        case "walmart-bedsheet":
            fileName = f"WALMART_MACRO_SHEET_BEDSHEET_{get_date()}.zip"
            walmart = WalmartBedsheet(mastersheet=mastersheet, po_file=po_files[0].file)
            zip_buffer = walmart.process()
            return fileName, zip_buffer
        case "kohls-po-mismatch":
            upload_folder = await save_files(customer, po_files)
            zip_buffer = KohlsPOMismatch(
                mastersheet=mastersheet, po_folder=upload_folder, filename=filename
            ).process()
            delete_folders([upload_folder])
            return filename, zip_buffer
