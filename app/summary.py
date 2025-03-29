from typing import List, BinaryIO
from app.buyers.kohls import KohlsPOSummary
from app.utils.format import get_date


async def create_summary(poNumbers: List[int], mastersheet: BinaryIO):
    kohlsPO = KohlsPOSummary(mastersheet=mastersheet, poNumbers=poNumbers)
    kohlsPO.process()
    fileName = f"KOHLS_PO_SUMMARY_{get_date()}.xlsx"
    return fileName, kohlsPO.get_bytes_buffer()
