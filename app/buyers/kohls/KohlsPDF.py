from .KohlsPyMuPDF import extract_metadata_pymupdf, extract_table_rows_pymupdf
from .KohlssPDFPlumber import extract_metadata_pdfplumber, extract_table_rows_pdfplumber


def extract_metadata(pdf_path: str, backend: str = "pdfplumber"):
    if backend == "pdfplumber":
        return extract_metadata_pdfplumber(pdf_path)

    return extract_metadata_pymupdf(pdf_path)


def extract_table_rows(pdf_path: str, backend: str = "pdfplumber"):
    if backend == "pdfplumber":
        return extract_table_rows_pdfplumber(pdf_path)

    return extract_table_rows_pymupdf(pdf_path)
