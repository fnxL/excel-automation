from typing import Union, Annotated, List
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from .macro import create_macro
from .summary import create_summary
import time

origins = [
    "http://localhost",
    "http://localhost:5173",
    "http://localhost:5173/",
    "http://localhost:4173",
    "https://automate-excel.pages.dev",
    "https://xl.fnxl.live",
]

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["*"],
)


@app.get("/")
def read_root():
    return {"Hello": "World"}


@app.get("/customers")
def get_customers():
    return {
        "customers": [
            {"label": "Kohl's Towel PDF", "value": "kohls-towel-pdf"},
            {"label": "Kohl's Towel", "value": "kohls-towel"},
            {"label": "Kohl's Rugs PDF", "value": "kohls-rugs-pdf"},
            {"label": "Kohl's Bedsheet", "value": "kohls-bedsheet"},
            {"label": "Kohl's PO Summary", "value": "kohls-po-summary"},
            {"label": "Walmart Bedsheet", "value": "walmart-bedsheet"},
            {"label": "Argos", "value": "argos"},
        ]
    }


@app.post("/posummary")
async def po_summary(
    poNumbers: Annotated[str, Form()],
    mastersheet: Annotated[UploadFile, File()],
    customer: str,
):
    start_time = time.time()

    poNumbersList = [int(s) for s in poNumbers.splitlines() if s != ""]
    filename, bytes_buffer = await create_summary(
        poNumbers=poNumbersList, mastersheet=mastersheet.file
    )
    print(f"Got final zip buffer: {time.time() - start_time} seconds")
    return StreamingResponse(
        bytes_buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/generate/macro")
async def generate_macro(
    mastersheet: UploadFile,
    customer: Union[str],
    files: list[UploadFile] = File(...),
    backend: str = "pdfplumber",
):
    start_time = time.time()
    filename, zip_buffer = await create_macro(
        customer=customer, mastersheet=mastersheet.file, po_files=files, backend=backend
    )
    print(f"Got final zip buffer: {time.time() - start_time} seconds")

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
