# ============================================================
# SCAN-TO-EXCEL
# PDF -> GEMINI VISION -> STRUCTURED JSON -> EXCEL
# ============================================================
#
# INSTALL:
#
# python -m pip install -U google-genai pymupdf pandas openpyxl pillow
#
#
# WINDOWS POWERSHELL:
#
# $env:GEMINI_API_KEY="YOUR_API_KEY"
#
#
# RUN:
#
# python gemini_llm_sol.py
#
#
# INPUT:
#
# invoice.pdf
#
#
# OUTPUT:
#
# invoice_output.json
# invoice_output.xlsx
# invoice_raw_response.txt
#
# ============================================================


import os
import io
import re
import json
import time
import traceback
from pathlib import Path


import pymupdf
import pandas as pd

from PIL import Image

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from google import genai
from google.genai import types


# ============================================================
# CONFIGURATION
# ============================================================

PDF_FILE = "6436760200-SWB.PDF"

JSON_OUTPUT = "invoice_output.json"

EXCEL_OUTPUT = "invoice_output.xlsx"

RAW_RESPONSE_OUTPUT = "invoice_raw_response.txt"


# ============================================================
# GEMINI MODELS
# ============================================================
#
# We try the newest model first.
#
# If Gemini temporarily returns:
#
#   503
#   429
#   500
#   502
#   504
#
# we retry and then move to the next model.
#
# This prevents one temporary Gemini outage from killing the
# entire Scan-to-Excel pipeline.
#
# ============================================================

GEMINI_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
]


# Number of attempts for each model.
MAX_RETRIES_PER_MODEL = 3


# Initial retry delay in seconds.
#
# Retry delays become:
#
# 5 sec
# 10 sec
# 20 sec
#
INITIAL_RETRY_DELAY = 5


# ============================================================
# PDF RENDERING
# ============================================================

PDF_DPI = 180


# ============================================================
# API KEY
# ============================================================
#
# IMPORTANT:
#
# NEVER hard-code your Gemini API key in this file.
#
# Use:
#
# PowerShell:
#
# $env:GEMINI_API_KEY="YOUR_API_KEY"
#
# ============================================================

API_KEY = os.getenv(
    "GEMINI_API_KEY",
    ""
).strip()


if not API_KEY:

    raise RuntimeError(
        "\n"
        "GEMINI_API_KEY was not found.\n\n"
        "Set it in PowerShell using:\n\n"
        '$env:GEMINI_API_KEY="YOUR_API_KEY"\n\n'
        "Then run the program again.\n"
    )


# ============================================================
# GEMINI CLIENT
# ============================================================

client = genai.Client(
    api_key=API_KEY
)


# ============================================================
# SYSTEM PROMPT
# ============================================================

SYSTEM_PROMPT = r"""
You are an advanced multimodal document-understanding system.

You specialize in:

- invoices
- bills
- receipts
- purchase orders
- statements
- quotations
- tax invoices
- business documents
- financial documents
- scanned documents

You will receive images representing pages of ONE document.

Your task is to understand the COMPLETE visual structure of the
document and convert it into structured JSON.

This is NOT a simple OCR task.

The most important objective is to reconstruct the semantic
structure of the document.

Do NOT assume a predefined invoice table.

Do NOT assume every document has the same columns.

Do NOT force unrelated information into one table.

============================================================
DOCUMENT UNDERSTANDING
============================================================

Analyze:

- text
- spatial position
- alignment
- whitespace
- borders
- visual grouping
- headings
- labels
- repeated rows
- repeated columns
- numerical patterns
- semantic relationships
- page continuation
- merged cells
- multi-line cells

Use these clues to determine:

1. Which information belongs together.
2. Which information is metadata.
3. Which regions are tables.
4. Where rows begin and end.
5. Where columns begin and end.
6. What each column means.
7. Which values belong to which cells.
8. Which regions are totals.
9. Which regions are tax summaries.
10. Which tables continue on later pages.

============================================================
VERY IMPORTANT
============================================================

The invoice structure is UNKNOWN.

Discover the structure from the document.

For example, one invoice may contain:

Description | HSN | Qty | Rate | CGST | SGST | Amount

Another may contain:

Item Code | Description | Unit | Quantity | Price | Total

Another may contain:

Description
Quantity
Unit Price
Taxable Value
CGST
SGST
Total

Another may contain several completely different tables.

Do NOT force all documents into one universal table.

============================================================
TABLE RECONSTRUCTION
============================================================

When identifying a table:

1. Determine its visual boundaries.
2. Determine its logical columns.
3. Determine its rows.
4. Determine the meaning of every column.
5. Preserve multi-line descriptions.
6. Preserve numerical values.
7. Preserve units.
8. Preserve tax information.
9. Do not merge unrelated tables.
10. Do not treat totals as line items.
11. If a table continues onto another page, merge it into
    one logical table when the structure clearly continues.

If the table has no visible header:

- infer column meaning from alignment and semantics
- set "inferred": true
- reduce confidence appropriately

============================================================
OCR
============================================================

Correct obvious OCR mistakes when the visual evidence is strong.

Examples:

"INV0ICE" -> "INVOICE"

"1O0.50" -> "100.50"

"6 30" -> "6:30"

However:

DO NOT silently change uncertain values.

If a value cannot be confidently read:

return:

""

and reduce confidence.

============================================================
NO HALLUCINATION
============================================================

Never invent:

- invoice numbers
- dates
- names
- addresses
- tax numbers
- products
- quantities
- prices
- totals
- bank information
- payment information

Only extract information visible in the document.

============================================================
PAGES
============================================================

Analyze ALL pages.

Every table and important extracted item should contain the
page number where it was found.

If a table continues across multiple pages, preserve the
relationship.

============================================================
CONFIDENCE
============================================================

Every important extracted item should have a confidence between
0 and 1.

Use lower confidence when:

- handwriting is unclear
- OCR is uncertain
- the value is partially obscured
- table structure is ambiguous
- a column meaning is inferred

Do not invent artificially high confidence values.

============================================================
DUPLICATES
============================================================

Avoid unnecessary duplication.

If the same invoice number appears in multiple places, do not
create several different invoice numbers.

============================================================
OUTPUT
============================================================

Return ONLY valid JSON.

Do not return:

- markdown
- ```json
- explanations
- comments
- text outside JSON
"""


# ============================================================
# JSON INSTRUCTION
# ============================================================

JSON_INSTRUCTION = r"""
Return exactly ONE JSON object using this structure:

{
    "document": {
        "type": "",
        "title": "",
        "currency": "",
        "language": "",
        "confidence": 0.0
    },

    "metadata": {
        "invoice_number": "",
        "invoice_date": "",
        "due_date": "",
        "po_number": "",
        "reference_number": "",
        "payment_terms": ""
    },

    "seller": {
        "name": "",
        "address": "",
        "tax_id": "",
        "phone": "",
        "email": ""
    },

    "buyer": {
        "name": "",
        "address": "",
        "tax_id": "",
        "phone": "",
        "email": ""
    },

    "tables": [
        {
            "table_id": "",
            "title": "",
            "pages": [1],
            "confidence": 0.0,

            "columns": [
                {
                    "name": "",
                    "semantic_type": "",
                    "inferred": false,
                    "confidence": 0.0
                }
            ],

            "rows": [
                {
                    "cells": {
                        "Column Name": "value"
                    },
                    "page": 1,
                    "confidence": 0.0
                }
            ]
        }
    ],

    "taxes": [
        {
            "name": "",
            "rate": "",
            "amount": "",
            "page": 1,
            "confidence": 0.0
        }
    ],

    "totals": {
        "subtotal": "",
        "discount": "",
        "shipping": "",
        "tax": "",
        "grand_total": "",
        "amount_paid": "",
        "balance_due": ""
    },

    "other_information": [
        {
            "label": "",
            "value": "",
            "page": 1,
            "confidence": 0.0
        }
    ],

    "pages": [
        {
            "page": 1,
            "summary": ""
        }
    ],

    "validation": {
        "needs_review": false,
        "issues": []
    }
}

Rules:

- Return one top-level object.
- Use arrays where specified.
- Use empty strings for missing values.
- Confidence must be between 0 and 1.
- Table columns must reflect the actual document.
- Do not force unrelated tables into one table.
- Preserve page numbers.
- Preserve row order.
- Do not invent values.
- Return valid JSON only.
"""


# ============================================================
# PDF -> IMAGES
# ============================================================

def pdf_to_images(
    pdf_path,
    dpi=180
):

    pdf_path = Path(pdf_path)

    if not pdf_path.exists():

        raise FileNotFoundError(
            f"\nPDF file not found:\n{pdf_path}\n"
        )


    print()
    print("=" * 60)
    print("READING PDF")
    print("=" * 60)


    document = pymupdf.open(
        pdf_path
    )


    print(
        f"File: {pdf_path}"
    )

    print(
        f"Total pages: {len(document)}"
    )


    pages = []


    scale = dpi / 72.0


    matrix = pymupdf.Matrix(
        scale,
        scale
    )


    for page_number, page in enumerate(
        document,
        start=1
    ):

        print(
            f"  Converting page {page_number}..."
        )


        pixmap = page.get_pixmap(
            matrix=matrix,
            alpha=False
        )


        image_bytes = pixmap.tobytes(
            "png"
        )


        image = Image.open(
            io.BytesIO(image_bytes)
        ).convert(
            "RGB"
        )


        pages.append(
            {
                "page": page_number,
                "image": image
            }
        )


    document.close()


    print(
        "PDF conversion complete."
    )


    return pages


# ============================================================
# TEMPORARY ERROR DETECTION
# ============================================================

def is_temporary_error(
    exception
):

    text = str(
        exception
    ).upper()


    temporary_codes = [

        "503",

        "429",

        "500",

        "502",

        "504",

        "UNAVAILABLE",

        "RESOURCE_EXHAUSTED",

        "DEADLINE_EXCEEDED",

        "INTERNAL",

        "TIMEOUT",
    ]


    return any(
        code in text
        for code in temporary_codes
    )


# ============================================================
# GEMINI DOCUMENT EXTRACTION
# ============================================================

def extract_document(
    pages
):

    print()
    print("=" * 60)
    print("SENDING DOCUMENT TO GEMINI")
    print("=" * 60)


    contents = []


    # --------------------------------------------------------
    # Prompt
    # --------------------------------------------------------

    contents.append(
        SYSTEM_PROMPT
    )


    contents.append(
        JSON_INSTRUCTION
    )


    contents.append(
        """
The following images are pages of ONE document.

Analyze ALL pages together.

First understand the document structure.

Then identify every logical section and table.

Do not assume that all pages have the same layout.

Finally return the complete JSON object.
"""
    )


    # --------------------------------------------------------
    # Images
    # --------------------------------------------------------

    for page in pages:

        contents.append(
            f"\n===== PAGE {page['page']} =====\n"
        )

        contents.append(
            page["image"]
        )


    # --------------------------------------------------------
    # Try models
    # --------------------------------------------------------

    last_exception = None


    for model_index, model_name in enumerate(
        GEMINI_MODELS,
        start=1
    ):

        print()
        print(
            "-" * 60
        )

        print(
            f"MODEL {model_index}/"
            f"{len(GEMINI_MODELS)}"
        )

        print(
            f"Using: {model_name}"
        )

        print(
            "-" * 60
        )


        # ----------------------------------------------------
        # Retry model
        # ----------------------------------------------------

        for attempt in range(
            1,
            MAX_RETRIES_PER_MODEL + 1
        ):

            print(
                f"Attempt "
                f"{attempt}/"
                f"{MAX_RETRIES_PER_MODEL}"
            )


            try:

                # ------------------------------------------------
                # Current Gemini SDK configuration.
                #
                # We deliberately do NOT specify temperature,
                # top_p or top_k because these parameters are
                # deprecated for newer Gemini 3 models.
                # ------------------------------------------------

                config = types.GenerateContentConfig(

                    response_mime_type=(
                        "application/json"
                    )
                )


                response = (
                    client.models.generate_content(

                        model=model_name,

                        contents=contents,

                        config=config
                    )
                )


                # ------------------------------------------------
                # Validate response
                # ------------------------------------------------

                if response is None:

                    raise RuntimeError(
                        "Gemini returned no response."
                    )


                response_text = (
                    getattr(
                        response,
                        "text",
                        None
                    )
                )


                if not response_text:

                    raise RuntimeError(
                        "Gemini returned an empty response."
                    )


                print()
                print(
                    "=" * 60
                )

                print(
                    "GEMINI SUCCESS"
                )

                print(
                    f"Model used: {model_name}"
                )

                print(
                    "=" * 60
                )


                return response_text


            except Exception as exc:

                last_exception = exc


                print()
                print(
                    f"Gemini error:"
                )

                print(
                    str(exc)
                )


                # ------------------------------------------------
                # Temporary error
                # ------------------------------------------------

                if is_temporary_error(
                    exc
                ):

                    if attempt < MAX_RETRIES_PER_MODEL:

                        delay = (
                            INITIAL_RETRY_DELAY
                            * (
                                2 ** (attempt - 1)
                            )
                        )


                        print()
                        print(
                            "Temporary Gemini "
                            "server/quota error."
                        )

                        print(
                            f"Retrying in "
                            f"{delay} seconds..."
                        )


                        time.sleep(
                            delay
                        )


                    else:

                        print()
                        print(
                            f"{model_name} "
                            "failed after all retries."
                        )

                        print(
                            "Trying next model..."
                        )


                else:

                    # ------------------------------------------------
                    # Non-temporary error
                    #
                    # Don't blindly retry invalid API requests.
                    # ------------------------------------------------

                    print()
                    print(
                        "This does not appear to be "
                        "a temporary Gemini error."
                    )

                    raise


    # ========================================================
    # ALL MODELS FAILED
    # ========================================================

    print()
    print("=" * 60)
    print("ALL GEMINI MODELS FAILED")
    print("=" * 60)


    if last_exception:

        print(
            f"Last error: {last_exception}"
        )


    raise RuntimeError(
        "Gemini was unavailable after trying "
        "all configured models."
    ) from last_exception


# ============================================================
# JSON PARSER
# ============================================================

def parse_json(
    text
):

    if not isinstance(
        text,
        str
    ):

        raise RuntimeError(
            "Gemini response is not text."
        )


    text = text.strip()


    # --------------------------------------------------------
    # Remove accidental markdown fences
    # --------------------------------------------------------

    text = re.sub(
        r"^```json\s*",
        "",
        text,
        flags=re.IGNORECASE
    )


    text = re.sub(
        r"^```\s*",
        "",
        text
    )


    text = re.sub(
        r"\s*```$",
        "",
        text
    )


    text = text.strip()


    # --------------------------------------------------------
    # Direct parse
    # --------------------------------------------------------

    try:

        data = json.loads(
            text
        )


    except json.JSONDecodeError:

        # ----------------------------------------------------
        # Try extracting the outer JSON object.
        # ----------------------------------------------------

        start = text.find(
            "{"
        )

        end = text.rfind(
            "}"
        )


        if start == -1 or end == -1:

            with open(
                RAW_RESPONSE_OUTPUT,
                "w",
                encoding="utf-8"
            ) as file:

                file.write(
                    text
                )


            raise RuntimeError(
                "Gemini response does not contain JSON."
            )


        candidate = text[
            start:end + 1
        ]


        try:

            data = json.loads(
                candidate
            )


        except json.JSONDecodeError as exc:

            with open(
                RAW_RESPONSE_OUTPUT,
                "w",
                encoding="utf-8"
            ) as file:

                file.write(
                    text
                )


            print()
            print("=" * 60)
            print("JSON PARSING FAILED")
            print("=" * 60)

            print(
                f"Raw response saved to: "
                f"{RAW_RESPONSE_OUTPUT}"
            )

            print()
            print(
                text[:3000]
            )


            raise RuntimeError(
                "Gemini response was not valid JSON."
            ) from exc


    if not isinstance(
        data,
        dict
    ):

        raise RuntimeError(
            "Gemini JSON must be a JSON object."
        )


    return data


# ============================================================
# STRUCTURE VALIDATION
# ============================================================

def validate_structure(
    data
):

    issues = []


    required_sections = [

        "document",

        "metadata",

        "seller",

        "buyer",

        "tables",

        "taxes",

        "totals",

        "pages",
    ]


    for section in required_sections:

        if section not in data:

            issues.append(
                f"Missing section: {section}"
            )


    if not isinstance(
        data.get(
            "tables",
            []
        ),
        list
    ):

        issues.append(
            "'tables' must be a list."
        )


    if not isinstance(
        data.get(
            "taxes",
            []
        ),
        list
    ):

        issues.append(
            "'taxes' must be a list."
        )


    if not isinstance(
        data.get(
            "pages",
            []
        ),
        list
    ):

        issues.append(
            "'pages' must be a list."
        )


    # --------------------------------------------------------
    # Validate tables
    # --------------------------------------------------------

    for table_index, table in enumerate(
        data.get(
            "tables",
            []
        ),
        start=1
    ):

        if not isinstance(
            table,
            dict
        ):

            issues.append(
                f"Table {table_index} "
                "is not an object."
            )

            continue


        if "columns" not in table:

            issues.append(
                f"Table {table_index} "
                "has no columns."
            )


        if "rows" not in table:

            issues.append(
                f"Table {table_index} "
                "has no rows."
            )


        columns = table.get(
            "columns",
            []
        )


        rows = table.get(
            "rows",
            []
        )


        if not isinstance(
            columns,
            list
        ):

            issues.append(
                f"Table {table_index} "
                "columns are not a list."
            )


        if not isinstance(
            rows,
            list
        ):

            issues.append(
                f"Table {table_index} "
                "rows are not a list."
            )


    return issues


# ============================================================
# NUMBER PARSER
# ============================================================

def parse_number(
    value
):

    if value is None:

        return None


    value = str(
        value
    ).strip()


    if not value:

        return None


    # Currency symbols
    value = re.sub(
        r"[₹$€£]",
        "",
        value
    )


    # Spaces
    value = value.replace(
        " ",
        ""
    )


    # Percentage
    value = value.replace(
        "%",
        ""
    )


    # Commas
    value = value.replace(
        ",",
        ""
    )


    # Parentheses indicate negative
    if (
        value.startswith("(")
        and value.endswith(")")
    ):

        value = (
            "-"
            + value[1:-1]
        )


    try:

        return float(
            value
        )


    except ValueError:

        return None


# ============================================================
# TABLE MATHEMATICAL VALIDATION
# ============================================================

def validate_table_math(
    table
):

    issues = []


    rows = table.get(
        "rows",
        []
    )


    for row_number, row in enumerate(
        rows,
        start=1
    ):

        if not isinstance(
            row,
            dict
        ):

            continue


        cells = row.get(
            "cells",
            {}
        )


        if not isinstance(
            cells,
            dict
        ):

            continue


        quantity = None

        rate = None

        amount = None


        for key, value in cells.items():

            key_lower = str(
                key
            ).lower().strip()


            # ------------------------------------------------
            # Quantity
            # ------------------------------------------------

            if (
                key_lower == "qty"
                or "quantity" in key_lower
            ):

                quantity = parse_number(
                    value
                )


            # ------------------------------------------------
            # Rate
            # ------------------------------------------------

            elif (
                key_lower == "rate"
                or "unit price" in key_lower
                or "unit_price" in key_lower
                or "price" in key_lower
            ):

                rate = parse_number(
                    value
                )


            # ------------------------------------------------
            # Amount
            # ------------------------------------------------

            elif (
                key_lower == "amount"
                or key_lower == "total"
                or "line total" in key_lower
                or "line amount" in key_lower
            ):

                amount = parse_number(
                    value
                )


        # ----------------------------------------------------
        # Check only when all values exist.
        # ----------------------------------------------------

        if (
            quantity is not None
            and rate is not None
            and amount is not None
        ):

            expected = (
                quantity * rate
            )


            tolerance = max(
                1.0,
                abs(expected) * 0.02
            )


            if abs(
                expected - amount
            ) > tolerance:

                issues.append(
                    {
                        "row": row_number,
                        "expected_amount": round(
                            expected,
                            2
                        ),
                        "extracted_amount": amount
                    }
                )


    return issues


# ============================================================
# COMPLETE VALIDATION
# ============================================================

def validate_invoice(
    data
):

    structural_issues = (
        validate_structure(
            data
        )
    )


    math_issues = []


    for table in data.get(
        "tables",
        []
    ):

        table_math_issues = (
            validate_table_math(
                table
            )
        )


        if table_math_issues:

            math_issues.append(
                {
                    "table": table.get(
                        "table_id",
                        "unknown"
                    ),
                    "issues": table_math_issues
                }
            )


    all_issues = []


    all_issues.extend(
        structural_issues
    )


    if math_issues:

        all_issues.append(
            {
                "type": "mathematical_validation",
                "details": math_issues
            }
        )


    if "validation" not in data:

        data["validation"] = {}


    data["validation"]["issues"] = (
        all_issues
    )


    data["validation"]["needs_review"] = (
        len(all_issues) > 0
    )


    return data


# ============================================================
# SAVE JSON
# ============================================================

def save_json(
    data,
    output_path
):

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False
        )


    print()
    print(
        f"JSON saved: {output_path}"
    )


# ============================================================
# SAFE EXCEL SHEET NAME
# ============================================================

def safe_sheet_name(
    name,
    used_names,
    fallback
):

    name = str(
        name or ""
    )


    # Remove illegal Excel characters
    name = re.sub(
        r"[\[\]\*:/\\?]",
        "",
        name
    )


    name = name.strip()


    if not name:

        name = fallback


    name = name[:31]


    original = name

    counter = 2


    while name in used_names:

        suffix = f"_{counter}"


        name = (
            original[
                :31 - len(suffix)
            ]
            + suffix
        )


        counter += 1


    used_names.add(
        name
    )


    return name


# ============================================================
# TABLE -> DATAFRAME
# ============================================================

def table_to_dataframe(
    table
):

    columns_definition = table.get(
        "columns",
        []
    )


    rows = table.get(
        "rows",
        []
    )


    # --------------------------------------------------------
    # First obtain columns from the explicit column definition.
    # --------------------------------------------------------

    column_names = []


    if isinstance(
        columns_definition,
        list
    ):

        for column in columns_definition:

            if not isinstance(
                column,
                dict
            ):

                continue


            name = str(
                column.get(
                    "name",
                    ""
                )
            ).strip()


            if name and name not in column_names:

                column_names.append(
                    name
                )


    # --------------------------------------------------------
    # Also collect any columns found in row cells.
    #
    # This protects us if Gemini returns a cell whose column
    # was not included in the columns array.
    # --------------------------------------------------------

    for row in rows:

        if not isinstance(
            row,
            dict
        ):

            continue


        cells = row.get(
            "cells",
            {}
        )


        if not isinstance(
            cells,
            dict
        ):

            continue


        for column_name in cells.keys():

            column_name = str(
                column_name
            )


            if (
                column_name
                and column_name not in column_names
            ):

                column_names.append(
                    column_name
                )


    # --------------------------------------------------------
    # Build rows according to discovered columns.
    # --------------------------------------------------------

    output_rows = []


    for row in rows:

        if not isinstance(
            row,
            dict
        ):

            continue


        cells = row.get(
            "cells",
            {}
        )


        if not isinstance(
            cells,
            dict
        ):

            continue


        output_row = {}


        for column_name in column_names:

            output_row[column_name] = (
                cells.get(
                    column_name,
                    ""
                )
            )


        # Preserve page number in Excel
        output_row["_page"] = row.get(
            "page",
            ""
        )


        # Preserve confidence in Excel
        output_row["_confidence"] = row.get(
            "confidence",
            ""
        )


        output_rows.append(
            output_row
        )


    if not output_rows:

        return pd.DataFrame()


    return pd.DataFrame(
        output_rows,
        columns=(
            column_names
            + [
                "_page",
                "_confidence"
            ]
        )
    )


# ============================================================
# JSON -> EXCEL
# ============================================================

def json_to_excel(
    data,
    output_path
):

    print()
    print("=" * 60)
    print("CREATING EXCEL")
    print("=" * 60)


    used_sheet_names = set()


    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:


        # ====================================================
        # INVOICE INFORMATION
        # ====================================================

        info_rows = []


        document = data.get(
            "document",
            {}
        )


        if isinstance(
            document,
            dict
        ):

            for key, value in document.items():

                info_rows.append(
                    {
                        "Section": "Document",
                        "Field": key,
                        "Value": value
                    }
                )


        metadata = data.get(
            "metadata",
            {}
        )


        if isinstance(
            metadata,
            dict
        ):

            for key, value in metadata.items():

                info_rows.append(
                    {
                        "Section": "Metadata",
                        "Field": key,
                        "Value": value
                    }
                )


        seller = data.get(
            "seller",
            {}
        )


        if isinstance(
            seller,
            dict
        ):

            for key, value in seller.items():

                info_rows.append(
                    {
                        "Section": "Seller",
                        "Field": key,
                        "Value": value
                    }
                )


        buyer = data.get(
            "buyer",
            {}
        )


        if isinstance(
            buyer,
            dict
        ):

            for key, value in buyer.items():

                info_rows.append(
                    {
                        "Section": "Buyer",
                        "Field": key,
                        "Value": value
                    }
                )


        if info_rows:

            info_df = pd.DataFrame(
                info_rows
            )


            info_df.to_excel(
                writer,
                sheet_name="Invoice Info",
                index=False
            )


            used_sheet_names.add(
                "Invoice Info"
            )


        # ====================================================
        # TABLES
        # ====================================================

        tables = data.get(
            "tables",
            []
        )


        for table_number, table in enumerate(
            tables,
            start=1
        ):

            if not isinstance(
                table,
                dict
            ):

                continue


            dataframe = table_to_dataframe(
                table
            )


            if dataframe.empty:

                continue


            title = (
                table.get(
                    "title"
                )
                or table.get(
                    "table_id"
                )
                or f"Table {table_number}"
            )


            sheet_name = safe_sheet_name(
                title,
                used_sheet_names,
                f"Table {table_number}"
            )


            dataframe.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )


            print(
                f"  Created sheet: {sheet_name}"
            )


        # ====================================================
        # TAXES
        # ====================================================

        taxes = data.get(
            "taxes",
            []
        )


        if taxes:

            tax_df = pd.DataFrame(
                taxes
            )


            tax_df.to_excel(
                writer,
                sheet_name="Taxes",
                index=False
            )


            used_sheet_names.add(
                "Taxes"
            )


        # ====================================================
        # TOTALS
        # ====================================================

        totals = data.get(
            "totals",
            {}
        )


        if isinstance(
            totals,
            dict
        ) and totals:

            total_rows = []


            for key, value in totals.items():

                total_rows.append(
                    {
                        "Field": key,
                        "Value": value
                    }
                )


            totals_df = pd.DataFrame(
                total_rows
            )


            totals_df.to_excel(
                writer,
                sheet_name="Totals",
                index=False
            )


            used_sheet_names.add(
                "Totals"
            )


        # ====================================================
        # OTHER INFORMATION
        # ====================================================

        other_information = data.get(
            "other_information",
            []
        )


        if other_information:

            other_df = pd.DataFrame(
                other_information
            )


            other_df.to_excel(
                writer,
                sheet_name="Other Info",
                index=False
            )


            used_sheet_names.add(
                "Other Info"
            )


        # ====================================================
        # PAGES
        # ====================================================

        pages = data.get(
            "pages",
            []
        )


        if pages:

            pages_df = pd.DataFrame(
                pages
            )


            pages_df.to_excel(
                writer,
                sheet_name="Pages",
                index=False
            )


            used_sheet_names.add(
                "Pages"
            )


        # ====================================================
        # VALIDATION
        # ====================================================

        validation = data.get(
            "validation",
            {}
        )


        validation_rows = []


        validation_rows.append(
            {
                "Field": "Needs Review",
                "Value": validation.get(
                    "needs_review",
                    False
                )
            }
        )


        issues = validation.get(
            "issues",
            []
        )


        if issues:

            for issue_number, issue in enumerate(
                issues,
                start=1
            ):

                validation_rows.append(
                    {
                        "Field": (
                            f"Issue {issue_number}"
                        ),
                        "Value": json.dumps(
                            issue,
                            ensure_ascii=False
                        )
                    }
                )


        validation_df = pd.DataFrame(
            validation_rows
        )


        validation_df.to_excel(
            writer,
            sheet_name="Validation",
            index=False
        )


# ============================================================
# EXCEL FORMATTING
# ============================================================

def format_excel(
    output_path
):

    print(
        "Formatting Excel..."
    )


    workbook = load_workbook(
        output_path
    )


    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )


    header_font = Font(
        color="FFFFFF",
        bold=True
    )


    for worksheet in workbook.worksheets:

        # ----------------------------------------------------
        # Freeze header
        # ----------------------------------------------------

        worksheet.freeze_panes = "A2"


        # ----------------------------------------------------
        # Style header
        # ----------------------------------------------------

        for cell in worksheet[1]:

            cell.fill = header_fill

            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )


        # ----------------------------------------------------
        # Wrap text
        # ----------------------------------------------------

        for row in worksheet.iter_rows():

            for cell in row:

                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )


        # ----------------------------------------------------
        # Automatic column width
        # ----------------------------------------------------

        for column_cells in worksheet.columns:

            maximum_length = 0


            column_letter = (
                column_cells[0]
                .column_letter
            )


            for cell in column_cells:

                if cell.value is not None:

                    maximum_length = max(
                        maximum_length,
                        len(
                            str(
                                cell.value
                            )
                        )
                    )


            worksheet.column_dimensions[
                column_letter
            ].width = min(
                maximum_length + 3,
                60
            )


    workbook.save(
        output_path
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 60)
    print("SCAN-TO-EXCEL")
    print("PDF -> GEMINI -> JSON -> EXCEL")
    print("=" * 60)


    try:

        # ====================================================
        # STEP 1
        # ====================================================

        pages = pdf_to_images(
            PDF_FILE,
            dpi=PDF_DPI
        )


        # ====================================================
        # STEP 2
        # ====================================================

        raw_response = extract_document(
            pages
        )


        # ====================================================
        # STEP 3
        # ====================================================

        data = parse_json(
            raw_response
        )


        # ====================================================
        # STEP 4
        # ====================================================

        data = validate_invoice(
            data
        )


        # ====================================================
        # STEP 5
        # ====================================================

        save_json(
            data,
            JSON_OUTPUT
        )


        # ====================================================
        # STEP 6
        # ====================================================

        json_to_excel(
            data,
            EXCEL_OUTPUT
        )


        # ====================================================
        # STEP 7
        # ====================================================

        format_excel(
            EXCEL_OUTPUT
        )


        # ====================================================
        # SUMMARY
        # ====================================================

        tables = data.get(
            "tables",
            []
        )


        validation = data.get(
            "validation",
            {}
        )


        print()
        print("=" * 60)
        print("PROCESSING COMPLETE")
        print("=" * 60)


        print(
            f"Pages processed : "
            f"{len(pages)}"
        )


        print(
            f"Tables detected : "
            f"{len(tables)}"
        )


        print(
            "Needs review    : "
            f"{validation.get('needs_review', False)}"
        )


        print()
        print(
            f"JSON : {JSON_OUTPUT}"
        )


        print(
            f"Excel: {EXCEL_OUTPUT}"
        )


        if validation.get(
            "issues"
        ):

            print()
            print(
                "WARNING: Validation issues detected:"
            )


            for issue in validation[
                "issues"
            ]:

                print(
                    " -",
                    issue
                )


        else:

            print()
            print(
                "No validation issues detected."
            )


        print()
        print(
            "Done."
        )


    except Exception as exc:

        print()
        print("=" * 60)
        print("PROGRAM FAILED")
        print("=" * 60)


        print(
            f"{type(exc).__name__}: {exc}"
        )


        print()
        print(
            "Full traceback:"
        )


        traceback.print_exc()


        raise


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()