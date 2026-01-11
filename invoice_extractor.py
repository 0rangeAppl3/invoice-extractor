"""
Invoice Extractor Module
Uses PyMuPDF to convert PDF pages to images and Gemini for OCR + structured extraction.
Supports batch processing for multiple invoices.
"""

import io
import json
import base64
import time
from typing import Optional, Callable

import fitz  # PyMuPDF
from google import genai
from google.genai import types
from PIL import Image

from config import GEMINI_MODEL, PDF_DPI, INVOICE_SCHEMA, ITEM_SCHEMA


EXTRACTION_PROMPT = """You are an expert at extracting data from Vietnamese VAT invoices (Hóa đơn giá trị gia tăng).

Analyze the invoice image(s) and extract ALL information into a structured JSON format.

IMPORTANT RULES:
1. Extract ALL line items from the invoice table
2. Numbers should be extracted as numeric values (not strings), removing thousand separators (dots in Vietnamese format)
3. VAT rate should be a percentage number (e.g., 5 for 5%, 10 for 10%)
4. Dates should be in format "DD/MM/YYYY"
5. If a field is empty or not found, use null
6. Calculate vat_amount and amount_after_vat for each item if not explicitly shown

Required JSON structure:
{
    "invoice_number": "string - Số hóa đơn (e.g., 94)",
    "invoice_series": "string - Ký hiệu (e.g., C25TDX)",
    "invoice_date": "string - Ngày hóa đơn in DD/MM/YYYY format",
    "mccqt": "string - Mã MCCQT code if present",

    "seller_name": "string - Tên người bán",
    "seller_tax_code": "string - MST bên bán",
    "seller_address": "string - Địa chỉ người bán",
    "seller_phone": "string - Điện thoại người bán",
    "seller_bank_account": "string - Số tài khoản người bán",
    "seller_bank_name": "string - Tên ngân hàng người bán",

    "buyer_name": "string - Tên người mua",
    "buyer_tax_code": "string - MST bên mua",
    "buyer_address": "string - Địa chỉ người mua",
    "payment_method": "string - Hình thức thanh toán",
    "currency": "string - Đơn vị tiền tệ (e.g., VND)",

    "items": [
        {
            "stt": "number - line item number",
            "item_name": "string - Tên hàng hóa, dịch vụ",
            "unit": "string - Đơn vị tính",
            "quantity": "number - Số lượng",
            "unit_price": "number - Đơn giá",
            "discount": "number or null - Chiết khấu",
            "vat_rate": "number - Thuế suất as percentage (5 for 5%)",
            "amount_before_vat": "number - Thành tiền chưa có thuế GTGT",
            "vat_amount": "number - Tiền thuế GTGT for this item",
            "amount_after_vat": "number - Thành tiền sau thuế for this item"
        }
    ],

    "total_before_vat": "number - Tổng tiền chưa thuế",
    "total_vat": "number - Tổng tiền thuế",
    "total_fee": "number or null - Tổng tiền phí",
    "total_discount": "number or null - Tổng tiền chiết khấu",
    "total_after_vat": "number - Tổng tiền thanh toán",
    "total_in_words": "string - Tổng tiền thanh toán bằng chữ"
}

Return ONLY valid JSON, no markdown code blocks or extra text.
"""


def convert_pdf_to_images(pdf_bytes: bytes) -> list[Image.Image]:
    """Convert PDF pages to PIL Images using PyMuPDF."""
    images = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Render page to image with specified DPI
            mat = fitz.Matrix(PDF_DPI / 72, PDF_DPI / 72)
            pix = page.get_pixmap(matrix=mat)

            # Convert to PIL Image
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            images.append(img)
        doc.close()
    except Exception as e:
        raise RuntimeError(f"Failed to convert PDF to images: {str(e)}")

    if not images:
        raise ValueError("No pages found in PDF")

    return images


def image_to_base64(img: Image.Image) -> str:
    """Convert PIL Image to base64 string."""
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def extract_invoice_data(images: list[Image.Image], api_key: str) -> dict:
    """Use Gemini to extract structured data from invoice images."""
    client = genai.Client(api_key=api_key)

    # Prepare content with all images
    content = [EXTRACTION_PROMPT]
    for img in images:
        # Convert PIL Image to bytes for the new API
        img_buffer = io.BytesIO()
        img.save(img_buffer, format="PNG")
        img_bytes = img_buffer.getvalue()
        content.append(types.Part.from_bytes(data=img_bytes, mime_type="image/png"))

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=content
        )
        response_text = response.text.strip()

        # Clean up response if wrapped in markdown
        if response_text.startswith("```"):
            lines = response_text.split("\n")
            # Remove first and last lines (```json and ```)
            response_text = "\n".join(lines[1:-1])

        # Parse JSON
        invoice_data = json.loads(response_text)
        return invoice_data

    except json.JSONDecodeError as e:
        raise ValueError(f"Failed to parse Gemini response as JSON: {str(e)}\nResponse: {response_text[:500]}")
    except Exception as e:
        raise RuntimeError(f"Gemini API error: {str(e)}")


def validate_invoice_data(data: dict) -> tuple[bool, list[str]]:
    """Validate extracted invoice data against expected schema."""
    errors = []

    # Check required header fields
    for field, expected_type in INVOICE_SCHEMA.items():
        if field not in data:
            errors.append(f"Missing required field: {field}")
        elif data[field] is not None:
            if isinstance(expected_type, tuple):
                if not isinstance(data[field], expected_type):
                    errors.append(f"Field '{field}' has wrong type: expected {expected_type}, got {type(data[field])}")
            elif not isinstance(data[field], expected_type):
                errors.append(f"Field '{field}' has wrong type: expected {expected_type}, got {type(data[field])}")

    # Check items
    if "items" in data and isinstance(data["items"], list):
        for i, item in enumerate(data["items"]):
            for field, expected_type in ITEM_SCHEMA.items():
                if field not in item:
                    errors.append(f"Item {i+1} missing field: {field}")
                elif item[field] is not None:
                    if isinstance(expected_type, tuple):
                        if not isinstance(item[field], expected_type):
                            errors.append(f"Item {i+1} field '{field}' has wrong type")
                    elif not isinstance(item[field], expected_type):
                        errors.append(f"Item {i+1} field '{field}' has wrong type")

    return len(errors) == 0, errors


def normalize_invoice_data(data: dict) -> dict:
    """Normalize and clean up extracted invoice data."""
    normalized = data.copy()

    # Ensure items list exists
    if "items" not in normalized or normalized["items"] is None:
        normalized["items"] = []

    # Normalize each item
    for item in normalized["items"]:
        # Calculate vat_amount if missing
        if item.get("vat_amount") is None and item.get("amount_before_vat") and item.get("vat_rate"):
            item["vat_amount"] = item["amount_before_vat"] * item["vat_rate"] / 100

        # Calculate amount_after_vat if missing
        if item.get("amount_after_vat") is None and item.get("amount_before_vat") and item.get("vat_amount"):
            item["amount_after_vat"] = item["amount_before_vat"] + item["vat_amount"]

        # Ensure numeric fields are numbers
        for field in ["quantity", "unit_price", "vat_rate", "amount_before_vat", "vat_amount", "amount_after_vat"]:
            if field in item and item[field] is not None:
                try:
                    if isinstance(item[field], str):
                        # Remove thousand separators and convert
                        item[field] = float(item[field].replace(".", "").replace(",", "."))
                except (ValueError, AttributeError):
                    pass

    # Normalize totals
    for field in ["total_before_vat", "total_vat", "total_after_vat"]:
        if field in normalized and normalized[field] is not None:
            try:
                if isinstance(normalized[field], str):
                    normalized[field] = float(normalized[field].replace(".", "").replace(",", "."))
            except (ValueError, AttributeError):
                pass

    return normalized


def process_invoice(pdf_bytes: bytes, api_key: str) -> tuple[dict, list[Image.Image]]:
    """
    Main function to process an invoice PDF.
    Returns extracted data and images for preview.
    """
    # Convert PDF to images
    images = convert_pdf_to_images(pdf_bytes)

    # Extract data using Gemini
    raw_data = extract_invoice_data(images, api_key)

    # Normalize data
    normalized_data = normalize_invoice_data(raw_data)

    # Validate
    is_valid, errors = validate_invoice_data(normalized_data)
    if errors:
        normalized_data["_validation_warnings"] = errors

    return normalized_data, images


def process_invoices_batch(
    pdf_files: list[tuple[str, bytes]],
    api_key: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    batch_delay: float = 0.5
) -> list[dict]:
    """
    Process multiple invoice PDFs in batch mode.
    Uses sequential processing with caching to optimize API usage.

    Args:
        pdf_files: List of tuples (filename, pdf_bytes)
        api_key: Google Gemini API key
        progress_callback: Optional callback(current, total, filename) for progress updates
        batch_delay: Delay between API calls to avoid rate limiting

    Returns:
        List of result dicts with keys: filename, success, data, images, error
    """
    results = []
    total = len(pdf_files)

    # Create client once for all requests (token saving)
    client = genai.Client(api_key=api_key)

    for idx, (filename, pdf_bytes) in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx + 1, total, filename)

        result = {
            "filename": filename,
            "success": False,
            "data": None,
            "images": None,
            "error": None
        }

        try:
            # Convert PDF to images
            images = convert_pdf_to_images(pdf_bytes)
            result["images"] = images

            # Prepare content for Gemini
            content = [EXTRACTION_PROMPT]
            for img in images:
                # Convert PIL Image to bytes for the new API
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="PNG")
                img_bytes = img_buffer.getvalue()
                content.append(types.Part.from_bytes(data=img_bytes, mime_type="image/png"))

            # Call Gemini API
            response = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=content
            )
            response_text = response.text.strip()

            # Clean up response if wrapped in markdown
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                response_text = "\n".join(lines[1:-1])

            # Parse and normalize
            raw_data = json.loads(response_text)
            normalized_data = normalize_invoice_data(raw_data)

            # Validate
            is_valid, errors = validate_invoice_data(normalized_data)
            if errors:
                normalized_data["_validation_warnings"] = errors

            result["data"] = normalized_data
            result["success"] = True

        except json.JSONDecodeError as e:
            result["error"] = f"Failed to parse response as JSON: {str(e)}"
        except Exception as e:
            result["error"] = str(e)

        results.append(result)

        # Add delay between requests to avoid rate limiting (except for last item)
        if idx < total - 1 and batch_delay > 0:
            time.sleep(batch_delay)

    return results
