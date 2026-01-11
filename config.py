"""
Configuration for Vietnamese VAT Invoice Extractor
Defines mapping between extracted invoice data and Excel template cells.
"""

# Excel column mapping for header fields (single values)
# Format: "json_field": "excel_column"
HEADER_MAPPING = {
    "invoice_number": "A",      # Số hóa đơn
    "invoice_date": "B",        # Ngày hóa đơn
    "seller_tax_code": "C",     # MST Bên bán
    "seller_name": "D",         # Tên người bán
    "buyer_tax_code": "E",      # MST bên mua
    "buyer_name": "F",          # Tên người mua
}

# Excel column mapping for line items (repeated for each item)
# Format: "json_field": "excel_column"
ITEMS_MAPPING = {
    "item_name": "G",           # Tên hàng hóa
    "unit": "H",                # Đơn vị
    "quantity": "I",            # Số lượng
    "unit_price": "J",          # Đơn giá
    "vat_rate": "K",            # % VAT
    "amount_before_vat": "L",   # Thành tiền trước VAT
    "vat_amount": "M",          # VAT
    "amount_after_vat": "N",    # Thành tiền sau VAT
}

# Row where items start (after header row)
ITEMS_START_ROW = 2

# Expected JSON schema for validation
INVOICE_SCHEMA = {
    "invoice_number": str,
    "invoice_date": str,
    "seller_tax_code": str,
    "seller_name": str,
    "buyer_tax_code": str,
    "buyer_name": str,
    "items": list,
    "total_before_vat": (int, float),
    "total_vat": (int, float),
    "total_after_vat": (int, float),
}

ITEM_SCHEMA = {
    "item_name": str,
    "unit": str,
    "quantity": (int, float),
    "unit_price": (int, float),
    "vat_rate": (int, float),
    "amount_before_vat": (int, float),
    "vat_amount": (int, float),
    "amount_after_vat": (int, float),
}

# Gemini model configuration
GEMINI_MODEL = "gemini-2.5-flash"

# PDF to image conversion settings
PDF_DPI = 200  # Resolution for PDF page rendering
