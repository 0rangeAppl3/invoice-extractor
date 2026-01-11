# Vietnamese VAT Invoice Extractor

A Streamlit application that extracts structured data from scanned Vietnamese VAT invoice PDFs using Google Gemini AI and exports to Excel.

## Features

- Upload scanned PDF invoices (single PDF at a time)
- Convert PDF pages to images using PyMuPDF
- OCR and structured data extraction using Google Gemini multimodal AI
- Validate and normalize extracted data
- Export to Excel using a configurable mapping
- Support for custom Excel templates

## Requirements

- Python 3.9+
- Google Gemini API key (get one at https://makersuite.google.com/app/apikey)

## Installation

1. Clone or download this project

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Start the Streamlit app:
```bash
streamlit run app.py
```

2. Open your browser to http://localhost:8501

3. Enter your Google Gemini API key in the sidebar

4. Upload a scanned Vietnamese VAT invoice PDF

5. Click "Extract Invoice Data" to process the invoice

6. Review the extracted data

7. Click "Generate Excel" to create the output file

8. Download the Excel file

## Project Structure

```
invoice-extractor/
├── app.py                 # Main Streamlit application
├── config.py              # Configuration and field mappings
├── invoice_extractor.py   # PDF processing and Gemini extraction
├── excel_writer.py        # Excel template filling
├── requirements.txt       # Python dependencies
├── README.md              # This file
└── sample/
    ├── invoice.pdf        # Sample invoice for testing
    └── extracted.xlsx     # Sample output template
```

## Configuration

The `config.py` file contains configurable mappings:

### Header Mapping
Maps invoice header fields to Excel columns:
```python
HEADER_MAPPING = {
    "invoice_number": "A",      # Số hóa đơn
    "invoice_date": "B",        # Ngày hóa đơn
    "seller_tax_code": "C",     # MST Bên bán
    "seller_name": "D",         # Tên người bán
    "buyer_tax_code": "E",      # MST bên mua
    "buyer_name": "F",          # Tên người mua
}
```

### Items Mapping
Maps line item fields to Excel columns:
```python
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
```

### Items Start Row
The row where line items begin (default: 2, after header row):
```python
ITEMS_START_ROW = 2
```

## Extracted Data Structure

The app extracts the following fields from invoices:

**Invoice Header:**
- Invoice number and series
- Invoice date
- MCCQT code
- Seller information (name, tax code, address, phone, bank)
- Buyer information (name, tax code, address)
- Payment method and currency

**Line Items:**
- Item name
- Unit of measure
- Quantity
- Unit price
- Discount (if any)
- VAT rate
- Amount before VAT
- VAT amount
- Amount after VAT

**Totals:**
- Total before VAT
- Total VAT
- Total after VAT
- Total in words

## Error Handling

The app includes robust error handling for:
- Invalid PDF files
- PDF conversion failures
- Gemini API errors
- JSON parsing errors
- Data validation issues
- Excel generation errors

## Limitations

- Single PDF processing only (no batch mode)
- No authentication or database storage
- Requires active internet connection for Gemini API
- Best results with clear, high-quality scanned invoices

## License

MIT License
