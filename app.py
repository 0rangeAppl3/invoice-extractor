"""
Vietnamese VAT Invoice Extractor
Streamlit app for extracting data from scanned invoice PDFs and exporting to Excel.
Supports multiple file upload and batch processing.
"""

import streamlit as st
from datetime import datetime

from invoice_extractor import process_invoices_batch, convert_pdf_to_images
from excel_writer import fill_excel_batch, format_invoice_number


# Page configuration
st.set_page_config(
    page_title="Vietnamese VAT Invoice Extractor",
    page_icon="📄",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .file-status {
        padding: 0.5rem;
        border-radius: 0.25rem;
        margin: 0.25rem 0;
    }
    .status-success {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
    }
    .status-error {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
    }
    .status-pending {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state variables."""
    if "extracted_results" not in st.session_state:
        st.session_state.extracted_results = []
    if "processing_complete" not in st.session_state:
        st.session_state.processing_complete = False
    if "template_bytes" not in st.session_state:
        st.session_state.template_bytes = None


def display_invoice_summary(data: dict, filename: str):
    """Display a compact summary of extracted invoice data."""
    with st.expander(f"📄 {filename}", expanded=False):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Invoice Details**")
            inv_num = format_invoice_number(data.get('invoice_series'), data.get('invoice_number'))
            st.write(f"- Number: {inv_num}")
            st.write(f"- Date: {data.get('invoice_date', 'N/A')}")

            st.markdown("**Seller**")
            st.write(f"- {data.get('seller_name', 'N/A')}")
            st.write(f"- Tax: {data.get('seller_tax_code', 'N/A')}")

        with col2:
            st.markdown("**Totals**")
            st.write(f"- Before VAT: {data.get('total_before_vat', 0):,.0f} VND")
            st.write(f"- VAT: {data.get('total_vat', 0):,.0f} VND")
            st.write(f"- **Total: {data.get('total_after_vat', 0):,.0f} VND**")

            st.markdown("**Buyer**")
            st.write(f"- {data.get('buyer_name', 'N/A')}")
            st.write(f"- Tax: {data.get('buyer_tax_code', 'N/A')}")

        # Line items
        items = data.get("items", [])
        if items:
            st.markdown("**Line Items**")
            table_data = []
            for item in items:
                table_data.append({
                    "Item": item.get("item_name", "")[:40],
                    "Qty": item.get("quantity", 0),
                    "Price": f"{item.get('unit_price', 0):,.0f}",
                    "VAT%": f"{item.get('vat_rate', 0)}%",
                    "Total": f"{item.get('amount_after_vat', 0):,.0f}",
                })
            st.table(table_data)

        # Validation warnings
        if "_validation_warnings" in data and data["_validation_warnings"]:
            st.warning(f"Warnings: {len(data['_validation_warnings'])} validation issues")


def display_results_summary(results: list):
    """Display summary of all processing results."""
    success_count = sum(1 for r in results if r["success"])
    error_count = len(results) - success_count

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Files", len(results))
    with col2:
        st.metric("Successful", success_count)
    with col3:
        st.metric("Failed", error_count)

    # Show each result
    for result in results:
        if result["success"]:
            display_invoice_summary(result["data"], result["filename"])
        else:
            st.markdown(
                f'<div class="file-status status-error">'
                f'<strong>{result["filename"]}</strong>: {result["error"]}'
                f'</div>',
                unsafe_allow_html=True
            )


def main():
    """Main application entry point."""
    init_session_state()

    # Header
    st.markdown('<div class="main-header">Vietnamese VAT Invoice Extractor</div>', unsafe_allow_html=True)
    st.markdown("Upload scanned Vietnamese VAT invoice PDFs to extract data and export to Excel.")

    # Sidebar for settings
    with st.sidebar:
        st.header("Settings")

        # API Key
        api_key = st.text_input(
            "Google Gemini API Key",
            type="password",
            help="Enter your Google Gemini API key. Get one at https://makersuite.google.com/app/apikey"
        )

        st.markdown("---")

        # Custom template option
        st.subheader("Excel Template")
        use_custom_template = st.checkbox(
            "Use custom template",
            help="Upload your own Excel template with headers in row 1"
        )

        if use_custom_template:
            template_file = st.file_uploader(
                "Upload Template (.xlsx)",
                type=["xlsx"],
                help="Excel file with column headers matching the expected format"
            )
            if template_file:
                st.session_state.template_bytes = template_file.read()
                st.success("Template loaded!")
            elif st.session_state.template_bytes:
                st.info("Using previously loaded template")
        else:
            st.session_state.template_bytes = None

        st.markdown("---")

        # Processing options
        st.subheader("Processing Options")
        batch_delay = st.slider(
            "Delay between files (sec)",
            min_value=0.0,
            max_value=2.0,
            value=0.5,
            step=0.1,
            help="Delay between API calls to avoid rate limiting"
        )

        st.markdown("---")
        st.markdown("**About**")
        st.markdown("""
        Extract data from Vietnamese VAT invoices using:
        - PyMuPDF for PDF processing
        - Google Gemini for OCR & extraction
        - openpyxl for Excel output
        """)

    # Main content area
    if not api_key:
        st.warning("Please enter your Gemini API key in the sidebar to continue.")
        return

    # File uploader - multiple files
    uploaded_files = st.file_uploader(
        "Upload Invoice PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload one or more scanned Vietnamese VAT invoice PDFs"
    )

    if uploaded_files:
        st.info(f"**{len(uploaded_files)} file(s) selected**")

        # Preview section
        with st.expander("Preview Files", expanded=False):
            cols = st.columns(min(len(uploaded_files), 3))
            for idx, uploaded_file in enumerate(uploaded_files[:6]):  # Show max 6 previews
                col_idx = idx % 3
                with cols[col_idx]:
                    st.markdown(f"**{uploaded_file.name}**")
                    try:
                        pdf_bytes = uploaded_file.read()
                        uploaded_file.seek(0)  # Reset for later use
                        images = convert_pdf_to_images(pdf_bytes)
                        st.image(images[0], caption=f"Page 1", use_container_width=True)
                    except Exception as e:
                        st.error(f"Preview error: {str(e)[:50]}")

            if len(uploaded_files) > 6:
                st.info(f"... and {len(uploaded_files) - 6} more files")

        # Process button
        if st.button("Extract All Invoices", type="primary", use_container_width=True):
            # Prepare files
            pdf_files = []
            for uploaded_file in uploaded_files:
                pdf_bytes = uploaded_file.read()
                pdf_files.append((uploaded_file.name, pdf_bytes))

            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(current, total, filename):
                progress = current / total
                progress_bar.progress(progress)
                status_text.text(f"Processing {current}/{total}: {filename}")

            # Process batch
            with st.spinner("Processing invoices..."):
                try:
                    results = process_invoices_batch(
                        pdf_files,
                        api_key,
                        progress_callback=update_progress,
                        batch_delay=batch_delay
                    )
                    st.session_state.extracted_results = results
                    st.session_state.processing_complete = True

                    progress_bar.progress(1.0)
                    status_text.text("Processing complete!")

                except Exception as e:
                    st.error(f"Batch processing failed: {str(e)}")
                    st.session_state.processing_complete = False

        # Display results
        if st.session_state.extracted_results and st.session_state.processing_complete:
            st.markdown("---")
            st.subheader("Extraction Results")
            display_results_summary(st.session_state.extracted_results)

    # Export section
    st.markdown("---")
    if st.session_state.extracted_results and st.session_state.processing_complete:
        st.subheader("Export to Excel")

        # Get successful results only
        successful_results = [r for r in st.session_state.extracted_results if r["success"]]

        if successful_results:
            st.info(f"**{len(successful_results)} invoice(s)** ready for export")

            if st.button("Generate Excel", type="secondary", use_container_width=True):
                try:
                    # Prepare data with formatted invoice numbers
                    export_data = []
                    for result in successful_results:
                        data = result["data"].copy()
                        data["invoice_number"] = format_invoice_number(
                            data.get("invoice_series"),
                            data.get("invoice_number")
                        )
                        export_data.append(data)

                    # Generate Excel with batch function
                    excel_bytes = fill_excel_batch(
                        export_data,
                        st.session_state.template_bytes
                    )

                    # Create filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"invoices_batch_{timestamp}.xlsx"

                    st.download_button(
                        label="Download Excel File",
                        data=excel_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.success(f"Excel file generated with {len(export_data)} invoice(s)!")

                except Exception as e:
                    st.error(f"Error generating Excel: {str(e)}")
        else:
            st.warning("No successfully extracted invoices to export.")

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #888;'>"
        "Vietnamese VAT Invoice Extractor | Built with Streamlit & Gemini AI"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
