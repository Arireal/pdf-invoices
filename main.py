import streamlit as st
import pandas as pd
from fpdf import FPDF
from pathlib import Path
import os
from datetime import datetime
import zipfile
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Invoice to PDF Converter",
    page_icon="📄",
    layout="wide"
)

# Title
st.title("📄 Excel to PDF Invoice Converter")
st.markdown("---")

# Sidebar for settings
with st.sidebar:
    st.header("⚙️ Settings")
    company_name = st.text_input("Company Name", value="PythonHow")
    logo_file = st.file_uploader("Upload Company Logo (optional)",
                                 type=['png', 'jpg', 'jpeg'])
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("""
    1. Upload your Excel invoice files
    2. Configure settings if needed
    3. Click 'Convert to PDF'
    4. Download individual PDFs or all as ZIP
    """)

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📤 Upload Excel Files")
    uploaded_files = st.file_uploader(
        "Choose Excel files (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload one or more Excel invoice files"
    )

with col2:
    st.header("📊 Summary")
    if uploaded_files:
        st.metric("Files Uploaded", len(uploaded_files))
    else:
        st.info("No files uploaded yet")

st.markdown("---")


def create_pdf_from_excel(file, company_name, logo_data=None):
    """Convert Excel file to PDF and return as bytes"""
    try:
        # Get filename without extension
        filename = Path(file.name).stem

        # Split filename to get invoice number and date
        parts = filename.split("-")
        if len(parts) >= 2:
            invoice_nr = parts[0]
            date = parts[1]
        else:
            invoice_nr = filename
            date = datetime.now().strftime("%Y%m%d")

        # Create PDF
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.add_page()

        # Add invoice header
        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

        # Read Excel file
        df = pd.read_excel(file, sheet_name="Sheet 1")

        # Add table header
        columns = df.columns
        columns = [item.replace("_", " ").title() for item in columns]
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=columns[0], border=1)
        pdf.cell(w=70, h=8, txt=columns[1], border=1)
        pdf.cell(w=30, h=8, txt=columns[2], border=1)
        pdf.cell(w=30, h=8, txt=columns[3], border=1)
        pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

        # Add table rows
        for index, row in df.iterrows():
            pdf.set_font(family="Times", size=10)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
            pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
            pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
            pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
            pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

        # Add total row
        total_sum = df["total_price"].sum()
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=70, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt="", border=1)
        pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

        # Add total sum text
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

        # Add company name and logo
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=25, h=8, txt=company_name)

        # Add logo if provided
        if logo_data is not None:
            # Save logo temporarily
            logo_path = f"temp_logo_{datetime.now().timestamp()}.png"
            with open(logo_path, "wb") as f:
                f.write(logo_data)
            pdf.image(logo_path, w=10)
            # Clean up temp file
            os.remove(logo_path)

        # Return PDF as bytes
        return pdf.output(dest='S').encode('latin-1'), filename

    except Exception as e:
        raise Exception(f"Error processing {file.name}: {str(e)}")


# Convert button and processing
if uploaded_files:
    if st.button("🔄 Convert to PDF", type="primary", use_container_width=True):

        # Get logo data if uploaded
        logo_data = None
        if logo_file is not None:
            logo_data = logo_file.read()
            logo_file.seek(0)  # Reset file pointer

        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Store PDFs
        pdf_files = {}
        errors = []

        # Process each file
        for idx, file in enumerate(uploaded_files):
            try:
                status_text.text(f"Processing: {file.name}")
                pdf_bytes, filename = create_pdf_from_excel(file, company_name,
                                                            logo_data)
                pdf_files[f"{filename}.pdf"] = pdf_bytes

            except Exception as e:
                errors.append(f"❌ {file.name}: {str(e)}")

            # Update progress
            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.empty()
        progress_bar.empty()

        # Show results
        st.markdown("---")
        st.header("✅ Conversion Complete!")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Files", len(uploaded_files))
        with col2:
            st.metric("Successful", len(pdf_files))
        with col3:
            st.metric("Errors", len(errors))

        # Show errors if any
        if errors:
            st.error("**Errors occurred:**")
            for error in errors:
                st.write(error)

        # Download section
        if pdf_files:
            st.markdown("---")
            st.subheader("📥 Download PDFs")

            # Individual downloads
            st.write("**Download Individual Files:**")
            cols = st.columns(3)
            for idx, (pdf_name, pdf_data) in enumerate(pdf_files.items()):
                with cols[idx % 3]:
                    st.download_button(
                        label=f"📄 {pdf_name}",
                        data=pdf_data,
                        file_name=pdf_name,
                        mime="application/pdf"
                    )

            # Download all as ZIP
            if len(pdf_files) > 1:
                st.markdown("---")
                st.write("**Download All as ZIP:**")

                # Create ZIP file
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w',
                                     zipfile.ZIP_DEFLATED) as zip_file:
                    for pdf_name, pdf_data in pdf_files.items():
                        zip_file.writestr(pdf_name, pdf_data)

                zip_buffer.seek(0)

                st.download_button(
                    label="📦 Download All PDFs (ZIP)",
                    data=zip_buffer,
                    file_name=f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    type="primary"
                )

else:
    st.info("👆 Please upload Excel files to get started")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    <small>Invoice to PDF Converter | Built with Streamlit</small>
    </div>
    """,
    unsafe_allow_html=True
)
