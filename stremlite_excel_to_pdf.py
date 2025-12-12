import streamlit as st
import win32com.client as win32
import pythoncom
import os
import tempfile
import zipfile
import shutil
import re
from io import BytesIO

st.title("📄 Fast Excel to PDF Converter (No Poppler, No Extra Tools)")
st.write("Upload one or more Excel files. Each sheet will be converted into a separate PDF.")

uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)

# User controls for PDF sizing
scale = st.slider("PDF scale (%)", 50, 200, 100)
fit_to_page = st.checkbox("Fit sheet to single page (may shrink content)", value=False)
orientation_choice = st.selectbox("Orientation", ("Portrait", "Landscape"))
fit_columns_wide = st.checkbox("Fit all columns to page width (shrink if needed)", value=False)
margins = st.slider("Page margins (inches)", 0.0, 1.0, 0.5, step=0.1)
paper_size = st.selectbox("Paper size", ("Letter (8.5x11)", "A4 (210x297mm)", "A3 (297x420mm)"))


def convert_excel_fast(input_file, output_folder, scale=100, fit_to_page=False, orientation_choice="Portrait", fit_columns_wide=False, margins=0.5, paper_size="Letter (8.5x11)"):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Ensure Excel receives an absolute path (COM can fail on relative paths)
    input_file = os.path.abspath(input_file)
    # Ensure output folder is absolute as well
    output_folder = os.path.abspath(output_folder)
    os.makedirs(output_folder, exist_ok=True)
    wb = excel.Workbooks.Open(input_file)

    for sheet in wb.Sheets:
        # sanitize sheet name and limit length to avoid long paths
        def _sanitize(name):
            # replace characters forbidden in filenames on Windows
            name = re.sub(r'[<>:"\\/|?*]', "_", name)
            return name[:120]

        sheet_name = _sanitize(sheet.Name)
        pdf_path = os.path.join(output_folder, f"{sheet_name}.pdf")

        ws = wb.Worksheets(sheet.Name)

        # Paper size
        paper_size_map = {
            "Letter (8.5x11)": 1,  # xlLetterSize
            "A4 (210x297mm)": 5,   # xlA4
            "A3 (297x420mm)": 8    # xlA3
        }
        ws.PageSetup.PaperSize = paper_size_map.get(paper_size, 1)

        # Orientation
        if orientation_choice == "Portrait":
            ws.PageSetup.Orientation = 1
        else:
            ws.PageSetup.Orientation = 2

        # Margins
        ws.PageSetup.LeftMargin = margins
        ws.PageSetup.RightMargin = margins
        ws.PageSetup.TopMargin = margins
        ws.PageSetup.BottomMargin = margins

        # Page setup: either fit-to-page (may shrink) or fit-columns or explicit zoom
        if fit_to_page:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1
        elif fit_columns_wide:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False  # let rows expand as needed
        else:
            ws.PageSetup.FitToPagesWide = False
            ws.PageSetup.FitToPagesTall = False
            ws.PageSetup.Zoom = int(scale)

        # Try normal export; on failure, try exporting to a short temp PDF then move
        try:
            ws.ExportAsFixedFormat(0, pdf_path)
        except Exception as e:
            try:
                tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                tmp_pdf.close()
                ws.ExportAsFixedFormat(0, tmp_pdf.name)
                # move into place (overwrite if necessary)
                try:
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)
                except OSError:
                    pass
                shutil.move(tmp_pdf.name, pdf_path)
            except Exception as e2:
                # cleanup temp and log error to Streamlit
                try:
                    if os.path.exists(tmp_pdf.name):
                        os.remove(tmp_pdf.name)
                except Exception:
                    pass
                st.error(f"Failed to export sheet '{sheet.Name}' to PDF: {e2}")
                continue

    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()



if st.button("Convert to PDF"):
    if not uploaded_files:
        st.error("Please upload at least one Excel file.")
    else:
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for file in uploaded_files:
                st.write(f"📘 Processing: **{file.name}**")

                # Save uploaded file temporarily (use safe absolute temp file)
                _, ext = os.path.splitext(file.name)
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
                try:
                    tmp.write(file.read())
                    tmp.flush()
                finally:
                    tmp.close()
                temp_excel = tmp.name

                # Create temporary output folder
                base_name = os.path.splitext(file.name)[0]
                output_folder = f"output_{base_name}"
                os.makedirs(output_folder, exist_ok=True)

                # Convert with user's sizing options
                convert_excel_fast(temp_excel, output_folder, scale=scale, fit_to_page=fit_to_page, orientation_choice=orientation_choice, fit_columns_wide=fit_columns_wide, margins=margins, paper_size=paper_size)

                # Add generated PDFs to ZIP
                for pdf_file in os.listdir(output_folder):
                    pdf_path = os.path.join(output_folder, pdf_file)
                    zipf.write(pdf_path, f"{base_name}/{pdf_file}")

                # Cleanup temp files
                try:
                    os.remove(temp_excel)
                except OSError:
                    pass
                for p in os.listdir(output_folder):
                    os.remove(os.path.join(output_folder, p))
                os.rmdir(output_folder)

        zip_buffer.seek(0)

        st.success("🎉 Conversion Completed!")
        st.download_button(
            "⬇ Download PDFs (ZIP File)",
            data=zip_buffer,
            file_name="Converted_PDFs.zip",
            mime="application/zip"
        )
