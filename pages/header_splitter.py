import streamlit as st
import io
import zipfile
import pandas as pd

st.set_page_config(
    page_title="Excel Header Filter",
    page_icon="🗂️"
)

st.title("🗂️ Excel Header Filter")
st.write(
    "Upload Excel files, filter them by column header, and download only the matching files."
)

# ----------------------------
# Clear / Reset function
# ----------------------------
def clear_all():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# ----------------------------
# Inputs
# ----------------------------
header_filter = st.text_input(
    "Filter files by column header",
    placeholder="e.g. Debtor, Date, Status",
    key="header_filter"
)

uploaded_files = st.file_uploader(
    "Upload one or more .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True,
    key="file_uploader"
)

# ----------------------------
# Action Buttons
# ----------------------------
col1, col2 = st.columns(2)

with col1:
    process_btn = st.button("Filter files")

with col2:
    st.button("Clear all", on_click=clear_all)

# ----------------------------
# Processing logic
# ----------------------------
if process_btn:
    if not uploaded_files:
        st.warning("Please upload at least one Excel file.")
    elif not header_filter:
        st.warning("Please enter a column header to filter by.")
    else:
        zip_buffer = io.BytesIO()
        processed = 0
        skipped = 0

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for uploaded_file in uploaded_files:
                try:
                    # Read only headers
                    df_headers = pd.read_excel(uploaded_file, nrows=0)
                    if header_filter not in df_headers.columns:
                        skipped += 1
                        continue

                    # Add matching file to ZIP
                    uploaded_file.seek(0)  # reset file pointer
                    zipf.writestr(uploaded_file.name, uploaded_file.read())
                    processed += 1

                except Exception as e:
                    st.error(f"❌ Failed: {uploaded_file.name} ({e})")

        # ----------------------------
        # Results
        # ----------------------------
        if processed > 0:
            st.success(f"✅ {processed} file(s) matched the filter.")
            st.download_button(
                "⬇ Download filtered Excel files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="filtered_excels.zip",
                mime="application/zip"
            )

        if skipped > 0:
            st.info(f"ℹ️ {skipped} file(s) skipped (header not found).")
