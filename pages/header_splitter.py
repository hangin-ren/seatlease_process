import streamlit as st
import io
import zipfile
import pandas as pd

st.set_page_config(
    page_title="Excel Header Viewer & Filter",
    page_icon="🗂️"
)

st.title("🗂️ Excel Header Viewer & Filter")
st.write(
    "Upload Excel files, view their headers, select a column to filter by, and download matching files."
)

# ----------------------------
# Clear / Reset function
# ----------------------------
def clear_all():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# ----------------------------
# File uploader
# ----------------------------
uploaded_files = st.file_uploader(
    "Upload one or more .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True,
    key="file_uploader"
)

st.button("Clear all", on_click=clear_all)

# ----------------------------
# Display headers
# ----------------------------
file_headers = {}  # store headers per file

if uploaded_files:
    st.subheader("Uploaded Files and Headers")
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, nrows=0)
            file_headers[uploaded_file.name] = list(df.columns)
            st.markdown(f"**{uploaded_file.name}**")
            st.write(df.columns.tolist())
        except Exception as e:
            st.error(f"❌ Failed to read {uploaded_file.name}: {e}")

# ----------------------------
# Column filter selection
# ----------------------------
if file_headers:
    all_headers = sorted({col for cols in file_headers.values() for col in cols})
    selected_column = st.selectbox(
        "Select a column to filter files by",
        options=all_headers
    )

    if st.button("Filter & Download"):
        zip_buffer = io.BytesIO()
        processed = 0
        skipped = 0

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for uploaded_file in uploaded_files:
                try:
                    df = pd.read_excel(uploaded_file, nrows=0)
                    if selected_column not in df.columns:
                        skipped += 1
                        continue

                    # Add matching file to ZIP
                    uploaded_file.seek(0)
                    zipf.writestr(uploaded_file.name, uploaded_file.read())
                    processed += 1

                except Exception as e:
                    st.error(f"❌ Failed: {uploaded_file.name} ({e})")

        if processed > 0:
            st.success(f"✅ {processed} file(s) matched the filter.")
            st.download_button(
                "⬇ Download filtered Excel files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="filtered_excels.zip",
                mime="application/zip"
            )

        if skipped > 0:
            st.info(f"ℹ️ {skipped} file(s) skipped (column not found).")
