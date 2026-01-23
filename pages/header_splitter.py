import streamlit as st
import io
import zipfile
import pandas as pd

st.set_page_config(
    page_title="Excel Header Multi-Select Filter",
    page_icon="🗂️"
)

st.title("🗂️ Excel Header Multi-Select Filter")
st.write(
    "Upload Excel files, view their headers, select multiple columns, "
    "and download files that contain ALL selected columns."
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
# Read & display headers
# ----------------------------
file_headers = {}

if uploaded_files:
    st.subheader("Detected Headers")

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, nrows=0)
            headers = list(df.columns)
            file_headers[uploaded_file.name] = headers

            st.markdown(f"**{uploaded_file.name}**")
            st.write(headers)

        except Exception as e:
            st.error(f"❌ Failed to read {uploaded_file.name}: {e}")

# ----------------------------
# Multi-select filter
# ----------------------------
if file_headers:
    all_headers = sorted(
        {col for cols in file_headers.values() for col in cols}
    )

    selected_columns = st.multiselect(
        "Select one or more columns (files must contain ALL selected columns)",
        options=all_headers
    )

    if st.button("Filter & Download"):
        if not selected_columns:
            st.warning("Please select at least one column.")
        else:
            zip_buffer = io.BytesIO()
            processed = 0
            skipped = 0

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for uploaded_file in uploaded_files:
                    try:
                        df = pd.read_excel(uploaded_file, nrows=0)
                        headers = set(df.columns)

                        # Check ALL selected columns exist
                        if not set(selected_columns).issubset(headers):
                            skipped += 1
                            continue

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
                st.info(f"ℹ️ {skipped} file(s) skipped (missing selected columns).")
