import streamlit as st
import msoffcrypto
import io
import zipfile

st.set_page_config(page_title="Excel Password Remover")

st.title("🔓 Excel Password Remover")
st.write("Upload password-protected Excel files and download them without a password.")

password = st.text_input("Enter Excel password", type="password")

uploaded_files = st.file_uploader(
    "Upload one or more .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files and password:
    if st.button("Remove password"):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for uploaded_file in uploaded_files:
                try:
                    office_file = msoffcrypto.OfficeFile(uploaded_file)
                    office_file.load_key(password=password)

                    decrypted = io.BytesIO()
                    office_file.decrypt(decrypted)

                    zipf.writestr(uploaded_file.name, decrypted.getvalue())

                except Exception as e:
                    st.error(f"Failed: {uploaded_file.name} ({e})")

        st.success("Done! Download your files below 👇")

        st.download_button(
            label="Download unprotected Excel files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="unprotected_excels.zip",
            mime="application/zip"
        )