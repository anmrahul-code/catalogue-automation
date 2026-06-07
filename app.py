import streamlit as st
import pandas as pd
import io
st.set_page_config(
    page_title="Catalogue Automation Portal",
    layout="wide"
)

st.title("📦 Catalogue Automation Portal")
master_file = st.file_uploader(
    "Upload Master File",
    type=["xlsx"]
)
marketplace = st.selectbox(
    "Select Marketplace",
    [
        "Flipkart",
        "Myntra",
        "Ajio"
    ]
)
instruction_file = (
    f"config/{marketplace}_instruction.xlsx"
)

dropdown_file = (
    f"config/{marketplace}_dropdown.xlsx"
)
try:

    instruction_excel = pd.ExcelFile(
        instruction_file
    )

    categories = [
        sheet.replace("Mapping-", "")
        for sheet in instruction_excel.sheet_names
        if sheet.startswith("Mapping-")
    ]

except Exception as e:

    st.error(str(e))
    st.stop()
    selected_category = st.selectbox(
    "Select Category",
    categories
)
    if st.button("Generate Template"):
            if master_file is None:
                st.warning(
            "Please upload Master File"
        )

        st.stop()
    try:
