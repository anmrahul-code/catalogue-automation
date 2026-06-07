import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Catalogue Automation Portal",
    layout="wide"
)

st.title("📦 Catalogue Automation Portal")

# =========================
# Upload Master File
# =========================

master_file = st.file_uploader(
    "Upload Master File",
    type=["xlsx"]
)

# =========================
# Marketplace Selection
# =========================

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

# =========================
# Load Categories
# =========================

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

    st.error(
        f"Instruction file error: {e}"
    )
    st.stop()

selected_category = st.selectbox(
    "Select Category",
    categories
)
# =========================
# Generate Template
# =========================

if st.button("Generate Template"):

    if master_file is None:

        st.warning(
            "Please upload Master File"
        )
        st.stop()

    try:

        # =====================
        # Load Master File
        # =====================

        master_df = pd.read_excel(
            master_file
        )

        # =====================
        # Load Mapping Sheet
        # =====================

        mapping_sheet = (
            f"Mapping-{selected_category}"
        )

        instruction_df = pd.read_excel(
            instruction_file,
            sheet_name=mapping_sheet
        )

        # =====================
        # Load Dropdown Sheet
        # =====================

        dropdown_sheet = (
            f"Dropdown-{selected_category}"
        )

        dropdown_df = pd.read_excel(
            dropdown_file,
            sheet_name=dropdown_sheet
        )

        # =====================
        # Debug
        # =====================

        st.success(
            "Files Loaded Successfully"
        )

        st.write(
            "Master File Rows",
            len(master_df)
        )

        st.write(
            "Instruction Rows",
            len(instruction_df)
        )

        st.write(
            "Dropdown Rows",
            len(dropdown_df)
        )

        st.write(
            "Instruction Preview"
        )

        st.dataframe(
            instruction_df.head()
        )

        st.write(
            "Dropdown Preview"
        )

        st.dataframe(
            dropdown_df.head()
        )

    except Exception as e:

        st.error(
            f"Error : {str(e)}"
        )
# =====================
# Generate Output
# =====================

output_df = pd.DataFrame()

template_column = f"{marketplace} Template Column"

for _, row in instruction_df.iterrows():

    base_col = str(
        row["Base file Column"]
    ).strip()

    output_col = str(
        row[template_column]
    ).strip()

    remarks = str(
        row.get("Remarks", "")
    ).strip().lower()

    if output_col == "" or output_col.lower() == "nan":
        continue

    # =====================
    # Dropdown Values
    # =====================

    if "dropdown" in remarks:

        key = (
            output_col.upper(),
            selected_category.upper()
        )

        mapped_value = dropdown_dict.get(
            key,
            ""
        )

        output_df[output_col] = [
            mapped_value
        ] * len(category_df)

# =====================
# Direct Master Mapping
# =====================

if (
    base_col != "None"
    and base_col in category_df.columns
):

    output_df[output_col] = (
        category_df[base_col]
        .fillna("")
        .astype(str)
        .values
    )

else:

    output_df[output_col] = ""
