import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Catalogue Automation Portal", layout="wide")

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
    ["Myntra", "Ajio"]
)

instruction_file = f"config/{marketplace}_instruction.xlsx"
dropdown_file = f"config/{marketplace}_dropdown.xlsx"

# =========================
# Load Categories
# =========================
try:
    instruction_excel = pd.ExcelFile(instruction_file)

    categories = [
        sheet.replace("Mapping-", "")
        for sheet in instruction_excel.sheet_names
        if sheet.startswith("Mapping-")
    ]

except Exception as e:
    st.error(f"Instruction file error: {e}")
    st.stop()

selected_category = st.selectbox(
    "Select Category",
    categories
)

# =========================
# Generate Button
# =========================
if st.button("Generate Template"):

    if master_file is None:
        st.warning("Please upload Master File")
        st.stop()

    try:

        # =====================
        # Load Master File
        # =====================
        master_df = pd.read_excel(master_file)

        if "Final Category" not in master_df.columns:
            st.error("Final Category column not found in Master File")
            st.stop()

        master_df["Final Category"] = (
            master_df["Final Category"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # =====================
        # Load Instruction Sheet
        # =====================
        sheet_name = f"Mapping-{selected_category}"

        instruction_df = pd.read_excel(
            instruction_file,
            sheet_name=sheet_name
        )

        # =====================
        # Load Dropdown File
        # =====================
        dropdown_excel = pd.ExcelFile(dropdown_file)

        dropdown_frames = []

        for sheet in dropdown_excel.sheet_names:
            temp = pd.read_excel(
                dropdown_file,
                sheet_name=sheet
            )
            dropdown_frames.append(temp)

        dropdown_data = pd.concat(
            dropdown_frames,
            ignore_index=True
        )

        dropdown_data.columns = (
            dropdown_data.columns.str.strip()
        )

        # =====================
        # Build Mapping Dictionary
        # =====================
        dropdown_dict = {}

        for _, row in dropdown_data.iterrows():

            try:

                category = str(
                    row["Final Category"]
                ).strip().upper()

                attribute = str(
                    row["Attribute"]
                ).strip()

                base_value = str(
                    row["Base Value"]
                ).strip().lower()

                mapped_value = row["Mapped Value"]

                dropdown_dict[
                    (category, attribute, base_value)
                ] = mapped_value

            except:
                pass

        # =====================
        # Filter Category
        # =====================
        category_df = master_df[
            master_df["Final Category"]
            == selected_category.upper()
        ].copy()

        # =====================
        # Generate Output
        # =====================
        output_df = pd.DataFrame()

        for _, row in instruction_df.iterrows():

            base_col = str(
                row["Base file Column"]
            ).strip()

            output_col = str(
                row["Myntra Template Column"]
            ).strip()

            if base_col in category_df.columns:

                data = category_df[base_col].apply(
                    lambda x: ""
                    if pd.isna(x)
                    else str(x).strip()
                )

                mapped_data = []

                for val in data:

                    key = (
                        selected_category.upper(),
                        output_col,
                        str(val).lower()
                    )

                    mapped_data.append(
                        dropdown_dict.get(key, val)
                    )

                output_df[output_col] = mapped_data

            else:
                output_df[output_col] = ""

        # =====================
        # Create Excel Download
        # =====================
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            output_df.to_excel(
                writer,
                index=False
            )

        output.seek(0)

        st.success("Template Generated Successfully")

        st.download_button(
            label="⬇ Download Template",
            data=output,
            file_name=f"{selected_category}_Template_Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
