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
    ["Myntra", "Ajio", "Flipkart"]
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
# Generate Template
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
        # Load Mapping Sheet
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
        # Create Dropdown Dictionary
        # =====================
        dropdown_dict = {}

        required_columns = [
            "Attribute",
            "Base Value",
            "Mapped Value",
            "Final Category"
        ]

        if all(col in dropdown_data.columns for col in required_columns):

            for _, row in dropdown_data.iterrows():

                try:

                    key = (
                        str(row["Final Category"]).strip().upper(),
                        str(row["Attribute"]).strip(),
                        str(row["Base Value"]).strip().lower()
                    )

                    dropdown_dict[key] = row["Mapped Value"]

                except:
                    pass

              # =====================
        # Filter Category
        # =====================

        # =====================
# Filter Category
# =====================

if marketplace == "Flipkart" and "Gender" in master_df.columns:gender_series = (
        master_df["Gender"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Kids Tshirts = Boys + Girls
    if selected_category == "Kids Tshirts":

        category_df = master_df[
            (master_df["Final Category"] == "TSHIRTS") &
            (
                gender_series.isin(
                    ["BOYS", "GIRLS"]
                )
            )
        ].copy()

    # Adult Tshirts = Men + Women
    elif selected_category == "Tshirts":

        category_df = master_df[
            (master_df["Final Category"] == "TSHIRTS") &
            (
                gender_series.isin(
                    ["MEN", "WOMEN"]
                )
            )
        ].copy()

    else:

        category_df = master_df[
            master_df["Final Category"]
            == selected_category.upper()
        ].copy()

else:

    category_df = master_df[
        master_df["Final Category"]
        == selected_category.upper()
    ].copy()

        else:

            category_df = master_df[
                master_df["Final Category"]
                == selected_category.upper()
            ].copy()

        # =====================
        # Detect Template Column
        # =====================
        template_column = None

        possible_columns = [
            f"{marketplace} Template Column",
            "Myntra Template Column"
        ]

        for col in possible_columns:
            if col in instruction_df.columns:
                template_column = col
                break

        if template_column is None:
            st.error("Template column not found in instruction file")
            st.stop()

        # =====================
        # Generate Output
        # =====================
        output_df = pd.DataFrame()

        for _, row in instruction_df.iterrows():

            base_col = str(
                row["Base file Column"]
            ).strip()

            output_col = str(
                row[template_column]
            ).strip()

            if output_col == "" or output_col.lower() == "nan":
                continue

            if base_col in category_df.columns:

                values = category_df[base_col].fillna("").astype(str)

                mapped_values = []

                for val in values:

                    key = (
                        selected_category.upper(),
                        output_col,
                        str(val).strip().lower()
                    )

                    mapped_values.append(
                        dropdown_dict.get(key, val)
                    )

                output_df[output_col] = mapped_values

            else:
                output_df[output_col] = ""

        # =====================
        # Create Excel File
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
            file_name=f"{marketplace}_{selected_category}_Template_Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
