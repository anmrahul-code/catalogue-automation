import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Catalogue Automation Portal",
    layout="wide"
)

st.title("📦 Catalogue Automation Portal")

# Upload Master
master_file = st.file_uploader(
    "Upload Master File",
    type=["xlsx"]
)

# Marketplace
marketplace = st.selectbox(
    "Select Marketplace",
    ["Flipkart", "Myntra", "Ajio"]
)

instruction_file = f"config/{marketplace}_instruction.xlsx"
dropdown_file = f"config/{marketplace}_dropdown.xlsx"

# Load Categories
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

# Generate
if st.button("Generate Template"):

    try:

        # =====================
        # Load Files
        # =====================

        master_df = pd.read_excel(
            master_file
        )

        instruction_df = pd.read_excel(
            instruction_file,
            sheet_name=f"Mapping-{selected_category}"
        )

        dropdown_df = pd.read_excel(
            dropdown_file,
            sheet_name=f"Dropdown-{selected_category}"
        )

        # =====================
# Filter Category
# =====================

if "Final Category" not in master_df.columns:

    st.error(
        "Final Category column not found in Master File"
    )
    st.stop()

master_df["Final Category"] = (
    master_df["Final Category"]
    .astype(str)
    .str.strip()
)

category_df = master_df[
    master_df["Final Category"]
    ==
    selected_category
].copy()

# DEBUG

st.write(
    "Selected Category",
    selected_category
)

st.write(
    "Master Categories",
    master_df["Final Category"]
    .dropna()
    .unique()
)

st.write(
    "Category Rows Found",
    len(category_df)
)

if len(category_df) == 0:

    st.error(
        "No rows found for selected category"
    )
    st.stop()
        # =====================
        # Dropdown Dictionary
        # =====================

        dropdown_dict = {}

        for _, row in dropdown_df.iterrows():

            try:

                key = (
                    str(row["Attribute"]).strip().upper(),
                    str(row["Base Value"]).strip().upper()
                )

                dropdown_dict[key] = (
                    str(row["Mapped Value"]).strip()
                )

            except:
                pass

        # =====================
        # Output
        # =====================

        output_df = pd.DataFrame()

        template_column = (
            f"{marketplace} Template Column"
        )

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

            if (
                output_col == ""
                or output_col.lower() == "nan"
            ):
                continue

            # Dropdown Value

            if "dropdown" in remarks:

                key = (
                    output_col.upper(),
                    selected_category.upper()
                )

                mapped_value = (
                    dropdown_dict.get(
                        key,
                        ""
                    )
                )

                output_df[output_col] = (
                    [mapped_value]
                    * len(category_df)
                )

            # Master Value

            elif (
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

        # =====================
        # Download
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

        st.success(
            "Template Generated Successfully"
        )

        st.download_button(
            "Download Template",
            data=output,
            file_name=f"{marketplace}_{selected_category}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(str(e))
