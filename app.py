import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Catalogue Automation Portal",
    layout="wide"
)

st.title("📦 Catalogue Automation Portal")

# ==================================
# Upload Master File
# ==================================

master_file = st.file_uploader(
    "Upload Master File",
    type=["xlsx"]
)

# ==================================
# Marketplace
# ==================================

marketplace = st.selectbox(
    "Select Marketplace",
    ["Flipkart", "Myntra", "Ajio"]
)

instruction_file = f"config/{marketplace}_instruction.xlsx"
dropdown_file = f"config/{marketplace}_dropdown.xlsx"

# ==================================
# Categories
# ==================================

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

# ==================================
# Generate
# ==================================

if st.button("Generate Template"):

    if master_file is None:

        st.warning(
            "Please upload Master File"
        )
        st.stop()

    try:

        # ----------------------------
        # Load Files
        # ----------------------------

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

        # ----------------------------
        # Filter Category
        # ----------------------------

        category_column = (
            f"{marketplace} Category"
        )

        category_df = master_df[
            master_df[category_column]
            .astype(str)
            .str.strip()
            ==
            selected_category
        ].copy()

        if len(category_df) == 0:

            st.error(
                "No rows found for selected category"
            )
            st.stop()

        # ----------------------------
        # Template Column
        # ----------------------------

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

            st.error(
                "Template column not found"
            )
            st.stop()

        # ----------------------------
        # Dropdown Dictionary
        # ----------------------------

        dropdown_dict = {}

        for _, row in dropdown_df.iterrows():

            try:

                key = (
                    str(row["Attribute"])
                    .strip()
                    .upper(),

                    str(row["Base Value"])
                    .strip()
                    .upper()
                )

                dropdown_dict[key] = (
                    str(row["Mapped Value"])
                    .strip()
                )

            except:
                pass

        # ----------------------------
        # Output
        # ----------------------------

        output_df = pd.DataFrame()

        category_based_fields = [
            "Listing Status",
            "Fullfilment by",
            "Procurement type",
            "Procurement SLA (DAY)",
            "Shipping provider",
            "Local handling fee (INR)",
            "Zonal handling fee (INR)",
            "National handling fee (INR)",
            "Country Of Origin",
            "Tax Code"
        ]

        for _, row in instruction_df.iterrows():

            base_col = str(
                row.get(
                    "Base file Column",
                    ""
                )
            ).strip()

            output_col = str(
                row.get(
                    template_column,
                    ""
                )
            ).strip()

            remarks = str(
                row.get(
                    "Remarks",
                    ""
                )
            ).strip().lower()

            if (
                output_col == ""
                or output_col.lower() == "nan"
            ):
                continue

            # ----------------------------
            # Dropdown
            # ----------------------------

            if "dropdown" in remarks:

                if output_col in category_based_fields:

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

                else:

                    mapped_values = []

                    if base_col in category_df.columns:

                        values = (
                            category_df[base_col]
                            .fillna("")
                            .astype(str)
                        )

                        for val in values:

                            key = (
                                output_col.upper(),
                                val.strip().upper()
                            )

                            mapped_values.append(
                                dropdown_dict.get(
                                    key,
                                    val
                                )
                            )

                        output_df[output_col] = (
                            mapped_values
                        )

                    else:

                        output_df[output_col] = (
                            [""] * len(category_df)
                        )

            # ----------------------------
            # Blank Column
            # ----------------------------

            elif (
                base_col == ""
                or base_col.lower() == "none"
                or base_col.lower() == "nan"
            ):

                output_df[output_col] = (
                    [""] * len(category_df)
                )

            # ----------------------------
            # Direct Mapping
            # ----------------------------

            elif base_col in category_df.columns:

                output_df[output_col] = (
                    category_df[base_col]
                    .fillna("")
                    .astype(str)
                    .values
                )

            # ----------------------------
            # Default Blank
            # ----------------------------

            else:

                output_df[output_col] = (
                    [""] * len(category_df)
                )

        # ----------------------------
        # Download
        # ----------------------------

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

        st.dataframe(
            output_df.head()
        )

        st.download_button(
            label="⬇ Download Template",
            data=output,
            file_name=f"{marketplace}_{selected_category}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error: {str(e)}"
        )
