import streamlit as st
import pandas as pd
import io

# ==================================
# Page Setup
# ==================================

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

# ==================================
# Load Categories
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

    st.error(
        f"Instruction File Error: {e}"
    )
    st.stop()

selected_category = st.selectbox(
    "Select Category",
    categories
)

# ==================================
# Generate Template
# ==================================

if st.button("Generate Template"):

    if master_file is None:

        st.warning(
            "Please upload Master File"
        )
        st.stop()

    try:

        # ==================================
        # Load Files
        # ==================================

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

        # ==================================
        # Category Filter
        # ==================================

        category_column = (
            f"{marketplace} Category"
        )

        if category_column not in master_df.columns:

            st.error(
                f"{category_column} not found in Master File"
            )
            st.stop()

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

        # ==================================
        # Template Column
        # ==================================

        template_column = (
            f"{marketplace} Template Column"
        )

        if template_column not in instruction_df.columns:

            st.error(
                f"{template_column} not found"
            )
            st.stop()

        # ==================================
        # Dropdown Dictionary
        # ==================================

        dropdown_dict = {}

        required_cols = [
            "Attribute",
            "Base Value",
            "Mapped Value"
        ]

        if all(
            col in dropdown_df.columns
            for col in required_cols
        ):

            for _, row in dropdown_df.iterrows():

                try:

                    key = (
                        str(
                            row["Attribute"]
                        ).strip().upper(),

                        str(
                            row["Base Value"]
                        ).strip().upper()
                    )

                    dropdown_dict[key] = (
                        str(
                            row["Mapped Value"]
                        ).strip()
                    )

                except:
                    pass

        # ==================================
        # Generate Output
        # ==================================

        output_df = pd.DataFrame()

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

            # ==================================
            # Dropdown Mapping
            # ==================================

            if "dropdown" in remarks:

                lookup_value = ""

                if (
                    base_col in category_df.columns
                    and len(category_df) > 0
                ):

                    lookup_value = str(
                        category_df.iloc[0][base_col]
                    ).strip()

                key = (
                    output_col.upper(),
                    lookup_value.upper()
                )

                mapped_value = dropdown_dict.get(
                    key,
                    ""
                )

                output_df[output_col] = (
                    [mapped_value]
                    * len(category_df)
                )

            # ==================================
            # Blank Column
            # ==================================

            elif (
                base_col == ""
                or base_col.lower() == "none"
                or base_col == "nan"
            ):

                output_df[output_col] = (
                    [""] * len(category_df)
                )

            # ==================================
            # Direct Mapping
            # ==================================

            elif base_col in category_df.columns:

                output_df[output_col] = (
                    category_df[base_col]
                    .fillna("")
                    .astype(str)
                    .values
                )

            # ==================================
            # Default Blank
            # ==================================

            else:

                output_df[output_col] = (
                    [""] * len(category_df)
                )

        # ==================================
        # Debug
        # ==================================

        st.success(
            "Template Generated Successfully"
        )

        st.write(
            "Output Rows",
            len(output_df)
        )

        st.write(
            "Output Columns",
            len(output_df.columns)
        )

        st.dataframe(
            output_df.head()
        )

        # ==================================
        # Excel Download
        # ==================================

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

        st.download_button(
            label="⬇ Download Template",
            data=output,
            file_name=f"{marketplace}_{selected_category}_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error: {str(e)}"
        )
