import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Catalogue Automation Portal",
    layout="wide"
)

st.title("📦 Catalogue Automation Portal")

master_file = st.file_uploader(
    "Upload Master File",
    type=["xlsx"]
)

instruction_file = st.file_uploader(
    "Upload Instruction File",
    type=["xlsx"]
)

dropdown_file = st.file_uploader(
    "Upload Dropdown File",
    type=["xlsx"]
)

if master_file and instruction_file and dropdown_file:

    try:

        # ==========================
        # LOAD MASTER FILE
        # ==========================
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

        categories = sorted(
            master_df["Final Category"]
            .dropna()
            .unique()
        )

        selected_category = st.selectbox(
            "Select Category",
            categories
        )

        if st.button("Generate Template"):

            # ==========================
            # LOAD INSTRUCTION FILE
            # ==========================
            instruction_excel = pd.ExcelFile(instruction_file)

            instruction_data = {
                sheet: pd.read_excel(
                    instruction_file,
                    sheet_name=sheet
                )
                for sheet in instruction_excel.sheet_names
            }

            # ==========================
            # FIND CORRECT MAPPING SHEET
            # ==========================
            sheet_name = None

            for sheet in instruction_data.keys():

                sheet_category = (
                    sheet.replace("Mapping-", "")
                    .strip()
                    .upper()
                )

                if sheet_category == selected_category:
                    sheet_name = sheet
                    break

            if sheet_name is None:

                st.error(
                    f"No Mapping Sheet Found For {selected_category}"
                )

                st.write("Available Mapping Sheets:")

                st.write(
                    list(instruction_data.keys())
                )

                st.stop()

            instruction_df = instruction_data[sheet_name]

            # ==========================
            # LOAD DROPDOWN FILE
            # ==========================
            dropdown_excel = pd.ExcelFile(dropdown_file)

            def clean_dropdown_sheet(df):

                df.columns = (
                    df.columns
                    .astype(str)
                    .str.strip()
                )

                rename_map = {
                    "Attribute Name of Out Put File": "Attribute",
                    "Base file value": "Base",
                    "Dropdown Value": "Mapped",
                    "Base Value": "Base",
                    "Mapped Value": "Mapped"
                }

                return df.rename(columns=rename_map)

            dropdown_frames = []

            for sheet in dropdown_excel.sheet_names:

                temp_df = pd.read_excel(
                    dropdown_file,
                    sheet_name=sheet
                )

                temp_df = clean_dropdown_sheet(temp_df)

                dropdown_frames.append(temp_df)

            dropdown_data = pd.concat(
                dropdown_frames,
                ignore_index=True
            )

            # ==========================
            # BUILD DROPDOWN DICTIONARY
            # ==========================
            dropdown_dict = {}

            for _, row in dropdown_data.iterrows():

                try:

                    final_category = str(
                        row["Final Category"]
                    ).strip().upper()

                    attribute = str(
                        row["Attribute"]
                    ).strip()

                    base_value = str(
                        row["Base"]
                    ).strip().lower()

                    mapped_value = row["Mapped"]

                    dropdown_dict[
                        (
                            final_category,
                            attribute,
                            base_value
                        )
                    ] = mapped_value

                except:
                    pass

            # ==========================
            # FILTER CATEGORY
            # ==========================
            cat_df = master_df[
                master_df["Final Category"]
                == selected_category
            ].copy()

            output_df = pd.DataFrame()

            # ==========================
            # GENERATE OUTPUT
            # ==========================
            for _, row in instruction_df.iterrows():

                try:

                    base_col = str(
                        row["Base file Column"]
                    ).strip()

                    out_col = str(
                        row["Myntra Template Column"]
                    ).strip()

                    if (
                        base_col
                        and base_col in cat_df.columns
                    ):

                        col_data = cat_df[
                            base_col
                        ].apply(
                            lambda x:
                            ""
                            if pd.isna(x)
                            else str(x).strip()
                        )

                        col_data = col_data.apply(
                            lambda val:
                            dropdown_dict.get(
                                (
                                    selected_category,
                                    out_col,
                                    val.lower()
                                ),
                                val
                            )
                        )

                        output_df[out_col] = col_data

                    else:

                        output_df[out_col] = ""

                except:
                    pass

            # ==========================
            # CREATE EXCEL FILE
            # ==========================
            output = BytesIO()

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
                f"{selected_category} Template Generated Successfully"
            )

            st.write(
                f"Rows Generated: {len(output_df)}"
            )

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name=f"{selected_category}_Template_Filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error("Error Found")

        st.code(str(e))
