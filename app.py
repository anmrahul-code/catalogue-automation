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
        master_df = pd.read_excel(master_file)

        if "Final Category" not in master_df.columns:
            st.error("Final Category column not found in Master File")
            st.stop()

        categories = sorted(
            master_df["Final Category"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.upper()
            .unique()
        )

        selected_category = st.selectbox(
            "Select Category",
            categories
        )

        if st.button("Generate Template"):

            master_df["Final Category"] = (
                master_df["Final Category"]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            instruction_excel = pd.ExcelFile(instruction_file)
            dropdown_excel = pd.ExcelFile(dropdown_file)

            instruction_data = {
                sheet: pd.read_excel(
                    instruction_file,
                    sheet_name=sheet
                )
                for sheet in instruction_excel.sheet_names
            }

            def clean_dropdown_sheet(df):
                df.columns = df.columns.str.strip()

                rename_map = {
                    "Attribute Name of Out Put File": "Attribute",
                    "Base file value": "Base",
                    "Dropdown Value": "Mapped",
                    "Base Value": "Base",
                    "Mapped Value": "Mapped"
                }

                return df.rename(columns=rename_map)

            dropdown_data = pd.concat(
                [
                    clean_dropdown_sheet(
                        pd.read_excel(
                            dropdown_file,
                            sheet_name=sheet
                        )
                    )
                    for sheet in dropdown_excel.sheet_names
                ],
                ignore_index=True
            )

            dropdown_dict = {}

            for _, row in dropdown_data.iterrows():

                try:
                    key = (
                        str(row["Final Category"]).strip().upper(),
                        str(row["Attribute"]).strip(),
                        str(row["Base"]).strip().lower()
                    )

                    dropdown_dict[key] = row["Mapped"]

                except:
                    pass

            sheet_name = f"Mapping-{selected_category}

sheet_name = None

for s in instruction_data.keys():
    if s.replace("Mapping-", "").strip().upper() == selected_category.strip().upper():
        sheet_name = s
        break

if sheet_name is None:
    st.error(
        f"Mapping sheet not found for category: {selected_category}"
    )
    st.stop()
            if sheet_name not in instruction_data:
                st.error(
                    f"{sheet_name} sheet not found in Instruction File"
                )
                st.stop()

            instruction_df = instruction_data[sheet_name]

            cat_df = master_df[
                master_df["Final Category"]
                == selected_category
            ].copy()

            output_df = pd.DataFrame()

            for _, row in instruction_df.iterrows():

                base_col = row["Base file Column"]
                out_col = row["Myntra Template Column"]

                if base_col in cat_df.columns:

                    col_data = cat_df[base_col].apply(
                        lambda x: ""
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

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name=f"{selected_category}_Template_Filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(str(e))
