import streamlit as st

st.set_page_config(page_title="Catalogue Automation")

st.title("Catalogue Automation Portal")

master = st.file_uploader("Upload Master File")

instruction = st.file_uploader("Upload Instruction File")

dropdown = st.file_uploader("Upload Dropdown File")

marketplace = st.selectbox(
    "Marketplace",
    ["Myntra", "Ajio", "Amazon", "Flipkart"]
)

if st.button("Generate"):
    st.success("Files Uploaded Successfully")
