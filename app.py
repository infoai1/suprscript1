import streamlit as st
from docx import Document
from io import BytesIO

def detect_superscripts_and_find_references(docx_bytes):
    document = Document(BytesIO(docx_bytes))
    superscripts = []

    # Find all superscript numbers in runs
    for para in document.paragraphs:
        for run in para.runs:
            if run.font.superscript and run.text.strip().isdigit():
                superscripts.append((run.text.strip(), para.text.strip()))

    # For each superscript, find all paragraphs containing that number anywhere in the text
    superscript_with_matches = []
    for num, sup_para_text in superscripts:
        matches = [para.text.strip() for para in document.paragraphs if num in para.text and para.text.strip() != sup_para_text]
        superscript_with_matches.append((num, sup_para_text, matches))

    return superscript_with_matches

st.title("Superscript Detection and Reference Finder in DOCX")

uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    results = detect_superscripts_and_find_references(uploaded_file.read())
    
    if results:
        for num, sup_text, matches in results:
            st.markdown(f"### Superscript {num} found in:")
            st.write(f"> {sup_text}")
            if matches:
                st.markdown(f"**Other places containing `{num}`:**")
                for match in matches:
                    st.write(f"- {match}")
            else:
                st.write("No other references found.")
            st.markdown("---")
    else:
        st.write("No superscript numbers found in the document.")
