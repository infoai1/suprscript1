import streamlit as st
from docx import Document
from io import BytesIO

def detect_superscripts_and_references(docx_bytes):
    document = Document(BytesIO(docx_bytes))
    superscripts = []

    # Find superscript numbers in runs
    for para in document.paragraphs:
        for run in para.runs:
            if run.font.superscript and run.text.strip().isdigit():
                superscripts.append((run.text.strip(), para.text.strip()))

    # Find reference paragraphs starting with the number + dot (e.g. "28.")
    def find_reference(number):
        for para in document.paragraphs:
            if para.text.strip().startswith(number + '. '):
                return para.text.strip()
        return None

    references = []
    for num, sup_text in superscripts:
        ref_text = find_reference(num)
        if ref_text:
            references.append((num, sup_text, ref_text))

    return superscripts, references

st.title("Superscript and Reference Detector in DOCX")

uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    superscripts, references = detect_superscripts_and_references(uploaded_file.read())
    if superscripts:
        st.subheader("Superscripts Found:")
        for num, para_text in superscripts:
            st.write(f"Superscript {num} found in: {para_text}")
    else:
        st.write("No superscript numbers found.")
    
    if references:
        st.subheader("Matching References:")
        for num, sup_text, ref_text in references:
            st.write(f"Reference {num}: {ref_text}")
    else:
        st.write("No matching references found.")
