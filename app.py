import streamlit as st
from docx import Document
from io import BytesIO

def detect_superscripts_and_find_exact_references(docx_bytes):
    document = Document(BytesIO(docx_bytes))
    superscripts = []

    # Find all superscript numbers in document runs
    for para in document.paragraphs:
        for run in para.runs:
            if run.font.superscript and run.text.strip().isdigit():
                superscripts.append((run.text.strip(), para.text.strip()))

    # For each superscript number, find paragraph(s) that start exactly with "number."
    superscript_with_exact_refs = []
    for num, sup_para_text in superscripts:
        # Find paragraphs starting exactly with "num."
        exact_refs = [para.text.strip() for para in document.paragraphs if para.text.strip().startswith(f"{num}.")]
        superscript_with_exact_refs.append((num, sup_para_text, exact_refs))

    return superscript_with_exact_refs

st.title("Superscript and Exact Reference Finder in DOCX")

uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    results = detect_superscripts_and_find_exact_references(uploaded_file.read())
    
    if results:
        for num, sup_text, exact_refs in results:
            st.markdown(f"### Superscript {num} found in paragraph:")
            st.write(f"> {sup_text}")
            if exact_refs:
                st.markdown(f"**Exact reference paragraph(s) starting with `{num}.`:**")
                for ref in exact_refs:
                    st.write(f"- {ref}")
            else:
                st.write(f"No exact reference starting with '{num}.' found.")
            st.markdown("---")
    else:
        st.write("No superscript numbers found.")
