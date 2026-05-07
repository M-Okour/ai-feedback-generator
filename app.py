import streamlit as st
import pandas as pd
from docx import Document
import zipfile
import io

st.title("AI Student Exam Feedback Generator")

classlist = st.file_uploader("Upload classlist Excel", type=["xlsx"])
template = st.file_uploader("Upload feedback template", type=["docx"])
rubric = st.file_uploader("Upload rubric / feedback bank", type=["xlsx", "docx", "pdf"])
pdfs = st.file_uploader(
    "Upload scanned student exam PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Generate Feedback Files"):
    if not classlist or not template or not pdfs:
        st.error("Please upload the classlist, template, and scanned PDFs.")
    else:
        df = pd.read_excel(classlist)
        st.write("Classlist loaded:")
        st.dataframe(df.head())

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for pdf in pdfs:
                student_id = pdf.name.replace(".pdf", "")

                # Temporary placeholder
                student_name = "Student Name Placeholder"

                doc = Document(template)

                for p in doc.paragraphs:
                    p.text = p.text.replace("Student Name", student_name)
                    p.text = p.text.replace("ID No.", student_id)

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)

                zip_file.writestr(f"{student_id}.docx", doc_buffer.getvalue())

        st.download_button(
            label="Download Feedback ZIP",
            data=zip_buffer.getvalue(),
            file_name="feedback_files.zip",
            mime="application/zip"
        )
