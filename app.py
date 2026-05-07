import io
import re
import json
import zipfile

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from openai import OpenAI


# -----------------------------
# Streamlit setup
# -----------------------------

st.set_page_config(page_title="AI Student Exam Feedback Generator", layout="wide")
st.title("AI Student Exam Feedback Generator")


# -----------------------------
# OpenAI client
# -----------------------------

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# -----------------------------
# Helper functions
# -----------------------------

def normalize_pc_code(text):
    """
    Converts variations like:
    E1 PC1, E1: PC1, E 1 : PC 1
    into:
    E1:PC1
    """
    text = text.upper()
    text = re.sub(r"\s+", "", text)
    match = re.search(r"E(\d+):?PC(\d+)", text)
    if match:
        return f"E{match.group(1)}:PC{match.group(2)}"
    return text


def grade_level(mark):
    mark = int(mark)

    if mark < 60:
        return "Not Yet Competent"
    elif mark < 70:
        return "Competent"
    elif mark < 85:
        return "Competent with Merit"
    else:
        return "Competent with Distinction"


def read_uploaded_file_as_text(uploaded_file):
    """
    Reads uploaded rubric file.
    Supports txt, csv, xlsx, docx, pdf.
    """
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".txt") or filename.endswith(".csv"):
        return uploaded_file.read().decode("utf-8", errors="ignore")

    if filename.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
        return df.to_string(index=False)

    if filename.endswith(".docx"):
        doc = Document(uploaded_file)
        text = []
        for p in doc.paragraphs:
            text.append(p.text)

        for table in doc.tables:
            for row in table.rows:
                text.append(" | ".join(cell.text for cell in row.cells))

        return "\n".join(text)

    if filename.endswith(".pdf"):
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""

        for page in doc:
            direct_text = page.get_text()
            if direct_text.strip():
                text += "\n" + direct_text
            else:
                pix = page.get_pixmap(dpi=220)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text += "\n" + pytesseract.image_to_string(img)

        return text

    return ""


def extract_text_from_scanned_pdf(uploaded_pdf):
    """
    OCR scanned exam PDF.
    """
    uploaded_pdf.seek(0)
    pdf_bytes = uploaded_pdf.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    full_text = ""

    for page_number, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=220)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(img)

        full_text += f"\n\n--- PAGE {page_number} ---\n{text}"

    return full_text


def extract_pc_sections(exam_text):
    """
    Extract answer sections based on E1:PC1, E1:PC2, etc.
    """
    pattern = r"(E\s*\d+\s*:?\s*PC\s*\d+)"
    matches = list(re.finditer(pattern, exam_text, flags=re.IGNORECASE))

    sections = {}

    for i, match in enumerate(matches):
        pc_code = normalize_pc_code(match.group(1))

        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(exam_text)

        section_text = exam_text[start:end].strip()

        if pc_code not in sections:
            sections[pc_code] = section_text
        else:
            sections[pc_code] += "\n" + section_text

    return sections


def grade_pc_with_ai(pc_code, answer_text, rubric_text):
    """
    Uses AI to grade one PC answer.
    """
    prompt = f"""
You are an assessor for MCT 122 Analyse Static Loads.

Grade the student's answer for the following performance criterion.

Performance Criterion:
{pc_code}

Rubric:
{rubric_text}

Student OCR answer:
{answer_text}

Return ONLY valid JSON in this exact format:
{{
  "pc": "{pc_code}",
  "mark": 0,
  "level": "Not Yet Competent",
  "feedback": "Short formal feedback explaining where marks were lost."
}}

Rules:
- mark must be an integer from 0 to 100.
- level must match the mark:
  0-59 = Not Yet Competent
  60-69 = Competent
  70-84 = Competent with Merit
  85-100 = Competent with Distinction
- feedback must be 1 to 2 sentences.
- feedback must be specific to the student's answer.
- Do not invent work that is not visible in the OCR text.
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )

    raw = response.output_text.strip()

    try:
        result = json.loads(raw)
    except Exception:
        result = {
            "pc": pc_code,
            "mark": 0,
            "level": "Not Yet Competent",
            "feedback": "The answer could not be reliably graded because the AI output was not readable."
        }

    mark = int(result.get("mark", 0))
    mark = max(0, min(100, mark))

    result["mark"] = mark
    result["level"] = grade_level(mark)

    return result


def find_student_by_id(df, student_id):
    """
    Finds a student row by searching all Excel cells for the ID.
    """
    student_id = str(student_id).strip()

    for _, row in df.iterrows():
        values = [str(v).strip() for v in row.values]
        if student_id in values:
            return row

    return None


def guess_student_name(student_row):
    """
    Tries to identify student name column automatically.
    """
    possible_columns = [
        "Student Name",
        "Name",
        "Full Name",
        "Student",
        "Learner Name"
    ]

    for col in possible_columns:
        if col in student_row.index:
            return str(student_row[col])

    # fallback: return first non-empty text cell
    for value in student_row.values:
        if isinstance(value, str) and not value.strip().isdigit():
            return value.strip()

    return "Unknown Student"


def clear_cell(cell):
    cell.text = ""


def fill_feedback_template(doc, student_name, student_id, pc_marks, feedback_rows):
    """
    Fills the Word feedback template.

    pc_marks:
    {
        "E1:PC1": 60,
        "E1:PC2": 65,
        ...
    }

    feedback_rows:
    {
        "PC1.1": ("Competent", "Feedback text"),
        ...
    }
    """

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = [cell.text.strip() for cell in cells]

            # Fill Student Name and ID No.
            for i, cell in enumerate(cells):
                text = cell.text.strip()

                if text == "Student Name" and i + 1 < len(cells):
                    cells[i + 1].text = student_name

                if text == "ID No." and i + 1 < len(cells):
                    cells[i + 1].text = student_id

            # Fill PC marks in Assessment Results table
            for pc_code, mark in pc_marks.items():
                if pc_code in row_text:
                    for cell in cells:
                        if cell.text.strip() == "60":
                            cell.text = str(mark)
                            break

            # Fill summative grade
            if "Summative Assessment Grade %:" in row_text:
                marks = list(pc_marks.values())

                if marks:
                    if any(m < 60 for m in marks):
                        summative = min(marks)
                    else:
                        summative = round(sum(marks) / len(marks))

                    for cell in cells:
                        if cell.text.strip() == "60":
                            cell.text = str(summative)
                            break

            # Fill feedback rows
            for pc_short, (level, feedback) in feedback_rows.items():
                if pc_short in row_text:
                    pc_index = row_text.index(pc_short)

                    if pc_index + 2 < len(cells):
                        cells[pc_index + 1].text = level
                        cells[pc_index + 2].text = feedback

    return doc


# -----------------------------
# Upload section
# -----------------------------

classlist = st.file_uploader("Upload classlist Excel", type=["xlsx"])
template = st.file_uploader("Upload feedback template DOCX", type=["docx"])
rubric = st.file_uploader("Upload rubric / feedback bank", type=["xlsx", "docx", "pdf", "txt", "csv"])
pdfs = st.file_uploader(
    "Upload scanned student exam PDFs",
    type=["pdf"],
    accept_multiple_files=True
)


# -----------------------------
# PC mapping
# -----------------------------

pc_map = {
    "E1:PC1": "PC1.1",
    "E1:PC2": "PC1.2",
    "E1:PC3": "PC1.3",
    "E2:PC1": "PC2.1",
    "E2:PC2": "PC2.2",
}


# -----------------------------
# Main process
# -----------------------------

if st.button("Generate Feedback Files"):
    if not classlist:
        st.error("Please upload the classlist Excel file.")
        st.stop()

    if not template:
        st.error("Please upload the feedback DOCX template.")
        st.stop()

    if not rubric:
        st.error("Please upload the rubric / feedback bank.")
        st.stop()

    if not pdfs:
        st.error("Please upload scanned student exam PDFs.")
        st.stop()

    df = pd.read_excel(classlist)

    st.subheader("Classlist Preview")
    st.dataframe(df.head())

    rubric_text = read_uploaded_file_as_text(rubric)

    if not rubric_text.strip():
        st.error("The rubric file could not be read.")
        st.stop()

    zip_buffer = io.BytesIO()
    report_rows = []

    progress = st.progress(0)

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, pdf in enumerate(pdfs):
            student_id = pdf.name.replace(".pdf", "").strip()

            st.write(f"Processing: {student_id}")

            student_row = find_student_by_id(df, student_id)

            if student_row is None:
                st.warning(f"Student ID not found in classlist: {student_id}")
                continue

            student_name = guess_student_name(student_row)

            exam_text = extract_text_from_scanned_pdf(pdf)
            pc_sections = extract_pc_sections(exam_text)

            pc_marks = {}
            feedback_rows = {}

            for exam_pc, template_pc in pc_map.items():
                answer_text = pc_sections.get(exam_pc, "")

                if not answer_text.strip():
                    result = {
                        "pc": exam_pc,
                        "mark": 0,
                        "level": "Not Yet Competent",
                        "feedback": f"No clear answer section was detected for {exam_pc}. The student should ensure the answer is clearly labelled and complete."
                    }
                else:
                    result = grade_pc_with_ai(
                        pc_code=exam_pc,
                        answer_text=answer_text,
                        rubric_text=rubric_text
                    )

                pc_marks[exam_pc] = result["mark"]
                feedback_rows[template_pc] = (
                    result["level"],
                    result["feedback"]
                )

                report_rows.append({
                    "Student ID": student_id,
                    "Student Name": student_name,
                    "PC": exam_pc,
                    "Template PC": template_pc,
                    "Mark": result["mark"],
                    "Level": result["level"],
                    "Feedback": result["feedback"]
                })

            template.seek(0)
            doc = Document(template)

            doc = fill_feedback_template(
                doc=doc,
                student_name=student_name,
                student_id=student_id,
                pc_marks=pc_marks,
                feedback_rows=feedback_rows
            )

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            output_name = f"{student_id}.docx"
            zip_file.writestr(output_name, doc_buffer.getvalue())

            progress.progress((index + 1) / len(pdfs))

    zip_buffer.seek(0)

    st.success("Feedback files generated successfully.")

    st.download_button(
        label="Download Feedback DOCX ZIP",
        data=zip_buffer.getvalue(),
        file_name="feedback_files.zip",
        mime="application/zip"
    )

    if report_rows:
        st.subheader("Generated Marks and Feedback Summary")
        report_df = pd.DataFrame(report_rows)
        st.dataframe(report_df)

        csv_buffer = io.StringIO()
        report_df.to_csv(csv_buffer, index=False)

        st.download_button(
            label="Download Feedback Summary CSV",
            data=csv_buffer.getvalue(),
            file_name="feedback_summary.csv",
            mime="text/csv"
        )
