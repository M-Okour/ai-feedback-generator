import io
import re
import json
import time
import zipfile

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from openai import OpenAI
from openai import RateLimitError, APIError, APITimeoutError


# =========================================================
# Streamlit setup
# =========================================================

st.set_page_config(page_title="AI Student Exam Feedback Generator", layout="wide")
st.title("AI Student Exam Feedback Generator")


# =========================================================
# OpenAI client
# =========================================================

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# =========================================================
# PC mapping
# =========================================================

pc_map = {
    "E1:PC1": "PC1.1",
    "E1:PC2": "PC1.2",
    "E1:PC3": "PC1.3",
    "E2:PC1": "PC2.1",
    "E2:PC2": "PC2.2",
}


# =========================================================
# General helpers
# =========================================================

def normalize_pc_code(text):
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


def find_student_by_id(df, student_id):
    student_id = str(student_id).strip()

    for _, row in df.iterrows():
        values = [str(v).strip() for v in row.values]
        if student_id in values:
            return row

    return None


def guess_student_name(student_row):
    possible_columns = [
        "Student Name",
        "Name",
        "Full Name",
        "Student",
        "Learner Name",
        "Student Full Name"
    ]

    for col in possible_columns:
        if col in student_row.index:
            value = str(student_row[col]).strip()
            if value and value.lower() != "nan":
                return value

    for value in student_row.values:
        value = str(value).strip()
        if value and value.lower() != "nan" and not value.isdigit():
            if not value.upper().startswith("H00"):
                return value

    return "Unknown Student"


# =========================================================
# File reading helpers
# =========================================================

def read_docx_as_text(uploaded_file):
    uploaded_file.seek(0)
    doc = Document(uploaded_file)

    text_parts = []

    for p in doc.paragraphs:
        if p.text.strip():
            text_parts.append(p.text.strip())

    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip():
                text_parts.append(row_text)

    return "\n".join(text_parts)


def read_pdf_as_text(uploaded_file):
    uploaded_file.seek(0)
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


def read_excel_as_text(uploaded_file):
    uploaded_file.seek(0)
    sheets = pd.read_excel(uploaded_file, sheet_name=None)

    text_parts = []

    for sheet_name, df in sheets.items():
        text_parts.append(f"\n--- Sheet: {sheet_name} ---")
        text_parts.append(df.to_string(index=False))

    return "\n".join(text_parts)


def read_uploaded_file_as_text(uploaded_file):
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".docx"):
        return read_docx_as_text(uploaded_file)

    if filename.endswith(".pdf"):
        return read_pdf_as_text(uploaded_file)

    if filename.endswith(".xlsx"):
        return read_excel_as_text(uploaded_file)

    if filename.endswith(".txt") or filename.endswith(".csv"):
        uploaded_file.seek(0)
        return uploaded_file.read().decode("utf-8", errors="ignore")

    return ""


# =========================================================
# Rubric extraction
# =========================================================

def extract_rubric_sections(rubric_text):
    """
    Extracts PC-specific rubric blocks from the uploaded rubric.

    Returns:
    {
        "E1:PC1": "...rubric text...",
        "E1:PC2": "...rubric text...",
        ...
    }
    """

    pc_patterns = {
        "E1:PC1": r"(PC\s*1\.1.*?)(?=PC\s*1\.2|PC1\.2|$)",
        "E1:PC2": r"(PC\s*1\.2.*?)(?=PC\s*1\.3|PC1\.3|$)",
        "E1:PC3": r"(PC\s*1\.3.*?)(?=Element\s*2|PC\s*2\.1|PC2\.1|$)",
        "E2:PC1": r"(PC\s*2\.1.*?)(?=PC\s*2\.2|PC2\.2|$)",
        "E2:PC2": r"(PC\s*2\.2.*?)(?=Document\s*Title|$)",
    }

    sections = {}

    for pc_code, pattern in pc_patterns.items():
        match = re.search(pattern, rubric_text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            sections[pc_code] = match.group(1).strip()
        else:
            sections[pc_code] = rubric_text

    return sections


# =========================================================
# OCR scanned student exam
# =========================================================

def extract_text_from_scanned_pdf(uploaded_pdf):
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
    Detect sections such as:
    E1:PC1
    E1 PC1
    E1: PC1
    E 1 : PC 1
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


# =========================================================
# AI grading
# =========================================================

def safe_json_loads(raw_text):
    raw_text = raw_text.strip()

    try:
        return json.loads(raw_text)
    except Exception:
        pass

    match = re.search(r"\{.*\}", raw_text, flags=re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            pass

    return None


def grade_pc_with_ai(pc_code, answer_text, pc_rubric_text, max_retries=5):
    answer_text = answer_text[:4500]
    pc_rubric_text = pc_rubric_text[:4500]

    prompt = f"""
You are an assessor for MCT 122 Analyse Static Loads.

Use ONLY the provided rubric section to assess the student's answer.

Performance Criterion:
{pc_code}

Rubric section:
{pc_rubric_text}

Student OCR answer:
{answer_text}

Task:
1. Decide which rubric level best matches the student's answer.
2. Assign a mark within the correct mark range.
3. Write short formal feedback explaining:
   - what the student did correctly
   - where marks were lost
   - what should be improved

Return ONLY valid JSON:
{{
  "pc": "{pc_code}",
  "mark": 0,
  "level": "Not Yet Competent",
  "feedback": "..."
}}

Mark ranges:
- Not Yet Competent: 0-59
- Competent: 60-69
- Competent with Merit: 70-84
- Competent with Distinction: 85-100

Important rules:
- Do not give Competent with Distinction unless the answer fully satisfies the distinction rubric.
- Do not invent missing calculations, diagrams, units, or steps.
- If OCR is unclear or the answer is missing, assign Not Yet Competent.
- Feedback must be 1 to 2 sentences.
- Feedback must be suitable for the Assessor Feedback section.
"""

    for attempt in range(max_retries):
        try:
            response = client.responses.create(
                model="gpt-4.1-mini",
                input=prompt,
                max_output_tokens=300
            )

            raw = response.output_text.strip()
            result = safe_json_loads(raw)

            if result is None:
                raise ValueError("AI response was not valid JSON.")

            mark = int(result.get("mark", 0))
            mark = max(0, min(100, mark))

            result["pc"] = pc_code
            result["mark"] = mark
            result["level"] = grade_level(mark)

            time.sleep(2)
            return result

        except RateLimitError:
            wait_time = 10 * (attempt + 1)
            st.warning(f"Rate limit reached for {pc_code}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

        except (APIError, APITimeoutError):
            wait_time = 5 * (attempt + 1)
            st.warning(f"Temporary API issue for {pc_code}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

        except Exception:
            return {
                "pc": pc_code,
                "mark": 0,
                "level": "Not Yet Competent",
                "feedback": "The answer could not be reliably graded because the extracted response or OCR text was unclear."
            }

    return {
        "pc": pc_code,
        "mark": 0,
        "level": "Not Yet Competent",
        "feedback": "The answer could not be graded because the API rate limit was reached after several retries."
    }


# =========================================================
# DOCX filling
# =========================================================

def fill_feedback_template(doc, student_name, student_id, pc_marks, feedback_rows):
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

            # Fill PC marks
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


# =========================================================
# Upload interface
# =========================================================

classlist = st.file_uploader("Upload classlist Excel", type=["xlsx"])
template = st.file_uploader("Upload feedback template DOCX", type=["docx"])
rubric = st.file_uploader("Upload grading rubric", type=["docx", "pdf", "xlsx", "txt", "csv"])

pdfs = st.file_uploader(
    "Upload scanned student exam PDFs",
    type=["pdf"],
    accept_multiple_files=True
)


# =========================================================
# Main processing
# =========================================================

if st.button("Generate Feedback Files"):
    if not classlist:
        st.error("Please upload the classlist Excel file.")
        st.stop()

    if not template:
        st.error("Please upload the feedback DOCX template.")
        st.stop()

    if not rubric:
        st.error("Please upload the grading rubric.")
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

    rubric_sections = extract_rubric_sections(rubric_text)

    with st.expander("Preview extracted rubric sections"):
        for pc_code, section in rubric_sections.items():
            st.markdown(f"### {pc_code}")
            st.text(section[:1500])

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

            with st.expander(f"OCR and detected PC sections for {student_id}"):
                st.markdown("### Detected PC sections")
                st.write(list(pc_sections.keys()))
                st.markdown("### OCR preview")
                st.text(exam_text[:3000])

            pc_marks = {}
            feedback_rows = {}

            for exam_pc, template_pc in pc_map.items():
                answer_text = pc_sections.get(exam_pc, "")
                pc_rubric_text = rubric_sections.get(exam_pc, rubric_text)

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
                        pc_rubric_text=pc_rubric_text
                    )

                pc_marks[exam_pc] = result["mark"]
                feedback_rows[template_pc] = (
                    result["level"],
                    result["feedback"]
                )

                report_rows.append({
                    "Student ID": student_id,
                    "Student Name": student_name,
                    "Exam PC": exam_pc,
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

            zip_file.writestr(f"{student_id}.docx", doc_buffer.getvalue())

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
