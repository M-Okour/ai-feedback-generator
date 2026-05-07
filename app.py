import io
import re
import json
import time
import zipfile

import streamlit as st
import pandas as pd
import fitz
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
# General helpers
# =========================================================

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


def normalize_any_pc(text):
    """
    Converts:
    PC3.1      -> E3:PC3.1
    3.1        -> E3:PC3.1
    E3:PC3.1   -> E3:PC3.1
    E 3 : PC 3.1 -> E3:PC3.1
    """

    text = str(text).upper().replace(" ", "")

    match = re.search(r"E(\d+):?PC(\d+\.\d+)", text)
    if match:
        return f"E{match.group(1)}:PC{match.group(2)}"

    match = re.search(r"PC(\d+\.\d+)", text)
    if match:
        pc_number = match.group(1)
        element = pc_number.split(".")[0]
        return f"E{element}:PC{pc_number}"

    match = re.search(r"\b(\d+\.\d+)\b", text)
    if match:
        pc_number = match.group(1)
        element = pc_number.split(".")[0]
        return f"E{element}:PC{pc_number}"

    return text


def template_pc_label(pc_code):
    """
    E3:PC3.1 -> PC3.1
    """
    return pc_code.split(":")[-1]


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
        "Student Full Name",
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
# File reading
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
# Dynamic PC extraction from rubric
# =========================================================

def extract_pc_list_from_rubric(rubric_text):
    """
    Extracts PCs from the rubric dynamically.

    Supports:
    3.1
    PC3.1
    PC 3.1
    E3:PC3.1
    E4:PC4.1
    """

    pcs = []

    patterns = [
        r"\bE\s*(\d+)\s*:?\s*PC\s*(\d+\.\d+)\b",
        r"\bPC\s*(\d+\.\d+)\b",
        r"\b(\d+\.\d+)\s+(?:Determine|Analyse|Analyze|Calculate|Use|Construct|Apply)\b",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, rubric_text, flags=re.IGNORECASE):
            if len(match.groups()) == 2:
                element = match.group(1)
                pc_number = match.group(2)
                pc_code = f"E{element}:PC{pc_number}"
            else:
                pc_number = match.group(1)
                element = pc_number.split(".")[0]
                pc_code = f"E{element}:PC{pc_number}"

            pc_code = normalize_any_pc(pc_code)

            if pc_code not in pcs:
                pcs.append(pc_code)

    return pcs


def extract_sections_dynamic(text, pc_list):
    """
    Extracts sections based on PCs detected from the rubric.
    """

    headings = []

    for pc_code in pc_list:
        pc_number = pc_code.split("PC")[-1]
        element = pc_number.split(".")[0]

        patterns = [
            rf"E\s*{element}\s*:?\s*PC\s*{re.escape(pc_number)}",
            rf"PC\s*{re.escape(pc_number)}",
            rf"\b{re.escape(pc_number)}\b",
        ]

        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                headings.append({
                    "pc": pc_code,
                    "start": match.start(),
                    "end": match.end(),
                    "matched": match.group(0),
                })

    headings = sorted(headings, key=lambda x: x["start"])

    clean_headings = []

    for h in headings:
        if not clean_headings:
            clean_headings.append(h)
        else:
            last = clean_headings[-1]

            # Remove duplicate detections close to each other
            if h["pc"] == last["pc"] and abs(h["start"] - last["start"]) < 50:
                continue

            clean_headings.append(h)

    sections = {}

    for i, h in enumerate(clean_headings):
        start = h["start"]
        end = clean_headings[i + 1]["start"] if i + 1 < len(clean_headings) else len(text)

        pc_code = h["pc"]
        section_text = text[start:end].strip()

        if pc_code not in sections:
            sections[pc_code] = section_text
        else:
            sections[pc_code] += "\n\n" + section_text

    for pc_code in pc_list:
        sections.setdefault(pc_code, "")

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


def grade_pc_with_ai(
    pc_code,
    answer_text,
    answer_key_text,
    pc_rubric_text,
    max_retries=5,
):
    answer_text = answer_text[:4500]
    answer_key_text = answer_key_text[:4500]
    pc_rubric_text = pc_rubric_text[:4500]

    prompt = f"""
You are an assessor for MCT 122 Analyse Static Loads.

Assess the student's scanned answer using:
1. The official answer key to judge technical correctness.
2. The official rubric section to decide competency level and mark range.

Performance Criterion:
{pc_code}

Official Answer Key Section:
{answer_key_text}

Official Rubric Section:
{pc_rubric_text}

Student OCR Answer:
{answer_text}

Assessment task:
1. Compare the student's answer with the official answer key.
2. Check method, formula use, calculations, diagrams/FBDs, units, sign conventions, and final answer.
3. Use the rubric section to decide the level.
4. Assign a mark within the correct range.
5. Write short formal feedback for the Assessor Feedback section.

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
- Do not invent missing calculations, diagrams, values, units, or solution steps.
- If OCR is unclear or the answer section is missing, assign Not Yet Competent.
- If the answer key section is missing, rely on the rubric and visible student work only.
- The level must match the mark range.
- Feedback must be 1 to 2 sentences.
- Feedback must be specific and suitable for a formal assessment feedback form.
"""

    for attempt in range(max_retries):
        try:
            response = client.responses.create(
                model="gpt-4.1-mini",
                input=prompt,
                max_output_tokens=300,
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
                "feedback": "The answer could not be reliably graded because the extracted response or OCR text was unclear.",
            }

    return {
        "pc": pc_code,
        "mark": 0,
        "level": "Not Yet Competent",
        "feedback": "The answer could not be graded because the API rate limit was reached after several retries.",
    }


# =========================================================
# DOCX filling
# =========================================================

def fill_feedback_template(doc, student_name, student_id, pc_marks, feedback_rows):
    """
    Fills:
    - Student Name
    - ID No.
    - PC marks
    - Summative grade
    - Assessor feedback rows

    feedback_rows:
    {
        "PC3.1": ("Competent", "Feedback..."),
        "PC4.2": ("Competent with Merit", "Feedback...")
    }
    """

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = [cell.text.strip() for cell in cells]

            # Fill Student Name and ID No.
            for i, cell in enumerate(cells):
                text = cell.text.strip()

                if text in ["Student Name", "Student Name:"] and i + 1 < len(cells):
                    cells[i + 1].text = student_name

                if text in ["ID No.", "Student ID", "ID No", "Student ID:"] and i + 1 < len(cells):
                    cells[i + 1].text = student_id

            # Fill PC marks in Assessment Results table
            for pc_code, mark in pc_marks.items():
                pc_label = template_pc_label(pc_code)

                if pc_code in row_text or pc_label in row_text:
                    placed = False

                    # Prefer empty cells in the row
                    for cell in cells:
                        if cell.text.strip() == "":
                            cell.text = str(mark)
                            placed = True
                            break

                    # Fallback: replace old 60 if template has default 60
                    if not placed:
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

                    placed = False

                    for cell in cells:
                        if cell.text.strip() == "":
                            cell.text = str(summative)
                            placed = True
                            break

                    if not placed:
                        for cell in cells:
                            if cell.text.strip() == "60":
                                cell.text = str(summative)
                                break

            # Fill feedback rows if template already has PC feedback rows
            for pc_short, (level, feedback) in feedback_rows.items():
                if pc_short in row_text:
                    pc_index = row_text.index(pc_short)

                    if pc_index + 2 < len(cells):
                        cells[pc_index + 1].text = level
                        cells[pc_index + 2].text = feedback

    return doc


def add_feedback_table_if_missing(doc, feedback_rows):
    """
    If the template has only 'Assessor Feedback:' but no PC rows,
    this appends a feedback table after the document content.
    """

    has_feedback_rows = False

    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells]
            for pc_short in feedback_rows.keys():
                if pc_short in row_text:
                    has_feedback_rows = True

    if has_feedback_rows:
        return doc

    doc.add_paragraph("")
    doc.add_paragraph("Assessor Feedback")

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    hdr[0].text = "PC"
    hdr[1].text = "Level"
    hdr[2].text = "Feedback"

    for pc_short, (level, feedback) in feedback_rows.items():
        row = table.add_row().cells
        row[0].text = pc_short
        row[1].text = level
        row[2].text = feedback

    return doc


# =========================================================
# Upload interface
# =========================================================

classlist = st.file_uploader("Upload classlist Excel", type=["xlsx"])
template = st.file_uploader("Upload feedback template DOCX", type=["docx"])
rubric = st.file_uploader("Upload grading rubric", type=["docx", "pdf", "xlsx", "txt", "csv"])
answer_key = st.file_uploader("Upload official answer key", type=["docx", "pdf", "xlsx", "txt", "csv"])

pdfs = st.file_uploader(
    "Upload scanned student exam PDFs",
    type=["pdf"],
    accept_multiple_files=True,
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

    if not answer_key:
        st.error("Please upload the official answer key.")
        st.stop()

    if not pdfs:
        st.error("Please upload scanned student exam PDFs.")
        st.stop()

    df = pd.read_excel(classlist)

    st.subheader("Classlist Preview")
    st.dataframe(df.head())

    rubric_text = read_uploaded_file_as_text(rubric)
    answer_key_text = read_uploaded_file_as_text(answer_key)

    if not rubric_text.strip():
        st.error("The rubric file could not be read.")
        st.stop()

    if not answer_key_text.strip():
        st.error("The answer key file could not be read.")
        st.stop()

    pc_list = extract_pc_list_from_rubric(rubric_text)

    if not pc_list:
        st.error("No PCs were detected from the rubric. Please check that the rubric includes PC codes such as PC3.1, 3.1, or E3:PC3.1.")
        st.stop()

    st.subheader("PCs detected from rubric")
    st.write(pc_list)

    rubric_sections = extract_sections_dynamic(rubric_text, pc_list)
    answer_key_sections = extract_sections_dynamic(answer_key_text, pc_list)

    with st.expander("Preview extracted rubric sections"):
        for pc_code, section in rubric_sections.items():
            st.markdown(f"### {pc_code}")
            st.text(section[:1500] if section else "No section detected.")

    with st.expander("Preview extracted answer key sections"):
        for pc_code, section in answer_key_sections.items():
            st.markdown(f"### {pc_code}")
            st.text(section[:1500] if section else "No section detected.")

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
            pc_sections = extract_sections_dynamic(exam_text, pc_list)

            with st.expander(f"OCR and detected PC sections for {student_id}"):
                st.markdown("### Detected PC sections from student scan")
                detected = [pc for pc, txt in pc_sections.items() if txt.strip()]
                st.write(detected)

                st.markdown("### OCR preview")
                st.text(exam_text[:3000])

            pc_marks = {}
            feedback_rows = {}

            for pc_code in pc_list:
                pc_label = template_pc_label(pc_code)

                student_answer_text = pc_sections.get(pc_code, "")
                pc_rubric_text = rubric_sections.get(pc_code, "")
                pc_answer_key_text = answer_key_sections.get(pc_code, "")

                if not student_answer_text.strip():
                    result = {
                        "pc": pc_code,
                        "mark": 0,
                        "level": "Not Yet Competent",
                        "feedback": f"No clear answer section was detected for {pc_label}. The student should ensure the answer is clearly labelled and complete.",
                    }
                else:
                    result = grade_pc_with_ai(
                        pc_code=pc_code,
                        answer_text=student_answer_text,
                        answer_key_text=pc_answer_key_text,
                        pc_rubric_text=pc_rubric_text,
                    )

                pc_marks[pc_code] = result["mark"]
                feedback_rows[pc_label] = (
                    result["level"],
                    result["feedback"],
                )

                report_rows.append({
                    "Student ID": student_id,
                    "Student Name": student_name,
                    "PC": pc_code,
                    "Template PC Label": pc_label,
                    "Mark": result["mark"],
                    "Level": result["level"],
                    "Feedback": result["feedback"],
                    "Student Answer Detected": bool(student_answer_text.strip()),
                    "Rubric Section Detected": bool(pc_rubric_text.strip()),
                    "Answer Key Section Detected": bool(pc_answer_key_text.strip()),
                })

            template.seek(0)
            doc = Document(template)

            doc = fill_feedback_template(
                doc=doc,
                student_name=student_name,
                student_id=student_id,
                pc_marks=pc_marks,
                feedback_rows=feedback_rows,
            )

            doc = add_feedback_table_if_missing(doc, feedback_rows)

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
        mime="application/zip",
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
            mime="text/csv",
        )
