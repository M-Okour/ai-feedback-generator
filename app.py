import io
import re
import time
import zipfile
import pandas as pd
import streamlit as st
from docx import Document
from openai import OpenAI, RateLimitError, APIError, APITimeoutError


st.set_page_config(page_title="AI Feedback Generator", layout="wide")
st.title("AI Student Feedback Generator from Marks and Rubric")

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# =========================================================
# Basic helpers
# =========================================================

def get_level(mark):
    mark = float(mark)

    if mark < 60:
        return "Not Yet Competent"
    elif mark < 70:
        return "Competent"
    elif mark < 85:
        return "Competent with Merit"
    else:
        return "Competent with Distinction"


def get_grade_column_index(mark):
    """
    Template table columns:
    0 = LO & PC
    1 = Grade Classification
    2 = Not Yet Competent
    3 = Competent
    4 = Competent with Merit
    5 = Competent with Distinction
    """
    mark = float(mark)

    if mark < 60:
        return 2
    elif mark < 70:
        return 3
    elif mark < 85:
        return 4
    else:
        return 5


def get_first_name(full_name):
    if pd.isna(full_name):
        return "Student"

    full_name = str(full_name).strip()

    if not full_name:
        return "Student"

    return full_name.split()[0]


def normalize_pc_for_matching(text):
    """
    Converts:
    E1:PC1   -> PC1.1
    E1:PC2   -> PC1.2
    E2:PC1   -> PC2.1
    E3:PC3.1 -> PC3.1
    PC1.1    -> PC1.1
    PC3.1    -> PC3.1
    3.1      -> PC3.1
    """
    text = str(text).upper().replace(" ", "")

    # Ex:PCy -> PCx.y
    match = re.search(r"E(\d+):?PC(\d+)$", text)
    if match:
        element = match.group(1)
        pc = match.group(2)
        return f"PC{element}.{pc}"

    # E3:PC3.1 -> PC3.1
    match = re.search(r"E\d+:?PC(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    # PC3.1 -> PC3.1
    match = re.search(r"PC(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    # 3.1 -> PC3.1
    match = re.fullmatch(r"(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    return text


# =========================================================
# Read rubric DOCX
# =========================================================

def read_docx_text(file):
    file.seek(0)
    doc = Document(file)
    parts = []

    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())

    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip():
                parts.append(row_text)

    return "\n".join(parts)


def extract_pc_list_from_rubric(rubric_text):
    """
    Detects PCs dynamically from rubric.
    Supports:
    E1:PC1
    E3:PC3.1
    PC3.1
    3.1 Determine...
    """

    pcs = []

    patterns = [
        r"\bE\s*(\d+)\s*:?\s*PC\s*(\d+\.\d+|\d+)\b",
        r"\bPC\s*(\d+\.\d+|\d+)\b",
        r"\b(\d+\.\d+)\s+(?:Determine|Analyse|Analyze|Calculate|Use|Construct|Apply)\b",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, rubric_text, flags=re.IGNORECASE):
            if len(match.groups()) == 2:
                element = match.group(1)
                pc_number = match.group(2)

                if "." in pc_number:
                    pc_code = f"PC{pc_number}"
                else:
                    pc_code = f"PC{element}.{pc_number}"
            else:
                pc_number = match.group(1)

                if "." in pc_number:
                    pc_code = f"PC{pc_number}"
                else:
                    pc_code = f"PC{pc_number}"

            pc_code = normalize_pc_for_matching(pc_code)

            if pc_code not in pcs:
                pcs.append(pc_code)

    return pcs


def extract_rubric_sections(rubric_text, pc_list):
    sections = {}
    headings = []

    for pc in pc_list:
        pc_num = pc.replace("PC", "")

        # PC3.1 can appear as PC3.1 or 3.1
        patterns = [
            rf"\bPC\s*{re.escape(pc_num)}\b",
            rf"\b{re.escape(pc_num)}\b",
        ]

        for pattern in patterns:
            for match in re.finditer(pattern, rubric_text, flags=re.IGNORECASE):
                headings.append({
                    "pc": pc,
                    "start": match.start()
                })

    headings = sorted(headings, key=lambda x: x["start"])

    clean = []

    for h in headings:
        if not clean:
            clean.append(h)
        else:
            last = clean[-1]

            if h["pc"] == last["pc"] and abs(h["start"] - last["start"]) < 50:
                continue

            clean.append(h)

    for i, h in enumerate(clean):
        start = h["start"]
        end = clean[i + 1]["start"] if i + 1 < len(clean) else len(rubric_text)
        sections[h["pc"]] = rubric_text[start:end].strip()

    for pc in pc_list:
        sections.setdefault(pc, "")

    return sections


# =========================================================
# Excel helpers
# =========================================================

def find_student_columns(df):
    name_col = None
    id_col = None

    for col in df.columns:
        c = str(col).lower()

        if "name" in c and name_col is None:
            name_col = col

        if "id" in c and id_col is None:
            id_col = col

    return name_col, id_col


def get_pc_columns(df):
    """
    Finds PC mark columns dynamically.
    Accepts:
    E1:PC1
    E3:PC3.1
    PC1.1
    PC3.1
    3.1
    """

    pc_cols = {}

    for col in df.columns:
        col_text = str(col).upper().replace(" ", "")

        normalized = normalize_pc_for_matching(col_text)

        if re.fullmatch(r"PC\d+\.\d+", normalized):
            pc_cols[normalized] = col

    return pc_cols


# =========================================================
# AI feedback
# =========================================================

def generate_ai_feedback(first_name, pc, level, rubric_section, max_retries=4):
    prompt = f"""
You are writing formal assessment feedback for a diploma student.

Student first name:
{first_name}

Performance Criterion:
{pc}

Student competency level:
{level}

Relevant rubric criteria:
{rubric_section}

Write feedback for the Assessor Feedback section.

Rules:
- Use the student's first name once at the beginning.
- Write 1 to 2 sentences only.
- Match the feedback to the competency level.
- Explain what the student achieved and what should be improved.
- Use clear academic language.
- Do not mention AI, rubric file, automated marking, or the exact numerical mark.
- Do not include bullet points.
"""

    for attempt in range(max_retries):
        try:
            response = client.responses.create(
                model="gpt-4.1-mini",
                input=prompt,
                max_output_tokens=140
            )

            return response.output_text.strip()

        except RateLimitError:
            wait_time = 8 * (attempt + 1)
            st.warning(
                f"Rate limit reached while generating feedback for {pc}. "
                f"Retrying in {wait_time} seconds..."
            )
            time.sleep(wait_time)

        except (APIError, APITimeoutError):
            wait_time = 5 * (attempt + 1)
            st.warning(
                f"Temporary API issue for {pc}. "
                f"Retrying in {wait_time} seconds..."
            )
            time.sleep(wait_time)

        except Exception:
            break

    return (
        f"{first_name}, you achieved {level} for {pc}. "
        f"Please review the relevant solution steps, accuracy, presentation, and final answer to improve your performance."
    )


# =========================================================
# Template filling helpers
# =========================================================

def fill_name_and_id_in_table(table, student_name, student_id):
    for row in table.rows:
        cells = row.cells

        for i, cell in enumerate(cells):
            text = cell.text.strip()

            if text in ["Student Name", "Student Name:"]:
                if i + 1 < len(cells):
                    cells[i + 1].text = str(student_name)

            if text in ["ID No.", "ID No", "Student ID", "Student ID:"]:
                if i + 1 < len(cells):
                    cells[i + 1].text = str(student_id)


def fill_marks_in_assessment_table(table, pc_marks):
    for row in table.rows:
        cells = row.cells

        for cell in cells:
            cell_pc = normalize_pc_for_matching(cell.text)

            if cell_pc in pc_marks:
                mark = pc_marks[cell_pc]
                target_col = get_grade_column_index(mark)

                if target_col < len(cells):
                    cells[target_col].text = str(int(mark))


def fill_summative_grade_in_table(table, pc_marks):
    marks = [float(m) for m in pc_marks.values()]

    if not marks:
        return

    if any(m < 60 for m in marks):
        summative = min(marks)
    else:
        summative = round(sum(marks) / len(marks))

    for row in table.rows:
        cells = row.cells

        for i, cell in enumerate(cells):
            if "Summative Assessment Grade" in cell.text:
                if i + 1 < len(cells):
                    cells[i + 1].text = str(int(summative))
                return


def fill_template(doc, student_name, student_id, feedback_rows):
    """
    Template has two major sections:

    1. Student Acknowledgment
       - fill Student Name
       - fill ID No.
       - do not fill marks

    2. Assessment Results and Feedback
       - fill Student Name
       - fill ID No.
       - fill PC marks in grade-band columns
       - fill Summative Assessment Grade
       - add feedback table
    """

    pc_marks = {
        normalize_pc_for_matching(row["PC"]): row["Mark"]
        for row in feedback_rows
    }

    for table in doc.tables:
        table_text = "\n".join(
            cell.text.strip()
            for row in table.rows
            for cell in row.cells
        )

        is_ack_table = (
            "Portfolio Evidence Requirements" in table_text
            or "Student Acknowledgment" in table_text
            or "Unit Title/s" in table_text
        )

        is_assessment_table = (
            "Assessment Results" in table_text
            or "PC Grade" in table_text
            or "Summative Assessment Grade" in table_text
            or "Grade Classification" in table_text
        )

        if is_ack_table:
            fill_name_and_id_in_table(table, student_name, student_id)
            continue

        if is_assessment_table:
            fill_name_and_id_in_table(table, student_name, student_id)
            fill_marks_in_assessment_table(table, pc_marks)
            fill_summative_grade_in_table(table, pc_marks)

    # Add Assessor Feedback table WITHOUT marks
    doc.add_paragraph("")
    doc.add_paragraph("Assessor Feedback:")

    feedback_table = doc.add_table(rows=1, cols=3)
    feedback_table.style = "Table Grid"

    header = feedback_table.rows[0].cells
    header[0].text = "PC"
    header[1].text = "Level"
    header[2].text = "Feedback"

    for row_data in feedback_rows:
        row = feedback_table.add_row().cells
        row[0].text = row_data["PC"]
        row[1].text = row_data["Level"]
        row[2].text = row_data["Feedback"]

    return doc


# =========================================================
# Streamlit upload interface
# =========================================================

rubric_file = st.file_uploader("Upload rubric.docx", type=["docx"])
classlist_file = st.file_uploader("Upload classlist.xlsx", type=["xlsx"])
template_file = st.file_uploader("Upload Template_Feedback.docx", type=["docx"])


# =========================================================
# Main process
# =========================================================

if st.button("Generate AI Feedback Files"):
    if not rubric_file or not classlist_file or not template_file:
        st.error("Please upload rubric, classlist, and feedback template.")
        st.stop()

    rubric_text = read_docx_text(rubric_file)
    pc_list = extract_pc_list_from_rubric(rubric_text)
    rubric_sections = extract_rubric_sections(rubric_text, pc_list)

    df = pd.read_excel(classlist_file)

    name_col, id_col = find_student_columns(df)
    pc_cols = get_pc_columns(df)

    st.subheader("Detected Setup")
    st.write("Name column:", name_col)
    st.write("ID column:", id_col)
    st.write("PCs from rubric:", pc_list)
    st.write("PC mark columns from classlist:", pc_cols)

    if not name_col or not id_col:
        st.error("Could not detect student name or ID column.")
        st.stop()

    if not pc_cols:
        st.error("Could not detect PC mark columns. Use headers like E1:PC1, PC1.1, PC3.1, or 3.1.")
        st.stop()

    zip_buffer = io.BytesIO()
    summary_rows = []
    progress = st.progress(0)

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, student in df.iterrows():
            student_name = student[name_col]
            student_id = student[id_col]
            first_name = get_first_name(student_name)

            feedback_rows = []

            for pc in pc_list:
                normalized_pc = normalize_pc_for_matching(pc)

                if normalized_pc not in pc_cols:
                    continue

                mark = student[pc_cols[normalized_pc]]

                if pd.isna(mark):
                    continue

                mark = int(round(float(mark)))
                level = get_level(mark)
                rubric_section = rubric_sections.get(pc, "")

                feedback = generate_ai_feedback(
                    first_name=first_name,
                    pc=pc,
                    level=level,
                    rubric_section=rubric_section
                )

                feedback_rows.append({
                    "PC": pc,
                    "Mark": mark,
                    "Level": level,
                    "Feedback": feedback
                })

                summary_rows.append({
                    "Student Name": student_name,
                    "Student ID": student_id,
                    "PC": pc,
                    "Mark": mark,
                    "Level": level,
                    "Feedback": feedback
                })

            template_file.seek(0)
            doc = Document(template_file)

            doc = fill_template(
                doc=doc,
                student_name=student_name,
                student_id=student_id,
                feedback_rows=feedback_rows
            )

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            output_name = f"{student_id}.docx"
            zip_file.writestr(output_name, doc_buffer.getvalue())

            progress.progress((index + 1) / len(df))

    zip_buffer.seek(0)

    st.success("AI feedback files generated successfully.")

    st.download_button(
        label="Download Feedback ZIP",
        data=zip_buffer.getvalue(),
        file_name="student_ai_feedback_files.zip",
        mime="application/zip"
    )

    if summary_rows:
        summary_df = pd.DataFrame(summary_rows)

        st.subheader("Feedback Summary")
        st.dataframe(summary_df)

        csv_buffer = io.StringIO()
        summary_df.to_csv(csv_buffer, index=False)

        st.download_button(
            label="Download Feedback Summary CSV",
            data=csv_buffer.getvalue(),
            file_name="feedback_summary.csv",
            mime="text/csv"
        )
