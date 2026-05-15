import io
import re
import time
import zipfile
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from openai import OpenAI, RateLimitError, APIError, APITimeoutError


st.set_page_config(page_title="AI Feedback Generator", layout="wide")
st.title("AI Student Feedback Generator from Marks and Rubric")

feedback_mode = st.radio(
    "Feedback generation mode",
    options=[
        "Generate feedback for all PCs",
        "Generate feedback only for Not Yet Competent PCs"
    ],
    index=0
)

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


def get_level(mark):
    mark = float(mark)
    if mark < 60:
        return "Not Yet Competent"
    elif mark < 70:
        return "Competent"
    elif mark < 85:
        return "Competent with Merit"
    return "Competent with Distinction"


def get_grade_column_index(mark):
    mark = float(mark)
    if mark < 60:
        return 2
    elif mark < 70:
        return 3
    elif mark < 85:
        return 4
    return 5


def get_first_name(full_name):
    if pd.isna(full_name):
        return "Student"
    full_name = str(full_name).strip()
    return full_name.split()[0] if full_name else "Student"


def normalize_pc_for_matching(text):
    text = str(text).upper().replace(" ", "")

    match = re.search(r"E(\d+):?PC(\d+)$", text)
    if match:
        return f"PC{match.group(1)}.{match.group(2)}"

    match = re.search(r"E\d+:?PC(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    match = re.search(r"PC(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    match = re.fullmatch(r"(\d+\.\d+)", text)
    if match:
        return f"PC{match.group(1)}"

    return text


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
                pc_code = f"PC{pc_number}" if "." in pc_number else f"PC{element}.{pc_number}"
            else:
                pc_number = match.group(1)
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
        patterns = [
            rf"\bPC\s*{re.escape(pc_num)}\b",
            rf"\b{re.escape(pc_num)}\b",
        ]

        for pattern in patterns:
            for match in re.finditer(pattern, rubric_text, flags=re.IGNORECASE):
                headings.append({"pc": pc, "start": match.start()})

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
    pc_cols = {}

    for col in df.columns:
        col_text = str(col).upper().replace(" ", "")
        normalized = normalize_pc_for_matching(col_text)

        if re.fullmatch(r"PC\d+\.\d+", normalized):
            pc_cols[normalized] = col

    return pc_cols


def generate_ai_feedback(first_name, pc, level, rubric_section, max_retries=4):
    prompt = f"""
You are an experienced engineering assessor writing feedback for a diploma student.

Student first name:
{first_name}

Performance Criterion:
{pc}

Student competency level:
{level}

Relevant rubric criteria:
{rubric_section}

Write feedback for the Assessor Feedback section.

Style requirements:
- Sound natural, human, and supportive.
- Keep a professional assessor tone.
- Use the student's first name once at the beginning.
- Write 1 to 2 sentences only.
- Do not sound robotic or generic.
- Mention one specific strength based on the level.
- Mention one clear improvement area.
- Do not mention AI, automated marking, rubric file, or the exact numerical mark.
- Do not use bullet points.
- Do not overpraise if the level is low.
- Keep the language suitable for official assessment feedback.
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
            st.warning(f"Rate limit reached while generating feedback for {pc}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

        except (APIError, APITimeoutError):
            wait_time = 5 * (attempt + 1)
            st.warning(f"Temporary API issue for {pc}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

        except Exception:
            break

    return (
        f"{first_name}, you achieved {level} for {pc}. "
        f"Please review the relevant solution steps, accuracy, presentation, and final answer to improve your performance."
    )


# =========================================================
# Template filling helpers - safer for merged cells
# =========================================================

def row_texts(row):
    return [cell.text.strip() for cell in row.cells]


def row_full_text(row):
    return " | ".join(row_texts(row))


def is_empty_cell(cell):
    return cell.text.strip() == ""


def fill_first_empty_cell_in_row(row, value, start_index=0):
    cells = row.cells

    for i in range(start_index, len(cells)):
        if is_empty_cell(cells[i]):
            cells[i].text = str(value)
            return True

    return False


def fill_adjacent_or_empty(row, label_keywords, value):
    """
    Safer than cells[i+1] for merged-cell templates.
    It tries:
    1. cell immediately after the label
    2. first empty cell after the label
    3. first empty cell in the row
    """

    cells = row.cells

    for i, cell in enumerate(cells):
        text = cell.text.strip()

        if any(keyword.lower() in text.lower() for keyword in label_keywords):

            # Try immediate right cell
            if i + 1 < len(cells) and is_empty_cell(cells[i + 1]):
                cells[i + 1].text = str(value)
                return True

            # Try any empty cell after label
            if fill_first_empty_cell_in_row(row, value, start_index=i + 1):
                return True

            # Try any empty cell in row
            if fill_first_empty_cell_in_row(row, value, start_index=0):
                return True

    return False


def fill_name_and_id_in_table(table, student_name, student_id):
    """
    Handles both:
    Student Name | [empty]
    ID No.       | [empty]

    Safer for merged tables because it searches empty target cells.
    """

    for row in table.rows:
        fill_adjacent_or_empty(
            row=row,
            label_keywords=["Student Name"],
            value=student_name
        )

        fill_adjacent_or_empty(
            row=row,
            label_keywords=["ID No.", "ID No", "Student ID"],
            value=student_id
        )


def get_row_pc(row, pc_marks):
    """
    Find whether this row belongs to one of the PCs.
    """

    for cell in row.cells:
        normalized = normalize_pc_for_matching(cell.text)

        if normalized in pc_marks:
            return normalized

    return None


def clear_possible_grade_cells(row):
    """
    Clears likely grade cells in a PC row before filling the correct one.
    This prevents old placeholder marks from remaining.
    Assumes columns 2-5 are grade-band columns when available.
    """

    cells = row.cells

    for idx in [2, 3, 4, 5]:
        if idx < len(cells):
            # Do not clear if the cell itself contains a PC code
            if not normalize_pc_for_matching(cells[idx].text).startswith("PC"):
                cells[idx].text = ""


def fill_marks_in_assessment_table(table, pc_marks):
    """
    Fills the mark in the correct grade-band column.

    Expected visual structure:
    col 0 = LO & PC
    col 1 = Grade Classification
    col 2 = Not Yet Competent
    col 3 = Competent
    col 4 = Competent with Merit
    col 5 = Competent with Distinction

    Merged header cells may exist, but PC rows usually still preserve row cells.
    """

    for row in table.rows:
        row_pc = get_row_pc(row, pc_marks)

        if row_pc is None:
            continue

        mark = pc_marks[row_pc]
        target_col = get_grade_column_index(mark)

        clear_possible_grade_cells(row)

        cells = row.cells

        if target_col < len(cells):
            cells[target_col].text = str(int(mark))
        else:
            # fallback: first empty cell after the PC cell
            fill_first_empty_cell_in_row(row, str(int(mark)), start_index=1)


def fill_summative_grade_in_table(table, pc_marks):
    marks = [float(m) for m in pc_marks.values()]

    if not marks:
        return

    if any(m < 60 for m in marks):
        summative = min(marks)
    else:
        summative = round(sum(marks) / len(marks))

    for row in table.rows:
        text = row_full_text(row)

        if "Summative Assessment Grade" in text:
            filled = fill_adjacent_or_empty(
                row=row,
                label_keywords=["Summative Assessment Grade"],
                value=int(summative)
            )

            if not filled:
                cells = row.cells
                if len(cells) > 1:
                    cells[-1].text = str(int(summative))

            return


def set_cell_width(cell, width_inches):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")

    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)

    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def build_feedback_table_in_cell(cell, feedback_rows):
    cell.text = ""

    table = cell.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.autofit = False

    header = table.rows[0].cells
    header[0].text = "PC"
    header[1].text = "Level"
    header[2].text = "Feedback"

    set_cell_width(header[0], 0.75)
    set_cell_width(header[1], 1.35)
    set_cell_width(header[2], 5.4)

    for row_data in feedback_rows:
        if not str(row_data["Feedback"]).strip():
            continue

        row = table.add_row().cells
        row[0].text = row_data["PC"]
        row[1].text = row_data["Level"]
        row[2].text = row_data["Feedback"]

        set_cell_width(row[0], 0.75)
        set_cell_width(row[1], 1.35)
        set_cell_width(row[2], 5.4)

    return table


def insert_feedback_table_at_assessor_feedback(doc, feedback_rows):
    feedback_inserted = False

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells

            for i, cell in enumerate(cells):
                if "Assessor Feedback:" in cell.text or "Assessor Feedback" in cell.text:
                    # Prefer an empty cell after the label
                    target_cell = None

                    for j in range(i + 1, len(cells)):
                        if is_empty_cell(cells[j]):
                            target_cell = cells[j]
                            break

                    if target_cell is None:
                        target_cell = cells[i]

                    build_feedback_table_in_cell(
                        target_cell,
                        feedback_rows
                    )

                    feedback_inserted = True
                    break

            if feedback_inserted:
                break

        if feedback_inserted:
            break

    if not feedback_inserted:
        doc.add_paragraph("")
        doc.add_paragraph("Assessor Feedback:")

        fallback_table = doc.add_table(rows=1, cols=3)
        fallback_table.style = "Table Grid"
        fallback_table.autofit = False

        header = fallback_table.rows[0].cells
        header[0].text = "PC"
        header[1].text = "Level"
        header[2].text = "Feedback"

        set_cell_width(header[0], 0.75)
        set_cell_width(header[1], 1.35)
        set_cell_width(header[2], 5.4)

        for row_data in feedback_rows:
            if not str(row_data["Feedback"]).strip():
                continue

            row = fallback_table.add_row().cells
            row[0].text = row_data["PC"]
            row[1].text = row_data["Level"]
            row[2].text = row_data["Feedback"]

            set_cell_width(row[0], 0.75)
            set_cell_width(row[1], 1.35)
            set_cell_width(row[2], 5.4)

    return doc


def fill_template(doc, student_name, student_id, feedback_rows):
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

    doc = insert_feedback_table_at_assessor_feedback(doc, feedback_rows)

    return doc


rubric_file = st.file_uploader("Upload rubric.docx", type=["docx"])
classlist_file = st.file_uploader("Upload classlist.xlsx", type=["xlsx"])
template_file = st.file_uploader("Upload Template_Feedback.docx", type=["docx"])


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

                should_generate_feedback = (
                    feedback_mode == "Generate feedback for all PCs"
                    or level == "Not Yet Competent"
                )

                if should_generate_feedback:
                    feedback = generate_ai_feedback(
                        first_name=first_name,
                        pc=pc,
                        level=level,
                        rubric_section=rubric_section
                    )
                else:
                    feedback = ""

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
