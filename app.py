
import os
import io
import re
import sqlite3
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from google import genai
from docx import Document
from rapidfuzz import fuzz
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

# ---------------- CONFIGURATION ----------------

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
MODEL = "gemini-3.5-flash-lite"

st.set_page_config(
    page_title="BizFlow AI",
    page_icon="🤖",
    layout="wide"
)

# ---------------- DESIGN ----------------

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: #F1F5FB;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#101B35,#17365F);
}

[data-testid="stSidebar"] * {
    color: white;
}

.hero {
    background: linear-gradient(120deg,#101B35,#1759A6,#087F8C);
    padding: 30px;
    border-radius: 18px;
    color: white;
    margin-bottom: 25px;
}

.hero h1 {
    color: white;
    font-size: 36px;
    font-weight: 800;
}

.hero p {
    color: #DCEBFF;
}

[data-testid="stMetric"] {
    background: white;
    padding: 18px;
    border-radius: 13px;
    border-top: 4px solid #087F8C;
    box-shadow: 0 4px 15px #DCE4EF;
}

.stButton button {
    background: linear-gradient(90deg,#1759A6,#087F8C);
    color: white;
    border-radius: 9px;
    font-weight: 700;
}

.section {
    color: #172B50;
    font-size: 24px;
    font-weight: 800;
    margin-top: 20px;
}

.footer {
    text-align: center;
    padding: 25px;
    color: #64748B;
}
</style>
""", unsafe_allow_html=True)


# ---------------- HELPER FUNCTIONS ----------------

def ensure_database():

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS requirements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer TEXT,
            requirement TEXT,
            status TEXT
        )
    """)

    # Safely upgrade existing database
    cursor.execute("PRAGMA table_info(requirements)")
    columns = [row[1] for row in cursor.fetchall()]

    if "priority" not in columns:
        cursor.execute("""
            ALTER TABLE requirements
            ADD COLUMN priority TEXT DEFAULT 'Must Have'
        """)

    conn.commit()
    conn.close()


def save_business_requirement(
    customer,
    requirement,
    status,
    priority="Must Have"
):

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO requirements
        (customer, requirement, status, priority)
        VALUES (?, ?, ?, ?)
    """, (customer, requirement, status, priority))

    conn.commit()
    conn.close()


def get_business_requirements(search_term=""):

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM requirements
        WHERE customer LIKE ?
        ORDER BY id DESC
    """, (f"%{search_term}%",))

    data = cursor.fetchall()
    conn.close()

    return data


def update_business_requirement(
    requirement_id,
    customer,
    requirement,
    status,
    priority
):

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE requirements
        SET customer = ?,
            requirement = ?,
            status = ?,
            priority = ?
        WHERE id = ?
    """, (
        customer,
        requirement,
        status,
        priority,
        requirement_id
    ))

    conn.commit()
    conn.close()


ensure_database()


def normalize(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def normalize_phone(value):
    if pd.isna(value):
        return ""
    return "".join(ch for ch in str(value) if ch.isdigit())


def find_column(df, aliases):
    for col in df.columns:
        cleaned = str(col).strip().lower().replace(" ", "_")
        if cleaned in aliases:
            return col
    return None


def detect_duplicates(df):

    results = []
    pair_reasons = {}

    email_col = find_column(df, ["email", "email_address"])
    phone_col = find_column(df, ["phone", "mobile", "mobile_number"])
    name_col = find_column(df, ["customer_name", "name", "full_name"])
    company_col = find_column(df, ["company", "organization", "organisation"])

    def add_match(i, j, reason):
        key = (min(i, j), max(i, j))
        pair_reasons.setdefault(key, set()).add(reason)

    # Exact email matches
    if email_col:
        emails = df[email_col].apply(normalize)

        for i in range(len(df)):
            if emails.iloc[i]:
                for j in range(i + 1, len(df)):
                    if emails.iloc[i] == emails.iloc[j]:
                        add_match(i, j, "Matching Email")

    # Exact phone matches
    if phone_col:
        phones = df[phone_col].apply(normalize_phone)

        for i in range(len(df)):
            if phones.iloc[i]:
                for j in range(i + 1, len(df)):
                    if phones.iloc[i] == phones.iloc[j]:
                        add_match(i, j, "Matching Phone")

    # Similar names in same company (limited for performance)
    if name_col and company_col and len(df) <= 500:

        for i in range(len(df)):
            name1 = normalize(df.iloc[i][name_col])
            company1 = normalize(df.iloc[i][company_col])

            if not name1 or not company1:
                continue

            for j in range(i + 1, len(df)):

                name2 = normalize(df.iloc[j][name_col])
                company2 = normalize(df.iloc[j][company_col])

                if company1 == company2 and name2:
                    similarity = fuzz.ratio(name1, name2)

                    if 90 <= similarity < 100:
                        add_match(i, j, "Similar Customer Name")

    for (i, j), reasons in pair_reasons.items():
        results.append({
            "Record 1": i + 1,
            "Record 2": j + 1,
            "Customer 1": df.iloc[i][name_col] if name_col else "Unknown",
            "Customer 2": df.iloc[j][name_col] if name_col else "Unknown",
            "Reason": ", ".join(sorted(reasons))
        })

    return pd.DataFrame(results)


def generate_ai(prompt):

    if not API_KEY:
        raise ValueError("Gemini API key is missing. Check your .env file.")

    client = genai.Client(api_key=API_KEY)

    response = client.models.generate_content(
        model=MODEL,
        contents=prompt
    )

    if not response.text:
        raise ValueError("Gemini returned an empty response.")

    return response.text



def create_brd(content):

    doc = Document()

    # PAGE SETTINGS
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    # NORMAL STYLE
    styles = doc.styles

    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].paragraph_format.space_after = Pt(6)

    # HEADING STYLES
    for style_name, size in [
        ("Heading 1", 16),
        ("Heading 2", 13),
        ("Heading 3", 11)
    ]:
        styles[style_name].font.name = "Calibri"
        styles[style_name].font.size = Pt(size)
        styles[style_name].font.bold = True

    # HEADER
    header = section.header
    header.paragraphs[0].text = "BIZFLOW AI | BUSINESS REQUIREMENTS DOCUMENT"
    header.paragraphs[0].style = "Caption"

    # FOOTER
    footer = section.footer
    footer.paragraphs[0].text = "Confidential Business Document | Generated by BizFlow AI"

    # TITLE PAGE
    title = doc.add_heading("BUSINESS REQUIREMENTS DOCUMENT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph("Prepared using BizFlow AI")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")
    doc.add_paragraph("")

    project_title = doc.add_paragraph()
    project_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = project_title.add_run("Business Analysis Report")
    run.bold = True
    run.font.size = Pt(20)

    doc.add_paragraph("")
    date_para = doc.add_paragraph(
        f"Generated on: {datetime.now().strftime('%d %B %Y')}"
    )
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # CONTENT FORMATTING
    lines = content.split("\n")
    i = 0

    while i < len(lines):

        line = lines[i].strip()

        if not line:
            i += 1
            continue

        # MARKDOWN TABLE
        if line.startswith("|"):

            table_lines = []

            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1

            rows = []

            for table_line in table_lines:

                cells = [
                    cell.strip()
                    for cell in table_line.strip("|").split("|")
                ]

                # Ignore Markdown separator rows
                if all(
                    set(cell.replace(" ", "")) <= {"-", ":"}
                    for cell in cells
                ):
                    continue

                rows.append(cells)

            if rows:

                column_count = max(len(row) for row in rows)

                table = doc.add_table(
                    rows=0,
                    cols=column_count
                )

                table.style = "Table Grid"

                for row_index, row_data in enumerate(rows):

                    cells = table.add_row().cells

                    for col_index in range(column_count):

                        value = (
                            row_data[col_index]
                            if col_index < len(row_data)
                            else ""
                        )

                        cells[col_index].text = value

                        if row_index == 0:
                            for run in cells[col_index].paragraphs[0].runs:
                                run.bold = True

            continue

        # HEADINGS
        if line.startswith("### "):
            doc.add_heading(line[4:], level=3)

        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)

        elif line.startswith("# "):
            doc.add_heading(line[2:], level=1)

        # BULLET POINTS
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(
                line[2:],
                style="List Bullet"
            )

        # NUMBERED LISTS
        elif re.match(r"^\d+[\.\)]\s+", line):

            text_value = re.sub(
                r"^\d+[\.\)]\s+",
                "",
                line
            )

            doc.add_paragraph(
                text_value,
                style="List Number"
            )

        # NORMAL PARAGRAPH
        else:
            doc.add_paragraph(line)

        i += 1

    # SAVE DOCUMENT
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer

def create_traceability_table():

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS requirement_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            requirement_id TEXT,
            requirement_type TEXT,
            description TEXT,
            priority TEXT,
            acceptance_criteria TEXT,
            test_case TEXT,
            test_status TEXT
        )
    """)

    conn.commit()
    conn.close()


def save_traceability_item(
    project_id,
    requirement_id,
    requirement_type,
    description,
    priority,
    acceptance_criteria,
    test_case,
    test_status
):

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO requirement_items (
            project_id,
            requirement_id,
            requirement_type,
            description,
            priority,
            acceptance_criteria,
            test_case,
            test_status
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        project_id,
        requirement_id,
        requirement_type,
        description,
        priority,
        acceptance_criteria,
        test_case,
        test_status
    ))

    conn.commit()
    conn.close()


def get_traceability_items(project_id):

    conn = sqlite3.connect("bizflow.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            requirement_id,
            requirement_type,
            description,
            priority,
            acceptance_criteria,
            test_case,
            test_status
        FROM requirement_items
        WHERE project_id = ?
        ORDER BY id DESC
    """, (project_id,))

    data = cursor.fetchall()
    conn.close()

    return data


create_traceability_table()



# ---------------- SIDEBAR ----------------

with st.sidebar:

    st.markdown("# 🤖 BizFlow AI")
    st.caption("INTELLIGENT BUSINESS ASSISTANT")

    st.divider()

    page = st.radio(
    "MAIN MENU",
    [
        "🏠 Dashboard",
        "📊 CRM Data Analyzer",
        "📝 Requirement Analyzer",
        "🤖 AI Assistant",
        "🔄 Process Flow Analyzer",
        "🗄️ Requirements Database"
    ]
)

    st.divider()

    st.markdown("### DATASET UPLOAD")

    uploaded_file = st.file_uploader(
        "Upload CSV or Excel",
        type=["csv", "xlsx"]
    )

    st.divider()

    st.markdown("### SYSTEM STATUS")

    st.success("Application Ready")

    if API_KEY:
        st.success("Gemini API Configured")
    else:
        st.warning("Gemini API Key Missing")


# ---------------- LOAD DATA ----------------

df = None

if uploaded_file:

    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        if df.empty:
            st.warning("The uploaded dataset is empty.")
            df = None

    except Exception as e:
        st.error(f"Unable to read the file: {e}")


# ---------------- DATA METRICS ----------------

total_records = 0
total_columns = 0
missing_values = 0
duplicate_rows = 0
completeness = 0

duplicate_matches = pd.DataFrame()

if df is not None:

    total_records = len(df)
    total_columns = len(df.columns)
    missing_values = int(df.isna().sum().sum())
    duplicate_rows = int(df.duplicated().sum())

    total_cells = total_records * total_columns

    if total_cells:
        completeness = round(
            (1 - missing_values / total_cells) * 100, 1
        )

    duplicate_matches = detect_duplicates(df)


# ---------------- HEADER ----------------

st.markdown("""
<div class="hero">
    <h1>🤖 BizFlow AI</h1>
    <p>AI-Powered Business Analysis & CRM Data Intelligence</p>
    <p>Analyze. Automate. Understand. Improve.</p>
</div>
""", unsafe_allow_html=True)


# =================================================
# 1. DASHBOARD
# =================================================
if page == "🏠 Dashboard":

    st.markdown(
        '<div class="section">📊 Business Overview</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Customer Records", total_records)
    c2.metric("Duplicate Rows", duplicate_rows)
    c3.metric("Missing Values", missing_values)
    c4.metric("Data Completeness", f"{completeness}%")

    st.write("")

    left, right = st.columns(2)

    with left:
        st.markdown("### 📁 Dataset Information")

        if df is not None:
            st.write(f"Total Records: {total_records}")
            st.write(f"Total Columns: {total_columns}")
            st.write(f"Duplicate Matches: {len(duplicate_matches)}")
            st.progress(completeness / 100)
        else:
            st.info("Upload a customer dataset from the sidebar.")

    with right:
        st.markdown("### 🚀 Platform Modules")

        st.info("📊 CRM Data Analyzer - Customer data quality")
        st.success("📝 Requirement Analyzer - Business documentation")
        st.warning("🤖 AI Assistant - AI-powered explanations")

    st.markdown("### 💡 Project Purpose")

    st.write(
        "BizFlow AI helps organizations analyze customer data, "
        "understand business requirements and reduce manual "
        "documentation work through AI-assisted analysis."
    )



elif page == "🔄 Process Flow Analyzer":

    st.markdown("## 🔄 Process Flow Analyzer")

    st.write(
        "Understand the existing business process and visualize "
        "how automation can improve it."
    )

    st.markdown("### 📌 Business Scenario")

    st.info(
        "A sales team manages customer leads manually using Excel. "
        "Some leads are missed because follow-ups are not tracked."
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🔴 AS-IS Process")
        st.caption("Current manual workflow")

        as_is = """
        flowchart TD
            A[Customer Enquiry] --> B[Employee Records Lead in Excel]
            B --> C[Employee Manually Checks Leads]
            C --> D{Follow-up Done?}
            D -->|No| E[Lead May Be Missed]
            D -->|Yes| F[Manual Status Update]
        """

        st.mermaid_chart(as_is)

    with col2:
        st.markdown("### 🟢 TO-BE Process")
        st.caption("Proposed automated workflow")

        to_be = """
        flowchart TD
            A[Customer Enquiry] --> B[CRM Stores Lead]
            B --> C[Automatic Follow-up Reminder]
            C --> D[Employee Contacts Customer]
            D --> E[Update Lead Status]
            E --> F[Manager Views Dashboard]
        """

        st.mermaid_chart(to_be)

    st.markdown("### 💡 Business Improvements")

    st.success("Reduced manual work and missed follow-ups.")

    st.write("- Automatic follow-up reminders")
    st.write("- Centralized customer information")
    st.write("- Better lead tracking")
    st.write("- Improved management visibility")





# =================================================
# 3. CRM DATA ANALYZER
# =================================================

elif page == "📊 CRM Data Analyzer":

    st.markdown("## 🔎 CRM Data Quality Analysis")

    if df is None:
        st.warning("Upload a CSV or Excel file using the sidebar.")
    else:

        st.markdown("### Customer Dataset Preview")

        st.dataframe(
            df.head(100),
            use_container_width=True
        )

        st.markdown("### Data Quality Metrics")

        a, b, c, d = st.columns(4)

        a.metric("Records", total_records)
        b.metric("Columns", total_columns)
        c.metric("Missing Cells", missing_values)
        d.metric("Completeness", f"{completeness}%")

        st.markdown("### Column-wise Missing Values")

        quality = pd.DataFrame({
            "Column": df.columns,
            "Missing Values": df.isna().sum().values,
            "Missing %": (
                df.isna().mean().values * 100
            ).round(2)
        })

        st.dataframe(quality, use_container_width=True)

        missing_chart = df.isna().sum()
        missing_chart = missing_chart[missing_chart > 0]

        if not missing_chart.empty:
            st.bar_chart(missing_chart)
        else:
            st.success("No missing values detected.")

        st.markdown("### 🔍 Intelligent Duplicate Detection")

        if not duplicate_matches.empty:

            st.warning(
                f"{len(duplicate_matches)} possible duplicate pairs found."
            )

            st.dataframe(
                duplicate_matches,
                use_container_width=True
            )

            duplicate_csv = duplicate_matches.to_csv(
                index=False
            ).encode("utf-8")

            st.download_button(
                "Download Duplicate Report",
                duplicate_csv,
                "duplicate_report.csv",
                "text/csv"
            )

        else:
            st.success("No matching email, phone or similar-name pairs found.")

        st.markdown("### 📥 Download Clean Dataset")

        clean_df = df.drop_duplicates()

        st.download_button(
            "Download CSV Without Exact Duplicate Rows",
            clean_df.to_csv(index=False).encode("utf-8"),
            "clean_customer_data.csv",
            "text/csv"
        )


# =================================================
# 4. REQUIREMENT ANALYZER
# =================================================

elif page == "📝 Requirement Analyzer":

    st.markdown("## 📝 AI Business Requirement Analyzer")

    st.write(
        "Convert customer meeting notes into structured "
        "business requirements using Gemini AI."
    )
    customer_name = st.text_input(
    "Customer / Project Name",
    placeholder="Example: ABC Online Booking System"
    )

    notes = st.text_area(
        "Customer Meeting Notes",
        height=230,
        placeholder=(
            "Example: Customer wants an online booking system. "
            "They need customer registration, payment options, "
            "booking confirmation and email notifications..."
        )
    )

    if st.button(
        "✨ Analyze Business Requirements",
        use_container_width=True
    ):

        if not customer_name.strip():
         st.warning("Please enter the customer or project name.")

        elif not notes.strip():
         st.warning("Please enter customer meeting notes.")

        else:

            
            prompt = f"""
You are an experienced Senior Business Analyst.

Analyze the customer meeting notes and generate
a complete professional Business Analysis Report.

IMPORTANT RULES:
- Use simple professional English.
- Do not invent customer requirements.
- Do not assume missing business information.
- Mark unavailable details as To Be Confirmed.
- Make requirements clear and testable.

CUSTOMER MEETING NOTES:
{notes}

Generate the following sections:

1. EXECUTIVE SUMMARY
Summarize the business requirement.

2. EXISTING BUSINESS PROBLEMS
Identify the current business challenges.

3. CUSTOMER NEEDS
List customer expectations.

4. FUNCTIONAL REQUIREMENTS
List system functionalities.
Assign IDs: FR01, FR02, FR03.

5. NON-FUNCTIONAL REQUIREMENTS
Include security, performance, usability
and reliability requirements.
Assign IDs: NFR01, NFR02.

6. USER STORIES
Use this format:
As a [user], I want [feature],
so that [business benefit].

Generate relevant user stories.

7. ACCEPTANCE CRITERIA
For each user story, generate testable conditions.
Use Given, When, Then format.

8. REQUIREMENT PRIORITY
Classify requirements as:
High, Medium or Low.
Explain the priority.

9. AS-IS PROCESS
Explain the current business workflow step by step.

10. TO-BE PROCESS
Explain the proposed improved workflow step by step.

11. BUSINESS RISKS
Identify risks and suggest solutions.

12. CUSTOMER CLARIFICATION QUESTIONS
List important questions the BA should ask.

13. ACTION ITEMS
Provide:
Action Item | Responsible Person | Priority | Status.

Use To Be Confirmed for unknown information.

14. BUSINESS KPIs
Suggest measurable success indicators.
Do not invent target values.

15. BUSINESS RECOMMENDATIONS
Provide practical improvement suggestions.

Prepare a structured professional report.
"""

            try:
                with st.spinner(
                    "Gemini AI is analyzing the requirements..."
                ):

                    result = generate_ai(prompt)

                    st.session_state["brd_result"] = result

                    save_business_requirement(
                        customer_name.strip(),
                        result,
                        "New"
                    )

                st.success(
                    "Analysis completed and saved to database!"
                )

            except Exception as e:
                st.error(f"AI analysis failed: {e}")
    if "brd_result" in st.session_state:

        st.divider()

        st.markdown("### 📋 Generated Business Analysis")

        st.markdown(st.session_state["brd_result"])

        st.download_button(
            "📄 Download Business Requirements Document",
            data=create_brd(st.session_state["brd_result"]),
            file_name="BizFlow_BRD.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


# =================================================
# 5. AI ASSISTANT
# =================================================

elif page == "🤖 AI Assistant":

    st.markdown("## 🤖 BizFlow AI Assistant")

    st.write(
        "Ask questions about your business data or "
        "Business Analyst concepts."
    )

    question = st.text_area(
        "Ask your question",
        height=130,
        placeholder="Example: Explain the quality of my customer dataset..."
    )

    if st.button("Generate AI Answer", use_container_width=True):

        if not question.strip():
            st.warning("Please enter a question.")

        else:

            if df is not None:

                dataset_summary = f"""
Dataset contains {total_records} records.
Number of columns: {total_columns}.
Missing values: {missing_values}.
Exact duplicate rows: {duplicate_rows}.
Possible duplicate pairs: {len(duplicate_matches)}.
Data completeness: {completeness}%.

Column names: {", ".join(map(str, df.columns))}
Missing values by column:
{df.isna().sum().to_dict()}
"""

            else:
                dataset_summary = "No customer dataset has been uploaded."

            prompt = f"""
You are BizFlow AI, a helpful Business Analysis Assistant.

Answer the user's question in simple professional English.
Use the dataset summary when relevant.
Do not invent dataset statistics.
Do not request or expose personal customer information.

DATASET SUMMARY:
{dataset_summary}

USER QUESTION:
{question}
"""

            try:

                with st.spinner("Gemini is preparing your answer..."):
                    answer = generate_ai(prompt)

                st.markdown("### 💡 AI Response")
                st.markdown(answer)

            except Exception as e:
                st.error(f"AI assistant error: {e}")


elif page == "🗄️ Requirements Database":

    st.markdown("## 🗄️ Requirements Management")
    st.write(
        "Search, edit and manage business requirements "
        "throughout the project lifecycle."
    )

    st.divider()

    # SEARCH REQUIREMENTS
    st.markdown("### 🔍 Search Requirements")

    search_term = st.text_input(
        "Search by Customer / Project Name",
        placeholder="Enter project name..."
    )

    records = get_business_requirements(search_term)

    if records:

        st.markdown("### 📋 Saved Requirements")

        requirement_df = pd.DataFrame(
            records,
            columns=[
                "ID",
                "Customer",
                "Requirement",
                "Status",
                "Priority"
            ]
        )

        st.dataframe(
            requirement_df[
                ["ID", "Customer", "Status", "Priority"]
            ],
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        # SELECT REQUIREMENT
        st.markdown("### ✏️ Edit Requirement")

        record_options = {
            record[0]: f"#{record[0]} - {record[1]}"
            for record in records
        }

        selected_id = st.selectbox(
            "Select a requirement",
            options=list(record_options.keys()),
            format_func=lambda x: record_options[x]
        )

        selected_record = next(
            (r for r in records if r[0] == selected_id),
            None
        )

        if selected_record:

            with st.form("edit_requirement_form"):

                customer = st.text_input(
                    "Customer / Project Name",
                    value=selected_record[1] or ""
                )

                requirement = st.text_area(
                    "Business Requirement",
                    value=selected_record[2] or "",
                    height=200
                )

                status = st.selectbox(
                    "Requirement Status",
                    ["New", "In Progress", "Completed"],
                    index=(
                        ["New", "In Progress", "Completed"].index(
                            selected_record[3]
                        )
                        if selected_record[3] in
                        ["New", "In Progress", "Completed"]
                        else 0
                    )
                )

                priorities = [
                    "Must Have",
                    "Should Have",
                    "Could Have",
                    "Won't Have"
                ]

                priority = st.selectbox(
                    "Requirement Priority",
                    priorities,
                    index=(
                        priorities.index(selected_record[4])
                        if selected_record[4] in priorities
                        else 0
                    )
                )

                save_changes = st.form_submit_button(
                    "💾 Update Requirement",
                    use_container_width=True
                )

                if save_changes:

                    if customer.strip() and requirement.strip():

                        update_business_requirement(
                            selected_id,
                            customer.strip(),
                            requirement.strip(),
                            status,
                            priority
                        )

                        st.success(
                            "Requirement updated successfully!"
                        )

                        st.rerun()

                    else:
                        st.warning(
                            "Customer name and requirement cannot be empty."
                        )

    else:

        st.info("No matching requirements found.")

    st.divider()

    st.markdown("## 🔗 Requirement Traceability Matrix")

    projects = get_business_requirements()

    if projects:

        project_options = {
            record[0]: record[1]
            for record in projects
        }

        selected_project = st.selectbox(
            "Select Project for Traceability",
            options=list(project_options.keys()),
            format_func=lambda x: project_options[x],
            key="rtm_project"
        )

        st.markdown("### ➕ Add Individual Requirement")

        with st.form("rtm_form"):

            req_id = st.text_input(
                "Requirement ID",
                placeholder="Example: FR01"
            )

            req_type = st.selectbox(
                "Requirement Type",
                ["Functional", "Non-Functional", "Business"]
            )

            description = st.text_area(
                "Requirement Description"
            )

            priority = st.selectbox(
                "MoSCoW Priority",
                [
                    "Must Have",
                    "Should Have",
                    "Could Have",
                    "Won't Have"
                ]
            )

            acceptance = st.text_area(
                "Acceptance Criteria",
                placeholder="How will we confirm this requirement is satisfied?"
            )

            test_case = st.text_input(
                "Test Case ID",
                placeholder="Example: TC01"
            )

            test_status = st.selectbox(
                "Test Status",
                ["Not Tested", "Passed", "Failed"]
            )

            submit_rtm = st.form_submit_button(
                "Save Requirement",
                use_container_width=True
            )

            if submit_rtm:

                if req_id.strip() and description.strip():

                    save_traceability_item(
                        selected_project,
                        req_id.strip(),
                        req_type,
                        description.strip(),
                        priority,
                        acceptance.strip(),
                        test_case.strip(),
                        test_status
                    )

                    st.success("Traceability record saved!")
                    st.rerun()

                else:
                    st.warning(
                        "Requirement ID and description are required."
                    )

        st.divider()

        st.markdown("### 📊 Saved Traceability Records")

        rtm_records = get_traceability_items(selected_project)

        if rtm_records:

            rtm_df = pd.DataFrame(
                rtm_records,
                columns=[
                    "Requirement ID",
                    "Type",
                    "Description",
                    "Priority",
                    "Acceptance Criteria",
                    "Test Case",
                    "Test Status"
                ]
            )

            st.dataframe(
                rtm_df,
                use_container_width=True,
                hide_index=True
            )

            csv_data = rtm_df.to_csv(index=False).encode("utf-8")

            st.download_button(
                "⬇️ Download Traceability Matrix",
                data=csv_data,
                file_name="requirement_traceability_matrix.csv",
                mime="text/csv"
            )

        else:
            st.info("No individual requirements added yet.")

    else:
        st.info("Create a project requirement first.")



# ---------------- FOOTER ----------------

st.markdown("""
<div class="footer">
    BIZFLOW AI | SMARTER DATA. BETTER BUSINESS DECISIONS.
    <br>
    Built with Python, Streamlit and Gemini AI.
</div>
""", unsafe_allow_html=True)
