from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException, Request, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import Optional, List
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.drawing.image import Image as XLImage
import os, uuid, httpx, json, re, io, zipfile, base64, csv, time, logging, asyncio
from datetime import datetime, timedelta
from pathlib import Path
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# ══════════════════════════════════════════════════════
# SETUP & CONFIG
# ══════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("xlforge.log", mode="a")
    ]
)
logger = logging.getLogger("xlforge")

app = FastAPI(title="XLforge API", version="2.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage directories
STORAGE_DIR = Path("storage")
INPUT_DIR = STORAGE_DIR / "input"
OUTPUT_DIR = STORAGE_DIR / "output"
TEMP_DIR = STORAGE_DIR / "temp"
for d in [INPUT_DIR, OUTPUT_DIR, TEMP_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# In-memory stores
conversation_memory: dict = {}   # session_id → list of messages
rate_limit_store: dict = {}      # ip → [timestamps]
jobs: dict = {}                  # job_id → status dict

MAX_FILE_SIZE = 20 * 1024 * 1024   # 20MB
RATE_LIMIT = 15                     # requests per minute
FILE_EXPIRY_HOURS = 24


# ══════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════

def get_db():
    conn = sqlite3.connect("xlforge.db", check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS jobs (
            job_id TEXT PRIMARY KEY,
            status TEXT DEFAULT 'pending',
            prompt TEXT,
            session_id TEXT,
            input_file TEXT,
            output_file TEXT,
            error TEXT,
            created_at TEXT,
            completed_at TEXT,
            processing_ms INTEGER
        );
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            category TEXT,
            prompt TEXT,
            icon TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS usage_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            session_id TEXT,
            ip TEXT,
            file_size INTEGER,
            ai_time_ms INTEGER,
            excel_time_ms INTEGER,
            success INTEGER,
            created_at TEXT
        );
    """)
    # Seed default templates
    existing = conn.execute("SELECT COUNT(*) FROM templates").fetchone()[0]
    if existing == 0:
        templates = [
            ("Invoice", "finance", "Create a professional invoice template with company details, line items (description, qty, unit price, total), subtotal, tax (18%), grand total, payment terms, bank details", "🧾"),
            ("Monthly Budget", "finance", "Create a personal monthly budget tracker with income sources, fixed expenses, variable expenses, savings goal, actual vs budget comparison, and summary chart", "💰"),
            ("Inventory Tracker", "business", "Create an inventory management sheet with product ID, name, category, quantity in stock, reorder level, unit cost, total value, supplier, last restocked date, and stock status alerts", "📦"),
            ("Employee Attendance", "HR", "Create an employee attendance sheet for a month with employee ID, name, department, 31 days columns (P/A/L/H), total present, total absent, total leave, attendance percentage", "👥"),
            ("Sales Report", "sales", "Create a monthly sales report with salesperson, region, product, units sold, unit price, revenue, target, achievement percentage, rank, and bar chart", "📈"),
            ("Project Timeline", "management", "Create a project timeline/gantt-style sheet with task name, owner, start date, end date, duration days, status (Not Started/In Progress/Complete), priority, dependencies, completion %", "📅"),
            ("Student Gradebook", "education", "Create a student gradebook with student name, roll number, 6 subjects scores, total, percentage, grade (A/B/C/D/F), rank, pass/fail status and class average row", "🎓"),
            ("Expense Report", "finance", "Create a business expense report with date, category, description, amount, receipt number, payment method, reimbursable yes/no, approval status, and monthly totals", "💳"),
            ("KPI Dashboard", "management", "Create a KPI dashboard with metrics (Revenue, Customers, Conversion Rate, Avg Order Value, Churn Rate, NPS), current value, previous period, target, variance, trend (↑↓), and status (On Track/At Risk/Behind)", "📊"),
            ("Quotation", "sales", "Create a business quotation template with client details, quote number, validity date, itemized list (item, specs, qty, unit price, discount, net price), terms & conditions, total, and signature section", "📋"),
            ("BOQ (Bill of Quantities)", "construction", "Create a BOQ sheet with item number, description of work, unit, quantity, rate, amount, GST %, GST amount, total amount, section subtotals and grand total", "🏗️"),
            ("Payroll Sheet", "HR", "Create a monthly payroll sheet with employee ID, name, designation, basic salary, HRA, DA, other allowances, gross salary, PF deduction, ESI, tax, total deductions, net salary", "💵"),
        ]
        conn.executemany(
            "INSERT INTO templates (name, category, prompt, icon, created_at) VALUES (?,?,?,?,?)",
            [(t[0], t[1], t[2], t[3], datetime.utcnow().isoformat()) for t in templates]
        )
    conn.commit()
    conn.close()

init_db()


# ══════════════════════════════════════════════════════
# RATE LIMITER
# ══════════════════════════════════════════════════════

def check_rate_limit(ip: str) -> bool:
    now = time.time()
    window = 60
    if ip not in rate_limit_store:
        rate_limit_store[ip] = []
    rate_limit_store[ip] = [t for t in rate_limit_store[ip] if now - t < window]
    if len(rate_limit_store[ip]) >= RATE_LIMIT:
        return False
    rate_limit_store[ip].append(now)
    return True


# ══════════════════════════════════════════════════════
# FILE READER
# ══════════════════════════════════════════════════════

async def read_any_file(file: UploadFile) -> tuple:
    raw = await file.read()

    if len(raw) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail=f"File too large. Max size is {MAX_FILE_SIZE // 1024 // 1024}MB")

    filename = (file.filename or "").lower()

    if filename.endswith(('.xlsx', '.xls')):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
            lines = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                lines.append(f"[Sheet: {sheet_name}]")
                for row in ws.iter_rows(values_only=True):
                    if any(c is not None for c in row):
                        lines.append(" | ".join(str(c) if c is not None else "" for c in row))
            return "\n".join(lines), "excel", None, raw
        except Exception as e:
            return f"Excel read error: {e}", "excel", None, raw

    if filename.endswith('.csv'):
        try:
            return raw.decode("utf-8", errors="ignore"), "csv", None, raw
        except:
            return "", "csv", None, raw

    if filename.endswith('.docx'):
        try:
            z = zipfile.ZipFile(io.BytesIO(raw))
            xml = z.read("word/document.xml").decode("utf-8")
            text = re.sub(r'<[^>]+>', ' ', xml)
            text = re.sub(r'\s+', ' ', text).strip()
            return text[:6000], "word", None, raw
        except:
            return "", "word", None, raw

    if filename.endswith('.pdf'):
        try:
            text = raw.decode("latin-1", errors="ignore")
            strings = re.findall(r'[A-Za-z0-9 \+\-\=\.\,\:\;\!\?\%\$\#\@\/\(\)]{4,}', text)
            extracted = " ".join(strings)[:6000]
            return extracted, "pdf", None, raw
        except:
            return "", "pdf", None, raw

    if filename.endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp')):
        b64 = base64.b64encode(raw).decode()
        ext = filename.split('.')[-1]
        mime = f"image/{'jpeg' if ext in ['jpg','jpeg'] else ext}"
        return "", "image", {"b64": b64, "mime": mime}, raw

    try:
        return raw.decode("utf-8", errors="ignore")[:6000], "text", None, raw
    except:
        return "", "unknown", None, raw


# ══════════════════════════════════════════════════════
# SYSTEM PROMPT
# ══════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are XLforge, the world's most advanced AI Excel expert. You read any file, understand it completely, and produce perfect, professional Excel spreadsheets.

OUTPUT RULE: Return ONLY a valid JSON object. Nothing else. No markdown. No explanation. No text before or after the JSON.

JSON FORMAT:
{
  "sheets": [
    {
      "name": "Sheet Name",
      "headers": ["Column A", "Column B", "Column C"],
      "rows": [
        ["real value", 100, "text"]
      ],
      "formulas": [
        {"cell": "D2", "formula": "=SUM(B2:C2)", "label": "Total"}
      ],
      "summary_rows": [
        {"label": "TOTAL", "col": 1, "formula": "=SUM(B2:B100)"}
      ],
      "conditional_formatting": [
        {"range": "B2:B20", "type": "colorscale"},
        {"range": "C2:C20", "type": "databar"},
        {"range": "D2:D20", "type": "highlight_high", "threshold": 500}
      ],
      "chart": {
        "type": "bar",
        "title": "Chart Title",
        "data_cols": [1],
        "category_col": 0
      },
      "protection": null
    }
  ],
  "metadata": {
    "title": "Report Title",
    "description": "What this spreadsheet does",
    "author": "XLforge AI"
  }
}

══════════════════════════════
ABSOLUTE RULES — NEVER BREAK:
══════════════════════════════

RULE 1 — UNDERSTAND THE FILE COMPLETELY:
- Read every single row without skipping
- Understand exactly what kind of data it contains
- Complete ALL missing values intelligently
- Never replace file data with fake generated data

RULE 2 — SOLVE EVERYTHING:
Math problems → compute exact answers
Questions with blanks → fill correct answers
Missing totals → calculate them
Empty grade columns → compute grades A/B/C/D/F
Blank status columns → determine correct status
Answer/Result columns → COMPUTED VALUES not formulas

RULE 3 — COPY ALL ROWS:
If file has 50 rows → output must have 50 rows
If file has 100 rows → output must have 100 rows
NEVER reduce rows

RULE 4 — AUTO-UNDERSTAND MODE:
Math problems → solve all, add Answer column
Student marks → add Total, Average, Grade, Pass/Fail, Rank
Financial data → add totals, summaries, trends, chart
Inventory → add Stock Status, Reorder Alert, Value
Employee data → add summaries, department totals
Survey data → add analysis, counts, percentages
Always add value beyond original file

RULE 5 — REAL NUMBERS:
Prices, salaries, scores → integers or floats, NEVER strings
BAD: ["John", "50000"]
GOOD: ["John", 50000]

RULE 6 — SAFE FORMULAS ONLY:
Always wrap VLOOKUP in IFERROR:
=IFERROR(VLOOKUP(A2,Sheet2!$A:$B,2,FALSE),"Not Found")

Valid formulas:
=SUM(B2:B10), =AVERAGE(B2:B10), =MAX(B2:B10), =MIN(B2:B10)
=IF(B2>90,"A",IF(B2>80,"B",IF(B2>70,"C",IF(B2>60,"D","F"))))
=COUNTIF(B2:B10,">100"), =SUMIF(A2:A10,"North",B2:B10)
=RANK(B2,$B$2:$B$100,0), =TODAY(), =TEXT(A2,"DD-MMM-YYYY")
=IFERROR(formula,"fallback")

RULE 7 — MULTIPLE SHEETS WHEN HELPFUL:
For complex data, create:
Sheet 1: Raw data / main table
Sheet 2: Summary / dashboard
Sheet 3: Charts data

RULE 8 — CHARTS:
type options: "bar", "line", "pie", "area"
data_cols: list of 0-indexed number columns
category_col: 0-indexed label column
Always include chart for comparative or trend data

RULE 9 — CONDITIONAL FORMATTING:
"colorscale" → red-yellow-green gradient on numeric ranges
"databar" → blue progress bars
"highlight_high" → green for above threshold

RULE 10 — NO PLACEHOLDERS EVER:
Never: val1, col1, value1, header1, item1, data1, sample, name1
Always use real, meaningful data

RULE 11 — PROFESSIONAL QUALITY:
Every spreadsheet must look like it was made by a professional consultant.
Add summary rows at the bottom.
Add a totals/averages row.
Use meaningful sheet names.
Include metadata."""


# ══════════════════════════════════════════════════════
# VALIDATION LAYER
# ══════════════════════════════════════════════════════

def validate_ai_response(data: dict) -> tuple[bool, str]:
    if not isinstance(data, dict):
        return False, "Response is not a dict"
    if "sheets" not in data:
        return False, "Missing 'sheets' key"
    if not data["sheets"]:
        return False, "Empty sheets list"
    for i, sheet in enumerate(data["sheets"]):
        if not sheet.get("headers"):
            return False, f"Sheet {i}: missing headers"
        if not sheet.get("rows"):
            return False, f"Sheet {i}: missing rows"
        if not isinstance(sheet["headers"], list):
            return False, f"Sheet {i}: headers must be a list"
        if not isinstance(sheet["rows"], list):
            return False, f"Sheet {i}: rows must be a list"
        header_len = len(sheet["headers"])
        for j, row in enumerate(sheet["rows"]):
            if not isinstance(row, list):
                return False, f"Sheet {i}, Row {j}: must be a list"
            # Pad short rows
            while len(row) < header_len:
                row.append("")
    return True, "OK"


def coerce_numeric(val):
    """Convert string numbers to int/float so Excel SUM/formulas work correctly."""
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        v = val.strip()
        try:
            if '.' in v:
                return float(v)
            return int(v)
        except (ValueError, TypeError):
            pass
    return val

def sanitize_ai_response(data: dict) -> dict:
    for sheet in data.get("sheets", []):
        header_len = len(sheet.get("headers", []))
        cleaned_rows = []
        for row in sheet.get("rows", []):
            if isinstance(row, list):
                row = row[:header_len]
                while len(row) < header_len:
                    row.append("")
                # Coerce numeric strings → real numbers so SUM/formulas work
                row = [coerce_numeric(v) for v in row]
                cleaned_rows.append(row)
        sheet["rows"] = cleaned_rows
        if not sheet.get("formulas"):
            sheet["formulas"] = []
        if not sheet.get("conditional_formatting"):
            sheet["conditional_formatting"] = []
        if not sheet.get("summary_rows"):
            sheet["summary_rows"] = []
    return data


# ══════════════════════════════════════════════════════
# AI CALL (GROQ)
# ══════════════════════════════════════════════════════

async def call_groq(
    prompt: str,
    file_content: str = "",
    file_type: str = "",
    image_data: dict = None,
    session_id: str = None,
) -> dict:

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    # Inject conversation history
    if session_id and session_id in conversation_memory:
        history = conversation_memory[session_id][-6:]  # last 3 exchanges
        messages.extend(history)

    # Build user message
    if image_data:
        messages.append({
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{image_data['mime']};base64,{image_data['b64']}"}
                },
                {
                    "type": "text",
                    "text": f"""Analyze this image carefully.
{"Task: " + prompt if prompt.strip() else "Understand this image and create the best Excel spreadsheet from it."}
Read all text, numbers, tables visible in the image.
Solve any math problems, complete any missing data.
Return only JSON."""
                }
            ]
        })
    elif file_content and not prompt.strip():
        messages.append({
            "role": "user",
            "content": f"""I uploaded a {file_type} file. Analyze it and create the perfect professional Excel output.

FILE CONTENT (process ALL rows, do not skip any):
{file_content[:8000]}

Instructions:
- Understand exactly what this data is about
- Complete every missing value intelligently
- Solve every math problem if present
- Add professional formulas and charts
- Copy ALL rows without skipping any
- Add summary/totals rows at the bottom
- Return only JSON"""
        })
    elif file_content and prompt.strip():
        messages.append({
            "role": "user",
            "content": f"""Task: {prompt}

FILE TYPE: {file_type}
FILE CONTENT (use ALL rows, do not skip any):
{file_content[:8000]}

IMPORTANT:
- Use every single row from the file
- Complete any missing values or answers
- Do the task described above using this real data
- Add professional formatting and formulas
- Return only JSON"""
        })
    else:
        messages.append({
            "role": "user",
            "content": f"""Create a professional Excel spreadsheet: "{prompt}"
- Real, meaningful data — minimum 15 rows
- Multiple sheets if it makes sense
- Professional formulas throughout
- Add chart if data is comparative or trending
- Add conditional formatting
- Summary row at bottom
- No placeholder values whatsoever
- Return only JSON"""
        })

    last_error = None
    for attempt in range(4):
        try:
            temperature = [0.1, 0.2, 0.35, 0.5][attempt]
            model = "meta-llama/llama-4-scout-17b-16e-instruct" if image_data else "llama-3.3-70b-versatile"

            async with httpx.AsyncClient(timeout=60) as client:
                response = await client.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {os.getenv('GROQ_API_KEY')}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": model,
                        "max_tokens": 8000,
                        "temperature": temperature,
                        "messages": messages
                    }
                )

            if response.status_code != 200:
                raise ValueError(f"Groq HTTP {response.status_code}: {response.text[:200]}")

            result = response.json()
            if "error" in result:
                raise ValueError(f"Groq error: {result['error']}")

            text = result["choices"][0]["message"]["content"].strip()
            text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
            text = re.sub(r"\n?```$", "", text)
            text = text.strip()

            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                text = json_match.group(0)

            data = json.loads(text)

            valid, reason = validate_ai_response(data)
            if not valid:
                raise ValueError(f"Validation failed: {reason}")

            data = sanitize_ai_response(data)

            # Save to conversation memory
            if session_id:
                user_msg = {"role": "user", "content": prompt or f"[{file_type} file uploaded]"}
                assistant_msg = {"role": "assistant", "content": f"[Generated Excel with {len(data['sheets'])} sheet(s)]"}
                if session_id not in conversation_memory:
                    conversation_memory[session_id] = []
                conversation_memory[session_id].extend([user_msg, assistant_msg])
                conversation_memory[session_id] = conversation_memory[session_id][-20:]  # keep last 10 exchanges

            return data

        except json.JSONDecodeError as e:
            last_error = f"JSON parse error: {e}"
            logger.warning(f"Attempt {attempt+1} JSON error: {e}")
            continue
        except ValueError as e:
            last_error = str(e)
            logger.warning(f"Attempt {attempt+1} value error: {e}")
            continue
        except httpx.TimeoutException:
            last_error = "AI service timed out"
            logger.warning(f"Attempt {attempt+1} timeout")
            continue
        except Exception as e:
            last_error = str(e)
            logger.error(f"Attempt {attempt+1} unexpected error: {e}")
            continue

    raise ValueError(f"AI generation failed after 4 attempts. Last error: {last_error}")


# ══════════════════════════════════════════════════════
# EXCEL BUILDER
# ══════════════════════════════════════════════════════

def build_excel(data: dict, output_path: str, password: str = None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    BLUE = "2563EB"
    LIGHT_BLUE = "EFF6FF"
    WHITE = "FFFFFF"
    DARK = "1E293B"
    SUMMARY_BG = "1E3A5F"

    def thin_border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def thick_border():
        s = Side(style="medium", color="2563EB")
        return Border(left=s, right=s, top=s, bottom=s)

    def style_header(cell):
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(bold=True, color=WHITE, size=11, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()

    def style_data(cell, row_idx, header=""):
        cell.border = thin_border()
        cell.alignment = Alignment(vertical="center")
        if row_idx % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        if isinstance(cell.value, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
            h = header.lower()
            if any(k in h for k in ["salary","revenue","price","cost","amount","budget",
                                      "sales","income","expense","pay","total","value","rate"]):
                cell.number_format = '#,##0.00'
            elif any(k in h for k in ["percent","%","rate","growth","margin","achievement"]):
                cell.number_format = '0.00'

    def style_summary(cell):
        cell.fill = PatternFill("solid", fgColor=SUMMARY_BG)
        cell.font = Font(bold=True, color=WHITE, size=11, name="Calibri")
        cell.border = thick_border()
        cell.alignment = Alignment(horizontal="right", vertical="center")

    for sheet_def in data["sheets"]:
        name = str(sheet_def.get("name", "Sheet"))[:31]
        ws = wb.create_sheet(title=name)

        headers  = sheet_def.get("headers", [])
        rows     = sheet_def.get("rows", [])
        formulas = sheet_def.get("formulas", [])
        cf_rules = sheet_def.get("conditional_formatting", [])
        chart_def = sheet_def.get("chart", None)
        summary_rows = sheet_def.get("summary_rows", [])

        # ── Tab color
        ws.sheet_properties.tabColor = BLUE

        # ── Write headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=str(header))
            style_header(cell)
            ws.column_dimensions[get_column_letter(col)].width = max(len(str(header)) + 6, 18)
        ws.row_dimensions[1].height = 32

        # ── Write data rows
        for row_idx, row in enumerate(rows, 2):
            for col_idx, val in enumerate(row, 1):
                header = headers[col_idx-1] if col_idx <= len(headers) else ""
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                style_data(cell, row_idx, header)
            ws.row_dimensions[row_idx].height = 20

        # ── Summary rows at bottom
        next_row = len(rows) + 2
        for sr in summary_rows:
            label = sr.get("label", "TOTAL")
            # AI sends 0-indexed col; convert to 1-indexed for openpyxl
            col_idx = int(sr.get("col", 1)) + 1
            formula = sr.get("formula", "")
            # Fill label across all columns up to value col with merged-style look
            for c in range(1, col_idx):
                lc = ws.cell(row=next_row, column=c, value=label if c == 1 else None)
                style_summary(lc)
            # Value cell
            if formula:
                val_cell = ws.cell(row=next_row, column=col_idx, value=formula)
            else:
                val_cell = ws.cell(row=next_row, column=col_idx, value=0)
            style_summary(val_cell)
            # Style remaining cells in row
            for c in range(col_idx + 1, len(headers) + 1):
                style_summary(ws.cell(row=next_row, column=c))
            next_row += 1

        # ── Inline formulas
        for f in formulas:
            addr = f.get("cell", "")
            formula = f.get("formula", "")
            label = f.get("label", "")
            if addr and formula:
                try:
                    ws[addr] = formula
                    ws[addr].font = Font(bold=True, color=DARK)
                    ws[addr].border = thick_border()
                    m = re.match(r"([A-Z]+)(\d+)", addr)
                    if m and label:
                        ci = column_index_from_string(m.group(1))
                        ri = int(m.group(2))
                        if ci > 1:
                            lc = ws.cell(row=ri, column=ci-1, value=label)
                            lc.font = Font(bold=True)
                            lc.border = thin_border()
                except:
                    pass

        # ── Freeze panes & auto-filter
        ws.freeze_panes = "A2"
        if headers:
            last_data_row = len(rows) + 1
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_data_row}"

        # ── Print settings
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.sheet_view.showGridLines = True

        # ── Conditional formatting
        for cf in cf_rules:
            r = cf.get("range", "")
            t = cf.get("type", "")
            if not r:
                continue
            try:
                if t == "colorscale":
                    ws.conditional_formatting.add(r, ColorScaleRule(
                        start_type="min",  start_color="F87171",
                        mid_type="percentile", mid_value=50, mid_color="FCD34D",
                        end_type="max",    end_color="4ADE80"
                    ))
                elif t == "databar":
                    ws.conditional_formatting.add(r, DataBarRule(
                        start_type="min", start_value=0,
                        end_type="max",   end_value=100,
                        color="2563EB"
                    ))
                elif t == "highlight_high":
                    threshold = cf.get("threshold", 0)
                    ds = DifferentialStyle(
                        fill=PatternFill(bgColor="4ADE80"),
                        font=Font(color="14532D", bold=True)
                    )
                    ws.conditional_formatting.add(r, CellIsRule(
                        operator="greaterThan",
                        formula=[str(threshold)],
                        stopIfTrue=True,
                        dxf=ds
                    ))
                elif t == "highlight_low":
                    threshold = cf.get("threshold", 0)
                    ds = DifferentialStyle(
                        fill=PatternFill(bgColor="FCA5A5"),
                        font=Font(color="7F1D1D", bold=True)
                    )
                    ws.conditional_formatting.add(r, CellIsRule(
                        operator="lessThan",
                        formula=[str(threshold)],
                        stopIfTrue=True,
                        dxf=ds
                    ))
            except Exception as e:
                logger.warning(f"CF error: {e}")

        # ── Chart
        if chart_def:
            try:
                ctype  = chart_def.get("type", "bar")
                dcols  = chart_def.get("data_cols", [1])
                catcol = chart_def.get("category_col", 0)
                ctitle = chart_def.get("title", "Chart")
                nrows  = len(rows)

                chart_map = {
                    "pie":  PieChart(),
                    "line": LineChart(),
                    "area": AreaChart(),
                    "bar":  BarChart()
                }
                chart = chart_map.get(ctype, BarChart())
                chart.title  = ctitle
                chart.style  = 10
                chart.height = 15
                chart.width  = 28

                if ctype == "bar":
                    chart.type = "col"

                if ctype == "pie":
                    data_ref = Reference(ws, min_col=dcols[0]+1, min_row=1, max_row=nrows+1)
                    chart.add_data(data_ref, titles_from_data=True)
                else:
                    for dc in dcols:
                        data_ref = Reference(ws, min_col=dc+1, min_row=1, max_row=nrows+1)
                        chart.add_data(data_ref, titles_from_data=True)

                cats = Reference(ws, min_col=catcol+1, min_row=2, max_row=nrows+1)
                chart.set_categories(cats)

                anchor_row = max(nrows + 4, next_row + 2)
                ws.add_chart(chart, f"A{anchor_row}")
            except Exception as e:
                logger.warning(f"Chart error: {e}")

    wb.save(output_path)
    return output_path


# ══════════════════════════════════════════════════════
# CSV CONVERTER
# ══════════════════════════════════════════════════════

def excel_to_csv(excel_path: str) -> str:
    csv_path = excel_path.replace(".xlsx", ".csv")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            writer.writerow([c if c is not None else "" for c in row])
    return csv_path


# ══════════════════════════════════════════════════════
# BACKGROUND JOB PROCESSOR
# ══════════════════════════════════════════════════════

async def process_job(
    job_id: str,
    prompt: str,
    file_content: str,
    file_type: str,
    image_data: dict,
    session_id: str,
    output_path: str,
    password: str = None,
    custom_filename: str = None,
):
    start = time.time()
    conn = get_db()
    try:
        jobs[job_id]["status"] = "processing"
        conn.execute("UPDATE jobs SET status='processing' WHERE job_id=?", (job_id,))
        conn.commit()

        ai_start = time.time()
        data = await call_groq(prompt, file_content, file_type, image_data, session_id)
        ai_ms = int((time.time() - ai_start) * 1000)

        excel_start = time.time()
        build_excel(data, output_path, password=password)
        excel_ms = int((time.time() - excel_start) * 1000)

        total_ms = int((time.time() - start) * 1000)

        jobs[job_id].update({
            "status": "done",
            "output_path": output_path,
            "custom_filename": custom_filename or "xlforge_output.xlsx",
            "sheet_count": len(data.get("sheets", [])),
            "metadata": data.get("metadata", {}),
            "processing_ms": total_ms
        })

        conn.execute("""
            UPDATE jobs SET status='done', output_file=?, completed_at=?, processing_ms=?
            WHERE job_id=?
        """, (output_path, datetime.utcnow().isoformat(), total_ms, job_id))
        conn.execute("""
            INSERT INTO usage_log (job_id, session_id, ai_time_ms, excel_time_ms, success, created_at)
            VALUES (?,?,?,?,1,?)
        """, (job_id, session_id, ai_ms, excel_ms, datetime.utcnow().isoformat()))
        conn.commit()

        logger.info(f"Job {job_id} done in {total_ms}ms (AI:{ai_ms}ms, Excel:{excel_ms}ms)")

    except Exception as e:
        jobs[job_id].update({"status": "failed", "error": str(e)})
        conn.execute("UPDATE jobs SET status='failed', error=? WHERE job_id=?", (str(e), job_id))
        conn.execute("""
            INSERT INTO usage_log (job_id, session_id, success, created_at)
            VALUES (?,?,0,?)
        """, (job_id, session_id, datetime.utcnow().isoformat()))
        conn.commit()
        logger.error(f"Job {job_id} failed: {e}")
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# FILE CLEANUP (run periodically)
# ══════════════════════════════════════════════════════

def cleanup_old_files():
    cutoff = time.time() - (FILE_EXPIRY_HOURS * 3600)
    for folder in [INPUT_DIR, OUTPUT_DIR, TEMP_DIR]:
        for f in folder.glob("*"):
            if f.stat().st_mtime < cutoff:
                try:
                    f.unlink()
                    logger.info(f"Cleaned up: {f}")
                except:
                    pass


# ══════════════════════════════════════════════════════
# ENDPOINTS
# ══════════════════════════════════════════════════════

@app.get("/")
def root():
    return {
        "status": "XLforge v2.0 running!",
        "features": ["async jobs", "conversation memory", "templates", "CSV export",
                      "multiple files", "rate limiting", "history", "validation"]
    }

@app.get("/health")
def health():
    return {
        "status": "healthy",
        "active_jobs": len([j for j in jobs.values() if j["status"] == "processing"]),
        "sessions": len(conversation_memory),
        "timestamp": datetime.utcnow().isoformat()
    }


# ── MAIN GENERATE ENDPOINT (async job)
@app.post("/generate")
async def generate_excel(
    background_tasks: BackgroundTasks,
    request: Request,
    prompt: str = Form(default=""),
    session_id: str = Form(default=""),
    custom_filename: str = Form(default=""),
    password: str = Form(default=""),
    files: List[UploadFile] = File(default=[]),
    file: UploadFile = File(default=None),
):
    ip = request.client.host
    if not check_rate_limit(ip):
        raise HTTPException(status_code=429, detail="Rate limit exceeded. Max 15 requests/minute.")

    if not session_id:
        session_id = str(uuid.uuid4())

    job_id = str(uuid.uuid4())
    output_path = str(OUTPUT_DIR / f"{job_id}.xlsx")

    # Handle single or multiple files
    all_files = [f for f in (files or []) if f and f.filename]
    if file and file.filename:
        all_files.insert(0, file)

    file_content = ""
    file_type = ""
    image_data = None

    if all_files:
        # Merge content from all files
        contents = []
        for uf in all_files[:5]:  # max 5 files
            fc, ft, img, raw = await read_any_file(uf)
            if img:
                image_data = img  # use last image
            elif fc:
                contents.append(f"[File: {uf.filename}]\n{fc}")
            file_type = ft

            # Save input file
            input_path = INPUT_DIR / f"{job_id}_{uf.filename}"
            with open(input_path, "wb") as f_out:
                f_out.write(raw)

        file_content = "\n\n".join(contents)

    # Create job record
    jobs[job_id] = {
        "status": "pending",
        "job_id": job_id,
        "session_id": session_id,
        "created_at": datetime.utcnow().isoformat()
    }

    conn = get_db()
    conn.execute("""
        INSERT INTO jobs (job_id, status, prompt, session_id, created_at)
        VALUES (?,?,?,?,?)
    """, (job_id, "pending", prompt, session_id, datetime.utcnow().isoformat()))
    conn.commit()
    conn.close()

    fname = custom_filename.strip() or "xlforge_output"
    if not fname.endswith(".xlsx"):
        fname += ".xlsx"

    background_tasks.add_task(
        process_job,
        job_id=job_id,
        prompt=prompt,
        file_content=file_content,
        file_type=file_type,
        image_data=image_data,
        session_id=session_id,
        output_path=output_path,
        password=password or None,
        custom_filename=fname
    )

    return {
        "job_id": job_id,
        "session_id": session_id,
        "status": "processing",
        "message": "Your Excel is being generated. Poll /status/{job_id} to check progress."
    }


# ── STATUS CHECK
@app.get("/status/{job_id}")
def job_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        conn = get_db()
        row = conn.execute("SELECT * FROM jobs WHERE job_id=?", (job_id,)).fetchone()
        conn.close()
        if not row:
            raise HTTPException(status_code=404, detail="Job not found")
        return dict(row)
    return job


# ── DOWNLOAD EXCEL
@app.get("/download/{job_id}")
def download_excel(job_id: str):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        raise HTTPException(status_code=404, detail="File not ready or job not found")
    output_path = job.get("output_path", "")
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="File not found on disk")
    filename = job.get("custom_filename", "xlforge_output.xlsx")
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
        headers={"Access-Control-Allow-Origin": "*"}
    )


# ── DOWNLOAD AS CSV
@app.get("/download-csv/{job_id}")
def download_csv(job_id: str):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        raise HTTPException(status_code=404, detail="File not ready")
    output_path = job.get("output_path", "")
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="File not found")
    csv_path = excel_to_csv(output_path)
    filename = job.get("custom_filename", "xlforge_output.xlsx").replace(".xlsx", ".csv")
    return FileResponse(
        csv_path,
        media_type="text/csv",
        filename=filename,
        headers={"Access-Control-Allow-Origin": "*"}
    )


# ── SYNC GENERATE (for simple/quick requests, backwards compatible)
@app.post("/generate-sync")
async def generate_sync(
    request: Request,
    prompt: str = Form(default=""),
    session_id: str = Form(default=""),
    custom_filename: str = Form(default=""),
    file: UploadFile = File(None)
):
    ip = request.client.host
    if not check_rate_limit(ip):
        raise HTTPException(status_code=429, detail="Rate limit exceeded.")

    if not session_id:
        session_id = str(uuid.uuid4())

    job_id = str(uuid.uuid4())
    output_path = str(OUTPUT_DIR / f"{job_id}.xlsx")

    file_content, file_type, image_data = "", "", None
    if file and file.filename:
        file_content, file_type, image_data, _ = await read_any_file(file)

    try:
        data = await call_groq(prompt, file_content, file_type, image_data, session_id)
        build_excel(data, output_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    fname = custom_filename.strip() or "xlforge_output"
    if not fname.endswith(".xlsx"):
        fname += ".xlsx"

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=fname,
        headers={"Access-Control-Allow-Origin": "*", "X-Session-Id": session_id}
    )


# ── TEMPLATES
@app.get("/templates")
def get_templates(category: str = None):
    conn = get_db()
    if category:
        rows = conn.execute(
            "SELECT * FROM templates WHERE category=? ORDER BY name", (category,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM templates ORDER BY category, name").fetchall()
    conn.close()
    return {"templates": [dict(r) for r in rows]}

@app.get("/templates/categories")
def get_categories():
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT category FROM templates ORDER BY category").fetchall()
    conn.close()
    return {"categories": [r["category"] for r in rows]}

@app.post("/templates/use/{template_id}")
async def use_template(
    template_id: int,
    background_tasks: BackgroundTasks,
    request: Request,
    session_id: str = Form(default=""),
    custom_data: str = Form(default=""),
):
    conn = get_db()
    tmpl = conn.execute("SELECT * FROM templates WHERE id=?", (template_id,)).fetchone()
    conn.close()
    if not tmpl:
        raise HTTPException(status_code=404, detail="Template not found")

    prompt = tmpl["prompt"]
    if custom_data.strip():
        prompt += f"\n\nAdditional data/context: {custom_data}"

    ip = request.client.host
    if not check_rate_limit(ip):
        raise HTTPException(status_code=429, detail="Rate limit exceeded.")

    if not session_id:
        session_id = str(uuid.uuid4())

    job_id = str(uuid.uuid4())
    output_path = str(OUTPUT_DIR / f"{job_id}.xlsx")
    fname = f"{tmpl['name'].lower().replace(' ','_')}.xlsx"

    jobs[job_id] = {"status": "pending", "job_id": job_id, "session_id": session_id,
                    "created_at": datetime.utcnow().isoformat()}

    background_tasks.add_task(
        process_job,
        job_id=job_id, prompt=prompt, file_content="", file_type="",
        image_data=None, session_id=session_id, output_path=output_path,
        custom_filename=fname
    )

    return {"job_id": job_id, "session_id": session_id, "status": "processing",
            "template": tmpl["name"]}


# ── CONVERSATION / SESSION
@app.get("/session/{session_id}/history")
def get_session_history(session_id: str):
    history = conversation_memory.get(session_id, [])
    conn = get_db()
    jobs_list = conn.execute(
        "SELECT job_id, status, prompt, created_at, processing_ms FROM jobs WHERE session_id=? ORDER BY created_at DESC LIMIT 20",
        (session_id,)
    ).fetchall()
    conn.close()
    return {
        "session_id": session_id,
        "message_count": len(history),
        "jobs": [dict(j) for j in jobs_list]
    }

@app.delete("/session/{session_id}")
def clear_session(session_id: str):
    if session_id in conversation_memory:
        del conversation_memory[session_id]
    return {"message": "Session cleared", "session_id": session_id}


# ── JOB HISTORY
@app.get("/history")
def get_history(limit: int = 20):
    conn = get_db()
    rows = conn.execute(
        "SELECT job_id, status, prompt, created_at, processing_ms FROM jobs ORDER BY created_at DESC LIMIT ?",
        (limit,)
    ).fetchall()
    conn.close()
    return {"jobs": [dict(r) for r in rows]}


# ── EMAIL ENDPOINT
@app.post("/email/{job_id}")
async def email_file(job_id: str, to_email: str = Form(...), subject: str = Form(default="Your Excel File from XLforge")):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        raise HTTPException(status_code=404, detail="File not ready")

    output_path = job.get("output_path", "")
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="File not found")

    gmail_user = os.getenv("GMAIL_USER")
    gmail_pass = os.getenv("GMAIL_APP_PASSWORD")
    if not gmail_user or not gmail_pass:
        raise HTTPException(status_code=503, detail="Email not configured. Set GMAIL_USER and GMAIL_APP_PASSWORD env vars.")

    try:
        msg = MIMEMultipart()
        msg["From"] = gmail_user
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText("Please find your Excel file generated by XLforge AI attached.", "plain"))

        with open(output_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        filename = job.get("custom_filename", "xlforge_output.xlsx")
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_user, gmail_pass)
            server.send_message(msg)

        return {"message": f"File sent to {to_email} successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Email failed: {str(e)}")


# ── STATS / MONITORING
@app.get("/stats")
def get_stats():
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) FROM jobs").fetchone()[0]
    done = conn.execute("SELECT COUNT(*) FROM jobs WHERE status='done'").fetchone()[0]
    failed = conn.execute("SELECT COUNT(*) FROM jobs WHERE status='failed'").fetchone()[0]
    avg_time = conn.execute("SELECT AVG(processing_ms) FROM jobs WHERE status='done'").fetchone()[0]
    today = conn.execute(
        "SELECT COUNT(*) FROM jobs WHERE date(created_at)=date('now')"
    ).fetchone()[0]
    conn.close()
    return {
        "total_jobs": total,
        "successful": done,
        "failed": failed,
        "today": today,
        "avg_processing_ms": round(avg_time or 0),
        "success_rate": f"{(done/total*100):.1f}%" if total > 0 else "N/A",
        "active_sessions": len(conversation_memory),
        "active_jobs": len([j for j in jobs.values() if j["status"] == "processing"])
    }


# ── CLEANUP ENDPOINT
@app.post("/admin/cleanup")
def run_cleanup():
    cleanup_old_files()
    return {"message": "Cleanup complete"}


# ── LIST ACTIVE JOBS
@app.get("/jobs")
def list_jobs():
    return {
        "jobs": [
            {k: v for k, v in j.items() if k != "output_path"}
            for j in list(jobs.values())[-50:]
        ]
    }
