from fastapi import FastAPI, UploadFile, File, Form, Optional
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule
from openpyxl.styles.differential import DifferentialStyle
import os, uuid, httpx, json, re, io, zipfile, base64

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "XLforge running!"}


# ══════════════════════════════════════════════════════
# FILE READER — reads ANY file type into text
# ══════════════════════════════════════════════════════

async def read_any_file(file: UploadFile) -> tuple:
    raw = await file.read()
    filename = (file.filename or "").lower()

    # ── Excel
    if filename.endswith(('.xlsx', '.xls')):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
            lines = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                lines.append(f"[Sheet: {sheet_name}]")
                for row in ws.iter_rows(values_only=True):
                    if any(c is not None for c in row):
                        lines.append(" | ".join(
                            str(c) if c is not None else "" for c in row
                        ))
            return "\n".join(lines), "excel", None
        except Exception as e:
            return f"Excel read error: {e}", "excel", None

    # ── CSV
    if filename.endswith('.csv'):
        try:
            return raw.decode("utf-8", errors="ignore"), "csv", None
        except:
            return "", "csv", None

    # ── Word docx
    if filename.endswith('.docx'):
        try:
            z = zipfile.ZipFile(io.BytesIO(raw))
            xml = z.read("word/document.xml").decode("utf-8")
            text = re.sub(r'<[^>]+>', ' ', xml)
            text = re.sub(r'\s+', ' ', text).strip()
            return text[:6000], "word", None
        except:
            return "", "word", None

    # ── PDF — extract text from raw bytes
    if filename.endswith('.pdf'):
        try:
            text = raw.decode("latin-1", errors="ignore")
            # Extract readable strings from PDF
            strings = re.findall(r'[A-Za-z0-9 \+\-\=\.\,\:\;\!\?\%\$\#\@\/\(\)]{4,}', text)
            extracted = " ".join(strings)[:6000]
            return extracted, "pdf", None
        except:
            return "", "pdf", None

    # ── Images — send as base64 to vision
    if filename.endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp')):
        b64 = base64.b64encode(raw).decode()
        ext = filename.split('.')[-1]
        mime = f"image/{'jpeg' if ext in ['jpg','jpeg'] else ext}"
        return "", "image", {"b64": b64, "mime": mime}

    # ── Plain text fallback
    try:
        return raw.decode("utf-8", errors="ignore")[:6000], "text", None
    except:
        return "", "unknown", None


# ══════════════════════════════════════════════════════
# SYSTEM PROMPT
# ══════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are an expert Excel AI. You read any file, understand it deeply, and produce perfect Excel spreadsheets.

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
      }
    }
  ]
}

══════════════════════════════
ABSOLUTE RULES — NEVER BREAK:
══════════════════════════════

RULE 1 — UNDERSTAND THE FILE:
When a file is uploaded you MUST:
- Read every single row without skipping any
- Understand exactly what kind of data it contains
- Complete ALL missing values intelligently
- Never ignore the file content
- Never replace file data with fake generated data

RULE 2 — SOLVE EVERYTHING IN THE FILE:
Math problems (20+14, 5*6, 100-30) → compute and put answer as a number
Questions with blank answers → fill in correct answers
Missing totals → calculate them
Empty grade columns → compute grades
Blank status columns → determine status
The Answer/Result column must contain COMPUTED VALUES not formulas

RULE 3 — COPY ALL ROWS:
If uploaded file has 50 rows → output must have 50 rows
If uploaded file has 100 rows → output must have 100 rows
NEVER reduce the number of rows

RULE 4 — AUTO-UNDERSTAND MODE (no prompt given):
Analyze the file and decide what to do:
- Math problems → solve all, add Answer column
- Student marks → add Total, Average, Grade, Pass/Fail
- Financial data → add totals, summaries, trends, chart
- Inventory → add Stock Status, Reorder Alert
- Employee data → add summaries, department totals
- Survey data → add analysis, counts, percentages
Always add value beyond what the original file had

RULE 5 — REAL NUMBERS:
Prices, salaries, scores → integers or floats, NEVER strings
BAD: ["John", "50000"]
GOOD: ["John", 50000]

RULE 6 — SAFE FORMULAS ONLY:
Always wrap VLOOKUP in IFERROR:
=IFERROR(VLOOKUP(A2,Sheet2!$A:$B,2,FALSE),"Not Found")

Valid formulas:
=SUM(B2:B10)
=AVERAGE(B2:B10)
=MAX(B2:B10)
=MIN(B2:B10)
=COUNT(B2:B10)
=IF(B2>100,"High","Low")
=IF(B2>90,"A",IF(B2>80,"B",IF(B2>70,"C","F")))
=COUNTIF(B2:B10,">100")
=SUMIF(A2:A10,"North",B2:B10)
=IFERROR(VLOOKUP(A2,Sheet2!$A:$B,2,FALSE),"Not Found")
=B2*C2
=B2-C2
=B2/C2*100
=TODAY()

RULE 7 — VLOOKUP SETUP:
Sheet 1: main data, IDs in column A
Sheet 2: reference table, same IDs in column A
IDs must match exactly between sheets
Always use $A:$B absolute references

RULE 8 — CHARTS:
type options: "bar", "line", "pie", "area"
data_cols: list of 0-indexed number columns to plot
category_col: 0-indexed label column
No chart needed: "chart": null
Include chart whenever data is visual/comparative

RULE 9 — CONDITIONAL FORMATTING:
"colorscale" → red-yellow-green gradient
"databar" → blue progress bars
"highlight_high" → green highlight above threshold

RULE 10 — NO PLACEHOLDERS EVER:
Never use: val1, col1, value1, header1, item1, data1, sample, foo, bar, test, name1"""


# ══════════════════════════════════════════════════════
# GROQ CALL
# ══════════════════════════════════════════════════════

async def call_groq(
    prompt: str,
    file_content: str = "",
    file_type: str = "",
    image_data: dict = None
) -> dict:

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    # Build user message
    if image_data:
        # Vision mode for images
        messages.append({
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{image_data['mime']};base64,{image_data['b64']}"
                    }
                },
                {
                    "type": "text",
                    "text": f"""Analyze this image carefully. 
{"Task: " + prompt if prompt.strip() else "Understand what this image contains and create the best Excel spreadsheet from it."}
Read all text, numbers, tables visible in the image.
Solve any math problems, complete any missing data.
Return only JSON."""
                }
            ]
        })

    elif file_content and not prompt.strip():
        # Auto-understand mode
        messages.append({
            "role": "user",
            "content": f"""I uploaded a {file_type} file. Analyze it and create the perfect Excel output.

FILE CONTENT (process ALL rows):
{file_content[:5000]}

- Understand what this data is about
- Complete every missing value
- Solve every math problem if present
- Add useful formulas and charts
- Copy ALL rows, do not skip any
- Return only JSON"""
        })

    elif file_content and prompt.strip():
        # File + prompt
        messages.append({
            "role": "user",
            "content": f"""Task: {prompt}

FILE TYPE: {file_type}
FILE CONTENT (use ALL rows, do not skip any):
{file_content[:5000]}

IMPORTANT:
- Use every single row from the file above
- Complete any missing values or answers
- Do the task described above using this real data
- Return only JSON"""
        })

    else:
        # Prompt only
        messages.append({
            "role": "user",
            "content": f"""Create a professional Excel spreadsheet: "{prompt}"
- Real meaningful data, minimum 10 rows
- Add formulas where useful
- Add chart if it makes sense
- No placeholder values
- Return only JSON"""
        })

    bad_values = [
        "val1","val2","val3","col1","col2","col3",
        "value1","value2","header1","header2",
        "item1","item2","data1","sample1","name1"
    ]

    last_error = None

    for attempt in range(4):
        try:
            temperature = [0.1, 0.2, 0.4, 0.6][attempt]

            # Use vision model for images
            model = "llama-3.3-70b-versatile"
            if image_data:
                model = "meta-llama/llama-4-scout-17b-16e-instruct"

            async with httpx.AsyncClient(timeout=58) as client:
                response = await client.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {os.getenv('GROQ_API_KEY')}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": model,
                        "max_tokens": 4000,
                        "temperature": temperature,
                        "messages": messages
                    }
                )

            result = response.json()

            if "error" in result:
                raise ValueError(f"Groq error: {result['error']}")

            text = result["choices"][0]["message"]["content"].strip()

            # Strip markdown
            text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
            text = re.sub(r"\n?```$", "", text)
            text = text.strip()

            # Extract JSON
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                text = json_match.group(0)

            data = json.loads(text)

            # Validate
            if "sheets" not in data or not data["sheets"]:
                raise ValueError("No sheets in response")

            for sheet in data["sheets"]:
                if not sheet.get("headers"):
                    raise ValueError("Missing headers")
                if not sheet.get("rows"):
                    raise ValueError("Missing rows")

            # Check placeholders (only for non-file requests)
            if not file_content and not image_data:
                json_lower = json.dumps(data).lower()
                found = [b for b in bad_values if b in json_lower]
                if found:
                    messages.append({
                        "role": "user",
                        "content": f"Do NOT use placeholders like {found}. Use real data."
                    })
                    raise ValueError(f"Placeholders: {found}")

            return data

        except json.JSONDecodeError as e:
            last_error = f"JSON error: {e}"
            continue
        except ValueError as e:
            last_error = str(e)
            continue
        except Exception as e:
            last_error = str(e)
            continue

    raise ValueError(f"Failed after 4 attempts. Last: {last_error}")


# ══════════════════════════════════════════════════════
# EXCEL BUILDER
# ══════════════════════════════════════════════════════

def build_excel(data: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def thin_border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def style_header(cell):
        cell.fill = PatternFill("solid", fgColor="2563EB")
        cell.font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()

    def style_data(cell, row_idx, header=""):
        cell.border = thin_border()
        cell.alignment = Alignment(vertical="center")
        if row_idx % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="EFF6FF")
        if isinstance(cell.value, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
            h = header.lower()
            if any(k in h for k in ["salary","revenue","price","cost","amount",
                                      "budget","sales","income","expense","pay","total"]):
                cell.number_format = '#,##0.00'
            elif any(k in h for k in ["percent","%","rate","growth","margin"]):
                cell.number_format = '0.00%'

    for sheet_def in data["sheets"]:
        name = str(sheet_def.get("name", "Sheet"))[:31]
        ws = wb.create_sheet(title=name)

        headers  = sheet_def.get("headers", [])
        rows     = sheet_def.get("rows", [])
        formulas = sheet_def.get("formulas", [])
        cf_rules = sheet_def.get("conditional_formatting", [])
        chart_def = sheet_def.get("chart", None)

        # ── Write headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=str(header))
            style_header(cell)
            ws.column_dimensions[get_column_letter(col)].width = 22
        ws.row_dimensions[1].height = 30

        # ── Write data
        for row_idx, row in enumerate(rows, 2):
            for col_idx, val in enumerate(row, 1):
                header = headers[col_idx-1] if col_idx <= len(headers) else ""
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                style_data(cell, row_idx, header)
            ws.row_dimensions[row_idx].height = 18

        # ── Write formulas
        for f in formulas:
            addr    = f.get("cell", "")
            formula = f.get("formula", "")
            label   = f.get("label", "")
            if addr and formula:
                try:
                    ws[addr] = formula
                    ws[addr].font = Font(bold=True)
                    ws[addr].border = thin_border()
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

        # ── Freeze & filter
        ws.freeze_panes = "A2"
        if headers:
            ws.auto_filter.ref = (
                f"A1:{get_column_letter(len(headers))}{len(rows)+1}"
            )

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
            except:
                pass

        # ── Chart
        if chart_def:
            try:
                ctype    = chart_def.get("type", "bar")
                dcols    = chart_def.get("data_cols", [1])
                catcol   = chart_def.get("category_col", 0)
                ctitle   = chart_def.get("title", "Chart")
                nrows    = len(rows)

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
                chart.width  = 25

                if ctype == "bar":
                    chart.type = "col"

                if ctype == "pie":
                    data_ref = Reference(ws,
                        min_col=dcols[0]+1, min_row=1,
                        max_row=nrows+1)
                    chart.add_data(data_ref, titles_from_data=True)
                else:
                    for dc in dcols:
                        data_ref = Reference(ws,
                            min_col=dc+1, min_row=1,
                            max_row=nrows+1)
                        chart.add_data(data_ref, titles_from_data=True)

                cats = Reference(ws,
                    min_col=catcol+1,
                    min_row=2, max_row=nrows+1)
                chart.set_categories(cats)
                ws.add_chart(chart, f"A{nrows+4}")
            except:
                pass

    wb.save(output_path)


# ══════════════════════════════════════════════════════
# ENDPOINT
# ══════════════════════════════════════════════════════

@app.post("/generate")
async def generate_excel(
    prompt: str = Form(default=""),
    file: UploadFile = File(None)
):
    job_id      = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"

    file_content = ""
    file_type    = ""
    image_data   = None

    if file and file.filename:
        file_content, file_type, image_data = await read_any_file(file)

    try:
        data = await call_groq(prompt, file_content, file_type, image_data)
        build_excel(data, output_path)
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="xlforge_output.xlsx",
        headers={"Access-Control-Allow-Origin": "*"}
            )
