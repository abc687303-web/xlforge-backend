from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, GradientFill
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule
from openpyxl.styles.differential import DifferentialStyle
import os, uuid, httpx, json, re

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

# ── PROMPT ──────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are an expert Excel spreadsheet generator. You create realistic, professional spreadsheets with real data.

OUTPUT: Return ONLY a valid JSON object. No markdown. No explanation. No code blocks.

JSON STRUCTURE:
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
        {"range": "D2:D20", "type": "highlight_high", "threshold": 1000}
      ],
      "chart": {
        "type": "bar",
        "title": "My Chart",
        "data_cols": [1, 2],
        "category_col": 0
      }
    }
  ]
}

═══════════════════════════════════════
CRITICAL RULES — NEVER BREAK THESE:
═══════════════════════════════════════

RULE 1 — REAL DATA ONLY:
NEVER use: val1, val2, col1, col2, value1, header1, item1, data1, sample, foo, bar, test
ALWAYS use real names, real numbers, real dates relevant to the user's request.

RULE 2 — NUMBERS MUST BE NUMBERS:
Revenue, prices, quantities, scores, ages → use actual integers or floats, NOT strings.
BAD:  ["John", "50000"]
GOOD: ["John", 50000]

RULE 3 — VLOOKUP SETUP (when user asks for VLOOKUP):
- Sheet 1: Main sheet with a lookup column (e.g. Employee ID, Product Code)
- Sheet 2: Reference table with matching IDs + extra info
- VLOOKUP formula: =VLOOKUP(A2,Sheet2Name!$A:$D,2,FALSE)
- IDs in sheet 1 MUST exist in sheet 2
- Example for employee salary lookup:
  Sheet1 headers: ["Emp ID", "Name", "Department", "Salary"]
  Sheet1 formula: {"cell": "D2", "formula": "=VLOOKUP(A2,SalaryTable!$A:$B,2,FALSE)"}
  Sheet2 name: "SalaryTable"
  Sheet2 headers: ["Emp ID", "Salary"]
  Sheet2 rows: [["E001", 55000], ["E002", 62000]]

RULE 4 — FORMULA REFERENCE:
=SUM(B2:B10)           → sum a range
=AVERAGE(B2:B10)       → average
=MAX(B2:B10)           → maximum
=MIN(B2:B10)           → minimum  
=COUNT(B2:B10)         → count numbers
=COUNTA(A2:A10)        → count non-empty
=IF(B2>100,"High","Low")  → conditional
=IF(B2>1000,"Excellent",IF(B2>500,"Good","Poor"))  → nested IF
=COUNTIF(B2:B10,">100")   → count with condition
=SUMIF(A2:A10,"North",B2:B10)  → sum with condition
=VLOOKUP(A2,Sheet2!$A:$D,2,FALSE)  → lookup from another sheet
=IFERROR(VLOOKUP(A2,Sheet2!$A:$D,2,FALSE),"Not Found")  → safe VLOOKUP
=B2*C2                 → multiply
=B2/C2*100             → percentage
=TODAY()               → today's date
=CONCATENATE(A2," ",B2) → join text

RULE 5 — CHARTS:
chart type options: "bar", "line", "pie", "area"
data_cols: list of 0-indexed column numbers containing numbers to plot
category_col: 0-indexed column containing labels (text)
If user says "chart" or "graph" → always include chart
If no chart needed → "chart": null

RULE 6 — CONDITIONAL FORMATTING:
"colorscale" → green-yellow-red color gradient on numbers
"databar"    → blue progress bars in cells
"highlight_high" → highlights cells above threshold in green

RULE 7 — MULTIPLE SHEETS:
If user asks for dashboard, report, or analysis → create 2-3 sheets
Sheet 1: Raw data
Sheet 2: Summary with formulas
Sheet 3 (optional): Chart data

RULE 8 — ROW COUNT:
Always generate at least 10 rows of real data unless user specifies otherwise.
More rows = more realistic = better.

EXAMPLES OF GOOD RESPONSES:

User: "sales dashboard"
→ Sheet 1 "Sales Data": columns [Month, Region, Product, Units Sold, Unit Price, Revenue]
  with 12 rows of real monthly data, Revenue formula =D2*E2
→ Sheet 2 "Summary": total revenue, avg per month, best month using SUM/AVERAGE/MAX formulas
→ Include bar chart on Sheet 1 showing Revenue by Month

User: "employee salary with vlookup"  
→ Sheet 1 "Employees": [Emp ID, Name, Department, Salary] — salary via VLOOKUP
→ Sheet 2 "SalaryTable": [Emp ID, Salary] — reference data

User: "budget tracker"
→ Sheet 1 "Budget": [Category, Budgeted, Actual, Variance, Status]
  Variance = =C2-B2, Status = =IF(D2>0,"Over Budget","Under Budget")
→ Include pie chart of budgeted amounts"""


# ── GROQ CALL ───────────────────────────────────────────────────────────────

async def call_groq(prompt: str, file_content: str = "") -> dict:
    user_msg = f"""Create a professional Excel spreadsheet for this request: "{prompt}"

Remember:
- Use REAL data, real names, real numbers
- At least 10 rows
- Include formulas where useful
- Include chart if it makes sense
- NO placeholder values ever"""

    if file_content:
        user_msg += f"\n\nUser uploaded this data — use it:\n{file_content[:4000]}"

    bad_values = [
        "val1", "val2", "val3", "col1", "col2", "col3",
        "value1", "value2", "header1", "header2",
        "item1", "item2", "data1", "sample1", "name1",
        "product1", "category1"
    ]

    last_error = None

    for attempt in range(4):
        try:
            # Escalate temperature on retries to get different output
            temperature = [0.2, 0.4, 0.6, 0.7][attempt]

            async with httpx.AsyncClient(timeout=58) as client:
                response = await client.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {os.getenv('GROQ_API_KEY')}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": "llama-3.3-70b-versatile",
                        "max_tokens": 4000,
                        "temperature": temperature,
                        "messages": [
                            {"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": user_msg}
                        ]
                    }
                )

            result = response.json()
            text = result["choices"][0]["message"]["content"].strip()

            # Strip markdown fences
            text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
            text = re.sub(r"\n?```$", "", text)
            text = text.strip()

            # Extract JSON if surrounded by other text
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                text = json_match.group(0)

            data = json.loads(text)

            # Validate structure
            if "sheets" not in data or not data["sheets"]:
                raise ValueError("No sheets in response")

            # Check for placeholder values
            json_lower = json.dumps(data).lower()
            found_bad = [b for b in bad_values if b in json_lower]
            if found_bad:
                user_msg += f"\n\nIMPORTANT: Do NOT use placeholder values like {found_bad}. Use real data!"
                raise ValueError(f"Placeholder values detected: {found_bad}")

            # Validate each sheet has real data
            for sheet in data["sheets"]:
                if not sheet.get("headers"):
                    raise ValueError("Sheet missing headers")
                if not sheet.get("rows") or len(sheet["rows"]) < 1:
                    raise ValueError("Sheet has no rows")

            return data

        except json.JSONDecodeError as e:
            last_error = f"JSON parse error: {e}"
            continue
        except ValueError as e:
            last_error = str(e)
            continue
        except Exception as e:
            last_error = str(e)
            if attempt == 3:
                raise

    raise ValueError(f"Failed after 4 attempts. Last error: {last_error}")


# ── EXCEL BUILDER ───────────────────────────────────────────────────────────

def build_excel(data: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Styles
    BLUE       = "2563EB"
    LIGHT_BLUE = "EFF6FF"
    WHITE      = "FFFFFF"
    DARK       = "1E293B"
    GREEN      = "16A34A"
    RED        = "DC2626"
    AMBER      = "D97706"

    def header_style(cell, color=BLUE):
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(bold=True, color=WHITE, size=11, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()

    def thin_border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def data_style(cell, row_idx):
        cell.border = thin_border()
        cell.alignment = Alignment(vertical="center")
        if row_idx % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    for sheet_def in data["sheets"]:
        ws = wb.create_sheet(title=str(sheet_def.get("name", "Sheet"))[:31])
        headers  = sheet_def.get("headers", [])
        rows     = sheet_def.get("rows", [])
        formulas = sheet_def.get("formulas", [])
        cf_rules = sheet_def.get("conditional_formatting", [])
        chart_def = sheet_def.get("chart", None)

        # ── Headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=str(header))
            header_style(cell)
            ws.column_dimensions[get_column_letter(col)].width = 20
        ws.row_dimensions[1].height = 30

        # ── Data rows
        for row_idx, row in enumerate(rows, 2):
            for col_idx, val in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                data_style(cell, row_idx)
                # Right-align numbers
                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    # Format as currency if header suggests it
                    header_name = headers[col_idx-1].lower() if col_idx <= len(headers) else ""
                    if any(k in header_name for k in ["salary", "revenue", "price", "cost", "amount", "budget", "sales", "income", "expense"]):
                        cell.number_format = '#,##0.00'
                    elif any(k in header_name for k in ["percent", "%", "rate", "growth"]):
                        cell.number_format = '0.00%'
            ws.row_dimensions[row_idx].height = 18

        # ── Formulas
        formula_row_start = len(rows) + 2
        for f in formulas:
            cell_addr = f.get("cell", "")
            formula   = f.get("formula", "")
            label     = f.get("label", "")
            if cell_addr and formula:
                ws[cell_addr] = formula
                ws[cell_addr].border = thin_border()
                ws[cell_addr].font = Font(bold=True, color=DARK)
                # Add label in column before if possible
                try:
                    col_letter = re.match(r"([A-Z]+)", cell_addr).group(1)
                    row_num = int(re.search(r"(\d+)", cell_addr).group(1))
                    col_idx = column_index_from_string(col_letter)
                    if col_idx > 1 and label:
                        label_cell = ws.cell(row=row_num, column=col_idx-1, value=label)
                        label_cell.font = Font(bold=True)
                except:
                    pass

        # ── Freeze panes
        ws.freeze_panes = "A2"

        # ── Auto filter
        if headers:
            last_col = get_column_letter(len(headers))
            ws.auto_filter.ref = f"A1:{last_col}{len(rows)+1}"

        # ── Conditional formatting
        for cf in cf_rules:
            cf_range = cf.get("range", "")
            cf_type  = cf.get("type", "")
            if not cf_range:
                continue
            try:
                if cf_type == "colorscale":
                    ws.conditional_formatting.add(cf_range, ColorScaleRule(
                        start_type="min", start_color="F87171",
                        mid_type="percentile", mid_value=50, mid_color="FCD34D",
                        end_type="max", end_color="4ADE80"
                    ))
                elif cf_type == "databar":
                    ws.conditional_formatting.add(cf_range, DataBarRule(
                        start_type="min", start_value=0,
                        end_type="max", end_value=100,
                        color="2563EB"
                    ))
                elif cf_type == "highlight_high":
                    threshold = cf.get("threshold", 0)
                    green_fill = PatternFill(bgColor="4ADE80")
                    green_font = Font(color="14532D", bold=True)
                    ds = DifferentialStyle(fill=green_fill, font=green_font)
                    ws.conditional_formatting.add(cf_range, CellIsRule(
                        operator="greaterThan",
                        formula=[str(threshold)],
                        stopIfTrue=True,
                        dxf=ds
                    ))
            except:
                pass  # Skip bad CF rules silently

        # ── Chart
        if chart_def:
            try:
                chart_type   = chart_def.get("type", "bar")
                data_cols    = chart_def.get("data_cols", [1])
                category_col = chart_def.get("category_col", 0)
                chart_title  = chart_def.get("title", "Chart")
                num_rows     = len(rows)

                if chart_type == "pie" and data_cols:
                    chart = PieChart()
                    chart.title = chart_title
                    chart.style = 10
                    chart.height = 15
                    chart.width  = 25
                    data_ref = Reference(ws,
                        min_col=data_cols[0]+1, min_row=1,
                        max_row=num_rows+1)
                    chart.add_data(data_ref, titles_from_data=True)
                    cats = Reference(ws,
                        min_col=category_col+1,
                        min_row=2, max_row=num_rows+1)
                    chart.set_categories(cats)

                elif chart_type == "line":
                    chart = LineChart()
                    chart.title  = chart_title
                    chart.style  = 10
                    chart.height = 15
                    chart.width  = 25
                    for dc in data_cols:
                        data_ref = Reference(ws,
                            min_col=dc+1, min_row=1,
                            max_row=num_rows+1)
                        chart.add_data(data_ref, titles_from_data=True)
                    cats = Reference(ws,
                        min_col=category_col+1,
                        min_row=2, max_row=num_rows+1)
                    chart.set_categories(cats)

                elif chart_type == "area":
                    chart = AreaChart()
                    chart.title  = chart_title
                    chart.style  = 10
                    chart.height = 15
                    chart.width  = 25
                    for dc in data_cols:
                        data_ref = Reference(ws,
                            min_col=dc+1, min_row=1,
                            max_row=num_rows+1)
                        chart.add_data(data_ref, titles_from_data=True)
                    cats = Reference(ws,
                        min_col=category_col+1,
                        min_row=2, max_row=num_rows+1)
                    chart.set_categories(cats)

                else:  # bar (default)
                    chart = BarChart()
                    chart.title  = chart_title
                    chart.style  = 10
                    chart.height = 15
                    chart.width  = 25
                    chart.type   = "col"
                    for dc in data_cols:
                        data_ref = Reference(ws,
                            min_col=dc+1, min_row=1,
                            max_row=num_rows+1)
                        chart.add_data(data_ref, titles_from_data=True)
                    cats = Reference(ws,
                        min_col=category_col+1,
                        min_row=2, max_row=num_rows+1)
                    chart.set_categories(cats)

                ws.add_chart(chart, f"A{num_rows + 4}")
            except:
                pass  # Skip bad chart silently

    wb.save(output_path)


# ── ENDPOINT ─────────────────────────────────────────────────────────────────

@app.post("/generate")
async def generate_excel(prompt: str = Form(...), file: UploadFile = File(None)):
    job_id      = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"

    file_content = ""
    if file:
        try:
            raw = await file.read()
            file_content = raw.decode("utf-8", errors="ignore")
        except:
            file_content = ""

    try:
        data = await call_groq(prompt, file_content)
        build_excel(data, output_path)
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"detail": f"Error: {str(e)}"}
        )

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="xlforge_output.xlsx",
        headers={"Access-Control-Allow-Origin": "*"}
)
