from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
import os, uuid, httpx, json

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
    return {"status": "XLforge backend running!"}

@app.post("/generate")
async def generate_excel(prompt: str = Form(...), file: UploadFile = File(None)):
    job_id = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"

    async with httpx.AsyncClient(timeout=60) as client:
        response = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {os.getenv('GROQ_API_KEY')}",
                "Content-Type": "application/json"
            },
            json={
                "model": "llama-3.3-70b-versatile",
                "max_tokens": 2000,
                "messages": [{"role": "user", "content": f"""You are an Excel data generator. Return ONLY a JSON object with this exact structure:
{{
  "sheets": [
    {{
      "name": "Sheet name",
      "headers": ["Col1", "Col2", "Col3"],
      "rows": [
        ["val1", "val2", "val3"],
        ["val1", "val2", "val3"]
      ],
      "chart": false
    }}
  ]
}}

If the user wants a chart, add a second sheet with "chart": true and the same headers/rows as the data sheet.

Request: {prompt}

Return ONLY valid JSON, no markdown, no explanation."""}]
            }
        )

    text = response.json()["choices"][0]["message"]["content"].strip()
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    data = json.loads(text)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    blue_fill = PatternFill("solid", fgColor="2563EB")
    white_bold = Font(bold=True, color="FFFFFF")

    for sheet_def in data["sheets"]:
        ws = wb.create_sheet(title=sheet_def["name"])

        if sheet_def.get("chart"):
            first_ws = wb.worksheets[0]
            for row in first_ws.iter_rows():
                for cell in row:
                    ws[cell.coordinate] = cell.value

            chart = BarChart()
            chart.title = sheet_def["name"]
            chart.style = 10
            chart.height = 15
            chart.width = 25
            data_ref = Reference(ws, min_col=2, min_row=1,
                                 max_col=len(sheet_def["headers"]),
                                 max_row=len(sheet_def["rows"]) + 1)
            cats = Reference(ws, min_col=1, min_row=2,
                             max_row=len(sheet_def["rows"]) + 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats)
            ws.add_chart(chart, "A" + str(len(sheet_def["rows"]) + 4))
        else:
            for col, header in enumerate(sheet_def["headers"], 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.fill = blue_fill
                cell.font = white_bold
                cell.alignment = Alignment(horizontal="center")
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18

            for row_idx, row in enumerate(sheet_def["rows"], 2):
                for col_idx, val in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=val)

    wb.save(output_path)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="xlforge_output.xlsx",
        headers={"Access-Control-Allow-Origin": "*"}
    )
