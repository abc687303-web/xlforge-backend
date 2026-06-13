from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import anthropic
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import os
import uuid

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
    return {"status":"XLforge backend running!"}

@app.options("/generate")
def options_generate():
    return JSONResponse(content={}, headers={
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "*",
    })

@app.post("/generate")
async def generate_excel(prompt: str = Form(...)):
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    job_id = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        messages=[{"role":"user","content":f"""Generate Python code using openpyxl to create an Excel file.
Save the file to: {output_path}
Request: {prompt}

Rules:
- Use openpyxl only
- Add blue header row with white bold text
- Add sample data rows
- Make it look professional
- Return ONLY Python code, no markdown, no explanation"""}]
    )

    code = message.content[0].text
    if "```" in code:
        parts = code.split("```")
        for part in parts:
            if part.startswith("python"):
                code = part[6:]
                break
            elif len(part) > 50 and "openpyxl" in part:
                code = part
                break

    local_vars = {
        "openpyxl": openpyxl,
        "PatternFill": PatternFill,
        "Font": Font,
        "Alignment": Alignment,
        "output_path": output_path,
        "os": os
    }
    exec(code, local_vars)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="xlforge_output.xlsx",
        headers={"Access-Control-Allow-Origin":"*"}
    )
