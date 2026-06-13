from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import os, uuid, httpx

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
async def generate_excel(prompt: str = Form(...)):
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
                "messages": [{"role": "user", "content": f"""Generate Python code using openpyxl to create an Excel file.
Save the file to: {output_path}
Request: {prompt}

Rules:
- Use openpyxl only
- Add blue header row with white bold text
- Add sample data rows
- Make it look professional
- Return ONLY Python code, no markdown, no explanation"""}]
            }
        )

    code = response.json()["choices"][0]["message"]["content"]
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
        headers={"Access-Control-Allow-Origin": "*"}
    )
