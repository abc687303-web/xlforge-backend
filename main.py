from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import uuid

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "XLforge backend running!"}

@app.post("/generate")
async def generate_excel(prompt: str = Form(...)):
    import anthropic
    import openpyxl
    from openpyxl.styles import PatternFill, Font
    
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    
    job_id = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"
    
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        messages=[{"role": "user", "content": f"Generate Python code using openpyxl to create an Excel file saved at '{output_path}' based on this request: {prompt}. Return ONLY executable Python code, no explanations, no markdown."}]
    )
    
    code = message.content[0].text
    if "```" in code:
        code = code.split("```")[1]
        if code.startswith("python"):
            code = code[6:]
    
    exec(compile(code, "<string>", "exec"), {
        "openpyxl": openpyxl,
        "PatternFill": PatternFill,
        "Font": Font,
        "output_path": output_path
    })
    
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="xlforge_output.xlsx"
    )
