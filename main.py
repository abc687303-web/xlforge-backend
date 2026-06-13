from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import anthropic
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
import os
import uuid
import json
from dotenv import load_dotenv

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

@app.get("/")
def root():
    return {"status": "XLforge backend running!"}

@app.post("/generate")
async def generate_excel(
    prompt: str = Form(...),
    file: UploadFile = File(None)
):
    file_info = ""
    file_data = None

    if file:
        content = await file.read()
        if file.filename.endswith(".csv"):
            import io
            df = pd.read_csv(io.BytesIO(content))
            file_info = f"User uploaded CSV file '{file.filename}' with {len(df)} rows and columns: {list(df.columns)}. First 5 rows: {df.head().to_string()}"
            file_data = df
        elif file.filename.endswith(".xlsx"):
            import io
            df = pd.read_excel(io.BytesIO(content))
            file_info = f"User uploaded Excel file '{file.filename}' with {len(df)} rows and columns: {list(df.columns)}. First 5 rows: {df.head().to_string()}"
            file_data = df

    system_prompt = """You are XLforge, an expert Excel automation AI. 
    Generate Python code using openpyxl and pandas to create Excel files.
    
    IMPORTANT RULES:
    1. Always save the file as: output_file = f'/tmp/{job_id}.xlsx'
    2. Use openpyxl to create and style the Excel file
    3. Add proper formatting, colors, and styles
    4. If user uploaded data, use the dataframe called 'file_data'
    5. Return ONLY executable Python code, no explanations
    6. The variable job_id is already defined
    7. The variable file_data contains the uploaded dataframe (or None)
    8. Always create at least one sheet with data
    9. Add headers with blue background and white text
    10. Make it look professional
    
    Return ONLY Python code that can be executed directly."""

    user_message = f"{prompt}"
    if file_info:
        user_message += f"\n\nFile information: {file_info}"

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}]
    )

    generated_code = message.content[0].text

    # Clean code
    if "```python" in generated_code:
        generated_code = generated_code.split("```python")[1].split("```")[0]
    elif "```" in generated_code:
        generated_code = generated_code.split("```")[1].split("```")[0]

    job_id = str(uuid.uuid4())
    output_path = f"/tmp/{job_id}.xlsx"

    exec_globals = {
        "job_id": job_id,
        "file_data": file_data,
        "pd": pd,
        "openpyxl": openpyxl,
        "PatternFill": PatternFill,
        "Font": Font,
        "Alignment": Alignment,
        "BarChart": BarChart,
        "Reference": Reference,
        "output_file": output_path
    }

    exec(generated_code, exec_globals)

    if os.path.exists(output_path):
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="xlforge_output.xlsx"
        )

    return {"error": "File generation failed"}
