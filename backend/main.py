from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
from models import EvaluationResult, EngineSnapshot
from parser import parse_docx_sheet
from evaluator import evaluate_engine

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/evaluate", response_model=EvaluationResult)
async def evaluate_file(file: UploadFile = File(...)):
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Strict Integrity: Only .docx allowed.")
    
    temp_path = f"temp_{file.filename}"
    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        raw_data = parse_docx_sheet(temp_path)
        snapshot = EngineSnapshot(**raw_data) # Pydantic Validation
        result = evaluate_engine(snapshot)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        os.remove(temp_path)

# Run with: uvicorn main:app --reload
