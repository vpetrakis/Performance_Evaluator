from pydantic import BaseModel, Field
from typing import List

class CylinderData(BaseModel):
    id: int = Field(..., ge=1, le=6)
    p_max: float = Field(..., gt=0, description="Max Firing Pressure (bar)")
    p_comp: float = Field(..., gt=0, description="Compression Pressure (bar)")
    exhaust_temp: float = Field(..., gt=100, lt=600, description="Exhaust Temp (°C)")

class EngineSnapshot(BaseModel):
    vessel_name: str
    rpm: int
    scavenge_temp: float
    average_exhaust: float
    cylinders: List[CylinderData]

class EvaluationResult(BaseModel):
    status: str # "GREEN", "YELLOW", "RED"
    alerts: List[str]
    snapshot: EngineSnapshot
