from pydantic import BaseModel
from typing import List, Dict


class ResumeAnalysis(BaseModel):
    filename: str
    overall_score_pct: float
    components: Dict[str, float]
    missing_keywords: List[str]
    top_action_items: List[str]
    resume_preview: str