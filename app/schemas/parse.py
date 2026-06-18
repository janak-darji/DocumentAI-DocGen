from pydantic import BaseModel, Field


class ParsedHazard(BaseModel):
    hazard: str
    riskLevel: str
    controls: list[str]
    postRiskLevel: str
    responsiblePerson: str


class ParsedStep(BaseModel):
    stepNo: str
    jobTaskElement: str
    sequencePosition: int = Field(ge=1)
    hazards: list[ParsedHazard]


class ParsedSwmsResponse(BaseModel):
    activityType: str
    steps: list[ParsedStep]
