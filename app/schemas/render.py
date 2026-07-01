"""Schemas for SWMS Word document rendering."""

from __future__ import annotations

from pydantic import BaseModel, Field


class RenderHazard(BaseModel):
    hazard: str
    riskLevel: str
    controls: list[str] = Field(default_factory=list)
    postRiskLevel: str
    responsiblePerson: str


class RenderStep(BaseModel):
    stepNo: str
    jobTaskElement: str
    sequencePosition: int = Field(ge=1)
    hazards: list[RenderHazard] = Field(min_length=1)


class RenderSwmsRequest(BaseModel):
    title: str
    siteLocation: str
    jobActivities: list[str] = Field(default_factory=list)
    plantEquipment: list[str] = Field(default_factory=list)
    tradeName: str
    activityType: str = "Not Applicable"
    swmsIssueDate: str | None = None
    companyName: str | None = None
    steps: list[RenderStep] = Field(min_length=1)
