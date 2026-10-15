"""Schemas for document section and general information extraction."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class ExtractedField(BaseModel):
    """Dynamic General Information field definition extracted from document."""

    fieldKey: str
    fieldName: str
    fieldType: str = "text"
    defaultValue: str | None = None
    placeholder: str | None = None
    helpText: str | None = None
    isRequired: bool = False
    orderIndex: int = 0


class ExtractedSectionNode(BaseModel):
    """Hierarchical document section node supporting arbitrary nesting depth."""

    number: str | None = None
    sectionNumber: str | None = None
    level: int = 1
    sectionKey: str
    title: str
    content: str = ""
    orderIndex: int = 0
    pageNumber: int | None = None
    confidence: float = 1.0
    isAmbiguous: bool = False
    children: list[ExtractedSectionNode] = Field(default_factory=list)
    subSections: list[ExtractedSectionNode] = Field(default_factory=list)

    def model_post_init(self, __context: Any) -> None:
        # Keep number and sectionNumber in sync
        if self.number is not None and self.sectionNumber is None:
            self.sectionNumber = self.number
        elif self.sectionNumber is not None and self.number is None:
            self.number = self.sectionNumber

        # Keep children and subSections in sync
        if self.children and not self.subSections:
            self.subSections = self.children
        elif self.subSections and not self.children:
            self.children = self.subSections


ExtractedSectionNode.model_rebuild()

# Backward-compatibility aliases
ExtractedMainSection = ExtractedSectionNode
ExtractedSubSection = ExtractedSectionNode


class ParsedDocumentSectionsResponse(BaseModel):
    """Full extraction response including title, fields, and section hierarchy."""

    documentTitle: str
    generalInformationFields: list[ExtractedField] = Field(default_factory=list)
    sections: list[ExtractedSectionNode] = Field(default_factory=list)
    tableOfContents: list[dict] = Field(default_factory=list)
