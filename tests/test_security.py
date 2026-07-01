"""Tests for docgen-service internal API key security."""

from __future__ import annotations

import io

import pytest
from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)

MINIMAL_RENDER_PAYLOAD = {
    "title": "Test SWMS",
    "siteLocation": "Brisbane",
    "jobActivities": ["Test activity"],
    "plantEquipment": [],
    "tradeName": "Electrician",
    "activityType": "Not Applicable",
    "steps": [
        {
            "stepNo": "1",
            "jobTaskElement": "Test task",
            "sequencePosition": 1,
            "hazards": [
                {
                    "hazard": "Test hazard",
                    "riskLevel": "2",
                    "controls": ["Test control"],
                    "postRiskLevel": "3",
                    "responsiblePerson": "Supervisor",
                }
            ],
        }
    ],
}


def test_health_endpoint_is_public() -> None:
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_parse_rejects_missing_api_key_when_configured(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCGEN_API_KEY", "test-secret")

    response = client.post(
        "/parse",
        data={"activity_type": "Not Applicable"},
        files={"file": ("sample.pdf", io.BytesIO(b"%PDF-1.4"), "application/pdf")},
    )

    assert response.status_code == 401


def test_render_rejects_invalid_api_key_when_configured(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCGEN_API_KEY", "test-secret")

    response = client.post(
        "/render",
        json=MINIMAL_RENDER_PAYLOAD,
        headers={"X-Docgen-Api-Key": "wrong-secret"},
    )

    assert response.status_code == 401


def test_render_accepts_valid_api_key_when_configured(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCGEN_API_KEY", "test-secret")

    response = client.post(
        "/render",
        json=MINIMAL_RENDER_PAYLOAD,
        headers={"X-Docgen-Api-Key": "test-secret"},
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    assert response.content.startswith(b"PK")
