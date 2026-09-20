"""Schemas for document AI chat assistant."""

from __future__ import annotations

from typing import Any, Literal
from pydantic import BaseModel, Field


class ChatMessage(BaseModel):
    role: Literal["user", "assistant", "system"]
    content: str
    timestamp: str | None = None


class ChatRequest(BaseModel):
    messages: list[ChatMessage] = Field(default_factory=list)
    query: str | None = None
    document: dict[str, Any] | None = None


class ChatResponse(BaseModel):
    message: str
    suggestions: list[str] = Field(default_factory=list)
    insights: dict[str, Any] = Field(default_factory=dict)

