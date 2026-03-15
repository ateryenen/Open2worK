from typing import Literal, Union, List
from pydantic import BaseModel, Field


class OpenAppStep(BaseModel):
    action: Literal["open_app"]
    target: str


class TypeTextStep(BaseModel):
    action: Literal["type_text"]
    text: str


class SaveFileStep(BaseModel):
    action: Literal["save_file"]
    path: str


class WaitStep(BaseModel):
    action: Literal["wait"]
    seconds: float = Field(gt=0)


class ClickImageStep(BaseModel):
    action: Literal["click_image"]
    template_path: str
    confidence: float = Field(default=0.9, gt=0, le=1)
    timeout_seconds: float = Field(default=8.0, gt=0)
    click_count: int = Field(default=1, ge=1, le=10)


PlanStep = Union[OpenAppStep, TypeTextStep, SaveFileStep, WaitStep, ClickImageStep]


class ExecutionPlan(BaseModel):
    goal: str
    steps: List[PlanStep]
