from typing import Literal, Union, List
from pydantic import BaseModel, Field


class OpenAppStep(BaseModel):
    action: Literal["open_app"]
    target: str


class TypeTextStep(BaseModel):
    action: Literal["type_text"]
    text: str


class KeyPressStep(BaseModel):
    action: Literal["key_press"]
    key: str
    repeat: int = Field(default=1, ge=1, le=50)
    interval: float = Field(default=0.05, ge=0, le=2)


class SaveFileStep(BaseModel):
    action: Literal["save_file"]
    path: str


class WaitStep(BaseModel):
    action: Literal["wait"]
    seconds: float = Field(gt=0)


class ClickImageStep(BaseModel):
    action: Literal["click_image"]
    template_path: str
    confidence: float = Field(default=0.8, gt=0, le=1)
    timeout_seconds: float = Field(default=8.0, gt=0)
    click_count: int = Field(default=1, ge=1, le=10)
    press_duration: float = Field(default=0.03, ge=0, le=2)
    click_mode: Literal["mouse_event", "send_input"] = "mouse_event"


class MoveMouseHorizontalStep(BaseModel):
    action: Literal["move_mouse_horizontal"]
    direction: Literal["left", "right"] = "right"
    pixels: int = Field(default=100, ge=1, le=5000)
    duration: float = Field(default=0.2, ge=0, le=10)


class MouseClickStep(BaseModel):
    action: Literal["mouse_click"]
    button: Literal["left", "right"] = "left"
    click_count: int = Field(default=1, ge=1, le=10)
    interval: float = Field(default=0.1, ge=0, le=2)
    press_duration: float = Field(default=0.03, ge=0, le=2)
    click_mode: Literal["mouse_event", "send_input"] = "mouse_event"


PlanStep = Union[
    OpenAppStep,
    TypeTextStep,
    KeyPressStep,
    SaveFileStep,
    WaitStep,
    ClickImageStep,
    MoveMouseHorizontalStep,
    MouseClickStep,
]


class ExecutionPlan(BaseModel):
    goal: str
    steps: List[PlanStep]
