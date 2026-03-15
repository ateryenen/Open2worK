import re
import json
import urllib.error
import urllib.request
from typing import Optional

from .config import (
    DEFAULT_FILENAME,
    DEFAULT_TEXT,
    DEFAULT_WAIT_OPEN_SECONDS,
    DEFAULT_WAIT_TYPE_SECONDS,
    SUPPORTED_APP,
    desktop_output_path,
)
from .schemas import ExecutionPlan, OpenAppStep, SaveFileStep, TypeTextStep, WaitStep


class RuleBasedPlanner:
    def __init__(self) -> None:
        self.last_fallback_reasons: list[str] = []

    def build_plan(self, user_input: str) -> ExecutionPlan:
        self.last_fallback_reasons = []
        normalized = user_input.lower()
        if "notepad" not in normalized:
            self.last_fallback_reasons.append("unsupported app intent, fallback to notepad")

        parsed_text = self._extract_text(user_input)
        text = parsed_text or DEFAULT_TEXT
        if parsed_text is None:
            self.last_fallback_reasons.append("text not parsed, fallback to default text")

        parsed_filename = self._extract_filename(user_input)
        filename = parsed_filename or DEFAULT_FILENAME
        if parsed_filename is None:
            self.last_fallback_reasons.append("filename not parsed, fallback to default filename")

        path = str(desktop_output_path(filename))

        return ExecutionPlan(
            goal="Open Notepad and save text",
            steps=[
                OpenAppStep(action="open_app", target=SUPPORTED_APP),
                WaitStep(action="wait", seconds=DEFAULT_WAIT_OPEN_SECONDS),
                TypeTextStep(action="type_text", text=text),
                WaitStep(action="wait", seconds=DEFAULT_WAIT_TYPE_SECONDS),
                SaveFileStep(action="save_file", path=path),
            ],
        )

    def _extract_text(self, user_input: str) -> Optional[str]:
        quoted = re.search(r'"([^"]+)"', user_input)
        if quoted:
            return quoted.group(1)

        type_pattern = re.search(r"type\s+(.+?)(?:,|\s+and\s+save|$)", user_input, re.IGNORECASE)
        if type_pattern:
            return type_pattern.group(1).strip().strip("\"')")

        return None

    def _extract_filename(self, user_input: str) -> Optional[str]:
        as_pattern = re.search(r"as\s+([a-zA-Z0-9_.-]+\.txt)", user_input, re.IGNORECASE)
        if as_pattern:
            return as_pattern.group(1)

        filename_pattern = re.search(r"([a-zA-Z0-9_.-]+\.txt)", user_input)
        if filename_pattern:
            return filename_pattern.group(1)

        return None


def local_llm_planner_interface(
    user_input: str,
    model: str = "qwen2.5:latest",
    endpoint: str = "http://127.0.0.1:11434/api/generate",
    timeout_seconds: int = 90,
) -> ExecutionPlan:
    schema_hint = {
        "goal": "string",
        "steps": [
            {
                "action": "open_app|wait|type_text|save_file|click_image",
                "target": "for open_app",
                "seconds": 1.0,
                "text": "for type_text",
                "path": "for save_file",
                "template_path": "for click_image",
                "confidence": 0.9,
                "timeout_seconds": 8.0,
                "click_count": 1,
            }
        ],
    }
    prompt = (
        "You are a desktop automation planner. "
        "Return ONLY valid JSON with no markdown. "
        "Build steps only from supported actions. "
        "User request: "
        f"{user_input}\n"
        "If user asks image click and template path is missing, use C:/GITHUB/Open2worK/assets/T357.png. "
        "If user asks to click twice, set click_count=2. "
        "JSON schema hint: "
        f"{json.dumps(schema_hint, ensure_ascii=False)}"
    )

    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0},
    }

    req = urllib.request.Request(
        url=endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
            body = response.read().decode("utf-8")
        data = json.loads(body)
        raw = _extract_json_text(str(data.get("response", "")).strip())
        plan_dict = json.loads(raw)
        normalized = _normalize_plan_dict(plan_dict, user_input)
        return ExecutionPlan.model_validate(normalized)
    except urllib.error.HTTPError as e:
        if e.code == 404 and endpoint.endswith("/api/generate"):
            return _local_llm_planner_chat(
                user_input=user_input,
                model=model,
                endpoint=endpoint,
                timeout_seconds=timeout_seconds,
            )
        raise RuntimeError(f"LLM planner endpoint error ({e.code}) at {endpoint}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"cannot reach local LLM endpoint {endpoint}: {e}")
    except Exception as e:
        raise RuntimeError(f"LLM planner parse/validation failed: {e}")


def _local_llm_planner_chat(
    user_input: str,
    model: str,
    endpoint: str,
    timeout_seconds: int,
) -> ExecutionPlan:
    chat_endpoint = endpoint.replace("/api/generate", "/api/chat")
    prompt = (
        "Return ONLY valid JSON plan. Supported actions: open_app, wait, type_text, save_file, click_image. "
        "If image template path missing, use C:/GITHUB/Open2worK/assets/T357.png. "
        "If user says click twice, set click_count=2. "
        f"User request: {user_input}"
    )
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "options": {"temperature": 0},
    }
    req = urllib.request.Request(
        url=chat_endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
        body = response.read().decode("utf-8")
    data = json.loads(body)
    raw = _extract_json_text(str(data.get("message", {}).get("content", "")).strip())
    plan_dict = json.loads(raw)
    normalized = _normalize_plan_dict(plan_dict, user_input)
    return ExecutionPlan.model_validate(normalized)


def _extract_json_text(raw: str) -> str:
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        text = text.strip()

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return text[start : end + 1]
    return text


def _normalize_plan_dict(plan_dict: dict, user_input: str) -> dict:
    goal = str(plan_dict.get("goal", "LLM generated plan")).strip() or "LLM generated plan"
    steps = plan_dict.get("steps", [])
    if not isinstance(steps, list):
        steps = []

    normalized_steps: list[dict] = []
    allowed_actions = {"open_app", "wait", "type_text", "save_file", "click_image"}

    for step in steps:
        if not isinstance(step, dict):
            continue

        action = str(step.get("action", "")).strip().lower()
        if action not in allowed_actions:
            if "template_path" in step:
                action = "click_image"
            elif "target" in step:
                action = "open_app"
            elif "seconds" in step:
                action = "wait"
            elif "text" in step:
                action = "type_text"
            elif "path" in step:
                action = "save_file"
            else:
                continue

        if action == "open_app":
            normalized_steps.append({"action": "open_app", "target": str(step.get("target", "notepad"))})
        elif action == "wait":
            normalized_steps.append({"action": "wait", "seconds": float(step.get("seconds", 1.0))})
        elif action == "type_text":
            normalized_steps.append({"action": "type_text", "text": str(step.get("text", ""))})
        elif action == "save_file":
            normalized_steps.append({"action": "save_file", "path": str(step.get("path", ""))})
        elif action == "click_image":
            template_path = str(step.get("template_path", "")).strip()
            if (
                not template_path
                or template_path.lower() == "for click_image"
                or template_path.lower().startswith("for ")
            ):
                template_path = "C:/GITHUB/Open2worK/assets/T357.png"
            normalized_steps.append(
                {
                    "action": "click_image",
                    "template_path": template_path,
                    "confidence": float(step.get("confidence", 0.9)),
                    "timeout_seconds": float(step.get("timeout_seconds", 8.0)),
                    "click_count": int(step.get("click_count", 1)),
                }
            )

    click_steps = [s for s in normalized_steps if s.get("action") == "click_image"]
    if click_steps:
        inferred_double = any(token in user_input for token in ["兩次", "2次", "twice", "double"])
        merged_click_count = max(max(1, int(s.get("click_count", 1))) for s in click_steps)
        if inferred_double:
            merged_click_count = 2

        first_click = click_steps[0].copy()
        first_click["click_count"] = merged_click_count
        normalized_steps = [s for s in normalized_steps if s.get("action") != "click_image"]
        normalized_steps.insert(0, first_click)

    return {"goal": goal, "steps": normalized_steps}
