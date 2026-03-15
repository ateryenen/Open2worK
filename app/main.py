import argparse
from pathlib import Path

from .executor import DeterministicExecutor
from .monitor import LocalLLMScreenMonitor
from .planner import RuleBasedPlanner, local_llm_planner_interface
from .schemas import ClickImageStep, ExecutionPlan
from .utils import log, verify_file_exists


def run(
    user_input: str,
    dry_run: bool = False,
    monitor_enabled: bool = False,
    monitor_interval: float = 2.0,
    monitor_model: str = "phi:latest",
    monitor_endpoint: str = "http://127.0.0.1:11434/api/generate",
    image_template: str = "",
    image_timeout: float = 8.0,
    image_confidence: float = 0.9,
    image_click_count: int = 1,
    use_llm_planner: bool = False,
    planner_model: str = "qwen2.5:latest",
    planner_endpoint: str = "http://127.0.0.1:11434/api/generate",
    planner_timeout: int = 90,
) -> int:
    log(f"input received: {user_input}")

    planner = RuleBasedPlanner()
    if image_template:
        plan = ExecutionPlan(
            goal="Image click POC",
            steps=[
                ClickImageStep(
                    action="click_image",
                    template_path=image_template,
                    timeout_seconds=image_timeout,
                    confidence=image_confidence,
                    click_count=image_click_count,
                )
            ],
        )
    elif use_llm_planner:
        try:
            plan = local_llm_planner_interface(
                user_input=user_input,
                model=planner_model,
                endpoint=planner_endpoint,
                timeout_seconds=planner_timeout,
            )
            log(f"llm planner used: model={planner_model}")
        except Exception as e:
            log(f"llm planner failed, fallback to rule planner: {e}")
            plan = planner.build_plan(user_input)
    else:
        plan = planner.build_plan(user_input)
    log(f"plan generated: {plan.model_dump_json()}")
    if not image_template and planner.last_fallback_reasons:
        log(f"planner fallback: {'; '.join(planner.last_fallback_reasons)}")

    if dry_run:
        log("dry-run enabled: skip desktop automation and verification")
        log("execution result: success (dry-run)")
        return 0

    monitor = LocalLLMScreenMonitor(
        enabled=monitor_enabled,
        interval_seconds=monitor_interval,
        model=monitor_model,
        endpoint=monitor_endpoint,
    )
    monitor.start()

    executor = DeterministicExecutor(monitor=monitor)
    try:
        executor.execute(plan)
    finally:
        monitor.stop()

    actions = [getattr(step, "action", "") for step in plan.steps]
    image_only_plan = bool(actions) and all(action == "click_image" for action in actions)

    if image_template or image_only_plan:
        log("execution result: success (image click poc)")
        return 0

    notepad_ok = executor.is_notepad_running()
    save_targets = [step.path for step in plan.steps if getattr(step, "action", "") == "save_file"]
    output_path = save_targets[-1] if save_targets else ""
    file_ok = bool(output_path) and verify_file_exists(output_path)

    log(f"verification - notepad running: {notepad_ok}")
    log(f"verification - file exists: {file_ok} ({Path(output_path)})")

    success = notepad_ok and file_ok
    if success:
        log("execution result: success")
        return 0

    log("execution result: failed")
    return 1


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="AI Desktop Agent POC")
    parser.add_argument(
        "-c",
        "--command",
        type=str,
        help="Natural language command for planner",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Build and print plan only, do not control Notepad",
    )
    parser.add_argument(
        "--monitor",
        action="store_true",
        help="Enable local LLM screen monitoring during execution",
    )
    parser.add_argument(
        "--monitor-interval",
        type=float,
        default=2.0,
        help="Screen monitoring interval in seconds",
    )
    parser.add_argument(
        "--monitor-model",
        type=str,
        default="phi:latest",
        help="Local vision model name in Ollama",
    )
    parser.add_argument(
        "--monitor-endpoint",
        type=str,
        default="http://127.0.0.1:11434/api/generate",
        help="Local LLM HTTP endpoint for vision inference",
    )
    parser.add_argument(
        "--image-template",
        type=str,
        default="",
        help="Run image-click POC with template image path",
    )
    parser.add_argument(
        "--image-timeout",
        type=float,
        default=8.0,
        help="Image locate timeout seconds",
    )
    parser.add_argument(
        "--image-confidence",
        type=float,
        default=0.9,
        help="Image locate confidence (0~1)",
    )
    parser.add_argument(
        "--image-click-count",
        type=int,
        default=1,
        help="How many times to click after image is found",
    )
    parser.add_argument(
        "--use-llm-planner",
        action="store_true",
        help="Use local LLM to generate execution plan",
    )
    parser.add_argument(
        "--planner-model",
        type=str,
        default="qwen2.5:latest",
        help="Local planner model name",
    )
    parser.add_argument(
        "--planner-endpoint",
        type=str,
        default="http://127.0.0.1:11434/api/generate",
        help="Local planner endpoint",
    )
    parser.add_argument(
        "--planner-timeout",
        type=int,
        default=90,
        help="LLM planner timeout in seconds",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    command = args.command.strip() if args.command else input("Enter command: ").strip()
    raise SystemExit(
        run(
            command,
            dry_run=args.dry_run,
            monitor_enabled=args.monitor,
            monitor_interval=args.monitor_interval,
            monitor_model=args.monitor_model,
            monitor_endpoint=args.monitor_endpoint,
            image_template=args.image_template,
            image_timeout=args.image_timeout,
            image_confidence=args.image_confidence,
            image_click_count=args.image_click_count,
            use_llm_planner=args.use_llm_planner,
            planner_model=args.planner_model,
            planner_endpoint=args.planner_endpoint,
            planner_timeout=args.planner_timeout,
        )
    )
