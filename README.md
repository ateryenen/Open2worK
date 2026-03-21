# AI Desktop Agent POC (Windows + Notepad)

## Purpose
This project is a local proof of concept that validates a deterministic desktop automation flow:

1. Receive user command from terminal
2. Build a rule-based action plan
3. Execute plan on Windows Notepad
4. Verify output file exists
5. Print step-by-step logs

## Version
- Current documented version: v0.2.0
- Updated at: 2026-03-17

## Scope
- Windows only
- Notepad only
- Rule-based planner
- Deterministic executor
- Local execution only

## Project Structure

```text
Open2worK/
├─ app/
│  ├─ __init__.py
│  ├─ main.py
│  ├─ planner.py
│  ├─ executor.py
│  ├─ schemas.py
│  ├─ config.py
│  └─ utils.py
├─ requirements.txt
├─ spec.md
└─ README.md
```

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python -m app.main
```

or non-interactive mode:

```powershell
python -m app.main --command "Open Notepad, type \"Hello Ater\", and save it to Desktop as test.txt"
```

planner-only dry-run:

```powershell
python -m app.main --command "Open Notepad, type hello, and save as test.txt" --dry-run
```

run with local LLM screen monitoring:

```powershell
python -m app.main --command "Open Notepad, type hello, and save as test.txt" --monitor
```

monitor strategy (hybrid):

```text
Python first analyzes screen/window state into structured JSON,
then LLM only interprets that JSON for risk hints.
```

custom monitor model / interval:

```powershell
python -m app.main --command "Open Notepad, type hello, and save as test.txt" --monitor --monitor-model gemma3:12b --monitor-interval 1.5
```

custom monitor endpoint:

```powershell
python -m app.main --command "Open Notepad, type hello, and save as test.txt" --monitor --monitor-endpoint http://127.0.0.1:11434/api/generate
```

image-based click POC (use a template screenshot):

```powershell
python -m app.main --command "image poc" --image-template "C:\\GITHUB\\Open2worK\\assets\\T357.png" --image-timeout 10 --image-confidence 0.9
```

This mode only runs image locate + click once, and reports success/failure in terminal.

local LLM planner -> image double-click:

```powershell
python -m app.main --command "請用圖片點擊桌面上的 T357 捷徑兩次" --use-llm-planner --planner-model phi:latest --planner-timeout 90
```

execution tree cache (for faster repeated image actions):

- image click now uses cache-first lookup, then falls back to full-screen scan.
- cache file: `C:/GITHUB/Open2worK/.cache/execution_tree.json`
- after successful click, object position is updated in cache for next run.
- if cached point misses repeatedly, cache entry is auto-evicted.

benchmark success rate:

```powershell
python -m app.benchmark --runs 10
```

Example command:

```text
Open Notepad, type "Hello Ater", and save it to Desktop as test.txt
```

## Expected Result
- Notepad opens
- Text is typed
- File is saved to `%USERPROFILE%\Desktop\test.txt`
- Terminal logs show planning, execution, and verification

## Notes
- If the command is partially unsupported, planner falls back to defaults.
- Fallback reasons are printed in terminal logs for debugging.
- Current implementation reserves a future interface for local LLM planner integration.
- Screen monitor uses local Ollama endpoint `http://127.0.0.1:11434/api/generate`.
- Default monitor model is `phi:latest`.
- Monitor does not require image-vision model now; it sends Python-generated JSON state to LLM.
- If monitor endpoint returns 404, code will auto fallback from `/api/generate` to `/api/chat` once.

## Current Automation Status (2026-03-17)
- Editor flow test now executes real actions via `/api/run-direct`.
- Step library currently supports:
	- `open_app`
	- `type_text`
	- `key_press`
	- `click_image`
	- `move_mouse_horizontal`
	- `mouse_click`
	- `save_file`
	- `wait`
- `mouse_click` and `click_image` support click mode switch:
	- `mouse_event` (legacy)
	- `send_input` (new)
- `click_image` and `mouse_click` support `press_duration`.
- Image matching uses fallback strategy (color/gray + reduced confidence) to reduce locate timeout.

### RDP / Remote Desktop Limitation
- When target app is running inside a remote desktop session, local injected input (`mouse_event`, `SendInput`, `PostMessage`) can be blocked or not forwarded by RDP client.
- This is a Windows session/desktop boundary behavior, not only a focus issue.
- If remote target still ignores click, prefer running automation in the same remote session or use app-native automation interfaces (e.g., COM/API) when available.
