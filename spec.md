\*\*在 VS Code 上做出一個可驗證的本地桌面自動化 POC\*\*



\* Windows only

\* Notepad 為目標 App

\* 規則式 planner 為主

\* 預留未來接本地 LLM 的接口

\* 先不做完整 UI tree，只保留後續擴充欄位



````markdown

\# POC SPEC.md

AI Desktop Agent - Proof of Concept



Version: v0.1  

Author: Ater Chen  

Status: POC Draft



---



\## 1. Purpose



This POC aims to validate the feasibility of a \*\*local desktop automation agent\*\* that can interpret a simple user instruction and perform deterministic actions on a Windows desktop application.



The initial objective is to prove that the following flow works end-to-end:



1\. Receive a natural language command

2\. Convert it into a structured action plan

3\. Execute the plan on Windows

4\. Complete the task successfully



This POC is intentionally limited in scope to reduce complexity and maximize validation speed.



---



\## 2. POC Goal



The primary goal of this POC is to automate a simple Notepad workflow on Windows.



Target scenario:



\- Open Notepad

\- Type text

\- Save the file to Desktop



Example user instruction:



Open Notepad, type "Hello Ater", and save it to Desktop as test.txt



---



\## 3. Scope



\### In Scope



This POC includes:



\- Windows desktop environment only

\- VS Code local development

\- Python-based implementation

\- Rule-based planner

\- Structured action schema

\- Deterministic executor

\- Basic task logging

\- Basic execution verification

\- Notepad automation only



\### Out of Scope



This POC does not include:



\- macOS support

\- full Local LLM integration

\- OCR

\- image-based UI tree generation

\- remote desktop support

\- advanced recovery engine

\- self-learning workflows

\- multi-application orchestration



---



\## 4. Success Criteria



The POC is considered successful if it can reliably perform the following:



1\. Launch Notepad

2\. Input predefined text

3\. Save the file to Desktop

4\. Confirm the file exists after saving

5\. Print execution logs in terminal



Minimum validation target:



\- successful execution rate >= 80% in local testing



---



\## 5. System Overview



The POC system consists of three core parts:



```text

User Input

&nbsp; ↓

Planner

&nbsp; ↓

Executor

&nbsp; ↓

Windows Notepad

````



Supporting components:



```text

Schema

Logging

Basic Verification

```



---



\## 6. Core Architecture



\### 6.1 Planner



The planner converts a user instruction into a structured execution plan.



For this POC, the planner is rule-based.



Responsibilities:



\* parse simple instruction patterns

\* generate action list

\* produce structured output



Example output:



```json

{

&nbsp; "goal": "Open Notepad and save text",

&nbsp; "steps": \[

&nbsp;   { "action": "open\_app", "target": "notepad" },

&nbsp;   { "action": "wait", "seconds": 1.5 },

&nbsp;   { "action": "type\_text", "text": "Hello Ater" },

&nbsp;   { "action": "wait", "seconds": 0.5 },

&nbsp;   { "action": "save\_file", "path": "%USERPROFILE%\\\\Desktop\\\\test.txt" }

&nbsp; ]

}

```



---



\### 6.2 Executor



The executor performs deterministic actions on the system.



Responsibilities:



\* open target application

\* send keyboard input

\* trigger save action

\* enter file path

\* finalize save flow



Execution must be deterministic and must not rely on model inference.



---



\### 6.3 Verification



Verification ensures the expected result exists after execution.



For this POC, verification includes:



\* checking whether Notepad was launched

\* checking whether the output file exists on Desktop



If verification fails, the system reports failure in terminal logs.



---



\### 6.4 Logging



The system must print execution logs to terminal.



Log examples:



\* received user input

\* generated plan

\* executing step 1

\* executing step 2

\* verification success/failure



Persistent file logging is not required in this POC.



---



\## 7. Functional Requirements



\### FR-001 User Input



The system shall accept a text command from terminal input.



Example:



```text

Open Notepad, type "Hello Ater", and save it to Desktop as test.txt

```



---



\### FR-002 Rule-Based Planning



The system shall convert supported user instructions into a structured action plan.



Supported intent in POC:



\* open Notepad

\* type text

\* save file



If the instruction is unsupported, the planner may fall back to a default test plan.



---



\### FR-003 Action Schema



Each executable step shall follow a normalized action schema.



Supported action types in POC:



\* `open\_app`

\* `type\_text`

\* `save\_file`

\* `wait`



Example schema:



```json

{

&nbsp; "action": "type\_text",

&nbsp; "text": "Hello Ater"

}

```



---



\### FR-004 Open Application



The executor shall support launching Notepad.



Supported target:



\* `notepad`



Any other app target is out of scope for this POC.



---



\### FR-005 Type Text



The executor shall support typing text into the active Notepad window.



The text may be:



\* predefined

\* parsed from user input

\* fallback test text



---



\### FR-006 Save File



The executor shall support saving content to a file path.



For this POC, the target path shall be limited to Desktop output.



Example:



```text

%USERPROFILE%\\Desktop\\test.txt

```



---



\### FR-007 Execution Verification



After the save operation, the system shall verify whether the file exists at the expected location.



If the file does not exist, execution shall be marked as failed.



---



\### FR-008 Terminal Logging



The system shall print step-by-step logs to terminal.



Minimum required logs:



\* input received

\* plan generated

\* action execution start

\* action execution result

\* verification result



---



\## 8. Non-Functional Requirements



\### NFR-001 Simplicity



The POC should remain small, easy to run, and easy to debug.



\### NFR-002 Local Execution



The entire POC shall run locally on the developer machine.



\### NFR-003 Deterministic Behavior



The executor must behave deterministically.



\### NFR-004 Fast Validation



The POC should be runnable in a local VS Code environment with minimal setup.



\### NFR-005 Extensibility



Although the POC is rule-based, the structure should allow future replacement of the planner with a Local LLM.



---



\## 9. Technical Stack



Recommended stack for the POC:



\### Language



\* Python 3.x



\### Development Environment



\* VS Code



\### Libraries



\* `pywinauto`

\* `pydantic`

\* `pyautogui` (fallback only)

\* `python-dotenv` (optional)



\### Future-compatible libraries



\* `ollama`

\* `pytesseract`

\* `Pillow`



---



\## 10. Project Structure



Recommended project structure:



```text

desktop-agent-poc/

├─ app/

│  ├─ main.py

│  ├─ planner.py

│  ├─ executor.py

│  ├─ schemas.py

│  ├─ config.py

│  └─ utils.py

├─ requirements.txt

├─ README.md

└─ SPEC.md

```



---



\## 11. Execution Flow



POC execution flow:



```text

1\. User enters command

2\. Planner builds plan

3\. Executor runs each action

4\. Verification checks result

5\. Terminal prints success/failure

```



Detailed loop:



```text

User Input

&nbsp; ↓

Build Plan

&nbsp; ↓

Execute Step 1

&nbsp; ↓

Execute Step 2

&nbsp; ↓

Execute Step 3

&nbsp; ↓

Verify Output

&nbsp; ↓

Done

```



---



\## 12. Limitations



Known limitations of this POC:



\* only supports Windows

\* only supports Notepad

\* planner is rule-based

\* no OCR

\* no UI tree

\* no screenshot understanding

\* no advanced retry logic

\* no permission/safety engine

\* limited natural language flexibility



These limitations are acceptable because this phase focuses on proving the basic architecture.



---



\## 13. Future Upgrade Path



After this POC succeeds, the next upgrades should be:



\### Phase 2



\* replace rule-based planner with Local LLM planner

\* support structured JSON plan generation



\### Phase 3



\* add screenshot capture

\* add OCR

\* add image-derived pseudo UI tree



\### Phase 4



\* add verification engine

\* add retry and recovery logic

\* add action confidence handling



\### Phase 5



\* expand beyond Notepad to other desktop applications



---



\## 14. Design Principles



This POC follows these principles:



1\. Keep the first version narrow

2\. Prefer deterministic execution over intelligence

3\. Validate architecture before scaling complexity

4\. Separate planning from execution

5\. Make future Local LLM integration easy



Core idea:



```text

Planner decides what to do

Executor decides how to do it

Verification checks whether it worked

```



---



\## 15. Exit Conditions



This POC phase can be considered complete when:



\* project runs successfully in VS Code

\* Notepad automation works end-to-end

\* file save verification works

\* logs are visible and understandable

\* code structure is ready for Local LLM integration



---



\# End of POC Specification



```

