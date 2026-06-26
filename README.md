# AutoGen CAP Admission Counsellor AI

An intelligent agentic system designed to parse institutional engineering cut-off database records from a CAP Round PDF and cross-evaluate student records against admissions parameters via a stateful Multi-Agent structure using AutoGen and Groq.

---

## 1. System Requirements

- **Python Version**: Python `3.10` or `3.11` is highly recommended.
- **Operating System**: Platform independent (Windows, macOS, or Linux).
- **Core Dependencies**: `autogen`, `pandas`, `openpyxl`, and `pdfplumber`.

---

## 2. Project Directory Setup

Ensure that your local directory follows this structure precisely. The raw cut-off dataset PDF file must be placed directly inside the main project directory folder.

```text
Counselor/
├── agentic_counsellor_app/
│   └── app.py                  # Main AutoGen application script
├── 2023ENGG_CAP1_CutOff.pdf     # Source cut-off document (Required)
├── requirements.txt            # Streamlined package requirements
└── .gitignore                  # Git tracking rules file

**3. Local Installation & Deployment**
Step 3.1: Initialize an Isolated Workspace
Open your target shell prompt at the project root directory and spin up a local Python environment instance:
# On Windows (PowerShell)
python -m venv env
.\env\Scripts\Activate.ps1

# On macOS / Linux
python3 -m venv env
source env/bin/activate

Step 3.2: Pull Package Dependencies
Run the installation command to populate your virtual environment with structural data components:

pip install -r requirements.txt

4. Environment Variables Configuration
The pipeline requires a secure API token path to hook into the backend LLM processing architecture. Ensure this is configured locally prior to launching execution commands:

# On Windows (PowerShell)
$env:GROQ_API_KEY="your_actual_groq_api_key_here"

# On macOS / Linux
export GROQ_API_KEY="your_actual_groq_api_key_here"

5. Execution
Trigger the intake loop application by calling the main execution file pathway:

python ./agentic_counsellor_app/app.py

Interacting with the Portal Console
When prompted by the console cursor (>), declare the metrics profile you wish to cross-reference:

Ask me anything about your admission:
> I scored 94.2 percentile, category is OBC, and I'm looking for CS or IT branches.

6. Output Generation
Once the Counsellor_Agent and User_Proxy confirm data mapping and complete execution routines, a styled premium report sheet will appear in your project root workspace:

Filename: CAP_Round_Allotment_Preferences.xlsx

Features: Distinct visual grid lines, corporate navy corporate headers, exact float precision mappings, auto-fitted margins, and step-by-step reasoning logic logs for up to 25 target options.

