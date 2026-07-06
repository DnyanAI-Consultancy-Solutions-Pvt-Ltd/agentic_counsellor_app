import os
import re
import time
import json
import pandas as pd
import pdfplumber
import autogen
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Optional
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Load local environment configurations safely into FastAPI process space
from dotenv import load_dotenv
load_dotenv()

app = FastAPI(title="MHT-CET & JEE Counseling Engine with Production PDF RAG")

@app.get("/")
def read_root():
    return {
        "status": "online",
        "message": "MHT-CET & JEE Counseling Engine API is running. Use the Streamlit interface to interact."
    }

# Lock down the absolute workspace directory where your files sit
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

class ProcessingConfig:
    PDF_PATH = os.path.join(ROOT_DIR, "2023ENGG_CAP1_CutOff.pdf")
    OUTPUT_DIR = os.path.join(ROOT_DIR, "outputs")

df_database = None

# --- Pydantic Schema Layers ---
class QueryPayload(BaseModel):
    student_percentile: float
    seat_category: str
    choice_stream: Optional[str] = "Any"
    is_pwd: Optional[bool] = False
    jee_percentile: Optional[float] = 0.0
    preferred_city: Optional[str] = "Any"
    selected_gender: Optional[str] = "General Pool"
    target_university: Optional[str] = "Any"

class ExcelCompilePayload(BaseModel):
    json_chosen_records: str
    student_percentile: float
    seat_category: str
    choice_stream: Optional[str] = "Any"
    jee_percentile: Optional[float] = 0.0

class ChatPayload(BaseModel):
    prompt: str
    ui_percentile: float
    ui_jee_percentile: float
    ui_category: str
    ui_stream: Optional[str] = "Any"
    ui_pwd: Optional[bool] = False
    ui_city: Optional[str] = "Any"
    ui_gender: Optional[str] = "General Pool"
    ui_university: Optional[str] = "Any"

# --- Structural PDF Extraction Engine ---
@app.on_event("startup")
def compile_database_on_boot():
    """Reads the root folder PDF dynamically on startup to populate the RAG framework."""
    global df_database
    print(f"Targeting PDF at: {ProcessingConfig.PDF_PATH}")
    
    if not os.path.exists(ProcessingConfig.PDF_PATH):
        print(f"CRITICAL: '{ProcessingConfig.PDF_PATH}' missing from root workspace.")
        return
        
    try:
        reader = pdfplumber.open(ProcessingConfig.PDF_PATH)
        structured_data = []
        current_college = "Unknown Institute"
        current_branch = "General"
        
        seat_token_pattern = re.compile(r'\b([GLPDRD][A-Z0-9]{3,8})\b')
        percentile_pattern = re.compile(r'\(\s*(\d{1,2}\.\d+)\s*\)')
        
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line_str = line.strip()
                if not line_str:
                    continue
                if re.search(r"^\d{4}\s*-", line_str):
                    current_college = line_str
                    continue
                if re.search(r"^\d{9}\s*-", line_str):
                    current_branch = line_str
                    continue
                if "GOPENS" in line_str or "GOPENH" in line_str or "LOPEN" in line_str:
                    seats = seat_token_pattern.findall(line_str)
                    if seats: active_headers = seats
                    continue
                    
                scores = percentile_pattern.findall(line_str)
                if scores and active_headers:
                    for idx, score_val in enumerate(scores):
                        if idx < len(active_headers):
                            structured_data.append({
                                "College": current_college,
                                "Course": current_branch,
                                "Seat Type": active_headers[idx],
                                "Cutoff Percentile": float(score_val)
                            })
                            
        df_database = pd.DataFrame(structured_data).drop_duplicates()
        print(f"RAG Layer Primed: Successfully indexed {len(df_database)} rows directly from the PDF.")
    except Exception as e:
        print(f"Failed to read data from PDF: {str(e)}")

# --- Core Matrix Retrieval Logic ---
def execute_matching_logic(payload: QueryPayload):
    global df_database
    if df_database is None or df_database.empty:
        return []

    cat = payload.seat_category.strip().upper()
    effective_score = max(payload.student_percentile, payload.jee_percentile)
    
    if payload.is_pwd:
        targets = [f"PWD{cat}S", f"PWD{cat}H", f"PWD{cat}", f"PWDR{cat}S"]
    else:
        if payload.selected_gender == "Female":
            targets = [f"L{cat}S", f"L{cat}H", f"G{cat}S", f"G{cat}H", cat]
        else:
            targets = [f"G{cat}S", f"G{cat}H", cat]

    base_filtered = df_database[df_database["Seat Type"].isin(targets)].copy()
    if base_filtered.empty:
        base_filtered = df_database[df_database["Seat Type"].str.contains(cat, case=False, na=False)].copy()
    if base_filtered.empty:
        return []

    if payload.preferred_city and payload.preferred_city.lower() != "any":
        base_filtered = base_filtered[base_filtered["College"].str.contains(payload.preferred_city.strip(), case=False, na=False)]

    if payload.target_university and payload.target_university.lower() != "any":
        base_filtered = base_filtered[base_filtered["College"].str.contains(payload.target_university.strip(), case=False, na=False)]

    if payload.choice_stream and payload.choice_stream.lower() != "any":
        clean_stream = payload.choice_stream.replace("/", " or ").replace("&", " or ")
        tokens = [t.strip() for t in re.split(r'\bor\b', clean_stream, flags=re.IGNORECASE) if t.strip()]
        expanded_tokens = []
        for tk in tokens:
            if tk.upper() in ["CS", "CSE", "COMPUTER"]:
                expanded_tokens.extend(["Computer Science", "Computer Engineering"])
            elif tk.upper() in ["IT", "INFO"]:
                expanded_tokens.append("Information Technology")
            else:
                expanded_tokens.append(tk)
        regex_q = "|".join([re.escape(e) for e in expanded_tokens])
        base_filtered = base_filtered[base_filtered["Course"].str.contains(regex_q, case=False, na=False)]

    # Dynamic pre-stratification mathematically bounds token counts to fit in generation boundaries
    dream_pool = base_filtered[(base_filtered["Cutoff Percentile"] > effective_score) & (base_filtered["Cutoff Percentile"] <= effective_score + 5.0)].sort_values(by="Cutoff Percentile", ascending=False).head(20)
    target_pool = base_filtered[(base_filtered["Cutoff Percentile"] >= effective_score - 1.5) & (base_filtered["Cutoff Percentile"] <= effective_score)].sort_values(by="Cutoff Percentile", ascending=False).head(25)
    safety_pool = base_filtered[(base_filtered["Cutoff Percentile"] >= effective_score - 5.0) & (base_filtered["Cutoff Percentile"] < effective_score - 1.5)].sort_values(by="Cutoff Percentile", ascending=False).head(20)
    
    stratified_df = pd.concat([dream_pool, target_pool, safety_pool]).drop_duplicates(subset=["College", "Course"])

    compacted_records = []
    for _, row in stratified_df.iterrows():
        compacted_records.append({
            "College": row["College"],
            "Course": row["Course"],
            "SeatType": row["Seat Type"],
            "CutoffPercentile": row["Cutoff Percentile"]
        })
    return compacted_records

def execute_excel_generation(payload: ExcelCompilePayload):
    try:
        data = json.loads(payload.json_chosen_records)
    except Exception:
        match = re.search(r"\[\s*\{.*\}\s*\]", payload.json_chosen_records, re.DOTALL)
        if match:
            data = json.loads(match.group(0))
        else:
            raise ValueError("Provided configuration string cannot be successfully loaded as valid JSON data.")

    if len(data) > 30:
        data = data[:30]
        
    reconstructed_rows = []
    for idx, item in enumerate(data):
        reconstructed_rows.append({
            "Preference No": idx + 1,
            "College": item.get("College", "Unknown Institution"),
            "Course": item.get("Course", "Unknown Course Branch"),
            "Seat Type": item.get("SeatType", item.get("Seat Type", "N/A")),
            "Cutoff Percentile": float(item.get("CutoffPercentile", item.get("Cutoff Percentile", 0.0))),
            "Reasoning Logic": item.get("Reasoning Logic", item.get("ReasoningLogic", "Portfolio Match"))
        })
        
    df = pd.DataFrame(reconstructed_rows)
    if not os.path.exists(ProcessingConfig.OUTPUT_DIR):
        os.makedirs(ProcessingConfig.OUTPUT_DIR)

    filename = f"CAP_Strict_{payload.seat_category.upper()}_{payload.student_percentile}_{int(time.time())}.xlsx"
    full_output_path = os.path.join(ProcessingConfig.OUTPUT_DIR, filename)

    writer = pd.ExcelWriter(full_output_path, engine='openpyxl')
    with writer:
        df.to_excel(writer, index=False, sheet_name="Preferences")
        worksheet = writer.sheets["Preferences"]
        worksheet.views.sheetView[0].showGridLines = True
        
        h_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        h_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        z_fill = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
        w_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))

        for c_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=c_idx)
            cell.fill = h_fill; cell.font = h_font; cell.border = border

        for r_idx in range(2, len(df) + 2):
            c_fill = z_fill if r_idx % 2 == 0 else w_fill
            for c_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=r_idx, column=c_idx)
                cell.fill = c_fill; cell.border = border
                col_name = df.columns[c_idx - 1]
                if col_name in ["Preference No", "Seat Type"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "Cutoff Percentile":
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col in worksheet.columns:
            letter = get_column_letter(col[0].column)
            max_len = max(len(str(c.value or '')) for c in col)
            worksheet.column_dimensions[letter].width = min(max(max_len + 3, 12), 75)

    return {"filepath": full_output_path, "preview_data": df.to_dict(orient="records")}

# --- REST App Interface Routes ---
@app.post("/api/v1/match-colleges")
def get_matching_colleges(payload: QueryPayload):
    return execute_matching_logic(payload)

@app.post("/api/v1/generate-excel")
def generate_excel_report(payload: ExcelCompilePayload):
    try:
        return execute_excel_generation(payload)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/v1/chat")
def process_counselor_chat(payload: ChatPayload):
    """Executes the standard LLM instruction set over PDF extracted data."""
    if df_database is None or df_database.empty:
        raise HTTPException(status_code=503, detail="RAG Engine Data Context is empty. Verify root PDF file is present.")

    groq_key = os.environ.get("GROQ_API_KEY", "missing_key")
    config_list = [{'model': 'llama-3.3-70b-versatile', 'api_key': groq_key, 'api_type': 'openai', 'base_url': 'https://api.groq.com/openai/v1'}]
    
    q_payload = QueryPayload(
        student_percentile=payload.ui_percentile, 
        seat_category=payload.ui_category,
        choice_stream=payload.ui_stream, 
        is_pwd=payload.ui_pwd, 
        jee_percentile=payload.ui_jee_percentile,
        preferred_city=payload.ui_city, 
        selected_gender=payload.ui_gender, 
        target_university=payload.ui_university
    )
    raw_matches = execute_matching_logic(q_payload)
    raw_csv_context = pd.DataFrame(raw_matches).to_csv(index=False)
    
    # Standardized Normal Instructions for the Counselor Agent[cite: 2]
    standard_instructions = f"""You are an helpful, expert AI College Admission Counselor for Maharashtra engineering colleges.
Your goal is to help students generate a balanced list of up to 30 college preferences based on their score of {payload.ui_percentile}%.[cite: 2]

Here is the data found in our database matching their profile:
{raw_csv_context}

INSTRUCTIONS:
1. Review the options above and pick up to 30 of the best recommendations.[cite: 2]
2. Categorize your choices based on their score:
   - "Dream Options": Cutoffs higher than the student's score.[cite: 2]
   - "Target Options": Cutoffs right around the student's score.[cite: 2]
   - "Safety Options": Cutoffs lower than the student's score acting as a safe backup.[cite: 2]
3. Sort the list from the highest cutoff percentile down to the lowest.[cite: 2]
4. For each selected choice, write a brief 'Reasoning Logic' explaining why it's a good fit.[cite: 2]

STEPS TO EXECUTE:
1. Format your selected 30 rows as a valid JSON array.[cite: 2]
2. Call the tool 'internal_compile_excel' with this JSON string to automatically create the downloadable sheet on the user interface.[cite: 2]
3. After the tool executes successfully, provide a friendly summary breakdown of your advice to the student and append 'TERMINATE' to end your thought process."""

    agent_llm_config = {
        "config_list": config_list, 
        "timeout": 120, 
        "temperature": 0.2
    }

    counsellor = autogen.AssistantAgent(
        name="Counsellor_Brain_Agent", 
        llm_config=agent_llm_config,
        system_message=standard_instructions
    )
    
    user_proxy = autogen.UserProxyAgent(
        name="User_Proxy", 
        human_input_mode="NEVER", 
        max_consecutive_auto_reply=3, 
        is_termination_msg=lambda x: "TERMINATE" in (x.get("content") or ""),
        code_execution_config={"use_docker": False}
    )
    
    execution_context = {"filepath": None, "preview_data": None}

    def internal_compile_excel(json_chosen_records: str, student_percentile: float, seat_category: str, choice_stream: str = "Any", jee_percentile: float = 0.0):
        p = ExcelCompilePayload(
            json_chosen_records=json_chosen_records, 
            student_percentile=student_percentile, 
            seat_category=seat_category, 
            choice_stream=choice_stream, 
            jee_percentile=jee_percentile
        )
        res = execute_excel_generation(p)
        execution_context["filepath"] = res["filepath"]
        execution_context["preview_data"] = res["preview_data"]
        return f"Success! Excel document compiled at: {res['filepath']}"
        
    autogen.register_function(
        internal_compile_excel,
        caller=counsellor,
        executor=user_proxy,
        name="internal_compile_excel",
        description="Compiles curated strategic records into an Excel sheet layout structure."
    )

    user_proxy.initiate_chat(counsellor, message=payload.prompt, silent=True)
    
    final_reply = counsellor.last_message()["content"].replace("TERMINATE", "").strip()

    return {
        "reply": "Preference sheet generated successfully!" if not final_reply else final_reply, 
        "filepath": execution_context["filepath"], 
        "preview_data": execution_context["preview_data"]
    }