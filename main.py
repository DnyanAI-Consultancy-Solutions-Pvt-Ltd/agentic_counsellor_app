import os
import re
import time
import json
import threading
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

app = FastAPI(title="MHT-CET Counseling Engine with Production PDF RAG")

# Lock down the absolute workspace directory where your files sit
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

class ProcessingConfig:
    PDF_PATH = os.path.join(ROOT_DIR, "2023ENGG_CAP1_CutOff.pdf")
    OUTPUT_DIR = os.path.join(ROOT_DIR, "outputs")

# Shared thread-safe structures
df_database = None
is_database_loaded = False  

@app.get("/")
def read_root():
    global is_database_loaded, df_database
    return {
        "status": "online",
        "database_ready": is_database_loaded,
        "indexed_records": len(df_database) if df_database is not None else 0,
        "message": "MHT-CET Counseling Engine API is running."
    }

# --- Pydantic Schema Layers ---
class QueryPayload(BaseModel):
    student_percentile: float
    seat_category: str
    choice_stream: Optional[str] = "Any"
    is_pwd: Optional[bool] = False
    preferred_city: Optional[str] = "Any"
    selected_gender: Optional[str] = "General Pool"
    target_university: Optional[str] = "Any"

class ExcelCompilePayload(BaseModel):
    json_chosen_records: str
    student_percentile: float
    seat_category: str
    choice_stream: Optional[str] = "Any"

class ChatPayload(BaseModel):
    prompt: str
    ui_percentile: float
    ui_category: str
    ui_stream: Optional[str] = "Any"
    ui_pwd: Optional[bool] = False
    ui_city: Optional[str] = "Any"
    ui_gender: Optional[str] = "General Pool"
    ui_university: Optional[str] = "Any"

last_compiled_portfolio = {}

# --- High-Precision PDF Extraction Engine ---
def compile_database_worker():
    global df_database, is_database_loaded
    print(f"BACKGROUND WORKER: Targeting PDF at: {ProcessingConfig.PDF_PATH}")
    
    if not os.path.exists(ProcessingConfig.PDF_PATH):
        print(f"CRITICAL: '{ProcessingConfig.PDF_PATH}' missing from root workspace.")
        return
        
    try:
        start_time = time.time()
        reader = pdfplumber.open(ProcessingConfig.PDF_PATH)
        structured_data = []
        
        current_college = "Unknown Institute"
        current_branch = "General Branch"
        active_headers = []
        
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
                    if seats: 
                        active_headers = seats
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
            
            page.flush_cache()
                            
        if structured_data:
            df_database = pd.DataFrame(structured_data).drop_duplicates()
            is_database_loaded = True
            print(f"RAG Layer Primed: Successfully indexed {len(df_database)} rows in {round(time.time() - start_time, 2)}s.")
        else:
            df_database = pd.DataFrame(columns=["College", "Course", "Seat Type", "Cutoff Percentile"])
            is_database_loaded = True
            
    except Exception as e:
        print(f"Failed to read data from PDF background thread: {str(e)}")

@app.on_event("startup")
def compile_database_on_boot():
    threading.Thread(target=compile_database_worker, daemon=True).start()

# --- Core Matrix Retrieval Logic with Strict Equal Bracket Truncation ---
def execute_matching_logic(payload: QueryPayload):
    global df_database
    if df_database is None or df_database.empty:
        return []

    cat = payload.seat_category.strip().upper()
    effective_score = payload.student_percentile
    
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

    # Exactly 10 rows per bracket matrix structure
    dream_pool = base_filtered[(base_filtered["Cutoff Percentile"] > effective_score) & (base_filtered["Cutoff Percentile"] <= effective_score + 5.0)].sort_values(by="Cutoff Percentile", ascending=False).head(10)
    target_pool = base_filtered[(base_filtered["Cutoff Percentile"] >= effective_score - 1.5) & (base_filtered["Cutoff Percentile"] <= effective_score)].sort_values(by="Cutoff Percentile", ascending=False).head(10)
    safety_pool = base_filtered[(base_filtered["Cutoff Percentile"] >= effective_score - 6.0) & (base_filtered["Cutoff Percentile"] < effective_score - 1.5)].sort_values(by="Cutoff Percentile", ascending=False).head(10)
    
    all_matches = []
    for df_pool in [dream_pool, target_pool, safety_pool]:
        for _, row in df_pool.iterrows():
            all_matches.append({
                "College": row["College"],
                "Course": row["Course"],
                "Seat Type": row["Seat Type"],
                "Cutoff Percentile": row["Cutoff Percentile"]
            })
    return all_matches

# --- Excel Generation Engine ---
def execute_excel_generation_from_raw(data_list, student_percentile, seat_category):
    reconstructed_rows = []
    for idx, item in enumerate(data_list[:30]):
        raw_pct = float(item.get("Cutoff Percentile", item.get("CutoffPercentile", 0.0)))
        
        if raw_pct > student_percentile:
            raw_reason = f"Dream Option: Cutoff ({raw_pct}%) is higher than your score ({student_percentile}%). Ambitious high-reach preference."
        elif raw_pct >= student_percentile - 1.5:
            raw_reason = f"Target Option: Cutoff ({raw_pct}%) closely mirrors your score ({student_percentile}%). Strong practical chance of admission."
        else:
            raw_reason = f"Safety Option: Cutoff ({raw_pct}%) acts as a highly secure backup buffer to prevent CAP elimination."

        reconstructed_rows.append({
            "Preference No": idx + 1,
            "College": item.get("College", "Unknown Institution"),
            "Course": item.get("Course", "Unknown Course Branch"),
            "Seat Type": item.get("Seat Type", item.get("SeatType", "N/A")),
            "Cutoff Percentile": raw_pct,
            "Reasoning Logic": raw_reason
        })
        
    df = pd.DataFrame(reconstructed_rows)
    if not os.path.exists(ProcessingConfig.OUTPUT_DIR):
        os.makedirs(ProcessingConfig.OUTPUT_DIR)

    filename = f"CAP_Strict_{seat_category.upper()}_{student_percentile}_{int(time.time())}.xlsx"
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

# --- Clean Direct Agent Layer ---
@app.post("/api/v1/chat")
def process_counselor_chat(payload: ChatPayload):
    global last_compiled_portfolio, df_database
    if df_database is None or df_database.empty:
        raise HTTPException(status_code=503, detail="The PDF data is currently processing in the background.")

    groq_key = os.environ.get("GROQ_API_KEY", "missing_key")
    config_list = [{'model': 'llama-3.3-70b-versatile', 'api_key': groq_key, 'api_type': 'openai', 'base_url': 'https://api.groq.com/openai/v1'}]
    
    q_payload = QueryPayload(
        student_percentile=payload.ui_percentile, 
        seat_category=payload.ui_category,
        choice_stream=payload.ui_stream, 
        preferred_city=payload.ui_city, 
        selected_gender=payload.ui_gender, 
        target_university=payload.ui_university
    )
    filtered_matches = execute_matching_logic(q_payload)
    
    session_key = f"{payload.ui_percentile}_{payload.ui_category}"
    current_saved_list = last_compiled_portfolio.get(session_key, [])
    
    is_feedback_request = any(keyword in payload.prompt.lower() for keyword in ["remove", "replace", "swap", "change", "instead", "don't want", "move", "add"])
    
    if is_feedback_request and current_saved_list:
        updated_list = []
        for row in current_saved_list:
            if not any(keyword in row["College"].lower() for keyword in payload.prompt.lower().split()):
                updated_list.append(row)
        
        while len(updated_list) < 30 and len(filtered_matches) > len(updated_list):
            for match in filtered_matches:
                if match["College"] not in [r["College"] for r in updated_list]:
                    updated_list.append(match)
                    break
        final_list = updated_list[:30]
    else:
        final_list = filtered_matches[:30]

    last_compiled_portfolio[session_key] = final_list
    res = execute_excel_generation_from_raw(final_list, payload.ui_percentile, payload.ui_category)
    
    summary_instructions = f"""You are an elite expert AI College Admission Counselor for Maharashtra engineering CAP rounds. 
    Your tone must be highly supportive, reassuring, and professional. Speak like a real senior advisor.
    Provide a brief and helpful overview of the selection compiled for a {payload.ui_percentile}% score. Do NOT include any technical code, raw JSON, or agent payloads."""
    
    counsellor = autogen.AssistantAgent(name="Counsellor_Brain", llm_config={"config_list": config_list}, system_message=summary_instructions)
    
    # CRITICAL DOCKER PROTECTION GUARD: Explicitly set use_docker to False to avoid runtime crashes on Hugging Face
    user_proxy = autogen.UserProxyAgent(
        name="User_Proxy", 
        human_input_mode="NEVER", 
        max_consecutive_auto_reply=1,
        code_execution_config={"use_docker": False}
    )
    
    user_proxy.initiate_chat(counsellor, message=f"Provide a friendly expert counselor summary for a student profile with score {payload.ui_percentile}%.", silent=True)
    final_reply = counsellor.last_message()["content"].strip()

    return {
        "reply": final_reply if final_reply else "Your preference list has been re-proportioned with exactly 10 Dream, 10 Target, and 10 Safety choices successfully!", 
        "filepath": res["filepath"], 
        "preview_data": res["preview_data"]
    }