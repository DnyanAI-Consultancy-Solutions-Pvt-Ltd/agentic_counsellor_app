import os
import re
import pandas as pd
import pdfplumber
import autogen
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# 1. STATEFUL ADAPTIVE DATABASE PARSER
# =====================================================================
def parse_cutoff_pdf(pdf_path):
    """
    Parses layout rows from the CAP PDF. Matches row category headers 
    and links them to respective percentiles.
    """
    structured_data = []
    current_college = "Unknown Institute"
    current_branch = "General"
    active_seat_headers = []
    
    print(f"Reading and compiling real records from '{pdf_path}'...")
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split("\n")
            for line in lines:
                cleaned_line = line.strip()
                
                if re.search(r"^\d{4}\s*-", cleaned_line):
                    current_college = cleaned_line
                    continue
                
                if re.search(r"^\d{9}\s*-", cleaned_line):
                    current_branch = cleaned_line
                    continue
                
                if "GOPENS" in cleaned_line or "GOPENH" in cleaned_line or "Stage" in line:
                    found_seats = re.findall(r"\b([A-Z0-9]{3,9})\b", cleaned_line)
                    if found_seats:
                        active_seat_headers = found_seats
                    continue
                
                percentiles = re.findall(r"\(\s*(\d{1,2}\.\d+)\s*\)", cleaned_line)
                if percentiles and active_seat_headers:
                    for idx, score in enumerate(percentiles):
                        seat = active_seat_headers[idx] if idx < len(active_seat_headers) else "GOPENS"
                        structured_data.append({
                            "College": current_college,
                            "Course": current_branch,
                            "Seat Type": seat,
                            "Cutoff Percentile": float(score)
                        })
                        
    df = pd.DataFrame(structured_data).drop_duplicates()
    print(f"Successfully loaded {len(df)} live data records directly from your document.\n")
    return df

# Initialize global database from your source PDF document
pdf_filename = "2023ENGG_CAP1_CutOff.pdf"
if not os.path.exists(pdf_filename):
    raise FileNotFoundError(f"Critical Error: Place your source document '{pdf_filename}' in this folder.")

df_database = parse_cutoff_pdf(pdf_filename)


# =====================================================================
# 2. THE DEDUPLICATED RETRIEVAL & HIGH-FORMATTING EXCEL ENGINE
# =====================================================================
def find_and_export_colleges(student_percentile: float, seat_category: str, choice_stream: str = None, is_pwd: bool = False, output_excel: str = "CAP_Round_Allotment_Preferences.xlsx") -> str:
    """
    Filters the parsed database matching category, PWD profiles, and stream keywords.
    Removes duplicated college options to deliver a clean preference list.
    """
    global df_database
    seat_category = seat_category.strip().upper()
    
    # 1. STRICT CATEGORY LOCKING: Match exact requirements
    if is_pwd:
        search_seat_patterns = [f"PWD{seat_category}S", f"PWD{seat_category}H", f"PWD{seat_category}", f"PWDR{seat_category}S", "PWDOPENS"]
    else:
        search_seat_patterns = [f"G{seat_category}S", f"G{seat_category}H", f"L{seat_category}S", f"L{seat_category}H", seat_category]

    matched_df = df_database[df_database["Seat Type"].isin(search_seat_patterns)].copy()
    
    # Fallback to safety category variations only if the strict set yields no matches
    if matched_df.empty:
        matched_df = df_database[df_database["Seat Type"].str.contains(seat_category, case=False, na=False)].copy()
    if matched_df.empty:
        matched_df = df_database[df_database["Seat Type"].str.contains("OPEN", case=False, na=False)].copy()
        
    # 2. STREAM FILTER: Resolves combined parameters like "IT or CS"
    if choice_stream and choice_stream.lower() != "any":
        clean_stream = choice_stream.replace("/", " or ").replace("&", " or ")
        keywords = [k.strip() for k in re.split(r'\bor\b', clean_stream, flags=re.IGNORECASE) if k.strip()]
        
        expanded_keywords = []
        for kw in keywords:
            if kw.upper() in ["CS", "CSE", "COMPUTER"]:
                expanded_keywords.extend(["Computer Science", "Computer Engineering"])
            elif kw.upper() in ["IT"]:
                expanded_keywords.append("Information Technology")
            else:
                expanded_keywords.append(kw)
        
        regex_query = "|".join([re.escape(kw) for kw in expanded_keywords])
        matched_df = matched_df[matched_df["Course"].str.contains(regex_query, case=False, na=False)]
        
    # 3. SCORE MATCHING & TARGET CUTOFF BUFFER
    eligible_df = matched_df[matched_df["Cutoff Percentile"] <= (student_percentile + 2.0)]
    
    # 4. STRICT DEDUPLICATION: Remove repeat College+Course entries, keeping the highest matching cutoff
    eligible_df = eligible_df.sort_values(by="Cutoff Percentile", ascending=False)
    eligible_df = eligible_df.drop_duplicates(subset=["College", "Course"], keep="first")
    
    # 5. ASSIGN FINAL SELECTIONS (Ensures at least 20 entries)
    final_list = eligible_df.head(25)
    if len(final_list) < 20:
        # If strict options under student score are few, draw slightly higher safe options to fill out the 20 rows
        final_list = matched_df.sort_values(by="Cutoff Percentile", ascending=False).drop_duplicates(subset=["College", "Course"]).head(22)

    # Re-index preferences smoothly
    final_list = final_list.copy().sort_values(by="Cutoff Percentile", ascending=False)
    final_list.insert(0, "Preference No", range(1, len(final_list) + 1))
    
    # Generate Custom Reasoning Logic
    reasoning_log = []
    for idx, row in final_list.iterrows():
        cutoff = row["Cutoff Percentile"]
        diff = student_percentile - cutoff
        if diff < 0:
            reasoning_log.append(f"Aggressive/Dream Option: Historical cutoff is slightly above your score by {abs(diff):.4f}%. Higher tier choice.")
        elif diff <= 2.5:
            reasoning_log.append(f"Highly Balanced Target: Your score clears the historical cutoff by a solid {diff:.4f}%. Solid option.")
        else:
            reasoning_log.append(f"Secure Safety Backup: Comfortable percentile buffer of {diff:.4f}%, providing high admission certainty.")
            
    final_list["Reasoning Logic"] = reasoning_log
    final_list = final_list[["Preference No", "College", "Course", "Seat Type", "Cutoff Percentile", "Reasoning Logic"]]

    # Prevent write crashes if the file is currently open in Excel
    try:
        writer = pd.ExcelWriter(output_excel, engine='openpyxl')
    except PermissionError:
        import time
        output_excel = f"CAP_Round_Preferences_{int(time.time())}.xlsx"
        writer = pd.ExcelWriter(output_excel, engine='openpyxl')

    final_list.to_excel(writer, index=False, sheet_name="Preferences")
    
    workbook = writer.book
    worksheet = writer.sheets["Preferences"]
    worksheet.views.sheetView[0].showGridLines = True
    
    # Executive Formatting Themes (Navy Blue Corporate Accents)
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    data_font = Font(name="Segoe UI", size=10, color="333333")
    bold_data_font = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
    thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))

    for col_idx in range(1, len(final_list.columns) + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    worksheet.row_dimensions[1].height = 28

    for row_idx in range(2, len(final_list) + 2):
        current_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, len(final_list.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.fill = current_fill
            cell.font = data_font
            cell.border = thin_border
            
            col_name = final_list.columns[col_idx - 1]
            if col_name in ["Preference No", "Seat Type"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_name == "Preference No":
                    cell.font = bold_data_font
            elif col_name == "Cutoff Percentile":
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '0.0000000'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        worksheet.row_dimensions[row_idx].height = 24

    for col in worksheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 75)

    writer.close()
    return f"Successfully generated your deduplicated premium choice report in Excel with {len(final_list)} entries saved to '{output_excel}'."


# =====================================================================
# 3. AUTOGEN AGENT CONFIGURATION (WITH CAPACITY FAILOVER REDUNDANCY)
# =====================================================================
groq_api_key = os.environ.get("GROQ_API_KEY", "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")

config_list = [
    {
        'model': 'llama-3.3-70b-versatile',
        'api_key': groq_api_key,
        'api_type': 'openai',
        'base_url': 'https://api.groq.com/openai/v1'
    },
    {
        'model': 'mixtral-8x7b-32768',
        'api_key': groq_api_key,
        'api_type': 'openai',
        'base_url': 'https://api.groq.com/openai/v1'
    }
]

llm_config = {
    "config_list": config_list,
    "timeout": 60,
    "temperature": 0.0,
    "functions": [
        {
            "name": "find_and_export_colleges",
            "description": "Queries the parsed cutoff database using parameters and saves uniquely deduplicated options into an attractive, styled Excel preference sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "student_percentile": {
                        "type": "number",
                        "description": "The exact MHT-CET percentile scored by the applicant (e.g. 94.2)."
                    },
                    "seat_category": {
                        "type": "string",
                        "description": "Caste category designation string like OPEN, OBC, SC, ST, EWS."
                    },
                    "choice_stream": {
                        "type": "string",
                        "description": "Target fields, fields phrases or abbreviations like 'IT or CS'."
                    },
                    "is_pwd": {
                        "type": "boolean",
                        "description": "Set to True if user explicitly states they belong to the PWD (Persons with Disabilities) category."
                    }
                },
                "required": ["student_percentile", "seat_category", "is_pwd"]
            }
        }
    ]
}

counsellor = autogen.AssistantAgent(
    name="Counsellor_Agent",
    llm_config=llm_config,
    system_message="""You are an expert engineering admissions counsellor agent for the CAP rounds.
Your job is to parse the user's natural query statement, extract their percentile, category code, preferred branches, and check for PWD status.
You must execute the 'find_and_export_colleges' tool immediately with these parameters to generate a professionally formatted Excel spreadsheet.
Once the spreadsheet tool confirms successful creation, notify the user that their premium spreadsheet report is generated and conclude your final response with 'TERMINATE'."""
)

user_proxy = autogen.UserProxyAgent(
    name="User_Proxy",
    human_input_mode="NEVER",
    max_consecutive_auto_reply=5, 
    is_termination_msg=lambda x: "TERMINATE" in (x.get("content") or ""),
    code_execution_config={"work_dir": "counsellor_workspace", "use_docker": False}
)

user_proxy.register_function(
    function_map={
        "find_and_export_colleges": find_and_export_colleges
    }
)

if __name__ == "__main__":
    print("\n=======================================================")
    print("  Welcome to the AutoGen CAP Admission Counsellor AI   ")
    print("=======================================================\n")
    
    user_query = input("Ask me anything about your admission (e.g., 'I scored 94.2 percentile, category is OBC, I am a PWD student and want IT or CS branches'):\n> ")
    
    print("\nProcessing request via AutoGen & Groq LLM optimization...")
    user_proxy.initiate_chat(
        counsellor,
        message=user_query
    )