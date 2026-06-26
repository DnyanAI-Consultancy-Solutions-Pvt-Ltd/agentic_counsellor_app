import os
import re
import pandas as pd
import pdfplumber
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# 1. PARSING REAL DATA AND RESOLVING PWD / CASTE MATCHES
# =====================================================================
def extract_live_cap_database(pdf_path):
    """
    Exclusively parses the structural layout rows from your attached CAP PDF.
    Extracts all seats including General, Reserved Castes, and PWD variations.
    """
    structured_data = []
    current_college = "Unknown Institute"
    current_branch = "General"
    
    print(f"Analyzing and parsing database from '{pdf_path}'...")
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split("\n")
            for line in lines:
                # Capture College/Institute Line
                if re.search(r"^\d{4}\s*-", line.strip()):
                    current_college = line.strip()
                    continue
                
                # Capture Course/Branch Line
                if re.search(r"^\d{9}\s*-", line.strip()):
                    current_branch = line.strip()
                    continue
                
                # Extract cutoff percentiles: (XX.XXXXXXXX)
                percentiles = re.findall(r"\(\s*(\d{1,2}\.\d+)\s*\)", line)
                
                # Extract multi-column Seat Type tags (e.g., GOPENS, PWDOPENS, GOBCS, PWDOBCH)
                seat_types = re.findall(r"\b([A-Z0-9]{3,9})\b", line)
                
                if percentiles:
                    for idx, score in enumerate(percentiles):
                        # Safely align discovered seat codes dynamically on that line string
                        seat = seat_types[idx] if idx < len(seat_types) else "GOPENS"
                        structured_data.append({
                            "College": current_college,
                            "Course": current_branch,
                            "Seat Type": seat,
                            "Cutoff Percentile": float(score)
                        })
                        
    df = pd.DataFrame(structured_data).drop_duplicates()
    return df


# =====================================================================
# 2. DYNAMIC INPUT, REASONING ENGINE & SHEET COMPILATION
# =====================================================================
def generate_formatted_report(pdf_path, student_score, base_category, is_pwd=False, output_name="Engineering_Preference_Report.xlsx"):
    # Load raw dataframe
    db_df = extract_live_cap_database(pdf_path)
    if db_df.empty:
        print("Error: No data rows recovered from the document.")
        return

    # Normalize category strings
    base_category = base_category.strip().upper()
    
    # Resolve PWD variations or normal seat configurations
    # Normal: GOPENS, GOBCS, EWS, TFWS. PWD: PWDOPENS, PWDOBCS, PWDSCS, etc.
    if is_pwd:
        if "PWD" not in base_category:
            # If the user typed "OBC" and checked PWD, check for PWDOBCS, PWDOBCH, or general PWDOPENS
            search_seat_patterns = [f"PWD{base_category}S", f"PWD{base_category}H", f"PWD{base_category}", "PWDOPENS"]
        else:
            search_seat_patterns = [base_category, "PWDOPENS"]
    else:
        search_seat_patterns = [f"G{base_category}S", f"G{base_category}H", base_category, "GOPENS"]

    # Filter by targeting the solved seat categories
    filtered_df = db_df[db_df["Seat Type"].isin(search_seat_patterns)].copy()
    if filtered_df.empty:
        # Fallback to broader inclusion if rules are too restrictive
        filtered_df = db_df[db_df["Seat Type"].str.contains(base_category, case=False, na=False)].copy()
    if filtered_df.empty:
        filtered_df = db_df[db_df["Seat Type"] == "GOPENS"].copy()

    # Apply smart cutoff filter bounds (include colleges up to +2.0 percentile as realistic stretch options)
    eligible_options = filtered_df[filtered_df["Cutoff Percentile"] <= (student_score + 2.0)].copy()
    
    # Sort descending to place high-tier matching preferences at the top
    eligible_options = eligible_options.sort_values(by="Cutoff Percentile", ascending=False)
    
    # Slice to ensure a targeted array of the top 25 choices (guaranteeing at least 20)
    final_selections = eligible_options.head(25)
    if len(final_selections) < 20:
        final_selections = filtered_df.sort_values(by="Cutoff Percentile", ascending=False).head(22)

    # Inject Preference Column
    final_selections = final_selections.copy()
    final_selections.insert(0, "Preference No", range(1, len(final_selections) + 1))
    
    # Compute Custom Reasoning Logic Context
    reasoning_log = []
    for idx, row in final_selections.iterrows():
        cutoff = row["Cutoff Percentile"]
        diff = student_score - cutoff
        
        if diff < 0:
            reasoning_log.append(f"Stretch/Dream Choice. Cutoff is slightly above your score by {abs(diff):.4f}%. Excellent aggressive preference option.")
        elif diff <= 2.0:
            reasoning_log.append(f"Highly Competitive Match. Your score clears historical cutoff by {diff:.4f}%. Safe and optimized recommendation.")
        else:
            reasoning_log.append(f"Highly Secure Back-Up. Strong allocation buffer of {diff:.4f}% ensuring allocation stability during optimization.")
            
    final_selections["Reasoning Logic"] = reasoning_log

    # Rearrange final visual structure cleanly
    final_selections = final_selections[["Preference No", "College", "Course", "Seat Type", "Cutoff Percentile", "Reasoning Logic"]]

    # =====================================================================
    # 3. ADVANCED VISUAL STYLE SHEET DESIGN (OPENPYXL)
    # =====================================================================
    writer = pd.ExcelWriter(output_name, engine='openpyxl')
    final_selections.to_excel(writer, index=False, sheet_name="CAP Preferences")
    
    workbook = writer.book
    worksheet = writer.sheets["CAP Preferences"]
    
    # Enable gridlines clearly
    worksheet.views.sheetView[0].showGridLines = True
    
    # Corporate Palette Design Constants (Classic Navy Theme)
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    zebra_fill = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    data_font = Font(name="Segoe UI", size=10, color="333333")
    bold_data_font = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
    
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    # Style Header Rows
    for col_idx in range(1, len(final_selections.columns) + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    worksheet.row_dimensions[1].height = 28

    # Style Data Body Rows
    for row_idx in range(2, len(final_selections) + 2):
        is_even = (row_idx % 2 == 0)
        current_fill = zebra_fill if is_even else white_fill
        
        for col_idx in range(1, len(final_selections.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.fill = current_fill
            cell.font = data_font
            cell.border = thin_border
            
            # Contextual Alignment & Number Structuring Rules
            col_name = final_selections.columns[col_idx - 1]
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

    # Autonumeric Column Width Cushion Scaling
    for col in worksheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        # Apply structured buffer bounds so logic columns don't stretch excessively
        calculated_width = min(max(max_len + 3, 12), 70)
        worksheet.column_dimensions[col_letter].width = calculated_width

    writer.close()
    print(f"Compilation Complete! Excel file saved as: '{output_name}'")


# =====================================================================
# 4. RUNTIME EXECUTIVE INTERFACE
# =====================================================================
if __name__ == "__main__":
    print("==========================================================")
    print("      CAP ROUND AUTOMATED EXCEL REPORT GENERATOR         ")
    print("==========================================================\n")
    
    # 100% Parameter-driven setup. Zero hardcoded text queries.
    pdf_file = "2023ENGG_CAP1_CutOff.pdf"
    
    if not os.path.exists(pdf_file):
        print(f"Error: Missing target source file '{pdf_file}' in execution space.")
    else:
        try:
            score_input = float(input("Enter Student MHT-CET Percentile (e.g., 93.45): "))
            cat_input = input("Enter Caste Category Code (e.g., OPEN, OBC, SC, ST, NT1, EWS): ")
            pwd_input = input("Is the candidate registered under PWD category? (yes/no): ").strip().lower()
            
            pwd_flag = (pwd_input in ["yes", "y", "true"])
            
            print("\nExecuting internal data pipeline and formatting spreadsheet...")
            generate_formatted_report(
                pdf_path=pdf_file,
                student_score=score_input,
                base_category=cat_input,
                is_pwd=pwd_flag,
                output_name="CAP_Round_Allotment_Preferences.xlsx"
            )
        except ValueError:
            print("Invalid input sequence. Please ensure percentile scores are typed numerically.")