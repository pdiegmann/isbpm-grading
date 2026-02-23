import pandas as pd
import xlsxwriter
from typing import Dict, Any, List

def setup_grade_mapping_sheet(workbook, sheet_name='GradeMapping'):
    """Sets up a hidden sheet with the grading scale for VLOOKUP."""
    mapping_sheet = workbook.add_worksheet(sheet_name)
    scale = [
        (0.00, "5.0"),
        (0.50, "4.0"),
        (0.55, "3.7"),
        (0.60, "3.3"),
        (0.65, "3.0"),
        (0.70, "2.7"),
        (0.75, "2.3"),
        (0.80, "2.0"),
        (0.85, "1.7"),
        (0.90, "1.3"),
        (0.95, "1.0")
    ]
    for row, (pct, grade) in enumerate(scale):
        mapping_sheet.write_number(row, 0, pct)
        mapping_sheet.write_string(row, 1, grade)
    # mapping_sheet.hide() # Uncomment to hide this sheet in production
    return f"'{sheet_name}'!$A$1:$B$11"

def write_excel(output_path: str, students_df: pd.DataFrame, all_tasks: Dict[str, pd.DataFrame], all_other: Dict[str, pd.DataFrame]):
    """Generates the main Excel grading report using native Excel formulas with a polished layout."""
    
    with xlsxwriter.Workbook(output_path) as workbook:
        # Colors
        COLOR_PRIMARY = '#4F81BD'
        COLOR_SECONDARY = '#DCE6F1'
        COLOR_TOTAL = '#B8CCE4'
        COLOR_WHITE = '#FFFFFF'
        
        # Formats
        header_format = workbook.add_format({'bold': True, 'bg_color': COLOR_PRIMARY, 'font_color': COLOR_WHITE, 'bottom': 1})
        percent_format = workbook.add_format({'num_format': '0.00%'})
        score_format = workbook.add_format({'num_format': '0.0', 'align': 'center'})
        grade_format = workbook.add_format({'bold': True, 'align': 'center'})
        
        text_wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})
        title_format = workbook.add_format({'bold': True, 'font_size': 16, 'bottom': 2})
        
        section_format = workbook.add_format({'bold': True, 'bg_color': COLOR_PRIMARY, 'font_color': COLOR_WHITE, 'font_size': 12, 'valign': 'vcenter'})
        col_header_format = workbook.add_format({'bold': True, 'bg_color': COLOR_SECONDARY, 'bottom': 1, 'valign': 'vcenter'})
        col_header_center_format = workbook.add_format({'bold': True, 'bg_color': COLOR_SECONDARY, 'bottom': 1, 'align': 'center', 'valign': 'vcenter'})
        
        total_format = workbook.add_format({'bold': True, 'bg_color': COLOR_TOTAL, 'top': 1, 'bottom': 1, 'valign': 'vcenter'})
        total_pct_format = workbook.add_format({'bold': True, 'bg_color': COLOR_TOTAL, 'top': 1, 'bottom': 1, 'num_format': '0.00%', 'align': 'center', 'valign': 'vcenter'})
        final_grade_format = workbook.add_format({'bold': True, 'bg_color': COLOR_TOTAL, 'top': 1, 'bottom': 1, 'align': 'center', 'valign': 'vcenter'})
        
        cell_left_vcenter = workbook.add_format({'valign': 'vcenter'})
        cell_center_vcenter = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        merge_format = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'text_wrap': True, 'bold': True})
        
        mapping_range = "'GradeMapping'!$A$1:$B$11"
        
        # --- Master Overview Sheet ---
        master_sheet = workbook.add_worksheet('Master Overview')
        master_sheet.set_column('A:A', 15)
        master_sheet.set_column('B:C', 20)
        master_sheet.set_column('D:G', 18, percent_format)
        master_sheet.set_column('H:H', 15, grade_format)
        
        headers = ['Username', 'First Name', 'Last Name', 'Formalities', 'Practical Tasks', 'Solution Report', 'Total Percentage', 'German Grade']
        for col_num, h in enumerate(headers):
            master_sheet.write(0, col_num, h, header_format)
            
        master_row = 1
        
        for index, row_data in students_df.iterrows():
            username = row_data['Username']
            fname = row_data['First name']
            lname = row_data['Last name']
            
            sheet_name = str(username)[:31]
            ind_sheet = workbook.add_worksheet(sheet_name)
            
            master_sheet.write(master_row, 0, username)
            master_sheet.write(master_row, 1, fname)
            master_sheet.write(master_row, 2, lname)
            
            tasks_df = all_tasks.get(username)
            other_df = all_other.get(username)
            
            # Layout Mapping:
            # A: Category / Task (width 25)
            # B: Item (width 30 for text)
            # C: Score 1 (width 12)
            # D: Max 1 (width 12)
            # E: Score 2 (width 12)
            # F: Max 2 (width 12)
            # G: Empty Spacer (width 2)
            # H: Notes (width 50)
            # I: Hidden column for Task %
            ind_sheet.set_column('A:A', 25)
            ind_sheet.set_column('B:B', 30)
            ind_sheet.set_column('C:F', 12)
            ind_sheet.set_column('G:G', 2)
            ind_sheet.set_column('H:H', 50, text_wrap_format)
            ind_sheet.set_column('I:I', None, None, {'hidden': True})
            
            row = 0
            ind_sheet.merge_range(row, 0, row, 7, f"Grading Report: {fname} {lname} ({username})", title_format)
            row += 2
            
            pct_cells = {}
            
            # -------------------------------------------------------------
            # 1. Overall Formalities
            # -------------------------------------------------------------
            ind_sheet.merge_range(row, 0, row, 7, "1. Overall Formalities (20%)", section_format)
            row += 1
            
            ind_sheet.merge_range(row, 0, row, 1, "Category", col_header_format)
            ind_sheet.write(row, 2, "Score", col_header_center_format)
            ind_sheet.write(row, 3, "Max", col_header_center_format)
            ind_sheet.write(row, 4, "", col_header_format)
            ind_sheet.write(row, 5, "", col_header_format)
            ind_sheet.write(row, 6, "", col_header_format)
            ind_sheet.write(row, 7, "Notes", col_header_format)
            row += 1
            
            formalities_data = [
                ("Formatting", 0, 0.3, "Overall", "Formatting"),
                ("Structure", 0, 0.5, "Overall", "Structure"),
                ("Style/Language", 0, 0.2, "Overall", "Style/Language")
            ]
            
            formalities_score_cells = []
            
            for label, default_score, weight, cat, item in formalities_data:
                score = default_score
                notes = ""
                if other_df is not None and not other_df.empty:
                    match = other_df[(other_df['Category'] == cat) & (other_df['Item'].str.contains(item, case=False, na=False))]
                    if not match.empty:
                        score = float(match['Score'].values[0]) if not pd.isna(match['Score'].values[0]) else 0.0
                        notes = str(match['Notes'].values[0]) if 'Notes' in match.columns else ""
                
                ind_sheet.merge_range(row, 0, row, 1, f"{label} (w: {weight*100:.0f}%)", cell_left_vcenter)
                ind_sheet.write_number(row, 2, score, score_format)
                ind_sheet.write_number(row, 3, 2.0, score_format)
                ind_sheet.write(row, 4, "")
                ind_sheet.write(row, 5, "")
                ind_sheet.write(row, 6, "")
                ind_sheet.write_string(row, 7, notes)
                
                formalities_score_cells.append( (row+1, weight) )
                row += 1
            
            f_formula = "=SUM(" + ",".join([f"((C{r}/D{r})*{w})" for r, w in formalities_score_cells]) + ")"
            ind_sheet.merge_range(row, 0, row, 3, "Formalities Sub-Total %", total_format)
            ind_sheet.write_formula(row, 4, f_formula, total_pct_format)
            ind_sheet.write(row, 5, "", total_format)
            ind_sheet.write(row, 6, "", total_format)
            ind_sheet.write(row, 7, "", total_format)
            pct_cells['formalities'] = f"'{sheet_name}'!E{row+1}"
            row += 2
            
            # -------------------------------------------------------------
            # 2. Solution Report
            # -------------------------------------------------------------
            ind_sheet.merge_range(row, 0, row, 7, "2. Solution Report (40%)", section_format)
            row += 1
            
            ind_sheet.write(row, 0, "Dimension", col_header_format)
            ind_sheet.write(row, 1, "Item", col_header_format)
            ind_sheet.write(row, 2, "Score", col_header_center_format)
            ind_sheet.write(row, 3, "Max", col_header_center_format)
            ind_sheet.write(row, 4, "", col_header_format)
            ind_sheet.write(row, 5, "", col_header_format)
            ind_sheet.write(row, 6, "", col_header_format)
            ind_sheet.write(row, 7, "Notes", col_header_format)
            row += 1
            
            sr_metrics = [
                {"dim": "Approach", "weight": 0.5, "sub": [
                    ("Correctness", 2, 0.5, 'Solution Report', 'Approach - Correctness'),
                    ("Convincingness", 2, 0.35, 'Solution Report', 'Approach - Convincingness'),
                    ("References", 1, 0.15, 'Solution Report', 'Approach - References'),
                ]},
                {"dim": "Context &\nSituationality", "weight": 0.25, "sub": [
                    ("Correctness", 2, 0.5, 'Solution Report', 'Context & Situationality - Correctness'),
                    ("Convincingness", 2, 0.35, 'Solution Report', 'Context & Situationality - Convincingness'),
                    ("References", 1, 0.15, 'Solution Report', 'Context & Situationality - References'),
                ]},
                {"dim": "Implications", "weight": 0.25, "sub": [
                    ("Correctness", 2, 0.5, 'Solution Report', 'Implications - Correctness'),
                    ("Convincingness", 2, 0.35, 'Solution Report', 'Implications - Convincingness'),
                    ("References", 1, 0.15, 'Solution Report', 'Implications - References'),
                ]}
            ]
            
            dim_cells = []
            for sr in sr_metrics:
                dim_row_start = row
                sub_cells = []
                for label, max_val, weight, cat, item in sr['sub']:
                    score = 0
                    notes = ""
                    if other_df is not None and not other_df.empty:
                        match = other_df[(other_df['Category'] == cat) & (other_df['Item'].str.contains(item, case=False, na=False))]
                        if not match.empty:
                            score = float(match['Score'].values[0]) if not pd.isna(match['Score'].values[0]) else 0.0
                            notes = str(match['Notes'].values[0]) if 'Notes' in match.columns else ""
                            
                    ind_sheet.write(row, 1, f"{label} (w: {weight*100:.0f}%)", cell_left_vcenter)
                    ind_sheet.write_number(row, 2, score, score_format)
                    ind_sheet.write_number(row, 3, max_val, score_format)
                    ind_sheet.write(row, 4, "")
                    ind_sheet.write(row, 5, "")
                    ind_sheet.write(row, 6, "")
                    ind_sheet.write_string(row, 7, notes)
                    
                    sub_cells.append(f"((C{row+1}/D{row+1})*{weight})")
                    row += 1
                    
                ind_sheet.merge_range(dim_row_start, 0, row-1, 0, f"{sr['dim']}\n(w: {sr['weight']*100:.0f}%)", merge_format)
                
                dim_formula = "=SUM(" + ",".join(sub_cells) + ")"
                ind_sheet.merge_range(row, 0, row, 3, f"{sr['dim'].replace(chr(10), ' ')} Sub-Total", total_format)
                ind_sheet.write_formula(row, 4, dim_formula, total_pct_format)
                ind_sheet.write(row, 5, "", total_format)
                ind_sheet.write(row, 6, "", total_format)
                ind_sheet.write(row, 7, "", total_format)
                dim_cells.append(f"(E{row+1}*{sr['weight']})")
                row += 1
                
            sr_formula = "=SUM(" + ",".join(dim_cells) + ")"
            ind_sheet.merge_range(row, 0, row, 3, "Solution Report Final Sub-Total %", total_format)
            ind_sheet.write_formula(row, 4, sr_formula, total_pct_format)
            ind_sheet.write(row, 5, "", total_format)
            ind_sheet.write(row, 6, "", total_format)
            ind_sheet.write(row, 7, "", total_format)
            pct_cells['solution_report'] = f"'{sheet_name}'!E{row+1}"
            row += 2
            
            # -------------------------------------------------------------
            # 3. Practical Tasks
            # -------------------------------------------------------------
            ind_sheet.merge_range(row, 0, row, 7, "3. Practical Tasks (40%)", section_format)
            row += 1
            
            ind_sheet.merge_range(row, 0, row, 1, "Task", col_header_format)
            ind_sheet.write(row, 2, "Corr. Score", col_header_center_format)
            ind_sheet.write(row, 3, "Corr. Max", col_header_center_format)
            ind_sheet.write(row, 4, "Det. Score", col_header_center_format)
            ind_sheet.write(row, 5, "Det. Max", col_header_center_format)
            ind_sheet.write(row, 6, "", col_header_format)
            ind_sheet.write(row, 7, "Notes", col_header_format)
            row += 1
            
            task_pct_cells = []
            if tasks_df is not None and not tasks_df.empty:
                for _, t_row in tasks_df.iterrows():
                    t_name = t_row['Task']
                    c_score = float(t_row.get('practicalTaskCorrect', 0)) if not pd.isna(t_row.get('practicalTaskCorrect', 0)) else 0
                    d_score = float(t_row.get('practicalTaskDetails', 0)) if not pd.isna(t_row.get('practicalTaskDetails', 0)) else 0
                    
                    ind_sheet.merge_range(row, 0, row, 1, t_name, cell_left_vcenter)
                    ind_sheet.write_number(row, 2, c_score, score_format)
                    ind_sheet.write_number(row, 3, 2.0, score_format)
                    ind_sheet.write_number(row, 4, d_score, score_format)
                    ind_sheet.write_number(row, 5, 1.0, score_format)
                    ind_sheet.write(row, 6, "")
                    ind_sheet.write(row, 7, "")
                    
                    task_form = f"=((C{row+1}/D{row+1})*0.65)+((E{row+1}/F{row+1})*0.35)"
                    ind_sheet.write_formula(row, 8, task_form) # Column I (hidden)
                    task_pct_cells.append(f"I{row+1}")
                    row += 1
                
                tasks_formula = f"=AVERAGE({task_pct_cells[0]}:{task_pct_cells[-1]})"
            else:
                tasks_formula = "=0"
                ind_sheet.merge_range(row, 0, row, 7, "No tasks data.", cell_left_vcenter)
                row += 1
            
            ind_sheet.merge_range(row, 0, row, 3, "Practical Tasks Average %", total_format)
            ind_sheet.write_formula(row, 4, tasks_formula, total_pct_format)
            ind_sheet.write(row, 5, "", total_format)
            ind_sheet.write(row, 6, "", total_format)
            ind_sheet.write(row, 7, "", total_format)
            pct_cells['practical_tasks'] = f"'{sheet_name}'!E{row+1}"
            row += 3
            
            # -------------------------------------------------------------
            # Final Section
            # -------------------------------------------------------------
            ind_sheet.merge_range(row, 0, row, 7, "--- FINAL AGGREGATION ---", section_format)
            row += 1
            
            f_cell = pct_cells['formalities'].split('!')[1]
            sr_cell = pct_cells['solution_report'].split('!')[1]
            pt_cell = pct_cells['practical_tasks'].split('!')[1]
            
            ind_sheet.merge_range(row, 0, row, 3, "Overall Formalities (20%)", cell_left_vcenter)
            ind_sheet.write_formula(row, 4, f"={f_cell}", percent_format)
            ind_sheet.write(row, 5, "")
            ind_sheet.write(row, 6, "")
            ind_sheet.write(row, 7, "")
            row +=1
            
            ind_sheet.merge_range(row, 0, row, 3, "Solution Report (40%)", cell_left_vcenter)
            ind_sheet.write_formula(row, 4, f"={sr_cell}", percent_format)
            ind_sheet.write(row, 5, "")
            ind_sheet.write(row, 6, "")
            ind_sheet.write(row, 7, "")
            row +=1
            
            ind_sheet.merge_range(row, 0, row, 3, "Practical Tasks (40%)", cell_left_vcenter)
            ind_sheet.write_formula(row, 4, f"={pt_cell}", percent_format)
            ind_sheet.write(row, 5, "")
            ind_sheet.write(row, 6, "")
            ind_sheet.write(row, 7, "")
            row +=1
            
            # Compute Total Percentage on the individual sheet
            ind_tot_pct_formula = f"={f_cell}*0.2 + {pt_cell}*0.4 + {sr_cell}*0.4"
            ind_sheet.merge_range(row, 0, row, 3, "Total Final Percentage", total_format)
            ind_sheet.write_formula(row, 4, ind_tot_pct_formula, total_pct_format)
            ind_sheet.write(row, 5, "", total_format)
            ind_sheet.write(row, 6, "", total_format)
            ind_sheet.write(row, 7, "", total_format)
            tot_pct_row = row + 1
            row += 1
            
            # Compute Final Grade on the individual sheet using VLOOKUP
            ind_grade_formula = f"=VLOOKUP(E{tot_pct_row}, {mapping_range}, 2, TRUE)"
            ind_sheet.merge_range(row, 0, row, 3, "Total Final German Grade", total_format)
            ind_sheet.write_formula(row, 4, ind_grade_formula, final_grade_format)
            ind_sheet.write(row, 5, "", total_format)
            ind_sheet.write(row, 6, "", total_format)
            ind_sheet.write(row, 7, "", total_format)
            grade_row = row + 1
            
            # Link Master Overview to Individual Sheet
            master_sheet.write_formula(master_row, 3, "=" + pct_cells['formalities'], percent_format)
            master_sheet.write_formula(master_row, 4, "=" + pct_cells['practical_tasks'], percent_format)
            master_sheet.write_formula(master_row, 5, "=" + pct_cells['solution_report'], percent_format)
            master_sheet.write_formula(master_row, 6, f"='{sheet_name}'!E{tot_pct_row}", percent_format)
            master_sheet.write_formula(master_row, 7, f"='{sheet_name}'!E{grade_row}", grade_format)
            
            master_row += 1

        # Move Mapping sheet to end
        setup_grade_mapping_sheet(workbook, 'GradeMapping')
