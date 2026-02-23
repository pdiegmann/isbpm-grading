import pandas as pd
import numpy as np
from typing import Dict, Any

def get_german_grade(percentage: float) -> str:
    """Maps a percentage (0.0 to 1.0) to the German grading scale 1.0 to 5.0"""
    # Using typical university bounds:
    # >= 0.95: 1.0
    # >= 0.90: 1.3
    # >= 0.85: 1.7
    # >= 0.80: 2.0
    # >= 0.75: 2.3
    # >= 0.70: 2.7
    # >= 0.65: 3.0
    # >= 0.60: 3.3
    # >= 0.55: 3.7
    # >= 0.50: 4.0
    # <  0.50: 5.0
    if percentage >= 0.95: return "1.0"
    if percentage >= 0.90: return "1.3"
    if percentage >= 0.85: return "1.7"
    if percentage >= 0.80: return "2.0"
    if percentage >= 0.75: return "2.3"
    if percentage >= 0.70: return "2.7"
    if percentage >= 0.65: return "3.0"
    if percentage >= 0.60: return "3.3"
    if percentage >= 0.55: return "3.7"
    if percentage >= 0.50: return "4.0"
    return "5.0"

def calculate_dim_percentage(correctness: float, convincingness: float, references: float) -> float:
    """Calculates percentage for a Solution Report Dimension (Approach, Context, Implications)"""
    return (correctness / 2.0) * 0.5 + (convincingness / 2.0) * 0.35 + (references / 1.0) * 0.15

def calculate_grades(other_df: pd.DataFrame, tasks_df: pd.DataFrame) -> Dict[str, Any]:
    res = {}
    
    # --- 1. Practical Tasks (40%) ---
    # For each task: (correctness/2)*0.65 + (detail/1)*0.35
    # Then average across all tasks
    if tasks_df is not None and not tasks_df.empty:
        tasks_df['task_pct'] = (tasks_df['practicalTaskCorrect'] / 2.0) * 0.65 + (tasks_df['practicalTaskDetails'] / 1.0) * 0.35
        practical_tasks_pct = tasks_df['task_pct'].mean()
        res['practical_tasks_pct'] = practical_tasks_pct
        
        # Calculate per-task Solution Report averages if needed for fallback
        # Just compute them as additional info
        tasks_df['approach_pct'] = calculate_dim_percentage(tasks_df['approachCorrect'], tasks_df['approachConvincing'], tasks_df['approachReferences'])
        tasks_df['sit_pct'] = calculate_dim_percentage(tasks_df['situationalityCorrect'], tasks_df['situationalityConvincing'], tasks_df['situationalityReferences'])
        tasks_df['imp_pct'] = calculate_dim_percentage(tasks_df['implicationsCorrect'], tasks_df['implicationsConvincing'], tasks_df['implicationsReferences'])
    else:
        res['practical_tasks_pct'] = 0.0

    # --- 2. Overall Formalities (20%) ---
    res['formalities_pct'] = 0.0
    res['solution_report_pct'] = 0.0
    
    if other_df is not None and not other_df.empty:
        # Exctract Formatting, Structure, Style
        def get_score(cat, item):
            row = other_df[(other_df['Category'] == cat) & (other_df['Item'].str.contains(item, case=False, na=False))]
            return row['Score'].values[0] if not row.empty else 0.0
            
        fmt = get_score('Overall', 'Formatting')
        struc = get_score('Overall', 'Structure')
        style = get_score('Overall', 'Style/Language')
        
        formalities_pct = (fmt / 2.0) * 0.3 + (struc / 2.0) * 0.5 + (style / 2.0) * 0.2
        res['formalities_pct'] = formalities_pct
        
        # --- 3. Solution Report (40%) ---
        app_corr = get_score('Solution Report', 'Approach - Correctness')
        app_conv = get_score('Solution Report', 'Approach - Convincingness')
        app_ref = get_score('Solution Report', 'Approach - References')
        
        sit_corr = get_score('Solution Report', 'Context & Situationality - Correctness')
        sit_conv = get_score('Solution Report', 'Context & Situationality - Convincingness')
        sit_ref = get_score('Solution Report', 'Context & Situationality - References')
        
        imp_corr = get_score('Solution Report', 'Implications - Correctness')
        imp_conv = get_score('Solution Report', 'Implications - Convincingness')
        imp_ref = get_score('Solution Report', 'Implications - References')
        
        approach_pct = calculate_dim_percentage(app_corr, app_conv, app_ref)
        sit_pct = calculate_dim_percentage(sit_corr, sit_conv, sit_ref)
        imp_pct = calculate_dim_percentage(imp_corr, imp_conv, imp_ref)
        
        solution_report_pct = approach_pct * 0.5 + sit_pct * 0.25 + imp_pct * 0.25
        res['solution_report_pct'] = solution_report_pct
    else:
        # Fallback to tasks_df if other_df is missing (unlikely based on example but safe)
        if tasks_df is not None and not tasks_df.empty:
            avg_app = tasks_df['approach_pct'].mean()
            avg_sit = tasks_df['sit_pct'].mean()
            avg_imp = tasks_df['imp_pct'].mean()
            res['solution_report_pct'] = avg_app * 0.5 + avg_sit * 0.25 + avg_imp * 0.25

    # Compute Total Percentage
    res['total_pct'] = res['formalities_pct'] * 0.20 + res['practical_tasks_pct'] * 0.40 + res['solution_report_pct'] * 0.40
    res['german_grade'] = get_german_grade(res['total_pct'])
    
    return res
