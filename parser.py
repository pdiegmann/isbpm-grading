import pandas as pd
from pathlib import Path
from typing import Dict, Any, List, Optional

def parse_students(filepath: str | Path) -> pd.DataFrame:
    """Parses the students CSV."""
    df = pd.read_csv(filepath, sep=';', dtype=str)
    # Filter out accepted/tutor accounts if needed, depending on 'Status' or 'Position'
    if 'Status' in df.columns:
        df = df[df['Status'].isin(['autor', 'accepted'])]
    return df

def parse_grading_other(filepath: str | Path) -> pd.DataFrame:
    """Parses the *-other.csv which contains Overall and Solution Report scores/comments."""
    df = pd.read_csv(filepath, sep=',', dtype=str)
    # Convert 'Score' to float where possible
    df['Score'] = pd.to_numeric(df['Score'], errors='coerce')
    # Fill NaN notes with empty string
    if 'Notes' in df.columns:
        df['Notes'] = df['Notes'].fillna("")
    return df

def parse_grading_tasks(filepath: str | Path) -> pd.DataFrame:
    """Parses the *-tasks.csv which contains per-task grading points."""
    df = pd.read_csv(filepath, sep=',', dtype=str)
    # Convert all columns except 'Task' to floats
    for col in df.columns:
        if col != 'Task':
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def find_student_grading_files(base_dir: Path, username: str) -> Dict[str, Optional[Path]]:
    """Tries to find the grading files for a specific student username."""
    other_file = base_dir / f"{username}-other.csv"
    tasks_file = base_dir / f"{username}-tasks.csv"
    
    # In case there's a prefix, or they follow the example `example-grading-*`
    # We will also support the direct username if that matches
    return {
        'other': other_file if other_file.exists() else None,
        'tasks': tasks_file if tasks_file.exists() else None
    }
