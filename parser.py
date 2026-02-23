import pandas as pd
from pathlib import Path
from typing import Dict, Any, List, Optional
from unidecode import unidecode

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

def find_student_grading_files(base_dir: Path, username: str, firstname: str = "", lastname: str = "") -> Dict[str, Optional[Path]]:
    """Tries to find the grading files for a specific student username or name."""
    
    # Standard format: {username}-...
    other_file = base_dir / f"{username}-other.csv"
    tasks_file = base_dir / f"{username}-tasks.csv"
    text_file = base_dir / f"{username}.txt"
    
    if other_file.exists() or tasks_file.exists() or text_file.exists():
        return {
            'other': other_file if other_file.exists() else None,
            'tasks': tasks_file if tasks_file.exists() else None,
            'text': text_file if text_file.exists() else None
        }
        
    # AI format: {LastName}-{FirstName}-...
    if firstname and lastname:
        # Some special character replacements might happen, so case-insensitive matching is safer
        name_prefix = unidecode(f"{lastname}-{firstname}").lower()
        
        for file in base_dir.iterdir():
            if not file.is_file():
                continue
                
            fname_lower = unidecode(file.name).lower()
            if fname_lower.startswith(name_prefix):
                if fname_lower.endswith('-other.csv'):
                    other_file = file
                elif fname_lower.endswith('-tasks.csv'):
                    tasks_file = file
                elif fname_lower.endswith('.txt'):
                    text_file = file
                    
    return {
        'other': other_file if other_file.exists() and other_file.is_file() else None,
        'tasks': tasks_file if tasks_file.exists() and tasks_file.is_file() else None,
        'text': text_file if text_file.exists() and text_file.is_file() else None
    }

def parse_grading_text(filepath: str | Path) -> str:
    """Parses a text file and returns its content as a single string."""
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()
