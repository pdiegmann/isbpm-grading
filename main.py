import argparse
from pathlib import Path
from parser import parse_students, parse_grading_other, parse_grading_tasks, parse_grading_text, find_student_grading_files
from excel_generator import write_excel

def main():
    parser = argparse.ArgumentParser(description="Grading Support Tool (Native Excel Formulas)")
    parser.add_argument('--input-dir', type=str, required=True, help="Directory containing students and grading CSVs")
    parser.add_argument('--output', type=str, default='grades_output.xlsx', help="Path to the output Excel file")
    
    args = parser.parse_args()
    input_path = Path(args.input_dir)
    output_path = args.output
    
    if not input_path.exists():
        print(f"Error: Input directory '{input_path}' does not exist.")
        return
        
    # Find students.csv
    students_file = input_path / 'students.csv'
    if not students_file.exists():
        students_file = input_path / 'example-students.csv'
    
    if not students_file.exists():
        print(f"Error: Could not find students.csv or example-students.csv in {input_path}")
        return
        
    print(f"Loading students from {students_file}...")
    students_df = parse_students(students_file)
    print(f"Identified {len(students_df)} students.")
    
    all_tasks = {}
    all_other = {}
    all_texts = {}
    
    for _, student in students_df.iterrows():
        username = student['Username']
        firstname = student.get('First name', "")
        lastname = student.get('Last name', "")
        
        found_files = find_student_grading_files(input_path, username, firstname, lastname)
        
        other_file = found_files['other']
        tasks_file = found_files['tasks']
        text_file = found_files['text']
        
        # Override for the purely example files
        if username == 'jakbrz': 
            if (input_path / "example-grading-other.csv").exists():
                other_file = input_path / "example-grading-other.csv"
            if (input_path / "example-grading-tasks.csv").exists():
                tasks_file = input_path / "example-grading-tasks.csv"
            if (input_path / "example-grading-text.txt").exists():
                text_file = input_path / "example-grading-text.txt"
        
        if other_file or tasks_file or text_file:
            print(f"Found grading files for {username}")
            
            if other_file:
                all_other[username] = parse_grading_other(other_file)
            if tasks_file:
                all_tasks[username] = parse_grading_tasks(tasks_file)
            if text_file:
                all_texts[username] = parse_grading_text(text_file)
                
    print(f"Finished parsing. Building Excel workbook at {output_path}...")
    write_excel(output_path, students_df, all_tasks, all_other, all_texts)
    print("Done!")

if __name__ == "__main__":
    main()
