import os
import pdfplumber
import pandas as pd
import re

def extract_data_from_pdf(pdf_path):
    """
    Extracts student information and grades from a single PDF file.
    """
    prn = os.path.basename(pdf_path).replace('.pdf', '')
    
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        
        student_data = {'PRN': prn}
        
        seat_no_match = re.search(r'SEAT NO. : (\w+)', text)
        student_data['Seat Number'] = seat_no_match.group(1) if seat_no_match else None
            
        name_match = re.search(r'NAME : ([A-Z ]+)', text)
        student_data['Name'] = name_match.group(1).strip() if name_match else None
            
        gpa_match = re.search(r'GPA:([\d.]+)', text)
        student_data['GPA'] = float(gpa_match.group(1)) if gpa_match else None
            
        tables = page.extract_tables()
        
        if tables and len(tables) > 0:
            table = tables[0]
            
            # Find header row index
            header_index = -1
            for i, row in enumerate(table):
                if row and 'COURSE' in (row[1] or ''):
                    header_index = i
                    break
            
            if header_index != -1:
                header = [h.replace('\n', ' ') if h else '' for h in table[header_index]]
                
                # Find the 'COURSE' and 'TOTAL GRADE' column indices
                try:
                    course_col_index = header.index('COURSE')
                    ca_col_index = header.index('CA GRADE')
                    te_col_index = header.index('TERM END GRADE')
                    prac_col_index = header.index('PRACTICAL GRADE')
                    total_col_index = header.index('TOTAL GRADE')
                except ValueError:
                    # Handle cases where a grade column might be missing
                    # For now, we'll just log and skip, but this could be made more robust
                    print(f"Warning: Could not find all expected columns in {pdf_path}")
                    return student_data

                for i in range(header_index + 1, len(table)):
                    row = table[i]
                    if row[0] and 'RESULT DATE' in (row[0] or ''):
                        break
                    
                    course_name_cell = row[course_col_index]
                    if course_name_cell is not None:
                        subject_name = course_name_cell.strip()
                        if subject_name:
                            student_data[f'{subject_name} - CA'] = row[ca_col_index]
                            student_data[f'{subject_name} - Term End'] = row[te_col_index]
                            student_data[f'{subject_name} - Practical'] = row[prac_col_index]
                            student_data[f'{subject_name} - Total'] = row[total_col_index]

    return student_data

def calculate_descriptive_statistics(df, writer):
    """
    Calculates descriptive statistics and adds them to a new sheet in the Excel file.
    """
    gpa_stats = df['GPA'].describe(percentiles=[.25, .5, .75, .9]).to_frame('GPA Statistics')
    gpa_stats.loc['median'] = df['GPA'].median()
    gpa_stats.loc['mode'] = df['GPA'].mode().to_string(index=False)
    gpa_stats.loc['variance'] = df['GPA'].var()
    gpa_stats.loc['skewness'] = df['GPA'].skew()
    gpa_stats.loc['kurtosis'] = df['GPA'].kurt()
    gpa_stats.loc['range'] = df['GPA'].max() - df['GPA'].min()
    
    gpa_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=0)

    # Grade Distribution Analysis
    grade_cols = [col for col in df.columns if col.endswith(' - Total')]
    all_grades = df[grade_cols].stack().dropna()
    
    grade_counts = all_grades.value_counts().to_frame('Frequency')
    grade_counts['Percentage'] = (grade_counts['Frequency'] / grade_counts['Frequency'].sum() * 100).round(2)
    
    grade_counts.to_excel(writer, sheet_name='Statistical Summary', startrow=len(gpa_stats)+3)
    
    # Write titles
    workbook = writer.book
    sheet = writer.sheets['Statistical Summary']
    sheet.cell(row=1, column=1, value="Overall Performance Metrics")
    sheet.cell(row=len(gpa_stats)+3, column=1, value="Grade Distribution Analysis")

def calculate_subject_wise_analysis(df, writer):
    """
    Performs subject-wise analysis and adds it to a new sheet.
    """
    grade_points = {'O': 10, 'A+': 9, 'A': 8, 'B+': 7, 'B': 6, 'C': 5, 'P': 4, 'F': 0}
    
    total_grade_cols = [col for col in df.columns if col.endswith(' - Total')]
    subjects = [col.replace(' - Total', '') for col in total_grade_cols]
    
    subject_analysis = []

    for subject in subjects:
        total_col = f'{subject} - Total'
        
        # Grade Distribution
        grade_counts = df[total_col].value_counts()
        
        # Pass/Fail
        pass_count = grade_counts.drop('F', errors='ignore').sum()
        fail_count = grade_counts.get('F', 0)
        
        # Difficulty Index (Avg Grade Points)
        valid_grades = df[total_col].dropna()
        grade_pts = valid_grades.map(grade_points)
        avg_grade_points = grade_pts.mean()

        analysis_row = {
            'Subject': subject,
            'Average Grade Points': round(avg_grade_points, 2),
            'Pass Count': pass_count,
            'Fail Count': fail_count,
        }
        
        for grade in grade_points.keys():
            analysis_row[f'Grade {grade}'] = grade_counts.get(grade, 0)
            
        subject_analysis.append(analysis_row)

    analysis_df = pd.DataFrame(subject_analysis)
    analysis_df.to_excel(writer, sheet_name='Subject Analysis', index=False)

def calculate_correlation_analysis(df, writer):
    """
    Performs correlation analysis and adds it to a new sheet.
    """
    grade_points = {'O': 10, 'A+': 9, 'A': 8, 'B+': 7, 'B': 6, 'C': 5, 'P': 4, 'F': 0}
    
    total_grade_cols = [col for col in df.columns if col.endswith(' - Total')]
    subjects = [col.replace(' - Total', '') for col in total_grade_cols]
    
    # Create a new DataFrame for correlation analysis
    correlation_df = df[['PRN', 'GPA']].copy()
    
    for subject in subjects:
        total_col = f'{subject} - Total'
        correlation_df[subject] = df[total_col].map(grade_points)
        
    # Drop PRN for correlation calculation
    correlation_df = correlation_df.drop(columns=['PRN'])
    
    # Compute the correlation matrix
    correlation_matrix = correlation_df.corr()
    
    correlation_matrix.to_excel(writer, sheet_name='Correlation Analysis')

def main():
    """
    Main function to orchestrate the PDF data extraction and Excel report generation.
    """
    results_folder = 'Results'
    pdf_files = [os.path.join(results_folder, f) for f in os.listdir(results_folder) if f.endswith('.pdf')]
    
    all_student_data = []
    
    print(f"Found {len(pdf_files)} PDF files to process.")
    
    for pdf_file in pdf_files:
        print(f"Processing {pdf_file}...")
        try:
            data = extract_data_from_pdf(pdf_file)
            all_student_data.append(data)
        except Exception as e:
            print(f"Could not process {pdf_file}. Error: {e}")
            
    if not all_student_data:
        print("No data was extracted. Exiting.")
        return

    df = pd.DataFrame(all_student_data)
    
    # Reorder columns to have student info first
    cols = ['PRN', 'Seat Number', 'Name', 'GPA']
    other_cols = [c for c in df.columns if c not in cols]
    df = df[cols + sorted(other_cols)] # Sort subject columns alphabetically
    
    print("\nFirst 5 rows of extracted data:")
    print(df.head())
    
    output_excel_path = 'semester_analysis.xlsx'
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Master Sheet', index=False)
        calculate_descriptive_statistics(df, writer)
        calculate_subject_wise_analysis(df, writer)
        calculate_correlation_analysis(df, writer)
        
    print(f"\nAnalysis complete. Results saved to: {output_excel_path}")

if __name__ == '__main__':
    main()