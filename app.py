import os
from flask import Flask, render_template, request, jsonify, send_file
import io
import csv
import re
from datetime import datetime
import pandas as pd  # New import for Excel handling

app = Flask(__name__)

# Load student database
def load_student_database():
    students = {}
    try:
        with open('studentnumbers.txt', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.rsplit('\t', 1) if '\t' in line else line.rsplit(' ', 1)
                if len(parts) == 2:
                    name, number = parts[0].strip(), parts[1].strip()
                    students[name] = number
        return students
    except Exception:
        return {}

student_db = load_student_database()

@app.route('/')
def index():
    return render_template('index.html')

def process_attendance_data(df, week):
    """Process attendance data from either CSV or Excel"""
    absent_students = []
    first_module = None
    
    for _, row in df.iterrows():
        if str(row.get('Attendance', '')).lower() == 'absent':
            student_name = str(row.get('Student Name', ''))
            course_code = str(row.get('Course Code', ''))
            section_name = str(row.get('Section Name', ''))
            
            # Store first module code for filename
            if first_module is None and course_code:
                first_module = course_code
            
            # Extract qualification (everything before "(202")
            qualification = section_name.split('(202')[0].strip()
            
            # Determine year from course code
            year_match = re.search(r'^\D*(\d)', course_code)
            year = f"Year{year_match.group(1)}" if year_match else "N/A"
            
            # Get student number
            student_number = student_db.get(student_name, "N/A")
            
            absent_students.append({
                "Student Number": student_number,
                "Student Name": student_name,
                "Module(s)": course_code,
                "Qualification": qualification,
                "Year": year,
                "Week": f"Week{week}",
                "Day": "N/A",
                "Assessment(s)": "N/A",
                "Marks Obtained": "N/A",
                "Reason for AR": "Class Attendance_007"
            })
    
    return absent_students, first_module

@app.route('/process', methods=['POST'])
def process_attendance():
    try:
        # Get uploaded file
        file = request.files.get('attendance_file')
        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        
        # Get selected week
        week = request.form.get('week')
        if not week or not week.isdigit() or int(week) < 1 or int(week) > 16:
            return jsonify({"error": "Please select a valid week (1-16)"}), 400
        
        # Check file extension
        filename = file.filename.lower()
        
        if filename.endswith('.csv'):
            # Process CSV file
            content = file.read().decode('utf-8')
            df = pd.read_csv(io.StringIO(content))
        elif filename.endswith(('.xlsx', '.xls')):
            # Process Excel file
            df = pd.read_excel(file)
        else:
            return jsonify({"error": "Unsupported file type. Please upload CSV or Excel file."}), 400
        
        # Process the data
        absent_students, first_module = process_attendance_data(df, week)
        
        # Generate output CSV
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=[
            "Student Number", "Student Name", "Module(s)", "Qualification", 
            "Year", "Week", "Day", "Assessment(s)", 
            "Marks Obtained", "Reason for AR"
        ])
        writer.writeheader()
        writer.writerows(absent_students)
        
        # Create download response
        mem_file = io.BytesIO(output.getvalue().encode('utf-8'))
        mem_file.seek(0)
        
        # Generate filename
        current_date = datetime.now().strftime('%Y%m%d')
        filename = f"{first_module}_AR_Report_Week_{week}_{current_date}.csv" if first_module else f"AR_Report_Week_{week}_{current_date}.csv"
        
        return send_file(
            mem_file,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
