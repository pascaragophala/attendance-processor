import os
from flask import Flask, render_template, request, jsonify, send_file
import io
import csv
import re
from datetime import datetime

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
        
        # Process CSV file
        content = file.read().decode('utf-8')
        reader = csv.DictReader(io.StringIO(content))
        
        absent_students = []
        for row in reader:
            if row.get('Attendance', '').lower() == 'absent':
                student_name = row.get('Student Name', '')
                course_code = row.get('Course Code', '')
                section_name = row.get('Section Name', '')
                
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
                    "Week": week,
                    "Day": "N/A",
                    "Assessment(s)": "N/A",
                    "Marks Obtained": "N/A",
                    "Reason for AR": "Class Attendance_007"
                })
        
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
        
        filename = f"AR_Report_Week_{week}_{datetime.now().strftime('%Y%m%d')}.csv"
        
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