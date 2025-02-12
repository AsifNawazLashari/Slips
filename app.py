from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
from fpdf import FPDF
from datetime import datetime

app = Flask(__name__)

# Directory Paths
UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = '/sdcard/generated_slips'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def get_excel_file(class_name):
    return f'students_class_{class_name}.xlsx'

def get_pdf_folder(class_name):
    class_pdf_folder = os.path.join(PDF_FOLDER, f'class_{class_name}')
    os.makedirs(class_pdf_folder, exist_ok=True)
    return class_pdf_folder

def save_to_excel(data, class_name):
    database_file = get_excel_file(class_name)
    if os.path.exists(database_file):
        df = pd.read_excel(database_file, engine='openpyxl')
    else:
        df = pd.DataFrame(columns=["GR No", "Name", "Father Name", "Caste", "DOB", "Roll Number", "Profile Picture", "Class"])

    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(database_file, index=False, engine='openpyxl')


def get_schedule(class_name):
    schedules = {
        "6": [  # Schedule for Class 6
            ['03-03-2025', 'Monday', 'English', '08:30 AM - 11:30 AM'],
            ['04-03-2025', 'Tuesday', 'Mathematics', '08:30 AM - 11:30 AM'],
            ['05-03-2025', 'Wednesday', 'General Science', '08:30 AM - 11:30 AM'],
            ['06-03-2025', 'Thursday', 'Social Studies', '08:30 AM - 11:30 AM'],
            ['07-03-2025', 'Friday', 'Islamiyat', '08:30 AM - 11:30 AM'],
            ['08-03-2025', 'Saturday', 'M.T', '08:30 AM - 11:30 AM'],
            ['10-03-2025', 'Monday', 'N.M.T', '08:30 AM - 11:30 AM'],
            ['11-03-2025', 'Tuesday', 'Computer', '08:30 AM - 11:30 AM'],
            ['12-03-2025', 'Wednesday', 'Drawing', '08:30 AM - 11:30 AM'],
            ['13-03-2025', 'Thursday', 'Arabic', '08:30 AM - 11:30 AM'],
        ],
        "7": [  # Schedule for Class 7 (same as Class 6)
            ['03-03-2025', 'Monday', 'English', '08:30 AM - 11:30 AM'],
            ['04-03-2025', 'Tuesday', 'Mathematics', '08:30 AM - 11:30 AM'],
            ['05-03-2025', 'Wednesday', 'General Science', '08:30 AM - 11:30 AM'],
            ['06-03-2025', 'Thursday', 'Social Studies', '08:30 AM - 11:30 AM'],
            ['07-03-2025', 'Friday', 'Islamiyat', '08:30 AM - 11:30 AM'],
            ['08-03-2025', 'Saturday', 'M.T', '08:30 AM - 11:30 AM'],
            ['10-03-2025', 'Monday', 'N.M.T', '08:30 AM - 11:30 AM'],
            ['11-03-2025', 'Tuesday', 'Computer', '08:30 AM - 11:30 AM'],
            ['12-03-2025', 'Wednesday', 'Drawing', '08:30 AM - 11:30 AM'],
            ['13-03-2025', 'Thursday', 'Arabic', '08:30 AM - 11:30 AM'],
        ],
        "8": [  # Schedule for Class 8 (Computer subject missing)
            ['03-03-2025', 'Monday', 'English', '08:30 AM - 11:30 AM'],
            ['04-03-2025', 'Tuesday', 'Mathematics', '08:30 AM - 11:30 AM'],
            ['05-03-2025', 'Wednesday', 'General Science', '08:30 AM - 11:30 AM'],
            ['06-03-2025', 'Thursday', 'Social Studies', '08:30 AM - 11:30 AM'],
            ['07-03-2025', 'Friday', 'Islamiyat', '08:30 AM - 11:30 AM'],
            ['08-03-2025', 'Saturday', 'M.T', '08:30 AM - 11:30 AM'],
            ['10-03-2025', 'Monday', 'N.M.T', '08:30 AM - 11:30 AM'],
            ['12-03-2025', 'Wednesday', 'Drawing', '08:30 AM - 11:30 AM'],
            ['13-03-2025', 'Thursday', 'Arabic', '08:30 AM - 11:30 AM'],
        ]
    }

    return schedules.get(class_name, [])

def generate_pdf(gr_no, name, father_name, caste, dob, roll_number, class_name, profile_path):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Add Logo
    logo_path = "static/logo.jpg"
    if os.path.exists(logo_path):
        pdf.image(logo_path, 10, 8, 22)

    # Heading: "GOVERNMENT HIGHER SECONDARY SCHOOL MIRPUR"
    pdf.set_font("Arial", "B", 14)  # Arial font, bold, large size
    pdf.set_xy(15, 10)  # Align heading with logo
    pdf.set_text_color(43, 191, 110)
    pdf.cell(0, 5, "GOVERNMENT HIGHER SECONDARY SCHOOL MIRPUR", 0, 1, 'C')

    # Heading: "ANNUAL EXAMINATION 2024-25"
    pdf.set_font("Arial", "B", 12)  # Arial font, bold, smaller size
    pdf.set_xy(0, 17)
    pdf.set_text_color(142, 54, 0)
    pdf.cell(0, 5, "ANNUAL EXAMINATION 2024-25", 0, 1, 'C')

    # Heading: "ROLL NUMBER SLIP"
    pdf.set_font("Arial", "B", 12)  # Arial font, bold, smaller size
    pdf.set_xy(0, 24)
    pdf.set_text_color(0, 0, 139)
    pdf.cell(0, 5, "ROLL NUMBER SLIP", 0, 1, 'C')

    # GR No
    pdf.set_font("Courier", "B", 14)
    pdf.set_text_color(255, 0, 0)
    pdf.set_xy(153, 45)
    pdf.cell(40, 10, f"GR NO: {gr_no}", 0, 1, 'R')

    # Student Information
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(139, 69, 19)
    student_info = [
        ("NAME:", name),
        ("FATHER NAME:", father_name),
        ("CASTE:", caste),
        ("DOB:", datetime.strptime(dob, "%Y-%m-%d").strftime("%d/%m/%Y")),
        ("CLASS:", f"{class_name}"),
        ("ROLL NUMBER:", roll_number),
        ("INSTITUTE:", "GOVERNMENT HIGHER SECONDARY SCHOOL MIRPUR SAKRO"),
        ("CENTER:", "GOVERNMENT HIGHER SECONDARY SCHOOL MIRPUR SAKRO")
    ]
    for label, value in student_info:
        pdf.cell(30, 5, label, 0, 0)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(50, 5, value, 0, 1)
        pdf.set_text_color(139, 69, 19)

    pdf.ln(3)

    # Profile Picture
    if os.path.exists(profile_path):
        x, y, w, h = 160, 55, 30, 36
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x - 2, y - 2, w + 4, h + 4)
        pdf.image(profile_path, x, y, w, h)                   
    pdf.ln(10)

    # Exam Schedule Table
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(190, 225, 174)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(35, 8, "Date", 1, 0, 'C', True)
    pdf.cell(35, 8, "Day", 1, 0, 'C', True)
    pdf.cell(55, 8, "Subject", 1, 0, 'C', True)
    pdf.cell(57, 8, "Morning Timing", 1, 1, 'C', True)
    # Schedule Data
    schedule = get_schedule(class_name)
    pdf.set_font("Arial", "", 10)
    for row in schedule:
        pdf.cell(35, 7, row[0], 1, 0, 'C')
        pdf.cell(35, 7, row[1], 1, 0, 'C')
        pdf.cell(55, 7, row[2], 1, 0, 'C')
        pdf.cell(57, 7, row[3], 1, 1, 'C')

    pdf.ln(25)



    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(0, 0, 0)
    
    candidate_signature_y = pdf.get_y()
    pdf.set_xy(10, candidate_signature_y)
    pdf.cell(0, 4, "CANDIDATE'S SIGNATURE:", ln=True)
    
    # Get width of the "CANDIDATE'S SIGNATURE:" text
    text_width = pdf.get_string_width("CANDIDATE'S SIGNATURE:")
    
    # Add space between the text and the line
    pdf.ln(1)
    
    signature_line_y = pdf.get_y()
    pdf.set_xy(10, signature_line_y)
    pdf.cell(text_width, 0, "", ln=True, border='B')  # Set line width based on text width
    
    center_incharge_signature_y = candidate_signature_y
    pdf.set_xy(130, center_incharge_signature_y)
    pdf.cell(0, 4, "CENTER INCHARGE'S SIGNATURE:", ln=True)
    
    # Get width of the "CENTER INCHARGE'S SIGNATURE:" text
    text_width = pdf.get_string_width("CENTER INCHARGE'S SIGNATURE:")
    
    pdf.ln(1)
    
    signature_line_y = pdf.get_y()
    pdf.set_xy(130, signature_line_y)
    pdf.cell(text_width, 0, "", ln=True, border='B')  # Set line width based on text width
    
    # Adding some space before Instructions
    pdf.ln(20)  # Adjust this value as per your need for spacing

    # Instructions Title (Red and Bold)
    pdf.set_font("Arial", "B", 10)  # Bold font
    pdf.set_text_color(255, 0, 0)  # Red color for title
    pdf.cell(0, 6, "INSTRUCTIONS:", ln=True)
    
    # Instructions Text (Bold)
    pdf.set_font("Arial", "B", 10)  # Bold font for instructions
    pdf.set_text_color(0, 0, 0)  # Black color for instructions text
    
    instructions = [
        "1) Please ensure that you bring this slip with you to the examination hall.",
        "2) The possession of unauthorized materials in the examination hall is strictly prohibited.",
        "3) You are required to bring your own writing materials to the examination.",
        "4) The examination timings must be adhered to without exception."
    ]
    
    # Loop through instructions and add them to the PDF
    for instr in instructions:
        pdf.cell(0, 6, instr, ln=True)

    # Save PDF
    pdf_folder = get_pdf_folder(class_name)
    pdf_filename = f"{roll_number}_roll_number_slip.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_filename)
    pdf.output(pdf_path)

    return pdf_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_slip', methods=['POST'])
def generate_slip():
    """Handles form submission, saves data, and generates a PDF slip."""
    gr_no = request.form['grNo'].upper()
    name = request.form['name'].upper()
    father_name = request.form['fatherName'].upper()
    caste = request.form['caste'].upper()
    dob = request.form['dob']
    roll_number = request.form['rollNumber'].upper()
    class_name = request.form['class']
    profile_picture = request.files['profilePicture']         
    profile_path = os.path.join(UPLOAD_FOLDER, f"{gr_no}.jpg")
    profile_picture.save(profile_path)

    save_to_excel({
        "GR No": gr_no,
        "Name": name,
        "Father Name": father_name,
        "Caste": caste,
        "DOB": dob,
        "Roll Number": roll_number,
        "Profile Picture": profile_path,
        "Class": class_name
    }, class_name)

    generate_pdf(gr_no, name, father_name, caste, dob, roll_number, class_name, profile_path)
    return jsonify({"message": "Roll Number Slip Generated Successfully!"})

if __name__ == '__main__':
    app.run(debug=True)
