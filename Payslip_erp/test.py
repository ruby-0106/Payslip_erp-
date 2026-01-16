import tkinter as tk
from tkinter import messagebox, ttk
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

# Import reportlab libraries for PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

from PIL import Image, ImageTk #added line


# ----------------------
# Theme
# ----------------------
BG_COLOR = "#FFFFFF"      # White
ACCENT_COLOR = "#314B9F"  # Dark Blue
TEXT_COLOR = "#000000"    # Black
BUTTON_TEXT_COLOR = "#FFFFFF" # White

# --- สร้าง Path ที่ถูกต้องไปยังไฟล์ต่างๆ ---
# script_dir คือโฟลเดอร์ที่ไฟล์ .py นี้ถูกรัน
script_dir = os.path.dirname(os.path.abspath(__file__))
#ICON_PATH = os.path.join(script_dir, "EITHeader.png") -original line
ICON_PATH = os.path.join(script_dir, "EIT Lasertechnik.png") #added line
EMPLOYEE_FILE = os.path.join(script_dir, 'employees.json')
SALARY_FILE = os.path.join(script_dir, 'salaries.json')
FONT_PATH = os.path.join(script_dir, 'Prompt-Regular.ttf')
FONT_BOLD_PATH = os.path.join(script_dir, 'Prompt-Bold.ttf')
EXCEL_DIR = os.path.join(script_dir, "excel_files")
LICEN_IMAGE_PATH = os.path.join(script_dir,'licen.jpg')

# --- ลงทะเบียนฟอนต์ (สำหรับ PDF) ---
try:
    pdfmetrics.registerFont(TTFont('Prompt-Regular', FONT_PATH))
    pdfmetrics.registerFont(TTFont('Prompt-Bold', FONT_BOLD_PATH))
    font_name = "Prompt-Regular"
    font_name_bold = "Prompt-Bold"
except Exception as e:
    print(f"ไม่สามารถลงทะเบียนฟอนต์ Prompt ได้ (จำเป็นต้องมีไฟล์ .ttf): {e}")
    font_name = "Helvetica"
    font_name_bold = "Helvetica-Bold"

# --- การตั้งค่า Email (สำคัญ: ต้องแก้ไขเป็นข้อมูลจริง) ---
SENDER_EMAIL = "eit@eitlaser.com"  # <--- ใส่อีเมลผู้ส่ง
SENDER_PASSWORD = "grsc gthh jnuy ixtc" # <--- ใส่ App Password ที่สร้างจาก Google
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ----------------------
# Helper Functions
# ----------------------
def center_window(window, width, height):
    """จัดหน้าต่างให้อยู่กึ่งกลางจอ"""
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
    window.configure(bg=BG_COLOR)

# ----------------------
# JSON File Handling (อ่าน/เขียน ข้อมูลพนักงานและเงินเดือน)
# ----------------------
def load_employees():
    """โหลดรายชื่อพนักงานจาก employees.json"""
    if os.path.exists(EMPLOYEE_FILE):
        try:
            with open(EMPLOYEE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            return [] # ถ้าไฟล์เสียหรือว่างเปล่า
    return [] # ถ้าไม่มีไฟล์

def save_employees_to_file(employees):
    """บันทึกข้อมูลพนักงานลง employees.json"""
    with open(EMPLOYEE_FILE, 'w', encoding='utf-8') as f:
        json.dump(employees, f, indent=2, ensure_ascii=False)

def load_salaries():
    """โหลดข้อมูลเงินเดือนจาก salaries.json"""
    if os.path.exists(SALARY_FILE):
        try:
            with open(SALARY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return [entry for entry in data if isinstance(entry, dict)]
        except json.JSONDecodeError:
            return []
    return []


def save_salaries_to_file(salaries):
    """บันทึกข้อมูลเงินเดือนลง salaries.json"""
    with open(SALARY_FILE, 'w', encoding='utf-8') as f:
        json.dump(salaries, f, indent=2, ensure_ascii=False)

# ----------------------
# PDF Generation
# ----------------------
def create_pay_slip_pdf(employee_data, salary_data):
    """สร้างไฟล์ PDF ใบสลิปเงินเดือน"""
    filename = f"PaySlip_{employee_data['name'].replace(' ', '_')}_{salary_data['date']}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=70, leftMargin=70, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    # ตั้งค่าสไตล์ให้ใช้ฟอนต์ภาษาไทย
    styles.add(ParagraphStyle(name='Normal_Thai', parent=styles['Normal'], fontName=font_name, fontSize=12))
    styles.add(ParagraphStyle(name='Heading1_Thai', parent=styles['Heading1'], fontName=font_name_bold, fontSize=16))
    styles.add(ParagraphStyle(name='Center_Thai', parent=styles['Normal_Thai'], alignment=TA_CENTER))

    # --- Header and Employee Info ---
    if os.path.exists(ICON_PATH):
        elements.append(Image(ICON_PATH, width=575, height=100))
    elements.append(Paragraph("<b>Pay Slip / ใบสรุปเงินเดือน</b>", styles['Heading1_Thai']))
    elements.append(Paragraph(f"Date: {datetime.date.today().strftime('%d-%b-%Y')}", styles['Normal_Thai']))
    elements.append(Spacer(1, 20))
    
    # --- ตารางข้อมูลพนักงาน ---
    emp_table_data = [
        [Paragraph("<b>ชื่อ-นามสกุล พนักงาน :</b>", styles['Normal_Thai']),
         Paragraph(employee_data['name'], styles['Normal_Thai']),
         Paragraph("<b>รอบเดือน :</b>", styles['Normal_Thai']),
         Paragraph(salary_data['date'], styles['Normal_Thai'])],
        [Paragraph("<b>ตำแหน่ง :</b>", styles['Normal_Thai']),
         Paragraph(employee_data['position'], styles['Normal_Thai']),
         Paragraph("<b>เลขประจำตัวพนักงาน :</b>", styles['Normal_Thai']),
         Paragraph(employee_data.get('employee_code', ''), styles['Normal_Thai'])],
    ]
    emp_table = Table(emp_table_data, colWidths=[120, 200, 150, 100])
    emp_table.setStyle(TableStyle([('ALIGN', (1,0), (1,1), 'LEFT'), ('ALIGN', (3,0), (3,1), 'LEFT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    elements.append(emp_table)
    elements.append(Spacer(1, 20))

    # --- Calculation and Display ---
    income = salary_data.get('income', {})
    deduction = salary_data.get('deduction', {})
    collection_income = salary_data.get('collection_income', 0)
    collection_social_wht = salary_data.get('collection_social_wht', 0)

    income_items = list(income.items())
    deduct_items = list(deduction.items())
    total_income = sum(income.values())
    total_deduct = sum(deduction.values())
    net_income = total_income - total_deduct
    max_rows = max(len(income_items), len(deduct_items))

    translation_map = {
        "salary": "เงินเดือน", "overtime": "ค่าล่วงเวลา", "living": "ค่าครองชีพ/เบี้ยขยัน",
        "commission": "ค่าคอมมิชชั่น", "other_income": "รายรับอื่นๆ", "tax": "ภาษี",
        "social_security": "ประกันสังคม", "other_deduct": "หักอื่นๆ"
    }
    
    # --- ตารางสรุปรายรับ-รายจ่าย (ใช้ Paragraph เพื่อรองรับภาษาไทย) ---
    header_style = styles['Normal_Thai']
    data_style = styles['Normal_Thai']
    
    thai_table_data = [[ 
        Paragraph("<b>รายรับ</b>", header_style), 
        Paragraph("<b>จำนวนเงิน</b>", header_style), 
        Paragraph("<b>รายการหัก</b>", header_style), 
        Paragraph("<b>จำนวนเงิน</b>", header_style) 
    ]]
    
    for i in range(max_rows):
        row = []
        if i < len(income_items):
            key, value = income_items[i]
            row.extend([Paragraph(translation_map.get(key, key.title()), data_style), Paragraph(f"{value:,.2f}", data_style)])
        else:
            row.extend(["", ""])
        if i < len(deduct_items):
            key, value = deduct_items[i]
            row.extend([Paragraph(translation_map.get(key, key.title()), data_style), Paragraph(f"{value:,.2f}", data_style)])
        else:
            row.extend(["", ""])
        thai_table_data.append(row)
    
    # --- ยอดรวม ---
    thai_table_data.append([
        Paragraph("<b>รายรับรวม</b>", header_style), Paragraph(f"<b>{total_income:,.2f}</b>", data_style), 
        Paragraph("<b>ยอดหักรวม</b>", header_style), Paragraph(f"<b>{total_deduct:,.2f}</b>", data_style)
    ])
    thai_table_data.append([
        Paragraph("<b>รายได้สะสม</b>", header_style), Paragraph(f"<b>{collection_income:,.2f}</b>", data_style), 
        Paragraph("<b>ประกันสังคม-สะสม</b>", header_style), Paragraph(f"<b>{collection_social_wht:,.2f}</b>", data_style)
    ])
    thai_table_data.append([
        "", "", 
        Paragraph("<b>รายรับสุทธิ</b>", header_style), Paragraph(f"<b>{net_income:,.2f}</b>", data_style)
    ])

    thai_table = Table(thai_table_data, colWidths=[150, 100, 150, 100])
    thai_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,1), (1,-1), 'RIGHT'), ('ALIGN', (3,1), (3,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(thai_table)
    elements.append(Spacer(1, 20)) 
    # ---------------------------------
# --- START: English Table (เพิ่มส่วนนี้เข้าไป) ---
# ---------------------------------

    # เพิ่มหัวข้อสำหรับตารางภาษาอังกฤษ
    elements.append(Paragraph("<b>Pay Slip Summary</b>", styles['Heading1_Thai']))
    elements.append(Spacer(1, 10))

    # สร้างข้อมูลสำหรับตารางภาษาอังกฤษ
    eng_table_data = [[ 
        Paragraph("<b>Description</b>", header_style), 
        Paragraph("<b>Amount</b>", header_style), 
        Paragraph("<b>Deduction</b>", header_style), 
        Paragraph("<b>Amount</b>", header_style) 
    ]]

    # วนลูปข้อมูลรายรับ/รายหัก (เหมือนเดิม)
    for i in range(max_rows):
        row = []
        if i < len(income_items):
            key, value = income_items[i]
            # แปลง key (เช่น 'other_income') เป็น 'Other Income'
            desc = key.replace('_', ' ').title()
            row.extend([Paragraph(desc, data_style), Paragraph(f"{value:,.2f}", data_style)])
        else:
            row.extend(["", ""])

        if i < len(deduct_items):
            key, value = deduct_items[i]
            desc = key.replace('_', ' ').title()
            row.extend([Paragraph(desc, data_style), Paragraph(f"{value:,.2f}", data_style)])
        else:
            row.extend(["", ""])
        eng_table_data.append(row)

    # --- แถวสรุปยอดภาษาอังกฤษ ---
    eng_table_data.append([
        Paragraph("<b>Total Income</b>", header_style), Paragraph(f"<b>{total_income:,.2f}</b>", data_style), 
        Paragraph("<b>Total Deduction</b>", header_style), Paragraph(f"<b>{total_deduct:,.2f}</b>", data_style)
    ])
    # เพิ่มยอดสะสม (เหมือนตารางไทย)
    eng_table_data.append([
        Paragraph("<b>Collection Income</b>", header_style), Paragraph(f"<b>{collection_income:,.2f}</b>", data_style), 
        Paragraph("<b>Collection Social WHT</b>", header_style), Paragraph(f"<b>{collection_social_wht:,.2f}</b>", data_style)
    ])
    # ยอดสุทธิ
    eng_table_data.append([
        "", "", 
        Paragraph("<b>Net Income</b>", header_style), Paragraph(f"<b>{net_income:,.2f}</b>", data_style)
    ])

    # สร้างตารางและใส่สไตล์ (เหมือนตารางไทย)
    eng_table = Table(eng_table_data, colWidths=[150, 100, 150, 100])
    eng_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), 
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (1,1), (1,-1), 'RIGHT'), 
        ('ALIGN', (3,1), (3,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))

    # เพิ่มตารางภาษาอังกฤษลงใน PDF
    elements.append(eng_table)

    # ---------------------------------
    # --- END: English Table ---
    # ---------------------------------
    
    # --- Signature Block ---
    elements.append(Spacer(1, 60))
    signature_img = Image(LICEN_IMAGE_PATH, width=150, height=50) if os.path.exists(LICEN_IMAGE_PATH) else Paragraph("...........................................", styles['Normal_Thai'])
    
    line1_table = Table([[Paragraph("ลงชื่อ", styles['Normal_Thai']), signature_img]], colWidths=[50, 250])
    line1_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    
    final_signature_table = Table([[line1_table], [Paragraph("( รวีวรรณ งอยภูธร )", styles['Center_Thai'])], [Paragraph("ฝ่ายบุคคล", styles['Center_Thai'])]], colWidths=[300])
    final_signature_table.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    final_signature_table.hAlign = 'CENTER'
    elements.append(final_signature_table)

    doc.build(elements)
    return filename

# ----------------------
# Email Sending
# ----------------------
def send_email_with_attachment(file_path, recipient_email, employee_name):
    """ส่งอีเมลพร้อมไฟล์ PDF ที่แนบไป"""
    if not SENDER_EMAIL or SENDER_EMAIL == "your_email@gmail.com" or not SENDER_PASSWORD or SENDER_PASSWORD == "your_app_password":
        messagebox.showerror("ตั้งค่าอีเมลผิดพลาด", "กรุณาแก้ไขข้อมูล SENDER_EMAIL และ SENDER_PASSWORD (App Password) ในโค้ดก่อนทำการส่งอีเมล ❌")
        return False
        
    try:
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient_email
        msg["Subject"] = f"Pay Slip (สลิปเงินเดือน) - {employee_name}"
        
        body = f"เรียน {employee_name}\n\n"
        body += "บริษัทฯ ขอส่งสลิปเงินเดือนของท่าน ดังรายละเอียดในไฟล์แนบ\n"
        body += "กรุณาตรวจสอบรายละเอียด\n\n"
        body += "ขอแสดงความนับถือ,\n"
        body += "ฝ่ายบุคคล"
        
        msg.attach(MIMEText(body, "plain", _charset="utf-8"))

        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        messagebox.showerror("ส่งอีเมลล้มเหลว", f"เกิดข้อผิดพลาด: {e}\n(อาจเกิดจาก App Password ไม่ถูกต้อง หรือ Google/Antivirus บล็อก)")
        return False
    finally:
        # ลบไฟล์ PDF ชั่วคราวหลังส่งเสร็จ
        if os.path.exists(file_path):
            os.remove(file_path)

# ----------------------
# Login and Main Windows
# ----------------------
def login():
    """ตรวจสอบการล็อกอิน"""
    username = entry_user.get()
    password = entry_pass.get()
    if username == "eit@eitlaser.com" and password == "payslip282895":
        messagebox.showinfo("Login", "เข้าสู่ระบบสำเร็จ ✅ (Admin)")
        root.withdraw() # ซ่อนหน้าต่าง Login
        open_Admin_main_window()
    elif username == "User" and password == "7894":
        messagebox.showinfo("Login", "เข้าสู่ระบบสำเร็จ ✅ (User)")
        root.withdraw() # ซ่อนหน้าต่าง Login
        open_User_main_window()
    else:   
        messagebox.showerror("Login", "❌ ชื่อผู้ใช้หรือรหัสผ่านผิด")

def open_Admin_main_window():
    """เปิดหน้าต่างหลักสำหรับ Admin"""
    main_win = tk.Toplevel(bg=BG_COLOR)
    main_win.title("Admin Main System")
    center_window(main_win, 400, 300)
    
    tk.Label(main_win, text="ยินดีต้อนรับสู่ระบบหลังบ้าน Admin", font=("Arial", 12), bg=BG_COLOR).pack(pady=20)
    
    tk.Button(main_win, text="จัดการพนักงาน (Add/Delete)", command=lambda: add_delete_employee_window(main_win), bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
    
    # --- MODIFIED: ปุ่มนี้จะเปิดหน้าจัดการเงินเดือนที่รวมหน้าจอแล้ว ---
    tk.Button(main_win, text="จัดการเงินเดือน (Salary Management)", command=lambda: create_salary_management_window(main_win), bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)

    tk.Button(main_win, text="Logout", command=lambda: [main_win.destroy(), root.deiconify()], bg="#FF6347", fg=BUTTON_TEXT_COLOR).pack(pady=20, ipadx=10, ipady=5)
    
    # เมื่อปิดหน้าต่างหลัก ให้กลับไปหน้า Login
    main_win.protocol("WM_DELETE_WINDOW", lambda: [main_win.destroy(), root.deiconify()])


def open_User_main_window():
    """เปิดหน้าต่างหลักสำหรับ User (ตัวอย่าง)"""
    main_win = tk.Toplevel(bg=BG_COLOR)
    main_win.title("User Main System")
    center_window(main_win, 400, 300)
    tk.Label(main_win, text="ยินดีต้อนรับสู่ระบบหลังบ้าน User", font=("Arial", 12), bg=BG_COLOR).pack(pady=50)
    
    # เมื่อปิดหน้าต่างหลัก ให้กลับไปหน้า Login
    main_win.protocol("WM_DELETE_WINDOW", lambda: [main_win.destroy(), root.deiconify()])

# ----------------------
# Employee Management (เพิ่ม/ลบ พนักงาน)
# ----------------------
def add_delete_employee_window(parent_win):
    """หน้าต่างสำหรับเลือก เพิ่ม หรือ ลบ พนักงาน"""
    parent_win.withdraw() # ซ่อนหน้าต่างก่อนหน้า
    add_delete_win = tk.Toplevel(bg=BG_COLOR)
    add_delete_win.title("Employee Management")
    center_window(add_delete_win, 400, 250)
    
    tk.Label(add_delete_win, text="จัดการข้อมูลพนักงาน", font=("Arial", 12), bg=BG_COLOR).pack(pady=10)
    tk.Button(add_delete_win, text="Add Employee", command=lambda: add_employee_form(add_delete_win), bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
    tk.Button(add_delete_win, text="Delete Employee", command=lambda: delete_employee(add_delete_win), bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
    tk.Button(add_delete_win, text="กลับ", command=lambda: [add_delete_win.destroy(), parent_win.deiconify()], bg="#FF6347", fg=BUTTON_TEXT_COLOR).pack(pady=20, ipadx=10, ipady=5)
    
    add_delete_win.protocol("WM_DELETE_WINDOW", lambda: [add_delete_win.destroy(), parent_win.deiconify()])


def add_employee_form(parent_win):
    """ฟอร์มสำหรับกรอกข้อมูลพนักงานใหม่"""
    parent_win.withdraw()
    add_form_win = tk.Toplevel(bg=BG_COLOR)
    add_form_win.title("Add Employee")
    center_window(add_form_win, 450, 450)
    
    tk.Label(add_form_win, text="กรอกข้อมูลพนักงาน", font=("Arial", 14, "bold"), bg=BG_COLOR).pack(pady=10)
    
    fields = {
        "ชื่อพนักงาน (คำนำหน้าชื่อ ชื่อ นามสกุล)": "name",
        "เลขประจำตัวพนักงาน:": "employee_code",
        "เลขประจำตัวประชาชน:": "id",
        "ตำแหน่ง (Position):": "position",
        "อีเมล (สำหรับรับสลิป):": "email",
        "วันที่เริ่มงาน (DD-MM-YYYY):": "start_date"
    }
    
    entries = {}
    for label_text, key in fields.items():
        tk.Label(add_form_win, text=label_text, bg=BG_COLOR).pack(pady=(10,0))
        entry = tk.Entry(add_form_win)
        entry.pack(pady=5, padx=20, fill='x')
        entries[key] = entry
    
    def on_save():
        employee_data = {key: entry.get() for key, entry in entries.items()}
        save_employee(employee_data, add_form_win, parent_win)

    tk.Button(add_form_win, text="บันทึก", command=on_save, bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=15, ipadx=10, ipady=5)
    tk.Button(add_form_win, text="กลับ", command=lambda: [add_form_win.destroy(), parent_win.deiconify()], bg="#FF6347", fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
    
    add_form_win.protocol("WM_DELETE_WINDOW", lambda: [add_form_win.destroy(), parent_win.deiconify()])


def save_employee(data, window, parent_win):
    """บันทึกข้อมูลพนักงานลง JSON"""
    if not all(data.values()):
        messagebox.showerror("ผิดพลาด", "กรุณากรอกข้อมูลให้ครบถ้วน ❌")
        return

    employees = load_employees()
    if any(emp["id"] == data["id"] for emp in employees):
        messagebox.showerror("ผิดพลาด", "เลขประจำตัวประชาชนนี้มีอยู่ในระบบแล้ว ❌")
        return
    if any(emp.get("employee_code") == data["employee_code"] for emp in employees):
        messagebox.showerror("ผิดพลาด", "เลขประจำตัวพนักงานนี้มีอยู่ในระบบแล้ว ❌")
        return

    employees.append(data)
    save_employees_to_file(employees)
    messagebox.showinfo("บันทึกสำเร็จ", "ข้อมูลพนักงานถูกบันทึกเรียบร้อยแล้ว ✅")
    window.destroy()
    parent_win.deiconify()

def delete_employee(parent_win):
    """หน้าต่างสำหรับเลือกลบพนักงาน"""
    parent_win.withdraw()
    employees = load_employees()
    if not employees:
        messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีข้อมูลพนักงานที่สามารถลบได้")
        parent_win.deiconify()
        return

    delete_win = tk.Toplevel(bg=BG_COLOR)
    delete_win.title("ลบพนักงาน")
    center_window(delete_win, 400, 200)

    tk.Label(delete_win, text="เลือกพนักงานที่ต้องการลบ:", font=("Arial", 12), bg=BG_COLOR).pack(pady=10)
    
    # ดึงรายชื่อมาแสดงใน combobox
    employee_names = sorted([f'{emp["name"]} ({emp["id"]})' for emp in employees])
    emp_combo = ttk.Combobox(delete_win, values=employee_names, state="readonly")
    emp_combo.pack(pady=5, padx=20, fill='x')
    if employee_names: emp_combo.set(employee_names[0])

    def confirm_delete():
        selected_text = emp_combo.get()
        if not selected_text: return
        
        # ยืนยันก่อนลบ
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบ '{selected_text.split(' (')[0]}' ใช่หรือไม่?"):
            return
            
        id_to_delete = selected_text.split(" (")[1].replace(")", "")
        
        employees_to_keep = [emp for emp in employees if emp["id"] != id_to_delete]
        save_employees_to_file(employees_to_keep)
        messagebox.showinfo("ลบสำเร็จ", f"ข้อมูลของ '{selected_text.split(' (')[0]}' ถูกลบเรียบร้อยแล้ว ✅")
        delete_win.destroy()
        parent_win.deiconify()

    tk.Button(delete_win, text="ยืนยันการลบ", command=confirm_delete, bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(pady=10, ipadx=10, ipady=5)
    tk.Button(delete_win, text="กลับ", command=lambda: [delete_win.destroy(), parent_win.deiconify()], bg="#FF6347", fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
    
    delete_win.protocol("WM_DELETE_WINDOW", lambda: [delete_win.destroy(), parent_win.deiconify()])


# ----------------------------------------------------
# NEW: Unified Salary Management Window (หน้าจอรวม)
# ----------------------------------------------------
def create_salary_management_window(parent_win):
    """สร้างหน้าต่างจัดการเงินเดือน (ที่รวมหน้ากรอกและสรุปไว้ด้วยกัน)"""
    parent_win.withdraw()
    salary_win = tk.Toplevel(bg=BG_COLOR)
    salary_win.title("Salary Management (จัดการเงินเดือน)")
    salary_win.geometry("1100x700")

    # --- Data Variables (ตัวแปรสำหรับเก็บค่าในช่อง Entry) ---
    string_vars = {
        "salary": tk.StringVar(value="0"), "overtime": tk.StringVar(value="0"),
        "living": tk.StringVar(value="0"), "commission": tk.StringVar(value="0"),
        "other_income": tk.StringVar(value="0"), "tax": tk.StringVar(value="0"),
        "social_security": tk.StringVar(value="0"), "other_deduct": tk.StringVar(value="0"),
        "collection_income": tk.StringVar(value="0"), "collection_social_wht": tk.StringVar(value="0")
    }

    # --- Main Layout (แบ่งหน้าจอ) ---
    top_frame = tk.Frame(salary_win, bg=BG_COLOR)
    top_frame.pack(pady=10, padx=20, fill='x')
    
    main_content_frame = tk.Frame(salary_win, bg=BG_COLOR)
    main_content_frame.pack(pady=10, padx=20, fill='both', expand=True)

    # --- Left Side (Data Entry) ---
    entry_frame = tk.LabelFrame(main_content_frame, text="ข้อมูลเงินเดือน (Data Entry)", padx=15, pady=15, bg=BG_COLOR)
    entry_frame.grid(row=0, column=0, padx=10, sticky="nsew")

    # --- Right Side (Summary) ---
    summary_frame = tk.LabelFrame(main_content_frame, text="สรุปยอด (Summary)", padx=15, pady=15, bg=BG_COLOR)
    summary_frame.grid(row=0, column=1, padx=10, sticky="nsew")
    
    main_content_frame.grid_columnconfigure(0, weight=1)
    main_content_frame.grid_columnconfigure(1, weight=1)

    # --- Top Frame Widgets (Selectors) ---
    tk.Label(top_frame, text="เลือกพนักงาน:", bg=BG_COLOR).pack(side=tk.LEFT, padx=(0, 5))
    emp_combo = ttk.Combobox(top_frame, state="readonly", width=30)
    emp_combo.pack(side=tk.LEFT, padx=5)

    tk.Label(top_frame, text="เลือกรอบเดือน (จาก Excel):", bg=BG_COLOR).pack(side=tk.LEFT, padx=(10, 5))
    month_combo = ttk.Combobox(top_frame, state="readonly", width=20)
    month_combo.pack(side=tk.LEFT, padx=5)
    
    # --- Helper: Sort Excel Sheets Chronologically (FIXED) ---
    def sort_sheets_chronologically(sheet_names):
        """
        ฟังก์ชันสำหรับเรียงชื่อ Sheet เช่น '62-Jan', '64-Feb' ตามลำดับเวลา
        """
        month_map = {
            'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
            'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
        }
        def get_sort_key(sheet_name):
            try:
                year_str, month_str = sheet_name.split('-')
                year = int(year_str) + 2000 # (เช่น 62 -> 2062, หรือปรับตามปี พ.ศ./ค.ศ. ที่ใช้)
                # ถ้าปีเป็น พ.ศ. เช่น 67 (2567) ให้ใช้
                # year = int(year_str) 
                month = month_map[month_str]
                return (year, month)
            except (ValueError, KeyError):
                # ถ้าชื่อ Sheet ไม่ตรงฟอร์แมต ให้ไปอยู่ท้ายสุด
                return (9999, 99) 
        
        return sorted(sheet_names, key=get_sort_key)

    # --- Helper: Import Data ---
    def import_excel_data(event=None):
        """
        ดึงข้อมูลจากไฟล์ Excel ตามพนักงานและ Sheet ที่เลือก
        และอัปเดตค่าในช่อง Entry (string_vars)
        """
        emp_name = emp_combo.get()
        sheet_name = month_combo.get()
        if not emp_name or not sheet_name:
            # ไม่ต้องแสดง Error ถ้าแค่เลือกพนักงานแต่ยังไม่มี Sheet
            return

        excel_path = os.path.join(EXCEL_DIR, f"{emp_name}_FormSlip.xlsx")
        if not os.path.exists(excel_path):
            messagebox.showerror("ผิดพลาด", f"ไม่พบไฟล์ Excel: {os.path.basename(excel_path)}\n\n(ไฟล์ต้องอยู่ในโฟลเดอร์ 'excel_files')")
            return
        
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
            # ตำแหน่ง Cell ใน Excel (แถว, คอลัมน์) - Index เริ่มที่ 0
            data_map = {
                'salary': (15, 1), 'overtime': (15, 2), 'living': (15, 3),
                'commission': (15, 4), 'other_income': (15, 5), 'tax': (19, 1),
                'social_security': (19, 2), 'collection_income': (23, 1),
                'collection_social_wht': (23, 2)
            }
            # Reset all fields
            for key in string_vars:
                string_vars[key].set("0")
            
            # Populate from Excel
            for key, (row, col) in data_map.items():
                value = df.iat[row, col]
                string_vars[key].set(str(value) if pd.notna(value) else "0")

            # 'หักอื่นๆ' อาจมาจากหลาย Cell
            other_deduct1 = float(df.iat[19, 3]) if pd.notna(df.iat[19, 3]) else 0
            other_deduct2 = float(df.iat[19, 4]) if pd.notna(df.iat[19, 4]) else 0
            string_vars['other_deduct'].set(str(other_deduct1 + other_deduct2))
            
            # อัปเดตสรุปยอดทันที
            update_summary()
            
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"อ่านไฟล์ Excel ล้มเหลว: {e}\n\n(อาจเกิดจากโครงสร้างไฟล์ Excel หรือชื่อ Sheet ไม่ถูกต้อง)")

    tk.Button(top_frame, text="Import Excel Data", command=import_excel_data, bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR).pack(side=tk.LEFT, padx=10, ipady=4)

    # --- Entry Frame Widgets (Left Side) ---
    income_frame = tk.LabelFrame(entry_frame, text="รายรับ (Income)", padx=10, pady=10, bg=BG_COLOR)
    income_frame.pack(fill='x', expand=True, pady=5)
    deduction_frame = tk.LabelFrame(entry_frame, text="รายการหัก (Deductions)", padx=10, pady=10, bg=BG_COLOR)
    deduction_frame.pack(fill='x', expand=True, pady=5)
    collection_frame = tk.LabelFrame(entry_frame, text="ยอดสะสม (Collections)", padx=10, pady=10, bg=BG_COLOR)
    collection_frame.pack(fill='x', expand=True, pady=5)

    income_fields = [("เงินเดือน:", "salary"), ("ค่าล่วงเวลา:", "overtime"), ("ค่าครองชีพ/เบี้ยขยัน:", "living"), ("ค่าคอมมิชชั่น:", "commission"), ("รายรับอื่นๆ:", "other_income")]
    deduction_fields = [("ภาษี:", "tax"), ("ประกันสังคม:", "social_security"), ("หักอื่นๆ:", "other_deduct")]
    collection_fields = [("รายได้สะสม:", "collection_income"), ("ประกันสังคม-สะสม:", "collection_social_wht")]

    for i, (text, key) in enumerate(income_fields):
        tk.Label(income_frame, text=text, bg=BG_COLOR).grid(row=i, column=0, sticky="w", pady=2)
        tk.Entry(income_frame, textvariable=string_vars[key]).grid(row=i, column=1, pady=2, sticky='ew')
    income_frame.grid_columnconfigure(1, weight=1)

    for i, (text, key) in enumerate(deduction_fields):
        tk.Label(deduction_frame, text=text, bg=BG_COLOR).grid(row=i, column=0, sticky="w", pady=2)
        tk.Entry(deduction_frame, textvariable=string_vars[key]).grid(row=i, column=1, pady=2, sticky='ew')
    deduction_frame.grid_columnconfigure(1, weight=1)
    
    for i, (text, key) in enumerate(collection_fields):
        tk.Label(collection_frame, text=text, bg=BG_COLOR).grid(row=i, column=0, sticky="w", pady=2)
        tk.Entry(collection_frame, textvariable=string_vars[key]).grid(row=i, column=1, pady=2, sticky='ew')
    collection_frame.grid_columnconfigure(1, weight=1)

    # --- Summary Frame Widgets (Right Side) ---
    summary_labels = {}
    summary_fields = ["รายรับรวม (Total Income)", "ยอดหักรวม (Total Deduction)", "รายรับสุทธิ (Net Income)"]
    for i, text in enumerate(summary_fields):
        tk.Label(summary_frame, text=f"{text}:", font=('Arial', 12, 'bold'), bg=BG_COLOR).grid(row=i, column=0, sticky="w", pady=10)
        lbl = tk.Label(summary_frame, text="0.00", font=('Arial', 12), anchor='e', width=20, bg=BG_COLOR)
        lbl.grid(row=i, column=1, sticky="e", pady=10, padx=5)
        summary_labels[text.split(" (")[0]] = lbl
    summary_frame.grid_columnconfigure(1, weight=1)

    def update_summary(*args):
        """
        คำนวณสรุปยอด (ฝั่งขวา) อัตโนมัติ เมื่อข้อมูลฝั่งซ้ายเปลี่ยน
        """
        try:
            total_income = sum(float(string_vars[k].get() or 0) for k in ["salary", "overtime", "living", "commission", "other_income"])
            total_deduction = sum(float(string_vars[k].get() or 0) for k in ["tax", "social_security", "other_deduct"])
            net_income = total_income - total_deduction

            summary_labels["รายรับรวม"].config(text=f"{total_income:,.2f}")
            summary_labels["ยอดหักรวม"].config(text=f"{total_deduction:,.2f}")
            summary_labels["รายรับสุทธิ"].config(text=f"{net_income:,.2f}", fg="blue" if net_income >= 0 else "red")
        except ValueError:
             # Handle case where entry is not a valid number, e.g., empty or contains text
            pass
    
    # ผูก event: เมื่อข้อมูลในช่อง Entry เปลี่ยน ให้เรียก update_summary
    for var in string_vars.values():
        var.trace_add("write", update_summary)
    
    # --- Button Frame & Actions (ด้านล่าง) ---
    button_frame = tk.Frame(salary_win, bg=BG_COLOR)
    button_frame.pack(pady=20)

    def get_current_data():
        """รวบรวมข้อมูลปัจจุบันจาก string_vars"""
        emp_name = emp_combo.get()
        date = month_combo.get()
        if not emp_name or not date:
            messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกพนักงานและรอบเดือน")
            return None, None
        
        try:
            data = {
                "name": emp_name, "date": date,
                "income": {k: float(string_vars[k].get() or 0) for k in ["salary", "overtime", "living", "commission", "other_income"]},
                "deduction": {k: float(string_vars[k].get() or 0) for k in ["tax", "social_security", "other_deduct"]},
                "collection_income": float(string_vars["collection_income"].get() or 0),
                "collection_social_wht": float(string_vars["collection_social_wht"].get() or 0)
            }
            
            employees = load_employees()
            emp_data = next((e for e in employees if e['name'] == emp_name), None)
            
            return data, emp_data
        
        except (ValueError, KeyError) as e:
            messagebox.showerror("ข้อมูลผิดพลาด", f"กรุณากรอกข้อมูลตัวเลขให้ถูกต้อง: {e}")
            return None, None

    def save_salary_data():
        """บันทึกข้อมูลลง salaries.json"""
        salary_data, _ = get_current_data()
        if not salary_data:
            return

        salaries = load_salaries()
        # ลบข้อมูลเก่าของเดือนนี้ (ถ้ามี) แล้วเพิ่มข้อมูลใหม่
        salaries = [r for r in salaries if not (r['name'] == salary_data['name'] and r['date'] == salary_data['date'])]
        salaries.append(salary_data)
        
        save_salaries_to_file(salaries)
        messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลเงินเดือนเรียบร้อยแล้ว ✅")

    def generate_and_send():
        """สร้าง PDF และส่งอีเมล"""
        salary_data, emp_data = get_current_data()
        
        if not salary_data or not emp_data:
            messagebox.showerror("ผิดพลาด", "ไม่พบข้อมูลพนักงานหรือข้อมูลเงินเดือน")
            return

        # ยืนยันก่อนส่ง
        if not messagebox.askyesno("ยืนยันการส่ง", f"คุณต้องการสร้างและส่งสลิปเงินเดือนให้ '{emp_data['name']}' (รอบเดือน {salary_data['date']}) ใช่หรือไม่?"):
            return
            
        # บันทึกข้อมูลล่าสุดก่อนส่ง
        save_salary_data()
        
        recipient_email = emp_data.get("email")
        if not recipient_email:
            messagebox.showerror("ผิดพลาด", f"ไม่พบข้อมูลอีเมลสำหรับ '{emp_data['name']}' ❌\nกรุณาเพิ่มอีเมลในหน้าจัดการพนักงาน")
            return

        try:
            # 1. สร้าง PDF
            pdf_file = create_pay_slip_pdf(emp_data, salary_data)
            
            # 2. ส่ง Email
            if send_email_with_attachment(pdf_file, recipient_email, emp_data['name']):
                messagebox.showinfo("สำเร็จ", f"✅ ส่งใบสรุปเงินเดือนให้ '{emp_data['name']}' ที่อีเมล {recipient_email} เรียบร้อยแล้ว")
            else:
                 messagebox.showerror("ผิดพลาด", "❌ ไม่สามารถส่งอีเมลได้ กรุณาตรวจสอบการตั้งค่า SENDER_EMAIL และ SENDER_PASSWORD (App Password)")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด (PDF/Email)", f"ไม่สามารถสร้างหรือส่ง Pay Slip ได้:\n{e}")

    tk.Button(button_frame, text="Save Data", command=save_salary_data, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)
    tk.Button(button_frame, text="Generate & Send Payslip", command=generate_and_send, bg=ACCENT_COLOR, fg="white").pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)
    tk.Button(button_frame, text="กลับไปหน้าหลัก", command=lambda: [salary_win.destroy(), parent_win.deiconify()], bg="#f44336", fg="white").pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)

    # --- Populate Employee ComboBox and Bind Events ---
    def on_employee_select(event=None):
        """
        เมื่อเลือกพนักงาน:
        1. ค้นหาไฟล์ Excel ของพนักงานคนนั้น
        2. อ่านชื่อ Sheet ทั้งหมด
        3. เรียงลำดับชื่อ Sheet (FIXED)
        4. ใส่ชื่อ Sheet ลงใน month_combo
        5. สั่ง Import ข้อมูลของ Sheet แรกอัตโนมัติ
        """
        emp_name = emp_combo.get()
        if not emp_name: return
        
        # Reset month combo
        month_combo['values'] = []
        month_combo.set('')
        
        excel_path = os.path.join(EXCEL_DIR, f"{emp_name}_FormSlip.xlsx")
        
        if os.path.exists(excel_path):
            try:
                xls = pd.ExcelFile(excel_path)
                # Sort sheets (FIXED)
                sorted_sheets = sort_sheets_chronologically(xls.sheet_names)
                
                month_combo['values'] = sorted_sheets
                if sorted_sheets:
                    month_combo.set(sorted_sheets[0]) # เลือก Sheet แรก (ที่เรียงแล้ว)
                    import_excel_data() # Auto-import for the first sheet
            except Exception as e:
                messagebox.showerror("ผิดพลาด", f"ไม่สามารถอ่าน sheet จากไฟล์ Excel: {e}")
        else:
            # ถ้าไม่พบไฟล์ Excel ของพนักงานคนนี้
            messagebox.showwarning("ไม่พบไฟล์", f"ไม่พบไฟล์ '{emp_name}_FormSlip.xlsx' ในโฟลเดอร์ 'excel_files'")

    # --- โหลดรายชื่อพนักงาน (FIXED) ---
    employees = load_employees()
    if employees:
        # เรียงตามตัวอักษร (Mr., Ms. จะถูกเรียงตามปกติ)
        employee_names = sorted([emp['name'] for emp in employees])
        emp_combo['values'] = employee_names
        
        emp_combo.bind("<<ComboboxSelected>>", on_employee_select)
        month_combo.bind("<<ComboboxSelected>>", import_excel_data) # เมื่อเลือกเดือนใหม่ ให้ import auto
        
        # เลือกคนแรกไว้เป็น default
        if employee_names:
            emp_combo.set(employee_names[0])
            on_employee_select() # Trigger initial load
    else:
        # ถ้าไม่มีพนักงานในระบบเลย
        tk.Label(top_frame, text="ไม่พบข้อมูลพนักงาน! กรุณาเพิ่มพนักงานก่อน", fg="red", bg=BG_COLOR).pack(side=tk.LEFT, padx=10)


    salary_win.protocol("WM_DELETE_WINDOW", lambda: [salary_win.destroy(), parent_win.deiconify()])

# ----------------------
# Main Application Setup (หน้า Login เริ่มต้น)
# ----------------------
root = tk.Tk()
try:
    if os.path.exists(ICON_PATH):
        root.iconbitmap(ICON_PATH)
except tk.TclError:
    print(f"ไม่สามารถโหลดไฟล์ไอคอนได้ที่: {ICON_PATH} (ข้ามไป)")

root.title("EIT Backoffice System")
center_window(root, 1000, 500)
root.configure(bg="#f0f0f0")

# --- Login Widgets ---
if os.path.exists(ICON_PATH):
    #header_img_data = tk.PhotoImage(file=ICON_PATH).subsample(2,2) #original line
    img = Image.open(ICON_PATH) #added line
    img = img.resize((img.width // 2, img.height // 2)) #added line
    header_img_data = ImageTk.PhotoImage(img) #added line
    #header_label = tk.Label(root, image=header_img_data, bg="#f0f0f0")
    #header_label.pack(pady=(20,10))
    header_label = tk.Label(root, image=header_img_data) ##added line
    header_label.image = header_img_data  #added line
    header_label.pack() #added line

else:
    tk.Label(root, text="ยินดีต้อนรับสู่หลังบ้าน EIT!", font=("Arial", 16, "bold"), bg="#f0f0f0").pack(pady=20)

tk.Label(root, text="ชื่อผู้ใช้ (Username):", bg="#f0f0f0").pack(pady=5)
entry_user = tk.Entry(root)
entry_user.pack(pady=5, padx=50, fill='x')

tk.Label(root, text="รหัสผ่าน (Password):", bg="#f0f0f0").pack(pady=5)
entry_pass = tk.Entry(root, show="*")
entry_pass.pack(pady=5, padx=50, fill='x')

btn_login = tk.Button(root, text="Login", command=login, bg=ACCENT_COLOR, fg=BUTTON_TEXT_COLOR, font=("Arial", 12, "bold"))
btn_login.pack(pady=20, ipadx=20, ipady=8)

# เริ่มการทำงานของโปรแกรม
root.mainloop()