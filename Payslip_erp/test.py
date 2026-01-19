import tkinter as tk
from tkinter import messagebox, ttk
import json #data loading
import os #data loading
import glob
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd #data loading

try:
    from PIL import Image as PILImage, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False


# Import reportlab libraries for PDF generation
from reportlab.lib.pagesizes import A4 #PDF generation
from reportlab.pdfgen import canvas #PDF generation
from reportlab.pdfbase import pdfmetrics #PDF generation
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image #PDF generation
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER




# ----------------------
# Theme
# ----------------------
BG_COLOR = "#F5F5F7"
ACCENT_COLOR = "#1F6FEB"
TEXT_COLOR = "#111827"
BUTTON_TEXT_COLOR = "#FFFFFF"
FIELD_BORDER_COLOR = "#D1D5DB"
FORM_BG_COLOR = "#F4F6F8"


# --- สร้าง Path ที่ถูกต้องไปยังไฟล์ต่างๆ ---
# script_dir คือโฟลเดอร์ที่ไฟล์ .py นี้ถูกรัน
script_dir = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(script_dir, "EITHeader.png")
EMP_FORM_IMAGE_PATH = os.path.join(script_dir, "EIT Lasertechnik.png")
EIT_ICON_PATH = os.path.join(script_dir, "EIT Lasertechnik.png")
EINSTEIN_ICON_PATH = os.path.join(script_dir, "Einstein.png")
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
def create_pay_slip_pdf(employee_data, salary_data, company=None):
    """สร้างไฟล์ PDF ใบสลิปเงินเดือน"""
    filename = f"PaySlip_{employee_data['name'].replace(' ', '_')}_{salary_data['date']}.pdf"
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, filename)
    doc = SimpleDocTemplate(file_path, pagesize=A4, rightMargin=70, leftMargin=70, topMargin=30, bottomMargin=30)
    elements = []
   
    styles = getSampleStyleSheet()
    # ตั้งค่าสไตล์ให้ใช้ฟอนต์ภาษาไทย
    styles.add(ParagraphStyle(name='Normal_Thai', parent=styles['Normal'], fontName=font_name, fontSize=12))
    styles.add(ParagraphStyle(name='Heading1_Thai', parent=styles['Heading1'], fontName=font_name_bold, fontSize=16))
    styles.add(ParagraphStyle(name='Center_Thai', parent=styles['Normal_Thai'], alignment=TA_CENTER))


    # --- Header and Employee Info ---
    header_img_path = None
    company_str = (company or "").lower()
    if "einstein" in company_str:
        if os.path.exists(EINSTEIN_ICON_PATH):
            header_img_path = EINSTEIN_ICON_PATH
    elif "lasertechnik" in company_str or "eit" in company_str:
        if os.path.exists(EIT_ICON_PATH):
            header_img_path = EIT_ICON_PATH

    if header_img_path is None and os.path.exists(ICON_PATH):
        header_img_path = ICON_PATH

    if header_img_path and os.path.exists(header_img_path):
        elements.append(Image(header_img_path, width=575, height=100))
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
   
    final_signature_table = Table(
        [
            [line1_table],
            [Paragraph("( ระวิวรรณ งอยภูธร )", styles['Center_Thai'])],
            [Paragraph("ฝ่ายบุคคล", styles['Center_Thai'])],
            [Paragraph("เอกสารนี้เป็นความลับส่วนบุคคล ห้ามเผยแพร่", styles['Center_Thai'])],
        ],
        colWidths=[300],
    )
    final_signature_table.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    final_signature_table.hAlign = 'CENTER'
    elements.append(final_signature_table)


    doc.build(elements)
    return file_path


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
    org = org_selection.get()
    username = entry_user.get().strip()
    password = entry_pass.get().strip()


    if not org:
        messagebox.showerror("เลือกองค์กร", "กรุณาเลือกองค์กรก่อนเข้าสู่ระบบ")
        return


    if org == "EIT Lasertechnik":
        if username == "eit@eitlaser.com" and password == "payslip282895":
            root.withdraw()
            open_Admin_main_window()
        else:
            messagebox.showerror("เข้าสู่ระบบล้มเหลว", "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้องสำหรับ EIT Lasertechnik")
    elif org == "Einstein Industrie Technik (EIT) Laser":
        if username == "eit@eitlaser.com" and password == "payslip282895":
            root.withdraw()
            open_Admin_main_window()
        else:
            messagebox.showerror("เข้าสู่ระบบล้มเหลว", "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้องสำหรับ Einstein Industrie Technik (EIT) Laser")
    else:
        messagebox.showerror("เลือกองค์กร", "กรุณาเลือกองค์กรที่ถูกต้อง")


def open_Admin_main_window():
    """เปิดหน้าต่างหลักสำหรับ Admin"""
    main_win = tk.Toplevel(bg=BG_COLOR)
    main_win.title("Admin Main System")
    center_window(main_win, 400, 300)
    tk.Label(main_win, text="ยินดีต้อนรับสู่ระบบหลังบ้าน Admin", font=("Arial", 12), bg=BG_COLOR).pack(pady=20)

    btn_frame = tk.Frame(main_win, bg=BG_COLOR)
    btn_frame.pack(pady=10)

    if CTK_AVAILABLE:
        ctk.CTkButton(
            btn_frame,
            text="จัดการพนักงาน (Add/Delete)",
            command=lambda: add_delete_employee_window(main_win),
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=260,
        ).pack(pady=6)

        ctk.CTkButton(
            btn_frame,
            text="จัดการเงินเดือน (Salary Management)",
            command=lambda: create_salary_management_window(main_win),
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=260,
        ).pack(pady=6)

        ctk.CTkButton(
            btn_frame,
            text="กลับ",
            command=lambda: [main_win.destroy(), root.deiconify()],
            fg_color="#E5E7EB",
            text_color="#374151",
            hover_color="#D1D5DB",
            border_color="#D1D5DB",
            border_width=1,
            font=("Arial", 12),
            corner_radius=12,
            height=36,
            width=160,
        ).pack(pady=(16, 4))

        ctk.CTkButton(
            btn_frame,
            text="Logout",
            command=lambda: [main_win.destroy(), root.deiconify()],
            fg_color="#EF4444",
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#DC2626",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=160,
        ).pack(pady=(4, 0))
    else:
        tk.Button(
            btn_frame,
            text="จัดการพนักงาน (Add/Delete)",
            command=lambda: add_delete_employee_window(main_win),
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
            activebackground=ACCENT_COLOR,
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
            bd=0,
            width=22,
        ).pack(pady=6, ipady=3)

        tk.Button(
            btn_frame,
            text="จัดการเงินเดือน (Salary Management)",
            command=lambda: create_salary_management_window(main_win),
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
            activebackground=ACCENT_COLOR,
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
            bd=0,
            width=22,
        ).pack(pady=6, ipady=3)

        tk.Button(
            btn_frame,
            text="กลับ",
            command=lambda: [main_win.destroy(), root.deiconify()],
            bg="#E5E7EB",
            fg="#374151",
            activebackground="#D1D5DB",
            activeforeground="#111827",
            font=("Arial", 12),
            relief="flat",
            bd=0,
            width=10,
        ).pack(pady=(16, 4), ipady=3)

        tk.Button(
            btn_frame,
            text="Logout",
            command=lambda: [main_win.destroy(), root.deiconify()],
            bg="#EF4444",
            fg=BUTTON_TEXT_COLOR,
            activebackground="#DC2626",
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
            bd=0,
            width=10,
        ).pack(pady=(4, 0), ipady=3)
    main_win.protocol("WM_DELETE_WINDOW", lambda: [main_win.destroy(), root.deiconify()])




def open_User_main_window():
    """เปิดหน้าต่างหลักสำหรับ User (ตัวอย่าง)"""
    main_win = tk.Toplevel(bg=BG_COLOR)
    main_win.title("User Main System")
    center_window(main_win, 400, 300)
    tk.Label(main_win, text="ยินดีต้อนรับสู่ระบบหลังบ้าน User", font=("Arial", 12), bg=BG_COLOR).pack(pady=50)


    # เพิ่มปุ่มกลับไปหน้าก่อนหน้า
    tk.Button(main_win, text="กลับ", command=lambda: [main_win.destroy(), root.deiconify()], bg="#FF6347", fg=BUTTON_TEXT_COLOR).pack(pady=5, ipadx=10, ipady=5)
   
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

    if CTK_AVAILABLE:
        ctk.CTkButton(
            add_delete_win,
            text="Add Employee",
            command=lambda: add_employee_form(add_delete_win),
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=200,
        ).pack(pady=5, ipadx=10, ipady=5)

        ctk.CTkButton(
            add_delete_win,
            text="Delete Employee",
            command=lambda: delete_employee(add_delete_win),
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=200,
        ).pack(pady=5, ipadx=10, ipady=5)

        ctk.CTkButton(
            add_delete_win,
            text="กลับ",
            command=lambda: [add_delete_win.destroy(), parent_win.deiconify()],
            fg_color="#E5E7EB",
            text_color="#374151",
            hover_color="#D1D5DB",
            border_color="#D1D5DB",
            border_width=1,
            font=("Arial", 12),
            corner_radius=12,
            height=36,
            width=160,
        ).pack(pady=20, ipadx=10, ipady=5)
    else:
        tk.Button(
            add_delete_win,
            text="Add Employee",
            command=lambda: add_employee_form(add_delete_win),
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
        ).pack(pady=5, ipadx=10, ipady=5)

        tk.Button(
            add_delete_win,
            text="Delete Employee",
            command=lambda: delete_employee(add_delete_win),
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
        ).pack(pady=5, ipadx=10, ipady=5)

        tk.Button(
            add_delete_win,
            text="กลับ",
            command=lambda: [add_delete_win.destroy(), parent_win.deiconify()],
            bg="#FF6347",
            fg=BUTTON_TEXT_COLOR,
        ).pack(pady=20, ipadx=10, ipady=5)

    add_delete_win.protocol("WM_DELETE_WINDOW", lambda: [add_delete_win.destroy(), parent_win.deiconify()])




def add_employee_form(parent_win):
    parent_win.withdraw()
    add_form_win = tk.Toplevel(bg=FORM_BG_COLOR)
    add_form_win.title("Add Employee")
    center_window(add_form_win, 900, 550)

    container = tk.Frame(add_form_win, bg="#FFFFFF")
    container.pack(expand=True, fill="both")

    scroll_container = tk.Frame(container, bg="#FFFFFF")
    scroll_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(scroll_container, bg="#FFFFFF", highlightthickness=0)
    scrollbar = tk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    content = tk.Frame(canvas, bg="#FFFFFF")
    content_window = canvas.create_window((0, 0), window=content, anchor="n")

    def on_content_config(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_config(event):
        canvas.itemconfig(content_window, width=event.width)

    content.bind("<Configure>", on_content_config)
    canvas.bind("<Configure>", on_canvas_config)

    content_inner = tk.Frame(content, bg="#FFFFFF")
    content_inner.pack(expand=True, fill="both", padx=32, pady=(20, 32))

    tk.Label(
        content_inner,
        text="EMPLOYEE INFORMATION",
        font=("Arial", 16, "bold"),
        bg="#FFFFFF",
        fg=TEXT_COLOR,
    ).pack(pady=(0, 4), anchor="w")

    tk.Frame(content_inner, bg=FIELD_BORDER_COLOR, height=1).pack(fill="x", pady=(0, 16))

    form = tk.Frame(content_inner, bg="#FFFFFF")
    form.pack(pady=(0, 24), fill="x")

    placeholders = {
        "name": "Enter employee full name",
        "employee_code": "Enter employee ID",
        "id": "Enter national ID number",
        "position": "Enter position",
        "email": "Enter email address",
        "start_date": "Enter start date (DD-MM-YYYY)"
    }

    entries = {}

    def make_placeholder(entry, key, wrapper):
        placeholder = placeholders.get(key, "")
        if not placeholder:
            return
        entry.insert(0, placeholder)

        def set_text_color(color):
            if CTK_AVAILABLE and isinstance(entry, ctk.CTkEntry):
                entry.configure(text_color=color)
            else:
                entry.config(fg=color)

        def set_border_color(color):
            if CTK_AVAILABLE and isinstance(entry, ctk.CTkEntry):
                try:
                    entry.configure(border_color=color)
                except tk.TclError:
                    pass
            elif wrapper is not None:
                wrapper.config(highlightbackground=color)

        set_text_color("#9CA3AF")
        set_border_color(FIELD_BORDER_COLOR)

        def on_focus_in(event):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                set_text_color("#111827")
            set_border_color("#000000")

        def on_focus_out(event):
            if entry.get().strip() == "":
                entry.delete(0, tk.END)
                entry.insert(0, placeholder)
                set_text_color("#9CA3AF")
            else:
                set_text_color("#111827")
            set_border_color(FIELD_BORDER_COLOR)

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    layout = [
        ("ชื่อพนักงาน (คำนำหน้าชื่อ ชื่อ นามสกุล)", "Employee full name", "name", 0, 0),
        ("เลขประจำตัวประชาชน:", "National ID number", "id", 0, 1),
        ("เลขประจำตัวพนักงาน:", "Employee ID", "employee_code", 1, 0),
        ("ตำแหน่ง", "Position", "position", 1, 1),
        ("อีเมล (สำหรับรับสลิป):", "Email for payslip", "email", 2, 0),
        ("วันที่เริ่มงาน", "Start date (DD-MM-YYYY)", "start_date", 2, 1),
    ]

    for th_label, en_label, key, row, col in layout:
        field_block = tk.Frame(form, bg="#FFFFFF")
        field_block.grid(row=row, column=col, sticky="ew", padx=6, pady=(12, 6))

        tk.Label(field_block, text=th_label, bg="#FFFFFF", fg=TEXT_COLOR, font=("Arial", 10, "bold")).pack(anchor="w", padx=12)
        tk.Label(field_block, text=en_label, bg="#FFFFFF", fg="#6B7280", font=("Arial", 9)).pack(anchor="w", pady=(0, 4), padx=12)

        if CTK_AVAILABLE:
            wrapper = tk.Frame(
                field_block,
                bg="#FFFFFF",
                bd=0,
                highlightthickness=0,
            )
        else:
            wrapper = tk.Frame(
                field_block,
                bg="#FFFFFF",
                highlightbackground=FIELD_BORDER_COLOR,
                highlightthickness=1,
                bd=0,
            )
        wrapper.pack(fill="x")

        if CTK_AVAILABLE:
            entry = ctk.CTkEntry(
                wrapper,
                corner_radius=10,
                fg_color="#FFFFFF",
                border_width=1,
                border_color=FIELD_BORDER_COLOR,
                height=40,
                font=("Arial", 12),
            )
            entry.pack(fill="x", padx=12, pady=4)
        else:
            entry = tk.Entry(wrapper, bd=0, relief="flat", bg="#FFFFFF", font=("Arial", 12), width=40)
            entry.pack(fill="x", padx=12, pady=10)

        make_placeholder(entry, key, wrapper)
        entries[key] = entry

    form.grid_columnconfigure(0, weight=1)
    form.grid_columnconfigure(1, weight=1)

    spacer = tk.Frame(content_inner, bg="#FFFFFF", height=8)
    spacer.pack(fill="x")

    actions = tk.Frame(content_inner, bg="#FFFFFF")
    actions.pack(pady=(12, 0), fill="x")

    def on_save():
        employee_data = {}
        for key, entry in entries.items():
            value = entry.get().strip()
            placeholder = placeholders.get(key)
            if placeholder and value == placeholder:
                value = ""
            employee_data[key] = value
        save_employee(employee_data, add_form_win, parent_win)

    if CTK_AVAILABLE:
        ctk.CTkButton(
            actions,
            text="บันทึก",
            command=on_save,
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=120,
        ).pack(side="right", padx=(6, 0), pady=(0, 4))

        ctk.CTkButton(
            actions,
            text="กลับ",
            command=lambda: [add_form_win.destroy(), parent_win.deiconify()],
            fg_color="#E5E7EB",
            text_color="#374151",
            hover_color="#D1D5DB",
            border_color="#D1D5DB",
            border_width=1,
            font=("Arial", 12),
            corner_radius=12,
            height=36,
            width=120,
        ).pack(side="right", padx=(0, 6), pady=(0, 4))
    else:
        tk.Button(
            actions,
            text="บันทึก",
            command=on_save,
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
            activebackground=ACCENT_COLOR,
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
            bd=0,
            width=8,
        ).pack(side="right", padx=(6, 0), ipady=3, ipadx=4)

        tk.Button(
            actions,
            text="กลับ",
            command=lambda: [add_form_win.destroy(), parent_win.deiconify()],
            bg="#E5E7EB",
            fg="#374151",
            activebackground="#D1D5DB",
            activeforeground="#111827",
            bd=0,
            font=("Arial", 12),
            relief="flat",
            width=8,
        ).pack(side="right", padx=(0, 6), ipady=3, ipadx=4)

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


    delete_win = tk.Toplevel(bg=FORM_BG_COLOR)
    delete_win.title("ลบพนักงาน")
    center_window(delete_win, 900, 550)

    container = tk.Frame(delete_win, bg="#FFFFFF")
    container.pack(expand=True, fill="both")

    content_inner = tk.Frame(container, bg="#FFFFFF")
    content_inner.pack(expand=True, fill="both", padx=32, pady=(20, 32))

    tk.Label(
        content_inner,
        text="DELETE EMPLOYEE",
        font=("Arial", 16, "bold"),
        bg="#FFFFFF",
        fg=TEXT_COLOR,
    ).pack(pady=(0, 4), anchor="w")

    tk.Frame(content_inner, bg=FIELD_BORDER_COLOR, height=1).pack(fill="x", pady=(0, 16))

    search_frame = tk.Frame(content_inner, bg="#FFFFFF")
    search_frame.pack(fill="x", pady=(0, 8))

    search_var = tk.StringVar()

    if CTK_AVAILABLE:
        search_wrapper = tk.Frame(
            search_frame,
            bg="#FFFFFF",
            bd=0,
            highlightthickness=0,
        )
        search_wrapper.pack(fill="x", expand=True)

        search_box = ctk.CTkFrame(
            search_wrapper,
            fg_color="#FFFFFF",
            corner_radius=10,
            border_width=1,
            border_color=FIELD_BORDER_COLOR,
        )
        search_box.pack(fill="x", expand=True, pady=4)

        icon_parent = search_box
        entry_parent = search_box
    else:
        search_wrapper = tk.Frame(
            search_frame,
            bg="#FFFFFF",
            highlightbackground=FIELD_BORDER_COLOR,
            highlightthickness=1,
            bd=0,
        )
        search_wrapper.pack(fill="x", expand=True)
        icon_parent = search_wrapper
        entry_parent = search_wrapper

    icon_label = tk.Label(
        icon_parent,
        text="🔍",
        bg="#FFFFFF",
        fg="#6B7280",
        font=("Arial", 12, "bold"),
    )
    icon_label.pack(side="left", padx=(8, 4), pady=4)

    if CTK_AVAILABLE:
        search_entry = ctk.CTkEntry(
            entry_parent,
            corner_radius=0,
            fg_color="#FFFFFF",
            border_width=0,
            height=36,
            font=("Arial", 12),
            textvariable=search_var,
        )
        search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=4)
    else:
        search_entry = tk.Entry(
            entry_parent,
            textvariable=search_var,
            bd=0,
            relief="flat",
            bg="#FFFFFF",
            font=("Arial", 12),
        )
        search_entry.pack(side="left", ipadx=4, ipady=4, fill="x", expand=True, pady=4)

    search_placeholder = "Search Employee Name..."
    search_entry.insert(0, search_placeholder)
    if CTK_AVAILABLE and isinstance(search_entry, ctk.CTkEntry):
        search_entry.configure(text_color="#9CA3AF")
    else:
        search_entry.config(fg="#9CA3AF")

    def set_search_border(color):
        if CTK_AVAILABLE:
            try:
                search_box.configure(border_color=color)
            except tk.TclError:
                pass
        else:
            search_wrapper.config(highlightbackground=color)

    set_search_border(FIELD_BORDER_COLOR)

    def on_search_focus_in(event):
        if search_entry.get() == search_placeholder:
            search_entry.delete(0, tk.END)
            if CTK_AVAILABLE and isinstance(search_entry, ctk.CTkEntry):
                search_entry.configure(text_color="#111827")
            else:
                search_entry.config(fg="#111827")
        set_search_border("#000000")

    def on_search_focus_out(event):
        if search_entry.get().strip() == "":
            search_entry.delete(0, tk.END)
            search_entry.insert(0, search_placeholder)
            if CTK_AVAILABLE and isinstance(search_entry, ctk.CTkEntry):
                search_entry.configure(text_color="#9CA3AF")
            else:
                search_entry.config(fg="#9CA3AF")
        set_search_border(FIELD_BORDER_COLOR)

    search_entry.bind("<FocusIn>", on_search_focus_in)
    search_entry.bind("<FocusOut>", on_search_focus_out)

    if CTK_AVAILABLE:
        list_frame_outer = ctk.CTkFrame(
            content_inner,
            fg_color="#FFFFFF",
            corner_radius=12,
            border_width=1,
            border_color=FIELD_BORDER_COLOR,
        )
        list_frame_outer.pack(pady=(0, 12), padx=2, fill="both")
        list_frame = tk.Frame(
            list_frame_outer,
            bg="#FFFFFF",
            bd=0,
            highlightthickness=0,
        )
        list_frame.pack(fill="both", expand=True, padx=2, pady=2)
    else:
        list_frame = tk.Frame(
            content_inner,
            bg="#FFFFFF",
            highlightbackground=FIELD_BORDER_COLOR,
            highlightthickness=1,
            bd=0,
        )
        list_frame.pack(pady=(0, 12), padx=2, fill="x")

    columns = ("no", "name", "id")

    style = ttk.Style()
    style.configure(
        "Rounded.Treeview.Heading",
        background=ACCENT_COLOR,
        foreground=TEXT_COLOR,
        font=("Arial", 10, "bold"),
    )
    style.configure(
        "Rounded.Treeview",
        font=("Arial", 10),
        rowheight=28,
        borderwidth=0,
        relief="flat",
    )

    emp_tree = ttk.Treeview(
        list_frame,
        columns=columns,
        show="headings",
        selectmode="browse",
        height=10,
        style="Rounded.Treeview",
    )
    emp_tree.heading("no", text="#", anchor="center")
    emp_tree.heading("name", text="Employee", anchor="w")
    emp_tree.heading("id", text="ID", anchor="center")

    emp_tree.column("no", width=40, anchor="center")
    emp_tree.column("name", width=260, anchor="w")
    emp_tree.column("id", width=200, anchor="center")

    emp_tree.tag_configure("odd", background="#F9FAFB")
    emp_tree.tag_configure("even", background="#FFFFFF")

    scroll_y = tk.Scrollbar(list_frame, orient="vertical", command=emp_tree.yview)
    emp_tree.configure(yscrollcommand=scroll_y.set)
    emp_tree.pack(side="left", fill="both", expand=True)
    scroll_y.pack(side="right", fill="y")

    sorted_emps = sorted(employees, key=lambda e: e.get("name", ""))

    def refresh_tree(filtered):
        emp_tree.delete(*emp_tree.get_children())
        for idx, emp in enumerate(filtered, start=1):
            tag = "odd" if idx % 2 == 1 else "even"
            emp_tree.insert("", "end", values=(idx, emp.get("name", ""), emp.get("id", "")), tags=(tag,))

    refresh_tree(sorted_emps)

    def apply_search(*_):
        query = search_var.get().strip().lower()
        if query == search_placeholder.lower():
            query = ""
        if not query:
            refresh_tree(sorted_emps)
            return
        filtered = [
            emp for emp in sorted_emps
            if query in emp.get("name", "").lower() or query in str(emp.get("id", "")).lower()
        ]
        refresh_tree(filtered)

    search_var.trace_add("write", apply_search)

    def confirm_delete():
        selection = emp_tree.selection()
        if not selection:
            messagebox.showerror("ผิดพลาด", "กรุณาเลือกพนักงานที่ต้องการลบ")
            return

        item_id = selection[0]
        values = emp_tree.item(item_id, "values")
        name = values[1]
        id_to_delete = values[2]

        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบ '{name}' ใช่หรือไม่?"):
            return

        employees_to_keep = [emp for emp in employees if emp.get("id") != id_to_delete]
        save_employees_to_file(employees_to_keep)
        messagebox.showinfo("ลบสำเร็จ", f"ข้อมูลของ '{name}' ถูกลบเรียบร้อยแล้ว ✅")
        delete_win.destroy()
        parent_win.deiconify()

    actions = tk.Frame(content_inner, bg="#FFFFFF")
    actions.pack(pady=(12, 0), fill="x")

    if CTK_AVAILABLE:
        ctk.CTkButton(
            actions,
            text="ยืนยันการลบ",
            command=confirm_delete,
            fg_color="#EF4444",
            hover_color="#DC2626",
            text_color=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=140,
        ).pack(side="right", padx=(6, 0), pady=(0, 4))

        ctk.CTkButton(
            actions,
            text="กลับ",
            command=lambda: [delete_win.destroy(), parent_win.deiconify()],
            fg_color="#E5E7EB",
            text_color="#374151",
            hover_color="#D1D5DB",
            border_color="#D1D5DB",
            border_width=1,
            font=("Arial", 12),
            corner_radius=12,
            height=36,
            width=120,
        ).pack(side="right", padx=(0, 6), pady=(0, 4))
    else:
        tk.Button(
            actions,
            text="ยืนยันการลบ",
            command=confirm_delete,
            bg="#FF4B4B",
            fg=BUTTON_TEXT_COLOR,
            activebackground="#FF3333",
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
            bd=0,
            width=12,
        ).pack(side="right", padx=(6, 0), ipady=5, ipadx=8)

        tk.Button(
            actions,
            text="กลับ",
            command=lambda: [delete_win.destroy(), parent_win.deiconify()],
            bg="#E5E7EB",
            fg="#374151",
            activebackground="#D1D5DB",
            activeforeground="#111827",
            bd=0,
            font=("Arial", 12),
            relief="flat",
            width=12,
        ).pack(side="right", padx=(0, 6), ipady=5, ipadx=8)

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
    salary_win.protocol("WM_DELETE_WINDOW", lambda: [salary_win.destroy(), parent_win.deiconify()])


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

    company_var = tk.StringVar(value=org_selection.get() or "EIT Lasertechnik")
    tk.Label(top_frame, text="เลือกบริษัท:", bg=BG_COLOR).pack(side=tk.LEFT, padx=(10, 5))
    company_combo = ttk.Combobox(top_frame, state="readonly", width=30, textvariable=company_var)
    company_combo['values'] = ["EIT Lasertechnik", "Einstein Industrie Technik (EIT) Laser"]
    company_combo.pack(side=tk.LEFT, padx=5)
   
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
        """สร้าง PDF แสดงตัวอย่าง แล้วค่อยถามเพื่อส่งอีเมล"""
        salary_data, emp_data = get_current_data()
       
        if not salary_data or not emp_data:
            messagebox.showerror("ผิดพลาด", "ไม่พบข้อมูลพนักงานหรือข้อมูลเงินเดือน")
            return

        # บันทึกข้อมูลล่าสุดก่อนสร้าง PDF
        save_salary_data()
       
        recipient_email = emp_data.get("email")
        if not recipient_email:
            messagebox.showerror("ผิดพลาด", f"ไม่พบข้อมูลอีเมลสำหรับ '{emp_data['name']}' ❌\nกรุณาเพิ่มอีเมลในหน้าจัดการพนักงาน")
            return

        try:
            # 1. สร้าง PDF
            pdf_company = company_var.get() or org_selection.get() or "EIT Lasertechnik"
            pdf_file = create_pay_slip_pdf(emp_data, salary_data, pdf_company)

            # 2. เปิดไฟล์ PDF ให้ดูตัวอย่าง
            try:
                os.startfile(pdf_file)
            except Exception as e:
                messagebox.showwarning("ไม่สามารถเปิดไฟล์", f"สร้างไฟล์ {pdf_file} แล้ว แต่ไม่สามารถเปิดไฟล์อัตโนมัติได้:\n{e}")

            # 3. ถามยืนยันก่อนส่งอีเมล
            if not messagebox.askyesno("ยืนยันการส่งอีเมล", f"ได้แสดงตัวอย่างสลิปเงินเดือนให้ '{emp_data['name']}' แล้ว\n\nต้องการส่งอีเมลไปที่ {recipient_email} หรือไม่?"):
                messagebox.showinfo("ยกเลิกการส่ง", "ระบบสร้างไฟล์ Pay Slip แล้ว แต่ยังไม่ได้ส่งอีเมล")
                return
           
            # 4. ส่ง Email
            if send_email_with_attachment(pdf_file, recipient_email, emp_data['name']):
                messagebox.showinfo("สำเร็จ", f"✅ ส่งใบสรุปเงินเดือนให้ '{emp_data['name']}' ที่อีเมล {recipient_email} เรียบร้อยแล้ว")
            else:
                messagebox.showerror("ผิดพลาด", "❌ ไม่สามารถส่งอีเมลได้ กรุณาตรวจสอบการตั้งค่า SENDER_EMAIL และ SENDER_PASSWORD (App Password)")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด (PDF/Email)", f"ไม่สามารถสร้างหรือส่ง Pay Slip ได้:\n{e}")


    if CTK_AVAILABLE:
        ctk.CTkButton(
            button_frame,
            text="Save Data",
            command=save_salary_data,
            fg_color="#22C55E",
            hover_color="#16A34A",
            text_color=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=150,
        ).pack(side=tk.LEFT, padx=10, pady=(0, 4))

        ctk.CTkButton(
            button_frame,
            text="Generate & Send Payslip",
            command=generate_and_send,
            fg_color=ACCENT_COLOR,
            hover_color="#1D4ED8",
            text_color=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
            width=200,
        ).pack(side=tk.LEFT, padx=10, pady=(0, 4))

        ctk.CTkButton(
            button_frame,
            text="กลับไปหน้าหลัก",
            command=lambda: [salary_win.destroy(), parent_win.deiconify()],
            fg_color="#E5E7EB",
            text_color="#374151",
            hover_color="#D1D5DB",
            border_color="#D1D5DB",
            border_width=1,
            font=("Arial", 12),
            corner_radius=12,
            height=36,
            width=160,
        ).pack(side=tk.LEFT, padx=10, pady=(0, 4))
    else:
        tk.Button(
            button_frame,
            text="Save Data",
            command=save_salary_data,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            relief="flat",
        ).pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)

        tk.Button(
            button_frame,
            text="Generate & Send Payslip",
            command=generate_and_send,
            bg=ACCENT_COLOR,
            fg="white",
            font=("Arial", 12, "bold"),
            relief="flat",
        ).pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)

        tk.Button(
            button_frame,
            text="กลับไปหน้าหลัก",
            command=lambda: [salary_win.destroy(), parent_win.deiconify()],
            bg="#f44336",
            fg="white",
            font=("Arial", 12, "bold"),
            relief="flat",
        ).pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)


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


    def refresh_employee_list(event=None):
        employees = load_employees()
        if not employees:
            emp_combo['values'] = []
            emp_combo.set('')
            month_combo['values'] = []
            month_combo.set('')
            return

        company = (company_var.get() or "").lower()
        if "einstein" in company:
            filtered = [emp for emp in employees if emp['name'] != "MissChatrawee Sungkaksem"]
        else:
            filtered = [emp for emp in employees if emp['name'] == "MissChatrawee Sungkaksem"]

        employee_names = sorted([emp['name'] for emp in filtered])
        emp_combo['values'] = employee_names

        month_combo['values'] = []
        month_combo.set('')

        if employee_names:
            emp_combo.set(employee_names[0])
            on_employee_select()
        else:
            emp_combo.set('')

    emp_combo.bind("<<ComboboxSelected>>", on_employee_select)
    month_combo.bind("<<ComboboxSelected>>", import_excel_data)
    company_combo.bind("<<ComboboxSelected>>", refresh_employee_list)
    refresh_employee_list()




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
root.configure(bg=BG_COLOR)


# --- Pre-login Organization Selection ---
org_selection = tk.StringVar(value="EIT Lasertechnik")




def show_org_selection():
    for w in root.winfo_children():
        w.destroy()


    if os.path.exists(ICON_PATH):
        header_img_data = tk.PhotoImage(file=ICON_PATH).subsample(2, 2)
        header_label = tk.Label(root, image=header_img_data, bg=BG_COLOR)
        header_label.image = header_img_data
        header_label.pack(pady=(20, 10))


    container = tk.Frame(root, bg=BG_COLOR)
    container.pack(expand=True, fill="both")

    shadow = tk.Frame(container, bg=BG_COLOR)
    shadow.place(relx=0.5, rely=0.5, anchor="center")

    card_shadow = tk.Frame(shadow, bg="#D1D5DB")
    card_shadow.pack(pady=4)

    card = tk.Frame(card_shadow, bg="#FFFFFF", bd=0, highlightbackground="#E5E7EB", highlightthickness=1)
    card.pack(padx=2, pady=2, ipadx=60, ipady=20)

    top_accent = tk.Frame(card, bg=ACCENT_COLOR, height=3)
    top_accent.pack(fill="x", side="top")

    tk.Label(card, text="Choose Organization", font=("Arial", 18, "bold"), fg=TEXT_COLOR, bg="#FFFFFF").pack(pady=(24, 4))
    tk.Label(card, text="Select the company to continue to login", font=("Arial", 10), fg="#6B7280", bg="#FFFFFF").pack(pady=(0, 12))

    divider = tk.Frame(card, bg="#E5E7EB", height=1)
    divider.pack(fill="x", padx=32, pady=(0, 12))

    btn_frame = tk.Frame(card, bg="#FFFFFF")
    btn_frame.pack(pady=8, padx=32, fill="x")


    tk.Button(
        btn_frame,
        text="EIT Lasertechnik",
        command=lambda: [org_selection.set("EIT Lasertechnik"), show_login_form()],
        bg=ACCENT_COLOR,
        fg=BUTTON_TEXT_COLOR,
        activebackground=ACCENT_COLOR,
        activeforeground=BUTTON_TEXT_COLOR,
        font=("Arial", 11, "bold"),
        relief="flat",
    ).pack(pady=6, ipadx=6, ipady=8, fill="x")


    tk.Button(
        btn_frame,
        text="Einstein Industrie Technik\n(EIT) Laser",
        command=lambda: [org_selection.set("Einstein Industrie Technik (EIT) Laser"), show_login_form()],
        bg=ACCENT_COLOR,
        fg=BUTTON_TEXT_COLOR,
        activebackground=ACCENT_COLOR,
        activeforeground=BUTTON_TEXT_COLOR,
        font=("Arial", 11, "bold"),
        relief="flat",
    ).pack(pady=6, ipadx=6, ipady=8, fill="x")




def show_login_form():
    for w in root.winfo_children():
        w.destroy()


    root.title(f"EIT Backoffice System - {org_selection.get()}")


    container = tk.Frame(root, bg=BG_COLOR)
    container.pack(expand=True, fill="both")

    banner = tk.Frame(container, bg=BG_COLOR)
    banner.pack(pady=(20, 10))

    tk.Label(banner, text="Payslip Management", font=("Arial", 18, "bold"), fg=TEXT_COLOR, bg=BG_COLOR).pack()

    shadow = tk.Frame(container, bg=BG_COLOR)
    shadow.pack(pady=4)

    card_shadow = tk.Frame(shadow, bg="#D1D5DB")
    card_shadow.pack(pady=4)

    card = tk.Frame(card_shadow, bg="#FFFFFF", bd=0, highlightbackground="#E5E7EB", highlightthickness=1)
    card.pack(padx=2, pady=2, ipadx=40)

    top_accent = tk.Frame(card, bg=ACCENT_COLOR, height=3)
    top_accent.pack(fill="x", side="top")

    tk.Label(card, text="Login", font=("Arial", 18, "bold"), bg="#FFFFFF", fg=TEXT_COLOR).pack(pady=(24, 4))
    tk.Label(card, text="Enter your credentials to continue", font=("Arial", 10), bg="#FFFFFF", fg="#6B7280").pack(pady=(0, 16))

    form = tk.Frame(card, bg="#FFFFFF")
    form.pack(padx=32, pady=(0, 16), fill="x")

    tk.Label(form, text="Email", bg="#FFFFFF", fg=TEXT_COLOR).pack(anchor="w", pady=(0, 4))
    global entry_user, entry_pass

    if CTK_AVAILABLE:
        entry_user = ctk.CTkEntry(
            form,
            corner_radius=10,
            fg_color="#FFFFFF",
            border_width=1,
            border_color=FIELD_BORDER_COLOR,
            height=36,
            font=("Arial", 12),
        )
        entry_user.pack(fill="x", pady=(0, 10))
    else:
        entry_user = tk.Entry(form, bd=0, relief="flat", bg="#FFFFFF", font=("Arial", 12))
        entry_user.pack(fill="x", pady=(0, 10))

    tk.Label(form, text="Password", bg="#FFFFFF", fg=TEXT_COLOR).pack(anchor="w", pady=(4, 4))

    password_row = tk.Frame(form, bg="#FFFFFF")
    password_row.pack(fill="x", pady=(0, 10))

    password_state = {"visible": False}

    if CTK_AVAILABLE:
        password_wrapper = ctk.CTkFrame(
            password_row,
            fg_color="#FFFFFF",
            border_color=FIELD_BORDER_COLOR,
            border_width=1,
            corner_radius=10,
        )
        password_wrapper.pack(fill="x")

        entry_pass = tk.Entry(
            password_wrapper,
            show="*",
            bd=0,
            relief="flat",
            bg="#FFFFFF",
            font=("Arial", 12),
        )
        entry_pass.pack(side="left", fill="x", expand=True, padx=(12, 4), pady=6)

        toggle_btn = ctk.CTkButton(
            password_wrapper,
            text="👁",
            command=lambda: None,
            fg_color="#FFFFFF",
            text_color="#6B7280",
            hover_color="#F3F4F6",
            corner_radius=16,
            border_width=0,
            width=32,
            height=32,
        )
        toggle_btn.pack(side="right", padx=(0, 8), pady=4)
    else:
        password_wrapper = tk.Frame(
            password_row,
            bg="#FFFFFF",
            bd=0,
            highlightbackground=FIELD_BORDER_COLOR,
            highlightthickness=1,
        )
        password_wrapper.pack(fill="x")

        entry_pass = tk.Entry(
            password_wrapper,
            show="*",
            bd=0,
            relief="flat",
            bg="#FFFFFF",
            font=("Arial", 12),
        )
        entry_pass.pack(side="left", fill="x", expand=True, padx=(10, 4), pady=6)

        toggle_btn = tk.Button(
            password_wrapper,
            text="👁",
            bg="#FFFFFF",
            fg="#6B7280",
            activebackground="#FFFFFF",
            activeforeground="#111827",
            relief="flat",
            bd=0,
            cursor="hand2",
            width=2,
            command=lambda: None,
        )
        toggle_btn.pack(side="right", padx=(0, 8), pady=4)

    def toggle_password():
        if password_state["visible"]:
            if CTK_AVAILABLE:
                entry_pass.configure(show="*")
                toggle_btn.configure(text="👁")
            else:
                entry_pass.config(show="*")
                toggle_btn.config(text="👁")
            password_state["visible"] = False
        else:
            if CTK_AVAILABLE:
                entry_pass.configure(show="")
                toggle_btn.configure(text="👁")
            else:
                entry_pass.config(show="")
                toggle_btn.config(text="👁")
            password_state["visible"] = True

    if CTK_AVAILABLE:
        toggle_btn.configure(command=toggle_password)
    else:
        toggle_btn.config(command=toggle_password)

    options_row = tk.Frame(card, bg="#FFFFFF")
    options_row.pack(padx=32, pady=(0, 12), fill="x")

    if CTK_AVAILABLE:
        ctk.CTkButton(
            options_row,
            text="Forgot password?",
            command=lambda: None,
            fg_color="#FFFFFF",
            text_color=ACCENT_COLOR,
            hover_color="#F3F4F6",
            corner_radius=8,
            border_width=0,
        ).pack(side="right")

        ctk.CTkButton(
            card,
            text="Login",
            command=login,
            fg_color=ACCENT_COLOR,
            text_color=BUTTON_TEXT_COLOR,
            hover_color="#1D4ED8",
            font=("Arial", 12, "bold"),
            corner_radius=12,
            height=36,
        ).pack(padx=32, pady=(0, 16), fill="x")

        # ไม่มีปุ่มย้อนกลับไปหน้าเลือกองค์กรอีกต่อไป
    else:
        tk.Button(
            options_row,
            text="Forgot password?",
            command=lambda: None,
            bg="#FFFFFF",
            fg=ACCENT_COLOR,
            activebackground="#FFFFFF",
            activeforeground=ACCENT_COLOR,
            relief="flat",
            cursor="hand2",
        ).pack(side="right")

        btn_login = tk.Button(
            card,
            text="Login",
            command=login,
            bg=ACCENT_COLOR,
            fg=BUTTON_TEXT_COLOR,
            activebackground=ACCENT_COLOR,
            activeforeground=BUTTON_TEXT_COLOR,
            font=("Arial", 12, "bold"),
            relief="flat",
        )
        btn_login.pack(padx=32, pady=(0, 16), ipadx=24, ipady=8, fill="x")

        # ไม่มีปุ่มย้อนกลับไปหน้าเลือกองค์กรอีกต่อไป


show_login_form()


# เริ่มการทำงานของโปรแกรม
root.mainloop()

