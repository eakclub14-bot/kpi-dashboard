import calendar
import io
import os
import re
from datetime import datetime
from flask import Flask, jsonify, render_template, send_file
import pandas as pd

try:
    import pythoncom
    import win32com.client
except ImportError:
    win32com = None
    pythoncom = None

try:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    pass

app = Flask(__name__)

LOADING_PROGRESS = 0

TARGET_ASSIGNEES = [
    "sasiprapa.ke@rsm.co.th",
    "panyakron.ch@rsm.co.th",
    "thapakorn.od@rsm.co.th",
    "husadin.sr@rsm.co.th",
    "nattapon.ni@rsm.co.th",
]

MONTH_NAME_TO_NUM = {
    'JAN': 1, 'JANUARY': 1, 'ม.ค.': 1, 'มกราคม': 1,
    'FEB': 2, 'FEBRUARY': 2, 'ก.พ.': 2, 'กุมภาพันธ์': 2,
    'MAR': 3, 'MARCH': 3, 'มี.ค.': 3, 'มีนาคม': 3,
    'APR': 4, 'APRIL': 4, 'เม.ย.': 4, 'เมษายน': 4,
    'MAY': 5, 'พ.ค.': 5, 'พฤษภาคม': 5,
    'JUN': 6, 'JUNE': 6, 'มิ.ย.': 6, 'มิถุนายน': 6,
    'JUL': 7, 'JULY': 7, 'ก.ค.': 7, 'กรกฎาคม': 7,
    'AUG': 8, 'AUGUST': 8, 'ส.ค.': 8, 'สิงหาคม': 8,
    'SEP': 9, 'SEPTEMBER': 9, 'ก.ย.': 9, 'กันยายน': 9,
    'OCT': 10, 'OCTOBER': 10, 'ต.ค.': 10, 'ตุลาคม': 10,
    'NOV': 11, 'NOVEMBER': 11, 'พ.ย.': 11, 'พฤศจิกายน': 11,
    'DEC': 12, 'DECEMBER': 12, 'ธ.ค.': 12, 'ธันวาคม': 12,
}

NUM_TO_MONTH_CODE = {
    1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN',
    7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC',
}


def calculate_due_date_info(month_num, year, report_type):
    """คำนวณ Due Date บังคับใช้ปี ค.ศ. 2026 เป็นอย่างน้อย"""
    if not year or year < 2026:
        year = 2026

    next_month = 1 if month_num == 12 else month_num + 1
    next_year = year + 1 if month_num == 12 else year

    due_day = 1 if report_type == 'MONTHEND_DSR' else 5

    due_dt = datetime(next_year, next_month, due_day, 17, 0, 0)
    
    # รูปแบบสำหรับหน้าเว็บ และ Excel
    due_str_iso = due_dt.strftime('%Y-%m-%d %H:%M:%S')
    due_str_excel = due_dt.strftime('%d/%m/%Y %H:%M:%S')
    
    return due_str_iso, due_str_excel, due_dt


def detect_report_details(subject, body, sent_date_str=None):
    full_text = f"{subject} {body}".upper()
    report_type = None
    category = "อื่นๆ"

    if re.search(r'(CONTROL\s*NUMBER|CONTROL)', full_text):
        report_type = 'CONTROL_NUMBER'
        category = 'Control Number'
    elif re.search(r'(MONTHEND\s*DSR|MONTHLY\s*DSR|DSR_TRACKING|DSR|\bTRACKING\b)', full_text):
        report_type = 'MONTHEND_DSR'
        category = 'Monthend DSR'
    elif re.search(r'(สินค้าแถม|ของแถม|แถม|PREMIUM|FREE)', full_text):
        report_type = 'FREE_PREMIUM'
        category = 'รายงานสินค้าแถม'

    detected_month = None
    detected_year = None

    month_regex = r'(JAN|JANUARY|ม\.ค\.|มกราคม|FEB|FEBRUARY|ก\.พ\.|กุมภาพันธ์|MAR|MARCH|มี\.ค\.|มีนาคม|APR|APRIL|เม\.ย\.|เมษายน|MAY|พ\.ค\.|พฤษภาคม|JUN|JUNE|มิ\.ย\.|มิถุนายน|JUL|JULY|ก\.ค\.|กรกฎาคม|AUG|AUGUST|ส\.ค\.|สิงหาคม|SEP|SEPTEMBER|ก\.ย\.|กันยายน|OCT|OCTOBER|ต\.ค\.|ตุลาคม|NOV|NOVEMBER|พ\.ย\.|พฤศจิกายน|DEC|DECEMBER|ธ\.ค\.|ธันวาคม)'
    year_regex = r"(?:\'|\s*\/|\s*)?(\d{4}|\d{2})?"

    match = re.search(rf"{month_regex}{year_regex}", full_text, re.IGNORECASE)
    if match:
        month_str = match.group(1).strip()
        year_str = match.group(2)

        detected_month = MONTH_NAME_TO_NUM.get(month_str)

        if year_str:
            parsed_year = int(year_str)
            if parsed_year > 2500:
                parsed_year -= 543
            elif parsed_year < 100:
                parsed_year += 2000
            
            if parsed_year < 2026:
                parsed_year = 2026
                
            detected_year = parsed_year

    sent_dt = parse_date_string(sent_date_str) if sent_date_str else None

    if not detected_month and sent_dt:
        detected_month = sent_dt.month - 1 if sent_dt.month > 1 else 12
        detected_year = sent_dt.year if sent_dt.month > 1 else sent_dt.year - 1
    elif detected_month and not detected_year:
        detected_year = sent_dt.year if sent_dt else 2026

    if not detected_month:
        detected_month = 9
    
    if not detected_year or detected_year < 2026:
        detected_year = 2026

    period_tag = NUM_TO_MONTH_CODE.get(detected_month, 'SEP')
    due_iso, due_excel, due_dt = calculate_due_date_info(detected_month, detected_year, report_type)

    return due_iso, due_excel, due_dt, period_tag, report_type, category, detected_month, detected_year


def parse_date_string(date_str):
    if not date_str:
        return None
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%d/%m/%Y %H:%M:%S',
        '%d/%m/%Y %H:%M',
    ]
    for fmt in formats:
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except ValueError:
            pass
    return None


def get_clean_sender_email(msg):
    try:
        sender_email = str(msg.SenderEmailAddress or '')
        if "@" in sender_email:
            return sender_email.lower().strip()

        if hasattr(msg, "Sender") and msg.Sender:
            sender_obj = msg.Sender
            if hasattr(sender_obj, "GetExchangeUser"):
                ex_user = sender_obj.GetExchangeUser()
                if ex_user and ex_user.PrimarySmtpAddress:
                    return str(ex_user.PrimarySmtpAddress).lower().strip()

        PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        try:
            smtp_addr = msg.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            if smtp_addr and "@" in smtp_addr:
                return str(smtp_addr).lower().strip()
        except Exception:
            pass

        sender_name = str(msg.SenderName or '').lower().strip()
        for assignee in TARGET_ASSIGNEES:
            prefix = assignee.split('@')[0].lower()
            if prefix in sender_name:
                return assignee.lower()

        return sender_email.lower().strip()
    except Exception:
        return ""


def fetch_emails_from_outlook():
    global LOADING_PROGRESS
    LOADING_PROGRESS = 0
    emails = []

    if not win32com:
        LOADING_PROGRESS = 100
        return emails

    if pythoncom:
        pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        report_pattern = re.compile(
            r'(CONTROL\s*NUMBER|CONTROL|DSR|TRACKING|MONTHEND\s*DSR|MONTHLY\s*DSR|สินค้าแถม|ของแถม|แถม)',
            re.IGNORECASE,
        )

        MAX_CHECK = 5000
        total_items = min(messages.Count, MAX_CHECK)
        target_list_lower = [e.lower().strip() for e in TARGET_ASSIGNEES]

        for i, msg in enumerate(messages):
            if i >= MAX_CHECK:
                break

            LOADING_PROGRESS = int(((i + 1) / total_items) * 100)

            try:
                subject = str(msg.Subject or '')
                body = str(msg.Body or '')
                full_text = f"{subject} {body}"

                if not report_pattern.search(full_text):
                    continue

                sender_email_clean = get_clean_sender_email(msg)

                if sender_email_clean not in target_list_lower:
                    continue

                sent_date = msg.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
                attachment_count = msg.Attachments.Count
                attachment_status = (
                    f"มีไฟล์แนบ ({attachment_count} ไฟล์)"
                    if attachment_count > 0
                    else "ไม่มีไฟล์แนบ"
                )

                emails.append({
                    "subject": subject,
                    "body": body,
                    "assignee": sender_email_clean,
                    "sent_date": sent_date,
                    "attachment_status": attachment_status,
                })
            except Exception:
                continue

    except Exception as e:
        print(f"Error fetching Outlook emails: {e}")
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
        LOADING_PROGRESS = 100

    return emails


def build_dashboard_payload():
    raw_emails = fetch_emails_from_outlook()

    if not raw_emails:
        raw_emails = [
            {
                "subject": "RE: รบกวนช่วยส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS)",
                "body": "คะแนน CD ร้านค้า HFS เดือน Sep 2026 RCB CD1: 11,439 CD2: 10,631 CD3: 9,914",
                "assignee": "sasiprapa.ke@rsm.co.th",
                "sent_date": "2026-10-05 09:50:00",
                "attachment_status": "ไม่มีไฟล์แนบ",
            },
            {
                "subject": "RCB_MAS_HFS_DSR_Tracking [30 Sep 26] MonthEnd",
                "body": "ไฟล์แนบ CBR_MAS HFS DSR Tracking_V15 (30 Sep 26) - MonthEnd.xlsx",
                "assignee": "sasiprapa.ke@rsm.co.th",
                "sent_date": "2026-10-01 15:16:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "ข้อมูลสินค้าแถมประจำเดือน SEP'2026",
                "body": "ส่งรายงานของแถมประจำเดือน กันยายน 2569 ครับ",
                "assignee": "husadin.sr@rsm.co.th",
                "sent_date": "2026-10-01 15:11:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "BON_MAS_HFS_DSR_Tracking_30Sep26 (MonthEnd)",
                "body": "ไฟล์แนบ BON_MAS_HFS_DSR_Tracking_V15_30Sep26(MonthEnd).xlsx",
                "assignee": "nattapon.ni@rsm.co.th",
                "sent_date": "2026-10-01 14:04:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "Monthend KKC_DSR MAS&HFS REPORT 30 Sep 2026",
                "body": "ไฟล์แนบ Monthend KKC_MAS_HFS_DSR_Tracking_V15 30 Sep 26.xlsx",
                "assignee": "thapakorn.od@rsm.co.th",
                "sent_date": "2026-10-01 10:14:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "RE: RSM MAS HFS DSR Tracking 31 Aug'26 [Month End]",
                "body": "ไฟล์แนบ RSM_MAS_HFS_DSR_Tracking_V15 Form (31 Aug'26)_MonthEND.xlsx",
                "assignee": "panyakron.ch@rsm.co.th",
                "sent_date": "2026-09-01 10:39:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "RE: รบกวนช่วยส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS)",
                "body": "คะแนน CD ร้านค้า HFS RSM AUG'26 ครับ CD1: 17,823 CD2: 17,661 CD3: 12,715",
                "assignee": "panyakron.ch@rsm.co.th",
                "sent_date": "2026-09-02 14:32:00",
                "attachment_status": "ไม่มีไฟล์แนบ",
            },
            {
                "subject": "ส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS) UTH",
                "body": "ส่งคะแนน CD AUG'2026 ครับ",
                "assignee": "husadin.sr@rsm.co.th",
                "sent_date": "2026-09-02 10:49:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "รบกวนช่วยส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS) KKC",
                "body": "คะแนน CD AUG'2026 ครับ",
                "assignee": "thapakorn.od@rsm.co.th",
                "sent_date": "2026-09-02 10:22:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "ส่งรายงานสินค้าแถมประจำเดือน MAY 2026 RSMGROUP",
                "body": "ข้อมูลของแถมประจำเดือน MAY'26 ครับ",
                "assignee": "husadin.sr@rsm.co.th",
                "sent_date": "2026-06-04 10:36:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "ส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS) UTH",
                "body": "คะแนน CD MAY'2026 ครับ",
                "assignee": "husadin.sr@rsm.co.th",
                "sent_date": "2026-06-04 08:44:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "ส่งรายงานสินค้าแถมประจำเดือน APR 2026 RSMGROUP",
                "body": "ข้อมูลของแถมประจำเดือน APR'26 ครับ",
                "assignee": "husadin.sr@rsm.co.th",
                "sent_date": "2026-05-04 10:22:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "RE:รบกวนช่วยส่ง Control Number ตรวจสอบยอดขาย และ คะแนน CD (HFS)",
                "body": "ไฟล์ Promotion IM Feb26.xlsx",
                "assignee": "nattapon.ni@rsm.co.th",
                "sent_date": "2026-03-04 16:24:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
            {
                "subject": "ส่ง Control Number ตรวจสอบยอดขาย JAN 2026",
                "body": "คะแนน CD JAN'2026 ครับ",
                "assignee": "panyakron.ch@rsm.co.th",
                "sent_date": "2026-02-04 09:15:00",
                "attachment_status": "มีไฟล์แนบ (1 ไฟล์)",
            },
        ]

    processed_reports = []
    user_stats = {
        user.lower(): {"points": 0, "total": 0, "ontime": 0}
        for user in TARGET_ASSIGNEES
    }

    for email in raw_emails:
        subject = email.get("subject", "")
        body = email.get("body", "")
        assignee = email.get("assignee", "Unknown").lower()
        sent_date_str = email.get("sent_date", "")

        (
            due_iso,
            due_excel,
            due_dt,
            period_tag,
            report_type,
            category,
            m_num,
            y_num,
        ) = detect_report_details(subject, body, sent_date_str)

        sent_dt = parse_date_string(sent_date_str)

        status_type = "ontime"
        status_label = "ตรงเวลา"
        points = 10

        if sent_dt and due_dt and sent_dt > due_dt:
            status_type = "late"
            status_label = "ส่งล่าช้า"
            points = 5

        if assignee in user_stats:
            user_stats[assignee]["total"] += 1
            if status_type in ["ontime", "early"]:
                user_stats[assignee]["ontime"] += 1
            user_stats[assignee]["points"] += points

        processed_reports.append({
            "report_name": subject,
            "subject": subject,
            "assignee": assignee,
            "submitted_date": sent_date_str,
            "sent_date": sent_date_str,
            "due_date": due_iso,
            "due_date_display": due_iso,
            "due_date_excel": due_excel,
            "period": period_tag,
            "category": category,
            "status": status_label,
            "status_type": status_type,
            "points": points,
            "attachment_status": email.get("attachment_status", "ไม่มีไฟล์แนบ"),
        })

    leaderboard = []

    #  จุดที่ปรับแก้ไข: จัดอันดับแบบ ยุติธรรม (Fair Ranking)
    # เรียงลำดับจาก:
    # 1. เปอร์เซ็นต์ความตรงเวลา (% On-time Rate)
    # 2. คะแนนเฉลี่ยต่อรายงาน (Avg Points per Report)
    # 3. คะแนนรวมสะสม (Total Points)
    sorted_users = sorted(
        user_stats.items(),
        key=lambda x: (
            (x[1]["ontime"] / x[1]["total"]) if x[1]["total"] > 0 else 0,
            (x[1]["points"] / x[1]["total"]) if x[1]["total"] > 0 else 0,
            x[1]["points"]
        ),
        reverse=True
    )

    for rank, (email, stats) in enumerate(sorted_users, 1):
        rate_val = (stats['ontime'] / stats['total'] * 100) if stats['total'] > 0 else 0.0
        rate_str = f"{rate_val:.1f}%"
        avg_points = (stats['points'] / stats['total']) if stats['total'] > 0 else 0.0

        leaderboard.append({
            "rank": rank,
            "assignee": email,
            "points": stats["points"],
            "avg_points": f"{avg_points:.1f}",
            "ontime_rate": rate_str,
            "total": stats["total"],
            "ontime": stats["ontime"],
        })

    return {"reports": processed_reports, "leaderboard": leaderboard}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/progress')
def get_progress():
    global LOADING_PROGRESS
    return jsonify({"progress": LOADING_PROGRESS})


@app.route('/api/dashboard', methods=['GET'])
def get_dashboard_data():
    payload = build_dashboard_payload()
    return jsonify(payload)


@app.route('/api/export-excel', methods=['GET'])
def export_excel():
    dashboard_data = build_dashboard_payload()
    reports_data = dashboard_data.get("reports", [])
    leaderboard_data = dashboard_data.get("leaderboard", [])

    df_reports = pd.DataFrame(reports_data)
    df_leaderboard = pd.DataFrame(leaderboard_data)

    if not df_reports.empty:
        df_reports['due_date'] = df_reports['due_date_excel']
        df_reports = df_reports[
            [
                'subject',
                'assignee',
                'period',
                'category',
                'sent_date',
                'due_date',
                'status',
                'points',
                'attachment_status',
            ]
        ]
        df_reports.columns = [
            'ชื่อรายงาน',
            'ผู้รับผิดชอบ',
            'ประจำเดือน',
            'หมวดหมู่',
            'วันที่ส่ง',
            'วันกำหนดส่ง',
            'สถานะ',
            'คะแนน',
            'ไฟล์แนบ',
        ]

    if not df_leaderboard.empty:
        df_leaderboard = df_leaderboard[
            ['rank', 'assignee', 'points', 'avg_points', 'ontime_rate', 'total', 'ontime']
        ]
        df_leaderboard.columns = [
            'อันดับ',
            'ผู้รับผิดชอบ',
            'คะแนนรวม',
            'คะแนนเฉลี่ย/รายงาน',
            'อัตราส่งตรงเวลา',
            'รายงานทั้งหมด',
            'ส่งตรงเวลา (ครั้ง)',
        ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_leaderboard.to_excel(
            writer, sheet_name='Leaderboard', index=False, startrow=3
        )
        df_reports.to_excel(
            writer, sheet_name='รายงานทั้งหมด', index=False, startrow=3
        )

        workbook = writer.book

        navy_header_fill = PatternFill(
            start_color="1F4E79", end_color="1F4E79", fill_type="solid"
        )
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

        title_font = Font(name="Segoe UI", size=15, bold=True, color="1F4E79")
        data_font = Font(name="Segoe UI", size=10)

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9'),
        )

        status_ontime_fill = PatternFill(
            start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
        )
        status_ontime_font = Font(
            name="Segoe UI", size=10, color="375623", bold=True
        )

        status_late_fill = PatternFill(
            start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
        )
        status_late_font = Font(
            name="Segoe UI", size=10, color="C65911", bold=True
        )

        ws_lb = workbook['Leaderboard']
        ws_lb['A1'] = "📊 รายงานสรุปอันดับผู้ส่งรายงานประจำปี 2026 (Fair Ranking)"
        ws_lb['A1'].font = title_font

        ws_rep = workbook['รายงานทั้งหมด']
        ws_rep['A1'] = "📋 รายละเอียดการส่งรายงานทั้งหมดประจำปี 2026"
        ws_rep['A1'].font = title_font

        for ws in [ws_lb, ws_rep]:
            ws.views.sheetView[0].showGridLines = True

            ws.row_dimensions[4].height = 28
            for cell in ws[4]:
                cell.fill = navy_header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
                ws.row_dimensions[row[0].row].height = 22
                for cell in row:
                    cell.border = thin_border
                    cell.font = data_font
                    cell.alignment = Alignment(
                        vertical="center", horizontal="left"
                    )

                    val_str = str(cell.value or '').strip()

                    if cell.column_letter in ['A', 'C', 'D', 'E', 'F', 'G', 'H']:
                        if ws == ws_rep and cell.column_letter not in ['A']:
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )
                        elif ws == ws_lb and cell.column_letter in [
                            'A',
                            'C',
                            'D',
                            'E',
                            'F',
                            'G'
                        ]:
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )

                    if val_str == "ตรงเวลา":
                        cell.fill = status_ontime_fill
                        cell.font = status_ontime_font
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
                    elif val_str == "ส่งล่าช้า":
                        cell.fill = status_late_fill
                        cell.font = status_late_font
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row >= 4 and cell.value:
                        val = str(cell.value)
                        length = sum(
                            2 if ord(char) > 255 else 1 for char in val
                        )
                        if length > max_len:
                            max_len = length
                adjusted_width = min(max(max_len + 4, 12), 60)
                ws.column_dimensions[col_letter].width = adjusted_width

    output.seek(0)

    return send_file(
        output,
        download_name="Report_Tracking_2026.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)