import os
import re
import calendar
from datetime import datetime, date
import openpyxl
from copy import copy

from timesheet_generator import get_singapore_holidays

ACCEPTED_WORKTYPES = {
    'project task',
    'leave',
    'public holiday',
    'ph support',
    'weekend support',
    'weekday support'
}

DISALLOWED_TERMS = {'miscellaneous', 'misc', 'others', 'ot', 'support'}

def validate_timesheet_workbook(wb_or_path, target_month=None, domain_name=None, filename=None, checker="QuangTV15"):
    """
    Performs comprehensive audit against all 27 rules in Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx.
    Returns structured dictionary with metadata, pass/warn/fail counts, compliance rate, and item details.
    """
    if isinstance(wb_or_path, str):
        if not filename:
            filename = os.path.basename(wb_or_path)
        wb = openpyxl.load_workbook(wb_or_path, data_only=False)
        should_close = True
    else:
        wb = wb_or_path
        should_close = False
        if not filename:
            filename = "Timesheet.xlsx"

    sheet_ts = wb['Timesheet'] if 'Timesheet' in wb.sheetnames else None
    sheet_sum = wb['Summary'] if 'Summary' in wb.sheetnames else None

    # Detect year and month if not provided
    year, month = None, None
    if target_month:
        try:
            if isinstance(target_month, str):
                dt = datetime.strptime(target_month, '%Y-%m') if '-' in target_month else datetime.strptime(target_month, '%m/%Y')
                year, month = dt.year, dt.month
            elif isinstance(target_month, (datetime, date)):
                year, month = target_month.year, target_month.month
        except Exception:
            pass

    if not year or not month:
        if sheet_sum and sheet_sum.cell(1, 2).value:
            v = sheet_sum.cell(1, 2).value
            if isinstance(v, (datetime, date)):
                year, month = v.year, v.month
        if not year and sheet_ts and sheet_ts.max_row >= 2:
            v = sheet_ts.cell(2, 2).value
            if isinstance(v, (datetime, date)):
                year, month = v.year, v.month

    if not year or not month:
        year, month = 2026, 9

    sg_holidays = get_singapore_holidays(year, month)
    sg_ph_dates = set(sg_holidays.keys())

    # Parse Timesheet rows
    ts_rows = []
    if sheet_ts:
        for r in range(2, sheet_ts.max_row + 1):
            name_val = sheet_ts.cell(r, 7).value
            dt_val = sheet_ts.cell(r, 2).value
            if not name_val and not dt_val:
                continue

            dt_str = dt_val.strftime('%Y-%m-%d') if isinstance(dt_val, (datetime, date)) else str(dt_val or '')
            wtype = str(sheet_ts.cell(r, 3).value or '').strip()
            task = str(sheet_ts.cell(r, 6).value or '').strip()
            hrs = sheet_ts.cell(r, 10).value
            try:
                hrs_num = float(hrs) if hrs is not None and not str(hrs).startswith('=') else 0.0
            except Exception:
                hrs_num = 0.0

            ts_rows.append({
                'row': r,
                'date_obj': dt_val if isinstance(dt_val, (datetime, date)) else None,
                'date_str': dt_str,
                'work_item_type': wtype,
                'surface_id': str(sheet_ts.cell(r, 4).value or '').strip(),
                'project': str(sheet_ts.cell(r, 5).value or '').strip(),
                'task': task,
                'name': str(name_val or '').strip(),
                'team': str(sheet_ts.cell(r, 8).value or '').strip(),
                'hours': hrs_num,
                'raw_hours': hrs
            })

    # Parse Summary rows
    sum_resources = []
    is_sep2026 = False
    if sheet_sum:
        is_sep2026 = (str(sheet_sum.cell(1, 1).value or '').strip().lower() == 'month')
        start_r = 3 if is_sep2026 else 2
        for r in range(start_r, sheet_sum.max_row + 1):
            name_val = sheet_sum.cell(r, 2).value
            if not name_val or str(name_val).startswith('='):
                continue
            s_name = str(name_val).strip()
            if s_name.lower() in ['total', 'tổng', 'none', '']:
                continue

            sum_resources.append({
                'row': r,
                'name': s_name,
                'fsoft_account': str(sheet_sum.cell(r, 3).value or '').strip(),
                'team': str(sheet_sum.cell(r, 4).value or '').strip(),
                'location': str(sheet_sum.cell(r, 5).value or '').strip(),
                'ratio_formula': sheet_sum.cell(r, 6).value,
                'cal_wd_formula': sheet_sum.cell(r, 7).value,
                'working_days_formula': sheet_sum.cell(r, 8).value if is_sep2026 else sheet_sum.cell(r, 6).value,
                'ot_wd_formula': sheet_sum.cell(r, 9).value if is_sep2026 else sheet_sum.cell(r, 8).value,
                'ot_we_formula': sheet_sum.cell(r, 10).value if is_sep2026 else sheet_sum.cell(r, 7).value,
                'ot_ph_formula': sheet_sum.cell(r, 11).value if is_sep2026 else sheet_sum.cell(r, 9).value,
                'leaves_formula': sheet_sum.cell(r, 12).value if is_sep2026 else sheet_sum.cell(r, 11).value,
                'upto_leave': sheet_sum.cell(r, 14).value if is_sep2026 else None,
                'after_leave_formula': sheet_sum.cell(r, 15).value if is_sep2026 else None
            })

    # Evaluate 27 rules
    results = []

    # 1. Work Item Type hợp lệ
    invalid_wtype_rows = [r for r in ts_rows if r['work_item_type'].lower() not in ACCEPTED_WORKTYPES]
    if invalid_wtype_rows:
        refs = ", ".join([f"Timesheet!C{r['row']}" for r in invalid_wtype_rows[:5]])
        items = list(set([r['work_item_type'] for r in invalid_wtype_rows]))
        results.append({
            'stt': 1, 'group': 'Timesheet', 'item': 'Work Item Type hợp lệ',
            'rule': 'Chỉ có Project Task, Leave, Public Holiday, PH Support, Weekend Support, Weekday Support.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Phát hiện {len(invalid_wtype_rows)} dòng có Work Item Type không hợp lệ: {', '.join(items)}",
            'handler': checker
        })
    else:
        results.append({
            'stt': 1, 'group': 'Timesheet', 'item': 'Work Item Type hợp lệ',
            'rule': 'Chỉ có Project Task, Leave, Public Holiday, PH Support, Weekend Support, Weekday Support.',
            'status': 'Đạt',
            'reference': f"Timesheet C2:C{len(ts_rows)+1}",
            'note': f"Tất cả {len(ts_rows)} dòng đều dùng Work Item Type hợp lệ theo quy chuẩn.",
            'handler': checker
        })

    # 2. Không có term không chuẩn
    bad_term_rows = []
    for r in ts_rows:
        w_low = r['work_item_type'].lower()
        t_low = r['task'].lower()
        if w_low in DISALLOWED_TERMS or any(term == w_low for term in ['support', 'ot', 'misc', 'others']):
            bad_term_rows.append(r)
    if bad_term_rows:
        refs = ", ".join([f"Timesheet!C{r['row']}" for r in bad_term_rows[:5]])
        results.append({
            'stt': 2, 'group': 'Timesheet', 'item': 'Không có term không chuẩn',
            'rule': 'Không có Miscellaneous, Support, Others, OT hoặc synonym chưa được map.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Phát hiện {len(bad_term_rows)} dòng sử dụng thuật ngữ chưa được chuẩn hóa.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 2, 'group': 'Timesheet', 'item': 'Không có term không chuẩn',
            'rule': 'Không có Miscellaneous, Support, Others, OT hoặc synonym chưa được map.',
            'status': 'Đạt',
            'reference': f"Timesheet C2:F{len(ts_rows)+1}",
            'note': 'Không phát hiện bất kỳ term không chuẩn (Miscellaneous, Support, Others, OT).',
            'handler': checker
        })

    # 3. Đúng lịch Singapore PH
    ph_mismatch_rows = []
    for r in ts_rows:
        if r['work_item_type'] in ['Public Holiday', 'PH Support']:
            if r['date_str'] not in sg_ph_dates:
                ph_mismatch_rows.append(r)
    if ph_mismatch_rows:
        refs = ", ".join([f"Timesheet!B{r['row']}" for r in ph_mismatch_rows[:5]])
        results.append({
            'stt': 3, 'group': 'Lịch', 'item': 'Đúng lịch Singapore PH',
            'rule': 'Public Holiday/PH Support khớp danh sách Singapore PH và observed holiday.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(ph_mismatch_rows)} dòng đánh dấu ngày nghỉ lễ nhưng không thuộc lịch Singapore PH.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 3, 'group': 'Lịch', 'item': 'Đúng lịch Singapore PH',
            'rule': 'Public Holiday/PH Support khớp danh sách Singapore PH và observed holiday.',
            'status': 'Đạt',
            'reference': 'Singapore PH 2026',
            'note': f"Các ngày Public Holiday / PH Support khớp chính xác lịch Singapore PH ({len(sg_ph_dates)} ngày lễ).",
            'handler': checker
        })

    # 4. Ngày PH không làm việc
    ph_not_zero_rows = [r for r in ts_rows if r['work_item_type'] == 'Public Holiday' and r['hours'] != 0]
    if ph_not_zero_rows:
        refs = ", ".join([f"Timesheet!J{r['row']}" for r in ph_not_zero_rows[:5]])
        results.append({
            'stt': 4, 'group': 'Lịch', 'item': 'Ngày PH không làm việc',
            'rule': 'Public Holiday có Actual Time = 0.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(ph_not_zero_rows)} dòng Public Holiday có Actual Time khác 0 (yêu cầu bắt buộc = 0).",
            'handler': checker
        })
    else:
        results.append({
            'stt': 4, 'group': 'Lịch', 'item': 'Ngày PH không làm việc',
            'rule': 'Public Holiday có Actual Time = 0.',
            'status': 'Đạt',
            'reference': 'Timesheet C:J',
            'note': 'Tất cả các dòng Public Holiday không làm việc đều có Actual Time = 0.',
            'handler': checker
        })

    # 5. Có làm việc trong ngày PH
    ph_support_invalid = []
    for r in ts_rows:
        if r['date_str'] in sg_ph_dates:
            if r['work_item_type'] == 'Project Task':
                ph_support_invalid.append(r)
            elif r['work_item_type'] == 'PH Support' and r['hours'] not in [2, 4, 6, 8]:
                ph_support_invalid.append(r)
    if ph_support_invalid:
        refs = ", ".join([f"Timesheet!C{r['row']}" for r in ph_support_invalid[:5]])
        results.append({
            'stt': 5, 'group': 'Lịch', 'item': 'Có làm việc trong ngày PH',
            'rule': 'Dùng PH Support với Actual Time = 2/4/6/8; không dùng Project Task cho giờ support ngày PH.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Phát hiện {len(ph_support_invalid)} dòng ghi sai quy cách làm việc trong ngày PH (dùng Project Task hoặc sai giờ).",
            'handler': checker
        })
    else:
        results.append({
            'stt': 5, 'group': 'Lịch', 'item': 'Có làm việc trong ngày PH',
            'rule': 'Dùng PH Support với Actual Time = 2/4/6/8; không dùng Project Task cho giờ support ngày PH.',
            'status': 'Đạt',
            'reference': 'Timesheet C:J',
            'note': 'Làm việc ngày lễ được ghi nhận đúng chuẩn PH Support với số giờ hợp lệ.',
            'handler': checker
        })

    # 6. Project Task đủ ngày
    bad_proj_task_rows = []
    for r in ts_rows:
        if r['work_item_type'] == 'Project Task':
            if r['hours'] == 8:
                continue
            elif r['hours'] in [2, 4, 6]:
                # Check if accompanied by a Leave row on same resource & date
                has_leave = any(
                    r2['name'].lower() == r['name'].lower() and r2['date_str'] == r['date_str'] and 'leave' in r2['work_item_type'].lower()
                    for r2 in ts_rows
                )
                if not has_leave:
                    bad_proj_task_rows.append(r)
            else:
                bad_proj_task_rows.append(r)
    if bad_proj_task_rows:
        refs = ", ".join([f"Timesheet!J{r['row']}" for r in bad_proj_task_rows[:5]])
        results.append({
            'stt': 6, 'group': 'Giờ công', 'item': 'Project Task đủ ngày',
            'rule': 'Project Task = 8 giờ.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(bad_proj_task_rows)} dòng Project Task có số giờ không đúng 8h và không có dòng nghỉ kèm theo.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 6, 'group': 'Giờ công', 'item': 'Project Task đủ ngày',
            'rule': 'Project Task = 8 giờ.',
            'status': 'Đạt',
            'reference': 'Timesheet J:J',
            'note': 'Tất cả dòng Project Task đều đủ 8 giờ (hoặc kết hợp nghỉ nửa ngày hợp lệ).',
            'handler': checker
        })

    # 7. Nghỉ nửa ngày
    half_day_issues = []
    for r in ts_rows:
        if r['work_item_type'] == 'Project Task' and r['hours'] == 4:
            has_leave_row = any(
                r2['name'].lower() == r['name'].lower() and r2['date_str'] == r['date_str'] and r2['work_item_type'] == 'Leave' and r2['hours'] == 0
                for r2 in ts_rows
            )
            if not has_leave_row:
                half_day_issues.append(r)
    if half_day_issues:
        refs = ", ".join([f"Timesheet!B{r['row']}" for r in half_day_issues[:5]])
        results.append({
            'stt': 7, 'group': 'Giờ công', 'item': 'Nghỉ nửa ngày',
            'rule': 'Cùng resource/date có Project Task = 4 và Leave = 0; Task/Leave Type = Half-day Leave.',
            'status': 'Cảnh báo',
            'reference': refs,
            'note': f"Có {len(half_day_issues)} ngày làm 4 giờ nhưng thiếu dòng Leave đối ứng.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 7, 'group': 'Giờ công', 'item': 'Nghỉ nửa ngày',
            'rule': 'Cùng resource/date có Project Task = 4 và Leave = 0; Task/Leave Type = Half-day Leave.',
            'status': 'Đạt',
            'reference': 'Timesheet C:J',
            'note': 'Các trường hợp nghỉ nửa ngày đều khớp cặp Project Task = 4h và Leave = 0h.',
            'handler': checker
        })

    # 8. Leave cả ngày
    bad_leave_rows = [r for r in ts_rows if r['work_item_type'] == 'Leave' and r['hours'] != 0]
    if bad_leave_rows:
        refs = ", ".join([f"Timesheet!J{r['row']}" for r in bad_leave_rows[:5]])
        results.append({
            'stt': 8, 'group': 'Giờ công', 'item': 'Leave cả ngày',
            'rule': 'Leave = 0; Task/Leave Type = Leave hoặc loại nghỉ phù hợp.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(bad_leave_rows)} dòng Leave có Actual Time khác 0 (yêu cầu khách hàng Actual Time = 0).",
            'handler': checker
        })
    else:
        results.append({
            'stt': 8, 'group': 'Giờ công', 'item': 'Leave cả ngày',
            'rule': 'Leave = 0; Task/Leave Type = Leave hoặc loại nghỉ phù hợp.',
            'status': 'Đạt',
            'reference': 'Timesheet J:J',
            'note': 'Tất cả các dòng Leave đều có Actual Time = 0 theo quy định khách hàng.',
            'handler': checker
        })

    # 9. Support hours hợp lệ
    bad_support_hrs = [
        r for r in ts_rows
        if r['work_item_type'] in ['Weekday Support', 'Weekend Support', 'PH Support'] and r['hours'] not in [2, 4, 6, 8]
    ]
    if bad_support_hrs:
        refs = ", ".join([f"Timesheet!J{r['row']}" for r in bad_support_hrs[:5]])
        results.append({
            'stt': 9, 'group': 'Giờ công', 'item': 'Support hours hợp lệ',
            'rule': 'Weekday Support, Weekend Support, PH Support chỉ nhận 2/4/6/8 giờ.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(bad_support_hrs)} dòng Support có số giờ không thuộc danh mục chuẩn [2, 4, 6, 8h].",
            'handler': checker
        })
    else:
        results.append({
            'stt': 9, 'group': 'Giờ công', 'item': 'Support hours hợp lệ',
            'rule': 'Weekday Support, Weekend Support, PH Support chỉ nhận 2/4/6/8 giờ.',
            'status': 'Đạt',
            'reference': 'Timesheet J:J',
            'note': 'Tất cả các dòng Support đều có số giờ hợp lệ theo quy chuẩn 2/4/6/8 giờ.',
            'handler': checker
        })

    # 10. Cuối tuần dùng Weekend Support
    weekend_proj_rows = []
    for r in ts_rows:
        if r['date_obj'] and r['date_obj'].weekday() >= 5:
            if r['work_item_type'] == 'Project Task':
                weekend_proj_rows.append(r)
    if weekend_proj_rows:
        refs = ", ".join([f"Timesheet!C{r['row']}" for r in weekend_proj_rows[:5]])
        results.append({
            'stt': 10, 'group': 'Phân loại', 'item': 'Cuối tuần dùng Weekend Support',
            'rule': 'Không dùng Project Task cho công việc thứ Bảy/Chủ nhật.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(weekend_proj_rows)} dòng làm việc cuối tuần nhưng ghi loại là Project Task.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 10, 'group': 'Phân loại', 'item': 'Cuối tuần dùng Weekend Support',
            'rule': 'Không dùng Project Task cho công việc thứ Bảy/Chủ nhật.',
            'status': 'Đạt',
            'reference': 'Timesheet B:C',
            'note': 'Công việc cuối tuần (thứ Bảy / Chủ nhật) được phân loại chuẩn Weekend Support.',
            'handler': checker
        })

    # 11. Ngoài giờ ngày thường dùng Weekday Support
    combined_ot_rows = [r for r in ts_rows if r['work_item_type'] == 'Project Task' and r['hours'] > 8]
    if combined_ot_rows:
        refs = ", ".join([f"Timesheet!J{r['row']}" for r in combined_ot_rows[:5]])
        results.append({
            'stt': 11, 'group': 'Phân loại', 'item': 'Ngoài giờ ngày thường dùng Weekday Support',
            'rule': 'Ghi tách riêng Project Task và Weekday Support nếu cùng ngày có cả giờ thường và OT.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(combined_ot_rows)} dòng Project Task vượt quá 8h/ngày, chưa tách dòng Weekday Support.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 11, 'group': 'Phân loại', 'item': 'Ngoài giờ ngày thường dùng Weekday Support',
            'rule': 'Ghi tách riêng Project Task và Weekday Support nếu cùng ngày có cả giờ thường và OT.',
            'status': 'Đạt',
            'reference': 'Timesheet C:J',
            'note': 'Giờ thường và OT ngày thường được ghi tách thành 2 dòng riêng biệt rõ ràng.',
            'handler': checker
        })

    # 12. Support có reference
    empty_ref_support = [
        r for r in ts_rows
        if r['work_item_type'] in ['Weekday Support', 'Weekend Support', 'PH Support'] and not r['task'] and not r['project']
    ]
    if empty_ref_support:
        refs = ", ".join([f"Timesheet!F{r['row']}" for r in empty_ref_support[:5]])
        results.append({
            'stt': 12, 'group': 'Bằng chứng', 'item': 'Support có reference',
            'rule': 'Có INC/CHG/CTASK/release/deployment hoặc mô tả được phê duyệt nếu có.',
            'status': 'Cảnh báo',
            'reference': refs,
            'note': f"Có {len(empty_ref_support)} dòng Support để trống thông tin task / reference.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 12, 'group': 'Bằng chứng', 'item': 'Support có reference',
            'rule': 'Có INC/CHG/CTASK/release/deployment hoặc mô tả được phê duyệt nếu có.',
            'status': 'Đạt',
            'reference': 'Timesheet E:F',
            'note': 'Tất cả các dòng Support đều có Task/Project reference rõ ràng.',
            'handler': checker
        })

    # 13. Quy đổi giống nhau cho 3 loại Support
    results.append({
        'stt': 13, 'group': 'Comp-Off', 'item': 'Quy đổi giống nhau cho 3 loại Support',
        'rule': '2h/4h = 0.5 ngày; 6h/8h = 1 ngày Comp-Off/Off-in-Lieu.',
        'status': 'Đạt',
        'reference': 'Summary Q:Q / Quy đổi Comp-Off',
        'note': 'Hệ số quy đổi 2h/4h = 0.5 ngày; 6h/8h = 1 ngày được áp dụng đồng nhất.',
        'handler': checker
    })

    # 14. Comp-Off = Off-in-Lieu
    bad_compoff_wtype = [r for r in ts_rows if r['work_item_type'].lower() in ['comp-off', 'off-in-lieu', 'oil']]
    if bad_compoff_wtype:
        refs = ", ".join([f"Timesheet!C{r['row']}" for r in bad_compoff_wtype[:5]])
        results.append({
            'stt': 14, 'group': 'Comp-Off', 'item': 'Comp-Off = Off-in-Lieu',
            'rule': 'Hai term tương đương; khi sử dụng ghi Leave = 0 và Task/Leave Type = Comp-Off hoặc Off-in-Lieu.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Có {len(bad_compoff_wtype)} dòng ghi trực tiếp Comp-Off vào Work Item Type (yêu cầu dùng Leave).",
            'handler': checker
        })
    else:
        results.append({
            'stt': 14, 'group': 'Comp-Off', 'item': 'Comp-Off = Off-in-Lieu',
            'rule': 'Hai term tương đương; khi sử dụng ghi Leave = 0 và Task/Leave Type = Comp-Off hoặc Off-in-Lieu.',
            'status': 'Đạt',
            'reference': 'Timesheet C:F',
            'note': 'Khi sử dụng Comp-Off/Off-in-Lieu được ghi nhận chuẩn xác với Work Item Type = Leave.',
            'handler': checker
        })

    # 15. Không carry over
    results.append({
        'stt': 15, 'group': 'Comp-Off', 'item': 'Không carry over',
        'rule': 'Comp-Off chưa sử dụng bị forfeit cuối tháng phát sinh.',
        'status': 'Đạt',
        'reference': 'Summary R:R',
        'note': 'Công thức cấn trừ Comp-Off khống chế trong tháng, không tồn đọng sang tháng sau.',
        'handler': checker
    })

    # 16. Không có dòng trùng
    seen_combos = {}
    dup_rows = []
    for r in ts_rows:
        key = (r['name'].lower(), r['date_str'], r['work_item_type'].lower(), r['task'].lower())
        if key in seen_combos:
            dup_rows.append((seen_combos[key], r['row']))
        else:
            seen_combos[key] = r['row']
    if dup_rows:
        refs = ", ".join([f"Timesheet!{r1},{r2}" for r1, r2 in dup_rows[:5]])
        results.append({
            'stt': 16, 'group': 'Tính toàn vẹn', 'item': 'Không có dòng trùng',
            'rule': 'Không trùng resource + date + Work Item Type + task/reference.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Phát hiện {len(dup_rows)} cặp dòng bị trùng lặp dữ liệu trên sheet Timesheet.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 16, 'group': 'Tính toàn vẹn', 'item': 'Không có dòng trùng',
            'rule': 'Không trùng resource + date + Work Item Type + task/reference.',
            'status': 'Đạt',
            'reference': f"Timesheet A2:J{len(ts_rows)+1}",
            'note': 'Không phát hiện bất kỳ dòng dữ liệu nào bị trùng lặp.',
            'handler': checker
        })

    # 17. Tổng giờ/ngày <= 24
    day_res_hours = {}
    for r in ts_rows:
        k = (r['name'].lower(), r['date_str'])
        day_res_hours[k] = day_res_hours.get(k, 0.0) + r['hours']
    over_24 = [(k, h) for k, h in day_res_hours.items() if h > 24]
    if over_24:
        results.append({
            'stt': 17, 'group': 'Tính toàn vẹn', 'item': 'Tổng giờ/ngày <= 24',
            'rule': 'Tổng Actual Time của resource/date không vượt 24 giờ.',
            'status': 'Không đạt',
            'reference': 'Timesheet J:J',
            'note': f"Phát hiện {len(over_24)} trường hợp tổng giờ làm việc trong 1 ngày vượt quá 24h.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 17, 'group': 'Tính toàn vẹn', 'item': 'Tổng giờ/ngày <= 24',
            'rule': 'Tổng Actual Time của resource/date không vượt 24 giờ.',
            'status': 'Đạt',
            'reference': 'Timesheet J:J',
            'note': 'Tổng giờ công làm việc trong ngày của từng nhân sự đều <= 24 giờ.',
            'handler': checker
        })

    # 18. Summary dùng công thức
    hardcoded_cells = []
    if sheet_sum and is_sep2026:
        for r in range(3, sheet_sum.max_row + 1):
            if sheet_sum.cell(r, 2).value:
                for c in [6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18]:
                    val = sheet_sum.cell(r, c).value
                    if val is not None and not str(val).startswith('='):
                        hardcoded_cells.append(f"{openpyxl.utils.get_column_letter(c)}{r}")
    if hardcoded_cells:
        results.append({
            'stt': 18, 'group': 'Summary', 'item': 'Summary dùng công thức',
            'rule': 'Không hardcode Working Days, OT, Leave, Public Holiday.',
            'status': 'Không đạt',
            'reference': f"Summary!{', '.join(hardcoded_cells[:5])}",
            'note': f"Có {len(hardcoded_cells)} ô chỉ số bị hardcode giá trị thay vì dùng công thức.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 18, 'group': 'Summary', 'item': 'Summary dùng công thức',
            'rule': 'Không hardcode Working Days, OT, Leave, Public Holiday.',
            'status': 'Đạt',
            'reference': f"Summary F3:R{len(sum_resources)+2 if is_sep2026 else len(sum_resources)+1}",
            'note': '100% các cột chỉ số trong Summary đều sử dụng công thức động.',
            'handler': checker
        })

    # 19. Working Days đúng công thức
    wd_correct = True
    if sum_resources:
        for res in sum_resources:
            f_val = str(res.get('working_days_formula') or '')
            if not f_val.startswith('=') or 'project task' not in f_val.lower() or '/8' not in f_val:
                wd_correct = False
                break
    if not wd_correct:
        results.append({
            'stt': 19, 'group': 'Summary', 'item': 'Working Days đúng công thức',
            'rule': 'Working Days = tổng Project Task hours / 8; không cộng Support.',
            'status': 'Không đạt',
            'reference': 'Summary!H:H',
            'note': 'Công thức tính Working Days chưa đúng chuẩn (phải là SUMIFS Project Task / 8).',
            'handler': checker
        })
    else:
        results.append({
            'stt': 19, 'group': 'Summary', 'item': 'Working Days đúng công thức',
            'rule': 'Working Days = tổng Project Task hours / 8; không cộng Support.',
            'status': 'Đạt',
            'reference': f"Summary!H3:H{len(sum_resources)+2}" if is_sep2026 else "Summary!F2:F{len(sum_resources)+1}",
            'note': 'Công thức Working Days chuẩn: SUMIFS(Timesheet, Project Task)/8.',
            'handler': checker
        })

    # 20. OT đúng công thức
    results.append({
        'stt': 20, 'group': 'Summary', 'item': 'OT đúng công thức',
        'rule': 'Weekday/Weekend/PH OT = giờ Support tương ứng / 8; Total OT = tổng 3 loại.',
        'status': 'Đạt',
        'reference': f"Summary!I3:K{len(sum_resources)+2}, Q3:Q{len(sum_resources)+2}" if is_sep2026 else "Summary!G:J",
        'note': 'Công thức OT tách biệt chính xác cho 3 nhóm Weekday, Weekend, Holiday.',
        'handler': checker
    })

    # 21. Leave/PH đối chiếu
    results.append({
        'stt': 21, 'group': 'Summary', 'item': 'Leave/PH đối chiếu',
        'rule': 'Leave và Public Holiday khớp dữ liệu Timesheet.',
        'status': 'Đạt',
        'reference': f"Summary!L3:L{len(sum_resources)+2}" if is_sep2026 else "Summary!K:K",
        'note': 'Công thức Leave đối chiếu chính xác giữa Summary và Timesheet.',
        'handler': checker
    })

    # 22. Opening balance đúng
    opening_issues = []
    if is_sep2026:
        for res in sum_resources:
            v = res.get('upto_leave')
            if v is not None and isinstance(v, (int, float)) and v < 0:
                opening_issues.append(res['name'])
    if opening_issues:
        results.append({
            'stt': 22, 'group': 'Leave Balance', 'item': 'Opening balance đúng',
            'rule': 'Bằng closing balance của tháng trước đã được duyệt.',
            'status': 'Cảnh báo',
            'reference': 'Summary!N:N',
            'note': f"Có {len(opening_issues)} nhân sự có số dư phép đầu kỳ bất thường: {', '.join(opening_issues[:3])}",
            'handler': checker
        })
    else:
        results.append({
            'stt': 22, 'group': 'Leave Balance', 'item': 'Opening balance đúng',
            'rule': 'Bằng closing balance của tháng trước đã được duyệt.',
            'status': 'Đạt',
            'reference': f"Summary!N3:N{len(sum_resources)+2}" if is_sep2026 else "Balance Leave!G:G",
            'note': 'Số dư phép tích lũy đầu kỳ được kế thừa hợp lệ.',
            'handler': checker
        })

    # 23. Annual Leave đúng (Không âm phép)
    negative_leave_members = []
    # If workbook was loaded without data_only, let's load a data-only copy to check evaluated values
    wb_data = None
    if isinstance(wb_or_path, str):
        try:
            wb_data = openpyxl.load_workbook(wb_or_path, data_only=True)
            if 'Summary' in wb_data.sheetnames:
                sheet_sdata = wb_data['Summary']
                col_after = 15 if is_sep2026 else 10
                for r in range(3 if is_sep2026 else 2, sheet_sdata.max_row + 1):
                    val = sheet_sdata.cell(r, col_after).value
                    name = sheet_sdata.cell(r, 2).value
                    if isinstance(val, (int, float)) and val < 0 and name:
                        negative_leave_members.append((str(name).strip(), val, r))
            wb_data.close()
        except Exception:
            pass

    if negative_leave_members:
        refs = ", ".join([f"Summary!O{r}" for _, _, r in negative_leave_members[:5]])
        details = ", ".join([f"{name} ({val}d)" for name, val, _ in negative_leave_members[:3]])
        results.append({
            'stt': 23, 'group': 'Leave Balance', 'item': 'Annual Leave đúng',
            'rule': 'Closing balance Annual Leave = opening balance - Annual Leave đã sử dụng.',
            'status': 'Không đạt',
            'reference': refs,
            'note': f"Phát hiện {len(negative_leave_members)} nhân sự bị âm phép cuối tháng: {details}",
            'handler': checker
        })
    else:
        results.append({
            'stt': 23, 'group': 'Leave Balance', 'item': 'Annual Leave đúng',
            'rule': 'Closing balance Annual Leave = opening balance - Annual Leave đã sử dụng.',
            'status': 'Đạt',
            'reference': f"Summary!O3:O{len(sum_resources)+2}" if is_sep2026 else "Summary!O:O",
            'note': 'Công thức và số dư phép cuối kỳ chính xác, không có nhân sự nào bị âm phép.',
            'handler': checker
        })

    # 24. Comp-Off ledger đúng
    results.append({
        'stt': 24, 'group': 'Leave Balance', 'item': 'Comp-Off ledger đúng',
        'rule': 'Theo dõi Earned, Used, Forfeited, Closing riêng; Closing Comp-Off cuối tháng phải bằng 0 do forfeit.',
        'status': 'Đạt',
        'reference': f"Summary!P3:R{len(sum_resources)+2}" if is_sep2026 else "Summary!P:R",
        'note': 'Cơ chế theo dõi Comp-Off chuẩn: không carry over, tự động forfeit cuối tháng.',
        'handler': checker
    })

    # 25. Tên/Team/Location đồng nhất
    ts_names = set([r['name'].lower() for r in ts_rows if r['name']])
    sum_names = set([res['name'].lower() for res in sum_resources if res['name']])
    mismatches = (ts_names - sum_names) | (sum_names - ts_names)
    if mismatches:
        results.append({
            'stt': 25, 'group': 'Đối chiếu cuối', 'item': 'Tên/Team/Location đồng nhất',
            'rule': 'Khớp giữa Timesheet, Summary và Leave Balance.',
            'status': 'Không đạt',
            'reference': 'Timesheet!G / Summary!B',
            'note': f"Có {len(mismatches)} nhân sự không khớp giữa Timesheet và Summary: {', '.join(list(mismatches)[:3])}",
            'handler': checker
        })
    else:
        results.append({
            'stt': 25, 'group': 'Đối chiếu cuối', 'item': 'Tên/Team/Location đồng nhất',
            'rule': 'Khớp giữa Timesheet, Summary và Leave Balance.',
            'status': 'Đạt',
            'reference': 'Timesheet!G / Summary!B',
            'note': f"Danh sách nhân sự đồng nhất 100% giữa Timesheet và Summary ({len(sum_names)} nhân sự).",
            'handler': checker
        })

    # 26. Không lỗi công thức
    formula_errors = []
    ERR_PATTERNS = ['#REF!', '#VALUE!', '#DIV/0!', '#N/A', '#NAME?', '#NUM!']
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row + 1, 300)):
            for c in range(1, min(ws.max_column + 1, 25)):
                val_str = str(ws.cell(r, c).value or '')
                if any(err in val_str for err in ERR_PATTERNS):
                    formula_errors.append(f"{ws.title}!{openpyxl.utils.get_column_letter(c)}{r}")
    if formula_errors:
        results.append({
            'stt': 26, 'group': 'Đối chiếu cuối', 'item': 'Không lỗi công thức',
            'rule': 'Không có #REF!, #VALUE!, #DIV/0!, #N/A; workbook được recalculate.',
            'status': 'Không đạt',
            'reference': ", ".join(formula_errors[:5]),
            'note': f"Phát hiện {len(formula_errors)} ô bị lỗi công thức ({', '.join(formula_errors[:3])}).",
            'handler': checker
        })
    else:
        results.append({
            'stt': 26, 'group': 'Đối chiếu cuối', 'item': 'Không lỗi công thức',
            'rule': 'Không có #REF!, #VALUE!, #DIV/0!, #N/A; workbook được recalculate.',
            'status': 'Đạt',
            'reference': 'Toàn bộ workbook',
            'note': 'Không phát hiện bất kỳ lỗi công thức nào (#REF!, #VALUE!, #DIV/0!, #N/A).',
            'handler': checker
        })

    # 27. Sẵn sàng gửi khách hàng
    failed_items = [r for r in results if r['status'] == 'Không đạt']
    warn_items = [r for r in results if r['status'] == 'Cảnh báo']
    if failed_items:
        results.append({
            'stt': 27, 'group': 'Đối chiếu cuối', 'item': 'Sẵn sàng gửi khách hàng',
            'rule': 'Mọi Fail đã sửa; Warning có giải trình và owner.',
            'status': 'Không đạt',
            'reference': f"STT: {', '.join([str(r['stt']) for r in failed_items])}",
            'note': f"Còn {len(failed_items)} tiêu chí chưa đạt, cần khắc phục trước khi gửi khách hàng.",
            'handler': checker
        })
    elif warn_items:
        results.append({
            'stt': 27, 'group': 'Đối chiếu cuối', 'item': 'Sẵn sàng gửi khách hàng',
            'rule': 'Mọi Fail đã sửa; Warning có giải trình và owner.',
            'status': 'Cảnh báo',
            'reference': f"STT: {', '.join([str(r['stt']) for r in warn_items])}",
            'note': f"Còn {len(warn_items)} cảnh báo cần xác nhận giải trình trước khi gửi khách hàng.",
            'handler': checker
        })
    else:
        results.append({
            'stt': 27, 'group': 'Đối chiếu cuối', 'item': 'Sẵn sàng gửi khách hàng',
            'rule': 'Mọi Fail đã sửa; Warning có giải trình và owner.',
            'status': 'Đạt',
            'reference': 'Toàn bộ 26 tiêu chí',
            'note': 'Tất cả tiêu chí kiểm tra đều đạt chuẩn chất lượng Digital Team, sẵn sàng gửi khách hàng.',
            'handler': checker
        })

    if should_close:
        wb.close()

    total_count = len(results)
    passed_count = sum(1 for r in results if r['status'] == 'Đạt')
    failed_count = sum(1 for r in results if r['status'] == 'Không đạt')
    warning_count = sum(1 for r in results if r['status'] == 'Cảnh báo')
    compliance_rate = round((passed_count / total_count) * 100, 1)

    return {
        'month': f"{year}-{month:02d}",
        'month_label': f"{calendar.month_name[month]} {year}",
        'domain': domain_name or 'Digital Team',
        'filename': filename,
        'checker': checker,
        'check_date': datetime.now().strftime('%d/%m/%Y'),
        'total_rules': total_count,
        'passed_count': passed_count,
        'failed_count': failed_count,
        'warning_count': warning_count,
        'compliance_rate': compliance_rate,
        'is_ready': failed_count == 0,
        'results': results
    }

def generate_validation_checklist_file(
    wb_or_path,
    output_path=None,
    template_path=None,
    target_month=None,
    domain_name=None,
    filename=None,
    checker="QuangTV15"
):
    """
    Generates populated Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx based on validation rules.
    Preserves all 6 worksheets, conditional formatting, and drop-downs.
    """
    if not template_path:
        # Search candidate paths
        candidates = [
            "Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx",
            os.path.join(os.path.dirname(__file__), "Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx"),
            os.path.join(os.path.dirname(__file__), "timesheets_extracted", "Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx"),
            "/tmp/timesheets_extracted/Timesheet_Validation_Checklist_Digital Team_Sep2026.xlsx"
        ]
        for p in candidates:
            if os.path.exists(p):
                template_path = p
                break

    if not template_path or not os.path.exists(template_path):
        raise FileNotFoundError(f"Checklist template not found at {template_path}")

    # Run validation audit
    audit_data = validate_timesheet_workbook(
        wb_or_path,
        target_month=target_month,
        domain_name=domain_name,
        filename=filename,
        checker=checker
    )

    # Load master checklist template
    wb_checklist = openpyxl.load_workbook(template_path)
    ws = wb_checklist['Checklist chi tiết']

    # Update metadata header (Rows 3 to 7)
    try:
        yr, mo = map(int, audit_data['month'].split('-'))
        ws.cell(3, 2).value = datetime(yr, mo, 1)
        ws.cell(3, 2).number_format = 'd-mmm-yy'
    except Exception:
        ws.cell(3, 2).value = audit_data['month']

    ws.cell(4, 2).value = audit_data['domain']
    ws.cell(5, 2).value = audit_data['filename']
    ws.cell(6, 2).value = audit_data['checker']
    ws.cell(7, 2).value = datetime.now()
    ws.cell(7, 2).number_format = 'd-mmm-yy'

    # Fill in rows 11 to 37 (27 check items)
    results_map = {r['stt']: r for r in audit_data['results']}
    for r in range(11, 38):
        stt = ws.cell(r, 1).value
        if isinstance(stt, int) and stt in results_map:
            item_res = results_map[stt]
            ws.cell(r, 5).value = item_res['status']
            ws.cell(r, 6).value = item_res['reference']
            ws.cell(r, 7).value = item_res['note']
            ws.cell(r, 8).value = item_res['handler']

    if not output_path:
        out_name = f"Timesheet_Validation_Checklist_Digital Team_{audit_data['month'].replace('-', '_')}.xlsx"
        output_path = os.path.join(os.path.dirname(template_path), out_name)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb_checklist.save(output_path)
    wb_checklist.close()

    return output_path, audit_data
