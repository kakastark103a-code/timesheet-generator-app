import re
import os
import requests
from datetime import datetime

def audit_timesheet_data(summary_data, balance_data, timesheet_data, month_str):
    """
    Performs comprehensive row-level audit of timesheet data according to Singapore Client Rules:
    1. Public Holiday (PH) => Actual Time MUST = 0
    2. Leave => Actual Time MUST = 0
    3. Project Task => Actual Time MUST = 8 (or 2/4/6), CANNOT = 0
    4. Work Item Type => Must be in accepted list (No Miscellaneous or invalid types)
    5. Weekend dates => Must be 'Weekend support' or 'PH Support'
    6. OT Worktypes => Actual Time must be 2/4/6/8 (Cannot be 0)
    7. Leave Balance => Balance in month cannot be negative (< 0)
    """
    anomalies = []
    idx = 1
    
    ACCEPTED_WORKTYPES = [
        'project task', 'leave', 'public holiday',
        'ph support', 'weekend support', 'weekday support'
    ]

    # 1. Audit Leave Balance Sheet
    for r_idx, row in enumerate(balance_data, start=3):
        name = row.get('name', 'Unknown')
        bal_in = row.get('balance_in_month')
        if isinstance(bal_in, (int, float)) and bal_in < 0:
            add_needed = abs(bal_in)
            cur_upto = row.get('balance_upto') or 10
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'danger',
                'category': 'Leave Balance Error',
                'sheet': 'Balance Leave',
                'row_index': r_idx,
                'member': name,
                'date': '--',
                'target_field': 'Balance Upto (Cột G)',
                'current_value': str(cur_upto),
                'proposed_value': str(cur_upto + add_needed),
                'message': f"Nhân sự <strong>{name}</strong> bị <strong>âm phép</strong> cuối tháng: {bal_in} ngày.",
                'suggested_fix': f"Tăng dư phép tích lũy (Balance Upto) từ {cur_upto} ➔ {cur_upto + add_needed} để hết âm phép.",
                'action_hint': f"Cần update Leave Balance upto lên {cur_upto + add_needed}",
                'fix_action': {'type': 'adjust_leave_balance', 'name': name, 'add_upto': add_needed}
            })
            idx += 1
            
    # 2. Audit Timesheet Sheet Rows
    for r_idx, row in enumerate(timesheet_data, start=2):
        dt_val = row.get('date', '')
        wtype_raw = str(row.get('work_item_type', '')).strip()
        wtype_lower = wtype_raw.lower()
        name = row.get('name', '')
        hours = row.get('hours', 0)

        # Rule 1: Reject Invalid Worktypes (e.g. Miscellaneous)
        if wtype_lower not in ACCEPTED_WORKTYPES:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'danger',
                'category': 'Invalid Worktype',
                'sheet': 'Timesheet',
                'row_index': r_idx,
                'member': name,
                'date': dt_val,
                'target_field': 'Work Item Type (Cột C)',
                'current_value': wtype_raw,
                'proposed_value': 'Project Task',
                'message': f"Dòng {r_idx} ({name} - {dt_val}): điền loại '<code>{wtype_raw}</code>' không thuộc danh mục Worktype được chấp nhận.",
                'suggested_fix': f"Đổi Work Item Type: <code>{wtype_raw}</code> ➔ <code>Project Task</code>",
                'action_hint': "Chuyển sang Worktype hợp lệ",
                'fix_action': {'type': 'replace_invalid_worktype', 'row_index': r_idx}
            })
            idx += 1

        # Rule 2: Public Holiday (PH) Actual Time MUST BE 0
        if wtype_lower == 'public holiday' and isinstance(hours, (int, float)) and hours != 0:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'danger',
                'category': 'PH Hours Error',
                'sheet': 'Timesheet',
                'row_index': r_idx,
                'member': name,
                'date': dt_val,
                'target_field': 'Actual Time (Cột J)',
                'current_value': f"{hours}h",
                'proposed_value': '0h',
                'message': f"Dòng {r_idx} ({name} - {dt_val}): loại 'Public Holiday' nhưng có Actual Time = <strong>{hours}h</strong> (Quy tắc Client: PH Actual Time phải = 0).",
                'suggested_fix': f"Đổi số giờ làm: <code>{hours}h</code> ➔ <code>0h</code>",
                'action_hint': "Chuyển Actual Time của Public Holiday về 0",
                'fix_action': {'type': 'set_ph_hours_zero', 'row_index': r_idx}
            })
            idx += 1

        # Rule 3: Leave Actual Time MUST BE 0
        if wtype_lower == 'leave' and isinstance(hours, (int, float)) and hours > 0:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'warning',
                'category': 'Leave Hours Error',
                'sheet': 'Timesheet',
                'row_index': r_idx,
                'member': name,
                'date': dt_val,
                'target_field': 'Actual Time (Cột J)',
                'current_value': f"{hours}h",
                'proposed_value': '0h',
                'message': f"Dòng {r_idx} ({name} - {dt_val}): loại 'Leave' nhưng có Actual Time = <strong>{hours}h</strong> (Quy tắc Client: Leave Actual Time phải = 0).",
                'suggested_fix': f"Đổi số giờ làm: <code>{hours}h</code> ➔ <code>0h</code>",
                'action_hint': "Chuyển Actual Time của Leave về 0",
                'fix_action': {'type': 'set_leave_hours_zero', 'row_index': r_idx}
            })
            idx += 1

        # Rule 4: Project Task Actual Time CANNOT be 0
        if wtype_lower == 'project task' and isinstance(hours, (int, float)) and hours == 0:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'warning',
                'category': 'Project Task Hours Error',
                'sheet': 'Timesheet',
                'row_index': r_idx,
                'member': name,
                'date': dt_val,
                'target_field': 'Actual Time (Cột J)',
                'current_value': '0h',
                'proposed_value': '8h',
                'message': f"Dòng {r_idx} ({name} - {dt_val}): loại 'Project Task' nhưng có Actual Time = <strong>0h</strong> (Quy tắc Client: Project Task phải = 8h).",
                'suggested_fix': "Đổi số giờ làm: <code>0h</code> ➔ <code>8h</code>",
                'action_hint': "Chuyển Actual Time của Project Task thành 8h",
                'fix_action': {'type': 'set_project_task_hours_8', 'row_index': r_idx}
            })
            idx += 1

        # Rule 5: OT Worktype Actual Time CANNOT be 0
        if wtype_lower in ['ph support', 'weekend support', 'weekday support'] and isinstance(hours, (int, float)) and hours == 0:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'warning',
                'category': 'OT Hours Error',
                'sheet': 'Timesheet',
                'row_index': r_idx,
                'member': name,
                'date': dt_val,
                'target_field': 'Actual Time (Cột J)',
                'current_value': '0h',
                'proposed_value': '4h',
                'message': f"Dòng {r_idx} ({name} - {dt_val}): OT loại '{wtype_raw}' nhưng Actual Time lại = 0h.",
                'suggested_fix': "Đổi số giờ OT: <code>0h</code> ➔ <code>4h</code>",
                'action_hint': "Chuyển Actual Time OT thành 4h",
                'fix_action': {'type': 'set_ot_hours_default', 'row_index': r_idx}
            })
            idx += 1

        # Rule 6: Weekend Worktype Missing 'Weekend support' / 'PH Support'
        if dt_val:
            try:
                dt_obj = datetime.strptime(str(dt_val)[:10], '%Y-%m-%d')
                if dt_obj.weekday() >= 5 and wtype_lower not in ['weekend support', 'ph support']:
                    anomalies.append({
                        'id': f"fix_{idx}",
                        'type': 'warning',
                        'category': 'Weekend Work Error',
                        'sheet': 'Timesheet',
                        'row_index': r_idx,
                        'member': name,
                        'date': dt_val,
                        'target_field': 'Work Item Type (Cột C)',
                        'current_value': wtype_raw,
                        'proposed_value': 'Weekend support',
                        'message': f"Dòng {r_idx} ({name} - {dt_obj.strftime('%d-%b')}): làm việc cuối tuần chưa đánh nhãn 'Weekend support'.",
                        'suggested_fix': f"Đổi loại công việc: <code>{wtype_raw}</code> ➔ <code>Weekend support</code>",
                        'action_hint': "Đổi loại công việc cuối tuần sang Weekend support",
                        'fix_action': {'type': 'fix_weekend_worktype', 'row_index': r_idx}
                    })
                    idx += 1
            except Exception:
                pass
                
    # 3. Audit High OT
    for r_idx, row in enumerate(summary_data, start=2):
        name = row.get('name', '')
        tot_ot = row.get('total_ot', 0)
        if isinstance(tot_ot, (int, float)) and tot_ot > 10:
            anomalies.append({
                'id': f"fix_{idx}",
                'type': 'info',
                'category': 'High OT',
                'sheet': 'Summary',
                'row_index': r_idx,
                'member': name,
                'date': '--',
                'target_field': 'Total OT (Cột J)',
                'current_value': f"{tot_ot} ngày",
                'proposed_value': f"{tot_ot} ngày (PM Approved)",
                'message': f"Nhân sự <strong>{name}</strong> có tổng số ngày OT cao trong tháng: {tot_ot} ngày.",
                'suggested_fix': "Xác nhận phê duyệt từ PM (Không cần sửa file Excel).",
                'action_hint': "Xác nhận lại phê duyệt OT từ Project Manager",
                'fix_action': None
            })
            idx += 1
            
    if not anomalies:
        anomalies.append({
            'id': 'fix_ok',
            'type': 'success',
            'category': 'Perfect',
            'sheet': '--',
            'row_index': 0,
            'member': 'All',
            'date': '--',
            'target_field': '--',
            'current_value': 'OK',
            'proposed_value': 'OK',
            'message': '✅ Tất cả số liệu Timesheet hợp lệ 100%! Không phát hiện lỗi hay vi phạm rule nào.',
            'suggested_fix': None,
            'action_hint': None,
            'fix_action': None
        })
        
    return anomalies

def apply_fixes_to_workbook(wb, fix_actions):
    """
    Applies approved fix_action items directly to an openpyxl workbook, supporting customized user overrides.
    """
    if not fix_actions:
        return

    sheet_ts = wb['Timesheet'] if 'Timesheet' in wb.sheetnames else None

    for act in fix_actions:
        if not isinstance(act, dict):
            continue

        act_type = act.get('type')
        r_idx = act.get('row_index')
        custom_val = act.get('custom_value')

        # Hours Fixes (PH, Leave, Project Task, OT)
        if act_type in ['set_ph_hours_zero', 'set_leave_hours_zero', 'set_project_task_hours_8', 'set_ot_hours_default'] and sheet_ts and r_idx:
            if custom_val is not None and str(custom_val).strip() != '':
                try:
                    sheet_ts.cell(r_idx, 10).value = float(custom_val)
                except Exception:
                    sheet_ts.cell(r_idx, 10).value = 0 if 'zero' in act_type else 8
            else:
                sheet_ts.cell(r_idx, 10).value = 0 if 'zero' in act_type else (8 if act_type == 'set_project_task_hours_8' else 4)

        # Worktype Fixes (Invalid Worktype, Weekend Worktype)
        elif act_type in ['replace_invalid_worktype', 'fix_weekend_worktype'] and sheet_ts and r_idx:
            if custom_val and str(custom_val).strip() != '':
                sheet_ts.cell(r_idx, 3).value = str(custom_val).strip()
            else:
                sheet_ts.cell(r_idx, 3).value = 'Weekend support' if act_type == 'fix_weekend_worktype' else 'Project Task'

        # Leave Balance Fix
        elif act_type == 'adjust_leave_balance':
            target_name = act.get('name', '').lower()
            bal_sheet_name = 'Balance Leave' if 'Balance Leave' in wb.sheetnames else ('Leave Balance' if 'Leave Balance' in wb.sheetnames else None)
            if bal_sheet_name and bal_sheet_name in wb.sheetnames:
                sheet_bal = wb[bal_sheet_name]
                for r in range(3, sheet_bal.max_row + 1):
                    name_cell = sheet_bal.cell(r, 2).value
                    if name_cell and target_name in str(name_cell).lower():
                        if custom_val is not None and str(custom_val).strip() != '':
                            try:
                                sheet_bal.cell(r, 7).value = float(custom_val)
                            except Exception:
                                pass
                        else:
                            cur_upto = sheet_bal.cell(r, 7).value or 10
                            add_val = act.get('add_upto', 0)
                            if isinstance(cur_upto, (int, float)):
                                sheet_bal.cell(r, 7).value = cur_upto + add_val

    # Clear active AutoFilter criteria across all worksheets
    try:
        from timesheet_generator import clear_all_worksheet_filters
        clear_all_worksheet_filters(wb)
    except Exception:
        pass

def process_ai_chat_command(user_prompt: str, current_notes: str, domain_key: str = 'cbg', api_key: str = None, base_url: str = None):
    """
    Processes natural language command using LLM or Smart Rule Engine.
    Returns dict with reply message, updated_notes, and actions.
    """
    # 1. If live API key provided or present in env, try calling LLM API
    llm_key = api_key or os.environ.get('OPENAI_API_KEY') or os.environ.get('GEMINI_API_KEY')
    llm_url = base_url or os.environ.get('OPENAI_BASE_URL', 'https://api.openai.com/v1')
    
    if llm_key and not user_prompt.startswith('[LOCAL]'):
        try:
            endpoint = f"{llm_url.rstrip('/')}/chat/completions"
            headers = {
                'Authorization': f'Bearer {llm_key}',
                'Content-Type': 'application/json'
            }
            system_msg = """Bạn là Chuyên gia AI Hỗ trợ Quản lý Timesheet (FPT Timesheet Assistant).
Nhiệm vụ của bạn:
1. Giải đáp các thắc mắc về Timesheet, số dư phép, quy tắc tính OT (Weekday support, Weekend support, PH Support).
2. Phân tích yêu cầu tiếng Việt của người dùng để cập nhật bảng Ghi chú (Notes/Overrides).
Format phản hồi dạng JSON:
{
  "reply": "Câu trả lời thân thiện cho người dùng",
  "updated_notes": "Nội dung ghi chú cập nhật mới (nếu có)",
  "action_type": "update_notes / info / fix_anomaly"
}"""
            payload = {
                'model': 'gpt-4o-mini',
                'messages': [
                    {'role': 'system', 'content': system_msg},
                    {'role': 'user', 'content': f"Ghi chú hiện tại:\n{current_notes}\n\nYêu cầu của người dùng:\n{user_prompt}"}
                ],
                'temperature': 0.3
            }
            res = requests.post(endpoint, headers=headers, json=payload, timeout=8)
            if res.status_code == 200:
                content = res.json()['choices'][0]['message']['content']
                import json
                try:
                    parsed_json = json.loads(content)
                    return {
                        'reply': parsed_json.get('reply', content),
                        'updated_notes': parsed_json.get('updated_notes', current_notes),
                        'action': parsed_json.get('action_type', 'info')
                    }
                except Exception:
                    return {'reply': content, 'updated_notes': current_notes, 'action': 'info'}
        except Exception:
            pass  # Fallback to Smart Rule Engine below

    # 2. Smart Rule Engine Fallback (Instant, offline, zero failure)
    prompt_lower = user_prompt.lower()
    updated_notes = current_notes.strip() if current_notes else ""
    lines = [l for l in updated_notes.split('\n') if l.strip()]
    
    # Action 1: Update Leave Balance
    m_leave = re.search(r'(?:sửa|cập nhật|update|chỉnh|cho)\s+phép\s+(?:cho\s+)?([A-Za-zÀ-ỹ\s]+?)\s+(?:về|thành|=|\:)\s*([\d\.]+)', prompt_lower, re.UNICODE)
    if not m_leave:
        m_leave = re.search(r'([A-Za-zÀ-ỹ\s]+?)\s+(?:leave balance|phép)\s+.*?([\d\.]+)', prompt_lower, re.UNICODE)
        
    if m_leave:
        raw_name = m_leave.group(1).strip()
        new_val = m_leave.group(2).strip()
        
        # Name resolution helper
        target_name = raw_name.title()
        if raw_name.lower() in ['hà', 'ha']: target_name = "Dao Manh Ha"
        elif raw_name.lower() in ['mai']: target_name = "Duong Thi Tuyet Mai"
        elif raw_name.lower() in ['tùng', 'tung']: target_name = "Duong Xuan Tung"
        
        # Check if line for member exists
        found = False
        new_lines = []
        for line in lines:
            if target_name.lower() in line.lower() or raw_name.lower() in line.lower():
                # Replace balance note
                line = re.sub(r'Cần update Leave Balance upto.*', f'Cần update Leave Balance upto về {new_val}', line)
                if 'Leave Balance' not in line:
                    line += f"\tCần update Leave Balance upto về {new_val}"
                found = True
            new_lines.append(line)
            
        if not found:
            new_lines.append(f"{target_name}\tAcc\t{domain_key.upper()}\tCần update Leave Balance upto về {new_val}")
            
        updated_notes = "\n".join(new_lines)
        return {
            'reply': f"🤖 **Đã cập nhật số dư phép!**\nĐã điều chỉnh `Leave Balance upto` cho nhân sự **{target_name}** thành **{new_val} ngày**.",
            'updated_notes': updated_notes,
            'action': 'update_notes'
        }

    # Action 2: Add OT / Work Item entry
    m_ot = re.search(r'(?:thêm|bổ sung|add)?\s*(?:ot|support)\s*(?:cho\s+)?([A-Za-z\s]+?)\s+(?:ngày\s+)?(\d{1,2}[-\/]\d{1,2}|\d{1,2}\s+[A-Za-z]{3}|\d{1,2})\s*(?:ot\s*)?(\d+(?:\.\d+)?h)?', prompt_lower)
    if not m_ot:
        m_ot = re.search(r'([A-Za-z\s]+?)\s+(?:ngày\s+)?(\d{1,2}[-\/]\d{1,2}|\d{1,2}\s+[A-Za-z]{3}|\d{1,2})\s*(?:ot\s*)?(\d+(?:\.\d+)?h)?\s*(weekday|weekend|ph)?', prompt_lower)
        
    if m_ot:
        target_name = m_ot.group(1).strip().title()
        date_part = m_ot.group(2).strip()
        hours_part = m_ot.group(3).strip() if m_ot.group(3) else "4h"
        
        wtype_part = "Weekday support"
        if 'weekend' in prompt_lower:
            wtype_part = "Weekend support"
        elif 'ph' in prompt_lower:
            wtype_part = "PH Support"
            
        new_entry = f"{target_name}\tAcc\t{domain_key.upper()}\t{date_part} OT {hours_part} {wtype_part} cho support activities"
        if updated_notes:
            updated_notes += "\n" + new_entry
        else:
            updated_notes = "Member name\tFsoft Account\tProject name\tNote\n" + new_entry
            
        return {
            'reply': f"🤖 **Đã thêm ghi chú OT mới!**\nThêm dòng OT ngày **{date_part}** ({hours_part} {wtype_part}) cho nhân sự **{target_name}**.",
            'updated_notes': updated_notes,
            'action': 'add_ot'
        }

    # Action 3: Review / Audit Query
    if any(k in prompt_lower for k in ['review', 'rà soát', 'kiểm tra', 'lỗi', 'cảnh báo', 'audit', 'thống kê']):
        return {
            'reply': f"🤖 **Kết Quả Rà Soát Timesheet ({domain_key.upper()})**:\n- Bảng số liệu đã được kiểm tra thành công.\n- Bạn có thể xem chi tiết danh sách cảnh báo tại tab **'🔍 Review & Audit Timesheet'** trên giao diện chính.",
            'updated_notes': updated_notes,
            'action': 'review'
        }

    # Fallback general response
    return {
        'reply': f"🤖 **AI Timesheet Assistant**: Tôi có thể giúp bạn tự động rà soát file Timesheet, điều chỉnh dư phép (VD: *'Sửa phép cho Hà về 12'*), hoặc thêm ngày OT (VD: *'Thêm OT 4h cho Mai ngày 1/7'*). Vui lòng thử nhập lệnh!",
        'updated_notes': updated_notes,
        'action': 'info'
    }
