import os
import json
import re
import requests
import xlrd
from datetime import datetime, timedelta

BASE_URL = "https://altask.ru/DO/"

def download_file(date_str):
    filename = f"{date_str}.xls"
    url = f"{BASE_URL}{filename}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            file_path = f"/tmp/{filename}"
            with open(file_path, 'wb') as f:
                f.write(response.content)
            return file_path
    except Exception:
        pass
    return None

def parse_teacher_info(cell_text):
    if not cell_text:
        return "", "", None
    room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', cell_text)
    room = room_match.group(1) if room_match else ""
    subgroup = 1 if "(п/п)" in cell_text or "подгруппа" in cell_text.lower() else None
    if "(подгруппа)" in cell_text.lower():
        subgroup = 1 
    teacher = re.sub(r'\s*\d+[а-яА-Я]?\s*(?:\(п/п\)|\(подгруппа\))?\s*$', '', cell_text).strip()
    return teacher, room, subgroup

def parse_excel_for_day(file_path, target_date_key):
    """Парсит один файл и возвращает данные только для одного дня"""
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        
        title = sheet.cell_value(0, 8)
        week_match = re.search(r'(\d+)\s*неделя', title)
        week_num = int(week_match.group(1)) if week_match else 1
        
        # Пытаемся вытащить название дня недели из заголовка
        day_name_raw = "четверг"
        for d in ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье']:
            if d in title.lower():
                day_name_raw = d
                break
        
        groups_data = {}
        block_starts = []
        for r in range(sheet.nrows):
            if str(sheet.cell_value(r, 0)).strip() == "№ пары":
                block_starts.append(r)
                
        for b_idx, start_row in enumerate(block_starts):
            end_row = block_starts[b_idx + 1] if b_idx + 1 < len(block_starts) else sheet.nrows
            groups_in_block = []
            for c in range(1, sheet.ncols):
                group_name = str(sheet.cell_value(start_row, c)).strip()
                if group_name and group_name != "":
                    groups_in_block.append((c, group_name))
            
            r = start_row + 1
            while r < end_row:
                para_num_str = str(sheet.cell_value(r, 0)).strip()
                if para_num_str.isdigit() and 1 <= int(para_num_str) <= 7:
                    para_num = int(para_num_str)
                    subjects_row = r
                    teachers_row = r + 1
                    if teachers_row >= end_row:
                        break
                    for col_idx, group_name in groups_in_block:
                        subject = str(sheet.cell_value(subjects_row, col_idx)).strip()
                        teacher_info = str(sheet.cell_value(teachers_row, col_idx)).strip()
                        if subject or teacher_info:
                            teacher, room, subgroup = parse_teacher_info(teacher_info)
                            item = {"n": para_num, "s": subject, "t": teacher, "r": room}
                            if subgroup:
                                item["p"] = subgroup
                            if group_name not in groups_data:
                                groups_data[group_name] = []
                            if not any(d["n"] == para_num and d["s"] == subject for d in groups_data[group_name]):
                                groups_data[group_name].append(item)
                    r += 2
                else:
                    r += 1

        return {
            "weekday": day_name_raw,
            "week": week_num,
            "groups": groups_data
        }
    except Exception as e:
        print(f"Ошибка при парсинге {file_path}: {e}")
        return None

def main():
    today = datetime.now()
    all_days = {}
    all_groups = set()
    
    print("Начинаем сбор расписания на неделю...")
    
    # Пробуем скачать расписание на 7 дней вперед (включая сегодня)
    for i in range(7):
        target_date = today + timedelta(days=i)
        date_str = target_date.strftime("%d%m")
        date_key = target_date.strftime("%Y-%m-%d")
        
        print(f"Ищем файл за {date_str} ({date_key})...")
        file_path = download_file(date_str)
        
        if file_path:
            day_data = parse_excel_for_day(file_path, date_key)
            if day_data and day_data["groups"]:
                all_days[date_key] = day_data
                all_groups.update(day_data["groups"].keys())
                print(f"  -> Успешно добавлен день {date_key} (групп: {len(day_data['groups'])})")
        else:
            print(f"  -> Файл за {date_str} не найден.")

    if not all_days:
        print("Не удалось скачать ни одного дня. Прерываем.")
        return

    result = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(all_groups)),
        "days": all_days
    }
    
    output_path = 'schedule.json' 
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"Успешно сохранено в {output_path}. Всего дней: {len(all_days)}")

if __name__ == "__main__":
    main()
