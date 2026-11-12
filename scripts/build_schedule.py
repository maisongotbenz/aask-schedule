import os
import json
import re
import requests
import xlrd
from datetime import datetime, timedelta

# ВАЖНО: Если файл лежит в корне и называется 1311.xls
# Если вы закинули его в другую папку, измените ссылку (добавьте /имя_папки/)
LOCAL_EXCEL_URL = "https://raw.githubusercontent.com/maisongotbenz/aask-schedule/main/1311.xls"

def download_file(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            file_path = "/tmp/schedule.xls"
            with open(file_path, 'wb') as f:
                f.write(response.content)
            return file_path
    except Exception as e:
        print(f"Ошибка скачивания: {e}")
    return None

def parse_teacher_info(cell_text):
    if not cell_text: return "", "", None
    room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', cell_text)
    room = room_match.group(1) if room_match else ""
    subgroup = 1 if "(п/п)" in cell_text or "подгруппа" in cell_text.lower() else None
    if "(подгруппа)" in cell_text.lower(): subgroup = 1 
    teacher = re.sub(r'\s*\d+[а-яА-Я]?\s*(?:\(п/п\)|\(подгруппа\))?\s*$', '', cell_text).strip()
    return teacher, room, subgroup

def parse_excel_for_day(file_path, target_date_key):
    if not file_path or not os.path.exists(file_path): return None
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        
        # Вытаскиваем день и неделю из заголовка
        title = str(sheet.cell_value(0, 8))
        week_match = re.search(r'(\d+)\s*неделя', title)
        week_num = int(week_match.group(1)) if week_match else 1
        
        day_name_raw = "четверг"
        for d in ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота']:
            if d in title.lower():
                day_name_raw = d
                break
        
        groups_data = {}
        block_starts = [r for r in range(sheet.nrows) if str(sheet.cell_value(r, 0)).strip() == "№ пары"]
        
        for b_idx, start_row in enumerate(block_starts):
            end_row = block_starts[b_idx + 1] if b_idx + 1 < len(block_starts) else sheet.nrows
            groups_in_block = []
            for c in range(1, sheet.ncols):
                group_name = str(sheet.cell_value(start_row, c)).strip()
                if group_name: groups_in_block.append((c, group_name))
            
            r = start_row + 1
            while r < end_row:
                para_num_str = str(sheet.cell_value(r, 0)).strip()
                if para_num_str.isdigit() and 1 <= int(para_num_str) <= 7:
                    para_num = int(para_num_str)
                    for col_idx, group_name in groups_in_block:
                        subject = str(sheet.cell_value(r, col_idx)).strip()
                        teacher_info = str(sheet.cell_value(r + 1, col_idx)).strip()
                        if subject or teacher_info:
                            teacher, room, subgroup = parse_teacher_info(teacher_info)
                            item = {"n": para_num, "s": subject, "t": teacher, "r": room}
                            if subgroup: item["p"] = subgroup
                            if group_name not in groups_data: groups_data[group_name] = []
                            if not any(d["n"] == para_num and d["s"] == subject for d in groups_data[group_name]):
                                groups_data[group_name].append(item)
                    r += 2
                else:
                    r += 1
        return {"weekday": day_name_raw, "week": week_num, "groups": groups_data}
    except Exception as e:
        print(f"Ошибка парсинга Excel: {e}")
        return None

def main():
    print("Скачиваем Excel из вашего репозитория...")
    file_path = download_file(LOCAL_EXCEL_URL)
    
    if not file_path:
        print("Не удалось скачать файл. Проверьте ссылку!")
        return

    print("Файл скачан, парсим...")
    # Берем дату из имени файла (если 1311, то 13 ноября 2026)
    date_key = "2026-11-13" 
    day_data = parse_excel_for_day(file_path, date_key)
    
    if not day_data or not day_data["groups"]:
        print("Не удалось извлечь данные из файла.")
        return

    print(f"Успешно распарсено групп: {len(day_data['groups'])}")
    
    # Формируем итоговый JSON
    result = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(day_data["groups"].keys())),
        "days": { date_key: day_data }
    }
    
    with open('schedule.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("Успешно сохранено в schedule.json")

if __name__ == "__main__":
    main()
