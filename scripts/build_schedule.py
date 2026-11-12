import os
import json
import re
import requests
import xlrd
from datetime import datetime

# Список файлов, которые мы скачаем из вашего репозитория
FILES = [
    {"file": "1211.xls", "date": "2026-11-12"},
    {"file": "1311.xls", "date": "2026-11-13"},
]

def download_file(filename):
    url = f"https://raw.githubusercontent.com/maisongotbenz/aask-schedule/main/{filename}"
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            path = f"/tmp/{filename}"
            with open(path, 'wb') as f: f.write(r.content)
            return path
    except Exception as e:
        print(f"Ошибка загрузки {filename}: {e}")
    return None

def parse_teacher(cell):
    if not cell: return "", "", None
    room = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', cell)
    room = room.group(1) if room else ""
    p = 1 if "(п/п)" in cell or "подгруппа" in cell.lower() else None
    teacher = re.sub(r'\s*\d+[а-яА-Я]?\s*(?:\(п/п\)|\(подгруппа\))?\s*$', '', cell).strip()
    return teacher, room, p

def parse_excel(path, date_key):
    if not path or not os.path.exists(path): return None
    try:
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        
        # Берем заголовок из первой строки
        title = str(sh.cell_value(0, 8))
        week = int(re.search(r'(\d+)\s*неделя', title).group(1)) if re.search(r'(\d+)\s*неделя', title) else 1
        
        weekday = "четверг"
        for d in ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота']:
            if d in title.lower(): weekday = d; break
        
        groups = {}
        starts = [r for r in range(sh.nrows) if str(sh.cell_value(r, 0)).strip() == "№ пары"]
        
        for i, start in enumerate(starts):
            end = starts[i+1] if i+1 < len(starts) else sh.nrows
            cols = [(c, str(sh.cell_value(start, c)).strip()) for c in range(1, sh.ncols) if str(sh.cell_value(start, c)).strip()]
            
            r = start + 1
            while r < end:
                num_str = str(sh.cell_value(r, 0)).strip()
                if num_str.isdigit() and 1 <= int(num_str) <= 7:
                    num = int(num_str)
                    for c, gname in cols:
                        subj = str(sh.cell_value(r, c)).strip()
                        t_info = str(sh.cell_value(r+1, c)).strip()
                        if subj or t_info:
                            t, rm, p = parse_teacher(t_info)
                            item = {"n": num, "s": subj, "t": t, "r": rm}
                            if p: item["p"] = p
                            if gname not in groups: groups[gname] = []
                            if not any(x["n"] == num and x["s"] == subj for x in groups[gname]):
                                groups[gname].append(item)
                    r += 2
                else: r += 1
                
        return {"weekday": weekday, "week": week, "groups": groups}
    except Exception as e:
        print(f"Ошибка парсинга {path}: {e}")
        return None

def main():
    all_days = {}
    all_groups = set()
    
    for item in FILES:
        print(f"Скачиваем {item['file']}...")
        path = download_file(item['file'])
        if path:
            day_data = parse_excel(path, item['date'])
            if day_data and day_data['groups']:
                all_days[item['date']] = day_data
                all_groups.update(day_data['groups'].keys())
                print(f"  -> {item['date']} успешно добавлен (групп: {len(day_data['groups'])})")
    
    if not all_days:
        print("Нет данных.")
        return

    res = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(all_groups)),
        "days": all_days
    }
    
    with open('schedule.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(f"Готово! Сохранено дней: {len(all_days)}")

if __name__ == "__main__":
    main()
