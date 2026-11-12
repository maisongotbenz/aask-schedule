import os
import json
import re
import requests
import pandas as pd
from datetime import datetime
from io import StringIO

FILES = [
    {"file": "1211.xls", "date": "2026-11-12"},
    {"file": "1311.xls", "date": "2026-11-13"},
]

def download_file(filename):
    url = f"https://raw.githubusercontent.com/maisongotbenz/aask-schedule/main/{filename}"
    print(f"Пытаемся скачать: {url}")
    try:
        r = requests.get(url, timeout=15)
        print(f"Статус ответа: {r.status_code}")
        if r.status_code == 200:
            path = f"/tmp/{filename}"
            with open(path, 'wb') as f: 
                f.write(r.content)
            print(f"Файл сохранен: {path} (размер: {len(r.content)} байт)")
            return path
        else:
            print(f"Ошибка! Сервер вернул статус {r.status_code}")
    except Exception as e:
        print(f"Исключение при скачивании {filename}: {e}")
    return None

def parse_teacher(cell):
    if not cell or pd.isna(cell): return "", "", None
    cell = str(cell).strip()
    room = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', cell)
    room = room.group(1) if room else ""
    p = 1 if "(п/п)" in cell or "подгруппа" in cell.lower() else None
    teacher = re.sub(r'\s*\d+[а-яА-Я]?\s*(?:\(п/п\)|\(подгруппа\))?\s*$', '', cell).strip()
    return teacher, room, p

def parse_file(path, date_key):
    if not path or not os.path.exists(path): return None
    try:
        # Пытаемся прочитать как Excel
        try:
            df = pd.read_excel(path, header=None)
        except Exception:
            # Если не получилось, пробуем как HTML (частая проблема .xls с сайтов)
            print("Не удалось прочитать как Excel, пробуем как HTML...")
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                html_content = f.read()
            dfs = pd.read_html(StringIO(html_content), header=None)
            if not dfs:
                raise ValueError("Не найдено таблиц в HTML")
            df = dfs[0] # Берем первую таблицу
        
        # Ищем заголовок
        title = ""
        for i in range(min(10, len(df))):
            row_str = " ".join([str(x) for x in df.iloc[i].values if not pd.isna(x)])
            if "неделя" in row_str.lower():
                title = row_str
                break
        
        week = 1
        week_match = re.search(r'(\d+)\s*неделя', title)
        if week_match: week = int(week_match.group(1))
        
        weekday = "четверг"
        for d in ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота']:
            if d in title.lower(): weekday = d; break
        
        groups = {}
        # Ищем строки, где первый столбец содержит "№ пары"
        starts = []
        for r in range(len(df)):
            if str(df.iloc[r, 0]).strip() == "№ пары":
                starts.append(r)
        
        for i, start in enumerate(starts):
            end = starts[i+1] if i+1 < len(starts) else len(df)
            
            # Получаем названия групп из строки start
            cols = []
            for c in range(1, len(df.columns)):
                gname = str(df.iloc[start, c]).strip()
                if gname and gname != "nan": 
                    cols.append((c, gname))
            
            r = start + 1
            while r < end - 1: # -1, так как смотрим r+1
                num_str = str(df.iloc[r, 0]).strip()
                if num_str.isdigit() and 1 <= int(num_str) <= 7:
                    num = int(num_str)
                    for c, gname in cols:
                        subj = str(df.iloc[r, c]).strip() if not pd.isna(df.iloc[r, c]) else ""
                        t_info = str(df.iloc[r+1, c]).strip() if not pd.isna(df.iloc[r+1, c]) else ""
                        
                        if subj or t_info:
                            t, rm, p = parse_teacher(t_info)
                            item = {"n": num, "s": subj, "t": t, "r": rm}
                            if p: item["p"] = p
                            if gname not in groups: groups[gname] = []
                            if not any(x["n"] == num and x["s"] == subj for x in groups[gname]):
                                groups[gname].append(item)
                    r += 2
                else: 
                    r += 1
        return {"weekday": weekday, "week": week, "groups": groups}
    except Exception as e:
        print(f"Критическая ошибка парсинга {path}: {e}")
        return None

def main():
    all_days = {}
    all_groups = set()
    
    for item in FILES:
        print(f"--- Обработка {item['file']} ---")
        path = download_file(item['file'])
        if path:
            day_data = parse_file(path, item['date'])
            if day_data and day_data['groups']:
                all_days[item['date']] = day_data
                all_groups.update(day_data['groups'].keys())
                print(f"  -> Успешно добавлен день {item['date']} (групп: {len(day_data['groups'])})")
            else:
                print(f"  -> Не удалось распарсить данные из {item['file']}")

    if not all_days:
        print("!!! Нет данных для сохранения.")
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
