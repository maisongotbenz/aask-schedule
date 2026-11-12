import os
import json
import re
import requests
import pandas as pd
from datetime import datetime
from io import StringIO

def download_and_parse(date_str):
    # Прямая ссылка на файл дневного отделения (DO)
    url = f"https://www.altask.ru/images/raspisanie/DO/{date_str}.xls"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    
    print(f"Пытаемся скачать: {url}")
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            print(f"Файл {date_str}.xls не найден (Статус: {r.status_code})")
            return None
        
        print(f"Файл скачан, размер: {len(r.content)} байт")
        
        # Читаем как Excel
        try:
            df = pd.read_excel(StringIO(r.content.decode('cp1251')), header=None)
            print("Excel успешно прочитан.")
        except Exception as e:
            print(f"Не удалось прочитать как Excel: {e}")
            return None
            
        return df
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        return None

def parse_dataframe(df, date_key):
    try:
        # Ищем заголовок с датой
        title = ""
        for i in range(min(10, len(df))):
            row_str = " ".join([str(x) for x in df.iloc[i].values if not pd.isna(x)])
            if "неделя" in row_str.lower():
                title = row_str
                break
        
        week = int(re.search(r'(\d+)\s*неделя', title).group(1)) if re.search(r'(\d+)\s*неделя', title) else 1
        weekday = "четверг"
        for d in ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота']:
            if d in title.lower(): weekday = d; break
        
        groups = {}
        
        # Ищем все блоки с группами (где "№ пары")
        block_starts = [r for r in range(len(df)) if str(df.iloc[r, 0]).strip() == "№ пары"]
        print(f"Найдено блоков с группами: {len(block_starts)}")
        
        for i, start in enumerate(block_starts):
            end = block_starts[i+1] if i+1 < len(block_starts) else len(df)
            
            # Собираем названия групп
            cols = []
            for c in range(1, len(df.columns)):
                val = df.iloc[start, c]
                if not pd.isna(val):
                    gname = str(val).strip()
                    if gname and gname != "nan":
                        cols.append((c, gname))
            
            print(f"  В блоке {i+1} найдено групп: {len(cols)}")
            
            # Собираем пары
            r = start + 1
            while r < end - 1:
                num_str = str(df.iloc[r, 0]).strip()
                if num_str.isdigit() and 1 <= int(num_str) <= 7:
                    num = int(num_str)
                    for c, gname in cols:
                        subj = str(df.iloc[r, c]).strip() if not pd.isna(df.iloc[r, c]) else ""
                        t_info = str(df.iloc[r+1, c]).strip() if not pd.isna(df.iloc[r+1, c]) else ""
                        
                        if subj or t_info:
                            t = t_info
                            rm = ""
                            p = None
                            if "(п/п)" in t or "подгруппа" in t.lower(): p = 1
                            room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', t)
                            if room_match:
                                rm = room_match.group(1)
                                t = t.replace(room_match.group(0), '').strip()
                            t = t.replace('(п/п)', '').replace('(подгруппа)', '').strip()
                            
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
        print(f"Ошибка парсинга: {e}")
        return None

def main():
    all_days = {}
    all_groups = set()
    
    now = datetime.now()
    # Пробуем скачать расписание на 3 дня вперед
    for i in range(3):
        target = now + pd.Timedelta(days=i)
        date_str = target.strftime("%d%m")
        date_key = target.strftime("%Y-%m-%d")
        
        print(f"\n--- Пробуем {date_str} ({date_key}) ---")
        df = download_and_parse(date_str)
        
        if df is not None:
            day_data = parse_dataframe(df, date_key)
            if day_data and day_data['groups']:
                all_days[date_key] = day_data
                all_groups.update(day_data['groups'].keys())
                print(f"  -> Успешно добавлен день {date_key} (групп: {len(day_data['groups'])})")
            else:
                print(f"  -> Не удалось извлечь данные из файла {date_str}.xls")

    if not all_days:
        print("\n!!! Нет данных для сохранения.")
        return

    res = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(all_groups)),
        "days": all_days
    }
    
    with open('schedule.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(f"\nГотово! Сохранено дней: {len(all_days)}")

if __name__ == "__main__":
    main()
