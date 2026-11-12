import os
import json
import re
import requests
import pandas as pd
from datetime import datetime
from io import StringIO

# Базовый URL, откуда сайт колледжа отдает файлы
BASE_URL = "https://altask.ru/DO/"

def download_and_parse(date_str):
    """Скачивает файл и пытается его распарсить"""
    url = f"{BASE_URL}{date_str}.xls"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            print(f"Файл {date_str}.xls не найден (Статус: {r.status_code})")
            return None
            
        print(f"Файл {date_str}.xls скачан, размер: {len(r.content)} байт")
        
        # Пытаемся прочитать как Excel
        try:
            df = pd.read_excel(StringIO(r.content.decode('cp1251')), header=None)
        except Exception:
            # Если не получилось, пробуем как HTML (сайт часто отдает HTML под видом .xls)
            try:
                dfs = pd.read_html(StringIO(r.content.decode('cp1251', errors='ignore')), header=None)
                df = dfs[0]
            except Exception as e:
                print(f"Не удалось прочитать файл {date_str}.xls: {e}")
                return None
                
        return df
    except Exception as e:
        print(f"Ошибка при скачивании {date_str}.xls: {e}")
        return None

def parse_dataframe(df, date_key):
    try:
        # Ищем заголовок (дата, день недели)
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
        starts = []
        for r in range(len(df)):
            if str(df.iloc[r, 0]).strip() == "№ пары":
                starts.append(r)
        
        for i, start in enumerate(starts):
            end = starts[i+1] if i+1 < len(starts) else len(df)
            cols = [(c, str(df.iloc[start, c]).strip()) for c in range(1, len(df.columns)) if not pd.isna(df.iloc[start, c]) and str(df.iloc[start, c]).strip()]
            
            r = start + 1
            while r < end - 1:
                num_str = str(df.iloc[r, 0]).strip()
                if num_str.isdigit() and 1 <= int(num_str) <= 7:
                    num = int(num_str)
                    for c, gname in cols:
                        subj = str(df.iloc[r, c]).strip() if not pd.isna(df.iloc[r, c]) else ""
                        t_info = str(df.iloc[r+1, c]).strip() if not pd.isna(df.iloc[r+1, c]) else ""
                        
                        if subj or t_info:
                            # Простейшая очистка от аудитории и подгруппы
                            t = t_info
                            rm = ""
                            p = None
                            
                            # Ищем подгруппу
                            if "(п/п)" in t or "подгруппа" in t.lower():
                                p = 1
                            
                            # Ищем аудиторию (число в конце)
                            room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', t)
                            if room_match:
                                rm = room_match.group(1)
                                t = t.replace(room_match.group(0), '').strip()
                            
                            # Убираем лишнее из имени преподавателя
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
    # Пробуем скачать расписание на сегодня, завтра и послезавтра
    for i in range(3):
        target = now + pd.Timedelta(days=i)
        date_str = target.strftime("%d%m")
        date_key = target.strftime("%Y-%m-%d")
        
        print(f"--- Пробуем {date_str} ({date_key}) ---")
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
        print("!!! Нет данных. Возможно, сайт колледжа недоступен или файлы не выложены.")
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
