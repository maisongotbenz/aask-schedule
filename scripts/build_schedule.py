import os
import json
import re
import requests
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

def download_and_parse(date_str):
    url = f"https://www.altask.ru/images/raspisanie/DO/{date_str}.xls"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    
    print(f"Пытаемся скачать: {url}")
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            print(f"Файл {date_str}.xls не найден (Статус: {r.status_code})")
            return None
        
        print(f"Файл скачан, размер: {len(r.content)} байт")
        
        # 1. Пытаемся прочитать как Excel (бинарный формат)
        try:
            df = pd.read_excel(BytesIO(r.content), header=None)
            print("Excel успешно прочитан.")
            return df
        except Exception as e:
            print(f"Не удалось прочитать как Excel: {e}")
        
        # 2. Если не вышло - пробуем как HTML (перебираем кодировки)
        for enc in ['utf-8', 'cp1251', 'latin1', 'windows-1251']:
            try:
                html_content = r.content.decode(enc, errors='ignore')
                dfs = pd.read_html(html_content, header=None)
                if dfs:
                    print(f"HTML успешно прочитан (кодировка {enc}).")
                    return dfs[0]
            except Exception:
                continue
                
        print("Не удалось прочитать файл ни как Excel, ни как HTML.")
        return None
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        return None

def parse_dataframe(df, date_key):
    try:
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
        block_starts = [r for r in range(len(df)) if str(df.iloc[r, 0]).strip() == "№ пары"]
        
        for i, start in enumerate(block_starts):
            end = block_starts[i+1] if i+1 < len(block_starts) else len(df)
            
            cols = []
            for c in range(1, len(df.columns)):
                val = df.iloc[start, c]
                if not pd.isna(val):
                    gname = str(val).strip()
                    if gname and gname != "nan":
                        cols.append((c, gname))
            
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
    for i in range(5):
        target = now + timedelta(days=i)
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

    if not all_days:
        print("\n!!! Нет данных для сохранения.")
        return

    # --- ГЕНЕРАЦИЯ НОВОСТЕЙ ---
    news = []
    old_data = {}
    if os.path.exists('schedule.json'):
        try:
            with open('schedule.json', 'r', encoding='utf-8') as f:
                old_data = json.load(f)
        except: pass

    old_days = set(old_data.get('days', {}).keys())
    new_days = set(all_days.keys())
    
    added_days = new_days - old_days
    for day_key in sorted(added_days):
        day_info = all_days[day_key]
        date_obj = datetime.strptime(day_key, '%Y-%m-%d')
        date_str = date_obj.strftime('%d.%m')
        news.append({
            "date": datetime.now().strftime('%Y-%m-%d'),
            "text": f"Появилось расписание на {date_str} ({day_info['weekday']})."
        })
    
    if not old_data:
        news.append({"date": datetime.now().strftime('%Y-%m-%d'), "text": "Добро пожаловать! Расписание успешно загружено."})

    old_news = old_data.get('news', [])
    all_news = news + old_news
    all_news = all_news[:10] # Оставляем только 10 последних новостей
    
    # --- ФОРМИРУЕМ ИТОГОВЫЙ JSON ---
    res = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(all_groups)),
        "days": all_days,
        "news": all_news
    }
    
    with open('schedule.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(f"\nГотово! Сохранено дней: {len(all_days)}. Новостей: {len(all_news)}")

if __name__ == "__main__":
    main()
