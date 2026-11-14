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
        
        try:
            df = pd.read_excel(BytesIO(r.content), header=None)
            print("Excel успешно прочитан.")
            return df
        except Exception as e:
            print(f"Не удалось прочитать как Excel: {e}")
        
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
        
        # Проверяем, понедельник ли это
        is_monday = (weekday == 'понедельник')
        
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
                        
                        # --- ЛОГИКА ДЛЯ "РАЗГОВОРОВ О ВАЖНОМ" ---
                        # Если день - понедельник, и предмет пустой, но есть текст в ячейке преподавателя
                        if is_monday and not subj and t_info:
                            # Принудительно ставим название
                            subj = "Разговоры о важном"
                            # t_info оставляем как есть (там фамилия преподавателя)
                        # -----------------------------------------

                        if subj or t_info:
                            t = t_info
                            rm = ""
                            p = None
                            if "(п/п)" in t or "подгруппа" in t.lower(): p = 1
                            
                            # Ищем аудиторию (число в конце)
                            room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', t)
                            if room_match:
                                rm = room_match.group(1)
                                t = t.replace(room_match.group(0), '').strip()
                            
                            # Очищаем от подгрупп
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
    old_data = {}
    if os.path.exists('schedule.json'):
        try:
            with open('schedule.json', 'r', encoding='utf-8') as f:
                old_data = json.load(f)
        except: pass

    all_days = old_data.get('days', {}) 
    all_groups = set(old_data.get('groups', []))
    news = old_data.get('news', [])
    
    now = datetime.now()
    new_data_found = False

    for i in range(7):
        target = now + timedelta(days=i)
        date_str = target.strftime("%d%m")
        date_key = target.strftime("%Y-%m-%d")
        
        # ВАЖНО: Если мы хотим обновить существующие дни (из-за новой логики с Разговорами),
        # нужно удалить эту проверку. Но чтобы не перекачивать всё каждый раз, 
        # мы проверим, есть ли уже "Разговоры о важном" в этом дне.
        # Для простоты, давайте просто перекачивать всё заново при каждом запуске (это быстро).
        # if date_key in all_days:
        #     continue

        print(f"\n--- Пробуем {date_str} ({date_key}) ---")
        df = download_and_parse(date_str)
        
        if df is not None:
            day_data = parse_dataframe(df, date_key)
            if day_data and day_data['groups']:
                all_days[date_key] = day_data
                all_groups.update(day_data['groups'].keys())
                new_data_found = True
                print(f"  -> Успешно добавлен новый день {date_key} (групп: {len(day_data['groups'])})")
                
                date_obj = datetime.strptime(date_key, '%Y-%m-%d')
                news.insert(0, {
                    "date": datetime.now().strftime('%Y-%m-%d'),
                    "text": f"Появилось расписание на {date_obj.strftime('%d.%m')} ({day_data['weekday']})."
                })

    if not new_data_found:
        print("\n!!! Новых данных нет. Расписание не изменилось.")
        if not news:
            news.append({"date": datetime.now().strftime('%Y-%m-%d'), "text": "Добро пожаловать! Расписание загружено."})
    else:
        print(f"\nНайдены новые дни! Обновляем schedule.json.")

    news = news[:10]
    
    res = {
        "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
        "groups": sorted(list(all_groups)),
        "days": all_days,
        "news": news
    }
    
    with open('schedule.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(f"Готово! В файле дней: {len(all_days)}")

if __name__ == "__main__":
    main()
