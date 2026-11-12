import os
import json
import re
import requests
import pandas as pd
from datetime import datetime
from io import StringIO

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
        
        # 1. Пытаемся прочитать как Excel
        try:
            df = pd.read_excel(StringIO(r.content.decode('cp1251')), header=None)
            print("Excel успешно прочитан.")
            return df
        except Exception as e:
            print(f"Не удалось прочитать как Excel: {e}")
        
        # 2. Если не вышло - пробуем как HTML (часто сайты отдают HTML под видом .xls)
        try:
            # Пробуем разные кодировки
            for enc in ['cp1251', 'utf-8', 'latin1']:
                try:
                    html_content = r.content.decode(enc, errors='ignore')
                    dfs = pd.read_html(StringIO(html_content), header=None)
                    if dfs:
                        print(f"HTML успешно прочитан (кодировка {enc}).")
                        return dfs[0]
                except Exception:
                    continue
        except Exception as e:
            print(f"Не удалось прочитать как HTML: {e}")
            
        return None
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        return None

# ... остальной код парсинга (parse_dataframe и main) оставьте без изменений ...
