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
    print(f"Пытаемся скачать: {url}")
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            file_path = f"/tmp/{filename}"
            with open(file_path, 'wb') as f:
                f.write(response.content)
            return file_path
        else:
            print(f"Файл не найден на сайте. Статус: {response.status_code}")
            return None
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        return None

def parse_teacher_info(cell_text):
    """Разбирает строку вида 'Иванов И.И. 101 (подгруппа)' на составляющие"""
    if not cell_text:
        return "", "", None
    
    # Ищем аудиторию (обычно это число в конце, иногда с буквой)
    room_match = re.search(r'(\d+[а-яА-Я]?)\s*(?:\(|$)', cell_text)
    room = room_match.group(1) if room_match else ""
    
    # Ищем признак подгруппы
    subgroup = 1 if "(п/п)" in cell_text or "подгруппа" in cell_text.lower() else None
    if "(подгруппа)" in cell_text.lower():
        subgroup = 1 # Условно, можно уточнять, но в таблице не указано, какая именно
    
    # Очищаем от аудитории и подгруппы, чтобы получить имя преподавателя
    teacher = re.sub(r'\s*\d+[а-яА-Я]?\s*(?:\(п/п\)|\(подгруппа\))?\s*$', '', cell_text).strip()
    
    return teacher, room, subgroup

def parse_excel(file_path):
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        
        # 1. Извлекаем заголовок (дата, день недели, номер недели)
        title = sheet.cell_value(0, 8) # Ячейка I1
        date_match = re.search(r'на (\d+) (\w+)', title)
        week_match = re.search(r'(\d+)\s*неделя', title)
        
        day_map = {
            'понедельник': 'понедельник', 'вторник': 'вторник', 'среда': 'среда',
            'четверг': 'четверг', 'пятница': 'пятница', 'суббота': 'суббота', 'воскресенье': 'воскресенье'
        }
        
        if not date_match or not week_match:
            print("Не удалось найти дату или неделю в заголовке.")
            return None
            
        day_num = int(date_match.group(1))
        day_name_raw = date_match.group(2).lower()
        weekday = day_map.get(day_name_raw, day_name_raw)
        week_num = int(week_match.group(1))
        
        # Формируем дату в формате YYYY-MM-DD (берем текущий год)
        now = datetime.now()
        try:
            day_date = now.replace(day=day_num)
        except ValueError:
            day_date = now # Если дата некорректна, оставляем текущую
            
        date_key = day_date.strftime("%Y-%m-%d")
        
        # 2. Проходим по всем блокам таблицы
        all_groups = set()
        days_data = {date_key: {"weekday": weekday, "week": week_num, "groups": {}}}
        
        # Ищем строки, где начинается блок с группами (в колонке A написано "№ пары")
        block_starts = []
        for r in range(sheet.nrows):
            if str(sheet.cell_value(r, 0)).strip() == "№ пары":
                block_starts.append(r)
                
        for b_idx, start_row in enumerate(block_starts):
            # Определяем, где заканчивается этот блок
            end_row = block_starts[b_idx + 1] if b_idx + 1 < len(block_starts) else sheet.nrows
            
            # Считываем названия групп из строки start_row (начиная со столбца B)
            groups_in_block = []
            for c in range(1, sheet.ncols):
                group_name = str(sheet.cell_value(start_row, c)).strip()
                if group_name and group_name != "":
                    groups_in_block.append((c, group_name))
                    all_groups.add(group_name)
            
            # Идем по парам в этом блоке (пропускаем строку с группами)
            r = start_row + 1
            while r < end_row:
                # Проверяем, является ли строка началом пары (в колонке A номер от 1 до 7)
                para_num_str = str(sheet.cell_value(r, 0)).strip()
                if para_num_str.isdigit() and 1 <= int(para_num_str) <= 7:
                    para_num = int(para_num_str)
                    
                    # Строка с предметами
                    subjects_row = r
                    # Строка с преподавателями (следующая)
                    teachers_row = r + 1
                    
                    if teachers_row >= end_row:
                        break
                        
                    for col_idx, group_name in groups_in_block:
                        subject = str(sheet.cell_value(subjects_row, col_idx)).strip()
                        teacher_info = str(sheet.cell_value(teachers_row, col_idx)).strip()
                        
                        if subject or teacher_info:
                            teacher, room, subgroup = parse_teacher_info(teacher_info)
                            
                            item = {
                                "n": para_num,
                                "s": subject,
                                "t": teacher,
                                "r": room
                            }
                            if subgroup:
                                item["p"] = subgroup
                                
                            if group_name not in days_data[date_key]["groups"]:
                                days_data[date_key]["groups"][group_name] = []
                                
                            # Проверяем, нет ли уже такой пары (иногда в таблице дублируются строки)
                            if not any(d["n"] == para_num and d["s"] == subject for d in days_data[date_key]["groups"][group_name]):
                                days_data[date_key]["groups"][group_name].append(item)
                    
                    r += 2 # Перепрыгиваем через строку с преподавателями
                else:
                    r += 1 # Если это не номер пары, идем дальше

        # 3. Формируем итоговый JSON
        result = {
            "updated": datetime.now().astimezone().isoformat(timespec='minutes'),
            "groups": sorted(list(all_groups)),
            "days": days_data
        }
        
        return result
    except Exception as e:
        print(f"Ошибка при парсинге Excel: {e}")
        return None

def main():
    today = datetime.now()
    # Сайт обычно выкладывает расписание на завтра или на сегодня, пробуем оба варианта
    for days_offset in [0, 1]:
        target_date = today + timedelta(days=days_offset)
        date_str = target_date.strftime("%d%m")
        
        print(f"Пытаемся найти расписание на {date_str}...")
        file_path = download_file(date_str)
        
        if file_path:
            print("Файл скачан, начинаем парсинг...")
            parsed_data = parse_excel(file_path)
            
            if parsed_data:
                # Сохраняем в корень проекта
                output_path = 'schedule.json' 
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(parsed_data, f, ensure_ascii=False, indent=2)
                print(f"Успешно сохранено в {output_path}")
                return # Выходим, если успешно сохранили
            else:
                print("Не удалось распарсить данные из файла.")
        else:
            print(f"Файла за {date_str} нет, пробуем следующую дату...")

    print("Не удалось найти и обработать расписание.")

if __name__ == "__main__":
    main()
