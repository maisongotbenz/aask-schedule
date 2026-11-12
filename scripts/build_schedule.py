import os
import json
import requests
import xlrd
from datetime import datetime

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
            print(f"Файл не найден. Статус: {response.status_code}")
            return None
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        return None

def parse_excel(file_path):
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        data = []
        for row_idx in range(sheet.nrows):
            row_values = sheet.row_values(row_idx)
            row_values = [str(cell).strip() for cell in row_values if str(cell).strip()]
            if row_values:
                data.append(row_values)
        return data
    except Exception as e:
        print(f"Ошибка при парсинге Excel: {e}")
        return None

def main():
    today = datetime.now()
    date_str = today.strftime("%d%m")
    print(f"Ищем расписание на дату: {date_str}")

    file_path = download_file(date_str)

    if file_path:
        print("Файл скачан, начинаем парсинг...")
        parsed_data = parse_excel(file_path)

        if parsed_data:
            output_path = 'schedule.json'

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(parsed_data, f, ensure_ascii=False, indent=2)
            print(f"Успешно сохранено в {output_path}")
        else:
            print("Не удалось распарсить данные из файла.")
    else:
        print("Не удалось скачать файл расписания. Возможно, на сайте его еще нет.")

if __name__ == "__main__":
    main()
