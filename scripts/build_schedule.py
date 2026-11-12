
import json
import re
import sys
from io import BytesIO, StringIO
from pathlib import Path
from datetime import datetime, timedelta
from urllib.parse import urljoin

import requests
import pandas as pd
from bs4 import BeautifulSoup


# ==================================================
# НАСТРОЙКИ
# ==================================================

BASE_URL = "https://www.altask.ru"

# Страница календаря расписания колледжа
CALENDAR_URL = (
    "https://www.altask.ru/index.php/2019-10-16-18-25-52"
)

# Файл сайта находится в корне репозитория
OUTPUT_FILE = Path("schedule.json")

# Сколько дней вперёд искать расписание
DAYS_AHEAD = 14

SECTIONS = {
    "DO": "Дневное отделение",
    "PB": "Пожарная безопасность",
    "EAL": "ППКРС Э. Алексеевой",
    "PS": "ППКРС П. Сухова",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/120.0 Safari/537.36"
    )
}

session = requests.Session()
session.headers.update(HEADERS)


# ==================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==================================================

def clean_text(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() == "nan":
        return ""

    return text


def get_weekday(date_key):
    names = [
        "понедельник",
        "вторник",
        "среда",
        "четверг",
        "пятница",
        "суббота",
        "воскресенье",
    ]

    date = datetime.strptime(date_key, "%Y-%m-%d")

    return names[date.weekday()]


def get_week_number(df):
    parts = []

    for row_index in range(min(15, len(df))):
        for value in df.iloc[row_index].values:
            text = clean_text(value)

            if text:
                parts.append(text)

    text = " ".join(parts).lower()

    match = re.search(
        r"(\d+)\s*[-]?\s*недел",
        text
    )

    return int(match.group(1)) if match else 1


# ==================================================
# ПОЛУЧЕНИЕ СТРАНИЦЫ РАСПИСАНИЯ ПО ДАТЕ
# ==================================================

def find_schedule_page(date_key):
    target_date = datetime.strptime(
        date_key,
        "%Y-%m-%d"
    )

    date_text = target_date.strftime("%d.%m.%Y")

    print(f"Ищем расписание на {date_text}")

    try:
        response = session.get(
            CALENDAR_URL,
            timeout=30
        )

        response.raise_for_status()

    except requests.RequestException as error:
        print(f"Ошибка открытия календаря: {error}")
        return None

    soup = BeautifulSoup(
        response.content,
        "html.parser"
    )

    # Ищем ссылку именно на нужную дату.
    for anchor in soup.find_all("a", href=True):

        text = anchor.get_text(" ", strip=True)

        if date_text not in text:
            continue

        if "расписание" not in text.lower():
            continue

        page_url = urljoin(
            response.url,
            anchor["href"]
        )

        print(f"Найдена страница: {page_url}")

        return page_url

    print(
        f"На странице календаря нет опубликованной "
        f"ссылки на {date_text}"
    )

    return None


# ==================================================
# ПОИСК EXCEL-ФАЙЛОВ НА СТРАНИЦЕ ДАТЫ
# ==================================================

def find_excel_links(date_key):

    page_url = find_schedule_page(date_key)

    if not page_url:
        return None

    try:
        response = session.get(
            page_url,
            timeout=30
        )

        response.raise_for_status()

    except requests.RequestException as error:
        print(f"Ошибка открытия страницы расписания: {error}")
        return None

    soup = BeautifulSoup(
        response.content,
        "html.parser"
    )

    result = {
        section: []
        for section in SECTIONS
    }

    for anchor in soup.find_all("a", href=True):

        href = urljoin(
            response.url,
            anchor["href"].strip()
        )

        match = re.search(
            r"/(DO|PB|EAL|PS)/([^/?#]+\.xls)(?:[?#]|$)",
            href,
            re.I
        )

        if not match:
            continue

        section = match.group(1).upper()

        if section in result and href not in result[section]:
            result[section].append(href)

    print("Найденные файлы:")

    for section, urls in result.items():
        print(
            f"  {SECTIONS[section]}: {len(urls)}"
        )

    return result


# ==================================================
# ЗАГРУЗКА И ЧТЕНИЕ EXCEL
# ==================================================

def download_excel(url):

    try:
        response = session.get(
            url,
            timeout=45
        )

        response.raise_for_status()

    except requests.RequestException as error:
        print(f"Ошибка загрузки файла: {error}")
        return None

    content = response.content

    print(
        f"  Получено байт: {len(content)}; "
        f"тип: {response.headers.get('Content-Type', 'не указан')}"
    )

    if not content:
        print("  Файл пустой.")
        return None

    # Проверяем, не вернул ли сайт HTML-страницу ошибки.
    beginning = content[:500].lower()

    if (
        b"<html" in beginning
        or b"<!doctype html" in beginning
    ):
        try:
            html = content.decode(
                "cp1251",
                errors="replace"
            )

            tables = pd.read_html(
                StringIO(html),
                header=None
            )

            if tables:
                print("  Таблица прочитана как HTML.")
                return tables[0]

        except Exception as error:
            print(f"  HTML не удалось прочитать: {error}")

    # Настоящий старый Excel XLS.
    try:
        df = pd.read_excel(
            BytesIO(content),
            header=None,
            engine="xlrd"
        )

        print("  Excel XLS успешно прочитан.")
        return df

    except Exception as error:
        print(f"  Ошибка чтения XLS: {error}")

    # Дополнительная попытка чтения HTML-таблицы.
    for encoding in ("cp1251", "utf-8"):

        try:
            html = content.decode(
                encoding,
                errors="replace"
            )

            tables = pd.read_html(
                StringIO(html),
                header=None
            )

            if tables:
                print(
                    f"  Таблица прочитана через HTML ({encoding})."
                )

                return tables[0]

        except Exception:
            continue

    print("  Не удалось определить формат файла.")
    print("  Первые байты:", repr(content[:100]))

    return None


# ==================================================
# РАЗБОР EXCEL-ТАБЛИЦЫ
# ==================================================

def parse_dataframe(df):

    if df is None or df.empty:
        return {}, 1

    week = get_week_number(df)

    groups = {}

    block_starts = []

    # Ищем строки заголовков.
    for row_index in range(len(df)):

        first_cell = clean_text(
            df.iloc[row_index, 0]
        ).lower()

        normalized = (
            first_cell
            .replace("№", "")
            .replace(".", "")
            .strip()
        )

        if normalized in (
            "пары",
            "пара",
            "номер пары"
        ):
            block_starts.append(row_index)

        elif first_cell in (
            "№ пары",
            "№пары"
        ):
            block_starts.append(row_index)

    if not block_starts:
        print("  Не найдена строка с номерами пар.")
        return {}, week

    for block_index, start in enumerate(block_starts):

        if block_index + 1 < len(block_starts):
            end = block_starts[block_index + 1]
        else:
            end = len(df)

        columns = []

        # Получаем названия групп.
        for col in range(1, len(df.columns)):

            group_name = clean_text(
                df.iloc[start, col]
            )

            if not group_name:
                continue

            if group_name.lower() in (
                "итого",
                "примечание"
            ):
                continue

            columns.append(
                (col, group_name)
            )

        if not columns:
            continue

        row_index = start + 1

        while row_index < end:

            pair_text = clean_text(
                df.iloc[row_index, 0]
            )

            if (
                pair_text.isdigit()
                and 1 <= int(pair_text) <= 8
            ):

                pair_number = int(pair_text)

                for col, group_name in columns:

                    subject = (
                        clean_text(df.iloc[row_index, col])
                        if col < len(df.columns)
                        else ""
                    )

                    teacher_info = ""

                    if (
                        row_index + 1 < end
                        and col < len(df.columns)
                    ):
                        teacher_info = clean_text(
                            df.iloc[row_index + 1, col]
                        )

                    if not subject and not teacher_info:
                        continue

                    teacher = teacher_info
                    room = ""

                    subgroup = None

                    if (
                        "(п/п)" in teacher.lower()
                        or "подгруппа" in teacher.lower()
                    ):
                        subgroup = 1

                    room_match = re.search(
                        r"(?:ауд\.?\s*)?(\d+[а-яА-Я]?)\s*(?:\(|$)",
                        teacher,
                        re.I
                    )

                    if room_match:
                        room = room_match.group(1)

                        teacher = (
                            teacher[:room_match.start()]
                            + teacher[room_match.end():]
                        ).strip()

                    teacher = re.sub(
                        r"\s*\((?:п/п|подгруппа)\)\s*",
                        " ",
                        teacher,
                        flags=re.I
                    ).strip()

                    lesson = {
                        "n": pair_number,
                        "s": subject,
                        "t": teacher,
                        "r": room
                    }

                    if subgroup:
                        lesson["p"] = subgroup

                    groups.setdefault(
                        group_name,
                        []
                    )

                    if lesson not in groups[group_name]:
                        groups[group_name].append(lesson)

                row_index += 2

            else:
                row_index += 1

    # Убираем группы без занятий.
    groups = {
        name: lessons
        for name, lessons in groups.items()
        if lessons
    }

    for name in groups:
        groups[name].sort(
            key=lambda item: item["n"]
        )

    return groups, week


# ==================================================
# СБОР ВСЕХ ОТДЕЛЕНИЙ ЗА ДЕНЬ
# ==================================================

def build_one_day(date_key):

    links = find_excel_links(date_key)

    if links is None:
        return None

    missing_sections = [
        section
        for section in SECTIONS
        if not links[section]
    ]

    if missing_sections:
        print(
            "Не найдены файлы отделений: "
            + ", ".join(missing_sections)
        )

        return None

    combined_groups = {}
    weeks = []

    for section in SECTIONS:

        for url in links[section]:

            print()
            print(f"Загружаем {section}: {url}")

            df = download_excel(url)

            if df is None:
                print("Не удалось прочитать файл отделения.")
                return None

            groups, week = parse_dataframe(df)

            if not groups:
                print(
                    f"В файле {section} не найдены группы."
                )

                return None

            weeks.append(week)

            for group_name, lessons in groups.items():

                combined_groups.setdefault(
                    group_name,
                    []
                )

                for lesson in lessons:

                    if lesson not in combined_groups[group_name]:
                        combined_groups[group_name].append(
                            lesson
                        )

    if not combined_groups:
        return None

    for group_name in combined_groups:

        combined_groups[group_name].sort(
            key=lambda item: item["n"]
        )

    week_number = (
        max(set(weeks), key=weeks.count)
        if weeks
        else 1
    )

    print()
    print(
        f"Успешно обработано групп: "
        f"{len(combined_groups)}"
    )

    return {
        "weekday": get_weekday(date_key),
        "week": week_number,
        "groups": combined_groups
    }


# ==================================================
# ГЛАВНАЯ ПРОГРАММА
# ==================================================

def main():

    print("=" * 55)
    print("AASK SCHEDULE — ОБНОВЛЕНИЕ РАСПИСАНИЯ")
    print("=" * 55)

    today = datetime.now().date()

    all_days = {}
    all_groups = set()

    for offset in range(DAYS_AHEAD):

        target = today + timedelta(days=offset)

        # Воскресенье пропускаем.
        if target.weekday() == 6:
            continue

        date_key = target.strftime("%Y-%m-%d")

        print()
        print("=" * 55)
        print(f"ДАТА: {date_key}")
        print("=" * 55)

        day_data = build_one_day(date_key)

        if day_data is None:
            print("День пропущен.")
            continue

        all_days[date_key] = day_data

        all_groups.update(
            day_data["groups"].keys()
        )

    print()
    print("=" * 55)
    print("РЕЗУЛЬТАТ")
    print("=" * 55)

    print(f"Дней найдено: {len(all_days)}")
    print(f"Групп найдено: {len(all_groups)}")

    if not all_days:
        print("ОШИБКА: не найдено ни одного полного дня.")
        print("Старый schedule.json не изменён.")
        sys.exit(1)

    if len(all_groups) < 2:
        print("ОШИБКА: слишком мало групп.")
        print("Старый schedule.json не изменён.")
        sys.exit(1)

    result = {
        "updated": datetime.now().astimezone().isoformat(
            timespec="minutes"
        ),
        "groups": sorted(
            all_groups,
            key=lambda value: value.lower()
        ),
        "days": all_days
    }

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    temp_file = OUTPUT_FILE.with_suffix(".tmp")

    with open(
        temp_file,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            result,
            file,
            ensure_ascii=False,
            indent=2
        )

    # Заменяем старый файл только после успешной сборки.
    temp_file.replace(OUTPUT_FILE)

    print()
    print("РАСПИСАНИЕ УСПЕШНО СОХРАНЕНО")
    print(f"Файл: {OUTPUT_FILE}")
    print(f"Дней: {len(all_days)}")
    print(f"Групп: {len(all_groups)}")


if __name__ == "__main__":
    main()
