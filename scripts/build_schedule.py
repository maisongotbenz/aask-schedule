
import json
import re
import sys
from io import BytesIO, StringIO
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin

import requests
import pandas as pd
from bs4 import BeautifulSoup


# --------------------------------------------------
# НАСТРОЙКИ
# --------------------------------------------------

BASE_URL = "https://altask.ru"

# Страница расписания колледжа с параметром даты.
PAGE_URL_TEMPLATE = (
    "https://altask.ru/index.php/staff/"
    "1563-11-11-2026?date={date}-00-00"
)

OUTPUT_FILE = Path("public/schedule.json")

DAYS_AHEAD = 14

# Все четыре раздела официального расписания.
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


# --------------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# --------------------------------------------------

def cell_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def weekday_name(date_key):
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


def get_week_number(text):
    match = re.search(r"(\d+)\s*[-]?\s*недел", text.lower())
    return int(match.group(1)) if match else 1


# --------------------------------------------------
# ПОИСК ССЫЛОК НА ОФИЦИАЛЬНОЙ СТРАНИЦЕ
# --------------------------------------------------

def find_excel_links(date_key):
    page_url = PAGE_URL_TEMPLATE.format(date=date_key)

    print(f"Открываем страницу: {date_key}")

    try:
        response = session.get(page_url, timeout=30)
        response.raise_for_status()
    except requests.RequestException as error:
        print(f"Ошибка открытия страницы: {error}")
        return None

    soup = BeautifulSoup(response.content, "html.parser")

    result = {}

    for section in SECTIONS:
        result[section] = []

    # Ищем настоящие ссылки на Excel.
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()

        if not re.search(r"\.xls(?:\?|$)", href, re.I):
            continue

        full_url = urljoin(page_url, href)

        match = re.search(
            r"/(DO|PB|EAL|PS)/([^/?]+\.xls)",
            full_url,
            re.I,
        )

        if not match:
            continue

        section = match.group(1).upper()

        if section in result:
            if full_url not in result[section]:
                result[section].append(full_url)

    for section, links in result.items():
        print(
            f"  {SECTIONS[section]}: "
            f"{len(links)} файлов"
        )

    return result


# --------------------------------------------------
# СКАЧИВАНИЕ EXCEL
# --------------------------------------------------

def download_excel(url):
    try:
        response = session.get(url, timeout=45)
        response.raise_for_status()

        if not response.content:
            raise ValueError("Пустой файл")

        content = response.content

        # Сначала пробуем настоящий Excel XLS.
        try:
            return pd.read_excel(
                BytesIO(content),
                header=None,
                engine="xlrd",
            )
        except Exception:
            pass

        # Некоторые старые XLS на сайте фактически являются HTML.
        try:
            text = content.decode(
                "cp1251",
                errors="replace",
            )

            tables = pd.read_html(
                StringIO(text),
                header=None,
            )

            if tables:
                return tables[0]

        except Exception:
            pass

        # Дополнительная попытка UTF-8.
        try:
            text = content.decode(
                "utf-8",
                errors="replace",
            )

            tables = pd.read_html(
                StringIO(text),
                header=None,
            )

            if tables:
                return tables[0]

        except Exception:
            pass

        print(f"Не удалось прочитать таблицу: {url}")
        return None

    except Exception as error:
        print(f"Ошибка скачивания {url}: {error}")
        return None


# --------------------------------------------------
# ОБРАБОТКА ОДНОЙ EXCEL-ТАБЛИЦЫ
# --------------------------------------------------

def parse_dataframe(df):
    if df is None or df.empty:
        return {}, 1

    # Определяем номер недели.
    title_parts = []

    for row_index in range(min(15, len(df))):
        row = [
            cell_text(value)
            for value in df.iloc[row_index].values
        ]
        title_parts.extend(row)

    title = " ".join(title_parts)
    week = get_week_number(title)

    groups = {}

    # Находим все блоки с заголовками групп.
    block_starts = []

    for row_index in range(len(df)):
        first = cell_text(df.iloc[row_index, 0]).lower()
        first = first.replace("№", "").replace(".", "").strip()

        if first == "пары" or first == "пары ":
            block_starts.append(row_index)

        elif first in ("№ пары", "номер пары"):
            block_starts.append(row_index)

    if not block_starts:
        print("  Не найден заголовок «№ пары»")
        return {}, week

    for block_index, start in enumerate(block_starts):

        end = (
            block_starts[block_index + 1]
            if block_index + 1 < len(block_starts)
            else len(df)
        )

        columns = []

        # Получаем названия групп из заголовка блока.
        for col in range(1, len(df.columns)):
            group_name = cell_text(df.iloc[start, col])

            if not group_name:
                continue

            if group_name.lower() in (
                "nan",
                "итого",
                "примечание",
            ):
                continue

            columns.append((col, group_name))

        if not columns:
            continue

        row_index = start + 1

        while row_index < end:

            pair_text = cell_text(df.iloc[row_index, 0])

            if pair_text.isdigit() and 1 <= int(pair_text) <= 8:

                pair_number = int(pair_text)

                for col, group_name in columns:

                    subject = (
                        cell_text(df.iloc[row_index, col])
                        if col < len(df.columns)
                        else ""
                    )

                    teacher_info = (
                        cell_text(df.iloc[row_index + 1, col])
                        if row_index + 1 < end
                        and col < len(df.columns)
                        else ""
                    )

                    if not subject and not teacher_info:
                        continue

                    teacher = teacher_info
                    room = ""

                    subgroup = None

                    if (
                        "(п/п)" in teacher
                        or "подгруппа" in teacher.lower()
                    ):
                        subgroup = 1

                    # Ищем номер аудитории.
                    room_match = re.search(
                        r"(?:ауд\.?\s*)?(\d+[а-яА-Я]?)\s*(?:\(|$)",
                        teacher,
                        re.I,
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
                        flags=re.I,
                    ).strip()

                    item = {
                        "n": pair_number,
                        "s": subject,
                        "t": teacher,
                        "r": room,
                    }

                    if subgroup:
                        item["p"] = subgroup

                    groups.setdefault(group_name, [])

                    # Не теряем разные подгруппы и преподавателей.
                    if item not in groups[group_name]:
                        groups[group_name].append(item)

                row_index += 2

            else:
                row_index += 1

    return groups, week


# --------------------------------------------------
# СБОР РАСПИСАНИЯ НА ОДИН ДЕНЬ
# --------------------------------------------------

def build_one_day(date_key):
    links = find_excel_links(date_key)

    if links is None:
        return None

    available_sections = [
        section
        for section, urls in links.items()
        if urls
    ]

    if not available_sections:
        print("  На странице нет ссылок на Excel.")
        return None

    # Не публикуем день, если отсутствует хотя бы один раздел.
    missing = [
        section
        for section in SECTIONS
        if not links[section]
    ]

    if missing:
        print(
            "  Неполная публикация. Отсутствуют разделы: "
            + ", ".join(missing)
        )
        return None

    combined_groups = {}
    weeks = []

    for section in SECTIONS:

        for url in links[section]:

            print(f"  Скачиваем {section}: {url}")

            df = download_excel(url)

            if df is None:
                print("  Не удалось загрузить один из разделов.")
                return None

            groups, week = parse_dataframe(df)

            if not groups:
                print(
                    f"  В разделе {section} не найдены группы."
                )
                return None

            weeks.append(week)

            for group_name, lessons in groups.items():

                combined_groups.setdefault(group_name, [])

                for lesson in lessons:
                    if lesson not in combined_groups[group_name]:
                        combined_groups[group_name].append(lesson)

    if not combined_groups:
        return None

    # Сортируем занятия по номеру пары.
    for group_name in combined_groups:
        combined_groups[group_name].sort(
            key=lambda lesson: lesson["n"]
        )

    # Используем наиболее часто встречающийся номер недели.
    week_number = (
        max(set(weeks), key=weeks.count)
        if weeks
        else 1
    )

    print(
        f"  Успешно: {len(combined_groups)} групп"
    )

    return {
        "weekday": weekday_name(date_key),
        "week": week_number,
        "groups": combined_groups,
    }


# --------------------------------------------------
# ОСНОВНАЯ ПРОГРАММА
# --------------------------------------------------

def main():

    print("=" * 55)
    print("AASK SCHEDULE — АВТОМАТИЧЕСКОЕ ОБНОВЛЕНИЕ")
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
        print(f"--- {date_key} ---")

        day_data = build_one_day(date_key)

        if day_data is None:
            print("День пропущен: нет полной проверенной публикации.")
            continue

        all_days[date_key] = day_data

        all_groups.update(
            day_data["groups"].keys()
        )

    # Нельзя заменять расписание пустым результатом.
    if not all_days:
        print()
        print("ОШИБКА: не найдено ни одного полного дня.")
        print("Старый schedule.json оставлен без изменений.")
        sys.exit(1)

    # Дополнительная проверка количества групп.
    print()
    print(f"Всего найдено групп: {len(all_groups)}")
    print(f"Всего дней: {len(all_days)}")

    if len(all_groups) < 2:
        print("ОШИБКА: слишком мало групп. Файл не обновлён.")
        sys.exit(1)

    result = {
        "updated": datetime.now().astimezone().isoformat(
            timespec="minutes"
        ),
        "groups": sorted(
            all_groups,
            key=lambda value: value.lower()
        ),
        "days": all_days,
    }

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Сначала записываем во временный файл.
    temp_file = OUTPUT_FILE.with_suffix(".tmp")

    with open(temp_file, "w", encoding="utf-8") as file:
        json.dump(
            result,
            file,
            ensure_ascii=False,
            indent=2,
        )

    # Атомарно заменяем старый файл.
    temp_file.replace(OUTPUT_FILE)

    print()
    print("=" * 55)
    print("ОБНОВЛЕНИЕ УСПЕШНО")
    print(f"Дней: {len(all_days)}")
    print(f"Групп: {len(all_groups)}")
    print(f"Файл: {OUTPUT_FILE}")
    print("=" * 55)


if __name__ == "__main__":
    main()
