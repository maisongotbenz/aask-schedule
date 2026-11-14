#!/usr/bin/env python3
"""Скачивает расписание ААСК («Дневное отделение»), разбирает ВСЕ группы и пишет public/schedule.json.
Запуск: python scripts/build_schedule.py            (скачать файлы с сайта колледжа)
        python scripts/build_schedule.py --local 1211.xlsx   (разобрать свой файл)"""
import argparse, json, re, sys, tempfile, urllib.request, urllib.error
from datetime import datetime, timedelta, date
from pathlib import Path
from zoneinfo import ZoneInfo

URL = "https://altask.ru/images/raspisanie/DO/{:02d}{:02d}.xls"
OUT = Path("public/schedule.json")
KEEP_DAYS = 14
MONTHS = {"января":1,"февраля":2,"марта":3,"апреля":4,"мая":5,"июня":6,"июля":7,"августа":8,
          "сентября":9,"октября":10,"ноября":11,"декабря":12}

def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v not in (None, "") else ""

def load_rows(path):
    if path.suffix.lower() == ".xls":
        import xlrd
        ws = xlrd.open_workbook(str(path)).sheet_by_index(0)
        return [[ws.cell_value(r, c) or None for c in range(ws.ncols)] for r in range(ws.nrows)]
    from openpyxl import load_workbook
    return [list(r) for r in load_workbook(path, data_only=True).worksheets[0].iter_rows(values_only=True)]

def cell(rows, r, c):
    return rows[r][c] if r < len(rows) and c < len(rows[r]) else None

def parse_title(rows):
    for row in rows[:3]:
        for v in row:
            m = re.search(r"на\s+(\d{1,2})\s+([а-я]+)\s+([а-я]+),?\s*(\d+)\s*недел", clean(v), re.I)
            if m and m.group(2).lower() in MONTHS:
                return int(m.group(1)), MONTHS[m.group(2).lower()], m.group(3).lower(), int(m.group(4))
    raise ValueError("не найден заголовок «Расписание на 12 ноября четверг, 1 неделя»")

def teacher_room(raw):
    sub = bool(re.search(r"\((п/п|подгруппа)\)", raw, re.I))
    raw = re.sub(r"\((п/п|подгруппа)\)", "", raw, flags=re.I).strip()
    m = re.match(r"^(.+\.)\s*(\d+[а-яa-z]?)?\s*$", raw, re.I)
    return (m.group(1).strip(), m.group(2) or "", sub) if m else (raw, "", sub)

TEACHER = re.compile(r"^[А-ЯЁ][а-яё\-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]?\.?")
TALK = "Разговоры о важном"

def parse_file(path):
    rows = load_rows(path)
    d, m, weekday, week = parse_title(rows)
    heads = [r for r, row in enumerate(rows) if clean(row[0] if row else "").lower().startswith("№ пары")]
    groups = {}
    for i, h in enumerate(heads):
        end = heads[i + 1] if i + 1 < len(heads) else len(rows)
        for c in range(1, len(rows[h])):
            g = clean(rows[h][c])
            if not g:
                continue
            lessons, r = [], h + 1
            if weekday == "понедельник":   # строка с фамилией между названием группы и 1-й парой = «Разговоры о важном»
                rr = h + 1
                while rr < end and not isinstance(cell(rows, rr, 0), (int, float)):
                    v = clean(cell(rows, rr, c))
                    if TEACHER.match(v):
                        t, room, _ = teacher_room(v)
                        lessons.append({"n": 0, "s": TALK, "t": t, "r": room})
                        break
                    rr += 1
                if not lessons:   # в файле фамилии нет, но разговоры проходят у всех групп
                    lessons.append({"n": 0, "s": TALK, "t": "", "r": ""})
            while r < end:
                first = cell(rows, r, 0)
                if isinstance(first, (int, float)) and not isinstance(first, bool):
                    subj = clean(cell(rows, r, c))
                    if subj:
                        t, room, sub = teacher_room(clean(cell(rows, r + 1, c)))
                        les = {"n": int(first), "s": subj, "t": t, "r": room}
                        if sub:
                            les["p"] = 1
                        lessons.append(les)
                    r += 2
                else:
                    r += 1
            lessons.sort(key=lambda x: x["n"])
            groups[g] = lessons
    return d, m, weekday, week, groups

def resolve_year(d, m, today):
    return min((date(y, m, d) for y in (today.year - 1, today.year, today.year + 1)),
               key=lambda x: abs((x - today).days))

def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (aask-schedule-bot)"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read(), None
    except urllib.error.HTTPError as e:
        return None, (None if e.code == 404 else f"HTTP {e.code}")
    except Exception as e:
        return None, str(e)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--local", nargs="*")
    a = ap.parse_args()
    today = datetime.now(ZoneInfo("Asia/Barnaul")).date()

    files, errors = [], []
    if a.local:
        files = [Path(p) for p in a.local]
    else:
        tmp = Path(tempfile.mkdtemp())
        for off in range(-1, 8):
            d = today + timedelta(days=off)
            url = URL.format(d.day, d.month)
            data, err = fetch(url)
            if err:
                errors.append(f"{url}: {err}")
            elif data:
                ext = ".xls" if data[:8] == bytes.fromhex("D0CF11E0A1B11AE1") else ".xlsx" if data[:2] == b"PK" else None
                if not ext:
                    errors.append(f"{url}: это не Excel-файл")
                    continue
                f = tmp / f"{d.day:02d}{d.month:02d}{ext}"
                f.write_bytes(data)
                files.append(f)
                print("скачан", url)

    old = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {"days": {}}
    days = dict(old.get("days", {}))
    for f in files:
        try:
            d, m, weekday, week, groups = parse_file(f)
        except Exception as e:
            errors.append(f"{f.name}: не удалось разобрать ({e})")
            continue
        iso = resolve_year(d, m, today).isoformat()
        days[iso] = {"weekday": weekday, "week": week, "groups": groups}
        print(f"{iso}: групп {len(groups)}, пар {sum(len(v) for v in groups.values())}")

    cutoff = (today - timedelta(days=KEEP_DAYS)).isoformat()
    days = {k: v for k, v in sorted(days.items()) if k >= cutoff}
    for e in errors:
        print("ВНИМАНИЕ:", e, file=sys.stderr)
    if not files and errors:
        sys.exit("Не удалось получить ни одного файла")
    if days == old.get("days"):
        print("Изменений нет")
        return
    names = sorted({g for d in days.values() for g in d["groups"]})
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({"updated": datetime.now(ZoneInfo("Asia/Barnaul")).isoformat(timespec="minutes"),
                               "groups": names, "days": days}, ensure_ascii=False, separators=(",", ":")),
                   encoding="utf-8")
    print("Записано:", OUT, f"({len(names)} групп, {len(days)} дн.)")

if __name__ == "__main__":
    main()
