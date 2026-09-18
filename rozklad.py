#!/usr/bin/env python3
"""
Читає опублікований .ics календар групи.

З календаря беремо тільки зустрічі: час, назву і локацію.
Викладач, матеріали й решта довідкової інформації живуть у Notion —
календар на них не претендує.

    python3 rozklad.py --tyzhden 2026-09-14
        друкує блоки [[day]] для tyzhden.toml

    python3 rozklad.py --nazvy
        показує, які назви курсів трапляються в календарі —
        щоб вписати їх у поле «Аліаси в календарі» в базі Курси

Позначки формату беруться з квадратних дужок на початку назви:
[Л] — лекція, [П] — практичне, [С] — семінар.
"""

import argparse
import re
import sys
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from dateutil.rrule import rrulestr

CALENDAR_URL = (
    "https://outlook.office365.com/owa/calendar/"
    "d3f74e7c7c874a48a051108e642f5c08@ukma.edu.ua/"
    "eab3cfddc2954aa984fefb00890a997512287280916006072470/calendar.ics"
)

KYIV = ZoneInfo("Europe/Kyiv")
# Exchange пише назви таймзон по-віндовому; нам потрібна лише ця.
WINDOWS_TZ = {"FLE Standard Time": KYIV, "Ukraine Standard Time": KYIV}

KIND = {"Л": "Лекція", "П": "Практичне", "С": "Семінар"}
WEEKDAYS_SHORT = ["пн", "вт", "ср", "чт", "пт", "сб", "нд"]


# ── розбір .ics ──────────────────────────────────────────────

def unfold(text):
    """
    Нормалізує переноси і склеює згорнуті рядки.

    Exchange віддає .ics із CRLF, а довгі значення розбиває на кілька
    рядків, де продовження починається з пробілу або табуляції.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n[ \t]", "", text)


def unescape(value):
    return (value.replace("\\n", "\n").replace("\\,", ",")
                 .replace("\\;", ";").replace("\\\\", "\\"))


def parse_props(block):
    props = []
    for line in block.split("\n"):
        if ":" not in line:
            continue
        head, _, value = line.partition(":")
        parts = head.split(";")
        params = {}
        for p in parts[1:]:
            k, _, v = p.partition("=")
            params[k.upper()] = v
        props.append((parts[0].upper(), params, value.strip()))
    return props


def parse_dt(value, params):
    """DTSTART може бути датою, локальним часом із TZID або UTC із Z."""
    if params.get("VALUE") == "DATE":
        return datetime.strptime(value, "%Y%m%d").replace(tzinfo=KYIV), True
    if value.endswith("Z"):
        dt = datetime.strptime(value, "%Y%m%dT%H%M%SZ")
        return dt.replace(tzinfo=ZoneInfo("UTC")).astimezone(KYIV), False
    tz = WINDOWS_TZ.get(params.get("TZID", ""), KYIV)
    return datetime.strptime(value, "%Y%m%dT%H%M%S").replace(tzinfo=tz), False


def load_events(source):
    events = []
    for block in re.findall(r"BEGIN:VEVENT\n(.*?)END:VEVENT",
                            unfold(source), re.S):
        ev = {"exdates": [], "location": ""}
        for name, params, value in parse_props(block):
            if name == "SUMMARY":
                ev["summary"] = unescape(value)
            elif name == "LOCATION":
                ev["location"] = unescape(value)
            elif name == "DTSTART":
                ev["start"], ev["all_day"] = parse_dt(value, params)
            elif name == "DTEND":
                ev["end"], _ = parse_dt(value, params)
            elif name == "RRULE":
                ev["rrule"] = value
            elif name == "EXDATE":
                for piece in value.split(","):
                    ev["exdates"].append(parse_dt(piece, params)[0])
        if ev.get("summary") and ev.get("start"):
            events.append(ev)
    return events


def occurrences(ev, window_start, window_end):
    """Розгортає повторювані події у межах вікна, з урахуванням скасувань."""
    if not ev.get("rrule"):
        return [ev["start"]] if window_start <= ev["start"] <= window_end else []
    rule = rrulestr(ev["rrule"], dtstart=ev["start"])
    skip = {d.astimezone(KYIV) for d in ev["exdates"]}
    return [dt.astimezone(KYIV)
            for dt in rule.between(window_start, window_end, inc=True)
            if dt.astimezone(KYIV) not in skip]


# ── назва і формат заняття ───────────────────────────────────

def split_summary(summary):
    """«[Л] Проектування програмних систем» → ('Проектування…', 'Лекція')."""
    m = re.match(r"^\[(.)\]\s*(.+)$", summary.strip())
    if not m:
        return summary.strip(), ""
    name = m.group(2).strip()
    # Після назви інколи дописують примітку через крапку — вона не частина назви.
    name = re.split(r"\.\s+(?=[А-ЯЇІЄҐA-Z])", name)[0].strip().rstrip(".")
    return name, KIND.get(m.group(1).upper(), "")


def sessions_in_window(events, window_start, window_end):
    out = []
    for ev in events:
        if ev.get("all_day"):
            continue
        name, kind = split_summary(ev["summary"])
        if not kind:          # оголошення без позначки формату — не заняття
            continue
        link = ev["location"].strip()
        span = (ev["end"] - ev["start"]) if ev.get("end") else None
        for when in occurrences(ev, window_start, window_end):
            out.append({"when": when, "title": name, "kind": kind,
                        "until": when + span if span else None,
                        "link": link if link.startswith("http") else ""})
    return out


# ── режим 1: розклад тижня у формат tyzhden.toml ─────────────

def print_week(events, monday):
    window_start = datetime.combine(monday, datetime.min.time(), KYIV)
    window_end = window_start + timedelta(days=7)

    by_day = defaultdict(list)
    for s in sessions_in_window(events, window_start, window_end):
        by_day[s["when"].date()].append(s)

    if not by_day:
        print("# Цього тижня занять у календарі немає", file=sys.stderr)
        return

    print(f"# Згенеровано з календаря на тиждень від {monday.isoformat()}.")
    print("# Викладач і запасне посилання підставляються на кроці злиття з Notion.")
    for day in sorted(by_day):
        print(f'\n[[day]]\ndate = "{day.isoformat()}"')
        for s in sorted(by_day[day], key=lambda x: x["when"]):
            print("[[day.session]]")
            print(f'time    = "{s["when"].strftime("%H:%M")}"')
            print(f'title   = "{s["title"]}"')
            print(f'teacher = "{s["kind"]}"')
            if s["link"]:
                print(f'link    = "{s["link"]}"')
            else:
                print("# link — у події порожня локація, підставиться з Notion")


# ── режим 2: назви курсів для поля «Аліаси в календарі» ──────

def print_names(events):
    found = defaultdict(lambda: {"kinds": set(), "slots": set()})
    for ev in events:
        name, kind = split_summary(ev["summary"])
        if not kind:
            continue
        found[name]["kinds"].add(kind)
        found[name]["slots"].add(f'{WEEKDAYS_SHORT[ev["start"].weekday()]} '
                                 f'{ev["start"].strftime("%H:%M")}')

    print(f"Назв у календарі: {len(found)}\n")
    for name, info in sorted(found.items()):
        print(f"  {name}")
        print(f'      формати: {", ".join(sorted(info["kinds"]))}')
        print(f'      слоти:   {", ".join(sorted(info["slots"]))}\n')
    print("Ці назви мають дослівно потрапити в поле «Аліаси в календарі»")
    print("відповідних записів бази Курси — по них іде звʼязування.")

    unmarked = {ev["summary"].strip() for ev in events
                if not split_summary(ev["summary"])[1] and not ev.get("all_day")}
    if unmarked:
        print("\nПодії без позначки [Л]/[П]/[С] — у розклад не потраплять:")
        for s in sorted(unmarked):
            print(f"  · {s}")


# ── точка входу ──────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tyzhden", metavar="РРРР-ММ-ДД",
                    help="понеділок тижня, для якого друкувати розклад")
    ap.add_argument("--nazvy", action="store_true",
                    help="показати назви курсів, що трапляються в календарі")
    ap.add_argument("--file", help="локальний .ics замість завантаження")
    args = ap.parse_args()

    if not args.tyzhden and not args.nazvy:
        ap.error("вкажи --tyzhden або --nazvy")

    if args.file:
        source = open(args.file, encoding="utf-8").read()
    else:
        with urllib.request.urlopen(CALENDAR_URL, timeout=30) as r:
            source = r.read().decode("utf-8")

    events = load_events(source)

    if args.nazvy:
        print_names(events)
    if args.tyzhden:
        monday = date.fromisoformat(args.tyzhden)
        monday -= timedelta(days=monday.weekday())
        print_week(events, monday)


if __name__ == "__main__":
    main()
