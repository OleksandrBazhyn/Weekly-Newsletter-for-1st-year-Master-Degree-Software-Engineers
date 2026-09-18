#!/usr/bin/env python3
"""
Збирає тижневий лист і кладе чернетку в Outlook.

    python3 zbirka.py                 наступний тиждень, чернетка в пошті
    python3 zbirka.py --suho          нічого не надсилати, лише зберегти файли
    python3 zbirka.py --tyzhden 2026-09-14   конкретний тиждень

Джерела: календар (час, назва, локація) і Notion (все інше).
Модель пише лише вступний абзац; жоден факт нею не вигадується.
"""

import argparse
import os
import re
import sys
import unicodedata
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import build
import graph_mail
import notion_data
import rozklad

def load_env(path=".env"):
    """Читає .env, не перетираючи вже задані змінні оточення."""
    f = Path(path)
    if not f.exists():
        return
    for line in f.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


load_env()

KYIV = ZoneInfo("Europe/Kyiv")
GORYT = 3          # днів до дедлайну, після яких курс «потребує уваги»
# Номер тижня рахуємо від початку триместру, а не за ISO: «Тиждень 3»
# зрозуміло, «Тиждень 38» — ні.
TRYMESTR = date.fromisoformat(os.environ.get("TRYMESTR_START", "2026-08-31"))
NA_OBRII = 14      # горизонт, у межах якого дедлайн взагалі показуємо


# ── звʼязування назв ─────────────────────────────────────────

def normalize(text):
    text = unicodedata.normalize("NFKD", (text or "").casefold())
    return re.sub(r"[^\w]+", "", text)


def build_index(courses):
    index = {}
    for c in courses.values():
        for alias in [c["name"], *c["aliases"]]:
            index[normalize(alias)] = c
    return index


# ── стан курсу: обчислюється, а не зберігається ──────────────

def course_state(course, tasks, today):
    soonest = min((t["due"] for t in tasks), default=None)
    if soonest is not None:
        left = (soonest - today).days
        if left <= GORYT:
            return "увага"
        if left <= NA_OBRII:
            return "робота"
    return "робота" if course["now"] else "спокійно"


def course_text(course, tasks, today):
    if course["now"]:
        return course["now"]
    if tasks:
        nearest = tasks[0]
        left = (nearest["due"] - today).days
        koly = "сьогодні" if left == 0 else "завтра" if left == 1 else \
               f'до {nearest["due"].day} {build.MONTHS[nearest["due"].month - 1]}'
        return f'Найближче: {nearest["title"]}, {koly}.'
    return ""


# ── вступний абзац ───────────────────────────────────────────

PROMPT = """Ти допомагаєш старості групи магістрів-програмістів НаУКМА
писати вступний абзац до тижневої розсилки.

Ось факти про тиждень:
{facts}

Напиши 2–3 речення українською, які відкриють лист. Вимоги:
- спирайся ТІЛЬКИ на наведені факти, нічого не додавай і не уточнюй;
- жодних привітань, побажань удачі й закликів на кшталт «продуктивного тижня»;
- називай найважливіше першим реченням;
- пиши так, як пише колега колегам: спокійно, без пафосу й канцеляриту;
- якщо тиждень спокійний — так і скажи, не роби з нього подію.
Поверни лише текст абзацу, без лапок і пояснень."""


def _post_json(url, headers, payload):
    import json
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode(), method="POST",
        headers={"content-type": "application/json", **headers})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def _openai(prompt):
    data = _post_json(
        "https://api.openai.com/v1/chat/completions",
        {"authorization": f'Bearer {os.environ["OPENAI_API_KEY"]}'},
        {"model": os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
         "max_completion_tokens": 400,
         "messages": [{"role": "user", "content": prompt}]})
    return data["choices"][0]["message"]["content"].strip()


def _anthropic(prompt):
    data = _post_json(
        "https://api.anthropic.com/v1/messages",
        {"x-api-key": os.environ["ANTHROPIC_API_KEY"],
         "anthropic-version": "2023-06-01"},
        {"model": os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5"),
         "max_tokens": 400,
         "messages": [{"role": "user", "content": prompt}]})
    return "".join(b.get("text", "") for b in data["content"]).strip()


def write_lead(facts):
    """Абзац не критичний: якщо модель мовчить, лист усе одно збереться."""
    if os.environ.get("OPENAI_API_KEY"):
        provider, call = "OpenAI", _openai
    elif os.environ.get("ANTHROPIC_API_KEY"):
        provider, call = "Anthropic", _anthropic
    else:
        print("Ключа до моделі немає — абзац лишаю порожнім", file=sys.stderr)
        return ""
    try:
        return call(PROMPT.format(facts=facts))
    except Exception as err:
        print(f"{provider} не відповів ({err}); абзац лишаю порожнім", file=sys.stderr)
        return ""


def facts_for_lead(cfg, sessions_count):
    lines = [f"Занять цього тижня: {sessions_count}."]
    for d in cfg["deadline"]:
        lines.append(f'Термін: {d["title"]} — {d["date"].isoformat()}.')
    for c in cfg["course"]:
        if c["text"]:
            lines.append(f'{c["name"]} ({c["state"]}): {c["text"]}')
    for n in cfg["news"]:
        lines.append(f'Можливість: {n["title"]}.')
    return "\n".join(lines)


# ── збірка ───────────────────────────────────────────────────

def next_monday(today):
    return today + timedelta(days=(0 - today.weekday()) % 7)


def assemble(monday):
    sunday = monday + timedelta(days=6)

    with urllib.request.urlopen(rozklad.CALENDAR_URL, timeout=30) as r:
        events = rozklad.load_events(r.read().decode("utf-8"))
    window_start = datetime.combine(monday, datetime.min.time(), KYIV)
    sessions = rozklad.sessions_in_window(events, window_start,
                                          window_start + timedelta(days=7))

    courses = notion_data.courses()
    index = build_index(courses)
    tasks = notion_data.open_tasks()

    by_course = {}
    for t in tasks:
        for cid in t["course_ids"]:
            by_course.setdefault(cid, []).append(t)

    # ── розклад ──
    days, seen, missing = {}, set(), set()
    for s in sorted(sessions, key=lambda x: x["when"]):
        course = index.get(normalize(s["title"]))
        if course is None:
            missing.add(s["title"])
        else:
            seen.add(course["id"])
        meta = s["kind"]
        if course and course["teacher"]:
            meta = f'{s["kind"]}, {course["teacher"]}'
        days.setdefault(s["when"].date(), []).append({
            "time": s["when"].strftime("%H:%M"),
            "title": course["name"] if course else s["title"],
            "teacher": meta,
            "link": s["link"] or (course["link"] if course else ""),
        })

    # ── терміни ──
    horizon = monday + timedelta(days=NA_OBRII)
    deadlines = [{"date": t["due"], "title": t["title"], "where": t["where"]}
                 for t in tasks if t["due"] <= horizon]

    # ── предмети: тільки ті, де є що сказати ──
    course_rows = []
    for cid, course in courses.items():
        mine = sorted(by_course.get(cid, []), key=lambda t: t["due"])
        text = course_text(course, mine, monday)
        if not text and cid not in seen:
            continue
        if not text:
            continue
        course_rows.append({"name": course["name"],
                            "state": course_state(course, mine, monday),
                            "text": text})

    cfg = {
        "week": {"number": (monday - TRYMESTR).days // 7 + 1,
                 "start": monday, "end": sunday, "lead": ""},
        "day": [{"date": d, "session": days[d]} for d in sorted(days)],
        "deadline": deadlines,
        "course": course_rows,
        "news": notion_data.opportunities(monday, sunday),
        "footer": {
            "signature": os.environ.get("PIDPYS", "Олександр"),
            "contact": "Питання — у чат групи або мені особисто.",
            "ps": os.environ.get("PS", ""),
            "link": [
                {"label": "Календар занять", "url": rozklad.CALENDAR_URL},
                {"label": "База знань", "url": os.environ.get(
                    "NOTION_PAGE_URL",
                    "https://app.notion.com/p/3db6f11353328066be90f89c99e29cd2")},
            ],
        },
    }
    cfg["week"]["lead"] = write_lead(facts_for_lead(cfg, len(sessions)))
    return cfg, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tyzhden", metavar="РРРР-ММ-ДД")
    ap.add_argument("--suho", action="store_true",
                    help="не створювати чернетку, лише зберегти файли")
    args = ap.parse_args()

    today = datetime.now(KYIV).date()
    monday = (date.fromisoformat(args.tyzhden) if args.tyzhden
              else next_monday(today))
    monday -= timedelta(days=monday.weekday())

    cfg, missing = assemble(monday)

    if missing:
        print("УВАГА: ці заняття не звʼязалися з жодним курсом у Notion —",
              "додай назви в поле «Аліаси в календарі»:", file=sys.stderr)
        for name in sorted(missing):
            print(f"  · {name}", file=sys.stderr)

    subject, html = build.render(cfg)
    text = build.plain_text(cfg)

    out = Path("out"); out.mkdir(exist_ok=True)
    (out / f"lyst-{monday.isoformat()}.html").write_text(html, encoding="utf-8")
    print(f"Тема: {subject}")
    print(f"Занять: {sum(len(d['session']) for d in cfg['day'])}, "
          f"термінів: {len(cfg['deadline'])}, "
          f"предметів у листі: {len(cfg['course'])}, "
          f"можливостей: {len(cfg['news'])}")

    if args.suho:
        print("Сухий запуск — чернетку не створюю.")
        return

    to = [a.strip() for a in os.environ.get("ROZSYLKA", "").split(",") if a.strip()]
    if not to:
        print("ROZSYLKA порожня — чернетка буде без адресатів.", file=sys.stderr)
    graph_mail.create_draft(subject, html, text, to)
    print("Чернетка створена в Outlook.")


if __name__ == "__main__":
    main()
