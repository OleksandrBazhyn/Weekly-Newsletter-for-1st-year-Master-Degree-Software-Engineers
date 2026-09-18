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
from email.message import EmailMessage
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
TRYMESTR = date.fromisoformat(os.environ.get("TRYMESTR_POCHATOK", "2026-08-31"))
NA_OBRII = 14      # горизонт, у межах якого дедлайн взагалі показуємо


# ── звʼязування назв ─────────────────────────────────────────

def normalize(text):
    text = unicodedata.normalize("NFKD", (text or "").casefold())
    return re.sub(r"[^\w]+", "", text)


PREFIXES = ("проф.", "доц.", "ст. викл.", "ст.викл.", "викл.", "асист.")


def short_name(full):
    """«Салата Кирило Володимирович» → «К.В. Салата». Формат інший — лишаємо як є."""
    prefix = ""
    rest = full.strip()
    for p in PREFIXES:
        if rest.lower().startswith(p):
            prefix, rest = rest[:len(p)] + " ", rest[len(p):].strip()
            break
    parts = rest.split()
    # Три слова з великої літери — це ПІБ. «кафедра англійської мови» — ні.
    if len(parts) == 3 and all(len(x) > 1 and x[0].isupper() for x in parts):
        surname, first, patronymic = parts
        rest = f"{first[0]}.{patronymic[0]}. {surname}"
    return (prefix + rest).strip()


def teachers_line(raw):
    """Прізвища через кому; якщо їх більше двох — перший і «та ін.»."""
    names = [short_name(n) for n in raw.split(",") if n.strip()]
    if not names:
        return ""
    if len(names) > 2:
        return f"{names[0]} та ін."
    return ", ".join(names)


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


LEAD_PROMPT = """Ти пишеш перший абзац тижневої розсилки для групи
магістрів-програмістів НаУКМА. Нижче — розклад і терміни цього тижня.

{facts}

Напиши одне–два речення, разом не довше 40 слів.

Чого НЕ робити:
- не перелічувати заняття, предмети чи дедлайни: вони є в таблицях нижче;
- не додавати нічого, чого немає у фактах вище, і не робити припущень
  про те, що там не написано;
- не давати порад на кшталт «варто підготуватися заздалегідь» —
  читачі дорослі люди;
- не писати побажань продуктивного тижня.

Що сказати: яка форма тижня. Де густо, де порожньо, чи є день,
у який збігається кілька справ. Якщо тиждень рівний — одне речення
про це, і все.

Пиши як колега колегам. Поверни лише текст абзацу."""

NEWS_PROMPT = """Нижче — оголошення, скопійовані в базу знань як є.
Стисни кожне до одного речення (до 150 символів) українською.

Залиш те, що потрібно для рішення йти чи ні: що це, коли, і умову
участі, якщо вона є (донат, реєстрація, дедлайн подачі). Прибери
вступи, біографії спікерів, заклики й емодзі. Нічого не додавай від
себе.

Поверни рівно {count} рядків, по одному на оголошення, у тому ж
порядку, без нумерації й лапок.

{items}"""


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
         "max_completion_tokens": 2000,
         "messages": [{"role": "user", "content": prompt}]})
    return data["choices"][0]["message"]["content"].strip()


def _anthropic(prompt):
    data = _post_json(
        "https://api.anthropic.com/v1/messages",
        {"x-api-key": os.environ["ANTHROPIC_API_KEY"],
         "anthropic-version": "2023-06-01"},
        {"model": os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5"),
         "max_tokens": 1000,
         "messages": [{"role": "user", "content": prompt}]})
    return "".join(b.get("text", "") for b in data["content"]).strip()


def ask(prompt, navischo):
    """Модель необовʼязкова: якщо мовчить, лист усе одно збереться."""
    if os.environ.get("OPENAI_API_KEY"):
        provider, call = "OpenAI", _openai
    elif os.environ.get("ANTHROPIC_API_KEY"):
        provider, call = "Anthropic", _anthropic
    else:
        print(f"Ключа до моделі немає — {navischo} пропускаю", file=sys.stderr)
        return ""
    try:
        answer = call(prompt)
    except Exception as err:
        print(f"{provider} не відповів ({err}) — {navischo} пропускаю", file=sys.stderr)
        return ""
    if not answer.strip():
        print(f"{provider} повернув порожню відповідь — {navischo} пропускаю. "
              f"Найчастіше це замалий ліміт токенів для моделі з міркуваннями.",
              file=sys.stderr)
    return answer


def write_lead(facts):
    return ask(LEAD_PROMPT.format(facts=facts), "вступний абзац")


def first_sentences(text, limit=160):
    """Запасний варіант, якщо модель недоступна."""
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    dot = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
    return (cut[:dot + 1] if dot > 60 else cut.rsplit(" ", 1)[0] + "…").strip()


def shorten_news(items):
    """Оголошення в базі — сирі копіпасти; у лист має йти один рядок."""
    long_ones = [n for n in items if len(n["text"]) > 180]
    if not long_ones:
        return items
    listing = "\n\n".join(f'{i + 1}. {n["title"]}\n{n["text"]}'
                            for i, n in enumerate(long_ones))
    answer = ask(NEWS_PROMPT.format(count=len(long_ones), items=listing),
                 "скорочення оголошень")
    lines = [l.strip(" -–—•\"") for l in answer.splitlines() if l.strip()]
    for n, short in zip(long_ones, lines):
        n["text"] = short
    for n in long_ones:                      # на що рядків не вистачило
        if len(n["text"]) > 180:
            n["text"] = first_sentences(n["text"])
    return items


def facts_for_lead(cfg):
    """Усі сім днів, включно з порожніми, щоб модель нічого не домислювала."""
    start = cfg["week"]["start"]
    by_day = {d["date"]: d["session"] for d in cfg["day"]}
    lines = []
    for i in range(7):
        day = start + timedelta(days=i)
        name = build.WEEKDAYS[day.weekday()]
        pairs = by_day.get(day)
        if pairs:
            lines.append(f'{name}: ' + "; ".join(
                f'{s["time"]} {s["title"]}' for s in pairs))
        else:
            lines.append(f"{name}: занять немає")
    for d in sorted(cfg["deadline"], key=lambda x: x["date"]):
        koly = build.WEEKDAYS[d["date"].weekday()] if start <= d["date"] <= cfg["week"]["end"] \
               else d["date"].isoformat()
        lines.append(f'Термін ({koly}): {d["title"]}')
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

    # ── розклад: заняття одного курсу в один день зливаємо в рядок ──
    grouped, seen, missing = {}, set(), set()
    for s in sorted(sessions, key=lambda x: x["when"]):
        course = index.get(normalize(s["title"]))
        if course is None:
            missing.add(s["title"])
        else:
            seen.add(course["id"])
        key = (s["when"].date(), course["id"] if course else s["title"])
        row = grouped.get(key)
        if row is None:
            grouped[key] = {
                "start": s["when"], "until": s["until"],
                "title": course["name"] if course else s["title"],
                "kinds": [s["kind"]],
                "teacher": teachers_line(course["teacher"]) if course else "",
                "link": s["link"] or (course["link"] if course else ""),
            }
        else:
            row["until"] = max(filter(None, [row["until"], s["until"]]), default=None)
            if s["kind"] not in row["kinds"]:
                row["kinds"].append(s["kind"])
            row["link"] = row["link"] or s["link"]

    days = {}
    for (day, _), row in grouped.items():
        span = row["start"].strftime("%H:%M")
        if row["until"] and row["until"] > row["start"]:
            span += "\u2013" + row["until"].strftime("%H:%M")
        kinds = " і ".join(k.lower() for k in row["kinds"]).capitalize()
        meta = f'{kinds}, {row["teacher"]}' if row["teacher"] else kinds
        days.setdefault(day, []).append({
            "time": span, "title": row["title"],
            "teacher": meta, "link": row["link"],
        })
    for day in days:
        days[day].sort(key=lambda x: x["time"])

    # ── терміни: заголовок веде на місце здачі, поруч предмет ──
    horizon = monday + timedelta(days=NA_OBRII)
    names = {cid: c["name"] for cid, c in courses.items()}
    deadlines = [{"date": t["due"], "title": t["title"], "link": t["where"],
                  "course": ", ".join(names[c] for c in t["course_ids"] if c in names)}
                 for t in tasks if t["due"] <= horizon]

    # ── предмети: лише те, що ти написав сам ──
    course_rows = [
        {"name": c["name"], "text": c["now"],
         "state": course_state(c, sorted(by_course.get(cid, []),
                                         key=lambda t: t["due"]), monday)}
        for cid, c in courses.items() if c["now"]
    ]

    cfg = {
        "week": {"number": (monday - TRYMESTR).days // 7 + 1,
                 "start": monday, "end": sunday, "lead": "",
                 "lessons": len(sessions)},
        "day": [{"date": d, "session": days[d]} for d in sorted(days)],
        "deadline": deadlines,
        "course": course_rows,
        "news": shorten_news(notion_data.opportunities(monday, sunday)),
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
    cfg["week"]["lead"] = write_lead(facts_for_lead(cfg))
    return cfg, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tyzhden", metavar="РРРР-ММ-ДД")
    ap.add_argument("--suho", action="store_true",
                    help="не створювати чернетку, лише зберегти файли")
    ap.add_argument("--nadislaty", action="store_true",
                    help="надіслати раніше створену чернетку цього тижня")
    args = ap.parse_args()

    today = datetime.now(KYIV).date()
    monday = (date.fromisoformat(args.tyzhden) if args.tyzhden
              else next_monday(today))
    monday -= timedelta(days=monday.weekday())

    if args.nadislaty:
        subject = (f"Тиждень {(monday - TRYMESTR).days // 7 + 1} — "
                   f"{build.date_range(monday, monday + timedelta(days=6))}")
        draft = graph_mail.find_draft(subject)
        if not draft:
            print(f"Чернетки «{subject}» у теці немає — нічого не надсилаю.")
            return
        graph_mail.send_draft(draft)
        print(f"Надіслано: {subject}")
        return

    cfg, missing = assemble(monday)

    if missing:
        print("УВАГА: ці заняття не звʼязалися з жодним курсом у Notion —",
              "додай назви в поле «Аліаси в календарі»:", file=sys.stderr)
        for name in sorted(missing):
            print(f"  · {name}", file=sys.stderr)

    if not cfg["day"]:
        print("УВАГА: у листі жодного заняття. Під час триместру це майже",
              "напевно збій розбору календаря, а не порожній тиждень.",
              file=sys.stderr)

    subject, html = build.render(cfg)
    text = build.plain_text(cfg)

    out = Path("out"); out.mkdir(exist_ok=True)
    (out / f"lyst-{monday.isoformat()}.html").write_text(html, encoding="utf-8")

    # .eml перетягується в теку «Чернетки» в Outlook — це запасний шлях,
    # коли доступ до Graph закритий політикою тенанту.
    msg = EmailMessage()
    msg["Subject"] = subject
    to = [a.strip() for a in os.environ.get("ROZSYLKA", "").split(",") if a.strip()]
    if to:
        msg["To"] = ", ".join(to)
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    (out / f"lyst-{monday.isoformat()}.eml").write_bytes(msg.as_bytes())
    print(f"Тема: {subject}")
    print(f"Занять: {sum(len(d['session']) for d in cfg['day'])}, "
          f"термінів: {len(cfg['deadline'])}, "
          f"предметів у листі: {len(cfg['course'])}, "
          f"можливостей: {len(cfg['news'])}")

    if args.suho or os.environ.get("BEZ_GRAPH"):
        print("Файли готові. Чернетку через Graph не створюю.")
        return

    to = [a.strip() for a in os.environ.get("ROZSYLKA", "").split(",") if a.strip()]
    if not to:
        print("ROZSYLKA порожня — чернетка буде без адресатів.", file=sys.stderr)
    graph_mail.create_draft(subject, html, text, to)
    print("Чернетка створена в Outlook.")


if __name__ == "__main__":
    main()
