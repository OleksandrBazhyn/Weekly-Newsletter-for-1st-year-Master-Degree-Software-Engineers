#!/usr/bin/env python3
"""
Збирає тижневий лист групи з tyzhden.toml.

    python3 build.py                 # бере tyzhden.toml поруч зі скриптом
    python3 build.py inshyi.toml     # або вказаний файл

На виході — два файли в теці out/:
    lyst-РРРР-ММ-ДД.html   відкрити в браузері, Ctrl+A / Ctrl+C, вставити в Outlook
    lyst-РРРР-ММ-ДД.eml    перетягнути в теку «Чернетки» в Outlook для ПК
"""

import html
import re
import sys
import tomllib
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path

# ── палітра ──────────────────────────────────────────────────
NAVY = "#152C52"         # глибокий синій, шапка й плашки днів
BLUE = "#2E5AA8"         # робочий синій, час занять і статуси
BLUE_SOFT = "#E7EEF9"    # мʼяка синя підкладка
BLUE_PALE = "#A9BBDA"    # текст на синьому тлі
BRASS = "#C99A3F"        # латунь, стрічки й можливості
BRASS_INK = "#8A6A22"    # латунь для тексту, щоб читалося
BRASS_SOFT = "#FBF3E1"   # підкладка блоку Могилянки
SIGNAL = "#A33B2A"       # терміни
SIGNAL_SOFT = "#FBEFEC"  # підкладка термінів
CALM = "#3F7A55"         # «спокійно»
CALM_SOFT = "#E9F2EC"    # підкладка «спокійно»
INK = "#1B2436"          # основний текст
MUTED = "#63708A"        # другорядний текст
LINE = "#DFE5EE"         # тонкі лінії
PAGE = "#E4E9F1"         # тло сторінки

SANS = "'Segoe UI', -apple-system, Roboto, Helvetica, Arial, sans-serif"
SERIF = "Georgia, 'Times New Roman', serif"

# ── назви секцій: міняй тут, якщо хочеться своїх формулювань ──
TITLE_SCHEDULE = "Розклад тижня"
TITLE_DEADLINES = "Що здаємо"
TITLE_COURSES = "Як справи з предметами"
TITLE_NEWS = "Цікаве в Могилянці"

STATES = {
    "робота": (BLUE, BLUE_SOFT, "у роботі"),
    "спокійно": (CALM, CALM_SOFT, "спокійно"),
    "увага": (SIGNAL, SIGNAL_SOFT, "потребує уваги"),
}

WEEKDAYS = ["понеділок", "вівторок", "середа", "четвер",
            "п'ятниця", "субота", "неділя"]
WEEKDAYS_ACC = ["понеділок", "вівторок", "середу", "четвер",
                "п'ятницю", "суботу", "неділю"]
WEEKDAYS_SHORT = ["пн", "вт", "ср", "чт", "пт", "сб", "нд"]
MONTHS = ["січня", "лютого", "березня", "квітня", "травня", "червня",
          "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]


def e(text):
    return html.escape(str(text).strip())


def plural(n, one, few, many):
    """Українські форми: 1 заняття, 2 заняття, 5 занять."""
    if n % 100 in range(11, 15):
        return many
    last = n % 10
    if last == 1:
        return one
    if last in (2, 3, 4):
        return few
    return many


def as_date(value):
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def date_range(start, end):
    if start.month == end.month:
        return f"{start.day}\u2013{end.day} {MONTHS[end.month - 1]}"
    return (f"{start.day} {MONTHS[start.month - 1]} \u2013 "
            f"{end.day} {MONTHS[end.month - 1]}")


def countdown(target, today):
    days = (target - today).days
    if days < 0:
        return "термін минув", SIGNAL
    if days == 0:
        return "сьогодні", SIGNAL
    if days == 1:
        return "завтра", SIGNAL
    if days <= 7:
        return f"у {WEEKDAYS_ACC[target.weekday()]}", SIGNAL if days <= 3 else MUTED
    return f"{target.day} {MONTHS[target.month - 1]}", MUTED


def paragraphs(text):
    text = re.sub(r"[ \t]+", " ", str(text)).strip()
    return [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]


def week_summary(cfg, end, today):
    """Рядок-резюме під датою: скільки занять і коли найближчий термін."""
    lessons = sum(len(d.get("session", [])) for d in cfg.get("day", []))
    bits = []
    if lessons:
        bits.append(f"{lessons} {plural(lessons, 'заняття', 'заняття', 'занять')}")

    upcoming = sorted(as_date(d["date"]) for d in cfg.get("deadline", []))
    upcoming = [d for d in upcoming if d >= today]
    if upcoming:
        nearest = upcoming[0]
        when = (countdown(nearest, today)[0] if nearest <= end
                else f"{nearest.day} {MONTHS[nearest.month - 1]}")
        bits.append(f"найближчий термін {when}")
    return ", ".join(bits)


# ── блоки листа ──────────────────────────────────────────────

def masthead(week, start, end, summary):
    return f"""
<tr><td bgcolor="{BRASS}" height="6" style="background:{BRASS};font-size:0;line-height:6px;">&nbsp;</td></tr>
<tr><td class="px" bgcolor="{NAVY}" style="background:{NAVY};padding:30px 38px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">
    <tr>
      <td valign="middle" style="font:600 14px/1.6 {SANS};color:{BLUE_PALE};">
        Інженерія програмного забезпечення<br>магістратура, факультет інформатики НаУКМА
      </td>
      <td valign="middle" align="right" style="white-space:nowrap;">
        <div style="font:600 12px/1.2 {SANS};color:{BLUE_PALE};padding-bottom:3px;">тиждень</div>
        <div style="font:400 54px/1 {SERIF};color:{BRASS};">{e(week.get('number', ''))}</div>
      </td>
    </tr>
  </table>
  <div style="font:400 29px/1.3 {SERIF};color:#FFFFFF;padding-top:22px;">
    {e(date_range(start, end))}
  </div>
  {f'<div style="font:400 15px/1.6 {SANS};color:{BLUE_PALE};padding-top:7px;">{e(summary)}</div>' if summary else ''}
</td></tr>"""


def lead_block(text):
    if not text or not str(text).strip():
        return ""
    body = "".join(
        f'<p style="margin:0 0 14px 0;font:400 17px/1.75 {SANS};color:{INK};">{e(p)}</p>'
        for p in paragraphs(text)
    )
    return f'<tr><td class="px" style="padding:30px 38px 4px 38px;">{body}</td></tr>'


def heading(title, color=NAVY):
    return f"""
<tr><td class="px" style="padding:30px 38px 0 38px;">
  <div style="font:700 21px/1.3 {SANS};color:{color};">{e(title)}</div>
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="46"
         style="border-collapse:collapse;margin-top:9px;">
    <tr><td bgcolor="{color}" height="3" style="background:{color};font-size:0;line-height:3px;">&nbsp;</td></tr>
  </table>
</td></tr>"""


def schedule(days):
    if not days:
        return ""
    rows = []
    for item in days:
        d = as_date(item["date"])
        sessions = []
        for j, s in enumerate(item.get("session", [])):
            title = e(s.get("title", ""))
            if s.get("link"):
                title = (f'<a href="{e(s["link"])}" style="color:{INK};'
                         f'text-decoration:none;border-bottom:2px solid {LINE};">{title}</a>')
            meta = e(s.get("teacher", ""))
            note = s.get("note")
            gap = "0" if j == 0 else "15px"
            sessions.append(f"""
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
       style="border-collapse:collapse;margin-top:{gap};">
  <tr>
    <td width="58" valign="top"
        style="font:700 15px/1.5 {SANS};color:{BLUE};">{e(s.get('time', ''))}</td>
    <td valign="top">
      <div style="font:600 16px/1.5 {SANS};color:{INK};">{title}</div>
      {f'<div style="font:400 14px/1.5 {SANS};color:{MUTED};padding-top:3px;">{meta}</div>' if meta else ''}
      {f'<div style="font:600 14px/1.55 {SANS};color:{SIGNAL};padding-top:5px;">{e(note)}</div>' if note else ''}
    </td>
  </tr>
</table>""")
        rows.append(f"""
<tr>
  <td width="58" valign="top" style="padding:0 16px 16px 0;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="58"
           style="border-collapse:collapse;">
      <tr><td bgcolor="{NAVY}" align="center"
              style="background:{NAVY};border-radius:10px;padding:10px 0 11px 0;">
        <div style="font:700 24px/1 {SANS};color:#FFFFFF;">{d.day}</div>
        <div style="font:600 12px/1.3 {SANS};color:{BLUE_PALE};padding-top:4px;">{WEEKDAYS_SHORT[d.weekday()]}</div>
      </td></tr>
    </table>
  </td>
  <td valign="top" style="padding:2px 0 16px 0;">{''.join(sessions)}</td>
</tr>""")
    return f"""
<tr><td class="px" style="padding:20px 38px 0 38px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">{''.join(rows)}</table>
</td></tr>"""


def courses(items):
    if not items:
        return ""
    rows = []
    for i, c in enumerate(items):
        color, soft, label = STATES.get(
            str(c.get("state", "")).lower(), (MUTED, BLUE_SOFT, ""))
        top = "" if i == 0 else f"border-top:1px solid {LINE};"
        chip = (f'<table role="presentation" cellpadding="0" cellspacing="0" border="0"'
                f' style="border-collapse:collapse;display:inline-table;">'
                f'<tr><td bgcolor="{soft}" style="background:{soft};border-radius:20px;'
                f'padding:4px 12px 5px 12px;font:600 12px/1.3 {SANS};color:{color};">'
                f'{e(label)}</td></tr></table>') if label else ""
        rows.append(f"""
<tr><td style="{top}padding:18px 0;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">
    <tr>
      <td valign="middle" style="font:700 17px/1.4 {SANS};color:{INK};padding-right:10px;">
        {e(c.get('name', ''))}</td>
      <td valign="middle" align="right">{chip}</td>
    </tr>
  </table>
  <div style="font:400 15px/1.75 {SANS};color:{INK};padding-top:8px;">{e(c.get('text', ''))}</div>
</td></tr>""")
    return f"""
<tr><td class="px" style="padding:6px 38px 0 38px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">{''.join(rows)}</table>
</td></tr>"""


def deadlines(items, today):
    if not items:
        return ""
    rows = []
    for i, d in enumerate(sorted(items, key=lambda x: as_date(x["date"]))):
        target = as_date(d["date"])
        when, color = countdown(target, today)
        top = "" if i == 0 else "border-top:1px solid #F1DFDB;"
        where = (f'<span style="color:{MUTED};font-weight:400;">, {e(d["where"])}</span>'
                 if d.get("where") else "")
        rows.append(f"""
<tr><td style="{top}padding:14px 0;">
  <div style="font:600 16px/1.5 {SANS};color:{INK};">{e(d.get('title', ''))}</div>
  <div style="font:600 14px/1.5 {SANS};color:{color};padding-top:4px;">{e(when)}{where}</div>
</td></tr>""")
    return f"""
<tr><td class="px" style="padding:20px 38px 0 38px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">
    <tr><td bgcolor="{SIGNAL_SOFT}" style="background:{SIGNAL_SOFT};border-radius:12px;padding:6px 24px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
             style="border-collapse:collapse;">{''.join(rows)}</table>
    </td></tr>
  </table>
</td></tr>"""


def news(items):
    if not items:
        return ""
    rows = []
    for i, n in enumerate(items):
        top = "" if i == 0 else "border-top:1px solid #F0E3C8;"
        title = e(n.get("title", ""))
        if n.get("link"):
            title = (f'<a href="{e(n["link"])}" style="color:{BRASS_INK};'
                     f'text-decoration:none;border-bottom:2px solid #E8D3A6;">{title}</a>')
        rows.append(f"""
<tr><td style="{top}padding:15px 0;">
  <div style="font:600 16px/1.5 {SANS};">{title}</div>
  <div style="font:400 15px/1.7 {SANS};color:{INK};padding-top:5px;">{e(n.get('text', ''))}</div>
</td></tr>""")
    return f"""
<tr><td class="px" style="padding:20px 38px 0 38px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
         style="border-collapse:collapse;">
    <tr><td bgcolor="{BRASS_SOFT}" style="background:{BRASS_SOFT};border-radius:12px;padding:6px 24px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
             style="border-collapse:collapse;">{''.join(rows)}</table>
    </td></tr>
  </table>
</td></tr>"""


def ps_block(text):
    if not text or not str(text).strip():
        return ""
    return f"""
<tr><td class="px" style="padding:28px 38px 0 38px;">
  <div style="font:400 16px/1.7 {SANS};color:{MUTED};">{e(text)}</div>
</td></tr>"""


def footer_block(cfg):
    links = "".join(
        f'<td style="padding:0 10px 10px 0;">'
        f'<table role="presentation" cellpadding="0" cellspacing="0" border="0"'
        f' style="border-collapse:collapse;">'
        f'<tr><td bgcolor="#FFFFFF" style="background:#FFFFFF;border-radius:20px;'
        f'padding:9px 18px;">'
        f'<a href="{e(l["url"])}" style="font:600 14px/1.4 {SANS};color:{NAVY};'
        f'text-decoration:none;">{e(l["label"])}</a></td></tr></table></td>'
        for l in cfg.get("link", [])
    )
    links_row = (f'<table role="presentation" cellpadding="0" cellspacing="0" border="0">'
                 f'<tr>{links}</tr></table>') if links else ""
    return f"""
<tr><td class="px" bgcolor="{NAVY}" style="background:{NAVY};padding:26px 38px 28px 38px;">
  {links_row}
  <div style="font:400 15px/1.65 {SANS};color:{BLUE_PALE};padding-top:{'14px' if links_row else '0'};">
    {e(cfg.get('contact', ''))}
  </div>
  <div style="font:700 17px/1.6 {SANS};color:#FFFFFF;padding-top:8px;">
    {e(cfg.get('signature', ''))}
  </div>
</td></tr>
<tr><td bgcolor="{BRASS}" height="6" style="background:{BRASS};font-size:0;line-height:6px;">&nbsp;</td></tr>"""


def render(cfg):
    week = cfg.get("week", {})
    start, end = as_date(week["start"]), as_date(week["end"])
    today = date.today()
    if not (start <= today <= end):
        today = start

    summary = week_summary(cfg, end, today)
    parts = [masthead(week, start, end, summary), lead_block(week.get("lead"))]

    if cfg.get("day"):
        parts += [heading(TITLE_SCHEDULE), schedule(cfg["day"])]
    if cfg.get("deadline"):
        parts += [heading(TITLE_DEADLINES, SIGNAL), deadlines(cfg["deadline"], today)]
    if cfg.get("course"):
        parts += [heading(TITLE_COURSES), courses(cfg["course"])]
    if cfg.get("news"):
        parts += [heading(TITLE_NEWS, BRASS_INK), news(cfg["news"])]
    parts.append(ps_block(cfg.get("footer", {}).get("ps")))
    parts.append('<tr><td style="height:30px;font-size:0;line-height:30px;">&nbsp;</td></tr>')
    parts.append(footer_block(cfg.get("footer", {})))

    subject = f"Тиждень {week.get('number', '')} — {date_range(start, end)}"

    # Прехедер — те, що видно у списку листів поруч із темою.
    preheader = summary or (paragraphs(week.get("lead", "")) or [""])[0][:110]

    body = f"""<!DOCTYPE html>
<html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<meta name="supported-color-schemes" content="light">
<title>{e(subject)}</title>
<style>
  @media only screen and (max-width:620px) {{
    .px {{ padding-left:20px !important; padding-right:20px !important; }}
    .shell {{ border-radius:0 !important; }}
  }}
</style>
</head>
<body style="margin:0;padding:0;background:{PAGE};">
<div style="display:none;max-height:0;max-width:0;overflow:hidden;opacity:0;
     font-size:1px;line-height:1px;color:{PAGE};">{e(preheader)}</div>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
       style="border-collapse:collapse;background:{PAGE};">
  <tr><td align="center" style="padding:26px 10px;">
    <table role="presentation" class="shell" cellpadding="0" cellspacing="0" border="0" width="620"
           style="border-collapse:collapse;background:#FFFFFF;max-width:620px;width:100%;
                  border-radius:16px;overflow:hidden;">
      {''.join(parts)}
    </table>
  </td></tr>
</table>
</body></html>"""
    return subject, body


def plain_text(cfg):
    week = cfg.get("week", {})
    start, end = as_date(week["start"]), as_date(week["end"])
    out = [f"Тиждень {week.get('number', '')} — {date_range(start, end)}", ""]
    for p in paragraphs(week.get("lead", "")):
        out += [p, ""]
    if cfg.get("day"):
        out.append(TITLE_SCHEDULE.upper())
        for item in cfg["day"]:
            d = as_date(item["date"])
            out.append(f"  {d.day} {MONTHS[d.month - 1]}, {WEEKDAYS[d.weekday()]}")
            for s in item.get("session", []):
                out.append(f"    {s.get('time', '')}  {s.get('title', '')}"
                           f"{' — ' + s['teacher'] if s.get('teacher') else ''}")
                if s.get("note"):
                    out.append(f"           {s['note']}")
                if s.get("link"):
                    out.append(f"           {s['link']}")
        out.append("")
    if cfg.get("deadline"):
        out.append(TITLE_DEADLINES.upper())
        for d in sorted(cfg["deadline"], key=lambda x: as_date(x["date"])):
            t = as_date(d["date"])
            out.append(f"  {t.day} {MONTHS[t.month - 1]} — {d.get('title', '')}"
                       f"{' (' + d['where'] + ')' if d.get('where') else ''}")
        out.append("")
    if cfg.get("course"):
        out.append(TITLE_COURSES.upper())
        for c in cfg["course"]:
            state = STATES.get(str(c.get("state", "")).lower(), (None, None, ""))[2]
            out.append(f"  {c.get('name', '')}{' — ' + state if state else ''}")
            out.append(f"    {c.get('text', '')}")
        out.append("")
    if cfg.get("news"):
        out.append(TITLE_NEWS.upper())
        for n in cfg["news"]:
            out.append(f"  {n.get('title', '')}")
            out.append(f"    {n.get('text', '')}")
            if n.get("link"):
                out.append(f"    {n['link']}")
        out.append("")
    f = cfg.get("footer", {})
    if f.get("ps"):
        out += [f["ps"], ""]
    for l in f.get("link", []):
        out.append(f"{l['label']}: {l['url']}")
    out += ["", f.get("contact", ""), f.get("signature", "")]
    return "\n".join(out)


def main():
    here = Path(__file__).parent
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else here / "tyzhden.toml"
    cfg = tomllib.loads(source.read_text(encoding="utf-8"))

    subject, body = render(cfg)
    start = as_date(cfg["week"]["start"])
    out_dir = here / "out"
    out_dir.mkdir(exist_ok=True)

    html_path = out_dir / f"lyst-{start.isoformat()}.html"
    html_path.write_text(body, encoding="utf-8")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg.set_content(plain_text(cfg))
    msg.add_alternative(body, subtype="html")
    eml_path = out_dir / f"lyst-{start.isoformat()}.eml"
    eml_path.write_bytes(msg.as_bytes())

    print(f"Тема:  {subject}")
    print(f"HTML:  {html_path}")
    print(f"EML:   {eml_path}")


if __name__ == "__main__":
    main()
