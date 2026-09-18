#!/usr/bin/env python3
"""Читання Курсів, Завдань і Можливостей із Notion."""

import json
import os
import sys
import urllib.error
import urllib.request
from datetime import date

API = "https://api.notion.com/v1"
VERSION = "2022-06-28"

DB = {
    "kursy": "3db6f113533280a9b0fac3be4873aa4a",
    "zavdannia": "3db6f113533280d2ad2cf21195f5e4f1",
    "mozhlyvosti": "3db6f113533280799831e0691c2e1a0e",
}


def _call(path, payload=None):
    token = os.environ["NOTION_TOKEN"]
    req = urllib.request.Request(
        f"{API}/{path}",
        data=json.dumps(payload).encode() if payload is not None else None,
        method="POST" if payload is not None else "GET",
        headers={"Authorization": f"Bearer {token}",
                 "Notion-Version": VERSION,
                 "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except urllib.error.HTTPError as err:
        body = err.read().decode(errors="replace")
        raise RuntimeError(f"Notion {path}: {err.code} {body}") from None


def query_all(database_id, payload=None):
    """Проходить усі сторінки пагінації."""
    out, cursor = [], None
    while True:
        body = dict(payload or {})
        if cursor:
            body["start_cursor"] = cursor
        data = _call(f"databases/{database_id}/query", body)
        out.extend(data["results"])
        if not data.get("has_more"):
            return out
        cursor = data["next_cursor"]


# ── дістати значення властивості незалежно від типу ──────────

def value(page, name):
    prop = page["properties"].get(name)
    if not prop:
        return None
    kind = prop["type"]
    if kind in ("title", "rich_text"):
        return "".join(t["plain_text"] for t in prop[kind]).strip()
    if kind == "url":
        return (prop["url"] or "").strip()
    if kind == "email":
        return (prop["email"] or "").strip()
    if kind == "select":
        return prop["select"]["name"] if prop["select"] else ""
    if kind == "multi_select":
        return [o["name"] for o in prop["multi_select"]]
    if kind == "date":
        d = prop["date"]
        return date.fromisoformat(d["start"][:10]) if d and d.get("start") else None
    if kind == "relation":
        return [r["id"] for r in prop["relation"]]
    return None


def page_title(page_id, cache={}):
    """Назва повʼязаної сторінки — потрібна для Викладача."""
    if page_id in cache:
        return cache[page_id]
    page = _call(f"pages/{page_id.replace('-', '')}")
    for prop in page["properties"].values():
        if prop["type"] == "title":
            cache[page_id] = "".join(t["plain_text"] for t in prop["title"]).strip()
            return cache[page_id]
    cache[page_id] = ""
    return ""


# ── три вибірки ──────────────────────────────────────────────

def courses():
    out = {}
    for page in query_all(DB["kursy"]):
        name = value(page, "Курс")
        if not name:
            continue
        aliases = [a.strip() for a in (value(page, "Аліаси в календарі") or "").split("\n")
                   if a.strip()] or [name]
        teachers = [page_title(i) for i in (value(page, "Викладач") or [])]
        out[page["id"]] = {
            "id": page["id"],
            "name": name,
            "aliases": aliases,
            "teacher": ", ".join(t for t in teachers if t),
            "link": value(page, "Посилання на зустріч") or "",
            "materials": value(page, "Матеріали") or "",
            "now": value(page, "Що зараз") or "",
        }
    return out


def open_tasks():
    rows = query_all(DB["zavdannia"], {
        "filter": {"property": "Статус", "select": {"equals": "Відкрито"}}})
    out = []
    for page in rows:
        due = value(page, "Дедлайн")
        if not due:
            continue
        out.append({
            "title": value(page, "Завдання"),
            "due": due,
            "where": value(page, "Де здавати") or "",
            "type": value(page, "Тип") or "",
            "course_ids": value(page, "Курс") or [],
        })
    return sorted(out, key=lambda t: t["due"])


def opportunities(week_start, week_end):
    out = []
    for page in query_all(DB["mozhlyvosti"]):
        title = value(page, "Назва")
        since, until = value(page, "Показувати з"), value(page, "Показувати до")
        if not since or not until:
            print(f"Можливість «{title}» без дат показу — пропускаю",
                  file=sys.stderr)
            continue
        if until < since:
            print(f"Можливість «{title}»: «показувати до» ({until}) раніше за "
                  f"«показувати з» ({since}) — такий запис не покажеться ніколи",
                  file=sys.stderr)
            continue
        if since > week_end or until < week_start:
            continue
        out.append({
            "title": value(page, "Назва"),
            "text": value(page, "Опис") or "",
            "link": value(page, "Посилання") or "",
            "source": value(page, "Джерело") or "",
        })
    return out
