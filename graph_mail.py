#!/usr/bin/env python3
"""
Робота з поштою через Microsoft Graph.

Одноразово, на своєму ПК:
    python3 graph_mail.py --avtoryzatsiya
        покаже код, який треба ввести на microsoft.com/devicelogin,
        і надрукує refresh-токен для секрету GRAPH_REFRESH_TOKEN

Далі модуль використовується зі збірки:
    create_draft(subject, html, text, to=[...])
"""

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

TENANT = os.environ.get("GRAPH_TENANT", "b8cbfe43-c90c-4bea-84ae-be5d6d8a5f52")
SCOPES = "offline_access Mail.ReadWrite Mail.Send"
AUTH = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0"


def _post_form(url, fields):
    data = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except urllib.error.HTTPError as err:
        return json.loads(err.read().decode(errors="replace"))


def device_login(client_id):
    """Одноразовий вхід: користувач підтверджує доступ у браузері."""
    start = _post_form(f"{AUTH}/devicecode",
                       {"client_id": client_id, "scope": SCOPES})
    if "user_code" not in start:
        raise RuntimeError(f"Не вдалося почати вхід: {start}")

    print(f"\nВідкрий {start['verification_uri']}")
    print(f"Введи код: {start['user_code']}\n")

    deadline = time.time() + start.get("expires_in", 900)
    while time.time() < deadline:
        time.sleep(start.get("interval", 5))
        res = _post_form(f"{AUTH}/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id, "device_code": start["device_code"]})
        if "refresh_token" in res:
            return res["refresh_token"]
        if res.get("error") not in ("authorization_pending", "slow_down"):
            raise RuntimeError(f"Вхід не вдався: {res.get('error_description', res)}")
    raise RuntimeError("Час на підтвердження вичерпано")


def access_token():
    res = _post_form(f"{AUTH}/token", {
        "grant_type": "refresh_token",
        "client_id": os.environ["GRAPH_CLIENT_ID"],
        "refresh_token": os.environ["GRAPH_REFRESH_TOKEN"],
        "scope": SCOPES})
    if "access_token" not in res:
        raise RuntimeError(
            "Не вдалося оновити токен. Найімовірніше він протух — "
            f"перезапусти --avtoryzatsiya. Відповідь: {res.get('error_description', res)}")
    # Graph повертає новий refresh-токен; якщо є куди — зберігаємо.
    if res.get("refresh_token"):
        _persist_refresh(res["refresh_token"])
    return res["access_token"]


def _persist_refresh(token):
    """У GitHub Actions оновлюємо секрет, щоб токен не застарів."""
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"refresh_token={token}\n")


def create_draft(subject, html, text, to):
    body = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    req = urllib.request.Request(
        f"{GRAPH}/me/messages", data=json.dumps(body).encode(), method="POST",
        headers={"Authorization": f"Bearer {access_token()}",
                 "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except urllib.error.HTTPError as err:
        raise RuntimeError(
            f"Graph відмовив: {err.code} {err.read().decode(errors='replace')}") from None


def find_draft(subject):
    """Шукає чернетку з такою темою. Немає — значить її видалили."""
    query = urllib.parse.urlencode({
        "$filter": "subject eq '{}'".format(subject.replace("'", "''")),
        "$select": "id,subject"})
    req = urllib.request.Request(
        f"{GRAPH}/me/mailFolders/drafts/messages?{query}",
        headers={"Authorization": f"Bearer {access_token()}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        items = json.load(r).get("value", [])
    return items[0]["id"] if items else None


def send_draft(message_id):
    req = urllib.request.Request(
        f"{GRAPH}/me/messages/{message_id}/send", data=b"", method="POST",
        headers={"Authorization": f"Bearer {access_token()}",
                 "Content-Length": "0"})
    try:
        urllib.request.urlopen(req, timeout=30)
    except urllib.error.HTTPError as err:
        raise RuntimeError(
            f"Не вдалося надіслати: {err.code} "
            f"{err.read().decode(errors='replace')}") from None


if __name__ == "__main__":
    if "--avtoryzatsiya" not in sys.argv:
        print(__doc__)
        sys.exit(1)
    client_id = os.environ.get("GRAPH_CLIENT_ID") or input("GRAPH_CLIENT_ID: ").strip()
    token = device_login(client_id)
    print("Готово. Поклади це в секрет GRAPH_REFRESH_TOKEN:\n")
    print(token)
