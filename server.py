#!/usr/bin/env python3
"""
Alstom NDT — Skaner Prób v3.0
Hosting: Render.com
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
from io import BytesIO
from datetime import date
import os
import json
import urllib.request
import urllib.error
import openpyxl

app = Flask(__name__, static_folder=".", static_url_path="")

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Wczytaj szablony przy starcie
TEMPLATES = {}
for t in ["MT", "UT"]:
    path = os.path.join(BASE_DIR, f"{t}__wzor.xlsx")
    if os.path.exists(path):
        with open(path, "rb") as f:
            TEMPLATES[t] = f.read()
        print(f"✓ Wczytano {t}__wzor.xlsx")
    else:
        print(f"⚠ Brak {t}__wzor.xlsx!")


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/scan", methods=["POST"])
def scan():
    body = request.get_json()
    image_b64 = body.get("image_base64", "")

    if not API_KEY:
        return jsonify({"error": "Brak klucza ANTHROPIC_API_KEY na serwerze"}), 400

    prompt = """Analizujesz formularz badań nieniszczących spawów firmy Alstom.

Odczytaj z nagłówka formularza:
1. SPAWACZ - tylko samo nazwisko (lub imię i nazwisko) bez tytułów i cyfr
2. PROJEKT - TYLKO sam kod/numer projektu z pierwszej ramki pola PROJEKT (np. "ALP", "ALP2") — ignoruj wszystko co jest napisane obok lub po prawej stronie (np. "PODŁUŻNICA DOLNA" to nie jest część projektu)

Następnie dla każdego wiersza tabeli z numerem próby (np. X01, X05, X13, O3, X31 itp.):
- Kolumna MT: wpis (MT+, MtΦ, symbol z plusem/kółkiem) I komórka NIE zacieniona → dodaj do mt_proby
- Kolumna UT: wpis (UT+, utΦ, symbol z plusem/kółkiem) I komórka NIE zacieniona → dodaj do ut_proby
- Komórki ZACIENIONE/SZARE = pomijamy

Zwróć WYŁĄCZNIE JSON, zero tekstu przed ani po:
{"spawacz":"...","projekt":"...","mt_proby":["X01"],"ut_proby":["X53"]}"""

    payload = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 1000,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "x-api-key": API_KEY,
            "anthropic-version": "2023-06-01"
        }
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read())
        text = data["content"][0]["text"].strip().replace("```json", "").replace("```", "").strip()
        result = json.loads(text)
        return jsonify({"ok": True, "result": result})
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read()).get("error", {}).get("message", str(e))
        except:
            err = str(e)
        return jsonify({"error": err}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export", methods=["POST"])
def export():
    body = request.get_json()
    t       = body.get("type", "MT").upper()
    proby   = body.get("proby", [])
    spawacz = body.get("spawacz", "")
    projekt = body.get("projekt", "")

    if t not in TEMPLATES:
        return jsonify({"error": f"Brak szablonu {t}__wzor.xlsx"}), 400

    wb = openpyxl.load_workbook(BytesIO(TEMPLATES[t]))
    ws = wb["DANE"]

    for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=2):
        for cell in row:
            cell.value = None

    ws["A1"] = date.today()
    ws["A1"].number_format = "DD.MM.YYYY"
    ws["A2"] = ", ".join(proby)
    ws["A3"] = spawacz
    ws["A4"] = projekt

    for i, p in enumerate(proby):
        ws.cell(row=i + 1, column=2).value = p

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Nazwa pliku: TYP_próba1_próba2_..._spawacz.xlsx
    spawacz_safe = spawacz.replace(" ", "_").upper()
    proby_str = "_".join(proby)
    fname = f"{t}_{proby_str}_{spawacz_safe}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=fname
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port)
