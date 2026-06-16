#!/usr/bin/env python3
"""
Alstom NDT — Skaner Prób v3.2
Hosting: Render.com
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
from io import BytesIO
from datetime import date as d
import os
import json
import zipfile
import re
import urllib.request
import urllib.error

app = Flask(__name__, static_folder=".", static_url_path="")

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES = {}
for t in ["MT", "UT", "PT"]:
    path = os.path.join(BASE_DIR, f"{t}__wzor.xlsx")
    if os.path.exists(path):
        with open(path, "rb") as f:
            TEMPLATES[t] = f.read()
        print(f"✓ Wczytano {t}__wzor.xlsx")
    else:
        print(f"⚠ Brak {t}__wzor.xlsx!")


@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({"error": str(e)}), 500


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
2. PROJEKT - TYLKO sam kod/numer projektu z pierwszej ramki pola PROJEKT (np. "ALP", "ALP2") — ignoruj wszystko co jest napisane obok lub po prawej stronie

Następnie dla każdego wiersza tabeli z numerem próby:
- Kolumna MT: wpis (MT+, MtΦ itp.) I komórka NIE zacieniona → dodaj do mt_proby
- Kolumna UT: wpis (UT+, utΦ itp.) I komórka NIE zacieniona → dodaj do ut_proby
- Kolumna PT: wpis (PT+, PtΦ itp.) I komórka NIE zacieniona → dodaj do pt_proby
- Komórki ZACIENIONE/SZARE = pomijamy

Zwróć WYŁĄCZNIE JSON, zero tekstu przed ani po:
{"spawacz":"...","projekt":"...","mt_proby":[],"ut_proby":[],"pt_proby":[]}"""

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
        result.setdefault("pt_proby", [])
        return jsonify({"ok": True, "result": result})
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read()).get("error", {}).get("message", str(e))
        except:
            err = str(e)
        return jsonify({"error": err}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def get_dane_sheet_path(raw_bytes):
    wb_xml = zipfile.ZipFile(BytesIO(raw_bytes)).read('xl/workbook.xml').decode()
    sheets = re.findall(r'<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
    rels_xml = zipfile.ZipFile(BytesIO(raw_bytes)).read('xl/_rels/workbook.xml.rels').decode()
    rels = re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml)
    rid_to_target = {r[0]: r[1] for r in rels}
    for name, rid in sheets:
        if name == 'DANE':
            target = rid_to_target.get(rid, '')
            if not target.startswith('xl/'):
                target = 'xl/' + target
            return target
    raise Exception("Nie znaleziono arkusza DANE")


def patch_dane_sheet(xml_bytes, proby, spawacz, projekt):
    xml = xml_bytes.decode('utf-8')
    epoch = d(1899, 12, 30)
    serial = (d.today() - epoch).days
    proby_str = ', '.join(proby)

    def esc(v):
        return str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def inline_cell(ref, val):
        return f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>'

    def date_cell(ref, val):
        return f'<c r="{ref}" s="1"><v>{val}</v></c>'

    def replace_row(xml_content, row_num, cells_xml):
        row_pattern = rf'<row([^>]+r="{row_num}"[^>]*)>.*?</row>'
        m = re.search(row_pattern, xml_content, re.DOTALL)
        if m:
            new_row = f'<row{m.group(1)}>{cells_xml}</row>'
            return xml_content[:m.start()] + new_row + xml_content[m.end():]
        else:
            new_row = f'<row r="{row_num}">{cells_xml}</row>'
            return xml_content.replace('</sheetData>', new_row + '</sheetData>')

    # Wyczyść stare komórki B4-B30
    for r in range(4, 31):
        xml = re.sub(rf'<c r="B{r}"[^/]*/>', '', xml)
        xml = re.sub(rf'<c r="B{r}"[^>]*>.*?</c>', '', xml, flags=re.DOTALL)

    # Wiersze 1-4: A=dane, B=próby[0-3]
    b1 = inline_cell('B1', proby[0]) if len(proby) > 0 else ''
    b2 = inline_cell('B2', proby[1]) if len(proby) > 1 else ''
    b3 = inline_cell('B3', proby[2]) if len(proby) > 2 else ''
    b4 = inline_cell('B4', proby[3]) if len(proby) > 3 else ''

    xml = replace_row(xml, 1, date_cell('A1', serial) + b1)
    xml = replace_row(xml, 2, inline_cell('A2', proby_str) + b2)
    xml = replace_row(xml, 3, inline_cell('A3', spawacz) + b3)
    xml = replace_row(xml, 4, inline_cell('A4', projekt) + b4)

    # Próby B5+
    for i in range(4, len(proby)):
        row_num = i + 1
        new_cell = inline_cell(f'B{row_num}', proby[i])
        row_pattern = rf'<row([^>]+r="{row_num}"[^>]*)>.*?</row>'
        m = re.search(row_pattern, xml, re.DOTALL)
        if m:
            xml = xml[:m.start()] + f'<row{m.group(1)}>{new_cell}</row>' + xml[m.end():]
        else:
            xml = xml.replace('</sheetData>', f'<row r="{row_num}">{new_cell}</row></sheetData>')

    return xml.encode('utf-8')


def build_xlsx(t, proby, spawacz, projekt):
    raw = TEMPLATES[t]
    dane_file = get_dane_sheet_path(raw)
    out = BytesIO()
    with zipfile.ZipFile(BytesIO(raw), 'r') as zin, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == dane_file:
                data = patch_dane_sheet(data, proby, spawacz, projekt)
            zout.writestr(item, data)
    out.seek(0)
    return out


@app.route("/api/export", methods=["POST"])
def export():
    body = request.get_json()
    t       = body.get("type", "MT").upper()
    proby   = body.get("proby", [])
    spawacz = body.get("spawacz", "")
    projekt = body.get("projekt", "")

    if t not in TEMPLATES:
        return jsonify({"error": f"Brak szablonu {t}__wzor.xlsx"}), 400

    try:
        buf = build_xlsx(t, proby, spawacz, projekt)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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
