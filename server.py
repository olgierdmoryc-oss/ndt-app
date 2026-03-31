#!/usr/bin/env python3
"""
Alstom NDT — Skaner Prób v3.1
Hosting: Render.com
Obrazki w szablonach są w 100% zachowane (ZIP surgery).
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
from io import BytesIO
from datetime import date
import os, json, zipfile, io, re
import urllib.request, urllib.error

app = Flask(__name__, static_folder=".", static_url_path="")

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Wczytaj szablony przy starcie (surowe bajty)
TEMPLATES = {}
for t in ["MT", "UT"]:
    path = os.path.join(BASE_DIR, f"{t}__wzor.xlsx")
    if os.path.exists(path):
        with open(path, "rb") as f:
            TEMPLATES[t] = f.read()
        print(f"✓ Wczytano {t}__wzor.xlsx")
    else:
        print(f"⚠ Brak {t}__wzor.xlsx!")


# ── Shared strings helper ────────────────────────────────────────────────────
def _parse_shared_strings(ss_bytes):
    """Parsuje sharedStrings.xml → lista stringów."""
    strings = []
    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(ss_bytes)
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        for si in root.findall("x:si", ns):
            parts = []
            # <t> bezpośrednio w <si>
            t = si.find("x:t", ns)
            if t is not None and t.text:
                parts.append(t.text)
            # <r><t> runs
            for r in si.findall("x:r", ns):
                rt = r.find("x:t", ns)
                if rt is not None and rt.text:
                    parts.append(rt.text)
            strings.append("".join(parts))
    except Exception as e:
        print(f"Błąd parsowania sharedStrings: {e}")
    return strings


def _build_shared_strings_xml(strings):
    """Buduje sharedStrings.xml z listy stringów."""
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        f' count="{len(strings)}" uniqueCount="{len(strings)}">'
    ]
    for s in strings:
        escaped = (s
                   .replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace('"', "&quot;"))
        lines.append(f'<si><t xml:space="preserve">{escaped}</t></si>')
    lines.append("</sst>")
    return "".join(lines).encode("utf-8")


# ── ZIP Surgery: aktualizuj tylko sheet1.xml (DANE) ─────────────────────────
def patch_xlsx(template_bytes, today_date, proby, spawacz, projekt):
    """
    Wstrzykuje dane do arkusza DANE bez naruszania obrazków.
    Struktura DANE:
      A1 = data (liczba Excel)      → str1!I12
      A2 = numer próby (tekst)      → str1!S17  (wszystkie próby jako string)
      A3 = spawacz (tekst)          → str1!S21
      B1..B20 = próby (tekst)       → str2!B29..B48
    """
    src_zip = zipfile.ZipFile(io.BytesIO(template_bytes))
    
    # Wczytaj shared strings
    ss_bytes = src_zip.read("xl/sharedStrings.xml")
    strings = _parse_shared_strings(ss_bytes)
    
    # Przygotuj wartości
    excel_date = (today_date - date(1899, 12, 30)).days  # Excel serial date
    proby_str  = ", ".join(proby)
    spawacz_str = spawacz or ""
    
    # Znajdź lub dodaj stringi do shared strings
    def get_or_add(val):
        try:
            idx = strings.index(val)
        except ValueError:
            idx = len(strings)
            strings.append(val)
        return idx
    
    idx_proby   = get_or_add(proby_str)
    idx_spawacz = get_or_add(spawacz_str)
    idx_proby_list = [get_or_add(p) for p in proby]
    
    # Buduj nowy sheet1.xml (arkusz DANE)
    # Styl 105 = data, 106 = tekst (zachowujemy oryginalne style)
    rows_xml = []
    def b_cell(row, col_val_idx):
        if col_val_idx is None:
            return f'<c r="B{row}" s="106"/>'
        return f'<c r="B{row}" s="106" t="s"><v>{col_val_idx}</v></c>'

    b1 = idx_proby_list[0] if len(proby) > 0 else get_or_add("")
    rows_xml.append(
        f'<row r="1" spans="1:2">'
        f'<c r="A1" s="105"><v>{excel_date}</v></c>'
        + b_cell(1, b1) +
        f'</row>'
    )
    rows_xml.append(
        f'<row r="2" spans="1:2">'
        f'<c r="A2" s="106" t="s"><v>{idx_proby}</v></c>'
        + b_cell(2, idx_proby_list[1] if len(proby) > 1 else None) +
        f'</row>'
    )
    rows_xml.append(
        f'<row r="3" spans="1:2">'
        f'<c r="A3" s="106" t="s"><v>{idx_spawacz}</v></c>'
        + b_cell(3, idx_proby_list[2] if len(proby) > 2 else None) +
        f'</row>'
    )
    
    for i in range(3, 20):
        row_num = i + 1
        if i < len(proby):
            rows_xml.append(
                f'<row r="{row_num}" spans="2:2"><c r="B{row_num}" s="106" t="s">'
                f'<v>{idx_proby_list[i]}</v></c></row>'
            )
        else:
            rows_xml.append(f'<row r="{row_num}" spans="2:2"><c r="B{row_num}" s="106"/></row>')
    
    # Pozostałe wiersze 21-29 puste
    for row_num in range(21, 30):
        rows_xml.append(f'<row r="{row_num}" spans="2:2"><c r="B{row_num}" s="106"/></row>')
    
    sheet1_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac xr xr2 xr3"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"'
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"'
        ' xr:uid="{B657F5AE-D978-2546-81E1-E5CDD56975F5}">'
        '<dimension ref="A1:B29"/>'
        '<sheetViews><sheetView workbookViewId="0">'
        '<selection activeCell="C2" sqref="C2:C29"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr baseColWidth="10" defaultRowHeight="13"/>'
        '<sheetData>'
        + "".join(rows_xml) +
        '</sheetData>'
        '</worksheet>'
    ).encode("utf-8")
    
    # Nowy shared strings XML
    new_ss_xml = _build_shared_strings_xml(strings)
    
    # Złóż nowy ZIP — kopiuj wszystko, podmień sheet1.xml i sharedStrings.xml
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for item in src_zip.infolist():
            if item.filename == "xl/worksheets/sheet1.xml":
                out_zip.writestr(item.filename, sheet1_xml)
            elif item.filename == "xl/sharedStrings.xml":
                out_zip.writestr(item.filename, new_ss_xml)
            elif item.filename == "xl/calcChain.xml":
                pass  # Usuń calcChain - Excel odbuduje
            else:
                out_zip.writestr(item, src_zip.read(item.filename))
    
    src_zip.close()
    out_buf.seek(0)
    return out_buf


# ── Routes ───────────────────────────────────────────────────────────────────
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
    body    = request.get_json()
    t       = body.get("type", "MT").upper()
    proby   = body.get("proby", [])
    spawacz = body.get("spawacz", "") or ""
    projekt = body.get("projekt", "") or ""

    if t not in TEMPLATES:
        return jsonify({"error": f"Brak szablonu {t}__wzor.xlsx"}), 400

    try:
        buf = patch_xlsx(
            template_bytes=TEMPLATES[t],
            today_date=date.today(),
            proby=proby,
            spawacz=spawacz,
            projekt=projekt
        )
    except Exception as e:
        return jsonify({"error": f"Błąd generowania pliku: {e}"}), 500

    spawacz_safe = spawacz.replace(" ", "_").upper()
    proby_str    = "_".join(proby)
    fname        = f"{t}_{proby_str}_{spawacz_safe}.xlsx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=fname
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port)
