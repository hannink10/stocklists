#!/usr/bin/env python3
"""Erzeugt aus der aktuellen CSV im Ordner input/ eine Excel-Stockliste
auf Basis von template/stocklist_template.xlsx.

Aufruf:
    python generate_stocklist.py              # CSV aus input/ verwenden
    python generate_stocklist.py --csv X.csv  # bestimmte CSV verwenden
    python generate_stocklist.py --no-archive # CSV nach Erfolg nicht archivieren

Mapping und Regeln stehen in config/mapping.json.
"""

import argparse
import json
import re
import shutil
import sys
import unicodedata
from collections import OrderedDict
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config" / "mapping.json"


class StocklistError(Exception):
    """Kritischer Fehler: es wird keine Kundendatei erzeugt."""


# ---------------------------------------------------------------------------
# Eingabe
# ---------------------------------------------------------------------------

def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def find_input_csv(cfg, explicit):
    if explicit:
        path = Path(explicit)
        if not path.is_file():
            raise StocklistError(f"CSV-Datei nicht gefunden: {path}")
        return path

    input_dir = BASE_DIR / cfg["paths"]["input_dir"]
    csvs = sorted(p for p in input_dir.glob("*") if p.suffix.lower() == ".csv")
    if not csvs:
        raise StocklistError(f"Keine CSV-Datei im Ordner '{input_dir}' gefunden.")
    if len(csvs) > 1:
        names = "\n".join(f"  - {p.name}" for p in csvs)
        raise StocklistError(
            "Im Input-Ordner liegen mehrere CSV-Dateien:\n" + names + "\n"
            "Bitte nur die aktuelle CSV im Ordner lassen (alte nach input/archiv verschieben) "
            "oder die gewünschte Datei angeben: python generate_stocklist.py --csv input/DATEI.csv"
        )
    return csvs[0]


def read_csv(path, cfg):
    c = cfg["csv"]
    last_error = None
    for enc in c["encodings"]:
        try:
            raw = pd.read_csv(
                path, sep=c["delimiter"], encoding=enc, header=None,
                dtype=str, keep_default_na=False, skip_blank_lines=True,
            )
            # z.B. 'E' + Akzent als ein Zeichen 'É' speichern
            raw = raw.map(lambda v: unicodedata.normalize("NFC", v))
            return raw, enc
        except UnicodeDecodeError as e:
            last_error = e
        except pd.errors.EmptyDataError:
            raise StocklistError(f"Die CSV-Datei ist leer: {path}")
        except pd.errors.ParserError as e:
            raise StocklistError(f"CSV konnte nicht gelesen werden (Format/Trennzeichen?): {e}")
    raise StocklistError(f"CSV-Encoding nicht erkannt ({', '.join(c['encodings'])}): {last_error}")


def parse_number(text, cfg):
    """Deutsches Zahlenformat ('1.234,50') -> Decimal. Leer -> None."""
    s = text.strip()
    if s == "":
        return None
    s = s.replace(cfg["csv"]["thousands"], "").replace(cfg["csv"]["decimal"], ".").replace("\u2212", "-")
    if s.endswith("-"):  # nachgestelltes Minus, z.B. '3,00-'
        s = "-" + s[:-1].strip()
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(text)


# ---------------------------------------------------------------------------
# Prüfung & Aufbereitung
# ---------------------------------------------------------------------------

def check_header(raw, cfg, warnings):
    header = [h.strip() for h in raw.iloc[0].tolist()]
    while header and header[-1] == "":  # abschließendes ';' erzeugt leere Spalte
        header.pop()

    expected = cfg["csv"]["expected_header"]
    if header != expected:
        missing = [h for h in OrderedDict.fromkeys(expected) if header.count(h) < expected.count(h)]
        new = [h for h in OrderedDict.fromkeys(header) if header.count(h) > expected.count(h)]
        if missing:
            warnings.append("Spalten fehlen oder wurden umbenannt: " + ", ".join(missing))
        if new:
            warnings.append("Neue/unbekannte Spalten (werden ignoriert): " + ", ".join(new))
        if not missing and not new:
            warnings.append("Reihenfolge der CSV-Spalten hat sich geändert.")

    positions = {}
    errors = []
    for key, spec in cfg["columns"].items():
        hits = [i for i, h in enumerate(header) if h == spec["header"]]
        if len(hits) < spec["occurrence"]:
            errors.append(
                f"Pflichtspalte '{spec['header']}' (Vorkommen {spec['occurrence']}) fehlt "
                f"(gefunden: {len(hits)}x)."
            )
        else:
            positions[key] = hits[spec["occurrence"] - 1]
    # Mehrdeutige Spalten (z.B. 'UNIT PRICE' doppelt): Anzahl muss exakt wie erwartet sein,
    # sonst ist unklar, welche Spalte gemeint ist.
    for key, spec in cfg["columns"].items():
        n_exp = expected.count(spec["header"])
        n_act = header.count(spec["header"])
        if n_exp > 1 and n_act != n_exp:
            errors.append(
                f"Spalte '{spec['header']}' kommt {n_act}x statt {n_exp}x vor – "
                "Zuordnung nicht mehr eindeutig. Bitte mapping.json prüfen."
            )
    if errors:
        raise StocklistError("Spaltenprüfung fehlgeschlagen:\n  - " + "\n  - ".join(errors))
    return positions


def build_articles(raw, positions, cfg, template_sizes, warnings):
    sep = cfg["sizes"]["sku_separator"]
    no_size = cfg["sizes"]["no_size_column"]
    extra = [e["size"] for e in cfg["sizes"]["extra_sizes"]]
    known_sizes = set(template_sizes) | set(extra)

    errors = []
    negative = []
    seen_skus = {}
    articles = OrderedDict()

    for idx in range(1, len(raw)):
        row = raw.iloc[idx]
        line = idx + 1  # Zeilennummer in der CSV-Datei
        val = {k: str(row.iloc[p]).strip() for k, p in positions.items()}

        if all(v == "" for v in val.values()):
            continue
        for k in ("sku", "name", "uvp", "unit_price", "stock"):
            if val[k] == "":
                errors.append(f"Zeile {line}: Pflichtfeld '{cfg['columns'][k]['header']}' ist leer.")
        if any(val[k] == "" for k in ("sku", "name", "uvp", "unit_price", "stock")):
            continue

        sku = val["sku"]
        if sku in seen_skus:
            errors.append(f"Zeile {line}: SKU '{sku}' doppelt (bereits in Zeile {seen_skus[sku]}).")
            continue
        seen_skus[sku] = line

        try:
            uvp = parse_number(val["uvp"], cfg)
            price = parse_number(val["unit_price"], cfg)
            stock = parse_number(val["stock"], cfg)
        except ValueError as e:
            errors.append(f"Zeile {line} (SKU {sku}): ungültige Zahl '{e}'.")
            continue
        if stock != stock.to_integral_value():
            errors.append(f"Zeile {line} (SKU {sku}): ungültiger Bestand '{val['stock']}'.")
            continue
        if stock < 0:
            # negativer Bestand (z.B. überverkauft) zählt als 0
            negative.append(f"{sku} ({int(stock)})")
            stock = Decimal(0)
        if price <= 0 or uvp <= 0:
            errors.append(f"Zeile {line} (SKU {sku}): Preis <= 0.")
            continue
        if price > uvp:
            warnings.append(f"Zeile {line} (SKU {sku}): UNIT PRICE {price} ist höher als UVP {uvp}.")

        # Artikelnummer und Größe trennen
        if sep in sku:
            base, size = sku.rsplit(sep, 1)
            if size not in known_sizes:
                errors.append(
                    f"Zeile {line}: unbekannte Größe '{size}' in SKU '{sku}'. "
                    "Bitte in mapping.json unter sizes.extra_sizes ergänzen."
                )
                continue
            name = re.sub(r"\s*-?\s*" + re.escape(size) + r"$", "", val["name"]).strip()
            if name == val["name"]:
                warnings.append(f"Zeile {line}: Name '{val['name']}' endet nicht auf Größe '{size}'.")
        else:
            base, size, name = sku, no_size, val["name"]

        art = articles.get(base)
        if art is None:
            art = articles[base] = {
                "sku": base, "name": name, "uvp": uvp, "unit_price": price,
                "stock": OrderedDict(), "skus": [], "first_line": line,
                "image": None, "sold": 0,
            }
        else:
            for key, v in (("name", name), ("uvp", uvp), ("unit_price", price)):
                if art[key] != v:
                    errors.append(
                        f"Zeile {line}: Artikel {base} hat abweichende Angabe bei '{key}' "
                        f"('{v}' statt '{art[key]}' aus Zeile {art['first_line']})."
                    )
        if size in art["stock"]:
            errors.append(f"Zeile {line}: Artikel {base} hat Größe '{size}' mehrfach.")
        art["stock"][size] = int(stock)
        art["skus"].append(sku)

    if negative:
        warnings.append(f"{len(negative)} Größe(n) mit negativem Bestand auf 0 gesetzt: " + ", ".join(negative))
    if errors:
        shown = errors[:30]
        more = f"\n  ... und {len(errors) - 30} weitere" if len(errors) > 30 else ""
        raise StocklistError("Datenprüfung fehlgeschlagen:\n  - " + "\n  - ".join(shown) + more)
    if not articles:
        raise StocklistError("Die CSV enthält keine Artikel.")

    result = list(articles.values())
    if cfg["rules"]["skip_articles_without_stock"]:
        skipped = [a for a in result if sum(a["stock"].values()) == 0]
        result = [a for a in result if sum(a["stock"].values()) > 0]
        if skipped:
            warnings.append(
                f"{len(skipped)} Artikel ohne Bestand weggelassen: "
                + ", ".join(f"{a['sku']} ({a['name']})" for a in skipped)
            )
    return result


def sort_articles(articles, cfg):
    """Sortierung (stabil: bei Gleichstand bleibt die CSV-Reihenfolge)."""
    mode = cfg["rules"]["sort"]
    if mode == "stock_desc":
        articles.sort(key=lambda a: -sum(a["stock"].values()))
    elif mode == "bestseller":
        articles.sort(key=lambda a: (-a["sold"], -sum(a["stock"].values())))
    elif mode != "csv":
        raise StocklistError(f"Unbekannte Sortierung '{mode}' in mapping.json (stock_desc, bestseller, csv).")


def load_shops(cfg):
    try:
        import shopify_client as sc
    except ImportError:
        return []
    return sc.load_shops(BASE_DIR / "config" / "shopify.env", cfg["excel"]["default_brand"])


def select_shop(shops, wanted):
    """Shop/Marke für diesen Lauf: per --shop, bei mehreren Shops sonst per Auswahl im Fenster."""
    if wanted:
        key = normalize_name(wanted)
        for s in shops:
            if key in (normalize_name(s["name"]), normalize_name(s["shop"]), normalize_name(s["shop"].split(".")[0])):
                return s
        raise StocklistError(f"Shop '{wanted}' nicht in config/shopify.env. Vorhanden: "
                             + ", ".join(s["name"] for s in shops))
    if len(shops) <= 1:
        return shops[0] if shops else None
    if not sys.stdin.isatty():
        raise StocklistError("Mehrere Shops in config/shopify.env – bitte mit --shop <Marke> wählen: "
                             + ", ".join(s["name"] for s in shops))
    print("Für welche Marke soll die Stockliste erstellt werden?")
    for i, s in enumerate(shops, 1):
        print(f"  {i}  {s['name']}")
    while True:
        answer = input("Nummer eingeben und Enter drücken: ").strip()
        if answer.isdigit() and 1 <= int(answer) <= len(shops):
            print()
            return shops[int(answer) - 1]
        print(f"Bitte eine Zahl von 1 bis {len(shops)} eingeben.")


def normalize_name(name):
    """Name für den Abgleich mit Shopify: ohne Akzente, Satzzeichen und Groß-/Kleinschreibung."""
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", name).strip()


def pick_image(candidates, name):
    """Bild-URL aus mehreren Shopify-Treffern: gleicher Name vor anderem Namen, aktiv vor archiviert."""
    if not candidates:
        return None
    key = normalize_name(name)
    best = min(candidates, key=lambda v: (normalize_name(v["title"]) != key, v["status"] != "ACTIVE"))
    return best["url"]


def enrich_from_shopify(articles, cfg, warnings, shop):
    """Bilder und Verkaufszahlen aus Shopify. Fehler hier sind nie kritisch."""
    scfg = cfg.get("shopify", {})
    if not scfg.get("enabled") or shop is None or not (shop["client_id"] and shop["client_secret"]):
        if scfg.get("enabled"):
            warnings.append("Keine Shopify-Zugangsdaten (config/shopify.env) – Liste ohne Bilder, sortiert nach Bestand.")
        if cfg["rules"]["sort"] == "bestseller":
            cfg["rules"]["sort"] = "stock_desc"
        return
    try:
        import shopify_client as sc
    except ImportError as e:
        warnings.append(f"Shopify-Modul nicht ladbar ({e}) – Liste ohne Bilder/Bestseller.")
        cfg["rules"]["sort"] = "stock_desc" if cfg["rules"]["sort"] == "bestseller" else cfg["rules"]["sort"]
        return
    client = sc.ShopifyClient(shop, scfg["api_version"])

    if scfg.get("images"):
        try:
            variants = client.image_variants()
            by_sku, by_name = {}, {}
            for v in variants:
                by_sku.setdefault(v["sku"], []).append(v)
                by_name.setdefault(normalize_name(v["title"]), []).append(v)
            cache = BASE_DIR / scfg["image_cache_dir"]
            missing, via_name = [], []
            for art in articles:
                # 1. SKU mit Größe, 2. SKU ohne Größe (so sind viele Artikel in Shopify angelegt)
                cands = [v for s in art["skus"] + [art["sku"]] for v in by_sku.get(s, [])]
                if not cands and scfg.get("match_by_name", True):
                    # 3. Produktname
                    cands = by_name.get(normalize_name(art["name"]), [])
                    if cands:
                        via_name.append(f"{art['sku']} ({art['name']})")
                url = pick_image(cands, art["name"])
                if url is None:
                    missing.append(art["sku"])
                    continue
                try:
                    art["image"] = sc.thumbnail(url, cache, art["sku"], scfg["image_max_px"])
                except (sc.ShopifyError, OSError) as e:
                    missing.append(art["sku"])
                    warnings.append(f"Bild für {art['sku']} nicht ladbar: {e}")
            if via_name:
                warnings.append(f"{len(via_name)} Artikel-Bild(er) über den Produktnamen zugeordnet "
                                "(SKU nicht in Shopify): " + ", ".join(via_name))
            if missing:
                more = f" … (+{len(missing) - 10} weitere)" if len(missing) > 10 else ""
                warnings.append(f"{len(missing)} Artikel ohne Shopify-Bild: " + ", ".join(missing[:10]) + more)
        except sc.ShopifyError as e:
            warnings.append(f"Shopify-Bilder nicht abrufbar: {e}")

    if cfg["rules"]["sort"] == "bestseller":
        try:
            sold = client.units_sold_by_sku(scfg["bestseller_days"])
            for art in articles:
                # auch Verkäufe unter der SKU ohne Größe zählen
                art["sold"] = sum(sold.get(s, 0) for s in set(art["skus"]) | {art["sku"]})
        except sc.ShopifyError as e:
            warnings.append(f"Shopify-Verkaufszahlen nicht abrufbar ({e}) – sortiere nach Bestand.")
            cfg["rules"]["sort"] = "stock_desc"


def add_image(ws, path, row, col, box_w, box_h):
    """Bild zentriert in die (verbundene) Zelle setzen."""
    img = XLImage(str(path))
    scale = min(box_w / img.width, box_h / img.height, 1)
    w, h = int(img.width * scale), int(img.height * scale)
    img.width, img.height = w, h
    marker = AnchorMarker(col=col - 1, colOff=pixels_to_EMU((box_w - w) // 2 + 4),
                          row=row - 1, rowOff=pixels_to_EMU((box_h - h) // 2 + 4))
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
    ws.add_image(img)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def read_template_layout(ws, cfg):
    """Liest Spaltenrollen, Stile, Breiten und Höhen aus dem Template."""
    xc = cfg["excel"]
    hr, r1 = xc["header_row"], xc["first_data_row"]
    r2 = r1 + 1
    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(hr, col).value
        if v is not None:
            headers[str(v).strip()] = col

    col_of = {}
    for role, title in xc["column_headers"].items():
        if title not in headers:
            raise StocklistError(f"Template: Überschrift '{title}' nicht in Zeile {hr} gefunden.")
        col_of[role] = headers[title]
    first_size, last_size = col_of["uvp"] + 1, col_of["quantity"] - 1
    sizes = [str(ws.cell(hr, c).value).strip() for c in range(first_size, last_size + 1)]
    if cfg["sizes"]["no_size_column"] not in sizes:
        raise StocklistError(f"Template: Größenspalte '{cfg['sizes']['no_size_column']}' fehlt.")
    fixed = sorted(col_of[r] for r in ("image", "sku", "name", "unit_price", "uvp"))
    if fixed != list(range(1, len(fixed) + 1)) or col_of["total"] != col_of["quantity"] + 1:
        raise StocklistError("Template: unerwartete Spaltenanordnung.")

    def width(col):
        letter = get_column_letter(col)
        for dim in ws.column_dimensions.values():
            if dim.min and dim.max and dim.min <= col <= dim.max:
                return dim.width
        return ws.column_dimensions[letter].width

    def cell_info(row, col):
        c = ws.cell(row, col)
        return {"style": c._style, "value": c.value}

    roles = {}
    for role in ("image", "sku", "name", "unit_price", "uvp", "quantity", "total"):
        col = col_of[role]
        roles[role] = {
            "header": cell_info(hr, col), "top": cell_info(r1, col),
            "bottom": cell_info(r2, col), "width": width(col),
        }
    roles["size"] = {
        "header": cell_info(hr, first_size), "top": cell_info(r1, first_size),
        "bottom": cell_info(r2, first_size), "width": width(first_size),
    }
    return {
        "sizes": sizes,
        "fixed_order": [r for r in ("image", "sku", "name", "unit_price", "uvp")],
        "fixed_cols": {r: col_of[r] for r in ("image", "sku", "name", "unit_price", "uvp")},
        "roles": roles,
        "row_heights": {
            "title": ws.row_dimensions[1].height,
            "header": ws.row_dimensions[hr].height,
            "top": ws.row_dimensions[r1].height,
            "bottom": ws.row_dimensions[r2].height,
        },
        "title": ws.cell(1, 1).value,
    }


def final_size_list(template_sizes, articles, cfg):
    used = {s for a in articles for s in a["stock"]}
    sizes = list(template_sizes)
    for extra in cfg["sizes"]["extra_sizes"]:
        if extra["size"] in used and extra["size"] not in sizes:
            sizes.insert(sizes.index(extra["insert_after"]) + 1, extra["size"])
    return sizes


def apply(cell, info, value=None, keep_template_value=False):
    cell._style = info["style"]
    if keep_template_value:
        cell.value = info["value"]
    elif value is not None:
        cell.value = value


def write_workbook(template_path, out_path, articles, cfg, today, brand):
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]
    xc = cfg["excel"]
    layout = read_template_layout(ws, cfg)
    sizes = final_size_list(layout["sizes"], articles, cfg)
    roles = layout["roles"]
    hr, r_first = xc["header_row"], xc["first_data_row"]

    n_fixed = len(layout["fixed_order"])
    size_col = {s: n_fixed + 1 + i for i, s in enumerate(sizes)}
    q_col = n_fixed + len(sizes) + 1
    t_col = q_col + 1

    # Datenbereich ab Kopfzeile leeren (Titel in Zeile 1 bleibt erhalten)
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.delete_rows(hr, ws.max_row)
    for key in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[key]
    for r in [r for r in ws.row_dimensions if r >= hr]:
        del ws.row_dimensions[r]

    # Titel + Blattname
    date_str = today.strftime(xc["date_format"])
    if xc.get("title"):
        ws.cell(1, 1).value = xc["title"].format(BRAND=brand.upper(), brand=brand, date=date_str)
    elif layout["title"]:
        ws.cell(1, 1).value = str(layout["title"]).replace(xc["date_placeholder"], date_str)
    sheet_name = xc["sheet_name"].format(brand=brand, date=date_str)
    if len(sheet_name) > 31:
        # Excel erlaubt max. 31 Zeichen: Markenname kürzen
        cut = len(sheet_name) - 31
        sheet_name = xc["sheet_name"].format(brand=brand[:max(1, len(brand) - cut)].strip(), date=date_str)
    ws.title = sheet_name[:31]

    # Spaltenbreiten
    for role, col in layout["fixed_cols"].items():
        ws.column_dimensions[get_column_letter(col)].width = roles[role]["width"]
    for col in size_col.values():
        ws.column_dimensions[get_column_letter(col)].width = roles["size"]["width"]
    ws.column_dimensions[get_column_letter(q_col)].width = roles["quantity"]["width"]
    ws.column_dimensions[get_column_letter(t_col)].width = roles["total"]["width"]

    # Kopfzeile
    ws.row_dimensions[hr].height = layout["row_heights"]["header"]
    for role, col in layout["fixed_cols"].items():
        apply(ws.cell(hr, col), roles[role]["header"], keep_template_value=True)
    for s, col in size_col.items():
        apply(ws.cell(hr, col), roles["size"]["header"], s)
    apply(ws.cell(hr, q_col), roles["quantity"]["header"], keep_template_value=True)
    apply(ws.cell(hr, t_col), roles["total"]["header"], keep_template_value=True)

    first_size_letter = get_column_letter(size_col[sizes[0]])
    last_size_letter = get_column_letter(size_col[sizes[-1]])
    q_letter, t_letter = get_column_letter(q_col), get_column_letter(t_col)
    price_letter = get_column_letter(layout["fixed_cols"]["unit_price"])

    row = r_first
    for art in articles:
        top, bottom = row, row + 1
        ws.row_dimensions[top].height = layout["row_heights"]["top"]
        ws.row_dimensions[bottom].height = layout["row_heights"]["bottom"]

        values = {"image": None, "sku": art["sku"], "name": art["name"],
                  "unit_price": float(art["unit_price"]), "uvp": float(art["uvp"])}
        for role, col in layout["fixed_cols"].items():
            apply(ws.cell(top, col), roles[role]["top"], values[role])
            apply(ws.cell(bottom, col), roles[role]["bottom"])
        for s, col in size_col.items():
            qty = art["stock"].get(s)
            if qty == 0 and not cfg["rules"]["show_zero_stock"]:
                qty = None
            apply(ws.cell(top, col), roles["size"]["top"], qty)
            apply(ws.cell(bottom, col), roles["size"]["bottom"])
        apply(ws.cell(top, q_col), roles["quantity"]["top"], keep_template_value=True)
        apply(ws.cell(top, t_col), roles["total"]["top"], keep_template_value=True)
        apply(ws.cell(bottom, q_col), roles["quantity"]["bottom"])
        apply(ws.cell(bottom, t_col), roles["total"]["bottom"])
        if cfg["rules"]["order_formulas"]:
            ws.cell(bottom, q_col).value = (
                f'=IF(SUM({first_size_letter}{bottom}:{last_size_letter}{bottom})=0,"",'
                f"SUM({first_size_letter}{bottom}:{last_size_letter}{bottom}))"
            )
            ws.cell(bottom, t_col).value = f'=IF({q_letter}{bottom}="","",{q_letter}{bottom}*{price_letter}{top})'

        for col in layout["fixed_cols"].values():
            ws.merge_cells(start_row=top, start_column=col, end_row=bottom, end_column=col)
        if art.get("image"):
            # Platz = Spaltenbreite x Höhe beider Zeilen, abzüglich Rand
            box_w = int(roles["image"]["width"] * 8 + 5) - 8  # ca. 8 px je Zeichen bei Aptos Narrow 12
            box_h = int((layout["row_heights"]["top"] + layout["row_heights"]["bottom"]) * 4 / 3) - 8
            add_image(ws, art["image"], top, layout["fixed_cols"]["image"], box_w, box_h)
        row += 2

    last_row = row - 1
    if cfg["rules"]["grand_total_row"] and cfg["rules"]["order_formulas"]:
        ws.row_dimensions[row].height = layout["row_heights"]["bottom"]
        apply(ws.cell(row, q_col - 1), roles["quantity"]["bottom"], xc["grand_total_label"])
        apply(ws.cell(row, q_col), roles["quantity"]["bottom"],
              f"=SUM({q_letter}{r_first}:{q_letter}{last_row})")
        apply(ws.cell(row, t_col), roles["total"]["bottom"],
              f"=SUM({t_letter}{r_first}:{t_letter}{last_row})")

    ws.freeze_panes = ws.cell(r_first, n_fixed + 1).coordinate
    if cfg["rules"].get("hide_unused_sizes"):
        used = {s for a in articles for s in a["stock"]}
        for s, col in size_col.items():
            if s not in used:
                ws.column_dimensions[get_column_letter(col)].hidden = True

    if cfg["rules"]["landscape_print"]:
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "landscape"
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.5
        ws.page_margins.header = ws.page_margins.footer = 0.3
        ws.print_options.horizontalCentered = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = f"1:{hr}"

    wb.save(out_path)
    return sizes


def output_path(cfg, today, brand):
    out_dir = BASE_DIR / cfg["paths"]["output_dir"]
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_brand = re.sub(r"[^A-Za-z0-9]+", "-", normalize_name(brand).title()).strip("-")
    name = cfg["paths"]["output_filename"].format(brand=safe_brand, date_iso=today.isoformat())
    path = out_dir / name
    version = 2
    while path.exists():
        path = out_dir / f"{Path(name).stem}_v{version}{Path(name).suffix}"
        version += 1
    return path


def archive_csv(csv_path, cfg, today):
    archive_dir = BASE_DIR / cfg["paths"]["archive_dir"]
    archive_dir.mkdir(parents=True, exist_ok=True)
    target = archive_dir / f"{today.isoformat()}_{csv_path.name}"
    version = 2
    while target.exists():
        target = archive_dir / f"{today.isoformat()}_v{version}_{csv_path.name}"
        version += 1
    shutil.move(str(csv_path), target)
    return target


# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Excel-Stockliste aus CSV erzeugen")
    parser.add_argument("--csv", help="bestimmte CSV-Datei verwenden")
    parser.add_argument("--no-archive", action="store_true",
                        help="CSV nach Erfolg nicht nach input/archiv verschieben")
    parser.add_argument("--shop", help="Marke/Shop aus config/shopify.env (bei mehreren Shops)")
    args = parser.parse_args()

    cfg = load_config()
    today = date.today()
    warnings = []
    try:
        shop = select_shop(load_shops(cfg), args.shop)
        brand = shop["name"] if shop else cfg["excel"]["default_brand"]
        print(f"Marke:    {brand}" + (f" ({shop['shop']})" if shop else ""))
        csv_path = find_input_csv(cfg, args.csv)
        print(f"CSV:      {csv_path.name}")
        raw, enc = read_csv(csv_path, cfg)
        print(f"Encoding: {enc}, {len(raw) - 1} Datenzeilen")
        positions = check_header(raw, cfg, warnings)

        template_path = BASE_DIR / cfg["paths"]["template"]
        if not template_path.is_file():
            raise StocklistError(f"Template nicht gefunden: {template_path}")
        template_sizes = read_template_layout(load_workbook(template_path).worksheets[0], cfg)["sizes"]

        articles = build_articles(raw, positions, cfg, template_sizes, warnings)
        enrich_from_shopify(articles, cfg, warnings, shop)
        sort_articles(articles, cfg)
        out = output_path(cfg, today, brand)
        sizes = write_workbook(template_path, out, articles, cfg, today, brand)
    except (StocklistError, OSError) as e:
        print("\nFEHLER – es wurde KEINE Stockliste erstellt.\n")
        print(e)
        for w in warnings:
            print(f"Hinweis: {w}")
        return 1

    for w in warnings:
        print(f"Hinweis: {w}")
    extra = [s for s in sizes if s not in template_sizes]
    if extra:
        print(f"Hinweis: zusätzliche Größenspalten eingefügt: {', '.join(extra)}")
    total = sum(sum(a["stock"].values()) for a in articles)
    print(f"\nFertig: {out.relative_to(BASE_DIR)}  ({len(articles)} Artikel, {total} Teile)")
    if not args.no_archive and not args.csv:
        moved = archive_csv(csv_path, cfg, today)
        print(f"CSV archiviert: {moved.relative_to(BASE_DIR)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
