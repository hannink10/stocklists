"""Interne Auswertungen (Menüpunkt "Intern").

Sale-Empfehlung: liest eine Artikel-CSV aus JTL (Artikelnummer, Name, EK netto, UVP/VK, Bestand), holt aus Shopify
Verkäufe, Retouren, bezahlte Preise, Lagerbestand je Größe und das Shop-Datum und empfiehlt je Artikel einen Rabatt:
nach Abverkauf, Reichweite, Trend, Restgrößen und Retouren – aber nie unter die Mindestmarge (VK netto ÷ EK).
Ergebnis: eine Excel-Datei, sortiert nach Rohertrag-Potenzial.
"""

import math
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import generate_stocklist as g

DARK = "FF2F2F2F"
THIN = Side(style="thin", color="FFD9D9D9")
STEP_COLORS = {0: "FFF2F2F2", 20: "FFFFF2CC", 30: "FFFFE699", 40: "FFF8CBAD", 50: "FFF4B183", 70: "FFE06666"}


def run(mode, args, cfg, shop, brand, today):
    """Einstieg aus generate_stocklist.main(). Liefert den Exit-Code."""
    warnings = []
    try:
        if mode != "sale":
            raise g.StocklistError(f"Unbekannte Auswertung '{mode}'.")
        scfg = cfg["internal"]["sale"]
        print(f"Marke:    {brand}" + (f" ({shop['shop']})" if shop else ""))
        print("Auswertung: Sale-Empfehlung")
        if not shop:
            raise g.StocklistError("Für die Sale-Empfehlung werden Shopify-Daten gebraucht (config/shopify.env).")
        csv_path = g.find_input_csv(cfg, args.csv)
        print(f"CSV:      {csv_path.name}")
        raw, enc = g.read_csv(csv_path, cfg)
        header_idx, positions, col_info = g.detect_columns(raw, cfg, brand, warnings, scfg["columns"])
        print(f"Encoding: {enc}, {len(raw) - header_idx - 1} Datenzeilen")
        print("Spalten:  " + "\n          ".join(col_info))
        g.apply_brand_groups(cfg, brand)

        articles = read_articles(raw, header_idx, positions, cfg, warnings)
        excluded = ask_exclusions(args)
        if excluded:
            kept = [a for a in articles if not any(a["sku"].upper().startswith(p.upper()) for p in excluded)]
            print(f"Ausgeschlossen: {len(articles) - len(kept)} Artikel mit Nummer beginnend mit {', '.join(excluded)}")
            articles = kept
            if not articles:
                raise g.StocklistError("Nach dem Ausschluss sind keine Artikel mehr übrig.")
        scfg["_excluded"] = excluded
        print(f"Artikel:  {len(articles)} – lade Shopify-Daten (Verkäufe, Retouren, Größen) …")
        coverage = add_shopify_data(articles, cfg, shop, brand, scfg, today, warnings)
        recommend(articles, scfg, coverage, today)
        out = g.output_path(cfg, today, brand, scfg["output_filename"])
        write_report(out, articles, scfg, coverage, brand, today, cfg.get("product_groups"))
    except (g.StocklistError, OSError) as e:
        print("\nFEHLER – es wurde KEINE Auswertung erstellt.\n")
        print(e)
        for w in warnings:
            print(f"Hinweis: {w}")
        return 1

    for w in warnings:
        print(f"Hinweis: {w}")
    rated = [a for a in articles if a["discount"] is not None]
    on_sale = [a for a in rated if a["discount"] > 0]
    print(f"\nFertig: {out.relative_to(g.BASE_DIR)}")
    below = [a for a in rated if a.get("below_margin")]
    if below:
        print(f"  ⚠ {len(below)} Artikel mit {scfg['steps'][0]} % unter der Mindestmarge (rot markiert)")
    print(f"  {len(on_sale)} von {len(articles)} Artikeln mit Rabatt-Empfehlung, "
          f"Rohertrag bei Abverkauf {eur(sum(a['potential'] for a in on_sale))}")
    if not args.no_archive and not args.csv:
        moved = g.archive_csv(csv_path, cfg, today)
        print(f"CSV archiviert: {moved.relative_to(g.BASE_DIR)}")
    return 0


def ask_exclusions(args):
    """Artikelnummer-Anfänge, die nicht berücksichtigt werden (z.B. '100826 & 100726'): per --exclude oder Abfrage."""
    import re
    import sys
    text = getattr(args, "exclude", None)
    if text is None:
        if not sys.stdin.isatty():
            return []
        print("Artikel ausschließen? Artikelnummer-Anfänge eingeben (z.B. 100826 & 100726), Enter = keine:")
        text = input("  > ")
        print()
    return [t for t in re.findall(r"[A-Za-z0-9][A-Za-z0-9._-]*", text) if t.lower() not in ("und", "and")]


# ---------------------------------------------------------------------------
# Daten

def read_articles(raw, header_idx, positions, cfg, warnings):
    """Eine Zeile je Artikel (ohne Größen)."""
    articles, missing_ek, bad = [], [], []
    for idx in range(header_idx + 1, len(raw)):
        row = raw.iloc[idx]
        val = {k: str(row.iloc[p]).strip() for k, p in positions.items()}
        if not val.get("sku"):
            continue
        try:
            ek = g.parse_number(val.get("ek", ""), cfg) or g.Decimal(0)
            uvp = g.parse_number(val.get("uvp", ""), cfg) or g.Decimal(0)
            vk = g.parse_number(val.get("vk", ""), cfg) if "vk" in val else None
            stock = g.parse_number(val.get("stock", ""), cfg) or g.Decimal(0)
        except ValueError as e:
            bad.append(f"Zeile {idx + 1} ({val['sku']}): '{e}'")
            continue
        vk = vk or uvp
        if ek <= 0:
            missing_ek.append(val["sku"])
        articles.append({
            "sku": val["sku"], "name": val.get("name", ""), "ek": g.cents(ek), "uvp": g.cents(uvp),
            "vk": g.cents(vk), "stock": max(0, int(stock)), "line": idx + 1,
        })
    if bad:
        warnings.append(f"{len(bad)} Zeile(n) mit ungültigen Zahlen übersprungen: " + "; ".join(bad[:5]))
    if missing_ek:
        warnings.append(f"{len(missing_ek)} Artikel ohne EK – Marge nicht prüfbar, keine Empfehlung: "
                        + ", ".join(missing_ek[:10]) + (" …" if len(missing_ek) > 10 else ""))
    if not articles:
        raise g.StocklistError("Die CSV enthält keine Artikel.")
    return articles


def add_shopify_data(articles, cfg, shop, brand, scfg, today, warnings):
    """Verkäufe, Retouren, bezahlte Preise, Größen und Shop-Datum je Artikel aus Shopify."""
    import shopify_client as sc

    client = sc.ShopifyClient(shop, cfg["shopify"]["api_version"])
    try:
        orders = client.order_lines(scfg["days"])
        variants = client.variant_details()
    except sc.ShopifyError as e:
        raise g.StocklistError(f"Shopify-Daten nicht abrufbar: {e}")

    bases = {a["sku"] for a in articles}
    rule = g.brand_config(cfg, brand).get("sku_rule") or {"mode": "separator", "separator": cfg["sizes"]["sku_separator"]}
    template = g.load_workbook(g.BASE_DIR / cfg["paths"]["template"]).worksheets[0]
    known = set(g.read_template_layout(template, cfg)["sizes"]) | {e["size"] for e in cfg["sizes"]["extra_sizes"]}

    def base_of(sku):
        if sku in bases:
            return sku
        base, size, _ = g.split_sku(sku, rule, known)
        if base in bases:
            return base
        # Fallback: längste Artikelnummer, mit der die Varianten-SKU beginnt (z.B. ohne bekannte Größe)
        for n in range(len(sku) - 1, 2, -1):
            if sku[:n] in bases and len(sku) - n <= 6:
                return sku[:n]
        return None

    per = defaultdict(lambda: {"sold": 0, "returned": 0, "revenue": 0.0, "last": None, "recent": 0, "before": 0, "sold_recent": 0,
                               "sizes": 0, "sizes_stock": 0, "since": None, "on_sale": False, "found": False})
    oldest = None
    for line in orders["lines"]:
        b = base_of(line["sku"])
        oldest = line["date"] if oldest is None or line["date"] < oldest else oldest
        if not b:
            continue
        p = per[b]
        p["sold"] += line["qty"]
        p["revenue"] += line["paid"]
        p["last"] = line["date"] if p["last"] is None or line["date"] > p["last"] else p["last"]
        age = (today - line["date"]).days
        if age < scfg["recent_days"]:
            p["sold_recent"] += line["qty"]
        if age < scfg["trend_days"]:
            p["recent"] += line["qty"]
        elif age < 2 * scfg["trend_days"]:
            p["before"] += line["qty"]
    for sku, qty in orders["returned"].items():
        b = base_of(sku)
        if b:
            per[b]["returned"] += qty
    for v in variants:
        b = base_of(v["sku"])
        if not b:
            continue
        p = per[b]
        p["found"] = True
        p["sizes"] += 1
        p["sizes_stock"] += 1 if v["stock"] > 0 else 0
        p["on_sale"] |= bool(v["compare_at"] and v["compare_at"] > v["price"])
        if v["since"] and (p["since"] is None or v["since"] < p["since"]):
            p["since"] = v["since"]

    days = max(1, min(scfg["days"], (today - oldest).days if oldest else 1))
    not_found = []
    for a in articles:
        p = per[a["sku"]]
        a.update(p)
        if not p["found"]:
            not_found.append(a["sku"])
    if not_found:
        warnings.append(f"{len(not_found)} Artikel nicht in Shopify gefunden (keine Verkaufsdaten, keine Empfehlung): "
                        + ", ".join(not_found[:10]) + (" …" if len(not_found) > 10 else ""))
    full_year = days >= scfg["full_period_min_days"]
    if not full_year:
        warnings.append(f"Shopify liefert nur {days} Tage Bestellungen (ältere brauchen das App-Recht read_all_orders) – "
                        f"die Korrektur „übers Jahr gefragt“ entfällt.")
    total_sold = sum(a["sold"] for a in articles)
    total_ret = sum(a["returned"] for a in articles)
    return {"days": days, "oldest": oldest, "full_year": full_year, "sold": total_sold, "returned": total_ret}


# ---------------------------------------------------------------------------
# Empfehlung

def recommend(articles, scfg, coverage, today):
    """Ohne Rabatt bleiben nur Bestseller (laufen JETZT: Top-Verkäufe der letzten recent_days) und Neuheiten;
    alle anderen bekommen mindestens die kleinste Stufe. Die Stufe richtet sich nach dem aktuellen Abverkauf,
    das Jahr korrigiert nur (übers Jahr gefragt -> eine Stufe weniger)."""
    steps = sorted(scfg["steps"])
    tiers = scfg["tiers_recent"]
    vat = scfg["vat"]
    # Bestseller = läuft JETZT: Top x % nach verkauften Teilen der letzten recent_days (mind. bestseller_min_units)
    for a in articles:
        rr = a["returned"] / a["sold"] if a["sold"] else 0
        a["recent_net"] = round(a["sold_recent"] * (1 - rr))
    ranked = sorted((a["recent_net"] for a in articles if a.get("found")), reverse=True)
    top_n = max(1, int(len(ranked) * scfg["bestseller_top_pct"])) if ranked else 0
    best_limit = max(scfg["bestseller_min_units"], ranked[top_n - 1] if top_n else 0)
    scfg["_best_limit"] = best_limit
    weeks = coverage["days"] / 7

    def step_up(d, n=1):
        bigger = [s for s in steps if s > d]
        return bigger[min(n, len(bigger)) - 1] if bigger else d

    def step_down(d):
        # nie unter die kleinste Stufe: ohne Rabatt bleiben nur Bestseller und Neuheiten
        smaller = [s for s in steps if s < d]
        return smaller[-1] if smaller else steps[0]

    for a in articles:
        a["net_sold"] = max(0, a["sold"] - a["returned"])
        a["return_rate"] = a["returned"] / a["sold"] if a["sold"] else None
        a["str"] = a["net_sold"] / (a["net_sold"] + a["stock"]) if a["net_sold"] + a["stock"] else None
        a["str_recent"] = a["recent_net"] / (a["recent_net"] + a["stock"]) if a["recent_net"] + a["stock"] else None
        weekly = a["net_sold"] / weeks
        a["reach"] = a["stock"] / weekly if weekly else math.inf
        a["age"] = (today - a["since"]).days if a["since"] else None
        a["last_days"] = (today - a["last"]).days if a["last"] else None
        a["avg_paid"] = a["revenue"] / a["sold"] if a["sold"] else None
        a["trend"] = (a["recent"] / a["before"]) if a["before"] else (math.inf if a["recent"] else None)
        a["max_disc"] = (1 - scfg["min_margin"] * float(a["ek"]) * (1 + vat) / float(a["vk"])) if a["ek"] and a["vk"] else None
        reasons = []

        if not a["ek"] or not a["vk"]:
            a["discount"], a["reason"] = None, "kein EK/VK – Marge nicht prüfbar"
            continue
        if not a["found"]:
            a["discount"], a["reason"] = None, "nicht in Shopify gefunden"
            continue
        if a["age"] is not None and a["age"] < scfg["new_days"]:
            a["discount"], a["reason"] = 0, f"neu im Shop ({a['age']} Tage)"
            finish(a, scfg)
            continue

        days = scfg["recent_days"]
        # 1) Bestseller: läuft gerade -> kein Rabatt (außer bei massivem Überbestand)
        sells_well = a["recent_net"] >= best_limit
        if sells_well and (a["str_recent"] or 0) >= scfg["bestseller_min_str_recent"]:
            a["discount"] = 0
            a["reason"] = f"Bestseller: {a['recent_net']} verkauft in {days} Tagen (Top {scfg['bestseller_top_pct']:.0%}, ab {best_limit})"
            finish(a, scfg)
            continue
        # 2) Basis: Abverkauf der letzten recent_days
        desired = next(t["discount"] for t in tiers if (a["str_recent"] or 0) >= t["min_str"])
        reasons.append(f"{a['recent_net']} verkauft in {days} Tagen (Abverkauf {pct(a['str_recent'])})")
        if a["net_sold"] == 0 and a["age"] is not None and a["age"] >= scfg["slow_days"]:
            desired = scfg["slow_discount"]
            reasons = [f"Ladenhüter: 0 Verkäufe im Jahr, {a['age']} Tage im Shop"]
        elif sells_well:
            desired = step_down(desired)
            reasons.append(f"läuft gut, aber großer Bestand ({a['stock']})")
        # 3) Übers Jahr gefragt, läuft aber gerade nicht -> eine Stufe weniger (Nachfrage ist grundsätzlich da)
        elif coverage["full_year"] and (a["str"] or 0) >= scfg["year_popular_str"]:
            desired = step_down(desired)
            reasons.append(f"übers Jahr gefragt (Abverkauf 12 Mon. {pct(a['str'])})")
        # 4) Trend: zieht gerade an -> eine Stufe weniger
        if desired and a["trend"] and a["recent"] >= scfg["trend_min_units"] and a["trend"] >= scfg["trend_up"]:
            desired = step_down(desired)
            reasons.append("Trend steigend")
        # 5) Restgrößen: nur noch wenige Größen da -> eine Stufe mehr
        if desired and a["sizes"] >= 4 and a["sizes_stock"] <= scfg["broken_sizes_max"]:
            desired = step_up(desired)
            reasons.append(f"nur {a['sizes_stock']} von {a['sizes']} Größen")
        # 6) Hohe Retourenquote -> nicht tiefer gehen
        if desired and a["return_rate"] is not None and a["sold"] >= 5 and a["return_rate"] >= scfg["high_return_rate"]:
            desired = step_down(desired)
            reasons.append(f"Retourenquote {pct(a['return_rate'])}")
        # 7) Margen-Sperre
        allowed = [s for s in steps if s <= desired and s / 100 <= a["max_disc"] + 1e-9]
        chosen = max(allowed) if allowed else 0
        a["below_margin"] = False
        if chosen < desired:
            limit = f"max. {pct(max(0, a['max_disc']))}"
            if not chosen and scfg.get("force_min_step", True):
                # Sale des Jahres: trotzdem die kleinste Stufe, aber deutlich markieren
                chosen, a["below_margin"] = steps[0], True
                reasons = [f"⚠ UNTER MINDESTMARGE – Marge erlaubt nur {limit}"] + reasons
            elif not chosen:
                reasons = [f"Marge lässt keinen Rabatt zu ({limit})"] + reasons
            else:
                reasons.append(f"durch Marge {str(scfg['min_margin']).replace('.', ',')} begrenzt ({limit})")
        a["discount"] = chosen
        a["reason"] = ", ".join(reasons)
        finish(a, scfg)


def finish(a, scfg):
    vat = scfg["vat"]
    d = a["discount"] or 0
    a["sale_price"] = float(g.cents(a["vk"] * (1 - g.Decimal(d) / 100)))
    net = a["sale_price"] / (1 + vat)
    a["factor"] = net / float(a["ek"]) if a["ek"] else None
    rr = a["return_rate"] or 0
    a["profit_unit"] = net - float(a["ek"]) - rr * scfg["return_cost_eur"]
    a["potential"] = a["profit_unit"] * a["stock"]


def eur(x):
    return f"{x:,.0f} €".replace(",", ".")


def pct(x):
    return "–" if x is None else f"{x * 100:.0f} %"


# ---------------------------------------------------------------------------
# Excel

COLUMNS = [
    ("Rang", 6, None), ("SKU", 13, None), ("Artikel", 40, None), ("Kategorie", 12, None),
    ("Empf. Rabatt", 10, "0%"), ("Sale-Preis", 11, "#,##0.00 €"), ("Faktor danach", 9, "0.00"),
    ("Rohertrag/Stk", 11, "#,##0.00 €"), ("Rohertrag-Potenzial", 13, "#,##0 €"),
    ("Bestand", 8, "0"), ("Lagerwert EK", 12, "#,##0 €"), ("Verkauft 60 T.", 9, "0"), ("Abverkauf 60 T.", 9, "0%"),
    ("Verkauft 12 Mon.", 9, "0"), ("Abverkauf 12 Mon.", 9, "0%"), ("Retouren", 9, "0%"), ("Reichweite Wo.", 10, "0"), ("Trend 30 T.", 10, None), ("Letzter Verkauf (Tage)", 10, "0"),
    ("Ø bezahlt", 10, "#,##0.00 €"), ("Größen verfügbar", 10, None), ("Im Shop seit", 11, "DD.MM.YYYY"),
    ("Max. Rabatt (Marge)", 10, "0%"), ("EK netto", 10, "#,##0.00 €"), ("VK aktuell", 10, "#,##0.00 €"),
    ("Begründung", 60, None),
]


def row_values(rank, a, category):
    trend = None
    if a.get("trend") is not None:
        trend = "neu ↑" if math.isinf(a["trend"]) else f"{'↑' if a['trend'] > 1.1 else '↓' if a['trend'] < 0.9 else '→'} {a['recent']} vs {a['before']}"
    d = a.get("discount")
    return [
        rank, a["sku"], a["name"], category,
        d / 100 if d is not None else None, a.get("sale_price"), a.get("factor"), a.get("profit_unit"), a.get("potential"),
        a["stock"], float(a["ek"]) * a["stock"], a.get("recent_net"), a.get("str_recent"),
        a.get("net_sold"), a.get("str"), a.get("return_rate"),
        None if a.get("reach") is None or math.isinf(a["reach"]) else round(a["reach"]), trend, a.get("last_days"),
        a.get("avg_paid"), f"{a['sizes_stock']}/{a['sizes']}" if a.get("sizes") else None, a.get("since"),
        max(0, a["max_disc"]) if a.get("max_disc") is not None else None, float(a["ek"]), float(a["vk"]), a.get("reason"),
    ]


def write_report(out, articles, scfg, coverage, brand, today, groups_cfg=None):
    rated = [a for a in articles if a["discount"] is not None]
    unrated = [a for a in articles if a["discount"] is None]
    with_disc = sorted([a for a in rated if a["discount"] > 0], key=lambda a: -a["potential"])
    without = sorted([a for a in rated if a["discount"] == 0], key=lambda a: -(a.get("recent_net") or 0))
    ordered = with_disc + without + unrated

    wb = Workbook()
    ws = wb.active
    ws.title = "Sale-Empfehlung"
    period = f"{coverage['oldest'].strftime('%d.%m.%Y')}–{today.strftime('%d.%m.%Y')} ({coverage['days']} Tage)" if coverage["oldest"] else "–"
    ws["A1"] = f"SALE-EMPFEHLUNG {brand.upper()} – {today.strftime('%d.%m.%Y')}"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=DARK)
    excluded = scfg.get("_excluded") or []
    ws["A2"] = (f"Shopify-Daten: {period} · Mindestmarge {str(scfg['min_margin']).replace('.', ',')} (VK netto ÷ EK) · "
                + (f"ausgeschlossen: {', '.join(excluded)}… · " if excluded else "") + 
                f"Rabattstufen {', '.join(str(s) for s in scfg['steps'])} % · sortiert nach Rohertrag-Potenzial")
    ws["A2"].font = Font(name="Arial", size=10, color="FF7F7F7F")
    hr = 4
    for i, (title, width, _) in enumerate(COLUMNS, 1):
        c = ws.cell(hr, i, title)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor=DARK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[hr].height = 32

    for rank, a in enumerate(ordered, 1):
        r = hr + rank
        category = g.product_group(a["name"], groups_cfg, a["sku"]) if groups_cfg and groups_cfg.get("rules") else ""
        for i, v in enumerate(row_values(rank if a in with_disc else None, a, category), 1):
            c = ws.cell(r, i, v)
            c.font = Font(name="Arial", size=10, bold=(COLUMNS[i - 1][0] in ("Empf. Rabatt", "Sale-Preis")))
            c.border = Border(bottom=THIN)
            if COLUMNS[i - 1][2]:
                c.number_format = COLUMNS[i - 1][2]
            if i > 4 and COLUMNS[i - 1][0] != "Begründung":
                c.alignment = Alignment(horizontal="center")
        d = a["discount"]
        if d is not None:
            ws.cell(r, 5).fill = PatternFill("solid", fgColor=STEP_COLORS.get(d, "FFF8CBAD"))
        if a.get("below_margin"):
            warn = PatternFill("solid", fgColor="FFFFC7CE")
            for col in range(1, len(COLUMNS) + 1):
                ws.cell(r, col).fill = warn
            for col in (5, 7, len(COLUMNS)):
                ws.cell(r, col).font = Font(name="Arial", size=10, bold=True, color="FF9C0006")
        elif a.get("factor") is not None and a["factor"] < scfg["min_margin"] + 0.3:
            ws.cell(r, 7).font = Font(name="Arial", size=10, bold=True, color="FFC00000")
    ws.freeze_panes = ws.cell(hr + 1, 4)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(COLUMNS))}{hr + len(ordered)}"

    # Übersicht je Rabattstufe
    so = wb.create_sheet("Übersicht", 0)
    so["A1"] = f"SALE-ÜBERSICHT {brand.upper()} – {today.strftime('%d.%m.%Y')}"
    so["A1"].font = Font(name="Arial", size=16, bold=True, color=DARK)
    so["A2"] = f"Shopify-Daten: {period} · Retourenquote gesamt {pct(coverage['returned'] / coverage['sold']) if coverage['sold'] else '–'}"
    so["A2"].font = Font(name="Arial", size=10, color="FF7F7F7F")
    heads = ["Rabatt", "Artikel", "Bestand", "Lagerwert EK", "Umsatz netto bei Abverkauf", "Rohertrag bei Abverkauf", "Ø Faktor"]
    widths = [14, 10, 10, 14, 20, 20, 10]
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        c = so.cell(4, i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor=DARK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        so.column_dimensions[get_column_letter(i)].width = w
    r = 5
    totals = [0, 0, 0.0, 0.0, 0.0]
    for step in [0] + sorted(scfg["steps"]):
        grp = [a for a in rated if a["discount"] == step]
        if not grp:
            continue
        stock = sum(a["stock"] for a in grp)
        value = sum(float(a["ek"]) * a["stock"] for a in grp)
        revenue = sum(a["sale_price"] / (1 + scfg["vat"]) * a["stock"] for a in grp)
        profit = sum(a["potential"] for a in grp)
        factor = revenue / value if value else None
        vals = [f"{step} %" if step else "kein Rabatt", len(grp), stock, value, revenue, profit, factor]
        for i, v in enumerate(vals, 1):
            c = so.cell(r, i, v)
            c.font = Font(name="Arial", size=10)
            c.number_format = ["@", "0", "#,##0", "#,##0 €", "#,##0 €", "#,##0 €", "0.00"][i - 1]
        so.cell(r, 1).fill = PatternFill("solid", fgColor=STEP_COLORS.get(step, "FFF8CBAD"))
        for k, v in enumerate([len(grp), stock, value, revenue, profit]):
            totals[k] += v
        r += 1
    for i, v in enumerate(["Gesamt"] + totals + [totals[3] / totals[2] if totals[2] else None], 1):
        c = so.cell(r, i, v)
        c.font = Font(name="Arial", size=10, bold=True)
        c.number_format = ["@", "0", "#,##0", "#,##0 €", "#,##0 €", "#,##0 €", "0.00"][i - 1]
        c.border = Border(top=Side(style="medium", color=DARK))
    if unrated:
        so.cell(r + 2, 1, f"{len(unrated)} Artikel ohne Empfehlung (kein EK oder nicht in Shopify) – unten in „Sale-Empfehlung“.").font = Font(name="Arial", size=10, color="FF7F7F7F")
    below = [a for a in rated if a.get("below_margin")]
    if below:
        c = so.cell(r + 1, 1, f"⚠ {len(below)} Artikel mit {scfg['steps'][0]} % unter der Mindestmarge – in „Sale-Empfehlung“ rot markiert.")
        c.font = Font(name="Arial", size=10, bold=True, color="FF9C0006")
    if scfg.get("_excluded"):
        so.cell(r + 4, 1, "Ausgeschlossen (Artikelnummer beginnt mit): " + ", ".join(scfg["_excluded"])).font = Font(name="Arial", size=9, color="FF7F7F7F")
    so.cell(r + 3, 1, "Rohertrag = Sale-Preis netto − EK − erwartete Retourenkosten; „bei Abverkauf“ = wenn der ganze Bestand zum Sale-Preis verkauft wird.").font = Font(name="Arial", size=9, color="FF7F7F7F")
    wb.save(out)
