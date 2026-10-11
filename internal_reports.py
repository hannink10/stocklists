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
    loss = [a for a in rated if a.get("below_ek")]
    if loss:
        print(f"  ⛔ {len(loss)} Artikel würden unter EK verkauft (dunkelrot markiert)")
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
    try:
        received = client.inventory_received_by_month(scfg["restock_months"])
    except sc.ShopifyError as e:
        received = {}
        warnings.append(f"Wareneingänge nicht abrufbar (App-Recht read_reports?) – Nachbestellungen werden nicht erkannt: {e}")

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
                               "sizes": 0, "sizes_stock": 0, "since": None, "on_sale": False, "found": False,
                               "received": {}})
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
    for sku, by_month in received.items():
        b = base_of(sku)
        if b:
            for month, units in by_month.items():
                per[b]["received"][month] = per[b]["received"].get(month, 0) + units
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
        restocks = sorted(m for m, u in a.get("received", {}).items() if u >= scfg["restock_min_units"])
        a["last_restock"] = restocks[-1] if restocks else None
        a["received_12m"] = sum(a.get("received", {}).values())
        a["restocked"] = bool(a["last_restock"]) and (today - a["last_restock"]).days <= scfg["restock_protect_days"]
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
        # 0) Restposten (SKU-Reduktion): sehr alt und noch auf Lager -> fester Rabatt, egal wie er sich verkauft;
        #    liegt die Marge dann unter der Mindestmarge, bleibt der Rabatt – die Zeile wird markiert
        if (scfg.get("restposten_days") and a["age"] is not None and a["age"] >= scfg["restposten_days"]
                and a["stock"] > 0 and not a["restocked"]):
            a["discount"] = scfg["restposten_discount"]
            a["restposten"] = True
            a["below_margin"] = a["discount"] / 100 > a["max_disc"] + 1e-9
            a["reason"] = ((f"⚠ UNTER MINDESTMARGE – Marge erlaubt nur max. {pct(max(0, a['max_disc']))}, " if a["below_margin"] else "")
                           + f"Restposten: seit {a['since'].strftime('%d.%m.%Y')} im Shop ({a['age']} Tage) – SKU-Reduktion")
            finish(a, scfg)
            continue
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
        if a["restocked"]:
            # frisch nachbestellt = Dauerbrenner, kein Ladenhüter -> höchstens die kleinste Stufe
            desired = min(desired, steps[0])
            reasons = [f"Dauerbrenner: nachbestellt {a['last_restock'].strftime('%m/%Y')} "
                       f"(+{a['received'].get(a['last_restock'], 0)}), {a['recent_net']} verkauft in {days} Tagen"]
        elif a["net_sold"] == 0 and a["age"] is not None and a["age"] >= scfg["slow_days"]:
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
        if desired and not a["restocked"] and a["sizes"] >= 4 and a["sizes_stock"] <= scfg["broken_sizes_max"]:
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
    # Verkauf unter Einkaufspreis (nur möglich, wenn ein Rabatt die Mindestmarge bewusst übergeht)
    a["below_ek"] = bool(d) and a["factor"] is not None and a["factor"] < 1
    if a["below_ek"] and not (a.get("reason") or "").startswith("⛔"):
        a["reason"] = f"⛔ UNTER EK – Verlust {abs(net - float(a['ek'])):.2f} € je Teil, " + (a.get("reason") or "")


def eur(x):
    return f"{x:,.0f} €".replace(",", ".")


def pct(x):
    return "–" if x is None else f"{x * 100:.0f} %"


# ---------------------------------------------------------------------------
# Excel

COLUMNS = [
    # (Titel, Breite, Zahlenformat, Schlüssel)
    ("Rang", 6, None, "rank"), ("SKU", 13, None, "sku"), ("Artikel", 40, None, "name"), ("Kategorie", 12, None, "cat"),
    ("Rabatt (änderbar)", 11, "0%", "disc"), ("Vorschlag", 9, "0%", "suggest"),
    ("Sale-Preis", 11, "#,##0.00 €", "sale"), ("Faktor danach", 9, "0.00", "factor"),
    ("Rohertrag/Stk", 11, "#,##0.00 €", "profit"), ("Rohertrag-Potenzial", 13, "#,##0 €", "potential"),
    ("Umsatz netto bei Abverkauf", 13, "#,##0 €", "revenue"), ("Status", 15, None, "status"),
    ("Bestand", 8, "0", "stock"), ("Lagerwert EK", 12, "#,##0 €", "value"),
    ("Verkauft 60 T.", 9, "0", "recent"), ("Abverkauf 60 T.", 9, "0%", "str_recent"),
    ("Verkauft 12 Mon.", 9, "0", "year"), ("Abverkauf 12 Mon.", 9, "0%", "str"), ("Retouren", 9, "0%", "rr"),
    ("Reichweite Wo.", 10, "0", "reach"), ("Trend 30 T.", 10, None, "trend"), ("Letzter Verkauf (Tage)", 10, "0", "last"),
    ("Ø bezahlt", 10, "#,##0.00 €", "paid"), ("Größen verfügbar", 10, None, "sizes"), ("Im Shop seit", 11, "DD.MM.YYYY", "since"),
    ("Letzter Wareneingang", 11, "MM/YYYY", "restock"), ("Eingang 12 Mon.", 9, "0", "received"),
    ("Max. Rabatt (Marge)", 10, "0%", "maxdisc"), ("EK netto", 10, "#,##0.00 €", "ek"), ("VK aktuell", 10, "#,##0.00 €", "vk"),
    ("Begründung", 60, None, "reason"),
]
COL = {key: i for i, (_, _, _, key) in enumerate(COLUMNS, 1)}
L = {key: get_column_letter(i) for key, i in COL.items()}

# Einstellungen oben in "Sale-Empfehlung" (änderbar, Formeln rechnen damit)
P_MARGIN, P_VAT, P_RETURN = "$B$3", "$D$3", "$F$3"


def static_values(rank, a, category):
    trend = None
    if a.get("trend") is not None:
        trend = "neu ↑" if math.isinf(a["trend"]) else f"{'↑' if a['trend'] > 1.1 else '↓' if a['trend'] < 0.9 else '→'} {a['recent']} vs {a['before']}"
    d = a.get("discount")
    return {
        "rank": rank, "sku": a["sku"], "name": a["name"], "cat": category,
        "disc": d / 100 if d is not None else None, "suggest": d / 100 if d is not None else None,
        "stock": a["stock"], "value": float(a["ek"]) * a["stock"], "recent": a.get("recent_net"), "str_recent": a.get("str_recent"),
        "year": a.get("net_sold"), "str": a.get("str"), "rr": a.get("return_rate"),
        "reach": None if a.get("reach") is None or math.isinf(a["reach"]) else round(a["reach"]), "trend": trend,
        "last": a.get("last_days"), "paid": a.get("avg_paid"),
        "sizes": f"{a['sizes_stock']}/{a['sizes']}" if a.get("sizes") else None, "since": a.get("since"),
        "restock": a.get("last_restock"), "received": a.get("received_12m") or None,
        "ek": float(a["ek"]), "vk": float(a["vk"]), "reason": a.get("reason"),
    }


def row_formulas(r):
    """Alles, was vom Rabatt abhängt, als Excel-Formel – ändert man den Rabatt, rechnet die Zeile neu."""
    c = {k: f"${L[k]}{r}" for k in L}
    return {
        "sale": f'=IF(OR({c["disc"]}="",{c["vk"]}=0),"",ROUND({c["vk"]}*(1-{c["disc"]}),2))',
        "factor": f'=IF(OR({c["sale"]}="",{c["ek"]}=0),"",{c["sale"]}/(1+{P_VAT})/{c["ek"]})',
        "profit": f'=IF({c["sale"]}="","",{c["sale"]}/(1+{P_VAT})-{c["ek"]}-N({c["rr"]})*{P_RETURN})',
        "potential": f'=IF({c["profit"]}="","",{c["profit"]}*{c["stock"]})',
        "revenue": f'=IF({c["sale"]}="","",{c["sale"]}/(1+{P_VAT})*{c["stock"]})',
        "status": (f'=IF(OR({c["disc"]}="",{c["disc"]}=0,{c["factor"]}=""),"",IF({c["factor"]}<1,"⛔ UNTER EK",'
                   f'IF({c["factor"]}<{P_MARGIN},"⚠ UNTER MARGE","")))'),
        "maxdisc": f'=IF(OR({c["ek"]}=0,{c["vk"]}=0),"",MAX(0,1-{P_MARGIN}*{c["ek"]}*(1+{P_VAT})/{c["vk"]}))',
    }


def header_cell(c):
    c.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor=DARK)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_report(out, articles, scfg, coverage, brand, today, groups_cfg=None):
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation

    rated = [a for a in articles if a["discount"] is not None]
    unrated = [a for a in articles if a["discount"] is None]
    with_disc = sorted([a for a in rated if a["discount"] > 0], key=lambda a: -a["potential"])
    without = sorted([a for a in rated if a["discount"] == 0], key=lambda a: -(a.get("recent_net") or 0))
    ordered = with_disc + without + unrated

    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # Formeln beim Öffnen berechnen
    ws = wb.active
    ws.title = "Sale-Empfehlung"
    period = f"{coverage['oldest'].strftime('%d.%m.%Y')}–{today.strftime('%d.%m.%Y')} ({coverage['days']} Tage)" if coverage["oldest"] else "–"
    ws["A1"] = f"SALE-EMPFEHLUNG {brand.upper()} – {today.strftime('%d.%m.%Y')}"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=DARK)
    excluded = scfg.get("_excluded") or []
    ws["H1"] = (f"Shopify-Daten: {period} · " + (f"ausgeschlossen: {', '.join(excluded)}… · " if excluded else "")
                + "Spalte „Rabatt“ ist änderbar – Sale-Preis, Faktor, Rohertrag, Status und Übersicht rechnen mit.")
    ws["H1"].font = Font(name="Arial", size=10, color="FF7F7F7F")

    # Einstellungen (änderbar)
    edit_fill = PatternFill("solid", fgColor="FFFFF9C4")
    for label_col, value_col, label, value, fmt in (("A", "B", "Mindestmarge", scfg["min_margin"], "0.00"),
                                                    ("C", "D", "MwSt", scfg["vat"], "0%"),
                                                    ("E", "F", "Retourenkosten €", scfg["return_cost_eur"], "#,##0.00 €")):
        lc, vc = ws[f"{label_col}3"], ws[f"{value_col}3"]
        lc.value = label
        lc.font = Font(name="Arial", size=10, bold=True, color=DARK)
        lc.alignment = Alignment(horizontal="right")
        vc.value = value
        vc.number_format = fmt
        vc.fill = edit_fill
        vc.font = Font(name="Arial", size=10, bold=True)
        vc.border = Border(*(Side(style="thin", color="FF7F7F7F"),) * 4)

    hr = 4
    for i, (title, width, _, _) in enumerate(COLUMNS, 1):
        header_cell(ws.cell(hr, i, title))
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[hr].height = 32
    first, last = hr + 1, hr + len(ordered)

    for idx, a in enumerate(ordered, 1):
        r = hr + idx
        category = g.product_group(a["name"], groups_cfg, a["sku"]) if groups_cfg and groups_cfg.get("rules") else ""
        values = static_values(idx if a in with_disc else None, a, category)
        if a["discount"] is not None:
            values.update(row_formulas(r))
        else:
            values["maxdisc"] = row_formulas(r)["maxdisc"]
        for key, i in COL.items():
            c = ws.cell(r, i, values.get(key))
            c.font = Font(name="Arial", size=10, bold=key in ("disc", "sale", "status"))
            c.border = Border(bottom=THIN)
            fmt = COLUMNS[i - 1][2]
            if fmt:
                c.number_format = fmt
            if i > 4 and key != "reason":
                c.alignment = Alignment(horizontal="center")
        if a["discount"] is not None:
            ws.cell(r, COL["disc"]).fill = edit_fill
            ws.cell(r, COL["disc"]).border = Border(*(Side(style="thin", color="FFBFBFBF"),) * 4)

    if ordered:
        disc_rng = f"{L['disc']}{first}:{L['disc']}{last}"
        dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.95", allow_blank=True,
                            showInputMessage=True, promptTitle="Rabatt", prompt="Rabatt eingeben, z.B. 30 % – alles Weitere rechnet neu.",
                            showErrorMessage=True, errorTitle="Rabatt", error="Bitte einen Rabatt zwischen 0 % und 95 % eingeben.")
        dv.add(disc_rng)
        ws.add_data_validation(dv)
        full = f"A{first}:{get_column_letter(len(COLUMNS))}{last}"
        d, f = f"${L['disc']}{first}", f"${L['factor']}{first}"
        # Markierungen als bedingte Formatierung: passen sich an, wenn man den Rabatt ändert
        ws.conditional_formatting.add(full, FormulaRule(formula=[f'AND(N({d})>0,{f}<>"",{f}<1)'], stopIfTrue=True,
                                      fill=PatternFill("solid", fgColor="FFC00000", bgColor="FFC00000"),
                                      font=Font(bold=True, color="FFFFFFFF")))
        ws.conditional_formatting.add(full, FormulaRule(formula=[f'AND(N({d})>0,{f}<>"",{f}<{P_MARGIN})'], stopIfTrue=True,
                                      fill=PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE"),
                                      font=Font(color="FF9C0006")))
        ws.conditional_formatting.add(disc_rng, FormulaRule(formula=[f'AND({d}<>"",{d}<>${L["suggest"]}{first})'],
                                      font=Font(bold=True, color="FF1F4E79"),
                                      fill=PatternFill("solid", fgColor="FFDDEBF7", bgColor="FFDDEBF7")))
        for step, color in STEP_COLORS.items():
            ws.conditional_formatting.add(f"{L['suggest']}{first}:{L['suggest']}{last}",
                                          FormulaRule(formula=[f'${L["suggest"]}{first}={step / 100}'],
                                                      fill=PatternFill("solid", fgColor=color, bgColor=color)))
    ws.freeze_panes = ws.cell(first, COL["cat"] + 1)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(COLUMNS))}{last}"

    # Übersicht je Rabattstufe – per Formel aus "Sale-Empfehlung"
    so = wb.create_sheet("Übersicht", 0)
    so["A1"] = f"SALE-ÜBERSICHT {brand.upper()} – {today.strftime('%d.%m.%Y')}"
    so["A1"].font = Font(name="Arial", size=16, bold=True, color=DARK)
    so["A2"] = (f"Shopify-Daten: {period} · Retourenquote gesamt "
                f"{pct(coverage['returned'] / coverage['sold']) if coverage['sold'] else '–'} · rechnet mit den Rabatten aus „Sale-Empfehlung“")
    so["A2"].font = Font(name="Arial", size=10, color="FF7F7F7F")
    heads = ["Rabatt", "Artikel", "Bestand", "Lagerwert EK", "Umsatz netto bei Abverkauf", "Rohertrag bei Abverkauf", "Ø Faktor"]
    for i, (h, w) in enumerate(zip(heads, [14, 10, 10, 14, 20, 20, 10]), 1):
        header_cell(so.cell(4, i, h))
        so.column_dimensions[get_column_letter(i)].width = w
    sheet = "'Sale-Empfehlung'!"
    rng = {k: f"{sheet}${L[k]}${first}:${L[k]}${max(first, last)}" for k in ("disc", "stock", "value", "revenue", "potential", "status")}
    fmts = ["@", "0", "#,##0", "#,##0 €", "#,##0 €", "#,##0 €", "0.00"]
    steps = [0] + sorted(scfg["steps"])
    r = 5
    for step in steps:
        crit = f"{step / 100}"
        vals = [f"{step} %" if step else "kein Rabatt", f"=COUNTIF({rng['disc']},{crit})", f"=SUMIF({rng['disc']},{crit},{rng['stock']})",
                f"=SUMIF({rng['disc']},{crit},{rng['value']})", f"=SUMIF({rng['disc']},{crit},{rng['revenue']})",
                f"=SUMIF({rng['disc']},{crit},{rng['potential']})", f'=IF(D{r}=0,"",E{r}/D{r})']
        for i, v in enumerate(vals, 1):
            c = so.cell(r, i, v)
            c.font = Font(name="Arial", size=10)
            c.number_format = fmts[i - 1]
        so.cell(r, 1).fill = PatternFill("solid", fgColor=STEP_COLORS.get(step, "FFF8CBAD"))
        r += 1
    top, end = 5, r - 1
    # eigene Werte (z.B. 25 %) = alles, was nicht in den Stufen liegt
    other = ["andere %", f"=COUNT({rng['disc']})-SUM(B{top}:B{end})", f"=SUMIF({rng['disc']},\">=0\",{rng['stock']})-SUM(C{top}:C{end})",
             f"=SUMIF({rng['disc']},\">=0\",{rng['value']})-SUM(D{top}:D{end})", f"=SUM({rng['revenue']})-SUM(E{top}:E{end})",
             f"=SUM({rng['potential']})-SUM(F{top}:F{end})", f'=IF(D{r}=0,"",E{r}/D{r})']
    for i, v in enumerate(other, 1):
        c = so.cell(r, i, v)
        c.font = Font(name="Arial", size=10, italic=True)
        c.number_format = fmts[i - 1]
    r += 1
    total = ["Gesamt"] + [f"=SUM({col}{top}:{col}{r - 1})" for col in "BCDEF"] + [f'=IF(D{r}=0,"",E{r}/D{r})']
    for i, v in enumerate(total, 1):
        c = so.cell(r, i, v)
        c.font = Font(name="Arial", size=10, bold=True)
        c.number_format = fmts[i - 1]
        c.border = Border(top=Side(style="medium", color=DARK))

    notes = [
        (f'="⚠ "&COUNTIF({rng["status"]},"⚠*")&" Artikel unter der Mindestmarge (rot markiert)"', Font(name="Arial", size=10, bold=True, color="FF9C0006")),
        (f'="⛔ "&COUNTIF({rng["status"]},"⛔*")&" Artikel UNTER EK – Verlust bei Abverkauf "&TEXT(-SUMIF({rng["status"]},"⛔*",{rng["potential"]}),"#,##0")&" € (dunkelrot)"',
         Font(name="Arial", size=10, bold=True, color="FFC00000")),
    ]
    rest = [a for a in rated if a.get("restposten")]
    if rest:
        notes.append((f"Restposten (≥ {scfg['restposten_days'] / 365:.0f} Jahre im Shop, SKU-Reduktion): {len(rest)} Artikel vorgeschlagen mit "
                      f"{scfg['restposten_discount']} %.", Font(name="Arial", size=10, bold=True, color=DARK)))
    if unrated:
        notes.append((f"{len(unrated)} Artikel ohne Empfehlung (kein EK oder nicht in Shopify) – unten in „Sale-Empfehlung“.", Font(name="Arial", size=10, color="FF7F7F7F")))
    if excluded:
        notes.append(("Ausgeschlossen (Artikelnummer beginnt mit): " + ", ".join(excluded), Font(name="Arial", size=9, color="FF7F7F7F")))
    notes.append(("Rohertrag = Sale-Preis netto − EK − Retourenquote × Retourenkosten; „bei Abverkauf“ = der ganze Bestand wird zum Sale-Preis verkauft. "
                  "Mindestmarge, MwSt und Retourenkosten stehen änderbar oben in „Sale-Empfehlung“.", Font(name="Arial", size=9, color="FF7F7F7F")))
    for k, (text, font) in enumerate(notes, r + 2):
        so.cell(k, 1, text).font = font
    wb.save(out)
