"""Shopify-Anbindung (nur lesend): Produktbilder und Verkaufszahlen je SKU.

Zugangsdaten kommen aus config/shopify.env oder aus Umgebungsvariablen
(config/shopify.env ist in .gitignore und wird nie hochgeladen):

    SHOPIFY_CLIENT_ID=...
    SHOPIFY_CLIENT_SECRET=...

    # ein Shop pro Marke: SHOP_<Marke>=<adresse>.myshopify.com
    SHOP_RETERNITY=reternity.myshopify.com
    SHOP_SAINT_SASS=saint-sass.myshopify.com

    # nur falls ein Shop eine eigene App hat:
    SHOPIFY_CLIENT_ID_SAINT_SASS=...
    SHOPIFY_CLIENT_SECRET_SAINT_SASS=...

Älteres Format mit nur einem Shop (SHOPIFY_SHOP=...) funktioniert weiter.
"""

import hashlib
import io
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path

from PIL import Image as PILImage


class ShopifyError(Exception):
    pass


def _read_env(env_file):
    """Werte aus config/shopify.env, ergänzt um gleichnamige Umgebungsvariablen."""
    values = {}
    if env_file.is_file():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    for key, value in sorted(os.environ.items()):
        if key.startswith(("SHOP_", "SHOPIFY_")) and value:
            values[key] = value
    return values


def normalize_domain(value):
    """Shop-Adresse vereinheitlichen: auch Admin-Links und Kurzformen zu <name>.myshopify.com."""
    value = value.strip().replace("https://", "").replace("http://", "").strip("/")
    if value.startswith("admin.shopify.com/store/"):
        value = value.split("/")[2]
    value = value.split("/")[0]
    return value if "." in value else f"{value}.myshopify.com"


def load_shops(env_file, default_name):
    """Liste der konfigurierten Shops: [{name, shop, client_id, client_secret}, …]."""
    values = _read_env(env_file)
    entries = [(k[len("SHOP_"):], v) for k, v in values.items() if k.startswith("SHOP_") and v]
    if not entries and values.get("SHOPIFY_SHOP"):
        entries = [(default_name.upper().replace(" ", "_"), values["SHOPIFY_SHOP"])]
    shops = []
    for key, domain in entries:
        key = key.strip().upper().replace(" ", "_")
        shops.append({
            "name": key.replace("_", " ").title(),
            "shop": normalize_domain(domain),
            "client_id": values.get(f"SHOPIFY_CLIENT_ID_{key}") or values.get("SHOPIFY_CLIENT_ID"),
            "client_secret": values.get(f"SHOPIFY_CLIENT_SECRET_{key}") or values.get("SHOPIFY_CLIENT_SECRET"),
        })
    return shops


class ShopifyClient:
    def __init__(self, shop, api_version, timeout=30):
        self.shop = normalize_domain(shop["shop"])
        self.client_id = shop["client_id"]
        self.client_secret = shop["client_secret"]
        self.api_version = api_version
        self.timeout = timeout
        self._token = None

    # -- HTTP ---------------------------------------------------------------

    def _post(self, url, data, headers):
        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")[:300]
            raise ShopifyError(f"HTTP {e.code} von {url}: {body}")
        except urllib.error.URLError as e:
            raise ShopifyError(f"Shopify nicht erreichbar ({self.shop}): {e.reason}")

    def token(self):
        # Client-Credentials-Grant: Token gilt 24 h, wird pro Lauf neu geholt
        if self._token is None:
            data = urllib.parse.urlencode({
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            }).encode()
            result = self._post(
                f"https://{self.shop}/admin/oauth/access_token", data,
                {"Content-Type": "application/x-www-form-urlencoded"},
            )
            if "access_token" not in result:
                raise ShopifyError(f"Kein Access-Token erhalten: {result}")
            self._token = result["access_token"]
        return self._token

    def graphql(self, query, variables=None):
        url = f"https://{self.shop}/admin/api/{self.api_version}/graphql.json"
        payload = json.dumps({"query": query, "variables": variables or {}}).encode()
        for attempt in range(5):
            result = self._post(url, payload, {
                "Content-Type": "application/json",
                "X-Shopify-Access-Token": self.token(),
            })
            errors = result.get("errors")
            if errors and any(e.get("extensions", {}).get("code") == "THROTTLED" for e in errors):
                time.sleep(2 ** attempt)
                continue
            if errors:
                raise ShopifyError(f"GraphQL-Fehler: {errors}")
            return result["data"]
        raise ShopifyError("Shopify-API: zu viele Anfragen (gedrosselt).")

    # -- Daten --------------------------------------------------------------

    def image_variants(self):
        """Alle Varianten mit Bild: Liste von {sku, title, status, url}.

        Bild = Variantenbild, sonst Hauptbild des Produkts. Titel und Status werden
        mitgeliefert, damit Artikel auch über den Produktnamen gefunden werden können.
        """
        query = """
        query($after: String) {
          productVariants(first: 250, after: $after) {
            pageInfo { hasNextPage endCursor }
            nodes {
              sku
              image { url }
              product { title status featuredMedia { preview { image { url } } } }
            }
          }
        }"""
        variants, after = [], None
        while True:
            data = self.graphql(query, {"after": after})["productVariants"]
            for v in data["nodes"]:
                product = v.get("product") or {}
                url = (v.get("image") or {}).get("url")
                if not url:
                    media = product.get("featuredMedia") or {}
                    url = ((media.get("preview") or {}).get("image") or {}).get("url")
                if url:
                    variants.append({
                        "sku": (v.get("sku") or "").strip(),
                        "title": product.get("title") or "",
                        "status": product.get("status") or "",
                        "url": url,
                    })
            if not data["pageInfo"]["hasNextPage"]:
                return variants
            after = data["pageInfo"]["endCursor"]

    def units_sold_by_sku(self, days):
        """SKU -> verkaufte Stück der letzten `days` Tage (ohne stornierte Bestellungen)."""
        since = (date.today() - timedelta(days=days)).isoformat()
        query = """
        query($after: String, $q: String) {
          orders(first: 100, after: $after, query: $q) {
            pageInfo { hasNextPage endCursor }
            nodes {
              lineItems(first: 250) { nodes { sku quantity } }
            }
          }
        }"""
        sold, after = {}, None
        q = f"created_at:>={since} AND -status:cancelled"
        while True:
            data = self.graphql(query, {"after": after, "q": q})["orders"]
            for order in data["nodes"]:
                for li in order["lineItems"]["nodes"]:
                    sku = (li.get("sku") or "").strip()
                    if sku:
                        sold[sku] = sold.get(sku, 0) + int(li.get("quantity") or 0)
            if not data["pageInfo"]["hasNextPage"]:
                return sold
            after = data["pageInfo"]["endCursor"]


def thumbnail(url, cache_dir, key, max_px, timeout=30):
    """Lädt ein Bild (mit Cache) und liefert ein verkleinertes JPEG als BytesIO."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    safe_key = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in key)
    # URL-Hash im Namen: ändert sich das Bild in Shopify, wird es neu geladen
    url_hash = hashlib.sha1(url.encode()).hexdigest()[:10]
    cached = cache_dir / f"{safe_key}_{url_hash}_{max_px[0]}x{max_px[1]}.jpg"
    if not cached.is_file():
        sep = "&" if "?" in url else "?"
        src = f"{url}{sep}width={max_px[0] * 3}"  # Shopify-CDN liefert direkt eine kleinere Version
        try:
            with urllib.request.urlopen(src, timeout=timeout) as resp:
                raw = resp.read()
        except urllib.error.URLError as e:
            raise ShopifyError(f"Bild nicht ladbar: {e}")
        img = PILImage.open(io.BytesIO(raw))
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
            bg = PILImage.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        else:
            img = img.convert("RGB")
        # doppelte Auflösung speichern, damit es auf Retina-Bildschirmen scharf ist
        img.thumbnail((max_px[0] * 2, max_px[1] * 2))
        img.save(cached, "JPEG", quality=80, optimize=True)
    return cached
