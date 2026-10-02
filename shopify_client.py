"""Shopify-Anbindung (nur lesend): Produktbilder und Verkaufszahlen je SKU.

Zugangsdaten kommen aus Umgebungsvariablen oder aus config/shopify.env
(diese Datei ist in .gitignore und wird nie hochgeladen):

    SHOPIFY_SHOP=reternity.myshopify.com
    SHOPIFY_CLIENT_ID=...
    SHOPIFY_CLIENT_SECRET=...
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


def load_credentials(env_file):
    values = {}
    if env_file.is_file():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    creds = {}
    for key in ("SHOPIFY_SHOP", "SHOPIFY_CLIENT_ID", "SHOPIFY_CLIENT_SECRET"):
        creds[key] = os.environ.get(key) or values.get(key)
    if not all(creds.values()):
        return None
    return creds


class ShopifyClient:
    def __init__(self, creds, api_version, timeout=30):
        self.shop = creds["SHOPIFY_SHOP"].replace("https://", "").strip("/")
        self.client_id = creds["SHOPIFY_CLIENT_ID"]
        self.client_secret = creds["SHOPIFY_CLIENT_SECRET"]
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

    def image_urls_by_sku(self):
        """SKU -> Bild-URL (Variantenbild, sonst Hauptbild des Produkts)."""
        query = """
        query($after: String) {
          productVariants(first: 250, after: $after) {
            pageInfo { hasNextPage endCursor }
            nodes {
              sku
              image { url }
              product { featuredMedia { preview { image { url } } } }
            }
          }
        }"""
        urls, after = {}, None
        while True:
            data = self.graphql(query, {"after": after})["productVariants"]
            for v in data["nodes"]:
                sku = (v.get("sku") or "").strip()
                if not sku:
                    continue
                url = (v.get("image") or {}).get("url")
                if not url:
                    media = (v.get("product") or {}).get("featuredMedia") or {}
                    url = ((media.get("preview") or {}).get("image") or {}).get("url")
                if url:
                    urls[sku] = url
            if not data["pageInfo"]["hasNextPage"]:
                return urls
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
