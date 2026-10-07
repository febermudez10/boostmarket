from fastapi.staticfiles import StaticFiles
import os
import re
import requests
from datetime import datetime, timezone
from urllib.parse import urljoin

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed

import psycopg
from psycopg.rows import dict_row


# ============================================================
# CONFIGURACION
# ============================================================

DATABASE_URL = os.environ.get("DATABASE_URL")
AUTO_DEV_API_KEY = os.environ.get("AUTO_DEV_API_KEY")

TIMEOUT = 30

HEADERS = {
    "User-Agent": "South-Dade-Toyota-Inventory-Manager/1.0"
}

SOURCE = "https://www.southdadetoyota.com/llm/inventory/?type=used"


# ============================================================
# APP
# ============================================================

app = FastAPI(title="BoostMarket - South Dade Toyota")

app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)

templates = Jinja2Templates(directory="templates")


# ============================================================
# DATABASE
# ============================================================

def db():
    if not DATABASE_URL:
        raise RuntimeError(
            "DATABASE_URL no está configurado en Render."
        )

    return psycopg.connect(
        DATABASE_URL,
        row_factory=dict_row
    )


def init_db():
    with db() as c:
        with c.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS vehicles (
                    vin TEXT PRIMARY KEY,
                    vehicle TEXT,
                    condition TEXT,
                    mileage TEXT,
                    your_price DOUBLE PRECISION,
                    marketplace_price DOUBLE PRECISION,
                    url TEXT,
                    photo TEXT,
                    stock TEXT,
                    pic_count INTEGER DEFAULT 0,
                    status TEXT DEFAULT 'pending',
                    lead_count INTEGER DEFAULT 0,
                    last_seen TEXT,
                    updated_at TEXT
                )
            """)

            # Migraciones seguras para bases de datos existentes.
            cur.execute("ALTER TABLE vehicles ADD COLUMN IF NOT EXISTS stock TEXT")
            cur.execute("ALTER TABLE vehicles ADD COLUMN IF NOT EXISTS pic_count INTEGER DEFAULT 0")

            cur.execute("""
                CREATE TABLE IF NOT EXISTS vehicle_photos (
                    vin TEXT NOT NULL,
                    position INTEGER NOT NULL,
                    url TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (vin, position)
                )
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_vehicle_photos_vin ON vehicle_photos(vin)")

            # Migración de una sola vez:
            # borra los totales que anteriormente se importaron desde VINCUE.
            # Una vez aplicada, los futuros reinicios NO borrarán leads de Marketplace.
            cur.execute("""
                CREATE TABLE IF NOT EXISTS app_migrations (
                    migration_key TEXT PRIMARY KEY,
                    applied_at TEXT NOT NULL
                )
            """)
            cur.execute("""
                SELECT 1
                FROM app_migrations
                WHERE migration_key = %s
            """, ("reset_vincue_leads_v1",))

            if not cur.fetchone():
                cur.execute("UPDATE vehicles SET lead_count = 0")
                cur.execute("""
                    INSERT INTO app_migrations (migration_key, applied_at)
                    VALUES (%s, %s)
                """, (
                    "reset_vincue_leads_v1",
                    datetime.now(timezone.utc).isoformat()
                ))

        c.commit()


# ============================================================
# UTILIDADES
# ============================================================

def money(s):
    if not s:
        return None

    m = re.search(
        r'\$?\s*([\d,]+(?:\.\d{2})?)',
        s
    )

    if not m:
        return None

    return float(
        m.group(1).replace(",", "")
    )


# ============================================================
# WEBSITE INVENTORY COLLECTOR
# ============================================================

def get_primary_photo(detail_url):
    if not detail_url or detail_url == SOURCE:
        return ""

    try:
        response = requests.get(
            detail_url,
            headers=HEADERS,
            timeout=TIMEOUT
        )
        response.raise_for_status()

        page = BeautifulSoup(response.text, "lxml")

        # 1. Open Graph / Twitter image: normalmente es la foto principal.
        for attrs in (
            {"property": "og:image"},
            {"property": "og:image:secure_url"},
            {"name": "twitter:image"},
        ):
            meta = page.find("meta", attrs=attrs)
            if meta and meta.get("content"):
                candidate = urljoin(response.url, meta["content"].strip())
                if candidate.startswith(("http://", "https://")):
                    return candidate

        # 2. Algunos sitios declaran la imagen principal con image_src.
        image_src = page.find("link", rel="image_src")
        if image_src and image_src.get("href"):
            candidate = urljoin(response.url, image_src["href"].strip())
            if candidate.startswith(("http://", "https://")):
                return candidate

        # 3. Fallback: buscar una imagen grande de inventario/galería.
        bad_words = (
            "logo", "icon", "carfax", "pixel", "spinner",
            "placeholder", "avatar", "badge", "toyota-logo"
        )

        for img in page.find_all("img"):
            raw = (
                img.get("data-src")
                or img.get("data-lazy-src")
                or img.get("data-original")
                or img.get("src")
                or ""
            ).strip()

            if not raw or raw.startswith("data:"):
                continue

            candidate = urljoin(response.url, raw)
            low = candidate.lower()

            if any(word in low for word in bad_words):
                continue

            if any(ext in low for ext in (".jpg", ".jpeg", ".png", ".webp")):
                return candidate

    except Exception:
        # Una foto nunca debe hacer fallar toda la sincronización.
        return ""

    return ""


def collect():

    r = requests.get(
        SOURCE,
        headers=HEADERS,
        timeout=TIMEOUT
    )

    r.raise_for_status()

    soup = BeautifulSoup(
        r.text,
        "lxml"
    )

    found = {}

    for block in soup.find_all(
        ["article", "li", "div"]
    ):

        txt = " ".join(
            block.stripped_strings
        )

        vm = re.search(
            r'VIN:\s*([A-HJ-NPR-Z0-9]{17})',
            txt,
            re.I
        )

        if not vm:
            continue

        vin = vm.group(1).upper()

        if vin in found:
            continue

        pm = re.search(
            r'\$[\d,]+(?:\.\d{2})?',
            txt
        )

        mm = re.search(
            r'([\d,]+)\s*miles',
            txt,
            re.I
        )

        name = None

        for tag in block.find_all(
            ["h2", "h3", "h4", "a"]
        ):

            t = " ".join(
                tag.stripped_strings
            )

            if (
                re.search(r'\b20\d{2}\b', t)
                and len(t) < 160
            ):
                name = t
                break

        if not name:

            nm = re.search(
                r'\b20\d{2}\s+[A-Za-z0-9\- ]+'
                r'(?:LE|SE|XLE|SR5|EX|LX|Limited|Sport|Premium|Hybrid)?\b',
                txt
            )

            name = (
                nm.group(0).strip()
                if nm
                else "Used vehicle"
            )

        href = None

        for a in block.find_all(
            "a",
            href=True
        ):

            h = urljoin(
                r.url,
                a["href"]
            )

            if (
                "southdadetoyota.com" in h
                and (
                    "inventory" in h
                    or "/used-" in h
                )
            ):
                href = h
                break

        price = (
            money(pm.group(0))
            if pm
            else None
        )

        detail_url = href or SOURCE
        photo_url = get_primary_photo(detail_url)

        found[vin] = {
            "vin": vin,
            "vehicle": name,
            "condition": "Used",
            "mileage": (
                mm.group(1).replace(",", "")
                if mm
                else ""
            ),
            "your_price": price,
            "marketplace_price": (
                price + 1000
                if price is not None
                else None
            ),
            "url": detail_url,
            "photo": photo_url
        }


    now = datetime.now(
        timezone.utc
    ).isoformat()


    with db() as c:

        with c.cursor() as cur:

            for v in found.values():

                cur.execute(
                    """
                    SELECT status, lead_count, photo
                    FROM vehicles
                    WHERE vin = %s
                    """,
                    (v["vin"],)
                )

                old = cur.fetchone()

                status = (
                    old["status"]
                    if old
                    else "pending"
                )

                leads = (
                    old["lead_count"]
                    if old
                    else 0
                )

                # Si el website no devuelve una foto temporalmente,
                # conservar la que ya estaba guardada en Supabase.
                if not v["photo"] and old and old.get("photo"):
                    v["photo"] = old["photo"]


                cur.execute(
                    """
                    INSERT INTO vehicles
                    (
                        vin,
                        vehicle,
                        condition,
                        mileage,
                        your_price,
                        marketplace_price,
                        url,
                        photo,
                        status,
                        lead_count,
                        last_seen,
                        updated_at
                    )

                    VALUES
                    (
                        %s,%s,%s,%s,%s,%s,
                        %s,%s,%s,%s,%s,%s
                    )

                    ON CONFLICT (vin)

                    DO UPDATE SET

                        vehicle = EXCLUDED.vehicle,
                        condition = EXCLUDED.condition,
                        mileage = EXCLUDED.mileage,
                        your_price = EXCLUDED.your_price,
                        marketplace_price = EXCLUDED.marketplace_price,
                        url = EXCLUDED.url,
                        photo = EXCLUDED.photo,
                        last_seen = EXCLUDED.last_seen,
                        updated_at = EXCLUDED.updated_at
                    """,
                    (
                        v["vin"],
                        v["vehicle"],
                        v["condition"],
                        v["mileage"],
                        v["your_price"],
                        v["marketplace_price"],
                        v["url"],
                        v["photo"],
                        status,
                        leads,
                        now,
                        now
                    )
                )


            # Marcar vehículos que ya no aparecen
            # como unavailable, sin borrarlos.

            if found:

                vins = list(found.keys())

                cur.execute(
                    """
                    UPDATE vehicles
                    SET
                        status = 'unavailable',
                        updated_at = %s
                    WHERE NOT (vin = ANY(%s))
                    """,
                    (
                        now,
                        vins
                    )
                )

        c.commit()

    return len(found)


# ============================================================
# STARTUP
# ============================================================

@app.on_event("startup")
def startup():
    init_db()


# ============================================================
# DASHBOARD
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
def home(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="dashboard.html"
    )


# ============================================================
# INVENTORY API
# ============================================================

@app.get("/api/inventory")
def inventory():

    with db() as c:

        with c.cursor() as cur:

            cur.execute("""
                SELECT *
                FROM vehicles
                ORDER BY updated_at DESC
            """)

            rows = cur.fetchall()

    return rows


# ============================================================
# SUMMARY
# ============================================================

@app.get("/api/summary")
def summary():

    with db() as c:

        with c.cursor() as cur:

            cur.execute("""
                SELECT COUNT(*) AS n
                FROM vehicles
                WHERE status != 'unavailable'
            """)

            total = cur.fetchone()["n"]


            cur.execute("""
                SELECT COUNT(*) AS n
                FROM vehicles
                WHERE status != 'unavailable'
            """)

            available = cur.fetchone()["n"]


            cur.execute("""
                SELECT COUNT(*) AS n
                FROM vehicles
                WHERE status = 'published'
            """)

            published = cur.fetchone()["n"]


            cur.execute("""
                SELECT COUNT(*) AS n
                FROM vehicles
                WHERE status = 'pending'
            """)

            pending = cur.fetchone()["n"]


            cur.execute("""
                SELECT COUNT(*) AS n
                FROM vehicles
                WHERE status = 'error'
            """)

            errors = cur.fetchone()["n"]


            cur.execute("""
                SELECT COALESCE(
                    SUM(lead_count),
                    0
                ) AS n
                FROM vehicles
            """)

            leads = cur.fetchone()["n"]


    return {
        "vehicles": total,
        "available": available,
        "published": published,
        "pending": pending,
        "errors": errors,
        "total_leads": leads
    }


# ============================================================
# WEBSITE SYNC
# ============================================================

@app.post("/api/sync")
def sync():
    """
    La sincronización directa con southdadetoyota.com está desactivada
    porque el sitio rechaza las solicitudes provenientes de Render (403).
    El inventario se mantiene en Supabase y Vincue pasa a ser la fuente
    de actualización mediante /api/import-vincue.
    """
    try:
        with db() as c:
            with c.cursor() as cur:
                cur.execute("""
                    SELECT COUNT(*) AS n
                    FROM vehicles
                    WHERE status != 'unavailable'
                """)
                count = cur.fetchone()["n"]

        return {
            "ok": True,
            "count": count,
            "mode": "vincue",
            "message": (
                "La sincronización directa con el website está desactivada "
                "por el bloqueo 403. Usa Importar Vincue para actualizar "
                "el inventario. Los vehículos guardados en Supabase no se modificaron."
            )
        }

    except Exception as e:
        return JSONResponse(
            {
                "ok": False,
                "error": str(e)
            },
            status_code=500
        )


# ============================================================
# AUTO.DEV - FOTOS POR VIN
# ============================================================

def fetch_auto_dev_photos(vin: str):
    if not AUTO_DEV_API_KEY:
        return []
    try:
        r = requests.get(
            f"https://api.auto.dev/photos/{vin}",
            headers={"Authorization": f"Bearer {AUTO_DEV_API_KEY}", "Accept": "application/json"},
            timeout=15,
        )
        r.raise_for_status()
        retail = (r.json().get("data") or {}).get("retail") or []
        photos, seen = [], set()
        for url in retail:
            if isinstance(url, str) and url.startswith(("http://", "https://")) and url not in seen:
                seen.add(url)
                photos.append(url)
        return photos
    except Exception:
        return []

def fetch_photos_for_vins(vins, max_workers=8):
    results = {}
    vins = list(dict.fromkeys(vins))
    if not AUTO_DEV_API_KEY:
        return results
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        jobs = {executor.submit(fetch_auto_dev_photos, vin): vin for vin in vins}
        for job in as_completed(jobs):
            vin = jobs[job]
            try:
                results[vin] = job.result()
            except Exception:
                results[vin] = []
    return results

def save_vehicle_photos(cur, vin, photos, now):
    if not photos:
        return 0
    cur.execute("DELETE FROM vehicle_photos WHERE vin = %s", (vin,))
    for position, url in enumerate(photos, 1):
        cur.execute(
            """INSERT INTO vehicle_photos (vin, position, url, updated_at)
               VALUES (%s,%s,%s,%s)
               ON CONFLICT (vin, position)
               DO UPDATE SET url=EXCLUDED.url, updated_at=EXCLUDED.updated_at""",
            (vin, position, url, now),
        )
    cur.execute("UPDATE vehicles SET photo=%s, updated_at=%s WHERE vin=%s", (photos[0], now, vin))
    return len(photos)


# ============================================================
# FOTOS VINCUE
# ============================================================

def normalize_vincue_photo_url(value):
    """Normaliza una URL publica de imagen de VINCUE y elimina miniaturas sz160x."""
    if not value:
        return ""

    url = str(value).strip()
    if not url.startswith(("http://", "https://")):
        return ""

    # VINCUE envuelve algunas imagenes grandes con /image/opt-sz160x/.
    # Quitar solo esa transformacion conserva la URL publica de la foto grande.
    url = url.replace(
        "https://cdn-img.vincue.net/image/opt-sz160x/",
        "https://cdn-img.vincue.net/image/",
        1,
    )
    url = url.replace(
        "http://cdn-img.vincue.net/image/opt-sz160x/",
        "http://cdn-img.vincue.net/image/",
        1,
    )
    return url


def excel_cell_url(cell):
    """Devuelve URL desde valor o hyperlink de una celda Excel."""
    if cell is None:
        return ""

    if getattr(cell, "hyperlink", None) and cell.hyperlink.target:
        candidate = normalize_vincue_photo_url(cell.hyperlink.target)
        if candidate:
            return candidate

    return normalize_vincue_photo_url(cell.value)


# ============================================================
# IMPORTAR VINCUE
# ============================================================

@app.post("/api/import-vincue")
async def import_vincue(file: UploadFile = File(...)):
    """Importa el export completo de Used Inventory de VINCUE.

    VINCUE pasa a ser la fuente principal del inventario. El VIN es la llave.
    Se importan stock, millaje, precio y cantidad de fotos. Los leads de VINCUE
    NO se importan: lead_count queda reservado exclusivamente para Marketplace.
    """
    try:
        contents = await file.read()
        wb = load_workbook(BytesIO(contents), data_only=True)
        ws = wb.active

        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in ws[1]
        ]
        columns = {name: i for i, name in enumerate(headers)}

        required = ["VIN", "Year", "Model", "StockNo", "Odo", "Price"]
        missing = [name for name in required if name not in columns]
        if missing:
            return JSONResponse(
                {"ok": False, "error": "Faltan columnas: " + ", ".join(missing)},
                status_code=400,
            )

        photo_column = next(
            (
                name
                for name in [
                    "Photo", "PhotoURL", "PhotoUrl", "Photo URL",
                    "Image", "ImageURL", "ImageUrl", "Image URL",
                    "PrimaryPhoto", "Primary Photo", "PrimaryImage",
                    "Primary Image", "Picture", "PictureURL", "Picture URL",
                    "Thumbnail", "ThumbnailURL", "Thumbnail URL"
                ]
                if name in columns
            ),
            None,
        )

        imported = 0
        imported_vins = []
        now = datetime.now(timezone.utc).isoformat()

        with db() as c:
            with c.cursor() as cur:
                for cells in ws.iter_rows(min_row=2, values_only=False):
                    row = [cell.value for cell in cells]
                    vin = str(row[columns["VIN"]] or "").strip().upper()
                    if not vin:
                        continue

                    year = str(row[columns["Year"]] or "").strip()
                    model = str(row[columns["Model"]] or "").strip()
                    stock = str(row[columns["StockNo"]] or "").strip()
                    odo = row[columns["Odo"]]
                    raw_price = row[columns["Price"]]

                    try:
                        price = float(raw_price or 0)
                    except (TypeError, ValueError):
                        price = 0

                    marketplace_price = price + 1000 if price > 0 else 0
                    vehicle = f"{year} {model}".strip()

                    pic_count = 0
                    if "PicCount" in columns:
                        try:
                            pic_count = int(row[columns["PicCount"]] or 0)
                        except (TypeError, ValueError):
                            pic_count = 0


                    photo = ""
                    if photo_column:
                        photo = excel_cell_url(cells[columns[photo_column]])

                    # Mantener estado y leads de Marketplace existentes.
                    cur.execute(
                        "SELECT status, lead_count FROM vehicles WHERE vin = %s",
                        (vin,),
                    )
                    old = cur.fetchone()
                    status = old["status"] if old and old["status"] != "unavailable" else "pending"
                    leads = old["lead_count"] if old else 0

                    cur.execute(
                        """
                        INSERT INTO vehicles
                        (
                            vin, vehicle, condition, mileage,
                            your_price, marketplace_price, url, photo,
                            stock, pic_count, status, lead_count,
                            last_seen, updated_at
                        )
                        VALUES
                        (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (vin)
                        DO UPDATE SET
                            vehicle = EXCLUDED.vehicle,
                            condition = EXCLUDED.condition,
                            mileage = EXCLUDED.mileage,
                            your_price = EXCLUDED.your_price,
                            marketplace_price = EXCLUDED.marketplace_price,
                            stock = EXCLUDED.stock,
                            pic_count = EXCLUDED.pic_count,
                            status = EXCLUDED.status,
                            photo = CASE
                                WHEN EXCLUDED.photo IS NOT NULL AND EXCLUDED.photo <> ''
                                THEN EXCLUDED.photo
                                ELSE vehicles.photo
                            END,
                            last_seen = EXCLUDED.last_seen,
                            updated_at = EXCLUDED.updated_at
                        """,
                        (
                            vin, vehicle, "Used", str(odo or ""),
                            price, marketplace_price, "", photo,
                            stock, pic_count, status, leads,
                            now, now,
                        ),
                    )

                    imported += 1
                    imported_vins.append(vin)

                # El export completo de VINCUE define qué unidades siguen activas.
                if imported_vins:
                    cur.execute(
                        """
                        UPDATE vehicles
                        SET status = 'unavailable', updated_at = %s
                        WHERE NOT (vin = ANY(%s))
                        """,
                        (now, imported_vins),
                    )

            c.commit()

        return {
            "ok": True,
            "count": imported,
            "source": "VINCUE",
            "message": f"{imported} vehículos importados desde VINCUE"
        }

    except Exception as e:
        return JSONResponse(
            {"ok": False, "error": str(e)},
            status_code=500,
        )


# ============================================================
# PRUEBA AUTO.DEV - FOTOS POR VIN
# ============================================================

@app.get("/api/test-photos/{vin}")
def test_auto_dev_photos(vin: str):
    vin = (vin or "").strip().upper()
    if not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin):
        return JSONResponse({"ok": False, "error": "VIN inválido"}, status_code=400)
    if not AUTO_DEV_API_KEY:
        return JSONResponse({"ok": False, "error": "AUTO_DEV_API_KEY no está configurado en Render"}, status_code=500)
    try:
        response = requests.get(
            f"https://api.auto.dev/photos/{vin}",
            headers={"Authorization": f"Bearer {AUTO_DEV_API_KEY}", "Accept": "application/json"},
            timeout=TIMEOUT,
        )
        try:
            payload = response.json()
        except ValueError:
            payload = {"raw": response.text[:1000]}
        if response.status_code != 200:
            return JSONResponse(
                {"ok": False, "vin": vin, "auto_dev_status": response.status_code, "response": payload},
                status_code=response.status_code,
            )
        data = payload.get("data") or {} if isinstance(payload, dict) else {}
        retail = data.get("retail") or [] if isinstance(data, dict) else []
        photos = [u for u in retail if isinstance(u, str) and u.startswith(("http://", "https://"))]
        return {"ok": True, "vin": vin, "count": len(photos), "photos": photos}
    except requests.RequestException as exc:
        return JSONResponse({"ok": False, "vin": vin, "error": str(exc)}, status_code=502)


# ============================================================
# SINCRONIZAR FOTOS AUTO.DEV
# ============================================================

@app.get("/api/sync-photos")
def sync_auto_dev_photos():
    if not AUTO_DEV_API_KEY:
        return JSONResponse({"ok": False, "error": "AUTO_DEV_API_KEY no está configurado"}, status_code=500)

    with db() as c:
        with c.cursor() as cur:
            cur.execute("SELECT vin FROM vehicles WHERE status != 'unavailable' ORDER BY vin")
            vins = [row["vin"] for row in cur.fetchall()]

    results = fetch_photos_for_vins(vins)
    now = datetime.now(timezone.utc).isoformat()
    total_photos = 0
    with_photos = 0
    missing = []

    with db() as c:
        with c.cursor() as cur:
            for vin in vins:
                photos = results.get(vin, [])
                if photos:
                    total_photos += save_vehicle_photos(cur, vin, photos, now)
                    with_photos += 1
                else:
                    missing.append(vin)
        c.commit()

    return {
        "ok": True,
        "vehicles_checked": len(vins),
        "vehicles_with_photos": with_photos,
        "total_photos": total_photos,
        "vehicles_without_photos": len(missing),
        "missing_vins": missing
    }


# ============================================================
# PUBLICACIONES
# ============================================================

@app.get(
    "/publicaciones",
    response_class=HTMLResponse
)
def publicaciones(
    request: Request
):

    return templates.TemplateResponse(
        request=request,
        name="publicaciones.html"
    )


# ============================================================
# PHOTO PROXY - AUTO.DEV
# ============================================================

@app.get("/api/photo/{vin}/{position}")
def vehicle_photo(vin: str, position: int):
    vin = (vin or "").strip().upper()

    if not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin):
        return Response(status_code=404)

    if position < 1:
        return Response(status_code=404)

    with db() as c:
        with c.cursor() as cur:
            cur.execute(
                """
                SELECT url
                FROM vehicle_photos
                WHERE vin = %s AND position = %s
                """,
                (vin, position),
            )
            row = cur.fetchone()

    if not row or not row.get("url"):
        return Response(status_code=404)

    try:
        headers = {
            "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
            "User-Agent": HEADERS["User-Agent"],
        }

        if AUTO_DEV_API_KEY:
            headers["Authorization"] = f"Bearer {AUTO_DEV_API_KEY}"

        r = requests.get(
            row["url"],
            headers=headers,
            timeout=TIMEOUT,
        )
        r.raise_for_status()

        content_type = r.headers.get("content-type", "image/jpeg")
        if not content_type.lower().startswith("image/"):
            return Response(status_code=502)

        return Response(
            content=r.content,
            media_type=content_type.split(";")[0],
            headers={"Cache-Control": "public, max-age=86400"},
        )

    except requests.RequestException:
        return Response(status_code=502)


# ============================================================
# VEHICLE DETAIL API
# ============================================================

@app.get("/api/vehicle/{vin}")
def vehicle_detail(vin: str):

    with db() as c:

        with c.cursor() as cur:

            cur.execute(
                """
                SELECT *
                FROM vehicles
                WHERE vin = %s
                """,
                (vin,)
            )

            row = cur.fetchone()


    if not row:

        return JSONResponse(
            {
                "ok": False,
                "error":
                "Vehículo no encontrado"
            },
            status_code=404
        )


    with db() as c:
        with c.cursor() as cur:
            cur.execute(
                "SELECT position FROM vehicle_photos WHERE vin=%s ORDER BY position",
                (vin,),
            )
            photos = [
                f"/api/photo/{vin}/{item['position']}"
                for item in cur.fetchall()
            ]

    return {
        "ok": True,
        "vehicle": row,
        "photos": photos,
        "photo_count": len(photos)
    }


# ============================================================
# PREPARAR PUBLICACION
# ============================================================

@app.get(
    "/preparar/{vin}",
    response_class=HTMLResponse
)
def preparar_publicacion(
    request: Request,
    vin: str
):

    return templates.TemplateResponse(
        request=request,
        name="preparar.html"
    )
    # ============================================================
# INVENTARIO
# ============================================================

@app.get(
    "/inventario",
    response_class=HTMLResponse
)
def inventario_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="inventario.html"
    )


# ============================================================
# LEADS
# ============================================================

@app.get(
    "/leads",
    response_class=HTMLResponse
)
def leads_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="leads.html"
    )


# ============================================================
# ACTIVIDAD
# ============================================================

@app.get(
    "/actividad",
    response_class=HTMLResponse
)
def actividad_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="actividad.html"
    )


# ============================================================
# CONFIGURACION
# ============================================================

@app.get(
    "/configuracion",
    response_class=HTMLResponse
)
def configuracion_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="configuracion.html"
    )
