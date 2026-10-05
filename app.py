from fastapi.staticfiles import StaticFiles
import os
import re
import requests
from datetime import datetime, timezone
from urllib.parse import urljoin

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from io import BytesIO

import psycopg
from psycopg.rows import dict_row


# ============================================================
# CONFIGURACION
# ============================================================

DATABASE_URL = os.environ.get("DATABASE_URL")

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
                    status TEXT DEFAULT 'pending',
                    lead_count INTEGER DEFAULT 0,
                    last_seen TEXT,
                    updated_at TEXT
                )
            """)

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
# IMPORTAR VINCUE
# ============================================================

@app.post("/api/import-vincue")
async def import_vincue(
    file: UploadFile = File(...)
):

    try:

        contents = await file.read()

        wb = load_workbook(
            BytesIO(contents),
            data_only=True
        )

        ws = wb.active


        headers = [
            str(cell.value).strip()
            if cell.value is not None
            else ""
            for cell in ws[1]
        ]


        columns = {
            name: i
            for i, name
            in enumerate(headers)
        }


        # Vincue puede exportar la foto con distintos nombres.
        # Si existe una de estas columnas, BoostMarket la guardará.
        photo_column = next(
            (name for name in [
                "Photo", "PhotoURL", "PhotoUrl", "Image", "ImageURL",
                "ImageUrl", "PrimaryPhoto", "PrimaryImage", "Picture"
            ] if name in columns),
            None
        )


        required = [
            "VIN",
            "Year",
            "Model",
            "StockNo",
            "Odo",
            "Price"
        ]


        missing = [
            name
            for name in required
            if name not in columns
        ]


        if missing:

            return JSONResponse(
                {
                    "ok": False,
                    "error":
                    "Faltan columnas: "
                    + ", ".join(missing)
                },
                status_code=400
            )


        imported = 0

        now = datetime.now(
            timezone.utc
        ).isoformat()


        with db() as c:

            with c.cursor() as cur:

                for row in ws.iter_rows(
                    min_row=2,
                    values_only=True
                ):

                    vin = str(
                        row[columns["VIN"]]
                        or ""
                    ).strip()


                    if not vin:
                        continue


                    year = str(
                        row[columns["Year"]]
                        or ""
                    ).strip()


                    model = str(
                        row[columns["Model"]]
                        or ""
                    ).strip()


                    odo = row[
                        columns["Odo"]
                    ]


                    price = row[
                        columns["Price"]
                    ]


                    photo = ""
                    if photo_column:
                        photo = str(
                            row[columns[photo_column]] or ""
                        ).strip()


                    try:
                        price = float(
                            price or 0
                        )

                    except (
                        TypeError,
                        ValueError
                    ):
                        price = 0


                    marketplace_price = (
                        price + 1000
                        if price > 0
                        else 0
                    )


                    vehicle = (
                        f"{year} {model}"
                    ).strip()


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
                            photo = CASE
                                WHEN EXCLUDED.photo IS NOT NULL
                                     AND EXCLUDED.photo <> ''
                                THEN EXCLUDED.photo
                                ELSE vehicles.photo
                            END,
                            last_seen = EXCLUDED.last_seen,
                            updated_at = EXCLUDED.updated_at
                        """,
                        (
                            vin,
                            vehicle,
                            "Used",
                            str(odo or ""),
                            price,
                            marketplace_price,
                            "",
                            photo,
                            "pending",
                            0,
                            now,
                            now
                        )
                    )


                    imported += 1


            c.commit()


        return {
            "ok": True,
            "count": imported
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


    return {
        "ok": True,
        "vehicle": row
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
