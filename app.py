from fastapi.staticfiles import StaticFiles
import re, sqlite3, requests
from datetime import datetime, timezone
from urllib.parse import urljoin
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from bs4 import BeautifulSoup
from fastapi import UploadFile, File
from openpyxl import load_workbook
from io import BytesIO
DB=__import__("os").environ.get("DATABASE_PATH", "data/south_dade.db")
import os
os.makedirs(os.path.dirname(DB), exist_ok=True)
TIMEOUT=30
HEADERS={"User-Agent":"South-Dade-Toyota-Inventory-Manager/1.0"}
SOURCE="https://www.southdadetoyota.com/used-vehicles/"
app=FastAPI(title="South Dade Toyota Marketplace Manager")
app.mount("/static", StaticFiles(directory="static"), name="static")
templates=Jinja2Templates(directory="templates")

def db():
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    return c

def init_db():
    c=db()
    c.execute("""CREATE TABLE IF NOT EXISTS vehicles(
      vin TEXT PRIMARY KEY, vehicle TEXT, condition TEXT, mileage INTEGER,
      your_price REAL, marketplace_price REAL, url TEXT, photo TEXT,
      status TEXT DEFAULT 'pending', lead_count INTEGER DEFAULT 0,
      last_seen TEXT, updated_at TEXT
    )""")
    c.commit(); c.close()

def money(s):
    if not s: return None
    m=re.search(r'\$?\s*([\d,]+(?:\.\d{2})?)',s)
    return float(m.group(1).replace(",","")) if m else None

def collect():
    r=requests.get(SOURCE,headers=HEADERS,timeout=TIMEOUT)
    r.raise_for_status()
    soup=BeautifulSoup(r.text,"lxml")
    found={}
    # The public used inventory page exposes each used vehicle as a listing.
    for block in soup.find_all(["article","li","div"]):
        txt=" ".join(block.stripped_strings)
        vm=re.search(r'VIN:\s*([A-HJ-NPR-Z0-9]{17})',txt,re.I)
        if not vm: continue
        vin=vm.group(1).upper()
        if vin in found: continue
        pm=re.search(r'\$[\d,]+(?:\.\d{2})?',txt)
        mm=re.search(r'([\d,]+)\s*miles',txt,re.I)
        name=None
        for tag in block.find_all(["h2","h3","h4","a"]):
            t=" ".join(tag.stripped_strings)
            if re.search(r'\b20\d{2}\b',t) and len(t)<160:
                name=t; break
        if not name:
            name=re.search(r'\b20\d{2}\s+[A-Za-z0-9\- ]+(?:LE|SE|XLE|SR5|EX|LX|Limited|Sport|Premium|Hybrid)?\b',txt)
            name=name.group(0).strip() if name else "Used vehicle"
        href=None
        for a in block.find_all("a",href=True):
            h=urljoin(r.url,a["href"])
            if "southdadetoyota.com" in h and ("inventory" in h or "/used-" in h):
                href=h; break
        price=money(pm.group(0)) if pm else None
        found[vin]={
          "vin":vin,"vehicle":name,"condition":"used",
          "mileage":int(mm.group(1).replace(",","")) if mm else None,
          "your_price":price,
          "marketplace_price":price+1000 if price is not None else None,
          "url":href or SOURCE,
          "photo":None
        }

    now=datetime.now(timezone.utc).isoformat()
    c=db()
    for v in found.values():
        old=c.execute("SELECT status,lead_count FROM vehicles WHERE vin=?",(v["vin"],)).fetchone()
        status=old["status"] if old else "pending"
        leads=old["lead_count"] if old else 0
        c.execute("""INSERT INTO vehicles
          (vin,vehicle,condition,mileage,your_price,marketplace_price,url,photo,status,lead_count,last_seen,updated_at)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
          ON CONFLICT(vin) DO UPDATE SET
          vehicle=excluded.vehicle,condition=excluded.condition,mileage=excluded.mileage,
          your_price=excluded.your_price,marketplace_price=excluded.marketplace_price,
          url=excluded.url,last_seen=excluded.last_seen,updated_at=excluded.updated_at""",
          (v["vin"],v["vehicle"],v["condition"],v["mileage"],v["your_price"],
           v["marketplace_price"],v["url"],v["photo"],status,leads,now,now))
    # Vehicles no longer present are marked unavailable, not deleted.
    if found:
        qmarks=",".join("?"*len(found))
        c.execute(f"UPDATE vehicles SET status='unavailable',updated_at=? WHERE vin NOT IN ({qmarks})",
                  [now,*found.keys()])
    c.commit(); c.close()
    return len(found)

@app.on_event("startup")
def startup(): init_db()

@app.get("/",response_class=HTMLResponse)
def home(request:Request):
    return templates.TemplateResponse(request=request, name="dashboard.html")

@app.get("/api/inventory")
def inventory():
    c=db()
    rows=[dict(x) for x in c.execute("SELECT * FROM vehicles ORDER BY updated_at DESC")]
    c.close()
    return rows

@app.get("/api/summary")
def summary():
    c=db()
    total=c.execute("SELECT COUNT(*) n FROM vehicles WHERE status!='unavailable'").fetchone()["n"]
    available=c.execute("SELECT COUNT(*) n FROM vehicles WHERE status!='unavailable'").fetchone()["n"]
    published=c.execute("SELECT COUNT(*) n FROM vehicles WHERE status='published'").fetchone()["n"]
    pending=c.execute("SELECT COUNT(*) n FROM vehicles WHERE status='pending'").fetchone()["n"]
    errors=c.execute("SELECT COUNT(*) n FROM vehicles WHERE status='error'").fetchone()["n"]
    leads=c.execute("SELECT COALESCE(SUM(lead_count),0) n FROM vehicles").fetchone()["n"]
    c.close()
    return {"vehicles":total,"available":available,"published":published,"pending":pending,"errors":errors,"total_leads":leads}

@app.post("/api/sync")
def sync():
    try:
        count=collect()
        return {"ok":True,"count":count}
    except Exception as e:
        return JSONResponse({"ok":False,"error":str(e)},status_code=502)
@app.post("/api/import-vincue")
async def import_vincue(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        wb = load_workbook(BytesIO(contents), data_only=True)
        ws = wb.active

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
        columns = {name: i for i, name in enumerate(headers)}

        required = ["VIN", "Year", "Model", "StockNo", "Odo", "Price"]
        missing = [name for name in required if name not in columns]

        if missing:
            return JSONResponse(
                {"ok": False, "error": "Faltan columnas: " + ", ".join(missing)},
                status_code=400
            )

        c = db()
        imported = 0
        now = datetime.now(timezone.utc).isoformat()

        for row in ws.iter_rows(min_row=2, values_only=True):
            vin = str(row[columns["VIN"]] or "").strip()

            if not vin:
                continue

            year = str(row[columns["Year"]] or "").strip()
            model = str(row[columns["Model"]] or "").strip()
            stock = str(row[columns["StockNo"]] or "").strip()

            odo = row[columns["Odo"]]
            price = row[columns["Price"]]

            try:
                price = float(price or 0)
            except (TypeError, ValueError):
                price = 0

            marketplace_price = price + 1000 if price > 0 else 0
            vehicle = f"{year} {model}".strip()

            c.execute("""
                INSERT INTO vehicles
                (vin, vehicle, condition, mileage, your_price,
                 marketplace_price, url, photo, status,
                 lead_count, last_seen, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(vin) DO UPDATE SET
                    vehicle=excluded.vehicle,
                    condition=excluded.condition,
                    mileage=excluded.mileage,
                    your_price=excluded.your_price,
                    marketplace_price=excluded.marketplace_price,
                    last_seen=excluded.last_seen,
                    updated_at=excluded.updated_at
            """, (
                vin,
                vehicle,
                "Used",
                str(odo or ""),
                price,
                marketplace_price,
                "",
                "",
                "pending",
                0,
                now,
                now
            ))

            imported += 1

        c.commit()
        c.close()

        return {"ok": True, "count": imported}

    except Exception as e:
        return JSONResponse(
            {"ok": False, "error": str(e)},
            status_code=500
        )

@app.get("/publicaciones", response_class=HTMLResponse)
def publicaciones(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="publicaciones.html"
    )

@app.get("/api/vehicle/{vin}")
def vehicle_detail(vin: str):
    c = db()
    row = c.execute(
        "SELECT * FROM vehicles WHERE vin = ?",
        (vin,)
    ).fetchone()
    c.close()

    if not row:
        return JSONResponse(
            {"ok": False, "error": "Vehículo no encontrado"},
            status_code=404
        )

    return {"ok": True, "vehicle": dict(row)}
