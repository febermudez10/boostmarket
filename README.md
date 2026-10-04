# BoostMarket

South Dade Toyota used-inventory and Marketplace management web app.

## Production deployment — Render

This project is prepared for Render.

Build command:
pip install -r requirements.txt

Start command:
uvicorn app:app --host 0.0.0.0 --port $PORT

A persistent disk should be mounted at:
`/var/data`

For production, configure the app's database path to `/var/data/south_dade.db`.

## Mobile

The dashboard is responsive and works from desktop, tablet and phone browsers.

## Important

Marketplace publication remains a separate integration and must use an authorized Meta/business mechanism. This project does not bypass Facebook security or access controls.
