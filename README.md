# KKP Fireworks Catalogue — Railway deployment

Mobile-friendly Python/Flask catalogue based on the supplied KKP Diwali 2026 price list. It includes all 139 products across 12 categories, product search, category filters, quantity selection, estimated totals, customer details, and PDF order receipts.

The home page footer says **Created by Aditya Maganbhai Patel**.

## Deploy to Railway

1. Extract this ZIP on your computer.
2. Create a GitHub repository and upload the **contents** of this folder (do not upload only the ZIP file).
3. Open https://railway.app and create a new project from that GitHub repository.
4. Wait for Railway to deploy the app. The included `railway.json` uses Gunicorn and configures `/health` as a health check.
5. In Railway, open your app service → **Settings / Variables** and add:
   - `ORDER_DATA_DIR` = `/app/data`
6. Open the service's **Volumes** section and add a Volume mounted at exactly:
   - `/app/data`
7. Generate a public domain for the service from Railway's networking/domain settings.
8. Open the generated URL and test by submitting a sample order. Confirm that the PDF downloads and the order is stored.

**Important:** create and mount the Volume at `/app/data` before relying on live orders. Without a persistent Volume, local files can be lost when the container is replaced.

## What is saved

- Each submitted order PDF is saved to `/app/data/orders/<ORDER_ID>_order.pdf`.
- Order history and customer/order details are stored in `/app/data/orders.sqlite3` (SQLite).
- The customer also receives a PDF download in their browser.
- The app does **not** create `order.json` files.

The PDF is an order request, not proof of stock availability or a completed sale. Confirm stock and final prices with the customer.

## Local test

Python 3.12 is recommended.

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000. By default, local order files are stored in `./data`. To test the same path as Railway, set `ORDER_DATA_DIR=/app/data` before starting the app.

## Files

- `app.py` — Flask routes, PDF generation, persistent order storage
- `products.json` — product catalogue and prices
- `templates/index.html` — page template
- `public/style.css`, `public/app.js` — design and browser behavior
- `railway.json`, `Procfile` — Railway/Gunicorn startup configuration

## Data and security notes

The SQLite database and PDFs are stored on the mounted Volume. This is suitable for a single Railway app instance; avoid multiple replicas writing to the same SQLite database. Back up important order data periodically. This app does not include an admin/login screen to browse stored orders; the PDF files and database remain on the Volume.


## Login page

The site now requires login before the catalogue and order API can be used. Initial usernames/passwords requested by the owner:

- `admin` / `123456` (admin role)
- `navsariyuvakmandalpurv` / `123456`
- `suratyuvakmandal` / `123456`
- `valsadyuvukmandal` / `123456`

Invalid credentials display an “Invalid credentials” message. The first version uses the admin role only as an identity label; it does not yet add a separate order-management dashboard.

### Configure Railway variables

Set `FLASK_SECRET_KEY` to a long random secret. You can override login credentials without editing code using `LOGIN_ADMIN_USER`, `LOGIN_ADMIN_PASSWORD`, `LOGIN_USER_1`, `LOGIN_PASSWORD_1`, `LOGIN_USER_2`, `LOGIN_PASSWORD_2`, `LOGIN_USER_3`, and `LOGIN_PASSWORD_3`.

**Security warning:** `123456` is weak and is only included because it was explicitly requested. Before publishing publicly, change all four passwords to unique, strong values in Railway Variables. Do not share the ZIP publicly while these default credentials remain active.


## Login credentials (change before public launch)

- Admin username: `admin`
- Admin password: `dhavalpatel`
- `navsariyuvakmandalpurv` password: `Navsari@Kkp26!`
- `suratyuvakmandal` password: `Surat#Kkp84!`
- `valsadyuvukmandal` password: `Valsad@Kkp73!`

If Railway environment variables `LOGIN_ADMIN_PASSWORD`, `LOGIN_PASSWORD_1`, `LOGIN_PASSWORD_2`, or `LOGIN_PASSWORD_3` are set, those values override the defaults in `app.py`.


## Order receipts and consolidated order archive
- Each new order gets a PDF receipt and an Excel (`.xlsx`) receipt in `data/orders/<username>/`.
- Copies of each order's PDF and Excel receipt are also stored in `data/all_orders/`.
- `data/all_orders/all_orders.xlsx` is regenerated with an **All Orders** summary sheet and an **Order Items** detail sheet.
- Admin can download the consolidated workbook at `/admin/all-orders.xlsx`.
- On Railway, keep the Volume mounted at `/app/data` and set `ORDER_DATA_DIR=/app/data`.
