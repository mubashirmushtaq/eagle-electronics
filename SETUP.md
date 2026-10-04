# Eagle Electronics — Repair Shop Manager · Setup Guide

Fully managed repair-shop platform. **Zero hosting cost:** frontend on GitHub Pages,
Google Sheet as the database, Apps Script as the API. Git keeps the code *and* the
backend script versioned; the sheet keeps the data.

Files in this folder:

| File         | What it is                                              |
|--------------|---------------------------------------------------------|
| `index.html` | The whole app (dashboard, job cards, delivery, reports, printing) |
| `Code.gs`    | Apps Script backend — paste into the Sheet's script project |
| `SETUP.md`   | This guide                                               |

---

## Step 1 — Create the Google Sheet (the database)

1. Create a new Google Sheet, name it **Eagle Electronics Repair Shop**.
2. Create 4 tabs with these **exact** names and paste these header rows into row 1:

**Tab `JobCards`**
```
JobNo | Date | CustomerName | Phone | Phone2 | CNIC | DeviceType | BrandModel | SerialNo | Fault | Accessories | EstCharges | Advance | AtDelivery | Balance | Status | Technician | ReceivedBy | ExpectedDate | ReadyDate | DeliveredDate | DeliveryProof | CNICPhotoURL | Notes
```

**Tab `Payments`**
```
PaymentID | Date | JobNo | Amount | Type | ReceivedBy | Note
```

**Tab `Users`**
```
Username | Password | Name | Role | Active
```
Add one seed row: `admin | changeme123 | Administrator | admin | TRUE`
(Change the password after first login via Users → you can add staff users there.)

**Tab `Settings`**
```
Key | Value
```
Seed rows:
```
ShopName     | Eagle Electronics
Address      | 382-6-B1, Barkat Chowk Behind BFC, Township Lahore
Phone1       | 0321-6596227
Phone2       | 0312-6596227
Services     | Sale & Services of Electro Medical Equipments, UPS, Microwave Oven, Stabilizer & LCD, LED.
Currency     | Rs
JobPrefix    | EE-
NextJobNo    | 1001
```

## Step 2 — Install the backend (Apps Script)

1. In the sheet: **Extensions → Apps Script**.
2. Delete the default `Code.gs` content, paste in this repo's `Code.gs`, save.
3. **Deploy → New deployment → Web app**:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone**
   - Deploy, authorize, and copy the Web App URL (ends in `/exec`).

> If you ever change the script code later, use **Deploy → Manage deployments → Edit → New version**
> so the URL stays the same.

## Step 3 — Point the app at the backend

In `index.html`, find this line near the top of the `<script>` block:

```js
const API_URL = "PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE";
```

Replace with your `/exec` URL.

## Step 4 — Publish (no hosting needed)

Option A — GitHub Pages (recommended, free):
1. Create a new GitHub repo (e.g. `eagle-electronics`), upload `index.html`, `Code.gs`, `SETUP.md`.
2. Repo **Settings → Pages → Deploy from branch → main → /**.
3. Open `https://<username>.github.io/eagle-electronics/` — done.

Option B — just open `index.html` on the shop PC (works, but no shared link).

## Step 5 — First run

1. Log in as `admin` / `changeme123`.
2. **Settings** → verify shop name/phones → **add staff users** if needed.
3. **New Job** → create the first job card → **Print job card** (this replaces the paper card).

---

## How the data is maintained

- **Git** versions the app + the Apps Script backend (`Code.gs`). Commit every change.
- **Sheet** holds all live data. Never edit `Balance` / `NextJobNo` by hand — the script maintains them.
- **Job #s** auto-increment (`EE-1001`, `EE-1002`, …) from `Settings → NextJobNo`.
- **CNIC photos** (taken when a customer loses their job card) are saved to a Google Drive
  folder `EagleElectronics_CNIC`; the shareable link is stored on the job row.
- **Backups:** the sheet's version history is your backup. For extra safety,
  periodically **File → Download → .xlsx**.

## Daily workflow

1. **New Job** — customer drops a device → fill form → print the job card, hand it over.
2. **Job Cards** — search by **phone number**, card #, name or device to trace a device.
3. Status flow: **Received → In Repair → Mark ready → Deliver**.
4. **Deliver** — customer returns the card → collect balance → print receipt.
   Card lost → choose *CNIC proof*, enter CNIC # and/or take a photo.
5. **Dashboard** — see what's ready for pickup, today's intake, revenue, outstanding balances.
6. **Reports** — filter by date/status, export to Excel.
