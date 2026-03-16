# sstramkak

Smart 5G Dashboard — a single-page web application that syncs data with Google Sheets via a Google Apps Script Web App.

---

## Google Sheets Sync — Google Apps Script

All data is pushed to / pulled from Google Sheets through a Google Apps Script (GAS) Web App.  
The script source lives in [`gas/Code.gs`](gas/Code.gs).

### Sheets managed

| Sheet name    | Key fields (selection)                                      |
|---------------|-------------------------------------------------------------|
| Sales         | id, name, phone, amount, agent, branch, date, …            |
| Customers     | id, name, phone, agent, branch, date, …                    |
| **TopUp**     | id, customerId, name, phone, amount, agent, branch, date, **endDate**, tariff, remark, tuStatus, lat, lng |
| Terminations  | id, customerId, name, phone, reason, agent, branch, date, … |
| OutCoverage   | id, name, phone, agent, branch, date, …                    |
| Promotions    | id, title, startDate, endDate, remark, …                   |
| Deposits      | id, amount, agent, branch, date, …                         |
| KPI           | id, agent, branch, date, …                                 |
| Items         | id, name, price, …                                         |
| Coverage      | id, lat, lng, …                                            |
| Staff         | id, username, password, role, status, …                    |

### TopUp — `endDate` column (expire date)

The TopUp form has an **Expiry Date** field (`tu-end-date`).  
The value is saved in the record as **`endDate`** and must be stored in the **`endDate`** column of the *TopUp* sheet.

The Apps Script (`gas/Code.gs`) automatically adds the `endDate` column if it is missing when the first sync is performed — no manual sheet editing is required.

---

## How to deploy / redeploy the Apps Script Web App

1. Open [Google Apps Script](https://script.google.com) and open the project linked to your Google Spreadsheet (or create a new standalone script and bind it to the sheet via the **Triggers** menu in the left sidebar).
2. Replace the contents of `Code.gs` with the file at [`gas/Code.gs`](gas/Code.gs) in this repository.
3. Click **Deploy → Manage deployments → New deployment** (or edit an existing one):
   - **Type:** Web App
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy** and copy the generated Web App URL.
5. If the URL has changed, update `GS_URL` in `app.js` (line ~79):
   ```js
   const GS_URL = 'https://script.google.com/macros/s/<deployment-id>/exec';
   ```
6. Save and reload the web app — sync should work immediately.

> **Note:** When updating an existing deployment, select **"New version"** from the version dropdown in the deployment editor for the changes to take effect.

---

## Sync overview

The web app communicates with the GAS Web App via `POST` requests:

```json
{ "sheet": "TopUp", "action": "sync", "data": [ { "id": "...", "endDate": "2025-12-31", ... } ] }
```

Supported `action` values:

| Action   | Description                                    |
|----------|------------------------------------------------|
| `sync`   | Replace all rows in the sheet with `data`      |
| `read`   | Return all rows as a JSON array of objects     |
| `delete` | Delete the row whose `id` matches `data.id`    |

The Apps Script automatically creates missing sheets and adds new header columns on the first sync, so it is **backward-compatible** with existing spreadsheet data.