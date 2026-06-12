# Restaurant Time Clock

A self-contained kiosk web app for clocking employees in and out with a
personal PIN. Every punch writes a row to a Google Sheet in real time.
No subscriptions, no backend server — Google Apps Script (free) is the backend.

```
[Tablet kiosk]  →  [Google Apps Script web app]  →  [Google Sheet]
```

| File | What it is |
|------|------------|
| `index.html` | The entire kiosk app (HTML + CSS + JS, one file) |
| `Code.gs` | The Apps Script backend, pasted into your Google Sheet |
| `.env.example` | Template for `.env` — holds your `SCRIPT_URL` config |
| `.nojekyll` | Makes GitHub Pages serve the `.env` dotfile |
| `timeclock-plan.md` | Full design document |

---

## Setup (one time, ~15 minutes)

### 1. Create the Google Sheet

Create a new Google Sheet with **three tabs**, named exactly:

**`Punches`** — row 1 headers:

| Timestamp | Employee Name | Employee ID | Action | Date |
|---|---|---|---|---|

**`Employees`** — row 1 headers, row 2 is the reserved admin row:

| ID | Name | PIN Hash |
|---|---|---|
| 0 | \_\_admin\_\_ | *(paste admin PIN hash — see step 2)* |

> **The admin row must be exactly this shape — all three cells filled:**
> - **A2 must be the number `0`.** The backend recognizes the admin row
>   *only* by `ID = 0`. If you leave the ID blank, the admin PIN won't
>   work; if you use any other number, the row shows up on the kiosk as a
>   regular employee.
> - **B2** is just a label — `__admin__` or anything readable.
> - **C2** is the 64-character **hash** from step 2, not the PIN itself.
>   Paste it with no spaces before or after.
>
> Don't add regular employees by hand — the kiosk list stays empty until
> you add them through the in-app admin panel (step 6), which fills in
> IDs and hashes for you.

**`PayPeriod`** — config cells and headers:

- `A1`: `PAY_START`, `B1`: first day of the current pay period (e.g. `2026-06-08`)
- `A2`: `PAY_END`, `B2`: last day of the period (e.g. `2026-06-21`)
- Row 3 headers: `Employee Name | Total Hours | Regular Hours | Overtime Hours | Notes`

### 2. Generate your admin PIN hash

Open any browser, press F12 → Console, paste this with your chosen 4-digit PIN:

```js
crypto.subtle.digest("SHA-256", new TextEncoder().encode("1234"))
  .then(b => console.log(Array.from(new Uint8Array(b))
  .map(x => x.toString(16).padStart(2, "0")).join("")));
```

Copy the 64-character output into cell `C2` of the `Employees` sheet.
Raw PINs are never stored anywhere — only SHA-256 hashes.

### 3. Deploy the Apps Script

1. In the Sheet: **Extensions → Apps Script**
2. Delete the placeholder code, paste in the contents of `Code.gs`, save
3. **Project Settings (gear icon) → Time zone** — set to the restaurant's
   timezone. All punch timestamps use the **server clock in this timezone**,
   so a misconfigured tablet can't record wrong times.
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Authorize when prompted, then copy the **Web app URL**

> Updating the script later: use **Deploy → Manage deployments → ✏️ → New
> version** on the *existing* deployment. Creating a brand-new deployment
> changes the URL and breaks the frontend.

### 4. Add the pay period trigger

In Apps Script: **Triggers (clock icon) → Add Trigger**
- Function: `updatePayPeriod`
- Event source: Time-driven → Minutes timer → **Every 30 minutes**

### 5. Configure and host the frontend

1. Copy `.env.example` to `.env` and paste the Web app URL after
   `SCRIPT_URL=`. `index.html` itself never needs editing.
2. Host the files anywhere static — GitHub Pages is free: create a repo,
   upload `index.html`, `.env`, and `.nojekyll`, enable Pages in repo
   settings. (`.nojekyll` is required — without it GitHub Pages skips
   dotfiles and the app can't load `.env`.)
3. Open the page — you should see the (empty) home screen

> **What the `.env` file is (and isn't).** Browsers have no environment
> variables, so the app fetches `.env` from the host at runtime. Keeping the
> URL there (with `.env` in `.gitignore`) keeps it out of your committed
> source — useful if the repo is public. It is **not** hidden from visitors:
> anything the page can fetch, a visitor can fetch. That's fine for
> `SCRIPT_URL` (the browser needs it anyway, and all actions are PIN-gated
> server-side), but never put real secrets in it.

### 6. Add employees

Tap the **gear icon** (bottom-left) → enter your admin PIN → **Add employee**.
Each employee gets a name and their own 4-digit PIN. Done.

### 7. Tablet kiosk mode

- **Android**: install Fully Kiosk Browser (free), point it at your page URL,
  enable *Keep Screen On* and *Motion Detection / Wake*
- **iPad**: open the page in Safari, then enable **Guided Access**
  (Settings → Accessibility → Guided Access, triple-click to lock)

---

## Daily use

- **Clock out** (home screen): shows everyone currently on the clock with
  live elapsed time. Tap your name → enter PIN.
- **Clock in**: tap the orange **+ Clock In** button → tap your name → PIN.
- Wrong PIN: the dots shake; 3 misses in a row shows "see a manager".

## Admin panel (gear icon)

- **Employees** — see live status, reset anyone's PIN, delete (two-tap
  confirm; an open shift is auto-closed with an OUT punch so hours count)
- **Add employee** — name + PIN
- **Change admin PIN**
- **Purge punch history** — permanently deletes punches *before* a chosen
  date. Run only after payroll for those periods is finalized and exported.

## Payroll

The `PayPeriod` tab refreshes every 30 minutes with each employee's total,
regular, and overtime hours for the window in `B1`/`B2`.

- **Overtime** is computed per 7-day workweek anchored to `PAY_START`
  (federal FLSA rule), not on the period total. If your legal workweek
  starts on a fixed weekday, set `PAY_START` on that weekday.
- The **Notes** column flags anyone still on the clock (or who missed a
  clock-out). Fix a missed clock-out by adding a correcting `OUT` row
  directly in the `Punches` sheet.
- At period end: export/copy the `PayPeriod` data, then update `B1`/`B2`
  for the new period.

## Security model (short version)

- Only SHA-256 PIN hashes are stored; raw PINs never leave the device.
- Hashes are **never sent to the browser** — the server verifies every
  punch and every admin action against the `Employees` sheet.
- The deployment URL is unguessable but treat it like a password anyway —
  it lives in `.env` (gitignored) rather than in the page source, so it
  never lands in a public repo.

## Troubleshooting

**Admin PIN doesn't unlock the panel**
- Check cell `A2` in `Employees` is the number `0` — a blank or different
  ID is the usual cause.
- Check `C2` holds the 64-character hash (not the raw PIN), with no
  leading/trailing spaces.
- Sheet edits take effect immediately — no redeploy needed. Just retry.

**The admin shows up as an employee on the kiosk**
- Its ID isn't `0`. Set `A2` to `0` and it disappears from the kiosk.

**Kiosk shows no employees**
- Normal until you add them via the admin panel (gear icon → Add
  employee). The hand-made admin row is the only row that should ever be
  entered directly in the sheet.

## Known limitations (by design)

- No daily/California-style overtime (weekly 40h only)
- No offline queue — if wifi drops, the app shows "No connection" and the
  punch is not recorded
- No break tracking
- Employees can't change their own PIN (admin does it)
