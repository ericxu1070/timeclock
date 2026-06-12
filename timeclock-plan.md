# Restaurant Employee Time Clock — Build Plan

## What We're Building

A self-contained kiosk web app for a restaurant tablet that lets employees clock
in and out using a personal PIN. Every clock event writes a row to a Google Sheet
in real time. No subscriptions, no third-party services, no session expiry.

---

## Architecture

```
[Tablet in kiosk/fullscreen mode]
        |
        |  HTTP POST (clock event)
        v
[Google Apps Script Web App]   <-- free, deployed from your Google account
        |
        |  Sheets API (built-in, no extra auth)
        v
[Google Sheet]   <-- one row per clock event, owner sees all data live
```

**No backend server required.** The Google Apps Script acts as the backend.
The frontend is a single HTML file that can be hosted anywhere (GitHub Pages,
Netlify free tier, or even served from a local file on the tablet).

---

## Google Sheet Structure

One sheet named `Punches` with these columns:

| A | B | C | D | E |
|---|---|---|---|---|
| Timestamp | Employee Name | Employee ID | Action | Date |

- **Timestamp**: ISO datetime string of when the punch happened (from the tablet)
- **Employee Name**: full name string
- **Employee ID**: internal numeric ID matching the frontend config
- **Action**: `"IN"` or `"OUT"`
- **Date**: date-only string (YYYY-MM-DD) for easy filtering

A second sheet named `Employees` stores the employee roster:

| A | B | C |
|---|---|---|
| ID | Name | PIN Hash |

- **PIN Hash**: SHA-256 hash of the PIN (never store raw PINs)
- This sheet is **not** used by the app at runtime — it's just a reference.
  Employee data is baked into the frontend config (see below).

A third sheet named `PayPeriod` provides a live summary of hours worked by
each employee within the current pay period (see section below).

---

## Pay Period Summary Sheet

### How hours are calculated

The `Punches` sheet is a raw log — one row per punch, not one row per shift.
To get total hours per employee per pay period, a separate `PayPeriod` sheet
pairs each `IN` punch with the next `OUT` punch for the same employee and sums
the durations. This is done entirely with Google Sheets formulas — no extra
code required.

### Pay period configuration

Two named cells at the top of the `PayPeriod` sheet control the window:

| Cell | Name | Example Value |
|------|------|---------------|
| B1 | `PAY_START` | `2025-06-09` |
| B2 | `PAY_END`   | `2025-06-22` |

Update these two cells at the start of each new pay period. Everything else
recalculates automatically.

Common pay period lengths:
- **Weekly**: 7-day window, update every Monday
- **Bi-weekly**: 14-day window, update every other Monday
- **Semi-monthly**: 1st–15th and 16th–last day of month

### PayPeriod sheet layout

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| Employee Name | Total Hours | Total (h:m) | Regular Hours (≤40) | Regular (h:m) | Overtime Hours (>40) | Overtime (h:m) | Notes |

Each duration appears twice: a decimal-hours column (e.g. `38.50`, for
multiplying by wage) and a readable `Xh Ym` column (e.g. `38h 30m`).
Row 1–2 hold the `PAY_START` / `PAY_END` config cells. Data starts at row 4.

### How the Apps Script computes hours

Rather than using complex array formulas in Sheets (which get brittle with
paired IN/OUT logic), the `PayPeriod` sheet is populated by a second Apps
Script function that runs on a time-based trigger (e.g., every 30 minutes).

The function:

1. Reads all rows from `Punches` where the date falls within `PAY_START`–`PAY_END`
2. For each employee, walks their punches in chronological order, pairing each
   `IN` with the next `OUT` to compute shift duration in hours
3. Sums all shift durations per employee
4. Clears and rewrites the data rows in `PayPeriod`
5. Flags any unpaired `IN` (employee currently clocked in or forgot to clock out)

### Code outline for the hours calculation

```javascript
function updatePayPeriod() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const punches   = ss.getSheetByName("Punches");
  const summary   = ss.getSheetByName("PayPeriod");

  const payStart  = new Date(summary.getRange("B1").getValue());
  const payEnd    = new Date(summary.getRange("B2").getValue());
  payEnd.setHours(23, 59, 59); // include the full end day

  // Read all punch rows (skip header row 1)
  const rows = punches.getDataRange().getValues().slice(1);

  // Group punches by employee, filtered to pay period
  const byEmployee = {};
  for (const row of rows) {
    const [timestamp, name, id, action] = row;
    const ts = new Date(timestamp);
    if (ts < payStart || ts > payEnd) continue;
    if (!byEmployee[name]) byEmployee[name] = [];
    byEmployee[name].push({ ts, action });
  }

  // Pair IN/OUT and sum hours per employee, BUCKETED BY WORKWEEK.
  // Overtime under FLSA is computed per 7-day workweek, NOT on the
  // pay-period total — otherwise a bi-weekly period would wrongly flag
  // someone working 35h/week (70h, zero real OT) as having 30h of OT.
  // Each shift's hours are attributed to the workweek its IN falls in.
  // Workweeks are anchored to PAY_START (day 0, 7, 14, ...).
  const MS_PER_WEEK = 7 * 24 * 3600000;
  const results = [];
  for (const [name, punches] of Object.entries(byEmployee)) {
    punches.sort((a, b) => a.ts - b.ts);
    const weekHours = {}; // weekIndex -> hours
    let openIn = null;
    for (const punch of punches) {
      if (punch.action === "IN")  { openIn = punch.ts; }
      if (punch.action === "OUT" && openIn) {
        const hours = (punch.ts - openIn) / 3600000;
        const week  = Math.floor((openIn - payStart) / MS_PER_WEEK);
        weekHours[week] = (weekHours[week] || 0) + hours;
        openIn = null;
      }
    }
    let totalHours = 0, regularHours = 0, overtimeHours = 0;
    for (const h of Object.values(weekHours)) {
      totalHours    += h;
      regularHours  += Math.min(h, 40);
      overtimeHours += Math.max(0, h - 40);
    }
    results.push([name, totalHours.toFixed(2), regularHours.toFixed(2), overtimeHours.toFixed(2)]);
  }

  // Sort by name, write to PayPeriod sheet
  results.sort((a, b) => a[0].localeCompare(b[0]));
  const dataRange = summary.getRange(4, 1, Math.max(results.length, 1), 4);
  dataRange.clearContent();
  if (results.length > 0) dataRange.setValues(results);
}
```

This function is added to the same `Code.gs` file as `doPost`. Set it to run
on a time-based trigger: Apps Script → Triggers → Add trigger →
`updatePayPeriod` → Time-driven → Minutes timer → Every 30 minutes.

### What the PayPeriod sheet looks like (example)

| Employee Name | Total Hours | Total (h:m) | Regular Hours | Regular (h:m) | Overtime Hours | Overtime (h:m) | Notes |
|---|---|---|---|---|---|---|---|
| Alex Johnson | 38.50 | 38h 30m | 38.50 | 38h 30m | 0.00 | 0h 0m | |
| James Lee | 43.25 | 43h 15m | 40.00 | 40h 0m | 3.25 | 3h 15m | |
| Maria Garcia | 21.00 | 21h 0m | 21.00 | 21h 0m | 0.00 | 0h 0m | |

### Edge cases handled

- **Employee currently clocked in** (no matching OUT yet): their in-progress
  shift is excluded from the total until they clock out. Add a note column
  flagging this if desired.
- **Missed clock-out** (IN with no following OUT before the next IN): the
  orphaned IN is skipped. The owner should manually add a correcting OUT row
  to the `Punches` sheet.
- **Overtime threshold**: 40 hours **per workweek** (FLSA federal rule),
  computed by bucketing each shift into the 7-day workweek its clock-in falls
  in, anchored to `PAY_START`. This is correct for weekly and multi-week pay
  periods alike. If your state uses **daily** overtime rules (e.g., California:
  OT after 8 hours/day, double-time after 12), the calculation needs a further
  per-day breakdown — note this requirement when building.
- **Workweek anchor assumption**: weeks are counted from `PAY_START`. If your
  legal workweek starts on a fixed weekday (e.g., Sunday) that differs from
  `PAY_START`'s weekday, set `PAY_START` to land on that weekday.

---

## Google Apps Script (Backend)

A single `Code.gs` file deployed as a **Web App** (Execute as: Me, Access:
Anyone).

### What it does

Accepts `POST` requests with a JSON body:

```json
{
  "employeeId": 3,
  "employeeName": "Maria Garcia",
  "action": "IN",
  "timestamp": "2025-06-11T14:32:00",
  "date": "2025-06-11"
}
```

Appends one row to the `Punches` sheet and returns `{ "status": "ok" }`.

### Code outline

```javascript
function recordPunch(data) {
  // Serialize concurrent punches so two tablets/tabs can't interleave
  // appendRow and corrupt the sheet. Wait up to 10s for the lock.
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Punches");

    // Use the SERVER clock as the source of truth, not the tablet's clock,
    // so a misconfigured tablet can't write wrong times. The tablet's
    // timestamp is ignored for the stored value. Format in the restaurant's
    // timezone (set it in Apps Script: File → Project Settings → Time zone).
    const now  = new Date();
    const tz   = Session.getScriptTimeZone();
    const ts   = Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ss");
    const date = Utilities.formatDate(now, tz, "yyyy-MM-dd");

    sheet.appendRow([ts, data.employeeName, data.employeeId, data.action, date]);
    return jsonOut({ status: "ok", timestamp: ts });
  } finally {
    lock.releaseLock();
  }
}
```

The frontend still sends a `timestamp` for display optimism, but the stored
value comes from the server clock. The success response echoes the canonical
`timestamp` so the frontend can show the authoritative time.

### Deployment steps

1. Open your Google Sheet → Extensions → Apps Script
2. Paste the `Code.gs` content
3. Click Deploy → New deployment → Web app
4. Set **Execute as**: Me
5. Set **Who has access**: Anyone
6. Copy the deployment URL — this goes into the frontend config

> **Security note**: "Anyone" means anyone with the URL can reach the script.
> The URL is a long random string and is not publicly listed. Punches and
> admin actions are still gated by server-side PIN verification (see Security
> Model), so this is fine for a small restaurant. A shared secret key in each
> request adds another layer if desired.

> **CORS — important, the frontend will silently fail without this.** When the
> page is hosted on a different origin (GitHub Pages, Netlify) and POSTs to
> `script.google.com`, the browser sends a CORS preflight that Apps Script
> cannot answer with custom headers. Two rules avoid it:
> 1. **Send POST bodies as `Content-Type: text/plain`** (not
>    `application/json`). A `text/plain` body is a "simple request" and skips
>    preflight. The body is still a JSON string; `JSON.parse(e.postData.contents)`
>    on the server is unchanged.
> 2. **Do not set custom request headers** from the frontend.
>
> `GET_STATE` is a simple GET and works without special handling. Redeploying
> the script creates a **new URL** unless you choose "Manage deployments →
> edit → new version" on the existing deployment — always update the existing
> deployment so `SCRIPT_URL` stays stable.

---

## Frontend (Kiosk Web App)

A single `index.html` file. No framework, no build step, no dependencies.

---

### App State

On page load, the app fetches the current clock status for all employees from
the Apps Script (`GET` request, see backend section). This returns each
employee's `id`, `name`, and current status (`"IN"` / `"OUT"`) based on their
most recent punch. This state is held in memory and updated after every
successful punch — no full page reload needed.

Employee data (roster) is **no longer baked into the HTML file**. It is stored
in the `Employees` sheet and fetched on load alongside clock status. This is
what makes the in-browser admin panel possible without redeploying files.

**PIN hashes are never sent to the browser.** `GET_STATE` returns the roster
without hashes. PIN verification happens **server-side**: the frontend hashes
the entered PIN and sends the hash with the action; the Apps Script compares it
against the stored hash in the `Employees` sheet and only then records the
punch (or executes the admin action). This is essential — a 4-digit PIN has
only 10,000 possible values, so any hash exposed to the client can be reversed
instantly. See the Security Model section below.

---

### Screens / Views

There are three views rendered inside one page. Only one is visible at a time.

#### 1. Clock-Out View (default home screen)

Shows only employees who are **currently clocked in**. Each card displays:
- Employee name
- A green pulsing dot indicator (clocked in)
- How long they have been clocked in (e.g. "3h 14m"), updated live every minute

Tapping a card opens the PIN popup for that employee. A successful PIN clocks
them **out**.

A small `+` or "Clock In" button (bottom corner, unobtrusive) switches to the
Clock-In View.

#### 2. Clock-In View

Shows only employees who are **not currently clocked in**. Cards are visually
muted/neutral (no green dot). Tapping a card opens the PIN popup. A successful
PIN clocks them **in**. A back arrow returns to the Clock-Out View.

Rationale for two views: the default screen is the busy one during a shift
(people clocking out). Clock-in happens at the start of shift, so it's a
deliberate secondary action.

#### 3. Admin Panel

Accessed via a small discreet button (e.g. gear icon, bottom corner of home
screen). Opens a full-screen overlay that requires the admin PIN before showing
any content. Once authenticated, shows the employee management interface (see
below). A "Lock" button returns to the home screen and clears the admin session.

---

### PIN Popup (shared across all employee actions)

A single modal overlay used for all PIN entry. It is **not** a separate screen
— it slides or fades in over the current view without any navigation.

**Layout:**
- Dark translucent backdrop covering the full screen
- Centered card: employee name at the top, subtle subtitle showing action
  ("Clocking out" or "Clocking in")
- 4 large dot indicators showing how many digits have been entered
- 3×3 numpad grid + 0 + backspace, large tap targets (minimum 72px)
- No submit button — PIN auto-submits when the 4th digit is entered
- A small ✕ in the top-right corner to cancel

**Submit behavior:** when the 4th digit is entered, the frontend hashes the PIN
and POSTs the punch with the hash. The popup shows a brief loading state while
the server verifies. Success vs. wrong-PIN is decided by the server response
(`status: "ok"` vs. `error: "bad_pin"`), not by any client-side comparison —
the frontend never holds employee hashes.

**Success behavior** (on `status: "ok"`):
- The backdrop flashes green briefly (300ms)
- The card text updates to "See you, Maria!" (clock out) or "Welcome, Maria!"
  (clock in) with the current time
- Auto-dismisses after 1.5 seconds, returning to the home view
- Home screen updates immediately (employee card appears/disappears)

**Wrong PIN behavior** (on `error: "bad_pin"`):
- The dot indicators shake (CSS animation, ~400ms)
- Dots reset to empty — no message, no screen change, just clear and retry
- After 3 consecutive wrong attempts for the same employee: popup closes, a
  subtle toast notification appears at the bottom of the screen for 3 seconds:
  "Too many attempts — see a manager"

No separate error screen. No separate confirmation screen. Everything happens
inside the popup.

---

### Visual Design

Minimal, light-mode, sleek. White and grey surfaces with orange accents.
No sharp edges anywhere — generous border radii on every surface.

- Background: soft off-white (`#f4f4f6`), cards pure white with subtle shadow
- Text: near-black (`#1b1b1f`), secondary text muted grey (`#86868b`)
- Accent: orange (`#f97316`) for primary buttons, highlights, and active states
- Clocked-in indicator: small filled green circle with a soft pulse animation
  (green kept for status semantics; orange reserved for interactive accents)
- Duration badge: small rounded pill showing elapsed time (e.g. `3h 14m`)
- PIN popup backdrop: `rgba(0,0,0,0.35)` with blur; white card,
  `border-radius: 28px`, soft shadow
- Numpad buttons: large light-grey (`#ececf1`) keys, dark text,
  `border-radius: 18px`, press feedback via `transform: scale(0.96)` on
  `:active` with a ~120ms transition — buttons must feel responsive
- Fonts: system font stack (`-apple-system, BlinkMacSystemFont, "Segoe UI"`) —
  no external font load
- Transitions: all UI changes use CSS transitions ≤300ms (no janky flashes)

---

### Admin Panel

Accessed from the home screen via a small gear icon. Requires admin PIN to unlock.

The admin PIN is a separate hash stored in the `Employees` sheet in a reserved
row with `id: 0` and `name: "__admin__"`. **It is never sent to the browser.**
To unlock, the frontend hashes the entered admin PIN and sends it with a
`VERIFY_ADMIN` request; the server compares against the stored admin hash and
returns ok/unauthorized. On success the frontend keeps the admin PIN hash in
memory for the duration of the admin session and attaches it as `adminPinHash`
on every admin action (`ADD_EMPLOYEE`, `DELETE_EMPLOYEE`, `PURGE_PUNCHES`,
`CHANGE_ADMIN_PIN`). Each of those is re-verified server-side. The "Lock"
button clears the in-memory hash.

#### Employee list view

Shows all employees in a scrollable list. Each row shows:
- Employee name
- Current clock status (IN / OUT badge)
- A delete button (trash icon) with a confirmation tap to prevent accidents

#### Add employee form

A simple inline form at the bottom of the list:
- Name text input
- PIN input (4-digit, masked)
- Confirm PIN input
- "Add Employee" button

On submit:
1. Validates name is not empty and PINs match
2. Hashes the PIN in the browser
3. POSTs to the Apps Script with `action: "ADD_EMPLOYEE"`
4. Apps Script appends a row to the `Employees` sheet with the next available ID
5. Frontend refreshes its in-memory employee list

#### Delete employee

Tapping delete on a row:
1. Shows an inline "Are you sure?" confirm step on that row
2. If confirmed, POSTs to Apps Script with `action: "DELETE_EMPLOYEE"` and the
   employee ID
3. Apps Script removes that row from `Employees` and posts a final `OUT` punch
   if the employee is currently clocked in (to close their open shift)
4. Frontend removes the employee from the in-memory list

All historical punch rows for the deleted employee remain in the `Punches` sheet
permanently. Their hours will still appear in any `PayPeriod` calculation that
covers dates when they worked. Use the Purge tool (below) to remove old punch
data when it is no longer needed.

#### Purge punch history

A dedicated section in the admin panel, visually separated and styled in red to
signal destructive action.

**UI:**
- A date input labeled "Delete all punches before:"
- A "Purge" button, red, requires a second confirmation tap ("Tap again to
  confirm — this cannot be undone")

**Behavior:**
1. Admin selects a cutoff date
2. POSTs to Apps Script with `action: "PURGE_PUNCHES"` and the cutoff date
3. Apps Script deletes all rows from `Punches` where the `Date` column is
   strictly before the cutoff date
4. Returns a count of deleted rows; frontend shows a brief toast: "Deleted
   142 punch records"

**Suggested use:** Run at the end of each pay period after payroll is finalized
and the period data has been exported or copied. Never run mid-period.

**Apps Script logic:**

```javascript
function purgePunches(data) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Punches");
  const cutoff  = new Date(data.before);
  const rows    = sheet.getDataRange().getValues();
  const toDelete = [];

  // Collect rows to delete (walk backwards to preserve row indices)
  for (let i = rows.length - 1; i >= 1; i--) {
    const rowDate = new Date(rows[i][4]); // column E = Date
    if (rowDate < cutoff) toDelete.push(i + 1); // 1-indexed
  }

  toDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", deleted: toDelete.length }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

Walking backwards through the rows before deleting is critical — deleting rows
top-down shifts all subsequent row indices and causes the wrong rows to be
deleted.

#### Change admin PIN

A separate section at the bottom of the admin panel: current admin PIN input,
new PIN, confirm new PIN, submit. Same flow as adding an employee.

---

### PIN Hashing (unchanged)

```javascript
async function hashPin(pin) {
  const encoded = new TextEncoder().encode(pin);
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}
```

Runs in the browser before any data is sent to the backend. Raw PINs never
leave the device.

---

### Kiosk / Fullscreen Behavior

- On Android tablet: use **Fully Kiosk Browser** (free tier) pointed at the
  hosted URL. Enable "Keep Screen On" and "Motion Detection / Wake" so the
  screen reactivates when someone approaches.
- On iPad: use **Guided Access** (Settings → Accessibility → Guided Access) to
  lock Safari to the page.
- The app requests fullscreen via the Fullscreen API on first tap as a fallback.
- Page auto-reloads nightly at 3am (via `setTimeout`) to pick up any roster
  changes made in the admin panel or sheet directly.

---

## Apps Script — Additional Endpoints

The `Code.gs` file now handles four request types, dispatched on the `action`
field of the POST body (or a `?action=` query param for GET requests).

| Action | Method | Auth | Description |
|--------|--------|------|-------------|
| `GET_STATE` | GET | none | Return all employees + current clock status (no hashes) |
| `PUNCH` | POST | employee PIN hash | Record a clock IN or OUT punch |
| `VERIFY_ADMIN` | POST | admin PIN hash | Validate admin PIN to unlock the admin panel |
| `ADD_EMPLOYEE` | POST | admin PIN hash | Append a new employee row to `Employees` |
| `DELETE_EMPLOYEE` | POST | admin PIN hash | Remove employee row; close open shift if needed |
| `PURGE_PUNCHES` | POST | admin PIN hash | Delete all punch rows before a cutoff date |
| `CHANGE_ADMIN_PIN` | POST | admin PIN hash | Replace the stored admin PIN hash |

### GET_STATE response shape

```json
{
  "employees": [
    { "id": 1, "name": "Alex Johnson", "status": "OUT" },
    { "id": 2, "name": "Maria Garcia", "status": "IN",
      "clockedInAt": "2025-06-11T10:15:00" }
  ]
}
```

**No `pinHash` field is ever returned.** The admin row (`id: 0`) is filtered
out server-side and not included in the response at all. The `clockedInAt`
field is only present when `status` is `"IN"`. It is used by the frontend to
compute and display the live elapsed duration.

### Code outline for doGet / doPost dispatcher

```javascript
function doGet(e) {
  if (e.parameter.action === "GET_STATE") return getState();
  return jsonOut({ status: "error", error: "unknown action" });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  // PUNCH verifies the employee's own PIN hash; admin actions verify the
  // admin PIN hash. verifyPin / verifyAdmin read the stored hash from the
  // Employees sheet and compare — hashes are never trusted from GET_STATE.
  switch (data.action) {
    case "PUNCH":
      if (!verifyPin(data.employeeId, data.pinHash))
        return jsonOut({ status: "error", error: "bad_pin" });
      return recordPunch(data);

    case "VERIFY_ADMIN":
      return jsonOut({ status: verifyAdmin(data.adminPinHash) ? "ok" : "error" });

    case "ADD_EMPLOYEE":
      if (!verifyAdmin(data.adminPinHash)) return jsonOut({ status: "error", error: "unauthorized" });
      return addEmployee(data);

    case "DELETE_EMPLOYEE":
      if (!verifyAdmin(data.adminPinHash)) return jsonOut({ status: "error", error: "unauthorized" });
      return deleteEmployee(data);

    case "PURGE_PUNCHES":
      if (!verifyAdmin(data.adminPinHash)) return jsonOut({ status: "error", error: "unauthorized" });
      return purgePunches(data);

    case "CHANGE_ADMIN_PIN":
      if (!verifyAdmin(data.adminPinHash)) return jsonOut({ status: "error", error: "unauthorized" });
      return changeAdminPin(data);

    default:
      return jsonOut({ status: "error", error: "unknown action" });
  }
}

// Shared JSON-output helper used by every handler.
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

> **Why PIN verification is server-side:** the client sends only the SHA-256
> hash of the entered PIN. The Apps Script looks up the stored hash for that
> employee (or the admin row) and compares. Because hashes are never returned
> by `GET_STATE`, an attacker with the URL cannot read them, and because the
> destructive actions all gate on `verifyAdmin`, they cannot purge or modify
> data without the admin PIN. Note this is defense-in-depth, not perfect: a
> determined attacker who knows a single valid PIN hash could still replay it.
> For a small restaurant this is an acceptable trade-off; a shared secret key
> in the request adds another layer if desired.

### Verification helpers

```javascript
function getEmployeesSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employees");
}

// Compare the submitted hash against the stored hash for one employee id.
function verifyPin(employeeId, pinHash) {
  if (!pinHash) return false;
  const rows = getEmployeesSheet().getDataRange().getValues().slice(1);
  for (const [id, name, storedHash] of rows) {
    if (String(id) === String(employeeId)) return storedHash === pinHash;
  }
  return false;
}

// Admin is the reserved row with id 0 / name "__admin__".
function verifyAdmin(adminPinHash) {
  if (!adminPinHash) return false;
  const rows = getEmployeesSheet().getDataRange().getValues().slice(1);
  for (const [id, name, storedHash] of rows) {
    if (String(id) === "0") return storedHash === adminPinHash;
  }
  return false;
}

// Already gated by verifyAdmin in the dispatcher; just writes the new hash.
function changeAdminPin(data) {
  const sheet = getEmployeesSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === "0") {
      sheet.getRange(i + 1, 3).setValue(data.newAdminPinHash); // col C = PIN Hash
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", error: "admin_row_missing" });
}
```

`addEmployee`, `deleteEmployee`, and `purgePunches` follow the same pattern
(read the sheet, mutate, return `jsonOut(...)`); `purgePunches` is outlined in
the Admin Panel section above. Wrap `addEmployee` / `deleteEmployee` in the
same `LockService` pattern as `recordPunch` since they also append/delete rows.

---

## File Structure

```
timeclock/
├── index.html          # entire app (HTML + CSS + JS in one file)
├── Code.gs             # Apps Script backend (pasted into the Google Sheet)
├── .env                # SCRIPT_URL config — fetched by the app at runtime,
│                       #   gitignored (copy from .env.example)
├── .env.example        # template for .env
├── .nojekyll           # lets GitHub Pages serve the .env dotfile
└── README.md           # setup instructions
```

The admin panel, PIN popup, both home views, and all JS live inside
`index.html`. The app has no build step, so `SCRIPT_URL` is read from a
plain `.env` file fetched at runtime rather than from real environment
variables; the file is publicly served (not secret), it just keeps the
deployment URL out of committed source.

---

## Google Sheet Structure (updated)

### Employees sheet

| A  | B             | C                  |
|----|---------------|--------------------|
| ID | Name          | PIN Hash           |
| 0  | `__admin__`   | `<admin pin hash>` |
| 1  | Alex Johnson  | `<hash>`           |
| 2  | Maria Garcia  | `<hash>`           |

Row with `ID = 0` and `Name = __admin__` is the reserved admin PIN row. The
app filters it out of the employee display.

### Punches sheet (unchanged)

| Timestamp | Employee Name | Employee ID | Action | Date |

### PayPeriod sheet (unchanged)

Config cells `PAY_START` / `PAY_END` in B1/B2. Summary table from row 4.

---

## Setup Checklist

### One-time setup (owner)

- [ ] Create a new Google Sheet with three tabs: `Punches`, `Employees`, `PayPeriod`
- [ ] Add column headers to `Punches`: Timestamp, Name, ID, Action, Date
- [ ] Add column headers to `Employees`: ID, Name, PIN Hash
- [ ] Add the admin row to `Employees`: ID=0, Name=`__admin__`, PIN Hash=hash
      of your chosen admin PIN (generate in browser console)
- [ ] Add labels to `PayPeriod` rows 1–2 for `PAY_START` / `PAY_END` config
      cells, and headers in row 3: Employee Name, Total Hours, Total (h:m),
      Regular Hours, Regular (h:m), Overtime Hours, Overtime (h:m), Notes
- [ ] Open Apps Script from the sheet, paste `Code.gs`, deploy as web app,
      copy the deployment URL
- [ ] Set up a time-based trigger for `updatePayPeriod` (every 30 minutes)
- [ ] Edit `index.html`: paste the Apps Script URL into `SCRIPT_URL`
- [ ] Host `index.html` (GitHub Pages is free and takes 2 minutes to set up)
- [ ] Open the URL on the tablet, set up kiosk mode
- [ ] Use the in-app admin panel to add all employees and set their PINs

### Adding / removing employees after launch

Use the admin panel in the app. No file editing or redeployment needed.

### Changing a PIN after launch

Employees cannot change their own PIN. The admin changes PINs via the admin
panel. If an employee forgets their PIN, the admin can set a new one for them.

---

## Google Sheet Usage

The `Punches` sheet is the raw log — one row per clock event, never edited by
the owner except to correct missed clock-outs. The `PayPeriod` sheet
auto-updates every 30 minutes and shows totals per employee for the current pay
period. At the end of a pay period, copy the `PayPeriod` data to a new sheet or
export it as CSV for payroll, then update the `PAY_START` / `PAY_END` cells for
the new period.

---

## What Is NOT Included (by design)

- No daily overtime calculation (only weekly 40-hour threshold) — state-specific
  daily OT rules (e.g. California) require restructuring `updatePayPeriod` to
  compute per-day totals
- No separate manager permission level — admin PIN is the only elevated access
- No offline support — if wifi drops, punches fail silently (can be improved by
  queuing in memory and retrying on reconnect)
- No photo capture or GPS verification
- No break tracking (can be added as a third punch action: `"BREAK_START"`,
  `"BREAK_END"`)
- No employee self-service PIN change (admin only by design)

---

## Potential Improvements (post-MVP)

- Offline punch queue: store failed punches in memory, retry when connectivity
  resumes, show a "No connection" banner
- Break tracking: third button in the PIN popup success state
- Scheduled auto-reload at 3am to flush any stale state
- Daily overtime support for California-style rules
- Export button in admin panel that downloads current pay period as CSV
