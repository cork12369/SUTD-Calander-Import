# SUTD MyPortal Weekly Schedule → `.ics` Export (Console Script)

## What this does
Exports your **My Weekly Schedule** timetable into an **iCalendar (`.ics`)** file you can import into Google Calendar / Apple Calendar / Outlook.

--
- Open **Enrollment → My Weekly Schedule** (weekly grid view).
- Works best on desktop Chrome/Firefox.

---

## Quick Start
1. Go to **My Weekly Schedule** (weekly view).
2. Open Developer Tools:
   - **Chrome/Edge**: `F12` or `Ctrl+Shift+I` / `Cmd+Opt+I`
   - Click **Console**
3. Paste the full script (the one that installs `exportSUTDToICS`) and press **Enter**.
4. Run:


```js
exportSUTDToICS({ weeks: 14, direction: "forward", clickDelayMs: 2500 });
```

## Export Previous 14 Weeks (if needed)

```js
exportSUTDToICS({ weeks: 14, direction: "backward", clickDelayMs: 2500 });
```

---

## Settings Reference

| Option           | Meaning                            | Typical Values            |
| ---------------- | ---------------------------------- | ------------------------- |
| `weeks`          | number of weeks to export          | `1`, `14`                 |
| `direction`      | move weeks forward/back            | `"forward"`, `"backward"` |
| `clickDelayMs`   | wait after clicking next/prev week | `1800`–`3500`             |
| `tzid`           | timezone in ICS                    | `"Asia/Singapore"`        |
| `filenamePrefix` | downloaded filename prefix         | `"sutd_schedule"`         |

Example:

```js
exportSUTDToICS({
  weeks: 10,
  direction: "forward",
  clickDelayMs: 3000,
  filenamePrefix: "my_timetable"
});
```

---

## Troubleshooting

### 1) 14-week file downloads but has no events

This usually means the script clicked **Next Week** but scraped too early.

Fix:

```js
exportSUTDToICS({ weeks: 14, clickDelayMs: 3500 });
```

### 2) Script can’t find the timetable iframe / cross-origin error

Fix:

* Right-click inside the timetable area → **Open frame in new tab**
* Run the script again in that tab.


---

## Importing the `.ics`

* **Google Calendar**: Settings → Import & Export → Import `.ics`
* **Apple Calendar**: File → Import
* **Outlook**: File → Open & Export → Import/Export
