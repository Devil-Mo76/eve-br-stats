# EVE Online BR Combat Statistics System · Detailed Operation Manual

> **English version** ｜ 中文版 → [操作手册.md](操作手册.md)
> Applies to: v1.1 Generic Edition
> Intended for: deployers (first-time users), daily operators, and developers extending the tool

---

## How to read this manual

| Who you are | Read this |
| --- | --- |
| First time here, just want to get it running | Part 1 → Part 2 Ch. 6–9 → Part 3 |
| Need to rebrand, reconfigure, change ports | Part 2 Ch. 10–11 |
| Want to build an EXE and share it | Part 4 |
| Something broke | Part 5 Ch. 21–22 |
| Want to modify the code | Part 6 |

Every `[MUST EDIT]` marker below corresponds to a real `TODO(通用版)` / `【需修改】`
comment in the source. Search for those strings to jump straight to the code.

---

## Table of Contents

**Part 1 · Understanding the System**
- [1. What this system is](#1-what-this-system-is)
- [2. Feature list](#2-feature-list)
- [3. How it works — data flow](#3-how-it-works--data-flow)
- [4. Directory structure](#4-directory-structure)

**Part 2 · Deployment and Configuration**
- [5. Requirements](#5-requirements)
- [6. Installation (from source)](#6-installation-from-source)
- [7. Starting and stopping](#7-starting-and-stopping)
- [8. Three deployment modes](#8-three-deployment-modes)
- [9. Mandatory first-time setup](#9-mandatory-first-time-setup)
- [10. Complete configuration reference (config.py)](#10-complete-configuration-reference-configpy)
- [11. What you must edit yourself](#11-what-you-must-edit-yourself)

**Part 3 · Feature Walkthroughs**
- [12. Combat statistics](#12-combat-statistics)
- [13. Attendance statistics](#13-attendance-statistics)
- [14. Admin pages](#14-admin-pages)
- [15. Report files explained](#15-report-files-explained)

**Part 4 · Packaging and Publishing**
- [16. Building an EXE](#16-building-an-exe)
- [17. Distribution and migration](#17-distribution-and-migration)
- [18. Publishing to GitHub](#18-publishing-to-github)

**Part 5 · Operations and Troubleshooting**
- [19. Backup and restore](#19-backup-and-restore)
- [20. Upgrade procedure](#20-upgrade-procedure)
- [21. Error message reference](#21-error-message-reference)
- [22. FAQ](#22-faq)

**Part 6 · Developer Guide**
- [23. Code structure and key functions](#23-code-structure-and-key-functions)
- [24. Extending the system](#24-extending-the-system)
- [25. Generic edition changelog](#25-generic-edition-changelog)

**Appendices**
- [A. Data file formats](#appendix-a-data-file-formats)
- [B. Configuration cheat sheet](#appendix-b-configuration-cheat-sheet)
- [C. Glossary](#appendix-c-glossary)

---

# Part 1 · Understanding the System

## 1. What this system is

**EVE Online BR Combat Statistics System** is a local web tool that runs on your own computer.

Its core job: **turn a list of BR (Battle Report) links into a damage leaderboard.**

Combat records in EVE Online live on third-party sites, and a single large engagement can
generate dozens or hundreds of BR links. Manually working out who contributed and who
did nothing is simply not practical. This tool automates that:

```
You paste a batch of BR links  →  the app fetches each combat record  →  keeps only members of the corporations you specified  →  produces an Excel report
```

**It is not**: an official EVE tool, a game modification, an account reader, or anything that
requires you to log in. It does exactly one thing — **aggregate public data**.

## 2. Feature list

| # | Feature | Where | Output |
| --- | --- | --- | --- |
| 1 | Combat statistics | Home page → "战斗统计分析" | Excel with charts (rank / pilot / sorties / kills / damage / ships) |
| 2 | Attendance statistics | Navbar → "出勤统计分析" | Attendance ranking + Excel (with a history sheet) |
| 3 | Corporation management | Manage → Corporation | Maintains *which corporations to track* |
| 4 | Character management | Manage → Character | Maintains character ID ↔ name mapping |
| 5 | Report history | Manage → History | View / download / delete past reports |
| 6 | System log | Manage → System log | Troubleshooting source |

> The web UI ships in Chinese. English-speaking users can still use every feature —
> this chapter names each menu so you can find it. If you want an English UI,
> see [24.6](#246-translating-the-ui-into-another-language).

## 3. How it works — data flow

### 3.1 Combat statistics pipeline

```
 ① You paste BR links
        │
        ▼
 ② Format validation (regex, checked on both the client and the server)
        │  invalid → red inline message, no requests fired
        ▼
 ③ GET https://br.evetools.org/api/v1/composition/get/{BR_ID}
        │  returns this fight's relateds → kms[] (list of killmails)
        ▼
 ④ For each killmail: GET https://kb.evetools.org/api/v1/killmails/{KILL_ID}
        │
        ├── Extract atts[] (attackers) and victim
        ├── Compare each corp (corporation ID) against data/corp_id.txt
        │      ├─ match    → accumulate sorties / damage for that pilot
        │      └─ no match → skip entirely (this is how "only our own" works)
        ├── If an attacker has blow (final blow) = true → that pilot's kills +1
        └── Unknown pilot encountered → auto-appended to data/char_id.txt
        ▼
 ⑤ generate_final_data() aggregates: sorties / kills / total damage / ships (deduplicated)
        ▼
 ⑥ openpyxl writes the workbook: data sheet + Damage Top10 chart + Kills Top10 chart
        ▼
 ⑦ Saved to data/reports/, and a "download latest report" button appears
```

**Key metrics**

| Metric | Meaning |
| --- | --- |
| Sorties | Number of times the pilot appears across the BRs you submitted (once per fight) |
| Kills | Number of **final blows**, not number of fights participated in |
| Total damage | Sum of damage dealt across all killmails |
| Ships used | Deduplicated list of ship names seen |

> ⚠️ **Important**: the scope is determined entirely by `data/corp_id.txt`.
> If it is empty, the app refuses and points you to the configuration page.
> This is **intentional** — it prevents generating a silently blank report that makes
> you think the tool is broken.

### 3.2 Attendance statistics pipeline

```
 ① Enter date + event type + member list (one name per line)
        │
        ▼
 ② Front-end fetch POST → /attendance/submit_fleet
        │
        ├── Deduplicate in order (a name repeated within one list counts once)
        └── Append to data/attendance.json
        ▼
 ③ Front-end immediately GETs /attendance/get_stats to refresh the live table
        ▼
 ④ Click "导出Excel" → /attendance/export_stats
        └── Workbook built in memory and streamed to the browser — no temp file on disk
```

> Attendance data is **persisted to disk**. Close the app, reopen it — your data is still there.

### 3.3 Third-party endpoints used

| Purpose | URL | Notes |
| --- | --- | --- |
| BR detail | `https://br.evetools.org/api/v1/composition/get/{id}` | Killmail list for a fight |
| Killmail detail | `https://kb.evetools.org/api/v1/killmails/{id}` | Attacker / victim breakdown |
| Ship data | `https://esi.evetech.net/latest` | Official ESI; used only by `ship_id.py` |

All of these live in `config.py` and can be swapped if an endpoint changes.

## 4. Directory structure

```
br-web/
│
├── app.py                      # Main app: routes + combat logic + Excel generation
├── config.py                   # ★ Single configuration entry point
├── attendance.py               # Attendance blueprint (with JSON persistence)
├── ship_id.py                  # Ship ID → name updater
├── build.py                    # PyInstaller build script
├── version_info.txt            # EXE version resource (company / copyright)
├── requirements.txt            # Dependency list
│
├── 操作手册.md                 # ★ Detailed manual (Chinese)
├── OPERATION_MANUAL_EN.md      # ★ This file (English)
├── README.md                   # Project overview
├── app_icon.ico                # EXE icon
│
├── tools/
│   ├── make_icon.py            # Generates favicon / EXE icon / PNG logo
│   └── smoke_test.py           # Smoke test (25 assertions, isolated data dir)
│
├── data/                       # Runtime data (see data/README.md)
│   ├── ship_id.txt             # Ship mapping (public data, 6000+ entries preloaded)
│   ├── corp_id.txt             # Corporation config ← you fill this in first
│   ├── char_id.txt             # Character mapping (written automatically)
│   ├── br_links.txt            # Last submitted BR links (saved automatically)
│   ├── attendance.json         # Attendance records (created automatically)
│   ├── reports/                # Generated Excel reports
│   └── README.md
│
├── static/
│   ├── site_logo.svg           # Navbar logo (neutral placeholder, replaceable)
│   ├── site_logo.png           # Same, PNG variant
│   ├── favicon.ico             # Browser tab icon
│   └── style.css
│
└── templates/                  # Jinja2 templates
    ├── base.html               # Layout: navbar, footer, background, styles
    ├── index.html              # Combat statistics home page
    ├── attendance.html         # Attendance page
    ├── docs.html               # Built-in help page
    ├── manage_corps.html       # Corporation management
    ├── manage_chars.html       # Character management
    ├── history.html            # Report history
    └── logs.html               # System log
```

---

# Part 2 · Deployment and Configuration

## 5. Requirements

| Item | Minimum | Recommended |
| --- | --- | --- |
| OS | Windows 10 / 11, Linux, macOS | Windows 11 |
| Python | 3.9 | 3.10 – 3.12 |
| RAM | 512 MB | 2 GB |
| Disk | 200 MB (incl. ship data) | 1 GB (room for reports) |
| Network | Reachable `br.evetools.org`, `kb.evetools.org` | Direct, **VPN/proxy off** |
| Browser | Chrome / Edge / Firefox | Chrome / Edge |

**About VPNs**: this matters. The BR endpoints rate-limit frequent requests, and going through
a proxy makes throttling or suspicious-traffic blocks much more likely. The upstream site
itself advises against using a VPN for BR statistics.

## 6. Installation (from source)

```bash
# 1) Enter the project directory
cd br-web

# 2) Create a virtual environment (recommended)
python -m venv .venv

# 3) Activate it
.venv\Scripts\activate           # Windows
# source .venv/bin/activate      # Linux / macOS

# 4) Install dependencies
pip install -r requirements.txt

# 5) Run
python app.py
```

If pip is slow, use a mirror (mainland China):

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

Dependencies:

| Package | Purpose | Required |
| --- | --- | --- |
| Flask | Web framework | ✅ |
| requests | Calls to third-party endpoints | ✅ |
| openpyxl | Excel + chart generation | ✅ |
| pandas | Attendance Excel export | ✅ |
| pyinstaller | Building the EXE | Build only |
| pillow | Generating icons | Icon work only |

## 7. Starting and stopping

```bash
python app.py
```

- The browser **opens automatically** (controlled by `config.AUTO_OPEN_BROWSER`);
- The console prints the data directory and the URL — **do not close this window**;
- To stop: press `Ctrl + C` in that window, or just close it.

**If the port is taken**, change `PORT` in `config.py` (5001 and 8000 both work).
To find what is holding 5000:

```bash
netstat -ano | findstr :5000     # Windows
lsof -i :5000                    # Linux / macOS
```

## 8. Three deployment modes

### Mode 1 — Local only (default, simplest)

```python
# config.py
HOST = "127.0.0.1"
PORT = 5000
```

Only you can reach it. No security considerations.

### Mode 2 — Shared on a LAN

```python
# config.py
HOST = "0.0.0.0"           # listen on all interfaces
PORT = 5000
SECRET_KEY = "<your own random string>"
```

Others browse to `http://<your-LAN-IP>:5000`.
Find your LAN IP with `ipconfig` on Windows (look for "IPv4 Address").
You also need to **allow the port through Windows Firewall** — usually just clicking
"Allow access" the first time Windows prompts about Python.

### Mode 3 — Public server (careful)

Only if you really need it, and you must:

1. Change `SECRET_KEY` to a random string;
2. Put it behind Nginx with HTTPS;
3. Understand that **this app has no authentication** — exposing it publicly means anyone
   can generate reports and read your history. If you need auth, see
   [24.4](#244-how-do-i-add-login-authentication).

## 9. Mandatory first-time setup

Once installed, **one step remains**: tell the app which corporation to track.

### 9.1 Configure via the web UI (recommended)

1. Open <http://127.0.0.1:5000>
2. Navbar → **管理 (Manage) → 军团管理 (Corporation)**
3. Fill in:
   - **军团ID (Corporation ID)** — digits only, e.g. `98769610`
   - **军团名称 (Corporation name)** — display only, anything you like
4. Click **添加军团 (Add)**

### 9.2 How to find your corporation ID

Open <https://zkillboard.com> → search your corporation → open its page and look at the URL:

```
https://zkillboard.com/corporation/98769610/
                                └────┬───┘
                              This is the corporation ID
```

You can also read it off the corporation page on <https://br.evetools.org>.

### 9.3 Editing the file directly (optional)

Edit `data/corp_id.txt`, one `ID_name` per line:

```
98769610_Example Corporation
12345678_Another Corporation
```

Blank lines and lines starting with `#` are ignored.

### 9.4 Verifying it worked

Go back to the home page and submit any BR link. If no corporation is configured, the page
shows a red message — *"尚未配置任何军团 ID，请先到「军团管理」页面添加你要统计的军团"*
("No corporation ID configured yet — add the corporation you want to track on the
Corporation page first"). Seeing that message means `corp_id.txt` is still empty.

## 10. Complete configuration reference (config.py)

Everything configurable lives in `config.py`. **Items marked ★ are the ones you will
normally need to change.**

### 10.1 Branding

| Variable | Default | Description |
| --- | --- | --- |
| ★ `SITE_NAME` | `EVE Online BR 战斗统计系统` | Used for the browser title, navbar, and footer |
| ★ `SITE_VERSION` | `1.1` | Rendered as "<name> v1.1" |
| ★ `FOOTER_TEXT` | same as `SITE_NAME` | Footer text; final output is "<text> ©<year>" |
| ★ `CORP_LOGO` | `site_logo.svg` | Navbar logo. Set to `""` to hide the logo entirely |
| `BACKGROUND_IMAGE` | `""` | Page background. Empty string = built-in deep-space gradient |

**Using your own logo**: drop the image into `static/`, then:

```python
CORP_LOGO = "my_logo.png"      # file name only, not a path
```

**Using your own background**:

```python
BACKGROUND_IMAGE = "background.jpg"
```

> ⚠️ Use artwork you actually have the rights to. Copying official game art carries
> copyright risk.

### 10.2 Web service

| Variable | Default | Description |
| --- | --- | --- |
| ★ `HOST` | `127.0.0.1` | Bind address; `0.0.0.0` makes it LAN-reachable |
| ★ `PORT` | `5000` | Port; change if taken |
| `AUTO_OPEN_BROWSER` | `True` | Open the browser on startup |
| ★ `SECRET_KEY` | placeholder | Flask session key. **Must be changed for public deployments** |

Generate a random key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

### 10.3 Data directory

| Variable | Default | Description |
| --- | --- | --- |
| `DATA_FOLDER` | `""` | Custom data directory; empty = default |

Resolution order in `get_data_folder()`:

```
1. Environment variable BR_DATA_DIR        ← highest priority
2. config.DATA_FOLDER (when non-empty)
3. Frozen EXE       → <dir of exe>/data
4. Running from source → <project root>/data
```

**Example**: keep data and reports inside a OneDrive folder for automatic backup:

```bash
# Windows (cmd, session-scoped)
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

### 10.4 Data source endpoints

| Variable | Description |
| --- | --- |
| `BR_API_BASE` | BR detail endpoint; `{}` is replaced with the BR ID |
| `KM_API_BASE` | Killmail endpoint; `{}` is replaced with the killmail ID |
| `ESI_BASE` | Official EVE ESI root, used by `ship_id.py` |
| `BR_LINK_PATTERN` | BR link validation regex, **shared by client and server** |
| `REQUEST_TIMEOUT` | Per-request timeout in seconds; raise to 20–30 on slow links |

`BR_LINK_PATTERN` is special: it is used by `re.match` in Python **and** injected into the
front-end JavaScript via `| tojson` as `new RegExp(...)`.
**Change it once and both sides stay in sync** — no drift between validators.

### 10.5 Report output

| Variable | Default | Description |
| --- | --- | --- |
| `REPORT_FILE_PREFIX` | `战斗报告` | Report file name prefix |
| `CHART_TOP_N` | `10` | How many entries the charts show |

Resulting file name looks like: `战斗报告_20250925_test.xlsx`

### 10.6 Attendance

| Variable | Default | Description |
| --- | --- | --- |
| ★ `EVENT_TYPES` | `["反收割", "小队活动", "会战", "集结待命"]` | Options in the "event type" dropdown |

Customise to match your own corporation's terminology:

```python
EVENT_TYPES = ["Counter-roam", "Small gang", "Fleet fight", "Staging"]
```

## 11. What you must edit yourself

**How to locate them** — run this from the project root:

```bash
grep -rn "TODO(通用版)\|【需修改】" --include="*.py" --include="*.html" --include="*.txt" .
```

> Line numbers below reflect the initial v1.1 release and may drift after edits —
> **always trust the search results over the numbers.**

### 11.1 Must change (affects correctness)

| # | Location | File | Line | What to do |
| --- | --- | --- | --- | --- |
| ① | Web UI → Corporation | — | — | Add your corporation ID and name, or nothing will be counted |
| ② | Branding | `config.py` | 24 | `SITE_NAME` → your system name |
| ③ | Branding | `config.py` | 30 | `FOOTER_TEXT` → your footer |
| ④ | Web service | `config.py` | 55 | `SECRET_KEY` → a random string (mandatory if public) |

### 11.2 Recommended

| # | File | Line | What to do |
| --- | --- | --- | --- |
| ⑤ | `config.py` | 34 | `CORP_LOGO` → your logo, or `""` to hide it |
| ⑥ | `config.py` | 38 | `BACKGROUND_IMAGE` → your background |
| ⑦ | `config.py` | 46 / 49 | `HOST` / `PORT` |
| ⑧ | `config.py` | 125 | `EVENT_TYPES` — attendance event types |
| ⑨ | `version_info.txt` | from line 4 | Company name / copyright shown in EXE properties |
| ⑩ | `build.py` | 27 / 31 / 34 / 37 | EXE name, icon, app name, version-info file |

### 11.3 Optional

| File | Line / location | Description |
| --- | --- | --- |
| `config.py` | 64 | `DATA_FOLDER` — custom data directory |
| `config.py` | 94 / 97 / 100 | Endpoint URLs |
| `config.py` | 103 | `BR_LINK_PATTERN` validation rule |
| `config.py` | 106 | `REQUEST_TIMEOUT` |
| `config.py` | 114 | `REPORT_FILE_PREFIX` |
| `config.py` | 117 | `CHART_TOP_N` |
| `attendance.py` | line 25, `STORE_FILENAME` | Attendance data file name |
| `templates/base.html` | around line 13 | CDN URLs (swap for local files on an isolated network) |
| `templates/attendance.html` | 24 | Event type dropdown |
| `tools/make_icon.py` | 22 | Icon colours |
| `app.py` | 316 | "No corporation configured" message text |
| `app.py` | 474 | Character management extension point (whitelists, etc.) |

### 11.4 Asset replacement cheat sheet

| Want to change | Steps |
| --- | --- |
| Navbar logo | Put image in `static/` → set `config.CORP_LOGO` to the file name |
| Page background | Put image in `static/` → set `config.BACKGROUND_IMAGE` |
| Favicon + EXE icon | Edit colours in `tools/make_icon.py` → `python tools/make_icon.py`; or replace `static/favicon.ico` and `app_icon.ico` directly |
| Attendance event types | Edit `config.EVENT_TYPES` |

---

# Part 3 · Feature Walkthroughs

## 12. Combat statistics

Entry point: the home page (navbar → "战斗统计分析").

### 12.1 Form fields

| Field | Required | Description |
| --- | --- | --- |
| 选择报告日期 — report date | ✅ | Used in the file name, format `YYYYMMDD` |
| 自定义报告名称 — custom name | ✅ | File name suffix, max 50 chars; illegal characters become `_` |
| 粘贴BR链接 — BR links | ✅ | One per line; must be `https://br.evetools.org/br/<ID>` |

### 12.2 Steps

1. Pick a date and type a name;
2. Paste BR links (one per line) — below the box a live counter shows
   "共 N 条链接，格式全部正确 / 其中 M 条无效" (N links, all valid / M invalid);
3. Click **生成报告 (Generate report)** — the progress bar advances through
   *initialising → analysing combat data → generating the report*, then the page reloads;
4. When the green banner *"BR报告生成完毕"* appears, click **下载最新报告 (Download latest report)**.

### 12.3 Counting rules

Only members of corporations listed under Corporation management are counted.
For each killmail:

- A friendly pilot appears in the **attacker** list → their ship and damage are recorded;
- A friendly pilot is the **victim** → a sortie and their ship are recorded;
- An attacker with `blow = true` (final blow) → that pilot's kill count +1;
- Within one fight, a pilot's sortie is counted once.

### 12.4 Output

See [Chapter 15](#15-report-files-explained).

## 13. Attendance statistics

### 13.1 Important note about the URL

Because of Flask blueprint prefix stacking (`url_prefix='/attendance'` plus
`@route('/attendance')` inside the blueprint), the attendance page's **real URL is**:

```
http://127.0.0.1:5000/attendance/attendance
```

**Opening `/attendance` directly returns 404. This is expected behaviour, not a bug.**
Normally you just click "出勤统计分析" in the navbar.

### 13.2 Form fields

| Field | Required | Description |
| --- | --- | --- |
| 报告日期 — report date | ✅ | Defaults to today |
| 报告类型 — event type | ✅ | Options from `config.EVENT_TYPES` |
| 舰队说明 — note | ❌ | Free text, e.g. "Hauling escort" |
| 角色列表 — member list | ✅ | One character name per line |

### 13.3 Steps

1. Pick a date and an event type;
2. Paste the members who joined (one per line);
3. Click **添加舰队角色 (Add fleet members)** → a modal reports "已成功添加 N 个角色",
   the textarea clears, and focus returns to it;
4. The right-hand table **refreshes live** with the attendance ranking; the footer shows the total;
5. Repeat 1–4 for the next fleet;
6. Click **导出Excel (Export Excel)** when you want an archive.

**Deduplication**: a name repeated inside a single submission counts once.
So "Pilot Alpha / Pilot Beta / Pilot Alpha" will report 2 members added.

### 13.4 Persistence

Attendance lives in `data/attendance.json` — **close and reopen the app, data is still there**.
On page load the app checks for existing data and, if found, enables the export button and
refreshes the table immediately.

## 14. Admin pages

Navbar → **管理 (Manage)** dropdown.

### 14.1 Corporation management

- **Purpose**: maintain the list of corporations to track
- **File**: `data/corp_id.txt`
- **Note**: empty by default in the Generic Edition; you must add at least one

Fields: corporation ID (digits) and corporation name (display only).
Existing entries are listed with a count and can be deleted individually (with confirmation).

### 14.2 Character management

- **Purpose**: supplement or correct the character ID ↔ name mapping
- **File**: `data/char_id.txt`

> **Usually no manual work needed.** During combat statistics, the app automatically appends
> every friendly character it encounters in the BRs into `char_id.txt`.
>
> You only need this page when a report shows `Unknown(1234567890)`, or when you want to
> pre-seed a name for someone who has not appeared yet.

### 14.3 Report history

- **Purpose**: view, download, and delete generated reports
- **Location**: `data/reports/`
- Sorted newest first; deletion asks for confirmation

### 14.4 System log

- **Purpose**: troubleshooting
- **File**: `data/system.log`
- Records every `SUCCESS` / `ERROR` / `INFO` entry with a timestamp

Example:

```
[2025-09-25 10:12:31.220000] SUCCESS: Report generated: 战斗报告_20250925_test.xlsx
[2025-09-25 10:15:02.100000] ERROR: Invalid BR link: https://br.evetools.org/br/xxx
```

A more verbose runtime log is written to `data/detailed.log` by Python's `logging` module.

## 15. Report files explained

### 15.1 Combat statistics report

File name: `战斗报告_<date>_<custom name>.xlsx`
Default location: `data/reports/`

**Sheet "战斗统计"**

| Col | Field | Description |
| --- | --- | --- |
| A | Rank | Descending by total damage |
| B | Pilot | Character name |
| C | Sorties | Times the pilot appeared in the combat records |
| D | Kills | Final blows |
| E | Total damage | Accumulated damage |
| F | Ships | Deduplicated ship list, separated by an ideographic comma |

**Charts**

- At `H2`: "总伤害排名前10" — Total damage top 10 (column chart)
- At `H18`: "击杀数排名前10" — Kills top 10 (column chart)

How many entries the charts show is controlled by `config.CHART_TOP_N`.

### 15.2 Attendance report

File name: `舰队出勤统计_<timestamp>.xlsx`
Downloaded to your browser's default download folder.

| Sheet | Contents |
| --- | --- |
| 出勤统计 | Character name, attendance count (descending) |
| 历史记录 | Submission time, event date, event type, member count, note, member list |

---

# Part 4 · Packaging and Publishing

## 16. Building an EXE

Useful when handing the tool to teammates who do not have Python.

```bash
# 1) Install build dependencies
pip install pyinstaller pillow

# 2) (Optional) regenerate the icon
python tools/make_icon.py

# 3) Build
python build.py
```

Output: `dist/BRStatisticsTool.exe`

Build configuration sits at the top of `build.py`:

| Constant | Description |
| --- | --- |
| `EXE_NAME` | Output file name |
| `ICON_FILE` | Icon file name (`.ico`) |
| `APP_NAME` | Application name shown in Windows properties (defaults to `config.SITE_NAME`) |
| `VERSION_FILE` | Version resource file (default `version_info.txt`) |

The build script automatically:

1. Cleans `build/`, `dist/`, and any stale `.spec` file;
2. Bundles `templates/`, `static/`, and `data/` into the EXE;
3. Declares hidden imports (`jinja2`, `werkzeug`, `flask`, `openpyxl`, `pandas`, …);
4. Applies the icon and version information.

**First launch of the EXE**: Windows SmartScreen may warn you — click
"More info → Run anyway".

> 💡 After packaging, **data is written to the `data/` folder next to the EXE**, not to a
> temp directory — so nothing is lost when you close the app. This was one of the bugs
> fixed in the Generic Edition.

## 17. Distribution and migration

When sharing, copy:

```
BRStatisticsTool.exe
data/                    # ship_id.txt alone is enough
```

**Migrating to a new machine**: copy the whole `data/` directory to the same relative
location on the new machine. Historical reports and attendance records travel with it.

## 18. Publishing to GitHub

The project ships with a `.gitignore` that already excludes:

- `build/`, `dist/`, `*.spec` (build artifacts)
- `data/corp_id.txt`, `data/char_id.txt`, `data/br_links.txt` (your data)
- `data/attendance.json`, `data/*.log`, `data/reports/*.xlsx`
- `.venv/`, `.env`, `*.pem`, `*.key`

Standard flow:

```bash
git init
git add .
git commit -m "feat: generic edition"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

### ⚠️ Token security

**Never put a token in code, scripts, `.git/config`, or a commit.**

If a token ever appeared in a chat log, terminal output, or a screenshot, immediately:

1. Open GitHub → `Settings → Developer settings → Personal access tokens`
2. Find the token → **Revoke**
3. Generate a new one with only the scopes you need (`Contents: Read and write`)

After pushing, verify nothing lingered in `.git/config`:

```bash
grep -c "github_pat\|ghp_" .git/config || echo "clean"
```

### If `git push` fails

Errors such as `schannel: server closed abruptly`, `SSL_ERROR_SYSCALL`, or a long hang
usually mean **the git endpoint on `github.com` is blocked by your network**
(while `api.github.com` still works). In that case push via the GitHub REST API —
see the bundled skill `github-push-via-api`.

---

# Part 5 · Operations and Troubleshooting

## 19. Backup and restore

### What to back up

| File | Priority | Notes |
| --- | --- | --- |
| `data/attendance.json` | 🔴 High | Attendance records — painful to re-enter |
| `data/corp_id.txt` | 🔴 High | Corporation configuration |
| `data/char_id.txt` | 🟡 Medium | Character mapping; can be rebuilt |
| `data/reports/*.xlsx` | 🟡 Medium | Historical reports |
| `data/ship_id.txt` | 🟢 Low | Public data; re-fetchable |

**Simplest approach**: copy the entire `data/` directory.

**Automatic backup**: point the data directory at a cloud-synced folder:

```bash
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

### Restore

Copy the backed-up `data/` back into place. Nothing else to do.

## 20. Upgrade procedure

1. Back up `data/` (previous chapter);
2. Overwrite everything **except `data/`** with the new version;
3. **If `config.py` changed upstream**, re-apply your own values;
4. Re-run `pip install -r requirements.txt` (dependencies may have changed);
5. Start it and run `python tools/smoke_test.py` to confirm everything works.

## 21. Error message reference

| Page message / console output | Cause | Fix |
| --- | --- | --- |
| `Invalid links: xxx` | BR link format invalid | Must be `https://br.evetools.org/br/<ID>` with no extra parameters |
| `Invalid BR link: xxx` | Format fine but endpoint returned non-200 | BR expired or endpoint temporarily down — retry |
| `Failed to access: xxx` | Same as above | Same as above |
| `尚未配置任何军团 ID，请先到「军团管理」页面添加…` | `corp_id.txt` is empty | Add a corporation under Corporation management |
| `Date and name are required` | Date or name missing | Both are mandatory |
| `Please enter valid BR links` | Links box empty | Enter at least one link |
| `Report generation failed: xxx` | Unexpected error during generation | Check the system log for details |
| `[提示] 数据文件 corp_id.txt 不存在，已按空数据处理` | First run; file not created yet | **Normal.** Created automatically once you configure a corporation |
| `Unknown(2118200607)` in a report | Character mapping missing | Auto-completed next run; or add it under Character management |
| `Unknown(12345)` for a ship | New ship not in the mapping | Run `python ship_id.py` |
| Requests keep timing out | Network / proxy issue | Turn off VPN; or raise `REQUEST_TIMEOUT` |
| 404 at `/attendance` | Blueprint prefix stacking | Use `/attendance/attendance` or click through the navbar |
| Port already in use | Another process holds 5000 | Change `config.PORT` |

## 22. FAQ

**Q1: Clicking attendance shows `/attendance` and a 404 page?**

Blueprint prefix stacking — the real URL is `/attendance/attendance`.
Clicking through the navbar always works. See [13.1](#131-important-note-about-the-url).

**Q2: Why does the very first report fail saying no corporation is configured?**

The Generic Edition ships with an empty corporation list. This is deliberate — otherwise it
would silently produce an all-blank report and you would assume the tool was broken.
Add one under Manage → Corporation.

**Q3: Can I track several corporations at once?**

Yes. Add each one under Corporation management. The scope is the union of all configured
corporations.

**Q4: Do I have to paste BR links one by one?**

No — paste them as multiple lines, one per line. The app processes them in sequence and
tells you exactly which link is malformed.

**Q5: Why do the damage numbers differ from in-game figures?**

This tool reports **damage recorded in the BRs you supplied**, not the in-game damage
statistics. Differences come from: only the links you provided are covered; kills only count
final blows; and within one fight a pilot's sortie counts once.

**Q6: Can attendance data be lost?**

No. It is stored in `data/attendance.json` and survives restarts. Back that file up.

**Q7: How do others on my LAN access it?**

Set `HOST = "0.0.0.0"` and have them browse to `http://<your-LAN-IP>:5000`.
Allow the port through the firewall. **Change `SECRET_KEY` before exposing it.**

**Q8: Can I host it on the public internet?**

Technically yes, but **the app has no authentication** — public exposure means anyone can
use it. If you must, add authentication first
(see [24.4](#244-how-do-i-add-login-authentication)) and put it behind HTTPS.

**Q9: A ship shows up as a number in the report — why?**

New ship not yet in the mapping. Run `python ship_id.py`; the script writes directly to
`data/ship_id.txt` by default.

**Q10: Can I replace the background and logo?**

Yes. Put the image in `static/`, then set `config.CORP_LOGO` / `config.BACKGROUND_IMAGE`.
Make sure you have the rights to the artwork you use.

**Q11: Does it run on macOS / Linux?**

Yes, it runs. Note that `build.py` produces a Windows EXE; on Linux/macOS it produces a
native binary for that platform instead.

**Q12: Can it run in the background without a console window?**

On Windows you can use `pythonw app.py` (no console window), or add the packaged EXE to
your Startup folder. Bear in mind you lose the live log output — for troubleshooting,
running it from a terminal is still more convenient.

---

# Part 6 · Developer Guide

## 23. Code structure and key functions

### 23.1 app.py

| Function / route | Purpose |
| --- | --- |
| `inject_globals()` | Injects branding and navigation into every template |
| `data_path(*parts)` | Builds paths inside the data directory (**always use this, never `__file__`**) |
| `load_ids(filename)` | Reads `ID_name` mappings; skips blank lines and `#` comments |
| `save_ids(filename, data)` | Writes mappings back |
| `resolve_report_path(filename)` | **Security check**: only allows real `.xlsx` files inside `data/reports` |
| `process_killmail(...)` | Parses one killmail and accumulates the numbers |
| `generate_final_data(...)` | Aggregates into the final report structure |
| `/` | Home page: submit BR links, generate a report |
| `/download-report/<filename>` | Download (with whitelist validation) |
| `/delete-report/<filename>` | Delete (with whitelist validation) |
| `/manage-chars`, `/manage-corps` | Character / corporation management |
| `/history`, `/logs`, `/docs` | History / log / help pages |

### 23.2 attendance.py

| Function | Purpose |
| --- | --- |
| `_store_path()` | Absolute path of the attendance data file |
| `_load_submissions()` | Reads all records (returns `[]` if the file is missing or corrupt) |
| `_save_submissions()` | Writes all records back |
| `_build_stats()` | Aggregates and sorts attendance counts by name |
| `/submit_fleet` | Submit a member list (**deduplicated on submit**) |
| `/get_stats`, `/get_history` | Live stats / submission history |
| `/export_stats` | Excel export via an in-memory `BytesIO` stream — nothing hits disk |

### 23.3 Three security-relevant implementations (**do not regress these**)

1. **Path traversal protection** — `resolve_report_path()`
   Accepts only file names that contain no path separator, end in `.xlsx`, and appear in the
   actual listing of `data/reports`.
   > ⚠️ **Do not switch to `werkzeug.secure_filename`** — it strips non-ASCII characters
   > entirely and would break Chinese report file names, making downloads fail.

2. **XSS protection** — `escapeHtml()` in `templates/attendance.html`
   Character names come from user input; always escape before writing to the DOM.

3. **Single source of truth for validation** — `config.BR_LINK_PATTERN` is injected into
   JavaScript via `| tojson`, so client and server can never disagree.

## 24. Extending the system

### 24.1 Adding a new configuration option

1. Add the variable to the right section of `config.py`, with a `【需修改】` marker or a
   clear comment;
2. If templates need it, expose it in `inject_globals()` in `app.py`;
3. Update Chapter 10 / 11 of `操作手册.md` **and** the matching chapter in this file.

### 24.2 Adding a new page

```python
@app.route('/my-page')
def my_page():
    return render_template('my_page.html', ...)
```

Inherit from `base.html` to get the navbar, footer, and styling for free:

```jinja
{% extends "base.html" %}
{% block content %}
  ...your content...
{% endblock %}
```

Register the nav entry in `inject_globals()` in `app.py`.

### 24.3 Changing what gets counted

For example, to count "any participation in a kill" instead of only final blows, edit the
final-blow block in `process_killmail()`:

```python
# Current behaviour: only blow=true increments the counter
for attacker in kb_data.get('atts', []):
    if attacker.get('blow', True):
        ...
        char_dir[char_str][0][0] += 1
        break
```

Drop the `blow` check and the `break` to count every participation — but judge for yourself
whether that metric is meaningful for your use case.

### 24.4 How do I add login authentication?

The original code left a comment about a login endpoint but never implemented it.
A reasonable approach:

1. `pip install flask-login`;
2. Add a `LoginManager` in `app.py`;
3. Keep a simple user store (a JSON file under `data/` is fine to start);
4. Decorate your routes with `@login_required`.

**Do not** hardcode usernames and passwords in the source — that is the first thing scanners
look for.

### 24.5 Always verify after changing code

```bash
python -m py_compile app.py config.py attendance.py    # syntax check
python tools/smoke_test.py                             # 25 smoke assertions
```

`smoke_test.py` points the data directory at a temporary folder, so it
**never touches your real `data/`** — run it as often as you like.

### 24.6 Translating the UI into another language

The templates are plain Jinja2 HTML. To localise:

1. Replace the Chinese strings in `templates/*.html`
   (the nav labels, buttons, and help text are the main ones);
2. Update `config.SITE_NAME`, `SITE_VERSION`, `FOOTER_TEXT`, and `EVENT_TYPES`;
3. Update `templates/docs.html` (the built-in help page).

Avoid touching the `url_for(...)` endpoint names — those are code identifiers, not display
strings, and renaming them breaks routing.

## 25. Generic edition changelog

### 25.1 Content removed from the original customised build

| Item | How it was handled |
| --- | --- |
| Corporation logo `static/corp_logo.png` | Deleted; replaced with the neutral `site_logo.svg` |
| Corporation icon `cosmic_wanderers_icon.ico` | Deleted; replaced with `app_icon.ico` / `favicon.ico` |
| Corporation background (2.3 MB of official art) | Deleted; replaced with a CSS deep-space gradient |
| Corporation name `Cosmic-Wanderers` | Replaced everywhere with the configurable `config.SITE_NAME` |
| Corporation ID | `data/corp_id.txt` cleared |
| 260 real character IDs / names | `data/char_id.txt` cleared |
| 3 real BR links | `data/br_links.txt` cleared |
| Combat and report logs | `system.log`, `detailed.log` deleted |
| Author credits | Removed |
| Leftovers `output.xlsx`, `__pycache__/` | Deleted |

### 25.2 Bugs fixed

| # | Problem | Fix |
| --- | --- | --- |
| 1 | Attendance lived only in memory and was lost on restart | Persisted to `data/attendance.json` |
| 2 | Excel export left an `output.xlsx` in the working directory | Switched to an in-memory `BytesIO` stream |
| 3 | Report endpoints were vulnerable to path traversal | Added a whitelist check |
| 4 | Attendance table injected character names into `innerHTML` (XSS) | Added `escapeHtml()` |
| 5 | Packaged EXE wrote data into a temp directory, losing it on restart | Data directory now resolves next to the EXE; `BR_DATA_DIR` supported |
| 6 | File I/O did not specify an encoding | Standardised on `encoding='utf-8'` |
| 7 | Client and server each had their own BR validation regex | Unified via `config.BR_LINK_PATTERN` |
| 8 | Four admin pages had no navigation entry | Added a "Manage" dropdown to the navbar |
| 9 | Form contents were lost after a validation error | Date, name, and links are now re-displayed; last links saved |
| 10 | Bootstrap loaded twice (5.1.3 + 5.3.0), Font Awesome duplicated | Consolidated to a single version on bootcdn |
| 11 | Empty report generated silently when no corporation was configured | Now fails with a clear message pointing to the config page |
| 12 | Duplicate names in one attendance list were double counted | Deduplicated on submit, order preserved |
| 13 | Delete button built `onclick` by string concatenation (broke on quotes) | Now uses `data-*` attributes plus `encodeURIComponent` |
| 14 | Empty lists rendered as a blank page | Added empty states and item counts |
| 15 | Hardcoded personal absolute paths | Replaced with `config.get_data_folder()` |
| 16 | Malformed markup in `docs.html` (nested `<h3>`, unclosed `<ol>`) | Rewritten as valid HTML |

### 25.3 Added

- `config.py` — single configuration entry point
- `README.md` — project overview
- `操作手册.md` — detailed manual (Chinese)
- `OPERATION_MANUAL_EN.md` — detailed manual (English, this file)
- `data/README.md` — data file format reference
- `tools/make_icon.py` — icon generator
- `tools/smoke_test.py` — smoke test (25 assertions)
- `.gitignore`, `requirements.txt`
- Neutral logo / favicon / EXE icon

---

# Appendices

## Appendix A: Data file formats

Convention across all files: **UTF-8, one record per line, ID and name separated by the
first underscore `_`.** Blank lines and lines starting with `#` are ignored.

### A.1 `corp_id.txt`

```
# corporationID_corporationName
98769610_Example Corporation
12345678_Another Corporation
```

### A.2 `char_id.txt`

```
2118200607_Character Name
2113436255_Another Pilot
```

### A.3 `ship_id.txt`

```
582_Bantam
583_Condor
584_Griffin
```

Public data; 6000+ entries preloaded. Refresh with `python ship_id.py` when new ships appear.

### A.4 `br_links.txt`

```
https://br.evetools.org/br/68d956f22213cf0012ace2af
https://br.evetools.org/br/68d6bab14ff6700012d935c5
```

Used to restore the last submitted links into the textarea.

### A.5 `attendance.json`

```json
[
  {
    "timestamp": "2025-09-25T10:12:31.220000",
    "date": "2025-09-25",
    "event_type": "Small gang",
    "note": "Hauling escort",
    "members": ["Pilot Alpha", "Pilot Beta"],
    "member_count": 2
  }
]
```

> Do not corrupt this file manually — a malformed file is treated as "no attendance data".

## Appendix B: Configuration cheat sheet

```python
# ---- Branding ----
SITE_NAME         = "EVE Online BR 战斗统计系统"
SITE_VERSION      = "1.1"
FOOTER_TEXT       = "EVE Online BR 战斗统计系统"
CORP_LOGO         = "site_logo.svg"        # "" = hide
BACKGROUND_IMAGE  = ""                     # "" = built-in gradient

# ---- Service ----
HOST              = "127.0.0.1"            # "0.0.0.0" = LAN-reachable
PORT              = 5000
AUTO_OPEN_BROWSER = True
SECRET_KEY        = "please-change-this-secret-key"   # must change if public

# ---- Data directory ----
DATA_FOLDER       = ""                     # "" = default; BR_DATA_DIR env var wins

# ---- Endpoints ----
BR_API_BASE       = "https://br.evetools.org/api/v1/composition/get/{}"
KM_API_BASE       = "https://kb.evetools.org/api/v1/killmails/{}"
ESI_BASE          = "https://esi.evetech.net/latest"
BR_LINK_PATTERN   = r"^https://br\.evetools\.org/br/[\w-]+$"
REQUEST_TIMEOUT   = 10

# ---- Reports ----
REPORT_FILE_PREFIX = "战斗报告"
CHART_TOP_N        = 10

# ---- Attendance ----
EVENT_TYPES = ["反收割", "小队活动", "会战", "集结待命"]
```

## Appendix C: Glossary

| Term | Full form | Meaning |
| --- | --- | --- |
| BR | Battle Report | A link recording the participants and outcome of a fight |
| KM | Killmail | The record generated for every kill in EVE |
| Corporation | — | The basic organisational unit in EVE |
| Alliance | — | A larger organisation composed of multiple corporations |
| Sortie | — | How many times a pilot appears in the combat records |
| Final blow | — | The last hit of a kill; marked `blow = true` in a BR |
| Ship | — | A flyable vehicle in EVE |
| ESI | EVE Swagger Interface | EVE's official public API |
| Blueprint | — | The Flask component that groups routes under a URL prefix |

---

## Quick start, in one minute

```
1. pip install -r requirements.txt
2. python app.py
3. Browser → Manage → Corporation → add "<corpID>_<corpName>"
4. Home → fill in date and name → paste BR links → Generate report → Download
5. Attendance → navbar "出勤统计分析" → enter members → Export Excel
```

---

> This document mirrors [操作手册.md](操作手册.md) chapter for chapter.
> If the two ever disagree, **the Chinese version is authoritative** — it is the
> primary maintained document.
