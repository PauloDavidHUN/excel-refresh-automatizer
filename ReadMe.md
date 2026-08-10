# Excel Refresh Automatizer

Automatically opens Excel workbooks, refreshes all data connections and pivot tables,
saves the result, and closes -- fully unattended. Supports Box sync detection and
Windows Task Scheduler integration.

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Windows](https://img.shields.io/badge/OS-Windows-blue)
![pywin32](https://img.shields.io/badge/requires-pywin32-blue)

---

## What it does

| Feature | Description |
|---|---|
| Connection Refresh | Refreshes all OLEDB data connections with background query disabled for reliability |
| Pivot Table Refresh | Iterates every sheet and refreshes all pivot tables automatically |
| Box Sync Aware | Waits for Box file sync to complete before opening -- no corrupt reads |
| Task Scheduler Ready | Runs automatically on schedule via Windows Task Scheduler |
| Detailed Logging | Logs every step to console and \`refresh.log\` with timestamps |
| Easy Config | Add or remove Excel files in a single \`config.py\` -- no code changes needed |

---

## Requirements

- **Python 3.x** -- tested on 3.10+
- **Windows OS** -- uses Win32 COM to control Excel
- **Microsoft Excel** -- must be installed
- **pywin32** -- see \`requirements.txt\`

---

> **No Python yet?** If you don't have Python installed, check out my other project:
> (https://github.com/PauloDavidHUN/python-for-office-offline)
> -- a fully offline, portable Python environment ready to use.

## Quickstart

**1. Clone the repo**

\`\`\`bash
git clone https://github.com/PauloDavidHUN/excel-refresh-automatizer
\`\`\`

**2. Install dependencies**

\`\`\`bash
pip install -r requirements.txt
\`\`\`

**3. Configure your paths**

Rename \`config.example.py\` to \`config.py\` and update the paths to match your local environment.

**4. Run manually**

\`\`\`bash
python refresh_starter.py
\`\`\`

**5. Schedule it (optional)**

See the [Task Scheduler How-To](#task-scheduler-how-to) section below.

---

## Project Structure

\`\`\`
excel-refresh-automatizer/
│
├── config.py # ← gitignored, not tracked
├── config.example.py # ← template to copy from
├── refresh_starter.py # ← entry point, run this
├── refresh_helper.py # ← core logic
├── requirements.txt
├── .gitignore
├── HUN.txt # ← Hungarian description
├── ReadMe.html # ← visual README preview
└── README.md
\`\`\`

> **Note:** \`config.py\` is listed in \`.gitignore\` and will never be committed.
> Always use \`config.example.py\` as the reference.

---

## How it works

1. Reads the file list from \`config.py\`
2. Waits for Box cloud sync to finish (size-stability check)
3. Opens the workbook silently via Win32 COM
4. Disables background queries, then refreshes all connections
5. Waits for all background queries to finish
6. Refreshes every pivot table on every sheet
7. Saves and closes -- Excel exits cleanly

---

## Task Scheduler How-To

Open **Task Scheduler** and click **Create Task...** then fill in the tabs as follows:

### General tab

| Field | Value |
|---|---|
| Name | \`Excel Refresh Automatizer\` |
| Description | \`Auto-refreshes Excel workbooks daily\` |
| Security options | Run whether user is logged on or not + Run with highest privileges |

### Triggers tab

| Field | Value |
|---|---|
| Begin the task | On a schedule |
| Settings | Daily -- set your preferred start time (e.g. 06:00:00) |
| Recur every | 1 day |

### Actions tab

| Field | Value |
|---|---|
| Action | Start a program |
| Program/script | Full path to Python e.g. \`C:\\Python310\\python.exe\` (find it with \`where python\` in CMD) |
| Add arguments | \`refresh_starter.py\` |
| Start in | Full path to project folder e.g. \`C:\\Users\\YourName\\excel-refresh-automatizer\` |

### Conditions tab

| Field | Value |
|---|---|
| Power | Uncheck "Start only if on AC power" if running on a laptop |

### Settings tab

| Field | Value |
|---|---|
| If missed | Enable "Run task as soon as possible after a scheduled start is missed" |
| If already running | Select "Do not start a new instance" |

> **Tip:** To find your Python path, open Command Prompt and run \`where python\`.
> Copy the full path and paste it into the Program/script field.

> **Important:** The **Start in** field (Task Scheduler) must be set correctly -- without it,
> Python won't find \`config.py\` and the script will fail silently.
