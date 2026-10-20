# backlog-bulldozer

> Bulldoze your sprint planning spreadsheet into Backlog: automatically creates parent tasks (UCs), child tasks (Backend / Frontend), and fills in the fields for you.

**backlog-bulldozer** reads an Excel planning file and creates issues in Backlog through its API. Each UC becomes a parent task, and each row in the spreadsheet becomes a `[BE]` or `[FE]` child task, with assignee, priority, start date, due date, and estimated hours already filled in.

It can also read a Daily Report workbook, reuse or create one monthly issue, and add one Backlog comment per report date.

> **Status:** the Sprint importer has been used against the HRM Backlog project. The Daily Report workflow is new; always run `--check` and an API-backed dry-run before using `--run`.

## Features

- Reads the legacy `Backlog` layout and the Sprint 6 layout; use `--file` to select the workbook.
- By default **only creates tasks that are not Done** and skips finished ones. Add `--all` to create Done tasks as well.
- Each UC becomes one parent task containing its BE/FE child tasks.
- Matches the Assignee by name and warns if no match is found.
- Creates the Category and Milestone automatically if they do not exist.
- Saves created tasks to `created_issues.json`, so re-running does not create duplicates.
- **Dry-run by default**: nothing is created until you add `--run`.
- The API key is prompted at runtime (input is hidden), unless `BACKLOG_API_KEY` is already set. The script does not save the key.

## Getting started

### 1. Install

```bash
pip3 install openpyxl requests
```

### Chạy bằng menu đơn giản

Chỉ cần một lệnh:

```bash
python3 backlog.py
```

Sau đó chọn:

```text
1. Đẩy task mới lên Backlog
2. Đẩy Daily Report lên Backlog
```

Menu sẽ hỏi đường dẫn file Excel trên máy (có thể kéo-thả từ Finder hoặc dùng `file://`), tự kiểm tra file, yêu cầu xác nhận chạy thật và cho nhập API Key ở chế độ ẩn. Khi nhập task, Milestone được gợi ý từ tên file như `Sprint 6_Final.xlsx`; bạn có thể nhấn Enter để giữ hoặc nhập tên khác. Khi nhập Daily Report, menu hỏi riêng có thay thế các comment cũ do công cụ tạo hay không.

### 2. Get an API key

In Backlog: **Personal settings** -> **API** -> create a new key.

### 3. Check the Excel locally

```bash
python3 excel_to_backlog.py --file ~/Downloads/Sprint.xlsx --milestone "Sprint 6" --check
```

This validates the workbook and prints task, UC, status, and hour totals without asking for an API key or calling Backlog.

### 4. Dry-run one UC

```bash
python3 excel_to_backlog.py --file ~/Downloads/Sprint.xlsx --milestone "Sprint 6" --uc "UC-12.13"
```

The script asks for your API key, then prints the list of tasks it would create along with any warnings. **Nothing is created in Backlog.**

> The dry-run still calls the API to read project information and validate Assignee names, so an API key is still required.

### 5. Create that UC for real

```bash
python3 excel_to_backlog.py --file ~/Downloads/Sprint.xlsx --milestone "Sprint 6" --uc "UC-12.13" --run
```

After checking that UC in Backlog, remove `--uc "UC-12.13"` to process the remaining UCs. Keep `created_issues.json`; it prevents the pilot UC from being created again.

> **Common macOS (zsh) pitfall:** do not paste a trailing `# comment` after the command. zsh does not treat `#` as a comment by default, so the comment is passed to the script as arguments and you get `unrecognized arguments`. Paste only the command itself.

## Options

| Option | Description |
|---|---|
| `--file FILE` | Path to the Excel file (always pass this for Sprint 6) |
| `--space SPACE` | Backlog space, e.g. `d-soft-prj.backlog.com` |
| `--project PROJECT` | Project key, e.g. `HRM` |
| `--milestone MILESTONE` | Milestone name (default `Sprint 3`, created if missing) |
| `--all` | Also create tasks that are already Done (Done -> Closed) |
| `--uc UC` | Only process one exact UC, useful for a pilot run |
| `--check` | Validate the Excel locally without calling the Backlog API |
| `--run` | Create for real. Without this option the script only does a dry-run |

Example:

```bash
python3 excel_to_backlog.py --file ~/Downloads/Sprint_3.xlsx --milestone "Sprint 4" --run
```

## Field mapping

| Backlog field | Source in Excel |
|---|---|
| **Subject** | `[UC-4.7][BE] POST /projects/{id}/close` or `[UC-4.7][FE] screen name` |
| **Description** | Description, Actor, Endpoint/Method, original Priority, Est. SP, schedule including times |
| **Status** | Empty -> Open, In Progress -> In Progress, Done -> Closed (only with `--all`) |
| **Priority** | Critical and High -> High, Medium -> Normal, Low -> Low |
| **Assignee** | Matched by name (Tien, Bach, Tinh, ...) |
| **Category** | Module + Backend/Frontend |
| **Milestone** | `Sprint 3` (change with `--milestone`) |
| **Start Date / Due Date** | Start date / end date |
| **Estimated Hours** | `Estimate (giờ)` directly; legacy day-based files are converted with days x 8 |
| **Parent issue** | Each UC is a parent task containing its BE/FE tasks |

## Things to know

- **Fields left empty:** Version, Actual Hours, and Resolution. The Excel file has no data for them.
- **Times in the schedule:** Backlog only accepts dates, so times (08:30, 13:30, ...) are written into the Description.
- **Critical priority:** Backlog has no Critical level by default, so it is merged into High. The original priority is kept in the Description.
- **Parent tasks:** the Excel file has no UC name, so the Subject is built as `[UC] Module - screen name`, for example `[UC-4.7] Project Management - Confirm project closure`. The parent's dates and estimated hours are computed from its child tasks.
- **Unmatched Assignee:** if a name does not exist in the project, the script warns and leaves the Assignee empty.
- **Issue Type:** the script uses `Task`. If your project names it differently, the script fails with a list of available names.
- **Re-running:** to start over, delete `created_issues.json` (and delete the old issues in Backlog to avoid duplicates).

## Example dry-run result

With the sample Sprint 3 file, the script reads:

- 37 tasks that are not Done, across 17 UCs (36 with an empty status and 1 In Progress)
- 58 Done tasks skipped (add `--all` to create them too)

## Security

- Never commit your API key, state files, or unreviewed Excel exports. The Daily Report workbook in `reports/` is intentionally versioned for this workflow.
- Suggested `.gitignore`:

```
created_issues.json
daily_report_state.json
daily_report_comment_backup.json
*.xlsx
!reports/Daily_Report_Bach_2026-10-01_19.xlsx
.env
__pycache__/
```

## Requirements

- Python 3.8+
- A Backlog account with permission to add issues, categories, and milestones in the project
- Libraries: `openpyxl`, `requests`

> The `NotOpenSSLWarning` on macOS (Python built against LibreSSL) is only a version warning and does not affect execution.

## Daily Report

The included workbook contains the Backlog issue fields shown in the HRM Daily Report and one row for every date from 1–19 October 2026:

```text
reports/Daily_Report_Bach_2026-10-01_19.xlsx
```

It uses three sheets:

- `Backlog Setup`: Project, Summary, Issue Type, Status, Priority, Assignee, Category, Milestone, Version, Start Date, Due Date, Estimated Hours, Actual Hours, Resolution, and Existing Issue Key.
- `Daily Reports`: one or more task rows per date with Task, Details, Backlog ID, Duration Hours, Progress, Delay / Issues, Git evidence, an optional Backlog comment override, and posted issue key. The supplied workbook covers all 19 dates at 7 hours per day; dates without commits use explicit progress allocations.
- `Git Activity`: complete commit/MR evidence used to prepare the daily summaries; the importer treats this sheet as read-only supporting data.

The supplied workbook targets the existing monthly issue `HRM-319`. Remove `Existing Issue Key` or pass a different `--issue` only when you intentionally want another issue.

### Check locally

```bash
python3 daily_report_to_backlog.py \
  --file reports/Daily_Report_Bach_2026-10-01_19.xlsx \
  --check
```

This validates the two required import sheets without asking for an API key or calling Backlog; `Git Activity` remains read-only evidence.

### API-backed dry-run

```bash
python3 daily_report_to_backlog.py \
  --file reports/Daily_Report_Bach_2026-10-01_19.xlsx
```

The dry-run reads the project configuration and existing comments but does not create or update anything.

To review only one date:

```bash
python3 daily_report_to_backlog.py \
  --file reports/Daily_Report_Bach_2026-10-01_19.xlsx \
  --date 2026-10-14
```

### Replace the old generated comments and post the new format

Run an API-backed preview first. It lists the generated comments that would be removed, but does not change Backlog:

```bash
python3 daily_report_to_backlog.py \
  --file reports/Daily_Report_Bach_2026-10-01_19.xlsx \
  --issue HRM-319 \
  --replace-managed-comments
```

After reviewing the preview, run the replacement:

```bash
python3 daily_report_to_backlog.py \
  --file reports/Daily_Report_Bach_2026-10-01_19.xlsx \
  --issue HRM-319 \
  --replace-managed-comments \
  --run
```

Replacement only deletes comments carrying this tool's Daily Report marker within the dates selected from the workbook. It does not delete `HRM-319` or manual comments. Before deletion, it writes their full contents to `daily_report_comment_backup.json`; then it posts one comment per date in the Task / Details / Backlog ID / Duration / Progress format, followed by Finish and Delay/Issues sections.

The script resolves Issue Type, Priority, Assignee, Status, Milestone, and Version by exact name. It writes `daily_report_state.json` after every successful comment and embeds a date/content marker in each comment so rerunning does not post duplicates.

Backlog records the real time when a comment is posted. It cannot backdate comments through the API, so the historical report date is included in each comment heading.

API references: [Add Issue](https://developer.nulab.com/docs/backlog/api/2/add-issue/), [Update Issue](https://developer.nulab.com/docs/backlog/api/2/update-issue/), and [Add Comment](https://developer.nulab.com/docs/backlog/api/2/add-comment/).

## License

MIT
