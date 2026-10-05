# backlog-bulldozer

> Bulldoze your sprint planning spreadsheet into Backlog: automatically creates parent tasks (UCs), child tasks (Backend / Frontend), and fills in the fields for you.

**backlog-bulldozer** reads an Excel planning file and creates issues in Backlog through its API. Each UC becomes a parent task, and each row in the spreadsheet becomes a `[BE]` or `[FE]` child task, with assignee, priority, start date, due date, and estimated hours already filled in.

> **Status:** the Excel parsing works correctly. The task creation part has **not yet been tested against a real Backlog space** (no API key was available at the time of writing). Always run a dry-run first.

## Features

- Reads the Excel file from `~/Downloads`, or from a path given with `--file`.
- By default **only creates tasks that are not Done** and skips finished ones. Add `--all` to create Done tasks as well.
- Each UC becomes one parent task containing its BE/FE child tasks.
- Matches the Assignee by name and warns if no match is found.
- Creates the Category and Milestone automatically if they do not exist.
- Saves created tasks to `created_issues.json`, so re-running does not create duplicates.
- **Dry-run by default**: nothing is created until you add `--run`.
- The API key is prompted at runtime (input is hidden), so it is never stored in a file or environment variable.

## Getting started

### 1. Install

```bash
pip3 install openpyxl requests
```

### 2. Get an API key

In Backlog: **Personal settings** -> **API** -> create a new key.

### 3. Dry-run

```bash
python3 excel_to_backlog.py
```

The script asks for your API key, then prints the list of tasks it would create along with any warnings. **Nothing is created in Backlog.**

> The dry-run still calls the API to read project information and validate Assignee names, so an API key is still required.

### 4. Create for real

```bash
python3 excel_to_backlog.py --run
```

> **Common macOS (zsh) pitfall:** do not paste a trailing `# comment` after the command. zsh does not treat `#` as a comment by default, so the comment is passed to the script as arguments and you get `unrecognized arguments`. Paste only the command itself.

## Options

| Option | Description |
|---|---|
| `--file FILE` | Path to the Excel file (defaults to auto-detecting one in `~/Downloads`) |
| `--space SPACE` | Backlog space, e.g. `d-soft-prj.backlog.com` |
| `--project PROJECT` | Project key, e.g. `HRM` |
| `--milestone MILESTONE` | Milestone name (default `Sprint 3`, created if missing) |
| `--all` | Also create tasks that are already Done (Done -> Closed) |
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
| **Estimated Hours** | Estimate (days) x 8 |
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

- Never commit your API key, Excel files, or `created_issues.json` to Git.
- Suggested `.gitignore`:

```
created_issues.json
*.xlsx
.env
__pycache__/
```

## Requirements

- Python 3.8+
- A Backlog account with permission to add issues, categories, and milestones in the project
- Libraries: `openpyxl`, `requests`

> The `NotOpenSSLWarning` on macOS (Python built against LibreSSL) is only a version warning and does not affect execution.

## License

MIT
