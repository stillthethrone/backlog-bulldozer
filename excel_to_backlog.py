#!/usr/bin/env python3
"""
Đọc Sprint_3_Planning_Backlog_UC_lich_du_kien.xlsx (sheet "Backlog") -> tạo Parent (UC) + Child (BE/FE) trên Backlog.

Cài đặt:   pip install openpyxl requests
Chạy thử:  python3 excel_to_backlog.py              (dry-run: chỉ in ra, không tạo gì)
Chạy thật: python3 excel_to_backlog.py --run

Tuỳ chọn:
  --file PATH        đường dẫn Excel (mặc định ~/Downloads/Sprint_3_Planning_Backlog_UC_lich_du_kien.xlsx)
  --space HOST       mặc định d-soft-prj.backlog.com
  --project KEY      mặc định HRM
  --milestone NAME   mặc định "Sprint 3" (truyền "" để bỏ qua)
  --all              tạo cả task đã Done (mặc định bỏ qua Done)
API Key được hỏi khi chạy (không hiện trên màn hình), hoặc lấy từ biến BACKLOG_API_KEY.
"""
import argparse, getpass, json, os, sys, time, unicodedata
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import load_workbook

DEFAULT_FILE = Path.home() / "Downloads" / "Sprint_3_Planning_Backlog_UC_lich_du_kien.xlsx"
STATE_FILE = Path(__file__).with_name("created_issues.json")  # chạy lại không tạo trùng
HOURS_PER_DAY = 8
ISSUE_TYPE_NAME = "Task"
PRIORITY_MAP = {"critical": "High", "high": "High", "medium": "Normal", "low": "Low"}  # Backlog chỉ có High/Normal/Low
# Backlog mặc định: 1 Open, 2 In Progress, 3 Resolved, 4 Closed
STATUS_MAP = {"in progress": 2, "done": 4}


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


class Backlog:
    def __init__(self, space, key):
        self.base, self.key = f"https://{space}/api/v2", key

    def __call__(self, method, path, **kw):
        kw.setdefault("params", {})["apiKey"] = self.key
        r = requests.request(method, f"{self.base}{path}", timeout=30, **kw)
        if not r.ok:
            raise RuntimeError(f"{method} {path} -> {r.status_code}: {r.text}")
        time.sleep(0.25)  # tránh rate limit
        return r.json() if r.text else {}


def read_rows(path):
    ws = load_workbook(path, data_only=True)["Backlog"]
    header, hrow = None, None
    for row in ws.iter_rows(min_row=1, max_row=30):
        names = [norm(c.value) for c in row]
        if "id" in names and "assignee" in names:
            header, hrow = names, row[0].row
            break
    if not header:
        sys.exit("Không tìm thấy dòng tiêu đề (cần cột ID và Assignee) trong sheet Backlog")
    idx = {n: i for i, n in enumerate(header) if n}
    need = {"id": "id", "module": "module", "kind": "loai", "method": "method", "endpoint": "endpoint / man hinh",
            "desc": "mo ta", "actor": "actor", "priority": "priority", "assignee": "assignee",
            "status": "trang thai", "sp": "est. (sp)", "days": "uoc luong thoi gian (ngay)",
            "start": "ngay bat dau", "end": "ngay ket thuc"}
    missing = [v for v in need.values() if v not in idx]
    if missing:
        sys.exit(f"Thiếu cột trong Excel: {missing}")
    rows = []
    for r in ws.iter_rows(min_row=hrow + 1, values_only=True):
        if not r[idx["id"]]:
            continue
        d = {k: r[idx[v]] for k, v in need.items()}
        d["id"] = str(d["id"]).strip()
        rows.append(d)
    return rows


def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def fmt_dt(d):
    return d.strftime("%d/%m/%Y %H:%M") if isinstance(d, datetime) else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=str(DEFAULT_FILE))
    ap.add_argument("--space", default=os.environ.get("BACKLOG_SPACE", "d-soft-prj.backlog.com"))
    ap.add_argument("--project", default=os.environ.get("BACKLOG_PROJECT", "HRM"))
    ap.add_argument("--milestone", default="Sprint 3")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--run", action="store_true")
    a = ap.parse_args()

    rows = read_rows(a.file)
    if not a.all:
        rows = [r for r in rows if norm(r["status"]) != "done"]
    print(f"Đọc {a.file}: {len(rows)} task thuộc {len({r['id'] for r in rows})} UC "
          f"({'tạo thật' if a.run else 'DRY-RUN'})")

    key = os.environ.get("BACKLOG_API_KEY") or getpass.getpass("Nhập Backlog API Key: ").strip()
    api = Backlog(a.space, key)

    project = api("GET", f"/projects/{a.project}")
    pid = project["id"]
    types = {t["name"]: t["id"] for t in api("GET", f"/projects/{a.project}/issueTypes")}
    prios = {p["name"]: p["id"] for p in api("GET", "/priorities")}
    users = api("GET", f"/projects/{a.project}/users")
    cats = {c["name"]: c["id"] for c in api("GET", f"/projects/{a.project}/categories")}
    miles = {m["name"]: m["id"] for m in api("GET", f"/projects/{a.project}/versions")}  # "versions" API = Milestone
    if ISSUE_TYPE_NAME not in types:
        sys.exit(f"Không có Issue Type '{ISSUE_TYPE_NAME}'. Hiện có: {list(types)}")

    def find_user(name):
        n = norm(name)
        for u in users:
            if n and n in norm(u["name"]):
                return u["id"]
        print(f"  ! Không tìm thấy user '{name}' trong project -> bỏ trống Assignee")

    def ensure(table, path, name):
        if not name:
            return None
        if name not in table:
            if not a.run:
                return None
            table[name] = api("POST", path, data={"name": name})["id"]
        return table[name]

    milestone_id = ensure(miles, f"/projects/{a.project}/versions", a.milestone) if a.milestone else None
    if a.milestone and not milestone_id and not a.run:
        print(f"  (Milestone '{a.milestone}' chưa có, sẽ được tạo khi chạy thật)")

    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}

    def create(summary, r, desc, parent_id=None, start=None, end=None, hours=None, cat_names=()):
        data = {
            "projectId": pid, "summary": summary, "description": desc,
            "issueTypeId": types[ISSUE_TYPE_NAME],
            "priorityId": prios[PRIORITY_MAP.get(norm(r["priority"]), "Normal")],
        }
        if parent_id: data["parentIssueId"] = parent_id
        if start: data["startDate"] = start.strftime("%Y-%m-%d")
        if end: data["dueDate"] = end.strftime("%Y-%m-%d")
        if hours: data["estimatedHours"] = hours
        uid = find_user(r["assignee"])
        if uid: data["assigneeId"] = uid
        cids = [c for c in (ensure(cats, f"/projects/{a.project}/categories", n) for n in cat_names) if c]
        if cids: data["categoryId[]"] = cids
        if milestone_id: data["milestoneId[]"] = [milestone_id]
        if not a.run:
            print("    [dry-run]", {k: v for k, v in data.items() if k != "description"})
            return {"id": 0, "issueKey": "DRY-RUN"}
        return api("POST", "/issues", data=data)

    groups = {}
    for r in rows:
        groups.setdefault(r["id"], []).append(r)

    for uc, items in groups.items():
        first = items[0]
        starts = [i["start"] for i in items if isinstance(i["start"], datetime)]
        ends = [i["end"] for i in items if isinstance(i["end"], datetime)]
        total_h = sum((to_float(i["days"]) or 0) * HOURS_PER_DAY for i in items) or None
        print(f"\n== {uc} ({len(items)} task)")

        # ---- Parent ----
        st = state.setdefault(uc, {})
        if "id" in st:
            parent_id = st["id"]
            print("  parent đã có:", st["key"])
        else:
            screens = [i["endpoint"] for i in items if i["kind"] == "Frontend"]
            title = screens[0] if screens else items[0]["endpoint"]
            p = create(f"[{uc}] {first['module']} - {title}", first,
                       desc=f"UC: {uc}\nModule: {first['module']}\nActor: {first['actor'] or ''}\n"
                            f"Số task: {len(items)} ({', '.join(sorted({i['kind'] for i in items}))})",
                       start=min(starts) if starts else None, end=max(ends) if ends else None,
                       hours=total_h, cat_names=[first["module"]])
            parent_id = p["id"]
            st.update(id=p["id"], key=p["issueKey"], children={})
            print("  tạo parent:", p["issueKey"])

        # ---- Children: mỗi dòng Excel = 1 task ----
        for r in items:
            ckey = f"{r['kind']}|{r['method'] or ''}|{r['endpoint']}"
            if ckey in st.get("children", {}):
                print("  child đã có:", ckey)
                continue
            is_be = r["kind"] == "Backend"
            label = {"Backend": "BE", "Frontend": "FE"}.get(r["kind"], r["kind"])
            what = f"{r['method']} {r['endpoint']}" if r["method"] else r["endpoint"]
            desc = (f"**{r['kind']}** - {what}\n\n{r['desc'] or ''}\n\n"
                    f"- UC: {uc}\n- Module: {r['module']}\n- Actor: {r['actor'] or ''}\n"
                    f"- Priority (Excel): {r['priority']}\n- Est.: {r['sp']} SP\n"
                    f"- Lịch dự kiến: {fmt_dt(r['start'])} -> {fmt_dt(r['end'])}")
            days = to_float(r["days"])
            c = create(f"[{uc}][{label}] {what}", r, desc, parent_id=parent_id,
                       start=r["start"] if isinstance(r["start"], datetime) else None,
                       end=r["end"] if isinstance(r["end"], datetime) else None,
                       hours=days * HOURS_PER_DAY if days else None,
                       cat_names=[r["module"], r["kind"]])
            st.setdefault("children", {})[ckey] = c["issueKey"]
            print("  tạo child:", c["issueKey"], f"[{label}] {what}")

            sid = STATUS_MAP.get(norm(r["status"]))
            if sid and a.run:
                upd = {"statusId": sid}
                if sid == 4: upd["resolutionId"] = 0  # 0 = Fixed
                api("PATCH", f"/issues/{c['issueKey']}", data=upd)

        if a.run:
            STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2))
        else:
            state.pop(uc, None)

    print("\nXong." if a.run else "\nĐây là dry-run. Thêm --run để tạo thật.")


if __name__ == "__main__":
    main()
