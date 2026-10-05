# backlog-bulldozer

> Ủi bảng kế hoạch Sprint trong Excel lên Backlog: tự tạo task cha (UC), task con (Backend / Frontend) và điền sẵn các trường.

**backlog-bulldozer** đọc file Excel planning rồi tạo issue trên Backlog bằng API. Mỗi UC là một task cha, mỗi dòng trong Excel là một task con `[BE]` hoặc `[FE]`, kèm người làm, độ ưu tiên, ngày bắt đầu, hạn và giờ ước lượng.

> **Trạng thái:** phần đọc Excel đã chạy đúng. Phần gọi API tạo task **chưa được kiểm tra với Backlog thật** (do chưa có API Key lúc viết). Hãy luôn chạy dry-run trước.

## Tính năng

- Đọc file Excel trong `~/Downloads`, hoặc chỉ định bằng `--file`.
- Mặc định **chỉ tạo task chưa Done**, bỏ qua task đã xong. Thêm `--all` để tạo cả task Done.
- Mỗi UC là một task cha, chứa các task con BE/FE.
- Khớp Assignee theo tên, cảnh báo nếu không tìm thấy.
- Tự tạo Category và Milestone nếu chưa có.
- Lưu các task đã tạo vào `created_issues.json`, chạy lại không bị trùng.
- **Dry-run mặc định**, chỉ tạo thật khi thêm `--run`.
- API Key được hỏi khi chạy (gõ vào không hiện ra), không cần lưu trong file hay biến môi trường.

## Bắt đầu

### 1. Cài đặt

```bash
pip3 install openpyxl requests
```

### 2. Lấy API Key

Backlog → **Personal settings** → **API** → tạo key mới.

### 3. Chạy thử (dry-run)

```bash
python3 excel_to_backlog.py
```

Script hỏi API Key, rồi in danh sách task sẽ tạo cùng các cảnh báo. **Chưa tạo gì trên Backlog.**

> Dry-run vẫn gọi API để đọc thông tin project và kiểm tra tên Assignee, nên vẫn cần API Key.

### 4. Tạo thật

```bash
python3 excel_to_backlog.py --run
```

> **Lỗi hay gặp trên macOS (zsh):** đừng dán kèm chú thích `# ...` vào cuối lệnh. zsh mặc định không coi `#` là chú thích, phần đó sẽ bị hiểu là tham số của script và báo `unrecognized arguments`. Chỉ dán đúng lệnh.

## Tham số

| Tham số | Ý nghĩa |
|---|---|
| `--file FILE` | Đường dẫn file Excel (mặc định tự tìm trong `~/Downloads`) |
| `--space SPACE` | Tên space Backlog, ví dụ `d-soft-prj.backlog.com` |
| `--project PROJECT` | Project key, ví dụ `HRM` |
| `--milestone MILESTONE` | Tên Milestone (mặc định `Sprint 3`, tự tạo nếu chưa có) |
| `--all` | Tạo cả task đã Done (Done → Closed) |
| `--run` | Tạo thật. Không có tham số này thì chỉ dry-run |

Ví dụ:

```bash
python3 excel_to_backlog.py --file ~/Downloads/Sprint_3.xlsx --milestone "Sprint 4" --run
```

## Cách điền các trường

| Trường Backlog | Lấy từ Excel |
|---|---|
| **Subject** | `[UC-4.7][BE] POST /projects/{id}/close` hoặc `[UC-4.7][FE] tên màn hình` |
| **Description** | Mô tả, Actor, Endpoint/Method, Priority gốc, Est. SP, lịch có cả giờ |
| **Status** | Trống → Open, In Progress → In Progress, Done → Closed (chỉ khi dùng `--all`) |
| **Priority** | Critical và High → High, Medium → Normal, Low → Low |
| **Assignee** | Khớp theo tên (Tiến, Bách, Tỉnh...) |
| **Category** | Module + Backend/Frontend |
| **Milestone** | `Sprint 3` (đổi bằng `--milestone`) |
| **Start Date / Due Date** | Ngày bắt đầu / kết thúc |
| **Estimated Hours** | Ước lượng (ngày) × 8 |
| **Parent issue** | Mỗi UC là một task cha chứa các task BE/FE |

## Những điểm cần biết

- **Trường để trống:** Version, Actual Hours, Resolution. Excel không có dữ liệu này.
- **Giờ trong lịch:** Backlog chỉ nhận ngày, nên giờ (08:30, 13:30...) được ghi trong Description.
- **Mức Critical:** Backlog mặc định không có, nên được gộp vào High. Priority gốc vẫn được ghi trong Description.
- **Task cha:** Excel không có tên UC, nên Subject được đặt theo dạng `[UC] Module - tên màn hình`, ví dụ `[UC-4.7] Quản lý dự án - Xác nhận đóng dự án`. Ngày và giờ ước lượng của task cha được tính từ các task con.
- **Assignee không khớp:** nếu tên không có trong project, script cảnh báo và để trống Assignee.
- **Issue Type:** script dùng `Task`. Nếu project đặt tên khác, script báo lỗi kèm danh sách tên có sẵn.
- **Chạy lại:** muốn tạo lại từ đầu thì xoá `created_issues.json` (và xoá issue cũ trên Backlog để khỏi trùng).

## Ví dụ kết quả dry-run

Với file Sprint 3 mẫu, script đọc được:

- 37 task chưa Done thuộc 17 UC (36 task trống trạng thái, 1 task In Progress)
- 58 task Done bị bỏ qua (thêm `--all` để tạo cả chúng)

## Bảo mật

- Không commit API Key, file Excel hay `created_issues.json` lên Git.
- Gợi ý `.gitignore`:

```
created_issues.json
*.xlsx
.env
__pycache__/
```

## Yêu cầu

- Python 3.8+
- Tài khoản Backlog có quyền thêm issue, category và milestone trong project
- Thư viện: `openpyxl`, `requests`

> Cảnh báo `NotOpenSSLWarning` trên macOS (Python dùng LibreSSL) chỉ là cảnh báo phiên bản, không ảnh hưởng việc chạy.

## License

MIT
