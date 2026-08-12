# Worklog.xlsx — Hướng dẫn sử dụng

File `worklog.xlsx` gồm 5 bảng. **Bạn chỉ cần gõ dữ liệu vào Time Entries và KPIs**, 3 bảng còn lại tự động tính.

---

## 6 Bảng Tính

| Bảng | Việc của bạn |
|------|-------------|
| **Time Entries** | ✍️ Ghi mọi phiên làm việc (Date → Description) |
| **Projects** | 📁 Đăng ký dự án (làm 1 lần trước khi tạo KPI) |
| **KPIs** | 🎯 Đặt father task + hạn chót + trạng thái |
| **Weekly Summary** | 📊 Tự động — giờ theo task/tuần gần nhất |
| **Daily Detail** | 📋 Tự động — chi tiết theo ngày (sửa ô C1 để đổi ngày) |
| **Monthly** | 📈 Tự động — % hoàn thành KPI theo tháng |

---

## Cột trong Time Entries

| Cột | Tiêu đề | Bắt buộc | Định dạng |
|-----|---------|----------|-----------|
| A | Ngày | ✅ | `dd-mm-yyyy` |
| B | Thứ | ✅ | `Mon/Tue/Wed/Thu/Fri/Sat/Sun` |
| C | Dự án | ✅ | Văn bản |
| D | Father Task | ✅ | Khớp chính xác tên trong KPIs |
| E | Sub Task | ❌ | Văn bản |
| F | Sub Task Code | ❌ | Tự động khi dùng CLI |
| G | Danh mục | ✅ | VD: Development, Testing, Review |
| H | Giờ bắt đầu | ✅ | `hh:mm` 24h |
| I | Giờ kết thúc | ✅ | `hh:mm` 24h |
| J | Thời lượng (h) | ❌ | Tự động = `(I-H)×24` |
| K | Mô tả | ✅ | Nội dung công việc |
| L | Break Info | ❌ | `meeting(15m); lunch(30m)` — tự động khi dùng CLI |

> Cột M-R là cột phụ ẩn (truyền dữ liệu cho công thức) — không cần quan tâm.

---

## Sử dụng CLI Timer (khuyên dùng)

### Quản lý Project

Project là cấp cao nhất. Phải tạo project trước rồi mới tạo KPI.

```bash
# Thêm project
python -m tools.timer project add -n Backend -d "API development"

# Danh sách project
python -m tools.timer project list

# Đổi tên project (cập nhật KPIs + Time Entries)
python -m tools.timer project rename -n "Backend" --new-name "Backend Services"

# Archive project
python -m tools.timer project archive -n "Old Project"
```

### Quản lý Father Task (KPI)

KPI thuộc về project. Mã KPI có dạng `PRJ1-KPI-1`.

```bash
kpi add -t "API Auth" -p Backend -d 5   # Project Backend phải tồn tại
kpi list                                # Danh sách (group theo project)
kpi status                              # Tiến độ tổng thể (group theo project)
kpi edit -t "API Auth" -d 3             # Sửa deadline
kpi rename -t "Cũ" -n "Mới"            # Đổi tên (cập nhật cả Time Entries)
kpi done -t "API Auth"                  # Đánh dấu hoàn thành
kpi delete -t "API Auth"                # Xóa
```

### Timer

```bash
# Bắt đầu (auto: project=thư mục git, task=branch git)
python -m tools.timer start -p Backend -t "API Auth" -s "Login" -c Development

# Kết thúc + ghi vào worklog.xlsx
python -m tools.timer stop -d "Hoàn thành chức năng X"

# Xem hôm nay / tuần (group theo project)
python -m tools.timer today
python -m tools.timer week

# Tạm dừng / tiếp tục / hủy / trạng thái
python -m tools.timer pause -r meeting
python -m tools.timer continue
python -m tools.timer cancel
python -m tools.timer status

# Pomodoro
python -m tools.timer pomo --work 25 --rest 5
```

### Quản lý Sub Task

Subtask là cấp con của Father Task. Có thể quản lý độc lập (giống KPI) hoặc tạo tự động qua timer.

**Tạo tự động qua timer:** Khi `start -s "Tên"`, code tự động sinh theo format `{KPI_CODE}-ST-{N}` (vd: `PRJ-1-KPI-1-ST-1`).

**Cách auto-generation:**
1. Khi `stop`, nếu có `-s` nhưng không có `--subtask-code`
2. Tra KPI code từ father task trong sheet KPIs
3. Tìm số thứ tự lớn nhất → `{KPI_CODE}-ST-{max+1}`

**CRUD SubTask (quản lý trên sheet SubTasks):**
```bash
subtask add -s "Login" -t "API Auth"    # Thêm subtask mới
subtask list [-t "Father"]              # Danh sách subtask + tổng giờ + father task
subtask edit -s "Login" -p Frontend     # Sửa project/father task
subtask edit -s "Login" -n "Auth"       # Sửa tên
subtask rename -s "Cũ" -n "Mới"        # Đổi tên (tất cả entries + Google Sheets)
subtask delete -s "Tên"                 # Xóa subtask khỏi SubTasks sheet
subtask done -s "Login"                 # Đánh dấu hoàn thành
subtask status                          # Xem tiến độ subtask (group theo father task)
```

**Subtask trong báo cáo:**
```bash
tasks [-t "Father"]                     # entries gom theo father task, hiện subtask
today                                   # entries hôm nay, hiện subtask
```

### Xem entries theo Father Task

```bash
tasks                                   # Tất cả entries gom theo father task
tasks -t "API Integration"              # Lọc theo tên
```

---

## Break Tracking

Khi `pause`, thêm `-r REASON` để ghi lại lý do nghỉ (meeting, lunch, break, review, context-switch, ...).  
Các phiên nghỉ được ghi vào cột **Break Info** (L) trong Time Entries dạng `meeting(15m); lunch(30m)`.

## SQLite Backend (tùy chọn)

Đặt biến môi trường `WORKLOG_DB=sqlite` để tự động đồng bộ dữ liệu vào SQLite song song với Excel:

```bash
# PowerShell
$env:WORKLOG_DB = "sqlite"
python -m tools.timer start -p Backend -t "API Integration"
python -m tools.timer stop -d "Xong"
# Dữ liệu được ghi vào cả worklog.xlsx và ~/.worklog.db
```

Dùng `sqlite3 ~/.worklog.db` để truy vấn trực tiếp.

## Webhook Notifications (tùy chọn)

Đặt biến môi trường `WORKLOG_WEBHOOK_URL` để tự động gửi thông báo khi stop timer:

```bash
$env:WORKLOG_WEBHOOK_URL = "https://hooks.slack.com/services/xxx/yyy/zzz"
python -m tools.timer stop -d "Xong feature"
# POST JSON {project, task, duration_h, description, ...} đến webhook
```

Kiểm tra webhook:
```bash
python -m tools.timer webhook-test
```

## Google Sheets Sync (tùy chọn)

Tự động đồng bộ dữ liệu lên Google Sheets khi chạy timer và CRUD commands.

### Chuẩn bị

1. Tạo **service account** trên Google Cloud Console → IAM & Admin → Service Accounts
2. Tạo key JSON, tải về máy
3. Tạo Google Sheet, chia sẻ với email service account (quyền Editor)
4. Lấy Spreadsheet ID từ URL: `https://docs.google.com/spreadsheets/d/{ID}/edit`

### Sheet structure

Tạo 3 sheet trong Google Sheet với tên và header giống Excel:
- **Projects**: Code, Project Name, Description, Status, Created Date
- **KPIs**: Code, Father Task, Project, Date, Deadline (days), Deadline (h), Status, Completed Date, Notes
- **Time Entries**: Date, Day, Project, Father Task, Sub Task, Sub Task Code, Category, Start Time, End Time, Duration (h), Description, Break Info

### Cấu hình

```bash
$env:WORKLOG_GSHEETS_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
$env:WORKLOG_GSHEETS_KEY = "C:\path\to\service-account-key.json"
```

Sau đó dùng CLI bình thường — dữ liệu tự động đồng bộ:

```bash
python -m tools.timer start -p Backend -t "API"
python -m tools.timer stop -d "Xong"
# Dữ liệu được ghi vào cả Excel và Google Sheets

python -m tools.timer project add -n "Frontend"
python -m tools.timer kpi add -t "Setup" -p Frontend
# CRUD cũng đồng bộ lên Google Sheets
```

## Validate & Báo cáo

```bash
python -m tools.validate          # Kiểm tra dữ liệu
python -m tools.gamify            # Điểm/streak/level
python -m tools.report            # Báo cáo tổng hợp
python -m tools.auto              # Chạy tất cả (gồm init SQLite)
```

---

## VBA Timer (chỉ dành cho Excel)

File `worklog.xlsm` có nút Start Timer. Chọn dòng → bấm nút → làm việc → bấm nút lần nữa → tự động điền giờ. Nếu dùng `.xlsx`, import `TimerModule.bas` vào VBA Editor (`Alt+F11` → File → Import).

---

## Yêu cầu

- WPS Office hoặc Excel 2021+ (hỗ trợ AGGREGATE, INDEX/MATCH, SUMIFS)
- Python 3.8+, `pip install openpyxl pandas`
