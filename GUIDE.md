# WorkLog — Hướng dẫn đầy đủ theo Flow (Project → KPI → Sub Task → Timer)

---

## 1. CÀI ĐẶT LẦN ĐẦU

```bash
pip install openpyxl pandas
python -m tools.build
python -m tools.seed
```

Kết quả: tạo `worklog.xlsx` với 6 sheet, dữ liệu mẫu.

---

## 2. FLOW CHÍNH: Project → Father Task → Sub Task → Làm việc → Hoàn thành

```
                 ┌─────────────┐
                 │  PROJECT    │  ← project add
                 ├─────────────┤
                 │  FATHER TASK│  ← kpi add (thuộc project)
                 ├─────────────┤
                 │  SUB TASK   │  ← start -s (tự động sinh code)
                 ├─────────────┤
                 │  TIME ENTRY │  ← stop (ghi vào Excel)
                 └─────────────┘
```

---

### 2.1. TẠO PROJECT

**Lệnh:**
```bash
python -m tools.timer project add -n "Backend" -d "API development"
```

| Payload | Ý nghĩa | Bắt buộc |
|---------|---------|----------|
| `-n`, `--name` | Tên project | ✅ |
| `-d`, `--description` | Mô tả project | ❌ |

**Dữ liệu ghi vào sheet Projects:**
| Cột | Giá trị |
|-----|---------|
| A: Code | `PRJ-1` (tự động) |
| B: Name | `Backend` |
| C: Description | `API development` |
| D: Status | `Active` |
| E: Created Date | hôm nay |

**Các lệnh khác:**
```bash
# Danh sách
python -m tools.timer project list

# Đổi tên (cập nhật KPIs + Time Entries)
python -m tools.timer project rename -n "Backend" --new-name "Backend Services"

# Archive (ẩn project cũ)
python -m tools.timer project archive -n "Old Project"
```

---

### 2.2. TẠO FATHER TASK (KPI) TRONG PROJECT

**Lệnh:**
```bash
python -m tools.timer kpi add -t "API Auth" -p Backend -d 5
```

| Payload | Ý nghĩa | Bắt buộc |
|---------|---------|----------|
| `-t`, `--name` | Tên father task | ✅ |
| `-p`, `--project` | Project chứa KPI này | ✅ (phải tồn tại) |
| `-d`, `--deadline` | Hạn chót (số ngày), mặc định 7 | ❌ |

**Dữ liệu ghi vào sheet KPIs:**
| Cột | Giá trị |
|-----|---------|
| A: Code | `PRJ-1-KPI-1` (tự động, scope theo project) |
| B: Father Task | `API Auth` |
| C: Project | `Backend` |
| D: Date | hôm nay |
| E: Deadline (days) | `5` |
| F: Deadline (h) | `=E*7.5` (công thức) |
| G: Status | `In Progress` |
| H: Completed Date | (trống) |

**Các lệnh khác:**
```bash
# Danh sách (group theo project)
python -m tools.timer kpi list

# Xem tiến độ (group theo project)
python -m tools.timer kpi status

# Sửa deadline
python -m tools.timer kpi edit -t "API Auth" -d 10

# Sửa project
python -m tools.timer kpi edit -t "API Auth" -p Frontend

# Đổi tên (cập nhật cả Time Entries)
python -m tools.timer kpi rename -t "API Auth" -n "Authentication"

# Đánh dấu hoàn thành
python -m tools.timer kpi done -t "API Auth"

# Xóa KPI
python -m tools.timer kpi delete -t "API Auth"
```

---

### 2.3. BẮT ĐẦU PHIÊN LÀM VIỆC

**Lệnh:**
```bash
python -m tools.timer start -p Backend -t "API Auth" -s "Login endpoint" -c Development -d "Implement JWT"
```

| Payload | Ý nghĩa | Bắt buộc | Tự động nếu bỏ qua |
|---------|---------|----------|-------------------|
| `-p`, `--project` | Project | ❌ | Git folder name |
| `-t`, `--task` | Father task | ❌ | Git branch name |
| `-s`, `--subtask` | Sub task | ❌ | — |
| `--subtask-code` | Sub task code (tự sinh nếu có -s) | ❌ | `PRJ1-KPI-1-ST-N` |
| `-c`, `--category` | Danh mục | ❌ | `Development` |
| `-d`, `--description` | Mô tả | ❌ | Git commit message |

**Dữ liệu lưu vào state** (`~/.worklog_timer.json`):
```json
{
  "accumulated_seconds": 0,
  "segment_start": "2026-07-21T09:00:00",
  "paused": false,
  "project": "Backend",
  "task": "API Auth",
  "subtask": "Login endpoint",
  "subtask_code": "PRJ-1-KPI-1-ST-1",
  "category": "Development",
  "description": "Implement JWT"
}
```

---

### 2.4. TẠM DỪNG (GIẢI LAO / HỌP)

**Lệnh:**
```bash
python -m tools.timer pause -r meeting
```

| Payload | Ý nghĩa | Bắt buộc |
|---------|---------|----------|
| `-r`, `--reason` | Lý do nghỉ | ❌ (mặc định: `break`) |

**Các lý do thường dùng:** `meeting`, `lunch`, `break`, `review`, `context-switch`, `phone`, `email`, `personal`

**Tiếp tục:**
```bash
python -m tools.timer continue
```

**Cách hoạt động:**
- `pause` → ghi lại thời điểm bắt đầu nghỉ + lý do
- `continue` → tính thời gian nghỉ (phút) + lưu vào `pause_log` trong state
- `stop` → tổng hợp các phiên nghỉ thành chuỗi `meeting(15m); lunch(30m)` ghi vào cột L

---

### 2.5. KẾT THÚC PHIÊN

**Lệnh:**
```bash
python -m tools.timer stop -d "Xong login API, đã test JWT"
```

| Payload | Ý nghĩa | Bắt buộc |
|---------|---------|----------|
| `-d`, `--description` | Mô tả công việc | ✅ (nếu chưa có ở start) |
| `-p`, `--project` | Ghi đè project | ❌ |
| `-t`, `--task` | Ghi đè father task | ❌ |
| `-s`, `--subtask` | Ghi đè subtask | ❌ |
| `-c`, `--category` | Ghi đè danh mục | ❌ |

**Dữ liệu ghi vào sheet Time Entries:**
| Cột | Giá trị |
|-----|---------|
| A: Date | `21-07-2026` |
| B: Day | `Tue` |
| C: Project | `Backend` |
| D: Father Task | `API Auth` |
| E: Sub Task | `Login endpoint` |
| F: Sub Task Code | `PRJ-1-KPI-1-ST-1` |
| G: Category | `Development` |
| H: Start Time | `09:00` |
| I: End Time | `11:30` |
| J: Duration (h) | `=IF(AND(H<>'',I<>''),(I-H)*24,'')` |
| K: Description | `Xong login API, đã test JWT` |
| L: Break Info | `meeting(15m); lunch(30m)` |

**Webhook payload** (nếu có `WORKLOG_WEBHOOK_URL`):
```json
{
  "project": "Backend",
  "task": "API Auth",
  "duration_h": 2.5,
  "duration_m": 150,
  "description": "Xong login API, đã test JWT",
  "category": "Development",
  "subtask": "Login endpoint",
  "break_info": "meeting(15m); lunch(30m)",
  "start_time": "09:00",
  "end_time": "11:30",
  "source": "worklog"
}
```

---

### 2.6. SUB TASK TỰ ĐỘNG

Sub task mới được tạo tự động khi `start` với tham số `-s` mới:

```bash
# Lần 1 → sinh PRJ-1-KPI-1-ST-1
python -m tools.timer start -p Backend -t "API Auth" -s "Login endpoint"
python -m tools.timer stop -d "Xong login"

# Lần 2 → sinh PRJ-1-KPI-1-ST-2 (tự động)
python -m tools.timer start -p Backend -t "API Auth" -s "Register endpoint"
python -m tools.timer stop -d "Xong register"
```

**Quản lý subtask:**
```bash
# Danh sách
python -m tools.timer subtask list

# Lọc theo father task
python -m tools.timer subtask list -t "API Auth"

# Đổi tên (cập nhật Time Entries)
python -m tools.timer subtask rename -s "Login endpoint" -n "Authentication Module"

# Xóa tên subtask
python -m tools.timer subtask delete -s "Old Module"
```

---

### 2.7. XEM BÁO CÁO

```bash
# Hôm nay (group theo project)
python -m tools.timer today

# Tuần này (group theo project)
python -m tools.timer week

# Tất cả entries gom theo father task (group theo project)
python -m tools.timer tasks

# Lọc theo father task
python -m tools.timer tasks -t "API Auth"
```

---

### 2.8. HOÀN THÀNH KPI

```bash
# Khi làm xong hết các subtask
python -m tools.timer kpi done -t "API Auth"
```

Dữ liệu cập nhật trong sheet KPIs:
| Cột | Giá trị |
|-----|---------|
| G: Status | `Done` |
| H: Completed Date | hôm nay |

---

## 3. CÁC LỆNH KHÁC

### Pomodoro
```bash
python -m tools.timer pomo                      # 25' work + 5' rest
python -m tools.timer pomo --work 45 --rest 15  # 45' work + 15' rest
```

### Hủy timer
```bash
python -m tools.timer cancel   # Xóa state, không ghi gì
```

### Trạng thái timer
```bash
python -m tools.timer status   # Đang chạy / đang pause / elapsed
```

### Validate & Báo cáo
```bash
python -m tools.validate       # Kiểm tra dữ liệu
python -m tools.gamify         # Điểm/streak/level
python -m tools.report         # Báo cáo tổng hợp
python -m tools.auto            # Chạy tất cả (validate + gamify + report)
```

---

## 4. TÍNH NĂNG MỞ RỘNG

### SQLite Backend
```powershell
$env:WORKLOG_DB = "sqlite"
python -m tools.timer start -p Backend -t "API Auth" -s "Login"
python -m tools.timer stop -d "Xong"
# Dữ liệu ghi song song vào ~/.worklog.db
sqlite3 ~/.worklog.db "SELECT * FROM time_entries;"
```

### Webhook Notifications
```powershell
$env:WORKLOG_WEBHOOK_URL = "https://hooks.slack.com/services/xxx/yyy/zzz"
python -m tools.timer webhook-test              # Kiểm tra kết nối
python -m tools.timer stop -d "Xong feature"    # Tự động gửi webhook
```

---

## 5. CẤU TRÚC EXCEL

| Sheet | Vị trí | Mô tả |
|-------|--------|-------|
| Time Entries | 1 | Nhật ký thời gian (ghi tự động) |
| Projects | 2 | Danh sách dự án (ghi bằng `project add`) |
| KPIs | 3 | Father task + deadline (ghi bằng `kpi add`) |
| Weekly Summary | 4 | Tự động tổng hợp tuần |
| Daily Detail | 5 | Tự động chi tiết ngày |
| Monthly | 6 | Tự động % hoàn thành KPI |

### Cột Time Entries

| Cột | Tên | Ghi chú |
|-----|-----|---------|
| A | Date | `dd-mm-yyyy` |
| B | Day | Tự động |
| C | Project | |
| D | Father Task | |
| E | Sub Task | |
| F | Sub Task Code | `PRJ1-KPI-1-ST-1` |
| G | Category | Development, Testing, Review... |
| H | Start Time | `hh:mm` |
| I | End Time | `hh:mm` |
| J | Duration (h) | Công thức tự động |
| K | Description | |
| L | Break Info | `meeting(15m); lunch(30m)` |
| M→R | Helper | Cột ẩn (công thức) |

---

## 6. VÍ DỤ LUỒNG THỰC TẾ HOÀN CHỈNH

```bash
# === SÁNG ===

# Tạo project (1 lần)
python -m tools.timer project add -n "Frontend" -d "UI dashboard"
python -m tools.timer project add -n "Backend" -d "API services"

# Tạo KPI
python -m tools.timer kpi add -t "Dashboard UI" -p Frontend -d 3
python -m tools.timer kpi add -t "API Stats" -p Backend -d 2

# Làm subtask 1
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Layout" -c "UI/UX"
python -m tools.timer pause -r meeting
# ... họp 15 phút ...
python -m tools.timer continue
python -m tools.timer stop -d "Xong grid layout"

# === TRƯA ===
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Layout" -c "UI/UX"
python -m tools.timer pause -r lunch
# ... ăn trưa 45 phút ...
python -m tools.timer continue
python -m tools.timer stop -d "Responsive mobile done"

# === CHIỀU ===
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Widget charts" -c "UI/UX"
python -m tools.timer stop -d "Xong biểu đồ doanh thu"

python -m tools.timer start -p Backend -t "API Stats" -s "GET /stats" -c Development
python -m tools.timer stop -d "Xong API + caching"

# === KIỂM TRA CUỐI NGÀY ===
python -m tools.timer today
python -m tools.timer kpi status
python -m tools.auto
```
