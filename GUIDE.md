# Hướng dẫn sử dụng WorkLog từ A-Z

## 1. Cài đặt lần đầu

```bash
pip install openpyxl pandas
python -m tools.build
python -m tools.seed
```

---

## 2. Flow đầy đủ: Project → Father Task → Sub Task → Làm việc → Hoàn thành

### 2.1. Tạo Project (dự án)

Project là cấp cao nhất. Phải có project trước khi tạo KPI.

```bash
python -m tools.timer project add -n Backend -d "API development and maintenance"
```

Kết quả: tạo project mã **PRJ-1**, trạng thái Active.

```bash
python -m tools.timer project list
```

### 2.2. Tạo Father Task (mục tiêu lớn) trong Project

```bash
python -m tools.timer kpi add -t "Xây dựng API Auth" -p Backend -d 3
```

Kết quả: tạo KPI với mã **PRJ-1-KPI-1** (scope trong project Backend), deadline 3 ngày.

> **Lưu ý**: Project `Backend` phải tồn tại trong Projects sheet. Nếu chưa có, thêm trước bằng `project add`.

### 2.3. Tạo Sub Task (việc nhỏ)

Sub task được tạo **tự động** khi bạn `start` với tham số `-s`:

```bash
python -m tools.timer start -p Backend -t "Xây dựng API Auth" -s "Login API" -c Development
```

Mỗi lần start với `-s` mới, CLI tự động sinh mã **PRJ-1-KPI-1-ST-1**, **PRJ-1-KPI-1-ST-2**,...

> Cách này không cần lệnh tạo subtask riêng — cứ `start -s "Tên"` là có subtask mới.

### 2.3. Làm việc với timer

```bash
# Bắt đầu phiên (đã có subtask)
python -m tools.timer start -p Backend -t "Xây dựng API Auth" -s "Login API" -c Development

# Tạm dừng (nghỉ, họp...)
python -m tools.timer pause -r meeting
python -m tools.timer continue

# Kết thúc phiên — ghi vào worklog.xlsx
python -m tools.timer stop -d "Xong chức năng login, đã test JWT"
```

### 2.4. Tiếp tục subtask cũ

```bash
# Làm tiếp subtask cũ (dùng đúng tên)
python -m tools.timer start -p Backend -t "Xây dựng API Auth" -s "Login API" -c Development
python -m tools.timer stop -d "Sửa lỗi refresh token"
```

### 2.5. Thêm subtask mới

```bash
# Tự động tạo subtask mới + mã mới (KPI-1-ST-2)
python -m tools.timer start -p Backend -t "Xây dựng API Auth" -s "Logout API" -c Development
python -m tools.timer stop -d "Hoàn thành logout"
```

### 2.6. Xem danh sách subtask + giờ

```bash
# Tất cả subtask
python -m tools.timer subtask list

# Lọc theo father task
python -m tools.timer subtask list -t "Xây dựng API Auth"

# Xem tất cả entries gom theo task
python -m tools.timer tasks
python -m tools.timer tasks -t "Xây dựng API Auth"
```

### 2.7. Quản lý subtask

```bash
# Đổi tên subtask (cập nhật toàn bộ Time Entries)
python -m tools.timer subtask rename -s "Login API" -n "Authentication Module"

# Xóa tên subtask
python -m tools.timer subtask delete -s "Old Module"
```

### 2.8. Hoàn thành Father Task

```bash
# Khi làm xong hết các subtask
python -m tools.timer kpi done -t "Xây dựng API Auth"

# Kiểm tra tiến độ
python -m tools.timer kpi status
```

---

## 3. Ví dụ luồng thực tế (cả dự án)

```bash
# === NGÀY 1 ===

# Tạo project
python -m tools.timer project add -n Frontend -d "UI development"
python -m tools.timer project add -n Backend -d "API development"

# Tạo mục tiêu trong project
python -m tools.timer kpi add -t "Dashboard UI" -p Frontend -d 2
python -m tools.timer kpi add -t "API Dashboard" -p Backend -d 3

# Làm subtask 1 của Dashboard UI
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Layout chính" -c "UI/UX"
python -m tools.timer stop -d "Xong grid layout, sidebar, navbar"

# Ăn trưa
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Layout chính" -c "UI/UX"
python -m tools.timer pause -r lunch
python -m tools.timer continue
python -m tools.timer stop -d "Responsive xong mobile"

# Subtask mới
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Widget thống kê" -c "UI/UX"
python -m tools.timer stop -d "Xong biểu đồ doanh thu"

# Kiểm tra cuối ngày
python -m tools.timer today
python -m tools.auto

# === NGÀY 2 ===

# Làm tiếp subtask cũ
python -m tools.timer start -p Frontend -t "Dashboard UI" -s "Widget thống kê" -c "UI/UX"
python -m tools.timer stop -d "Fix responsive chart"

# Xem còn gì chưa xong
python -m tools.timer kpi status

# Xong KPI
python -m tools.timer kpi done -t "Dashboard UI"

# Chuyển sang Backend
python -m tools.timer start -p Backend -t "API Dashboard" -s "GET /stats" -c Development
python -m tools.timer stop -d "Xong API thống kê + caching"

# Cuối ngày
python -m tools.timer week
python -m tools.auto
```

---

## 4. Các lệnh khác

```bash
# Pomodoro
python -m tools.timer pomo --work 25 --rest 5

# Sửa/xóa KPI
python -m tools.timer kpi edit -t "Dashboard UI" -d 5
python -m tools.timer kpi rename -t "Cũ" -n "Mới"
python -m tools.timer kpi delete -t "Dashboard UI"

# Validate dữ liệu
python -m tools.validate
python -m tools.gamify
python -m tools.report
```

---

## 5. Tính năng mở rộng (tùy chọn)

```powershell
# SQLite backend
$env:WORKLOG_DB = "sqlite"
python -m tools.timer start -p Backend -t "API Auth" -s "Login" -c Development

# Webhook Slack/Discord
$env:WORKLOG_WEBHOOK_URL = "https://hooks.slack.com/services/xxx/yyy/zzz"
python -m tools.timer webhook-test
```

---

## 6. Cấu trúc Excel

| Sheet | Mô tả |
|-------|-------|
| Time Entries | Nhật ký thời gian (cột A-L) |
| Projects | Danh sách dự án (Code, Name, Status) |
| KPIs | Father task + deadline + trạng thái |
| Weekly Summary | Tự động tổng hợp tuần |
| Daily Detail | Tự động chi tiết ngày |
| Monthly | Tự động % KPI theo tháng |
