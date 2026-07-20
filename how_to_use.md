# Worklog.xlsx — Hướng dẫn sử dụng

File `worklog.xlsx` gồm 5 bảng. **Bạn chỉ cần gõ dữ liệu vào Time Entries và KPIs**, 3 bảng còn lại tự động tính.

---

## 5 Bảng Tính

| Bảng | Việc của bạn |
|------|-------------|
| **Time Entries** | ✍️ Ghi mọi phiên làm việc (Date → Description) |
| **KPIs** | 🎯 Đặt father task + hạn chót + trạng thái (làm 1 lần) |
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

> Cột L-Q là cột phụ ẩn (truyền dữ liệu cho công thức) — không cần quan tâm.

---

## Sử dụng CLI Timer (khuyên dùng)

```bash
# Bắt đầu (auto: project=thư mục git, task=branch git)
python -m tools.timer start -p Backend -t "API Integration" -c Development

# Kết thúc + ghi vào worklog.xlsx
python -m tools.timer stop -d "Hoàn thành chức năng X"

# Xem hôm nay / tuần
python -m tools.timer today
python -m tools.timer week

# Tạm dừng / tiếp tục / hủy / trạng thái
python -m tools.timer pause
python -m tools.timer continue
python -m tools.timer cancel
python -m tools.timer status

# Pomodoro
python -m tools.timer pomo --work 25 --rest 5
```

### Quản lý Father Task (CRUD)

```bash
kpi add -t "Tên" -p Project -d 5      # Thêm mới (tự động gán mã KPI-1...)
kpi list                                # Danh sách
kpi status                              # Tiến độ tổng thể
kpi edit -t "Tên" -d 3                  # Sửa deadline
kpi rename -t "Cũ" -n "Mới"            # Đổi tên (cập nhật cả Time Entries)
kpi done -t "Tên"                       # Đánh dấu hoàn thành
kpi delete -t "Tên"                     # Xóa
```

### Quản lý Sub Task

```bash
subtask list [-t "Father"]              # Danh sách subtask + giờ
subtask rename -s "Cũ" -n "Mới"        # Đổi tên (tất cả entries)
subtask delete -s "Tên"                 # Xóa tên subtask
```

### Xem entries theo Father Task

```bash
tasks                                   # Tất cả entries gom theo father task
tasks -t "API Integration"              # Lọc theo tên
```

---

## Validate & Báo cáo

```bash
python -m tools.validate          # Kiểm tra dữ liệu
python -m tools.gamify            # Điểm/streak/level
python -m tools.report            # Báo cáo tổng hợp
python -m tools.auto              # Chạy tất cả
```

---

## VBA Timer (chỉ dành cho Excel)

File `worklog.xlsm` có nút Start Timer. Chọn dòng → bấm nút → làm việc → bấm nút lần nữa → tự động điền giờ. Nếu dùng `.xlsx`, import `TimerModule.bas` vào VBA Editor (`Alt+F11` → File → Import).

---

## Yêu cầu

- WPS Office hoặc Excel 2021+ (hỗ trợ AGGREGATE, INDEX/MATCH, SUMIFS)
- Python 3.8+, `pip install openpyxl pandas`
