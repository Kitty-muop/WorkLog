# WorkLog

CLI tool + Excel workbook ghi nhật ký công việc cho lập trình viên. Tự động tính giờ, theo dõi KPI, phân tích hiệu suất. Tương thích WPS Office.

## Quick Start

```bash
# Build workbook + seed mẫu
python -m tools.build && python -m tools.seed

# Bắt đầu làm việc (auto-detect: project=thư mục, task=branch git)
python -m tools.timer start
# ... làm việc ...
python -m tools.timer stop -d "Mô tả công việc"

# Xem hôm nay
python -m tools.timer today

# KPI CRUD
python -m tools.timer kpi add -t "API" -p Backend -d 5
python -m tools.timer kpi list
python -m tools.timer kpi done -t "API"

# Sub Task CRUD
python -m tools.timer subtask rename -s "Auth" -n "Authentication"
python -m tools.timer subtask delete -s "Old Module"
```

## Cấu trúc

```
WorkLog/
├── worklog.xlsx         # File Excel chính (5 sheet, tự động tính)
├── worklog.xlsm         # Bản có macro (VBA timer button)
├── tools/               # Python CLI
│   ├── timer.py         # Timer + CRUD (start/stop/kpi/subtask/today...)
│   ├── build.py         # Xây dựng workbook
│   ├── seed.py          # Dữ liệu mẫu
│   ├── validate.py      # Kiểm tra dữ liệu
│   ├── gamify.py        # Điểm/streak/level
│   └── report.py        # Báo cáo tổng hợp
├── vba/                 # Macro VBA
└── docs/
```

## Yêu cầu

- Python 3.8+, `pip install openpyxl pandas`
- WPS Office hoặc Excel 2021+
