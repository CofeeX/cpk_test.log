"""生成测试用 Excel 文件，模拟 CPK-PCBA 测试记录格式"""
import openpyxl
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# 第1行: 顶部标题 (Program Name)
ws.append(["Program Name:", "SC-POWER-TEST", "", "", "", "", ""])

# 第2行: 空行
ws.append([""] * 7)

# 元数据行 (header_rows)
# 项目名称行
ws.append(["项目名称", "序列号", "产品型号", "工单号", "测试员", "工装编号", "电压", "电流", "温度"])
# 备注行
ws.append(["备注", "", "", "", "", "", "工作电压", "工作电流", "工作温度"])
# 参数变量行
ws.append(["参数变量", "", "", "", "", "", "V", "mA", "°C"])
# 规格上限行
ws.append(["规格上限", "", "", "", "", "", 5.5, 100, 85])
# 规格下限行
ws.append(["规格下限", "", "", "", "", "", 4.5, 50, 60])

# 数据行
data = [
    ["", "SN001", "MODEL-A", "WO202601", "张三", "F01", 5.0, 75, 70],
    ["", "SN002", "MODEL-A", "WO202601", "张三", "F01", 5.1, 78, 72],
    ["", "SN003", "MODEL-A", "WO202601", "李四", "F02", 4.9, 72, 68],
    ["", "SN004", "MODEL-A", "WO202601", "李四", "F02", 5.2, 80, 75],
    ["", "SN005", "MODEL-A", "WO202601", "王五", "F01", 4.8, 70, 65],
]
for row in data:
    ws.append(row)

wb.save("/workspace/test_cpk_data.xlsx")
print("测试 Excel 已生成: /workspace/test_cpk_data.xlsx")
