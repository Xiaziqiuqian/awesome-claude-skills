#!/usr/bin/env python3
"""
生产规划仓库需求计算 Excel 模板生成脚本
Generates an Excel workbook for multi-SKU manufacturing warehouse planning.

Sheets:
  1. 使用说明       - README / instructions
  2. 参数设置       - Common parameters
  3. 产品SKU        - Product-level inputs
  4. 原料BOM        - Raw-material bill-of-materials matrix
  5. 原料仓计算     - Raw-material warehouse stock calculation
  6. 过程品仓计算   - WIP warehouse stock calculation
  7. 成品仓计算     - Finished-goods warehouse stock calculation
  8. 汇总           - Summary across all warehouses
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
import os

# ---------------------------------------------------------------------------
# Color palette
# ---------------------------------------------------------------------------
C_HEADER_BG  = "1F4E79"   # dark blue  – section headers
C_HEADER_FG  = "FFFFFF"   # white
C_INPUT_BG   = "DEEAF1"   # light blue – user-input cells
C_CALC_BG    = "E2EFDA"   # light green – formula / calculated cells
C_TITLE_BG   = "2E75B6"   # medium blue – sheet title row
C_BORDER     = "8EA9C1"

THIN = Side(border_style="thin", color=C_BORDER)
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NUM_FMT_INT   = '#,##0'
NUM_FMT_DEC2  = '#,##0.00'
NUM_FMT_PCT   = '0.0%'

NUM_SKU_ROWS = 12   # data rows for SKUs (supports up to 12 products)
NUM_RM_ROWS  = 15   # data rows for raw materials
DEFAULT_BASE_LEAD_DAYS = 15
DEFAULT_RM_WITH_PRESET_LEAD = 8
MANUFACTURING_PROCESS_STAGES = ["粗破", "烘干", "磨粉", "低温碳化", "石墨化", "高温碳化", "成品筛分"]
NUM_PROCESSES = len(MANUFACTURING_PROCESS_STAGES)
GRAPHITIZATION_STAGE_INDEX = MANUFACTURING_PROCESS_STAGES.index("石墨化")
PROCESS_BLOCK_ROWS = NUM_SKU_ROWS + 2
MIN_FAULT_DENOMINATOR = 0.01


def hfont(bold=True, color=C_HEADER_FG, sz=10):
    return Font(name="微软雅黑", bold=bold, color=color, size=sz)

def bfont(bold=False, color="000000", sz=10):
    return Font(name="微软雅黑", bold=bold, color=color, size=sz)

def hfill(color=C_HEADER_BG):
    return PatternFill("solid", fgColor=color)

def ifill():
    return PatternFill("solid", fgColor=C_INPUT_BG)

def cfill():
    return PatternFill("solid", fgColor=C_CALC_BG)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def set_hdr(ws, row, col, value, bg=C_HEADER_BG, bold=True, wrap=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font  = Font(name="微软雅黑", bold=bold, color=C_HEADER_FG, size=10)
    c.fill  = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    c.border = THIN_BORDER
    return c

def set_inp(ws, row, col, value=None, fmt=None, note=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font  = Font(name="微软雅黑", color="00008B", size=10)  # dark blue = user input
    c.fill  = ifill()
    c.alignment = left()
    c.border = THIN_BORDER
    if fmt:
        c.number_format = fmt
    return c

def set_fml(ws, row, col, formula, fmt=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.font  = Font(name="微软雅黑", color="000000", size=10)  # black = formula
    c.fill  = cfill()
    c.alignment = left()
    c.border = THIN_BORDER
    if fmt:
        c.number_format = fmt
    return c

def set_txt(ws, row, col, value, bold=False, bg=None, fg="000000", wrap=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font  = Font(name="微软雅黑", bold=bold, color=fg, size=10)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    c.border = THIN_BORDER
    return c

def col_w(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def merge_hdr(ws, row, c1, c2, value, bg=C_HEADER_BG):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=value)
    c.font  = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=11)
    c.fill  = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = THIN_BORDER
    return c


# ===========================================================================
# Sheet 1: 使用说明
# ===========================================================================
def build_readme(wb):
    ws = wb.create_sheet("使用说明")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 60

    ws.row_dimensions[1].height = 36
    ws.merge_cells("A1:C1")
    t = ws.cell(1, 1, "仓库规划库存需求计算模板 — 使用说明")
    t.font  = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=14)
    t.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    stage_chain = "→".join(MANUFACTURING_PROCESS_STAGES)
    sections = [
        ("【工作簿结构】", [
            ("使用说明",     "本页面，说明各工作表用途及填写方法。"),
            ("参数设置",     "填写全局通用参数：年工作天数、安全系数、托盘尺寸、面积利用率等。"),
            ("产品SKU",      "逐行输入每个产品/SKU 的需求与生产参数（年需求量、最小批量、生产间隔、成品目标库存天数等）。"),
            ("原料BOM",      "填写产品-原料消耗矩阵，每行一种原料，列对应各SKU，输入单位消耗量及累计良率损耗系数。"),
            ("原料仓计算",   "自动按原料汇总日消耗量，并计算：\n设计库存 = MAX(日消耗×有效采购提前期, MOQ, MPQ)×安全系数 + 安全库存 + 下批备料库存 + 退料库存 + 待检库存 - 旧料可消耗量。"),
            ("过程品仓计算", "按工序位置和SKU 计算：\n设计WIP = 修正后基础WIP + 待检库存 + 尾批库存 + 不良/返工库存；其中修正后基础WIP考虑设备故障率与工序良率损失。"),
            ("成品仓计算",   "按SKU 计算：\n设计库存 = MAX(日需求×生产间隔, 日需求×目标库存天数, 最小生产批量) + 安全库存 + 返工补偿库存 + 待检放行库存。"),
            ("汇总",         "汇总三类仓库的设计库存量，并可选择计算托盘数和估算面积。"),
        ]),
        ("【填写说明】", [
            ("蓝色单元格",   "用户需要手动输入的参数（深蓝色字体、浅蓝色底色）。"),
            ("绿色单元格",   "由公式自动计算，请勿手动修改（黑色字体、浅绿色底色）。"),
            ("黄色单元格",   "当前模板未启用黄色专用输入样式，关键假设请优先检查参数设置中“需求与供应波动修正”分组。"),
            ("行数扩展",     "各计算表预留了足够行数（最多20行），需要更多行时可复制最后一行的公式向下填充。"),
        ]),
        ("【计算逻辑说明】", [
            ("原料仓",       "库存 = MAX(日均消耗×有效采购提前期天数, MOQ, MPQ)×安全系数 + 安全库存 + 下一批次备料 + 换产退料 + 待检库存 - 旧料可消耗量"),
            ("过程品仓",     "库存 = MAX(下游小时消耗×缓冲小时数, 批量大小×最大等待批次数) ÷(1-故障率)×(1+良率损失率) + 待检等待库存 + 尾批库存 + 不良品/返工/隔离库存"),
            ("成品仓",       "库存 = MAX(日需求量×生产间隔天数, 日需求量×目标库存天数, 最小生产批量) + 安全库存天数×日需求量 + 返工补偿库存 + 待检放行库存"),
            ("面积估算",     "仓库面积 = 库存量 ÷ 每托盘装载量 × 单托盘占地面积 ÷ 堆码层数 ÷ 面积利用率"),
        ]),
        ("【注意事项】", [
            ("多SKU换产",    "换产时会产生剩余退料（原料未用完退回），退料库存已在原料仓计算中单独考虑。"),
            ("良率损耗",     f"BOM 表中请填写单位成品实际原料消耗量（已包含{stage_chain}全流程良率损耗影响）。"),
            ("需求修正系数叠加", "Q3/Q4峰值、订单变更冗余、新品爬坡系数会乘法叠加，请结合历史波动谨慎设值，避免过度放大。"),
            ("动态安全库存", "当前模板暂按固定安全库存计算（未启用动态算法）；后续有数据后可在参数设置中启用。"),
            ("分层管理",     "当前模板未启用ABC/XYZ分层管理，建议后续有分层规则后再接入差异化安全库存。"),
            ("安全系数",     "参数设置中的安全系数会乘以基础库存量，建议原料仓1.1~1.3，过程品仓1.1~1.2，成品仓1.1~1.2。"),
            ("数量单位",     "本表不限定单位，请在各表表头的单位列中注明所使用的单位（件、箱、吨、kg等），并保持全表一致。"),
        ]),
    ]

    row = 3
    for section_title, items in sections:
        ws.row_dimensions[row].height = 28
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row, 1, section_title)
        c.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=11)
        c.fill = PatternFill("solid", fgColor=C_TITLE_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")
        row += 1
        for key, val in items:
            ws.row_dimensions[row].height = max(15, val.count('\n') * 15 + 18)
            set_txt(ws, row, 2, key, bold=True, bg="F0F8FF")
            c = ws.cell(row, 3, val)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.border = THIN_BORDER
            row += 1
        row += 1


# ===========================================================================
# Sheet 2: 参数设置
# ===========================================================================
def build_params(wb):
    ws = wb.create_sheet("参数设置")
    ws.sheet_view.showGridLines = False
    col_w(ws, 1, 4)
    col_w(ws, 2, 30)
    col_w(ws, 3, 18)
    col_w(ws, 4, 12)
    col_w(ws, 5, 30)

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:E1")
    t = ws.cell(1, 1, "参数设置 — 全局通用参数")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    headers_row = 2
    for col, hdr in enumerate(["", "参数名称", "数值", "单位", "说明"], 1):
        set_hdr(ws, headers_row, col, hdr)

    params = [
        # (name, value, unit, note)
        ("【生产基础参数】", None, None, None),
        ("年工作天数",         300,   "天/年",  "全年实际生产工作日，常用250~300天"),
        ("每天工作小时数",     20,    "小时/天", "含加班，用于计算小时产能"),
        ("年工作小时数",       "=参数设置!C4*参数设置!C5", "小时/年", "自动计算 = 年工作天数 × 每天工作小时数"),
        ("默认安全系数（原料仓）", 1.2, "—",    "库存量乘以该系数以应对波动，建议1.1~1.3"),
        ("默认安全系数（过程品仓）", 1.15, "—", "建议1.1~1.2"),
        ("默认安全系数（成品仓）", 1.2, "—",    "建议1.1~1.2"),

        ("【托盘与存储参数】", None, None, None),
        ("标准托盘长度",       1.2,   "m",      "常用1.0m或1.2m"),
        ("标准托盘宽度",       1.0,   "m",      "常用0.8m或1.0m"),
        ("单托盘占地面积",     "=参数设置!C11*参数设置!C12", "m²/托", "自动计算 = 长 × 宽"),
        ("地堆层数",           2,     "层",     "无货架时地面堆叠层数"),
        ("货架层数（如有）",   4,     "层",     "有货架时使用货架层数，无货架填0"),

        ("【面积利用率参数】", None, None, None),
        ("原料仓面积利用率",   0.45,  "—",      "通常35%~60%，含通道/叉车/消防等"),
        ("过程品仓面积利用率", 0.50,  "—",      "通常40%~60%"),
        ("成品仓面积利用率",   0.50,  "—",      "通常40%~60%"),

        ("【扩展余量】", None, None, None),
        ("面积扩展余量系数",   1.2,   "—",      "最终面积乘以该系数，预留10%~30%余量"),

        ("【需求与供应波动修正】", None, None, None),
        ("Q3/Q4需求峰值系数",      1.2,   "—",      "旺季需求放大系数，建议1.10~1.30（如仅部分SKU受影响，请在SKU层按需调整）"),
        ("订单变更/取消冗余系数",  1.03,  "—",      "可参考历史呆滞库存占比估算，建议1.00~1.10"),
        ("原料交期波动系数",       1.1,   "—",      "用于放大基础采购提前期，覆盖供应商波动"),
        ("春节停运附加天数",       18,    "天",     "节假日停运可按15~21天预估"),
        ("瓶颈设备故障率",         0.08,  "—",      "按瓶颈工序设备故障率估算产能波动"),
        ("石墨化良率损失率",       0.08,  "—",      "石墨化工序低良率补偿参数，建议5%~15%"),
        ("成品返工率",             0.08,  "—",      "成品工序返工比例参数，建议5%~10%"),
        ("新品爬坡系数",           1.05,  "—",      "新品爬坡阶段需求修正系数"),
        ("动态安全库存启用标记",   0,     "0/1",    "当前默认0=未启用，后续有数据后可切换"),
        ("ABC分层管理启用标记",    0,     "0/1",    "当前默认0=未启用，后续分层管理后可切换"),
    ]

    row = 3
    for item in params:
        name, val, unit, note = item
        ws.row_dimensions[row].height = 20
        if val is None:  # section header
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
            c = ws.cell(row, 2, name)
            c.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=10)
            c.fill = PatternFill("solid", fgColor=C_TITLE_BG)
            c.alignment = Alignment(horizontal="left", vertical="center")
        else:
            set_txt(ws, row, 2, name)
            if isinstance(val, str) and val.startswith("="):
                set_fml(ws, row, 3, val, fmt=NUM_FMT_DEC2)
            else:
                inp = set_inp(ws, row, 3, val)
                if isinstance(val, float) and val < 2:
                    inp.number_format = NUM_FMT_DEC2
                else:
                    inp.number_format = NUM_FMT_INT
            set_txt(ws, row, 4, unit or "")
            set_txt(ws, row, 5, note or "")
        row += 1

    # Named reference row map for other sheets to use (documented in note column)
    ws.cell(row + 1, 2, "★ 其他工作表通过[参数设置!C行号]引用本表数值，请勿删除或移动行。").font = Font(
        name="微软雅黑", color="FF0000", italic=True, size=9)


# 参数设置 row reference map (row 1=title, row 2=header, data starts row 3):
#   C4=年工作天数, C5=每天工作小时数, C6=年工作小时数,
#   C7=安全系数(原料), C8=安全系数(过程品), C9=安全系数(成品),
#   C11=托盘长, C12=托盘宽, C13=单托盘占地面积, C14=地堆层数, C15=货架层数,
#   C17=原料仓利用率, C18=过程品仓利用率, C19=成品仓利用率, C21=扩展余量系数,
#   C24=Q3/Q4需求峰值系数, C25=订单变更/取消冗余系数, C26=原料交期波动系数,
#   C27=春节停运附加天数, C28=瓶颈设备故障率, C29=石墨化良率损失率,
#   C30=成品返工率, C31=新品爬坡系数

def build_sku(wb):
    ws = wb.create_sheet("产品SKU")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:V1")
    t = ws.cell(1, 1, "产品SKU — 产品级输入参数")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        ("行号", 4),
        ("产品编号/SKU", 14),
        ("产品名称", 16),
        ("单位", 6),
        ("年需求量\n(件/年)", 12),
        ("月需求量\n(件/月)\n[自动]", 12),
        ("日需求量\n(件/天)\n[自动]", 12),
        ("最小生产\n批量(件)", 12),
        ("生产间隔\n(天)", 10),
        ("目标成品\n库存天数", 10),
        ("安全库存\n天数", 10),
        ("待检等待\n天数", 10),
        ("成品每托\n装载量(件)", 12),
        ("单件重量\n(kg)", 10),
    ]
    headers.extend([(f"过程品缓冲\n小时数\n({stage}后)", 12) for stage in MANUFACTURING_PROCESS_STAGES])
    headers.append(("备注", 18))

    col = 1
    for hdr, w in headers:
        set_hdr(ws, 2, col, hdr)
        col_w(ws, col, w)
        col += 1

    ws.row_dimensions[2].height = 48

    # Sample data rows（示例填充8个SKU，模板总计支持12个SKU）
    # 每行最后7个数值依次对应 MANUFACTURING_PROCESS_STAGES 顺序的缓冲小时数
    sample_skus = [
        ("SKU-001", "产品A", "件", 120000, "", "", 5000, 7, 14, 3, 2, 50, 0.5, 2, 3, 4, 5, 6, 7, 8),
        ("SKU-002", "产品B", "件", 96000, "", "", 4000, 7, 14, 3, 2, 40, 0.6, 2, 3, 4, 4, 5, 6, 7),
        ("SKU-003", "产品C", "件", 72000, "", "", 3000, 10, 15, 4, 2, 30, 0.8, 3, 4, 5, 5, 6, 7, 8),
        ("SKU-004", "产品D", "件", 60000, "", "", 3000, 10, 15, 4, 2, 60, 0.4, 2, 2, 3, 4, 5, 6, 7),
        ("SKU-005", "产品E", "件", 48000, "", "", 2000, 14, 20, 5, 3, 50, 0.5, 3, 4, 5, 6, 7, 8, 9),
        ("SKU-006", "产品F", "件", 36000, "", "", 2000, 14, 20, 5, 3, 40, 0.7, 2, 3, 3, 4, 5, 6, 7),
        ("SKU-007", "产品G", "件", 24000, "", "", 1500, 21, 21, 7, 3, 30, 1.0, 4, 5, 6, 7, 8, 9, 10),
        ("SKU-008", "产品H", "件", 18000, "", "", 1500, 21, 21, 7, 3, 25, 1.2, 3, 4, 5, 6, 6, 7, 8),
    ]

    for i, sku in enumerate(sample_skus):
        r = 3 + i
        ws.row_dimensions[r].height = 20
        fill = "F2F2F2" if i % 2 else "FFFFFF"
        ws.cell(r, 1, i + 1).font = bfont()

        name_code, name, unit, annual, mo, dy, minbatch, interval, tgt_days, safe_days, insp_days, pallet_qty, wt, *buffers = sku

        set_inp(ws, r, 2, name_code)
        set_inp(ws, r, 3, name)
        set_inp(ws, r, 4, unit)
        set_inp(ws, r, 5, annual, NUM_FMT_INT)
        # monthly = (annual / 12) × Q3/Q4 peak factor × order-change redundancy × ramp-up factor
        set_fml(ws, r, 6, f"=IFERROR((产品SKU!E{r}/12)*参数设置!$C$24*参数设置!$C$25*参数设置!$C$31,\"\")", NUM_FMT_DEC2)
        # daily = (annual / working days) × Q3/Q4 peak factor × order-change redundancy × ramp-up factor
        set_fml(ws, r, 7, f"=IFERROR((产品SKU!E{r}/参数设置!$C$4)*参数设置!$C$24*参数设置!$C$25*参数设置!$C$31,\"\")", NUM_FMT_DEC2)
        set_inp(ws, r, 8, minbatch, NUM_FMT_INT)
        set_inp(ws, r, 9, interval, NUM_FMT_INT)
        set_inp(ws, r, 10, tgt_days, NUM_FMT_INT)
        set_inp(ws, r, 11, safe_days, NUM_FMT_INT)
        set_inp(ws, r, 12, insp_days, NUM_FMT_INT)
        set_inp(ws, r, 13, pallet_qty, NUM_FMT_INT)
        set_inp(ws, r, 14, wt, NUM_FMT_DEC2)
        for offset, val in enumerate(buffers, start=15):
            set_inp(ws, r, offset, val, NUM_FMT_DEC2)
        set_inp(ws, r, 15 + NUM_PROCESSES, "")

    # Blank rows for extension
    for i in range(len(sample_skus), NUM_SKU_ROWS):
        r = 3 + i
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1, i + 1)
        for col in range(2, 16 + NUM_PROCESSES):
            set_inp(ws, r, col)
        # daily = (annual / working days) × Q3/Q4 peak factor × order-change redundancy × ramp-up factor
        set_fml(ws, r, 6, f"=IFERROR((产品SKU!E{r}/12)*参数设置!$C$24*参数设置!$C$25*参数设置!$C$31,\"\")", NUM_FMT_DEC2)
        set_fml(ws, r, 7, f"=IFERROR((产品SKU!E{r}/参数设置!$C$4)*参数设置!$C$24*参数设置!$C$25*参数设置!$C$31,\"\")", NUM_FMT_DEC2)


# ===========================================================================
# Sheet 4: 原料BOM
# ===========================================================================
def build_bom(wb):
    ws = wb.create_sheet("原料BOM")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells(f"A1:{get_column_letter(4 + NUM_SKU_ROWS)}1")
    t = ws.cell(1, 1, "原料BOM — 产品-原料消耗矩阵")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    # Sub-header: explanation row
    ws.merge_cells(f"A2:{get_column_letter(4 + NUM_SKU_ROWS)}2")
    note = ws.cell(2, 1, "填写说明：对应单元格填写每生产1件该SKU成品所消耗的该原料数量(已含良率损耗)，单位须与原料仓计算中一致；若该SKU不使用该原料，填0或留空。")
    note.font = Font(name="微软雅黑", italic=True, color="444444", size=9)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28

    # Header row 3: fixed columns + SKU columns
    fixed_hdrs = ["行号", "原料编号", "原料名称", "单位"]
    col_widths = [4, 14, 18, 8]
    for col, (h, w) in enumerate(zip(fixed_hdrs, col_widths), 1):
        set_hdr(ws, 3, col, h)
        col_w(ws, col, w)

    # SKU columns - reference product code from 产品SKU
    for sk in range(NUM_SKU_ROWS):
        col = 5 + sk
        sku_row = 3 + sk
        # Header: pull SKU code from 产品SKU sheet
        c = ws.cell(3, col, f"=IFERROR(产品SKU!B{sku_row},\"SKU-{sk+1:02d}\")")
        c.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=10)
        c.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER
        col_w(ws, col, 12)

    ws.row_dimensions[3].height = 36

    # Sample raw materials
    sample_rms = [
        ("RM-001", "原料甲", "kg",    [1.20, 1.10, 0.00, 0.00, 0.80, 0.80, 0.00, 0.00]),
        ("RM-002", "原料乙", "kg",    [0.50, 0.00, 0.60, 0.60, 0.00, 0.00, 0.40, 0.40]),
        ("RM-003", "原料丙", "m",     [0.30, 0.30, 0.30, 0.00, 0.30, 0.00, 0.30, 0.00]),
        ("RM-004", "原料丁", "片",    [2.00, 2.00, 2.00, 2.00, 0.00, 0.00, 0.00, 0.00]),
        ("RM-005", "辅料A",  "g",     [5.00, 5.00, 5.00, 5.00, 5.00, 5.00, 5.00, 5.00]),
        ("RM-006", "辅料B",  "mL",    [2.00, 2.00, 2.00, 2.00, 2.00, 2.00, 2.00, 2.00]),
        ("RM-007", "包装材料X", "个", [1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
        ("RM-008", "包装材料Y", "个", [0.00, 0.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
    ]

    for i, rm in enumerate(sample_rms):
        r = 4 + i
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1, i + 1).font = bfont()
        code, name, unit, consumptions = rm
        set_inp(ws, r, 2, code)
        set_inp(ws, r, 3, name)
        set_inp(ws, r, 4, unit)
        for sk, qty in enumerate(consumptions):
            col = 5 + sk
            set_inp(ws, r, col, qty if qty else None, NUM_FMT_DEC2)

    # Blank rows for extension
    for i in range(len(sample_rms), NUM_RM_ROWS):
        r = 4 + i
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1, i + 1)
        for col in range(2, 5):
            set_inp(ws, r, col)
        for sk in range(NUM_SKU_ROWS):
            set_inp(ws, r, 5 + sk, None, NUM_FMT_DEC2)


# ===========================================================================
# Sheet 5: 原料仓计算
# ===========================================================================
def build_rawmat(wb):
    ws = wb.create_sheet("原料仓计算")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:U1")
    t = ws.cell(1, 1, "原料仓计算 — 设计库存量")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    note_text = ("公式：设计库存 = MAX(日均消耗×有效采购提前期, MOQ, MPQ) × 安全系数 + 安全库存 + "
                 "下批备料库存 + 换产退料库存 + 待检库存 - 呆滞旧料可消耗量")
    ws.merge_cells("A2:U2")
    note = ws.cell(2, 1, note_text)
    note.font = Font(name="微软雅黑", italic=True, color="444444", size=9)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    headers = [
        ("行号",            4),
        ("原料编号",        14),
        ("原料名称",        18),
        ("单位",            6),
        ("日均总消耗量\n[自动]", 14),
        ("基础采购\n提前期(天)", 12),
        ("春节停运\n附加天数", 12),
        ("交期波动\n系数", 10),
        ("有效采购\n提前期(天)\n[自动]", 14),
        ("最小订购量\nMOQ",  12),
        ("最小到货批量\nMPQ", 12),
        ("安全系数",        10),
        ("安全库存量\n(件)", 12),
        ("下批备料\n库存量", 12),
        ("换产退料\n库存量", 12),
        ("待检库存量",       12),
        ("呆滞旧料\n可消耗量", 12),
        ("设计库存量\n[自动]", 14),
        ("每托盘装量",       10),
        ("设计托盘数\n[自动]", 12),
        ("估算占地面积\n(m²)[自动]", 14),
    ]

    col = 1
    for hdr, w in headers:
        set_hdr(ws, 3, col, hdr)
        col_w(ws, col, w)
        col += 1
    ws.row_dimensions[3].height = 48

    # For daily consumption: sum across all SKUs of (BOM consumption × daily demand)
    # BOM sheet: raw material i is at row 4+i, columns 5..5+NUM_SKU_ROWS-1
    # SKU daily demand is at 产品SKU!G(3+sk) for sk=0..NUM_SKU_ROWS-1

    # We build a helper formula for daily consumption of raw material row r_rm (4+i in BOM)
    # daily_consumption_i = SUMPRODUCT(原料BOM!E(bom_row):X(bom_row), 产品SKU!G3:G14)
    # where bom_row = 4+i for the i-th raw material

    bom_sku_start_col = get_column_letter(5)
    bom_sku_end_col   = get_column_letter(4 + NUM_SKU_ROWS)
    sku_daily_start   = "产品SKU!$G$3"
    sku_daily_end     = f"产品SKU!$G${2+NUM_SKU_ROWS}"  # G3:G14

    for i in range(NUM_RM_ROWS):
        r = 4 + i
        bom_row = 4 + i
        ws.row_dimensions[r].height = 20

        ws.cell(r, 1, i + 1).font = bfont()
        # Pull raw material info from BOM sheet
        set_fml(ws, r, 2, f"=IFERROR(原料BOM!B{bom_row},\"\")")
        set_fml(ws, r, 3, f"=IFERROR(原料BOM!C{bom_row},\"\")")
        set_fml(ws, r, 4, f"=IFERROR(原料BOM!D{bom_row},\"\")")

        # Daily total consumption = SUMPRODUCT of BOM row × SKU daily demands
        daily_fml = (f"=IFERROR(SUMPRODUCT("
                     f"原料BOM!{bom_sku_start_col}{bom_row}:{bom_sku_end_col}{bom_row},"
                     f"{sku_daily_start}:{sku_daily_end}),0)")
        set_fml(ws, r, 5, daily_fml, NUM_FMT_DEC2)

        # User inputs
        set_inp(ws, r, 6, DEFAULT_BASE_LEAD_DAYS if i < DEFAULT_RM_WITH_PRESET_LEAD else None, NUM_FMT_INT)  # base lead time
        set_fml(ws, r, 7, "=参数设置!$C$27", NUM_FMT_INT)          # spring festival holiday days
        set_fml(ws, r, 8, "=参数设置!$C$26", NUM_FMT_DEC2)         # lead-time variation factor
        # Effective lead time = (base lead time × supplier lead-time variation factor) + spring-festival shutdown days
        set_fml(ws, r, 9, f"=IFERROR((原料仓计算!F{r}*原料仓计算!H{r})+原料仓计算!G{r},0)", NUM_FMT_DEC2)
        set_inp(ws, r, 10, None, NUM_FMT_INT)                     # MOQ
        set_inp(ws, r, 11, None, NUM_FMT_INT)                     # MPQ
        set_fml(ws, r, 12, "=参数设置!$C$7", NUM_FMT_DEC2)        # safety factor
        set_inp(ws, r, 13, None, NUM_FMT_INT)                     # safety stock qty
        set_inp(ws, r, 14, None, NUM_FMT_INT)                     # next batch prep
        set_inp(ws, r, 15, None, NUM_FMT_INT)                     # changeover return stock
        set_inp(ws, r, 16, None, NUM_FMT_INT)                     # inspection pending
        set_inp(ws, r, 17, None, NUM_FMT_INT)                     # obsolete stock consumption offset

        # Design stock = MAX(daily×effective lead, MOQ, MPQ) × safety_factor + safety + next + return + inspection - obsolete use
        design_fml = (f"=IFERROR(MAX(0,"
                      f"MAX(原料仓计算!E{r}*原料仓计算!I{r},"
                      f"IF(原料仓计算!J{r}=\"\",0,原料仓计算!J{r}),"
                      f"IF(原料仓计算!K{r}=\"\",0,原料仓计算!K{r}))"
                      f"*原料仓计算!L{r}"
                      f"+IF(原料仓计算!M{r}=\"\",0,原料仓计算!M{r})"
                      f"+IF(原料仓计算!N{r}=\"\",0,原料仓计算!N{r})"
                      f"+IF(原料仓计算!O{r}=\"\",0,原料仓计算!O{r})"
                      f"+IF(原料仓计算!P{r}=\"\",0,原料仓计算!P{r})"
                      f"-IF(原料仓计算!Q{r}=\"\",0,原料仓计算!Q{r})),0)")
        set_fml(ws, r, 18, design_fml, NUM_FMT_DEC2)

        set_inp(ws, r, 19, None, NUM_FMT_INT)  # pallet capacity

        # Pallet count
        pallet_fml = (f"=IFERROR(CEILING(原料仓计算!R{r}/"
                      f"IF(原料仓计算!S{r}=0,1,原料仓计算!S{r}),1),0)")
        set_fml(ws, r, 20, pallet_fml, NUM_FMT_INT)

        # Estimated area (m²) = pallet_count × pallet_area / stack_layers / utilization × expansion
        # Parameter refs: C13=单托盘占地, C14=地堆层数, C17=原料仓利用率, C21=扩展系数
        area_fml = (f"=IFERROR(原料仓计算!T{r}*参数设置!$C$13"
                    f"/MAX(参数设置!$C$14,1)"
                    f"/参数设置!$C$17"
                    f"*参数设置!$C$21,0)")
        set_fml(ws, r, 21, area_fml, NUM_FMT_DEC2)

    # Total row
    total_r = 4 + NUM_RM_ROWS
    ws.row_dimensions[total_r].height = 22
    ws.merge_cells(f"A{total_r}:D{total_r}")
    merge_hdr(ws, total_r, 1, 4, "合计", C_TITLE_BG)
    set_fml(ws, total_r, 18, f"=SUM(R4:R{total_r-1})", NUM_FMT_DEC2)
    set_fml(ws, total_r, 20, f"=SUM(T4:T{total_r-1})", NUM_FMT_INT)
    set_fml(ws, total_r, 21, f"=SUM(U4:U{total_r-1})", NUM_FMT_DEC2)


# ===========================================================================
# Sheet 6: 过程品仓计算
# ===========================================================================
def build_wip(wb):
    ws = wb.create_sheet("过程品仓计算")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:S1")
    t = ws.cell(1, 1, "过程品仓计算 — 工序WIP设计库存量")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    note_text = ("公式：设计WIP = MAX(下游小时消耗×缓冲小时数, 批量×最大等待批次数)÷(1-设备故障率)×(1+良率损失率) "
                 "+ 待检库存 + 尾批库存 + 不良品/返工/隔离库存")
    ws.merge_cells("A2:S2")
    note = ws.cell(2, 1, note_text)
    note.font = Font(name="微软雅黑", italic=True, color="444444", size=9)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    headers = [
        ("行号",              4),
        ("工序位置",          12),
        ("产品SKU",           14),
        ("单位",              6),
        ("日产量\n(件/天)\n[自动]", 12),
        ("小时产量\n[自动]",  12),
        ("缓冲小时数",        12),
        ("按缓冲时间\n计算WIP\n[自动]", 14),
        ("批量大小\n(件)",    12),
        ("最大等待\n批次数",  12),
        ("按批量计算\nWIP\n[自动]", 12),
        ("修正后基础WIP\n[自动]", 14),
        ("设备故障率",        10),
        ("工序良率\n损失率", 10),
        ("待检等待\n库存",    12),
        ("尾批库存",          12),
        ("不良/返工/\n隔离库存", 12),
        ("设计WIP\n合计\n[自动]", 14),
        ("估算占地\n面积(m²)\n[自动]", 14),
    ]

    col = 1
    for hdr, w in headers:
        set_hdr(ws, 3, col, hdr)
        col_w(ws, col, w)
        col += 1
    ws.row_dimensions[3].height = 54

    # N process positions × up to NUM_SKU_ROWS SKUs
    processes = [
        f"{stage}后暂存（成品前）" if idx == NUM_PROCESSES - 1 else f"{stage}后暂存"
        for idx, stage in enumerate(MANUFACTURING_PROCESS_STAGES)
    ]
    # Buffer hour columns in 产品SKU start from col15
    buf_cols = list(range(15, 15 + NUM_PROCESSES))

    r = 4
    for proc_idx, (proc_name, buf_sku_col) in enumerate(zip(processes, buf_cols)):
        # Process section header
        ws.row_dimensions[r].height = 22
        ws.merge_cells(f"A{r}:S{r}")
        c = ws.cell(r, 1, f"【{proc_name}】")
        c.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=10)
        c.fill = PatternFill("solid", fgColor=C_TITLE_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")
        r += 1

        buf_col_letter = get_column_letter(buf_sku_col)

        for sk in range(NUM_SKU_ROWS):
            sku_row = 3 + sk  # row in 产品SKU sheet
            ws.row_dimensions[r].height = 20
            ws.cell(r, 1, sk + 1).font = bfont()

            set_fml(ws, r, 2, proc_name)
            # SKU code
            set_fml(ws, r, 3, f"=IFERROR(产品SKU!B{sku_row},\"\")")
            set_fml(ws, r, 4, f"=IFERROR(产品SKU!D{sku_row},\"\")")  # unit
            # Daily demand from SKU sheet G column (col7)
            set_fml(ws, r, 5, f"=IFERROR(产品SKU!G{sku_row},0)", NUM_FMT_DEC2)
            # Hourly = daily / working hours per day
            set_fml(ws, r, 6, f"=IFERROR(过程品仓计算!E{r}/参数设置!$C$5,0)", NUM_FMT_DEC2)
            # Buffer hours from SKU sheet
            set_fml(ws, r, 7, f"=IFERROR(产品SKU!{buf_col_letter}{sku_row},0)", NUM_FMT_DEC2)
            # WIP by buffer time = hourly × buffer hours
            set_fml(ws, r, 8, f"=IFERROR(过程品仓计算!F{r}*过程品仓计算!G{r},0)", NUM_FMT_DEC2)

            # Batch size and max waiting batches (user input)
            set_inp(ws, r, 9, None, NUM_FMT_INT)    # batch size
            set_inp(ws, r, 10, 2, NUM_FMT_INT)       # max waiting batches default 2
            # WIP by batch = batch_size × waiting_batches
            set_fml(ws, r, 11,
                    f"=IFERROR(IF(过程品仓计算!I{r}=\"\",0,过程品仓计算!I{r})"
                    f"*IF(过程品仓计算!J{r}=\"\",0,过程品仓计算!J{r}),0)",
                    NUM_FMT_DEC2)
            # Corrected base WIP = MAX(buffer, batch)/(1-fault)*(1+yield_loss)
            set_fml(ws, r, 12,
                    f"=IFERROR(MAX(过程品仓计算!H{r},过程品仓计算!K{r})"
                    f"/MAX(1-IF(过程品仓计算!M{r}=\"\",0,过程品仓计算!M{r}),{MIN_FAULT_DENOMINATOR})"
                    f"*(1+IF(过程品仓计算!N{r}=\"\",0,过程品仓计算!N{r})),0)",
                    NUM_FMT_DEC2)

            is_graphite_stage = proc_idx == GRAPHITIZATION_STAGE_INDEX
            if is_graphite_stage:
                set_fml(ws, r, 13, "=参数设置!$C$28", NUM_FMT_DEC2)
                set_fml(ws, r, 14, "=参数设置!$C$29", NUM_FMT_DEC2)
            else:
                # Non-graphitization stages default to 0 and remain editable for manual overrides.
                set_inp(ws, r, 13, 0, NUM_FMT_DEC2)
                set_inp(ws, r, 14, 0, NUM_FMT_DEC2)

            # User inputs: inspection, tail batch, defect/rework
            set_inp(ws, r, 15, None, NUM_FMT_INT)
            set_inp(ws, r, 16, None, NUM_FMT_INT)
            set_inp(ws, r, 17, None, NUM_FMT_INT)

            # Total design WIP
            set_fml(ws, r, 18,
                    f"=IFERROR(过程品仓计算!L{r}"
                    f"+IF(过程品仓计算!O{r}=\"\",0,过程品仓计算!O{r})"
                    f"+IF(过程品仓计算!P{r}=\"\",0,过程品仓计算!P{r})"
                    f"+IF(过程品仓计算!Q{r}=\"\",0,过程品仓计算!Q{r}),0)",
                    NUM_FMT_DEC2)

            # Estimated area – use成品 pallet qty for WIP (col13 in SKU sheet = M)
            # pallet qty for WIP: assume same as finished goods pallet qty (col13 SKU)
            # area = CEILING(WIP/pallet_qty, 1) × pallet_area / stack / utilization × expansion
            area_fml = (f"=IFERROR(CEILING(过程品仓计算!R{r}/"
                        f"MAX(产品SKU!M{sku_row},1),1)"
                        f"*参数设置!$C$13"
                        f"/MAX(参数设置!$C$14,1)"
                        f"/参数设置!$C$18"
                        f"*参数设置!$C$21,0)")
            set_fml(ws, r, 19, area_fml, NUM_FMT_DEC2)

            r += 1

        # Sub-total for this process
        sub_start = r - NUM_SKU_ROWS
        ws.row_dimensions[r].height = 20
        ws.merge_cells(f"A{r}:Q{r}")
        merge_hdr(ws, r, 1, 17, f"{proc_name} 小计", "305496")
        set_fml(ws, r, 18, f"=SUM(R{sub_start}:R{r-1})", NUM_FMT_DEC2)
        set_fml(ws, r, 19, f"=SUM(S{sub_start}:S{r-1})", NUM_FMT_DEC2)
        r += 1

    # Grand total: sum only the subtotal rows (every PROCESS_BLOCK_ROWS rows, at offset NUM_SKU_ROWS+1)
    ws.row_dimensions[r].height = 22
    ws.merge_cells(f"A{r}:Q{r}")
    merge_hdr(ws, r, 1, 17, "过程品仓 合计", C_TITLE_BG)
    # SUMPRODUCT + MOD 仅选择每个工序分块中的“小计行”：
    # 每个工序块长度为 PROCESS_BLOCK_ROWS（1行分组标题 + NUM_SKU_ROWS行数据 + 1行小计），
    # 小计行相对块起点偏移 NUM_SKU_ROWS+1，因此通过 MOD 条件筛出所有小计行求和。
    set_fml(ws, r, 18, f"=SUMPRODUCT((MOD(ROW(R4:R{r-1})-4,{PROCESS_BLOCK_ROWS})=({NUM_SKU_ROWS+1}))*R4:R{r-1})", NUM_FMT_DEC2)
    set_fml(ws, r, 19, f"=SUMPRODUCT((MOD(ROW(S4:S{r-1})-4,{PROCESS_BLOCK_ROWS})=({NUM_SKU_ROWS+1}))*S4:S{r-1})", NUM_FMT_DEC2)


# ===========================================================================
# Sheet 7: 成品仓计算
# ===========================================================================
def build_fg(wb):
    ws = wb.create_sheet("成品仓计算")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:Q1")
    t = ws.cell(1, 1, "成品仓计算 — 设计库存量")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    note_text = ("公式：设计库存 = MAX(日需求×生产间隔, 日需求×目标库存天数, 最小生产批量) "
                 "+ 安全库存(日需求×安全天数) + 返工补偿库存 + 待检放行库存")
    ws.merge_cells("A2:Q2")
    note = ws.cell(2, 1, note_text)
    note.font = Font(name="微软雅黑", italic=True, color="444444", size=9)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    headers = [
        ("行号",            4),
        ("产品编号/SKU",    14),
        ("产品名称",        16),
        ("单位",            6),
        ("日需求量\n[自动]", 12),
        ("生产间隔\n(天)\n[自动]", 12),
        ("目标库存\n天数\n[自动]", 12),
        ("最小生产\n批量\n[自动]",  12),
        ("按生产间隔\n计算库存\n[自动]", 14),
        ("按目标天数\n计算库存\n[自动]", 14),
        ("设计基础\n库存\n=MAX(...)\n[自动]", 14),
        ("安全库存\n[自动]", 12),
        ("返工率",          10),
        ("返工补偿\n库存\n[自动]", 12),
        ("待检放行\n库存",   12),
        ("设计库存\n合计\n[自动]", 14),
        ("估算占地\n面积(m²)\n[自动]", 14),
    ]

    col = 1
    for hdr, w in headers:
        set_hdr(ws, 3, col, hdr)
        col_w(ws, col, w)
        col += 1
    ws.row_dimensions[3].height = 54

    for i in range(NUM_SKU_ROWS):
        r = 4 + i
        sku_row = 3 + i
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1, i + 1).font = bfont()

        set_fml(ws, r, 2, f"=IFERROR(产品SKU!B{sku_row},\"\")")
        set_fml(ws, r, 3, f"=IFERROR(产品SKU!C{sku_row},\"\")")
        set_fml(ws, r, 4, f"=IFERROR(产品SKU!D{sku_row},\"\")")
        # Daily demand
        set_fml(ws, r, 5, f"=IFERROR(产品SKU!G{sku_row},0)", NUM_FMT_DEC2)
        # Production interval
        set_fml(ws, r, 6, f"=IFERROR(产品SKU!I{sku_row},0)", NUM_FMT_INT)
        # Target inventory days
        set_fml(ws, r, 7, f"=IFERROR(产品SKU!J{sku_row},0)", NUM_FMT_INT)
        # Min production batch
        set_fml(ws, r, 8, f"=IFERROR(产品SKU!H{sku_row},0)", NUM_FMT_INT)
        # By production interval
        set_fml(ws, r, 9, f"=IFERROR(成品仓计算!E{r}*成品仓计算!F{r},0)", NUM_FMT_DEC2)
        # By target days
        set_fml(ws, r, 10, f"=IFERROR(成品仓计算!E{r}*成品仓计算!G{r},0)", NUM_FMT_DEC2)
        # Base stock = MAX(interval, target days, min batch)
        set_fml(ws, r, 11,
                f"=IFERROR(MAX(成品仓计算!I{r},成品仓计算!J{r},成品仓计算!H{r}),0)",
                NUM_FMT_DEC2)
        # Safety stock = daily × safety_days
        set_fml(ws, r, 12,
                f"=IFERROR(成品仓计算!E{r}*产品SKU!K{sku_row}*参数设置!$C$9,0)",
                NUM_FMT_DEC2)
        # Rework rate
        set_fml(ws, r, 13, "=参数设置!$C$30", NUM_FMT_DEC2)
        # Rework compensation stock
        set_fml(ws, r, 14, f"=IFERROR(成品仓计算!K{r}*成品仓计算!M{r},0)", NUM_FMT_DEC2)
        # Inspection pending (user input)
        set_inp(ws, r, 15, None, NUM_FMT_INT)
        # Total design stock
        set_fml(ws, r, 16,
                f"=IFERROR(成品仓计算!K{r}+成品仓计算!L{r}"
                f"+成品仓计算!N{r}"
                f"+IF(成品仓计算!O{r}=\"\",0,成品仓计算!O{r}),0)",
                NUM_FMT_DEC2)
        # Area
        area_fml = (f"=IFERROR(CEILING(成品仓计算!P{r}/"
                    f"MAX(产品SKU!M{sku_row},1),1)"
                    f"*参数设置!$C$13"
                    f"/MAX(参数设置!$C$14,1)"
                    f"/参数设置!$C$19"
                    f"*参数设置!$C$21,0)")
        set_fml(ws, r, 17, area_fml, NUM_FMT_DEC2)

    # Total row
    total_r = 4 + NUM_SKU_ROWS
    ws.row_dimensions[total_r].height = 22
    ws.merge_cells(f"A{total_r}:D{total_r}")
    merge_hdr(ws, total_r, 1, 4, "合计", C_TITLE_BG)
    set_fml(ws, total_r, 16, f"=SUM(P4:P{total_r-1})", NUM_FMT_DEC2)
    set_fml(ws, total_r, 17, f"=SUM(Q4:Q{total_r-1})", NUM_FMT_DEC2)


# ===========================================================================
# Sheet 8: 汇总
# ===========================================================================
def build_summary(wb):
    ws = wb.create_sheet("汇总")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:J1")
    t = ws.cell(1, 1, "汇总 — 各仓库设计库存量与估算面积")
    t.font = Font(name="微软雅黑", bold=True, color=C_HEADER_FG, size=13)
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        ("仓库类别", 16),
        ("计算依据", 40),
        ("设计库存量\n(合计)", 16),
        ("库存单位", 10),
        ("设计托盘数\n(合计)", 14),
        ("估算占地\n面积(m²)", 14),
        ("含扩展余量\n面积(m²)\n[自动]", 14),
        ("估算库存\n单价(元)", 14),
        ("库存占用\n金额(元)\n[自动]", 14),
        ("备注", 24),
    ]
    col = 1
    for hdr, w in headers:
        set_hdr(ws, 2, col, hdr)
        col_w(ws, col, w)
        col += 1
    ws.row_dimensions[2].height = 42

    # Raw material warehouse total row reference
    rm_total_row = 4 + NUM_RM_ROWS
    # WIP grand total row calculation
    # each process uses (1 section header + NUM_SKU_ROWS data rows + 1 subtotal) = PROCESS_BLOCK_ROWS
    # grand total is at row = 4 + NUM_PROCESSES*PROCESS_BLOCK_ROWS
    wip_grand_row = 4 + NUM_PROCESSES * PROCESS_BLOCK_ROWS
    fg_total_row  = 4 + NUM_SKU_ROWS

    rows = [
        ("原料仓",     "MAX(日消耗×采购提前期, MOQ)×安全系数 + 安全库存 + 下批备料 + 退料 + 待检",
         f"=原料仓计算!R{rm_total_row}", "（各原料单位不同）",
         f"=原料仓计算!T{rm_total_row}", f"=原料仓计算!U{rm_total_row}"),
    ]
    for stage in MANUFACTURING_PROCESS_STAGES:
        rows.append((
            f"过程品仓\n（{stage}后）",
            "MAX(下游小时消耗×缓冲小时数, 批量×等待批次) + 待检 + 尾批 + 不良品",
            None,
            "件",
            None,
            None,
        ))
    rows.extend([
        ("过程品仓 合计", f"见过程品仓计算!R{wip_grand_row}",
         f"=过程品仓计算!R{wip_grand_row}", "件",
         None, f"=过程品仓计算!S{wip_grand_row}"),
        ("成品仓",     "MAX(日需求×生产间隔, 日需求×目标天数, 最小批量) + 安全库存 + 返工补偿 + 待检放行",
         f"=成品仓计算!P{fg_total_row}", "件",
         None, f"=成品仓计算!Q{fg_total_row}"),
    ])

    # Find WIP per-process subtotal rows
    proc_subtotal_rows = []
    for proc_idx in range(NUM_PROCESSES):
        # section header at: 4 + proc_idx*PROCESS_BLOCK_ROWS, data rows follow, subtotal at +NUM_SKU_ROWS+1
        subtotal_r = 4 + proc_idx * PROCESS_BLOCK_ROWS + NUM_SKU_ROWS + 1
        proc_subtotal_rows.append(subtotal_r)

    for proc_idx, stage in enumerate(MANUFACTURING_PROCESS_STAGES):
        rows[1 + proc_idx] = (
            f"过程品仓\n（{stage}后）",
            "MAX(下游小时消耗×缓冲小时数, 批量×等待批次)÷(1-故障率)×(1+良率损失率) + 待检 + 尾批 + 不良品",
            f"=过程品仓计算!R{proc_subtotal_rows[proc_idx]}",
            "件",
            None,
            f"=过程品仓计算!S{proc_subtotal_rows[proc_idx]}",
        )

    for i, (wh, basis, qty_fml, unit, pallet_fml, area_fml) in enumerate(rows):
        r = 3 + i
        ws.row_dimensions[r].height = 24
        bg = "F2F2F2" if i % 2 else "FFFFFF"
        is_total = "合计" in wh

        set_txt(ws, r, 1, wh, bold=is_total, bg="E8F4FD" if is_total else bg)
        set_txt(ws, r, 2, basis)
        if qty_fml:
            set_fml(ws, r, 3, qty_fml, NUM_FMT_DEC2)
        else:
            set_inp(ws, r, 3, None, NUM_FMT_DEC2)
        set_txt(ws, r, 4, unit)
        if pallet_fml:
            set_fml(ws, r, 5, pallet_fml, NUM_FMT_INT)
        else:
            set_inp(ws, r, 5, None, NUM_FMT_INT)
        if area_fml:
            set_fml(ws, r, 6, area_fml, NUM_FMT_DEC2)
        else:
            set_inp(ws, r, 6, None, NUM_FMT_DEC2)
        # With expansion factor
        set_fml(ws, r, 7,
                f"=IFERROR(汇总!F{r}*参数设置!$C$21,0)",
                NUM_FMT_DEC2)
        set_inp(ws, r, 8, None, NUM_FMT_DEC2)
        set_fml(ws, r, 9, f"=IFERROR(汇总!C{r}*汇总!H{r},0)", NUM_FMT_DEC2)
        set_txt(ws, r, 10, "")

    # Grand total
    total_r = 3 + len(rows)
    ws.row_dimensions[total_r].height = 26
    ws.merge_cells(f"A{total_r}:B{total_r}")
    merge_hdr(ws, total_r, 1, 2, "三类仓库合计（不含辅助区域）", C_HEADER_BG)
    raw_r = 3
    wip_total_r = 3 + 1 + NUM_PROCESSES  # row index for "过程品仓 合计"
    fg_r = wip_total_r + 1
    set_fml(ws, total_r, 7,
            f"=IFERROR(汇总!G{raw_r}+汇总!G{wip_total_r}+汇总!G{fg_r},0)",
            NUM_FMT_DEC2)
    set_fml(ws, total_r, 9,
            f"=IFERROR(汇总!I{raw_r}+汇总!I{wip_total_r}+汇总!I{fg_r},0)",
            NUM_FMT_DEC2)

    # Auxiliary area note
    aux_r = total_r + 2
    ws.merge_cells(f"A{aux_r}:J{aux_r}")
    c = ws.cell(aux_r, 1,
                "⚠ 上述面积为存储区估算，实际仓库还需加上：收发货区、质检区、不良品隔离区、"
                "退换货区、包装区、通道等辅助区域，通常再增加30%~50%的辅助面积。")
    c.font = Font(name="微软雅黑", color="C00000", size=10, italic=True)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[aux_r].height = 36

    # Parameter reference reminder
    ref_r = aux_r + 2
    ws.merge_cells(f"A{ref_r}:J{ref_r}")
    c = ws.cell(ref_r, 1, "★ 面积估算关键参数(可在[参数设置]工作表中修改)：")
    c.font = Font(name="微软雅黑", bold=True, size=10)
    c.alignment = Alignment(horizontal="left", vertical="center")

    param_refs = [
        (ref_r + 1, "单托盘占地面积", "=参数设置!C13&\" m²\""),
        (ref_r + 2, "地堆层数",       "=参数设置!C14&\" 层\""),
        (ref_r + 3, "原料仓面积利用率", "=TEXT(参数设置!C17,\"0.0%\")"),
        (ref_r + 4, "过程品仓面积利用率", "=TEXT(参数设置!C18,\"0.0%\")"),
        (ref_r + 5, "成品仓面积利用率",   "=TEXT(参数设置!C19,\"0.0%\")"),
        (ref_r + 6, "扩展余量系数",    "=参数设置!C21"),
    ]
    col_w(ws, 1, 28)
    col_w(ws, 2, 40)
    col_w(ws, 3, 18)
    col_w(ws, 4, 12)
    col_w(ws, 5, 14)
    col_w(ws, 6, 14)
    col_w(ws, 7, 16)
    col_w(ws, 8, 14)
    col_w(ws, 9, 16)
    col_w(ws, 10, 24)

    for pr_row, label, fml in param_refs:
        ws.row_dimensions[pr_row].height = 18
        set_txt(ws, pr_row, 1, label)
        set_fml(ws, pr_row, 2, fml)


# ===========================================================================
# Main entry point
# ===========================================================================
def main():
    wb = Workbook()
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    build_readme(wb)
    build_params(wb)
    build_sku(wb)
    build_bom(wb)
    build_rawmat(wb)
    build_wip(wb)
    build_fg(wb)
    build_summary(wb)

    # Tab colors
    tab_colors = {
        "使用说明":     "1F4E79",
        "参数设置":     "2E75B6",
        "产品SKU":      "70AD47",
        "原料BOM":      "ED7D31",
        "原料仓计算":   "FFC000",
        "过程品仓计算": "9DC3E6",
        "成品仓计算":   "A9D18E",
        "汇总":         "FF0000",
    }
    for sheet_name, color in tab_colors.items():
        if sheet_name in wb.sheetnames:
            wb[sheet_name].sheet_properties.tabColor = color

    out_dir = os.path.dirname(__file__)
    out_path = os.path.join(out_dir, "仓库规划库存需求计算模板.xlsx")
    out_path_7stage = os.path.join(out_dir, "仓库规划库存需求计算模板_七道工序版.xlsx")
    wb.save(out_path)
    # 按业务要求额外提供同内容副本，便于按“七道工序版”文件名直接分发使用
    wb.save(out_path_7stage)
    print(f"✅ 已生成: {out_path}")
    print(f"✅ 已生成: {out_path_7stage}")
    return out_path_7stage


if __name__ == "__main__":
    main()
