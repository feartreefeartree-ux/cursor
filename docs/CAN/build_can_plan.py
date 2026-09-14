#!/usr/bin/env python3
"""Convert 高压风扇开发计划.xlsx 工时(人天) into start/end dates.

Source: the uploaded WBS (编号 / 阶段 / 任务 / 人天 / 负责人 / 交付物).
Window: 2026-06-15 .. 2026-10-30, weekdays only, minus 端午/中秋/国庆.
Critical path (105 人天) is scaled onto 93 working days.
Phase 5 (信息安全文档, 7 人天) runs in parallel after design review.
"""

from datetime import date, timedelta
from shutil import copyfile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

PLAN_START = date(2026, 6, 15)
PLAN_END = date(2026, 10, 30)
AS_OF = date(2026, 9, 14)

HOLIDAYS = {
    date(2026, 6, 19),  # 端午
    date(2026, 9, 25),  # 中秋
    date(2026, 10, 1),
    date(2026, 10, 2),
    date(2026, 10, 3),
    date(2026, 10, 4),
    date(2026, 10, 5),
    date(2026, 10, 6),
    date(2026, 10, 7),
}

# (code, phase, task, pd, owner, deliverable, path)
# path: "main" sequential on the embedded/test delivery path; "p5" parallel docs
ROWS = [
    ("1.1", "需求分析与评审", "客户资料分析与需求拆解", 5, "系统工程师", "D01 系统需求规格书", "main"),
    ("1.2", "需求分析与评审", "需求拆解表编制", 2, "系统工程师", "D02 需求拆解表", "main"),
    ("1.3", "需求分析与评审", "需求评审会议", 1, "项目经理+系统工程师", "D03 需求评审记录表", "main"),
    ("1.4", "需求分析与评审", "芯片可行性确认", 1, "系统工程师", "芯片可行性分析报告", "main"),
    ("2.1", "架构与详细设计", "软件架构设计", 2, "系统工程师", "D04 软件架构设计文档", "main"),
    ("2.2", "架构与详细设计", "Bootloader 详细设计", 2, "嵌入式开发工程师", "D05 Bootloader 详细设计", "main"),
    ("2.3", "架构与详细设计", "UDS 诊断栈详细设计", 3, "嵌入式开发工程师", "D06 UDS 诊断栈详细设计", "main"),
    ("2.4", "架构与详细设计", "安全访问模块设计", 2, "嵌入式开发工程师", "D07 安全访问设计", "main"),
    ("2.5", "架构与详细设计", "Flash 分区与存储管理设计", 2, "嵌入式开发工程师", "D08 Flash 分区设计", "main"),
    ("2.6", "架构与详细设计", "CAN 通信设计", 2, "嵌入式开发工程师", "D09 CAN 通信设计", "main"),
    ("2.7", "架构与详细设计", "设计内部评审", 1, "项目团队", "设计评审记录", "main"),
    ("3.1", "软件开发 — Bootloader", "MCU 底层初始化及配置", 3, "嵌入式开发工程师", "凌鸥芯片适配", "main"),
    ("3.2", "软件开发 — Bootloader", "CAN 驱动开发", 3, "嵌入式开发工程师", "初始化/收发/滤波/中断", "main"),
    ("3.3", "软件开发 — Bootloader", "CAN-TP 传输层", 5, "嵌入式开发工程师", "单帧/多帧/流控/定时", "main"),
    ("3.4", "软件开发 — Bootloader", "UDS 诊断栈", 8, "嵌入式开发工程师", "10个SID服务实现", "main"),
    ("3.5", "软件开发 — Bootloader", "安全信息访问", 5, "嵌入式开发工程师", "含算法验证和测试向量比对", "main"),
    ("3.6", "软件开发 — Bootloader", "Flash 擦写驱动 + CRC32", 3, "嵌入式开发工程师", "含硬件CRC模块调用", "main"),
    ("3.7", "软件开发 — Bootloader", "刷写状态机 + NRC + 跳步约束", 4, "嵌入式开发工程师", "16条跳步规则实现", "main"),
    ("3.8", "软件开发 — Bootloader", "Bootloader 集成调试", 5, "嵌入式开发工程师", "端到端刷写流程联调", "main"),
    ("3.9", "软件开发 — Application", "App 初始化 + CAN 驱动", 2, "嵌入式开发工程师", "部分代码共用", "main"),
    ("3.10", "软件开发 — Application", "UDS 诊断栈", 5, "嵌入式开发工程师", "完整诊断服务", "main"),
    ("3.11", "软件开发 — Application", "DID 管理模块", 2, "嵌入式开发工程师", "10+ 个 DID 读写", "main"),
    ("3.12", "软件开发 — Application", "DTC 管理模块", 2, "嵌入式开发工程师", "0x14/0x19 服务", "main"),
    ("3.13", "软件开发 — Application", "编程前条件检查", 2, "嵌入式开发工程师", "CAN 信号接收+条件判断", "main"),
    ("3.14", "软件开发 — Application", "风扇控制功能集成", 6, "嵌入式开发工程师", "PWM+ADC", "main"),
    ("3.15", "软件开发 — Application", "Application 集成调试", 4, "嵌入式开发工程师", "端到端诊断功能联调", "main"),
    ("4.1", "测试验证", "单元测试（关键模块）", 3, "开发+测试", "AES/CRC/Flash/CAN-TP", "main"),
    ("4.2", "测试验证", "Bootloader 刷写流程测试", 2, "测试工程师", "正常+16条异常路径", "main"),
    ("4.3", "测试验证", "诊断功能测试", 3, "测试工程师", "全SID/全会话/全NRC", "main"),
    ("4.4", "测试验证", "安全访问测试", 2, "测试工程师", "算法正确性+锁定机制", "main"),
    ("4.5", "测试验证", "BT↔App 集成测试", 3, "测试工程师", "刷写闭环+会话切换", "main"),
    ("4.6", "测试验证", "Bug 修复与回归测试", 5, "开发工程师", "问题修复验证", "main"),
    ("5.1", "信息安全文档", "信息安全技术规范填充", 2, "系统工程师", "D10 CS100120-027C", "p5"),
    ("5.2", "信息安全文档", "信息安全测试用例编写", 2, "测试工程师", "D11 CS100120-030C", "p5"),
    ("5.3", "信息安全文档", "CIA 接口协议确认与修订", 2, "项目经理", "D12 CS100120-045C", "p5"),
    ("5.4", "信息安全文档", "信息安全评审", 1, "项目团队", "评审记录", "p5"),
    ("6.1", "交付与客户支持", "版本发布说明编写", 1, "开发工程师", "D17 版本发布说明", "main"),
    ("6.2", "交付与客户支持", "交付评审", 1, "项目团队+客户", "评审记录", "main"),
    ("6.3", "交付与客户支持", "客户侧验证技术支持", 3, "开发工程师", "现场/远程支持", "main"),
]

SUBTOTALS = [
    ("阶段一 小计", "需求分析与评审"),
    ("阶段二 小计", "架构与详细设计"),
    ("BT 开发 小计", "软件开发 — Bootloader"),
    ("App 开发 小计", "软件开发 — Application"),
    ("阶段三 合计", ("软件开发 — Bootloader", "软件开发 — Application")),
    ("阶段四 小计", "测试验证"),
    ("阶段五 小计", "信息安全文档"),
    ("阶段六 小计", "交付与客户支持"),
    ("项目合计", None),
]

PHASE_FILL = {
    "需求分析与评审": "D6EAF8",
    "架构与详细设计": "D5F5E3",
    "软件开发 — Bootloader": "FDEBD0",
    "软件开发 — Application": "FCF3CF",
    "测试验证": "E8DAEF",
    "信息安全文档": "FADBD8",
    "交付与客户支持": "D6DBDF",
}


def working_days(start, end):
    days = []
    d = start
    while d <= end:
        if d.weekday() < 5 and d not in HOLIDAYS:
            days.append(d)
        d += timedelta(days=1)
    return days


def allocate(weights, seats, min_each=1):
    n = len(weights)
    remaining = seats - min_each * n
    extra_w = [w - min_each for w in weights]
    total = sum(extra_w)
    exact = [min_each + remaining * ew / total for ew in extra_w]
    base = [int(x) for x in exact]
    leftover = seats - sum(base)
    fracs = sorted(((exact[i] - base[i], -weights[i], i) for i in range(n)), reverse=True)
    for k in range(leftover):
        base[fracs[k][2]] += 1
    if sum(base) != seats:
        raise RuntimeError(f"allocation {sum(base)} != {seats}")
    return base


def schedule(wdays):
    main = [r for r in ROWS if r[6] == "main"]
    seats = allocate([r[3] for r in main], len(wdays))
    dates = {}
    cur = 0
    p2_end_idx = None
    for row, n in zip(main, seats):
        dates[row[0]] = (wdays[cur], wdays[cur + n - 1], n)
        if row[0] == "2.7":
            p2_end_idx = cur + n
        cur += n
    if p2_end_idx is None:
        raise RuntimeError("phase 2 end not found")
    cur = p2_end_idx
    for row in (r for r in ROWS if r[6] == "p5"):
        n = row[3]  # 7 人天, parallel, no compress
        dates[row[0]] = (wdays[cur], wdays[cur + n - 1], n)
        cur += n
    return dates


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(size=10, bold=False, color="000000"):
    return Font(name="微软雅黑", size=size, bold=bold, color=color)


def align(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


THIN = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)


def put(ws, row, col, value, *, bg=None, bold=False, size=10, h="center", color="000000"):
    cell = ws.cell(row, col, value)
    cell.font = font(size=size, bold=bold, color=color)
    cell.alignment = align(h)
    cell.border = THIN
    if bg:
        cell.fill = fill(bg)
    return cell


def put_date(ws, row, col, d, *, bg=None, bold=False):
    cell = put(ws, row, col, d, bg=bg, bold=bold)
    cell.number_format = "YYYY-MM-DD"
    return cell


def merge_row(ws, row, c1, c2, value, **kwargs):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = put(ws, row, c1, value, **kwargs)
    for col in range(c1 + 1, c2 + 1):
        c = ws.cell(row, col)
        c.border = THIN
        if kwargs.get("bg"):
            c.fill = fill(kwargs["bg"])
        c.alignment = align(kwargs.get("h", "center"))
    return cell


def build():
    wdays = working_days(PLAN_START, PLAN_END)
    dates = schedule(wdays)

    wb = Workbook()
    ws = wb.active
    ws.title = "工作任务分解"

    widths = {"A": 8, "B": 22, "C": 28, "D": 12, "E": 12, "F": 20, "G": 26, "H": 10, "I": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    merge_row(ws, 1, 1, 9, "高压风扇开发计划", bg="1F4E79", bold=True, size=16, color="FFFFFF")
    ws.row_dimensions[1].height = 28

    merge_row(
        ws,
        2,
        1,
        9,
        "工时已改成日期    窗口 2026-06-15～2026-10-30    编号 FAN_F-SW-100    版本 V1.1    密级：内部    只填黄色格",
        bg="2E75B6",
        size=9,
        color="FFFFFF",
        h="left",
    )
    ws.row_dimensions[2].height = 18

    put(ws, 3, 1, "供应商", bg="D6DCE4", bold=True)
    merge_row(ws, 3, 2, 3, None, bg="FFF2CC")
    put(ws, 3, 4, "零件号", bg="D6DCE4", bold=True)
    merge_row(ws, 3, 5, 6, "FAN_F / ", bg="FFF2CC", h="left")
    put(ws, 3, 7, "编制 / 日期", bg="D6DCE4", bold=True)
    merge_row(ws, 3, 8, 9, None, bg="FFF2CC")

    put(ws, 4, 1, "审核", bg="D6DCE4", bold=True)
    merge_row(ws, 4, 2, 3, None, bg="FFF2CC")
    put(ws, 4, 4, "批准", bg="D6DCE4", bold=True)
    merge_row(ws, 4, 5, 6, None, bg="FFF2CC")
    put(ws, 4, 7, "福田交样日", bg="D6DCE4", bold=True)
    merge_row(ws, 4, 8, 9, None, bg="FFF2CC")
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 20

    merge_row(
        ws,
        5,
        1,
        9,
        "原表 112 人天；主路径 105 人天铺进 93 个工作日（已扣周末、端午 6/19、中秋 9/25、国庆 10/1–10/7），按人天比例压缩约 11%。"
        "阶段五与开发并行，不占嵌入式主路径。负责人列保留（岗位）；姓名不必每行都写，接口人见下。"
        "最右「人天」仅对内，不要发给主机厂。",
        bg="FFF2CC",
        size=9,
        h="left",
    )
    ws.row_dimensions[5].height = 36

    headers = [
        "编号",
        "阶段名称",
        "任务名称",
        "开始日期",
        "结束日期",
        "负责人",
        "交付物/说明",
        "人天(对内)",
        "状态",
    ]
    for col, h in enumerate(headers, 1):
        put(ws, 6, col, h, bg="D6DCE4", bold=True)
    ws.row_dimensions[6].height = 22
    ws.freeze_panes = "A7"
    ws.auto_filter.ref = "A6:I6"

    r = 7
    phase_rows = {p: [] for p in PHASE_FILL}
    task_status_cells = []

    for code, phase, task, pd, owner, deliverable, path in ROWS:
        start, end, _n = dates[code]
        bg = PHASE_FILL[phase]
        put(ws, r, 1, code, bg=bg)
        put(ws, r, 2, phase, bg=bg)
        put(ws, r, 3, task, bg=bg, h="left")
        put_date(ws, r, 4, start, bg=bg)
        put_date(ws, r, 5, end, bg=bg)
        put(ws, r, 6, owner, bg=bg)
        put(ws, r, 7, deliverable, bg=bg, h="left")
        put(ws, r, 8, pd, bg=bg)
        put(ws, r, 9, "未开始", bg="FFF2CC")
        ws.row_dimensions[r].height = 22
        phase_rows[phase].append(r)
        task_status_cells.append(r)
        r += 1
        # subtotal immediately after last task of a group, except 阶段三合计 after App
        last_of = {
            "1.4": "阶段一 小计",
            "2.7": "阶段二 小计",
            "3.8": "BT 开发 小计",
            "3.15": "App 开发 小计",
            "4.6": "阶段四 小计",
            "5.4": "阶段五 小计（与开发并行）",
            "6.3": "阶段六 小计",
        }
        if code in last_of:
            r = write_subtotal(ws, r, last_of[code], phase_rows, dates)
            if code == "3.15":
                r = write_subtotal(ws, r, "阶段三 合计", phase_rows, dates)
            if code == "6.3":
                r = write_subtotal(ws, r, "项目合计", phase_rows, dates)

    # merge phase name cells
    for phase, rows in phase_rows.items():
        if len(rows) >= 2:
            ws.merge_cells(start_row=rows[0], start_column=2, end_row=rows[-1], end_column=2)
            ws.cell(rows[0], 2).alignment = align()

    dv = DataValidation(type="list", formula1='"未开始,进行中,已完成,暂停,取消"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"I{task_status_cells[0]}:I{task_status_cells[-1]}")

    # contacts
    r += 1
    merge_row(ws, r, 1, 9, "接口人（交给福田时填姓名和电话；对内可只保留上面的岗位）", bg="1F4E79", bold=True, color="FFFFFF", h="left")
    ws.row_dimensions[r].height = 18
    r += 1
    for col, h in enumerate(["角色", "姓名", "电话", "邮箱", "备注"], 1):
        put(ws, r, col, h, bg="D6DCE4", bold=True)
    merge_row(ws, r, 5, 9, "备注", bg="D6DCE4", bold=True)
    ws.cell(r, 5).value = "备注"
    r += 1
    contacts = [
        ("项目经理", "进度 / CIA / 对福田"),
        ("系统工程师", "需求与设计"),
        ("嵌入式开发工程师", "Bootloader / App / CAN"),
        ("测试工程师", "刷写与诊断测试"),
    ]
    for role, note in contacts:
        put(ws, r, 1, role, bold=True)
        put(ws, r, 2, None, bg="FFF2CC")
        put(ws, r, 3, None, bg="FFF2CC")
        put(ws, r, 4, None, bg="FFF2CC")
        merge_row(ws, r, 5, 9, note, h="left")
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1
    merge_row(ws, r, 1, 9, "变更记录", bg="1F4E79", bold=True, color="FFFFFF", h="left")
    r += 1
    put(ws, r, 1, "版本", bg="D6DCE4", bold=True)
    put(ws, r, 2, "日期", bg="D6DCE4", bold=True)
    merge_row(ws, r, 3, 9, "说明", bg="D6DCE4", bold=True, h="left")
    ws.cell(r, 3).value = "说明"
    r += 1
    put(ws, r, 1, "V1.1")
    put_date(ws, r, 2, AS_OF)
    merge_row(
        ws,
        r,
        3,
        9,
        "按上传的《高压风扇开发计划.xlsx》把「工时(人天)」改成开始/结束日期；负责人列保留；补状态、签署、接口人。年份按 2026。",
        h="left",
    )
    ws.row_dimensions[r].height = 28

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_title_rows = "1:6"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "1F4E79"

    notes = wb.create_sheet("排期说明")
    notes.column_dimensions["A"].width = 8
    notes.column_dimensions["B"].width = 22
    notes.column_dimensions["C"].width = 90
    merge_row(notes, 1, 1, 3, "日期是怎么从工时算出来的", bg="1F4E79", bold=True, size=14, color="FFFFFF", h="left")
    notes.row_dimensions[1].height = 24
    put(notes, 2, 1, "序号", bg="D6DCE4", bold=True)
    put(notes, 2, 2, "项", bg="D6DCE4", bold=True)
    put(notes, 2, 3, "说明", bg="D6DCE4", bold=True)
    rules = [
        ("窗口", "2026-06-15（周一）至 2026-10-30（周五）。原表未写年，按今天 2026-09-14 取 2026。"),
        ("工作日", "只排周一到周五；不排 2026-06-19 端午、2026-09-25 中秋、2026-10-01～10-07 国庆。窗口内共 93 个工作日。"),
        ("主路径", "阶段一～四 + 阶段六共 105 人天、35 条任务，按人天比例铺满 93 个工作日（每条至少 1 天，最大余数法）。所以 8 人天的 UDS 会变成 7 个工作日，不是 1:1。"),
        ("阶段五", "信息安全文档 7 人天不进主路径：设计评审结束后并行（系统/项目/测试），人天不压缩。"),
        ("负责人", "原列保留，写的是岗位，这符合主机厂进度表。交给福田再在「接口人」补姓名、电话、邮箱。不必做 RASIC。"),
        ("人天列", "只对内核对工作量。对客打印可隐藏 H 列。"),
        ("状态", "先填「未开始」，按实际改。4.2 在国庆前结束，4.3 从 10-08 接着做。"),
    ]
    for i, (item, text) in enumerate(rules, 1):
        put(notes, 2 + i, 1, i)
        put(notes, 2 + i, 2, item, bg="FFF2CC", bold=True)
        put(notes, 2 + i, 3, text, h="left")
        notes.row_dimensions[2 + i].height = 36
    notes.sheet_view.showGridLines = False
    notes.page_setup.orientation = "landscape"
    notes.page_setup.fitToPage = True
    notes.sheet_properties.pageSetUpPr.fitToPage = True
    notes.sheet_properties.tabColor = "C65911"

    out1 = "/workspace/docs/CAN/高压风扇开发计划.xlsx"
    out2 = "/workspace/docs/CAN/FAN_F-SW-100_高压风扇CAN开发计划.xlsx"
    wb.save(out1)
    copyfile(out1, out2)
    return out1, dates, wdays


def write_subtotal(ws, r, label, phase_rows, dates):
    if label == "阶段三 合计":
        codes = [x[0] for x in ROWS if x[1] in ("软件开发 — Bootloader", "软件开发 — Application")]
        pd = sum(x[3] for x in ROWS if x[1] in ("软件开发 — Bootloader", "软件开发 — Application"))
    elif label == "项目合计":
        codes = [x[0] for x in ROWS]
        pd = sum(x[3] for x in ROWS)
    else:
        phase = {
            "阶段一 小计": "需求分析与评审",
            "阶段二 小计": "架构与详细设计",
            "BT 开发 小计": "软件开发 — Bootloader",
            "App 开发 小计": "软件开发 — Application",
            "阶段四 小计": "测试验证",
            "阶段五 小计（与开发并行）": "信息安全文档",
            "阶段六 小计": "交付与客户支持",
        }[label]
        codes = [x[0] for x in ROWS if x[1] == phase]
        pd = sum(x[3] for x in ROWS if x[1] == phase)
    starts = [dates[c][0] for c in codes]
    ends = [dates[c][1] for c in codes]
    bg = "BFBFBF"
    put(ws, r, 1, "", bg=bg, bold=True)
    put(ws, r, 2, "", bg=bg, bold=True)
    put(ws, r, 3, label, bg=bg, bold=True)
    put_date(ws, r, 4, min(starts), bg=bg, bold=True)
    put_date(ws, r, 5, max(ends), bg=bg, bold=True)
    put(ws, r, 6, "", bg=bg)
    put(ws, r, 7, "", bg=bg)
    put(ws, r, 8, pd, bg=bg, bold=True)
    put(ws, r, 9, "", bg=bg)
    ws.row_dimensions[r].height = 18
    return r + 1


if __name__ == "__main__":
    path, dates, wdays = build()
    print("wrote", path)
    print("working days", len(wdays), wdays[0], wdays[-1])
    for code, phase, task, pd, owner, deliverable, path_name in ROWS:
        s, e, n = dates[code]
        print(f"{code:5} {pd:2}pd {n:2}wd {s} ~ {e}  {task}")
