#!/usr/bin/env python3
"""Build FAN_F CAN development plan (OEM schedule format, dates not hours)."""

from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

PLAN_START = date(2026, 6, 15)
PLAN_END = date(2026, 10, 30)
AS_OF = date(2026, 9, 14)

# Task list: phase, task, deliverable, start, end, role, note
# Window is calendar dates 2026-06-15 .. 2026-10-30 (Mon .. Fri).
# Original 工时 sheet was not uploaded; durations follow a typical
# HV-fan CAN WBS on S32K144W / 500 kbit/s and stay inside that window.
TASKS = [
    (
        "需求",
        "接收并冻结整车 CAN 矩阵 / DBC",
        "已确认矩阵/DBC 版本 + 差异清单",
        date(2026, 6, 15),
        date(2026, 6, 26),
        "项目 / 软件",
        "无正式矩阵则先用临时信号表，冻结日写入备注",
    ),
    (
        "需求",
        "确认周期报文、事件报文、NM、诊断范围",
        "CAN 需求说明书（报文清单 + 周期 + 诊断范围）",
        date(2026, 6, 15),
        date(2026, 7, 3),
        "软件",
        "与福田接口人书面确认；无 NM/UDS 则在备注写“不适用”",
    ),
    (
        "需求",
        "CAN 软件方案（调度、故障映射、与 FOC 接口）",
        "软件方案 1 页 + 信号-应用对照表",
        date(2026, 6, 22),
        date(2026, 7, 10),
        "软件",
        "含 0x7AA 等本项目 ID 的最终分配（以矩阵为准）",
    ),
    (
        "底层",
        "FlexCAN 驱动：波特率 500 kbit/s、滤波、邮箱、总线off恢复",
        "驱动自测记录（发送/接收/BusOff）",
        date(2026, 6, 29),
        date(2026, 7, 17),
        "软件",
        "时钟改 80 MHz 后复核 CAN 时序",
    ),
    (
        "底层",
        "周期发送与事件发送调度",
        "周期误差记录（对照矩阵）",
        date(2026, 7, 13),
        date(2026, 7, 31),
        "软件",
        "与 1 ms 任务调度对齐，禁止阻塞 delay",
    ),
    (
        "底层",
        "接收解析与应用层接口",
        "RX API + 超时/默认值策略",
        date(2026, 7, 13),
        date(2026, 7, 31),
        "软件",
        "控制报文丢失时的风扇安全态写进备注",
    ),
    (
        "协议",
        "网络管理（矩阵要求时）",
        "NM 睡眠/唤醒测试记录，或“不适用”说明",
        date(2026, 7, 20),
        date(2026, 8, 14),
        "软件",
        "矩阵无 NM 则本行关闭并改状态为取消",
    ),
    (
        "协议",
        "诊断 UDS / ISO-TP（矩阵或售后要求时）",
        "DID/DTC 清单 + 台架诊断记录，或“不适用”说明",
        date(2026, 7, 27),
        date(2026, 8, 21),
        "软件",
        "与 CIA 刷写/解锁要求分开登记",
    ),
    (
        "应用",
        "状态上送：转速、温度、电压、故障字",
        "对照矩阵的信号精度/分辨率记录",
        date(2026, 8, 3),
        date(2026, 8, 21),
        "软件",
        "故障字与 FOC/保护模块同源",
    ),
    (
        "应用",
        "控制报文：使能、目标、运行模式",
        "控制路径台架记录（含非法值/超时）",
        date(2026, 8, 3),
        date(2026, 8, 21),
        "软件",
        "与高压互锁/故障降级策略一致",
    ),
    (
        "应用",
        "故障与 DTC / 故障字映射",
        "故障映射表",
        date(2026, 8, 17),
        date(2026, 9, 4),
        "软件 / 测试",
        "只映射本控制器真实故障，不编造码",
    ),
    (
        "应用",
        "调试口与整车 CAN 隔离（FreeMASTER / JTAG）",
        "量产态：调试口关闭或隔离的检查记录",
        date(2026, 8, 24),
        date(2026, 9, 11),
        "软件",
        "交样连续生产态见 FAN_F-CS-200，本行只覆盖 CAN 调试残留",
    ),
    (
        "台架",
        "PCAN 台架联调（500 kbit/s）",
        "联调记录（通过准则：矩阵内周期报文稳定收发）",
        date(2026, 8, 31),
        date(2026, 9, 18),
        "测试 / 软件",
        "与 FreeMASTER 通道分开，避免占同一 ID",
    ),
    (
        "台架",
        "周期、负载、丢帧、BusOff 恢复",
        "总线负载与异常恢复记录",
        date(2026, 9, 7),
        date(2026, 9, 24),
        "测试",
        "9-25 中秋，结束日收到 24 日（周四）",
    ),
    (
        "台架",
        "诊断与刷写相关验证（如有）",
        "诊断/刷写检查表，或“不适用”",
        date(2026, 9, 14),
        date(2026, 10, 16),
        "测试 / 软件",
        "10/1–10/7 国庆不排关闭节点；结束放在节后",
    ),
    (
        "交样",
        "交样软件冻结与配套材料",
        "冻结版本号 + 校验和 + 发布说明",
        date(2026, 9, 21),
        date(2026, 10, 16),
        "项目 / 软件",
        "对齐福田交样日（黄格节点）；本行结束不得晚于交样",
    ),
    (
        "实车",
        "实车问题关闭",
        "问题清单关闭状态",
        date(2026, 10, 8),
        date(2026, 10, 23),
        "软件 / 测试",
        "国庆后进车；影响交样的项升级为里程碑风险",
    ),
    (
        "发布",
        "CAN 开发关闭：发布包、测试报告、遗留项",
        "发布包 + 测试报告 + 遗留项（无则写无）",
        date(2026, 10, 19),
        date(2026, 10, 30),
        "项目 / 软件",
        "窗口结束日 10-30（周五）",
    ),
]

MILESTONES = [
    ("M1 矩阵/需求冻结", date(2026, 7, 3), "需求阶段关闭", "未开始"),
    ("M2 驱动与调度可用", date(2026, 7, 31), "底层阶段关闭", "未开始"),
    ("M3 应用报文闭环", date(2026, 9, 4), "应用阶段主路径可用", "未开始"),
    ("M4 台架通讯通过", date(2026, 9, 18), "可支持上车/交样联调", "未开始"),
    ("M5 交样软件冻结", date(2026, 10, 16), "对齐福田交样（请填客户节点）", "未开始"),
    ("M6 CAN 开发关闭", date(2026, 10, 30), "本窗口结束", "未开始"),
]

PHASE_FILL = {
    "需求": "D6EAF8",
    "底层": "D5F5E3",
    "协议": "FCF3CF",
    "应用": "FADBD8",
    "台架": "E8DAEF",
    "交样": "D6DBDF",
    "实车": "F5CBA7",
    "发布": "AED6F1",
}


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def font(size=11, bold=False, color="000000", name="Calibri"):
    return Font(name=name, size=size, bold=bold, color=color)


def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def style_range(ws, row, c1, c2, **kwargs):
    for col in range(c1, c2 + 1):
        cell = ws.cell(row, col)
        if "fill" in kwargs:
            cell.fill = kwargs["fill"]
        if "font" in kwargs:
            cell.font = kwargs["font"]
        if "alignment" in kwargs:
            cell.alignment = kwargs["alignment"]
        if "border" in kwargs:
            cell.border = kwargs["border"]


def merge_title(ws, row, value, fill_color, font_obj, height=22):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    cell = ws.cell(row, 1, value)
    cell.fill = fill(fill_color)
    cell.font = font_obj
    cell.alignment = align("left", "center")
    style_range(ws, row, 1, 10, fill=fill(fill_color), font=font_obj, alignment=align("left", "center"))
    ws.row_dimensions[row].height = height
    return cell


def put(ws, row, col, value, *, bg=None, bold=False, color="000000", h="left", size=10):
    cell = ws.cell(row, col, value)
    if bg:
        cell.fill = fill(bg)
    cell.font = font(size=size, bold=bold, color=color)
    cell.alignment = align(h)
    cell.border = THIN
    return cell


def put_date(ws, row, col, d, *, bg=None):
    cell = put(ws, row, col, d, bg=bg, h="center")
    cell.number_format = "YYYY-MM-DD"
    return cell


def build_plan_sheet(wb: Workbook):
    ws = wb.active
    ws.title = "CAN开发计划"

    widths = {
        "A": 6,
        "B": 10,
        "C": 38,
        "D": 36,
        "E": 13,
        "F": 13,
        "G": 14,
        "H": 12,
        "I": 10,
        "J": 28,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    merge_title(
        ws,
        1,
        "FAN_F 高压风扇控制器  ·  CAN 开发计划",
        "1F4E79",
        font(16, True, "FFFFFF"),
        26,
    )
    merge_title(
        ws,
        2,
        "给车厂看日期和节点，不看内部工时    窗口：2026-06-15 ～ 2026-10-30    密级：内部    只填黄色格    状态请按实际改",
        "2E75B6",
        font(10, False, "FFFFFF"),
        18,
    )

    merge_title(ws, 3, "一、文件与项目", "1F4E79", font(11, True, "FFFFFF"), 18)

    labels = [
        (4, "供应商", None, "产品 / 零件号", "FAN_F / "),
        (5, "文件编号", "FAN_F-SW-100", "版本 / 日期", "V1.0 / 2026-09-14"),
        (6, "计划开始", PLAN_START, "计划结束", PLAN_END),
        (7, "编制", None, "审核", None),
        (8, "批准", None, "年份说明", "原表未写年，按 2026（今天 2026-09-14）。若是 2025 把本表日期整年改掉"),
    ]
    yellow_cells = {
        (4, 2),
        (4, 5),
        (5, 5),
        (7, 2),
        (7, 5),
        (8, 2),
    }
    for row, a, b, c, d in labels:
        put(ws, row, 1, a, bg="D6DCE4", bold=True)
        put(ws, row, 2, b, bg="FFF2CC" if (row, 2) in yellow_cells else "FFFFFF")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row, 3).border = THIN
        if (row, 2) in yellow_cells:
            ws.cell(row, 3).fill = fill("FFF2CC")
        put(ws, row, 4, c, bg="D6DCE4", bold=True)
        put(ws, row, 5, d, bg="FFF2CC" if (row, 5) in yellow_cells else "FFFFFF")
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=10)
        for col in range(6, 11):
            ws.cell(row, col).border = THIN
            ws.cell(row, col).fill = fill("FFF2CC" if (row, 5) in yellow_cells else "FFFFFF")
        if row == 6:
            ws.cell(row, 2).number_format = "YYYY-MM-DD"
            ws.cell(row, 5).number_format = "YYYY-MM-DD"
        ws.row_dimensions[row].height = 18

    # customer gates
    put(ws, 9, 1, "福田节点", bg="D6DCE4", bold=True)
    put(ws, 9, 2, "A样 / 交样日（填客户日期）", bg="FFF2CC")
    ws.merge_cells("B9:C9")
    ws.cell(9, 3).border = THIN
    ws.cell(9, 3).fill = fill("FFF2CC")
    put(ws, 9, 4, "B样 / OTS / 其他", bg="D6DCE4", bold=True)
    put(ws, 9, 5, None, bg="FFF2CC")
    ws.merge_cells("E9:J9")
    for col in range(6, 11):
        ws.cell(9, col).border = THIN
        ws.cell(9, col).fill = fill("FFF2CC")
    ws.row_dimensions[9].height = 18

    merge_title(
        ws,
        10,
        "格式要点：①工时改成开始/结束日期，人天不要发给主机厂  ②负责人要写，写成「责任岗位 + 姓名」，交给福田再补电话/邮箱  ③每条任务必须有交付物和状态  ④子计划必须能对上整车交样日  ⑤编制/审核/批准要签  详见「格式说明」表",
        "FFF2CC",
        font(9, False, "333333"),
        28,
    )

    merge_title(ws, 11, "二、里程碑（对内关门日，须能托住上面的福田节点）", "1F4E79", font(11, True, "FFFFFF"), 18)

    for col in range(1, 11):
        put(ws, 12, col, "", bg="D6DCE4", bold=True, h="center")
    put(ws, 12, 1, "序号", bg="D6DCE4", bold=True, h="center")
    put(ws, 12, 2, "里程碑", bg="D6DCE4", bold=True, h="center")
    ws.merge_cells("B12:C12")
    ws.cell(12, 3).fill = fill("D6DCE4")
    ws.cell(12, 3).border = THIN
    put(ws, 12, 4, "目标日期", bg="D6DCE4", bold=True, h="center")
    put(ws, 12, 5, "状态", bg="D6DCE4", bold=True, h="center")
    put(ws, 12, 6, "含义", bg="D6DCE4", bold=True, h="center")
    ws.merge_cells("F12:J12")
    for col in range(7, 11):
        ws.cell(12, col).fill = fill("D6DCE4")
        ws.cell(12, col).border = THIN
    ws.row_dimensions[12].height = 18

    for i, (name, d, meaning, status) in enumerate(MILESTONES, 1):
        r = 12 + i
        put(ws, r, 1, i, h="center")
        put(ws, r, 2, name, bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(r, 3).border = THIN
        put_date(ws, r, 4, d)
        put(ws, r, 5, status, bg="FFF2CC", h="center")
        put(ws, r, 6, meaning)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col in range(7, 11):
            ws.cell(r, col).border = THIN
        ws.row_dimensions[r].height = 18

    task_header_row = 19
    merge_title(ws, task_header_row, "三、CAN 工作包（开始/结束日期已按窗口排完，不再列工时）", "1F4E79", font(11, True, "FFFFFF"), 18)

    headers = [
        "序号",
        "阶段",
        "任务",
        "交付物",
        "开始日期",
        "结束日期",
        "责任岗位",
        "责任人",
        "状态",
        "备注",
    ]
    hr = task_header_row + 1
    for col, h in enumerate(headers, 1):
        put(ws, hr, col, h, bg="D6DCE4", bold=True, h="center")
    ws.row_dimensions[hr].height = 20
    ws.auto_filter.ref = f"A{hr}:J{hr + len(TASKS)}"
    ws.freeze_panes = f"A{hr + 1}"

    first_task = hr + 1
    for i, (phase, task, deliverable, start, end, role, note) in enumerate(TASKS, 1):
        r = hr + i
        bg = PHASE_FILL[phase]
        put(ws, r, 1, i, bg=bg, h="center")
        put(ws, r, 2, phase, bg=bg, bold=True, h="center")
        put(ws, r, 3, task, bg=bg)
        put(ws, r, 4, deliverable, bg=bg)
        put_date(ws, r, 5, start, bg=bg)
        put_date(ws, r, 6, end, bg=bg)
        put(ws, r, 7, role, bg=bg, h="center")
        put(ws, r, 8, None, bg="FFF2CC", h="center")
        put(ws, r, 9, "未开始", bg="FFF2CC", h="center")
        put(ws, r, 10, note, bg=bg)
        ws.row_dimensions[r].height = 32
        if end < start:
            raise ValueError(f"task {i} end before start")
        if start < PLAN_START or end > PLAN_END:
            raise ValueError(f"task {i} outside window: {start} {end}")

    last_task = hr + len(TASKS)

    dv = DataValidation(
        type="list",
        formula1='"未开始,进行中,已完成,暂停,取消"',
        allow_blank=True,
    )
    dv.error = "请选：未开始 / 进行中 / 已完成 / 暂停 / 取消"
    dv.errorTitle = "状态"
    dv.prompt = "选择状态"
    dv.promptTitle = "状态"
    ws.add_data_validation(dv)
    dv.add(f"I{first_task}:I{last_task}")
    dv.add("E13:E18")

    # contacts
    contact_row = last_task + 1
    merge_title(ws, contact_row, "四、接口人（交给福田的进度表必须能找到人；对内可先写岗位）", "1F4E79", font(11, True, "FFFFFF"), 18)
    cr = contact_row + 1
    for col, h in enumerate(["角色", "姓名", "电话", "邮箱", "备注", "", "", "", "", ""], 1):
        if col <= 5:
            put(ws, cr, col, h, bg="D6DCE4", bold=True, h="center")
        else:
            put(ws, cr, col, "", bg="D6DCE4")
    ws.merge_cells(start_row=cr, start_column=5, end_row=cr, end_column=10)
    for col in range(6, 11):
        ws.cell(cr, col).fill = fill("D6DCE4")
        ws.cell(cr, col).border = THIN

    contacts = [
        ("项目经理", "对福田的进度和节点"),
        ("CAN / 应用软件", "技术接口"),
        ("测试", "台架/实车问题"),
        ("质量", "交样与文件"),
    ]
    for i, (role, note) in enumerate(contacts):
        r = cr + 1 + i
        put(ws, r, 1, role, bg="FFFFFF", bold=True)
        put(ws, r, 2, None, bg="FFF2CC")
        put(ws, r, 3, None, bg="FFF2CC")
        put(ws, r, 4, None, bg="FFF2CC")
        put(ws, r, 5, note)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=10)
        for col in range(6, 11):
            ws.cell(r, col).border = THIN
        ws.row_dimensions[r].height = 18

    last_contact = cr + len(contacts)

    # change log
    ch = last_contact + 1
    merge_title(ws, ch, "五、变更记录", "1F4E79", font(11, True, "FFFFFF"), 18)
    put(ws, ch + 1, 1, "版本", bg="D6DCE4", bold=True, h="center")
    put(ws, ch + 1, 2, "日期", bg="D6DCE4", bold=True, h="center")
    put(ws, ch + 1, 3, "变更说明", bg="D6DCE4", bold=True)
    ws.merge_cells(start_row=ch + 1, start_column=3, end_row=ch + 1, end_column=10)
    for col in range(4, 11):
        ws.cell(ch + 1, col).fill = fill("D6DCE4")
        ws.cell(ch + 1, col).border = THIN
    put(ws, ch + 2, 1, "V1.0", h="center")
    put_date(ws, ch + 2, 2, AS_OF)
    put(
        ws,
        ch + 2,
        3,
        "按主机厂进度表格式重建：工时改为起止日期；窗口 2026-06-15～2026-10-30；补交付物/岗位/责任人/状态/签署/接口人。原《高压风扇开发计划.xlsx》未随消息上传，工作包按高压风扇 CAN 常规分解，收到原表工时后可按人天比例重排。",
    )
    ws.merge_cells(start_row=ch + 2, start_column=3, end_row=ch + 2, end_column=10)
    for col in range(4, 11):
        ws.cell(ch + 2, col).border = THIN
    ws.row_dimensions[ch + 2].height = 36

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    ws.print_title_rows = "1:2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "1F4E79"

    return first_task, last_task


def build_notes_sheet(wb: Workbook):
    ws = wb.create_sheet("格式说明")
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 88
    for col in range(4, 11):
        ws.column_dimensions[get_column_letter(col)].width = 3

    ws.merge_cells("A1:C1")
    ws["A1"] = "这份 CAN 计划相对原「工时表」要改什么（按福田/主机厂常见要求，不是你们公司内部工时算法）"
    ws["A1"].fill = fill("1F4E79")
    ws["A1"].font = font(14, True, "FFFFFF")
    ws["A1"].alignment = align("left", "center")
    style_range(ws, 1, 1, 3, fill=fill("1F4E79"), font=font(14, True, "FFFFFF"), alignment=align("left", "center"))
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:C2")
    ws["A2"] = "结论先看第 1 条：负责人列要保留，但不要只写一个名字。工时列不要给车厂。"
    ws["A2"].fill = fill("2E75B6")
    ws["A2"].font = font(10, False, "FFFFFF")
    style_range(ws, 2, 1, 3, fill=fill("2E75B6"), font=font(10, False, "FFFFFF"), alignment=align("left", "center"))
    ws.row_dimensions[2].height = 18

    headers = ["序号", "项目", "怎么处理"]
    for col, h in enumerate(headers, 1):
        put(ws, 3, col, h, bg="D6DCE4", bold=True, h="center")
    ws.row_dimensions[3].height = 18

    rows = [
        (
            "负责人",
            "要写。福田 SOR / 供应商进度表都要能找到责任人和联系方式；IATF+APQP 也要求核心小组角色清楚。"
            "正确写法是「责任岗位 + 姓名」，交给福田的版本再加电话、邮箱（见本表第四节）。"
            "不要做成 10 列 RASIC——那是 APQP 总计划或 CIA 接口协议用的。"
            "对内周例会表可以只写岗位。芯片原厂（NXP）不必写进负责人，除非他们真的承接本行工作。",
        ),
        (
            "工时 → 日期",
            "主机厂要的是日历关门日，用来对整车 A/B/C/OTS/交样，不是你们内部人天。"
            "人天给出去会被拿去压价或压缩周期。本表已改成开始日期、结束日期。"
            "工期可用日期相减在内部另算，不要作为对客列。"
            "原表未上传，本表按 2026-06-15～2026-10-30 窗口和高压风扇 CAN 常规工作包排期；"
            "你把原表发来后，可按「人天占比 × 窗口工作日」重算每行起止，并保留搭接。",
        ),
        (
            "交付物",
            "车厂进度表每一行都要能验收。空任务名 + 一个日期过不了评审。"
            "本表每行都有交付物；没有产出的行不要放进来。",
        ),
        (
            "对齐整车节点",
            "福田/主机厂评审看的是你的关门日能不能托住交样日，不是看你忙不忙。"
            "第一节黄格填客户 A样/B样/OTS/交样日；第五节里程碑必须早于或等于对应交样。"
            "交样软件冻结不得晚于客户交样日。",
        ),
        (
            "文件头与签署",
            "按福田文件习惯补：编号、版本、密级、编制/审核/批准、变更记录。"
            "只有任务列表、没有签署，质量/项目不认这是受控文件。",
        ),
        (
            "状态与变更",
            "每次对客同步改状态和版本。节点移动要写变更记录，并通知福田项目接口，不要只改日期不留痕。",
        ),
        (
            "和 CIA 的关系",
            "本表是 CAN 功能开发子计划，不是 CIA 3-1-1 网络安全开发计划。"
            "调试口关闭、量产密钥、漏洞响应仍走 CIA 那套表（CS-200/CS-300/CS-310）。"
            "两套节点不要互相打架：交连续生产态样件的日期，应落在本表 M5 附近。",
        ),
        (
            "中英文",
            "CIA 8-3 要求正式网络安全交付双语。这份内部/项目进度表用中文即可。"
            "只有当福田明确要英文进度表时再补英文明细，不要和 CIA 手册绑在一起翻译。",
        ),
        (
            "不要写进对客表",
            "内部人天、成本、加班、未对客户承诺的缓冲、个人评价、与芯片原厂的商务条款。"
            "FreeMASTER 工程调试细节可对内保留，对客只写「调试口已隔离/关闭」。",
        ),
        (
            "年份",
            "你只写了 6 月 15 日～10 月 30 日。按今天 2026-09-14，本表用 2026。"
            "若实际窗口是 2025，整表日期改年即可，工作包顺序不用动。",
        ),
    ]

    for i, (item, how) in enumerate(rows, 1):
        r = 3 + i
        put(ws, r, 1, i, bg="FFFFFF", bold=True, h="center")
        put(ws, r, 2, item, bg="FFF2CC", bold=True)
        put(ws, r, 3, how, bg="FFFFFF")
        ws.row_dimensions[r].height = 68

    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "C65911"


def main():
    wb = Workbook()
    build_plan_sheet(wb)
    build_notes_sheet(wb)
    out = "/workspace/docs/CAN/FAN_F-SW-100_高压风扇CAN开发计划.xlsx"
    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main()
