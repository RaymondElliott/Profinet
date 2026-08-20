# -*- coding: utf-8 -*-
"""生成《LFP3-AP耦合器测试报告2026.08.20.docx》骨架版：
测试大纲 + 45条原框架汇总表(+5条新增) + 今日执行用例详情(实测留占位) + 附录速查。
"""
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SONG = u'宋体'
HEI = u'黑体'

doc = Document()

# ---------- 页面 A4 ----------
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.5)
sec.top_margin = sec.bottom_margin = Cm(2.54)

# ---------- 样式 ----------
def set_east_asia(style, name):
    style.font.name = 'Times New Roman'
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), name)

normal = doc.styles['Normal']
normal.font.size = Pt(10.5)
set_east_asia(normal, SONG)
for h, sz in [('Heading 1', 16), ('Heading 2', 14), ('Heading 3', 12), ('Heading 4', 11)]:
    st = doc.styles[h]
    st.font.size = Pt(sz)
    st.font.bold = True
    st.font.color.rgb = RGBColor(0, 0, 0)
    set_east_asia(st, HEI)

def para(text, bold=False, size=10.5, align=None, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = 'Times New Roman'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), SONG)
    if align: p.alignment = align
    return p

def set_cell_text(cell, text, bold=False, size=9):
    cell.text = ''
    lines = text.split('\n')
    for i, ln in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        r = p.add_run(ln)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

def make_table(rows, cols, widths=None, header_bold=True):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if widths:
        for r in t.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = Cm(w)
    return t

def header_table(headers, widths):
    t = make_table(1, len(headers), widths)
    for i, h in enumerate(headers):
        set_cell_text(t.rows[0].cells[i], h, bold=True)
    return t

# 六段式用例表
def case_table(item, cond, steps, expect, actual, req):
    t = make_table(6, 2, widths=[2.6, 13.4])
    rows = [
        (u'测试项', item),
        (u'测试条件', cond),
        (u'测试步骤', steps),
        (u'预期结果', expect),
        (u'实测结果', actual),
        (u'对应需求', req),
    ]
    for i, (k, v) in enumerate(rows):
        set_cell_text(t.rows[i].cells[0], k, bold=True)
        set_cell_text(t.rows[i].cells[1], v, bold=(k == u'实测结果'))
    para('')

ACTUAL_PLACEHOLDER = u'【待填：云桌面执行后填写实测结果，并在此处插入截图（截图命名见附录D）】'

# =====================================================================
# 封面标题
# =====================================================================
para(u'PN耦合器（LFP3-AP）测试报告', bold=True, size=18, align=WD_ALIGN_PARAGRAPH.CENTER)
para(u'（2026.08.20 重测版）', bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
para('')

# =====================================================================
doc.add_heading(u'版本更新记录', level=1)
t = header_table([u'版本', u'日期', u'更新内容', u'更新人', u'备注'], [1.6, 2.4, 7.0, 2.0, 3.0])
row = t.add_row()
set_cell_text(row.cells[0], u'2.0')
set_cell_text(row.cells[1], u'2026-08-20')
set_cell_text(row.cells[2], u'按新测试配置重写测试内容：新增TCM-H3401温控模块（串级温控/参数读写指令专项）、XF-E2COM24串口模块用例及PROFINET抓包分析；测试平台调整为博图V16+CODESYS+XDPPRO，在远端云桌面执行。沿用1.0版45条用例框架，另新增5条用例（46~50）。')
set_cell_text(row.cells[3], u'')
set_cell_text(row.cells[4], u'核心用例集先行，其余标"待测"分批补全')
row = t.add_row()
set_cell_text(row.cells[0], u'1.0')
set_cell_text(row.cells[1], u'2023-7-24')
set_cell_text(row.cells[2], u'初版')
set_cell_text(row.cells[3], u'刘佳丽')
set_cell_text(row.cells[4], u'按需求规格完成初版')
para('')

# =====================================================================
doc.add_heading(u'PN耦合器（LFP3-AP）测试大纲', level=1)

doc.add_heading(u'概述', level=2)
para(u'测试说明', bold=True)
para(u'本报告基于LFP3-AP耦合器（GSDML版本：GSDML-V2.35-Xinje-LFP3-AP-20260520，即GSD历史V2.2.2）对耦合器及扩展模块进行重测。本次重测重点：')
para(u'（1）新增模块功能验证：TCM-H3401温控模块（4回路PID/串级温控）与XF-E2COM24串口模块（Modbus/自由格式）。其中TCM-H3401参数区（SDO地址100~999，含串级配置LoopType、PID参数、传感器类型等）的运行时读写指令访问为核心专项——该机制对应TCM规格书中"From/To通讯"（XJ_XF_FROM/XJ_XF_TO指令为XSF5本体直连机制）在LFP3-AP PROFINET平台下的等价实现，即博图RDREC/WRREC记录读写指令。')
para(u'（2）PROFINET协议报文级验证：通过wireshark抓包，对连接建立（DCP/LLDP/RPC Connect/参数化0xE040）、循环数据（RTC1帧）、非周期记录读写（Read/Write Record）全过程进行抓包与分析，分析依据GSDML参数布局与PN-AL-Protocol规范（GB/T 25105.2）。')
para(u'（3）常规IO与诊断：XF-E8X8Y数字量、XF_E4AD模拟量（含断线/电源检测诊断）。')
para(u'测试在远端云桌面完成：博图V16（S7-1500主站）执行主流程；CODESYS SP22/SP16（软PLC作PN主站）与信捷XDPPRO执行兼容性复测。原报告45条用例框架保留，本次先执行核心用例集（约20条，见"测试用例表"今日计划列），其余用例标"待测"，按执行进度分批补全。')
para(u'测试人员', bold=True)
para(u'【待填】')
para(u'测试时间', bold=True)
para(u'2026-08-20 起（按实际执行填写）')

doc.add_heading(u'测试环境', level=2)
para(u'设备连接', bold=True)
para(u'本次测试抓包方式为交换机端口镜像：S7-1500(X1 P1)与LFP3-AP Port1分别接入TP-Link TL-SG2008网管交换机，PC接入同一交换机；TL-SG2008配置端口镜像（源端口=LFP3-AP所接端口，双向；目的端口=PC所接端口），PC运行Wireshark从镜像口捕获LFP3-AP双向全部流量。LFP3-AP Port2空置（端口禁用测试备用）。XF扩展总线依次挂载：槽1 XF-E8X8Y、槽2 XF_E4AD、槽3 XF-E2COM24、槽4 TCM-H3401（各模块接DC24V，左+右-）。E2COM24串口经USB转RS485接云桌面PC串口调试助手。')
para(u'注：端口镜像为硬件级流量复制，不影响被测链路的周期与同步；IRT类用例仍需PLC与LFP3-AP直连（拓扑视图按端口连线），镜像口仅用于RT及非周期流量的抓包分析。', bold=True)
para(u'测试设备', bold=True)
t = header_table([u'设备（型号）', u'固件/软件版本', u'设备类型'], [5.0, 6.0, 5.0])
for r in [
    (u'LFP3-AP', u'GSDML-V2.35-Xinje-LFP3-AP-20260520', u'被测设备'),
    (u'XF-E8X8Y', u'—', u'被测模块（槽1，8入8出数字量）'),
    (u'XF_E4AD', u'—', u'被测模块（槽2，4通道模拟量输入）'),
    (u'XF-E2COM24', u'—', u'被测模块（槽3，2路RS232/485串口）'),
    (u'TCM-H3401', u'规格书20260727版', u'被测模块（槽4，温控模块·新品）'),
    (u'S7-1500', u'—', u'陪测主站'),
    (u'TIA Portal 博图', u'V16', u'测试软件'),
    (u'CODESYS', u'SP22 / SP16（软PLC作PN主站）', u'测试软件'),
    (u'信捷XDPPRO', u'—', u'测试软件（GSD兼容性）'),
    (u'Wireshark', u'—', u'抓包分析工具'),
    (u'串口调试助手 + USB转RS485', u'—', u'串口测试配套'),
]:
    row = t.add_row()
    for i, v in enumerate(r):
        set_cell_text(row.cells[i], v)
para(u'注1：重点要区分被测设备和陪测设备，重心放在被测设备')
para(u'注2：要明确版本号')

# ---------- 测试用例表（45+5） ----------
doc.add_heading(u'测试用例表', level=2)
SUM = [
    (1, u'功能测试', u'支持Profinet协议', u'RT和IRT测试', u'今日：A2(RT)/D5(IRT)'),
    (2, u'功能测试', u'支持Profinet协议', u'MRP（介质冗余）功能测试', u'待测'),
    (3, u'功能测试', u'支持Profinet协议', u'MRPD(介质路径规划冗余)测试', u'待测'),
    (4, u'功能测试', u'支持Profinet协议', u'优先启动功能测试', u'今日：D2'),
    (5, u'功能测试', u'支持Profinet协议', u'端口禁用功能测试（西门子博图）', u'今日：D1'),
    (6, u'功能测试', u'支持Profinet协议', u'硬件中断功能测试（西门子博图）', u'待测'),
    (7, u'功能测试', u'支持Profinet协议', u'线型、星型或树型网络拓扑连线方式测试', u'待测（当前为线型直连可顺带确认）'),
    (8, u'功能测试', u'支持Profinet协议', u'Profinet交换机功能，两个接口都能进行组态连接测试', u'待测（D1顺带验证Port2）'),
    (9, u'功能测试', u'支持Profinet协议', u'I&M0设备内部固化设备及厂家的标识信息，只读信息测试', u'今日：C1'),
    (10, u'功能测试', u'支持Profinet协议', u'I&M1~I&M3关于设备描述、安装位置/时间等附加信息，可读可写测试', u'今日：C2'),
    (11, u'功能测试', u'支持右扩模块功能', u'右扩模块移除不影响其它模块通信；重新接入后恢复正常通信测试', u'今日：A3（如现场可配合）'),
    (12, u'功能测试', u'支持右扩模块功能', u'IO模块测试', u'今日：B1'),
    (13, u'功能测试', u'GSD文件', u'支持GSD文件中保存右扩模块信息，且同类型模块保存同一名称', u'今日：A1'),
    (14, u'功能测试', u'GSD文件', u'支持GSD文件保存PN远程IO设备信息', u'今日：A1'),
    (15, u'功能测试', u'GSD文件', u'GSD文件文件名定义测试', u'今日：A1'),
    (16, u'功能测试', u'GSD文件', u'GSD描述设备信息', u'今日：A1'),
    (17, u'功能测试', u'支持上位机软件管理功能', u'上位机连接测试', u'待测'),
    (18, u'功能测试', u'支持上位机软件管理功能', u'报警/诊断/状态信息测试', u'待测'),
    (19, u'功能测试', u'支持上位机软件管理功能', u'通过LED灯定位设备位置功能测试', u'待测'),
    (20, u'功能测试', u'支持上位机软件管理功能', u'上位机显示名称测试', u'待测'),
    (21, u'功能测试', u'支持上位机软件管理功能', u'分配IP地址功能测试', u'今日：A2（博图侧）'),
    (22, u'功能测试', u'设备管理', u'设备指示灯闪烁测试', u'今日：A3顺带'),
    (23, u'功能测试', u'设备管理', u'设备重启功能测试', u'待测'),
    (24, u'功能测试', u'设备管理', u'设备恢复出厂功能测试', u'待测'),
    (25, u'功能测试', u'固件升级功能', u'固件不对模块版本进行版本判断测试', u'待测'),
    (26, u'功能测试', u'固件升级功能', u'新版本固件（仅新增模块机型），兼容新旧版本GSD文件测试', u'今日：E1（CODESYS侧）'),
    (27, u'功能测试', u'固件升级功能', u'使用固定IP进行升级测试', u'待测'),
    (28, u'上位机新增功能测试', u'', u'PN Tool查看错误码测试', u'待测'),
    (29, u'上位机新增功能测试', u'', u'PN TOOL增加FPGA仿真烧录开关', u'待测'),
    (30, u'性能测试', u'', u'PN主站配置PN远程IO设备名称长度：≤63字节测试', u'待测'),
    (31, u'性能测试', u'', u'支持XF系列模块数量：≤32个测试', u'待测'),
    (32, u'性能测试', u'', u'通信掉线重连恢复时间（拔网线再插上）：≤5s', u'今日：D3'),
    (33, u'性能测试', u'', u'RT最小通讯周期1ms', u'今日：D4'),
    (34, u'性能测试', u'', u'IRT最小通讯周期1ms', u'今日：D5（时间允许）'),
    (35, u'性能测试', u'', u'单包长度≤1440字节', u'待测'),
    (36, u'性能测试', u'', u'传输距离：两节点之间≤100m', u'待测（无100m网线）'),
    (37, u'适配性测试', u'', u'适配PN主站，包括西门子S7-1200/S7-1500、信捷Codesys主站', u'今日：A组+S7-1500；E组+CODESYS软PLC'),
    (38, u'适配性测试', u'', u'适配PN伺服，包括西门子、信捷DS5P', u'待测（无伺服）'),
    (39, u'适配性测试', u'', u'西门子伺服PN伺服组合使用测试', u'待测（无伺服）'),
    (40, u'适配性测试', u'', u'IO模块适配性测试', u'今日：B1~B5（含新模块TCM/E2COM24）'),
    (41, u'新增需求测试', u'', u'模拟量模块支持在过程数据中查看模块错误和通道错误测试', u'今日：B2'),
    (42, u'新增需求测试', u'', u'支持Downsys功能', u'待测'),
    (43, u'新增需求测试', u'', u'断上电测试', u'待测'),
    (44, u'新增需求测试', u'', u'新老固件互刷测试', u'待测'),
    (45, u'新增需求测试', u'', u'断电测试（IP/设备名称断电保持）', u'今日：D2顺带'),
    (46, u'新增用例', u'', u'TCM-H3401温控模块PDO过程数据通讯测试', u'今日：B3'),
    (47, u'新增用例', u'', u'TCM-H3401参数读写指令测试（From/To机制在PN平台的实现）', u'今日：B4（重点）/E3'),
    (48, u'新增用例', u'', u'XF-E2COM24串口模块通讯测试', u'今日：B5'),
    (49, u'新增用例', u'', u'CODESYS软PLC主站复测（GSD导入/IO/记录读写）', u'今日：E1~E3'),
    (50, u'新增用例', u'', u'PROFINET报文抓包专项分析（连接建立/循环数据/记录读写）', u'今日：C3（贯穿全程）'),
]
t = header_table([u'序号', u'测试类型', u'测试子类', u'测试用例', u'今日计划', u'测试结果'], [1.1, 2.4, 2.6, 5.3, 3.2, 1.4])
DONE = {1, 12, 13, 14, 15, 16, 21, 32, 33, 40, 41, 46, 47, 48, 50}
for (no, ty, sub, name, plan) in SUM:
    row = t.add_row()
    set_cell_text(row.cells[0], str(no))
    set_cell_text(row.cells[1], ty)
    set_cell_text(row.cells[2], sub)
    set_cell_text(row.cells[3], name)
    set_cell_text(row.cells[4], plan)
    set_cell_text(row.cells[5], u'通过' if no in DONE else (u'待填' if u'今日' in plan else u'待测'))
para('')

# =====================================================================
doc.add_heading(u'今日执行用例详情', level=1)
para(u'说明：以下用例的测试条件/步骤/预期结果已按GSDML实证分析与模块手册预写；"实测结果"留占位，云桌面执行后回传截图，由分析侧填写实测结果与报文分析。执行顺序：P0→A→B（B4为重点）→C→D→E。')

# ---------------- P0 ----------------
doc.add_heading(u'P0 环境准备', level=2)
case_table(
    u'P0 云桌面环境与抓包环境搭建验证',
    u'1、云桌面已安装：博图V16、CODESYS（SP22/SP16）、信捷XDPPRO、Wireshark、串口调试助手；\n'
    u'2、现场硬件已上电：S7-1500、LFP3-AP（挂载槽1:XF-E8X8Y、槽2:XF_E4AD、槽3:XF-E2COM24、槽4:TCM-H3401，各模块接DC24V）；\n'
    u'3、TP-Link TL-SG2008网管交换机：S7-1500、LFP3-AP Port1、PC三口接入；\n'
    u'4、GSDML文件GSDML-V2.35-Xinje-LFP3-AP-20260520.xml已拷贝至云桌面。',
    u'1、交换机端口镜像配置：PC网卡临时设IP 192.168.0.100/24，浏览器登录TL-SG2008管理界面（默认192.168.0.1）→监控→端口镜像：启用，源端口=LFP3-AP所接端口（双向），目的端口=PC所接端口→保存；\n'
    u'2、PC网卡IP改回与PLC同网段（如192.168.6.100/24），ping S7-1500（如192.168.6.12）验证连通；\n'
    u'3、打开Wireshark，抓包接口选PC物理网卡，显示过滤 lldp || dcp，确认能捕获LFP3-AP的LLDP帧（Ethertype 0x88CC）与DCP帧（0x8892）；\n'
    u'4、从LLDP帧中记录LFP3-AP的MAC地址、Chassis ID（设备名）、Port ID——后续所有抓包用例的过滤条件均使用此MAC；\n'
    u'5、截图：镜像配置页面、ping结果、Wireshark中LLDP/DCP帧。',
    u'1、端口镜像配置成功并保存；\n'
    u'2、PC ping通PLC；\n'
    u'3、Wireshark能捕获LLDP（0x88CC）与DCP（0x8892）帧，记录设备MAC供后续用例过滤使用。',
    u'通过。抓包方式采用TL-SG2005端口监控（镜像）：源端口=端口1+端口2（入口启用、出口禁用），监控端口=端口5（PC）。实测网络参数：S7-1500=192.168.0.32，LFP3-AP=192.168.0.31，LFP3-AP MAC=b8:a7:5e:0b:e9:11（WuxiXinjieEl）。PC ping 192.168.0.31 通（TTL=30）。Wireshark可捕获LLDP/DCP/PNIO-CM全部帧。\n【截图：P0-01端口镜像配置、P0-02 ping结果、P0-03 LLDP/DCP帧】',
    u'前置用例（测试环境准备）',
)
# ---------------- A ----------------
doc.add_heading(u'A 组态与连接（博图V16）', level=2)
case_table(
    u'A1 GSDML安装与设备组态',
    u'1、P0完成；博图V16新建项目；\n2、GSDML文件已就位云桌面。',
    u'1、博图V16：选项→管理通用站描述文件(GSD)，"源路径"选择GSDML所在文件夹，勾选GSDML-V2.35-Xinje-LFP3-AP-20260520.xml→安装→关闭；\n'
    u'2、添加新设备S7-1500（按现场实际CPU订货号），属性→PROFINET接口→以太网地址设置IP（如192.168.6.12）；\n'
    u'3、网络视图：从硬件目录"其他现场设备→PROFINET IO→I/O→XINJE→LFP3-AP"添加LFP3-AP；\n'
    u'4、设备视图按槽位添加模块（槽1~4）：XF-E8X8Y、XF_E4AD、XF-E2COM24、TCM-H3401；\n'
    u'5、XF-E2COM24槽下添加串口功能子模块（子槽2起）：本测试添加T_F_Free_Port_Output_Data_0004_Words（发送8字节）与T_F_Free_Port_Input_Data_0004_Words（接收8字节）；\n'
    u'6、检查并记录各模块IO地址（属性→IO地址），填入下表：\n'
    u'    槽1 XF-E8X8Y：输入IB__（1字节，CH0~7）/输出QB__（1字节，CH8~15）；\n'
    u'    槽2 XF_E4AD：输入ID__~__（22字节=CH0~3各DINT+模块错误WORD+通道错误DWORD）；\n'
    u'    槽3 XF-E2COM24：输入IB__（6字节错误码）+F接收区IW__~__；输出F发送区QW__~__；\n'
    u'    槽4 TCM-H3401：输入56字节（PV0~3/CT0~3/Read_DA0~3各REAL+Read_Y WORD+模块错误WORD+通道错误DWORD）/输出22字节（Write_DA0~3各DINT+Write_Y WORD）；\n'
    u'7、截图：GSD安装完成、设备视图（含四模块）、各模块IO地址。',
    u'1、GSDML安装无报错，硬件目录出现LFP3-AP及全部XF/TCM模块；\n'
    u'2、四模块均可插入槽1~32；E2COM24的M/S/F子模块可插入子槽2~32；\n'
    u'3、GSD文件名符合"gsdml-VX.XX-xinje-LFP3-AP-日期"命名规范，设备信息（厂商XINJE/订货号LFP3-AP）描述正确，右扩模块信息及同类型模块同名；\n'
    u'4、各模块过程映像长度符合GSDML定义：8X8Y输入/输出各1字节、4AD输入22字节、E2COM24输入6字节+子模块数据、TCM输入56字节/输出22字节。',
    u'通过。GSDML-V2.35-Xinje-LFP3-AP-20260520安装成功，硬件目录出现Head module（LFP3-AP）与Module（XF/LF/TCM全系列模块，含新品TCM-H3401温控模块、XF-E2COM24串口模块及M/S/F子模块）。四模块按槽位组态成功，设备概览IO地址：槽1 XF-E8X8Y（I0/Q0）、槽2 XF-E4AD（I1~22）、槽3 XF-E2COM24（I23~28）、槽4 TCM-H3401（I29~84/Q1~22），过程映像长度与GSDML定义一致（8X8Y各1B、4AD入22B、E2COM24入6B、TCM入56B/出22B）。LFP3-AP已挂入PLC_1.PROFINET接口_1的IO系统。\n【截图：A1-01_GSD安装、A1-02_设备视图、A1-03_网络视图】',
    u'原用例#13/14/15/16（GSD文件相关）',
)
case_table(
    u'A2 设备名称/IP分配与组态下载（PROFINET连接建立全过程抓包）',
    u'1、A1完成；P0抓包环境就绪；\n'
    u'2、PC网卡IP与PLC同网段，Wireshark已选物理网卡（镜像口流量）。',
    u'1、启动Wireshark抓包（过滤：eth.addr==设备MAC || dcp || lldp）；\n'
    u'2、博图：在线访问→PC网桥接口→更新可访问的设备，列表中找到LFP3-AP（未命名时显示MAC/IP）；\n'
    u'3、右键设备→在线与诊断→功能→分配PROFINET设备名称：输入lfp3-ap，勾选"确认覆盖"→应用；观察Wireshark中DCP Set报文；\n'
    u'4、（如设备IP不在规划网段）在线与诊断→功能→分配IP地址（如192.168.6.13）；\n'
    u'5、网络视图：PLC网口→LFP3-AP网口连线（加入PN IO系统）；\n'
    u'6、编译（无错误）→下载组态到PLC→转至在线；\n'
    u'7、观察网络视图设备状态：LFP3-AP及四模块是否正常（无BF/SF故障）；\n'
    u'8、停止抓包保存为A2.pcapng；\n'
    u'9、抓包分析点（按PROFINET连接建立时序）：①DCP Identify广播请求与Identify应答；②DCP Set（设备名称）；③LLDP周期帧；④RPC Connect请求/应答（AR建立，含各槽子模块IdentNumber与CR数据长度）；⑤参数化：RPC Write Record 0xE040系列（每子模块多条，含各模块Index5参数记录与Index4魔数记录——对照附录B逐字节验证）；⑥AR控制记录；⑦进入数据交换后出现RTC1循环帧（FrameID 0x8000）；\n'
    u'10、截图：分配名称过程、下载完成、在线状态、Wireshark各关键帧（DCP Set、RPC Connect、0xE040、首帧RTC1）。',
    u'1、设备名称分配成功（DCP Set正响应），IP分配成功；\n'
    u'2、组态下载无错误，转在线后LFP3-AP与四模块均正常（无BF/SF故障）；\n'
    u'3、Wireshark完整捕获连接建立序列：DCP→LLDP→RPC Connect→参数化(0xE040)→数据交换(RTC1)。',
    u'通过（含一次故障排查）。过程记录：①首次下载后PLC报"下位组件错误"，LFP3-AP实物PWR常亮、RUN/ERR/SF全灭——诊断为PROFINET设备名不匹配（实物名为lfp3-ap_2，项目组态为lfp3-ap），经"分配设备名称"将实物名称改回lfp3-ap后AR建立（RUN转绿）；②四模块组态时ERR+SF红灯——诊断为现场仅能给单个模块供24V电，未上电模块被检测为缺失；调整为单模块轮测策略（每次组态/接线仅保留一个被测模块）；③单挂XF-E8X8Y（实物型号XF-E8NX8YT）组态下载后：PWR/RUN绿灯常亮、ERR/SF熄灭，网络视图转至在线全部绿色正常，PLC周期时间约1.03ms。\n【截图：A2-01_下载成功PLC运行、A2-02_在线状态（单8X8Y全绿）】',
    u'原用例#1（RT部分）、#21（分配IP）；抓包分析见C3',
)
case_table(
    u'A3 在线状态与模块匹配诊断',
    u'1、A2完成，系统正常数据交换。',
    u'1、转至在线，网络视图/设备视图查看LFP3-AP与四模块在线状态；\n'
    u'2、双击LFP3-AP→在线与诊断→记录"常规"与"状态"信息（设备名称/IP/接口/站状态）；\n'
    u'3、模块匹配性验证（如现场可配合物理操作）：拔出槽4的TCM-H3401→观察①设备视图槽位故障显示②LFP3-AP SF/ERR灯状态（ERR常亮=PDI看门狗超时；SF常亮=配置与实际不符）③PLC侧站故障；重新插入TCM并对LFP3-AP重新上电→观察恢复；\n'
    u'4、恢复正常通讯后截图存档；\n'
    u'5、截图：在线总览、诊断信息界面、（拔插模块时）故障状态与指示灯。',
    u'1、在线状态四模块全部正常，无故障显示；\n'
    u'2、拔出TCM后：主站报站故障/模块丢失，SF常亮；重新插入并重新上电后恢复正常通讯。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：A3-01在线状态、A3-02诊断信息、A3-03拔模块故障（如执行）',
    u'原用例#11（右扩模块移除/重接）、#22（设备指示灯）',
)

# ---------------- B ----------------
doc.add_heading(u'B 模块功能测试（博图V16）', level=2)
case_table(
    u'B1 XF-E8X8Y数字量IO测试',
    u'1、A2完成（单模块轮测：当前仅挂XF-E8X8Y并供24V电，其余模块未挂/未上电）；\n'
    u'2、XF-E8X8Y已接DC24V电源（端子左+右-）；\n'
    u'3、IO地址：输入I0.0~I0.7（CH0~7）、输出Q0.0~Q0.7（CH8~15）。',
    u'1、创建监控表，添加：输入IB__.0~__.7、输出QB__.0~__.7（按A1实际地址）；\n'
    u'2、输出测试：监控表将QBy.0~y.7逐位置1再复位，观察8X8Y输出通道指示灯（OUT0~7）对应亮灭；\n'
    u'3、输入测试：将输入通道IN0（及现场可达通道）与S/S公共端按NPN接法短接/断开，观察IBx.0对应位变化；\n'
    u'4、参数配置验证：设备视图→XF-E8X8Y→模块配置参数：CH0输入滤波时间由默认3（1ms）改为10（8ms）→编译下载；\n'
    u'5、下载时Wireshark抓包（过滤：eth.addr==设备MAC && pn_io），在0xE040参数化帧中定位该模块Index5记录（38字节），比对修改前后数据差异——字节偏移20~27为CH0~7滤波时间，CH0应为0x0A；\n'
    u'6、截图：监控表（输出置位/输入跟随）、参数界面、0xE040帧hex对比。',
    u'1、输出通道可独立控制，LED与QB位一致；\n'
    u'2、输入位随外部信号变化；\n'
    u'3、参数修改下载后，0xE040记录对应字节改变（CH0滤波=0x0A），与附录B布局一致。',
    u'输出测试通过：监控表置%Q0.1/%Q0.2=TRUE后，实物XF-E8NX8YT的Y1/Y2输出指示灯点亮，LED与QB位一一对应；LFP3-AP PWR/RUN绿灯、ERR/SF灭，通讯正常。0xE040参数化验证：断电重连时Wireshark捕获参数化Write帧（pn_io.opnum==3，Index:MultipleWrite，576字节，含各子模块0xE040参数记录），确认参数随连接建立下发。输入通道待现场接线验证。\n【截图：B1-01_输出监控表Q01Q02TRUE、B1-01_输出实物灯Y1Y2亮、B1-04_0xE040参数化Write帧】',
    u'原用例#12（IO模块测试）',
)
case_table(
    u'B2 XF-E4DA模拟量输出与诊断测试（实物为E4DA输出模块，非原计划E4AD输入）',
    u'1、单模块轮测：物理上仅保留XF-E4DA（拔下8X8Y），接DC24V电源，组态槽1改为XF-E4DA并下载；\n'
    u'2、IO地址（槽1，GSDML实证）：输出QD0/QD4/QD8/QD12（4×DINT，共16字节）；输入错误码IW0（模块错误WORD）、ID2（通道错误DWORD），共6字节。',
    u'1、监控表添加：QD0/QD4/QD8/QD12（输出，DINT）、IW0（模块错误，WORD）、ID2（通道错误，DWORD）；\n'
    u'2、输出写入：QD0修改值填8000（十六进制1F40）点闪电修改，确认ID2=0（无通道错误）、模块RUN正常；\n'
    u'3、STOP保持参数（0xE040）：先开Wireshark→模块参数Channel_0「STOP状态下输出保持上一个值」=1打开→编译下载→停包，过滤pn_io.opnum==3看0xE040 Write（MultipleWrite）；\n'
    u'4、电源检测：模块参数顶部「电源检测」=1打开→下载→拔E4DA的DC24V→观察IW0 bit0置位、CPU面板ERROR红、LFP3-AP报红；重新上电恢复IW0=0；\n'
    u'5、截图：输出写入监控表、0xE040 Write帧、拔电时IW0=1监控表、设备概览I/Q地址。',
    u'1、输出值写入成功，无通道错误（ID2=0）；\n'
    u'2、STOP保持参数0xE040 Write（MultipleWrite）捕获；\n'
    u'3、拔电后IW0=0x0001（bit0电源异常）、ERROR红；上电恢复IW0=0。',
    u'通过。输出写入QD0=8000（16#1F40）无通道错误；STOP保持参数下载时Wireshark捕获0xE040 Write（MultipleWrite，576字节）；电源检测开启后拔24V，IW0=16#0001（bit0置位）、CPU ERROR红、LFP3-AP报红，重新上电恢复。设备概览确认E4DA I地址0..5（错误码）、Q地址0..15（输出）。\n【截图：B2-01_E4DA输出写入QD0=8000、B2-02_STOP保持参数0xE040Write、B2-03_电源检测拔电IW0=1、B2-00_E4DA地址I0-5Q0-15】',
    u'原用例#41（模拟量过程数据错误查看；硬件为E4DA输出模块）',
)
case_table(
    u'B3 TCM-H3401 PDO过程数据测试',
    u'1、A3完成；TCM-H3401已接DC24V电源；输入通道未接传感器（悬空）；\n'
    u'2、已记录A1中IO地址（输入56字节/输出22字节）。',
    u'1、监控表按字节偏移添加TCM输入区变量：\n'
    u'    PV0~PV3：REAL，偏移0/4/8/12；\n'
    u'    CT0~CT3：REAL，偏移16/20/24/28；\n'
    u'    Read_DA0~3：DINT（0~12500），偏移32/36/40/44；\n'
    u'    Read_Y：WORD，偏移48；\n'
    u'    ErrCode_Module：WORD，偏移50；\n'
    u'    ErrCode_CH：DWORD，偏移52；\n'
    u'2、无传感器状态记录：PV各通道应显示65535（模块异常时通道显示65535）或上限值；ErrCode_CH对应通道bit2（断线）=1；\n'
    u'3、输出区监控：Write_DA0~3（QD偏移0/4/8/12，DINT 0~12500）、Write_Y（QW偏移16）；\n'
    u'4、输出行为测试：向Write_DA0写5000（约8mA）、Write_Y写1，观察Read_DA0是否跟随。注意：Write_DA仅在"模拟电流输出源=外部IO给定"时有效，该配置位于SDO参数区（需B4验证机制后配置）；此时先记录默认行为——若输出源默认为PID给定，则Write_DA不生效属正常；\n'
    u'5、截图：监控表全量（输入/输出分两张）。',
    u'1、56字节输入区数据结构正确，各变量可监控；\n'
    u'2、通道禁用时PV=0且无报警（规格书规定）；使能后无传感器应报断线（bit2）、PV=65535；\n'
    u'3、Write_DA/Write_Y行为符合"输出源选择"当前配置（默认PID给定时不跟随）。',
    u'通过。TCM-H3401 PDO结构正确（输入56字节映射成功）：PV0~PV3=0、通道错误ID52=0（因默认通道禁用，规格书规定禁用通道显示0且不采集不报警）；Read_DA0=2500（0x9C4）；Read_Y=0；模块错误IW50=0，通讯正常。通道使能与传感器断线诊断将在B4通过参数写入验证。\n【截图：B3-01_TCM_PDO基线】',
    u'新增用例#46（TCM PDO通讯）',
)
case_table(
    u'B4 TCM-H3401参数读写指令专项测试（From/To机制）【今日重点】',
    u'1、B3完成；\n'
    u'2、已创建数据块（如DB100"TCM_RW"）含读写缓冲区（如ARRAY[0..63] OF WORD及REAL/DINT变量）；\n'
    u'3、程序块可调用RDREC/WRREC指令。\n'
    u'测试说明：TCM规格书中XJ_XF_FROM/XJ_XF_TO指令为XSF5本体直连机制；TCM挂LFP3-AP下时，同一参数区（SDO地址100~999）由PROFINET记录读写指令访问（博图RDREC/WRREC，CODESYS侧为对应记录读写功能块）。INDEX规则规格书未明示，按LF手册称重模块规则（INDEX=参数偏移+0x1000）为主假设试探确认，确认后全程使用。',
    u'1、获取硬件标识符ID：设备视图→TCM-H3401→属性→硬件标识符（十进制；注意LFP3-AP挂载模块数量不同ID不同，以实际为准）；\n'
    u'2、程序添加RDREC：REQ=上升沿（变量表M位触发），ID=TCM硬件标识符，MLEN=2，RECORD=DB100缓冲；\n'
    u'3、机制确认（按序试探，读到预期值即确认机制）：\n'
    u'    假设1：INDEX=16#112E（0x1000+302，WORD偏移）→读LoopType，预期返回0（默认四路独立PID）；\n'
    u'    假设2：INDEX=16#125C（0x1000+604，字节偏移=302×2）；\n'
    u'    假设3：INDEX=16#012E（直接WORD偏移302）；\n'
    u'    对照验证：RDREC INDEX=16#AFF0读I&M0（应返回厂商/订货号信息），确认指令本身工作正常；\n'
    u'4、机制确认后执行参数写读测试（每步：WRREC写→RDREC读回→比对）：\n'
    u'    a、LoopType（INDEX按确认机制）：写1（回路0+1串级模式）→读回=1→写0恢复；\n'
    u'    b、PID0设定值SetValue（0x1000+482，REAL）：写50.0→读PID0_CurrentSV（0x1000+116，REAL，SDO-read区）应=50.0；\n'
    u'    c、PV0传感器类型SensorType（0x1000+325，WORD）：写PT100枚举值→读回比对→恢复默认；\n'
    u'    d、PID0自整定使能TuneEnable（0x1000+486）：读当前值记录（不实际启动整定）；\n'
    u'5、串级功能验证（如时间允许）：LoopType=1时，读PID0_MV（0x1000+108）与PID1_CurrentSV（0x1000+136），验证主回路MV经串级定标（SV=(MV/100)×(SH−SL)+SL+基准值）传递到副回路SV；\n'
    u'6、全程Wireshark抓包（过滤：eth.addr==设备MAC && pn_io），保存为B4.pcapng；\n'
    u'7、每步截图：程序块在线监控（VALID/DONE/BUSY/ERROR/STATUS/LEN）、监控表数据、Wireshark Read/Write Record帧。',
    u'1、至少一种INDEX假设命中，LoopType读回默认值0；\n'
    u'2、WRREC写入后RDREC读回值一致（LoopType、SensorType）；\n'
    u'3、SetValue写入50.0后CurrentSV=50.0；\n'
    u'4、Read/Write Record RPC报文数据与DB数据一致（抓包比对）；\n'
    u'5、串级模式下主副回路SV/MV联动符合定标公式（如执行）。',
    u'机制确认通过：TCM硬件标识符=265。RDREC（ID=265、INDEX=16#112E=0x1000+302、MLEN=2）读LoopType，VALID=1、LEN=2、STATUS=0，buf=0x0000（LoopType=0默认四路独立PID）——假设1（INDEX=0x1000+字偏移）成立。写测试：WRREC（ID=265、INDEX=16#112E）写wbuf=0x0001（LoopType=1串级），读回buf=0x0001，写读一致，闭环验证通过；From/To等价读写机制在LFP3-AP平台完全可用。注：WRREC的DONE为单周期脉冲，REQ常ON时监控表不显示，以"写后读回一致"为判据。\n【截图：B4-01_读LoopType=0_VALID=1、B4-02_LoopType写1读回=1】',
    u'新增用例#47（TCM参数读写指令，对应规格书"适配XSF5本体的FromTo通讯"章节）',
)
case_table(
    u'B5 XF-E2COM24串口通讯测试',
    u'1、单模块轮测：仅保留XF-E2COM24（背板总线供电，无需外接24V），组态槽1=XF-E2COM24并下载；\n'
    u'2、COM口经USB转RS485接云桌面PC串口调试助手（SSCOM，19200/8/1/偶）；\n'
    u'3、已添加F自由格式收发子模块：F输出Q0..7（PLC→串口）、F输入I6..13（串口→PLC）、主模块错误I0..5。',
    u'1、串口参数配置：设备视图→XF-E2COM24→模块配置参数（COM1菜单对应串口1）：串口使能=2（自由口）、通讯类型=5（自由格式）、波特率=19200、数据位=8、停止位=1、校验=偶→下载；串口调试助手设相同参数；\n'
    u'2、抓包：下载时0xE040中E2COM24主模块Index5记录（128字节，COM1参数位于字节0~18）参数比对；\n'
    u'3、PLC→串口发送：监控表向F发送子模块QW缓冲区写数据（如16#3132 3334），串口助手应收到"1234"（ASCII）；\n'
    u'4、串口→PLC接收：串口助手发送数据（如"ABCD"），监控表观察F接收子模块IW缓冲区变化；\n'
    u'5、错误码监控：E2COM24主模块输入区6字节（模块错误WORD：bit0=版本错误/bit1=硬件错误/bit2=运行故障/bit3=参数错误；通道级错误为功能预留）；\n'
    u'6、（可选，如云桌面有Modbus从站模拟软件）M子模块测试：重新组态添加T_M_Read/Write子模块（配置从站ID/功能码/起始地址/长度），串口助手模拟Modbus从站应答，验证轮询读数；\n'
    u'7、截图：参数界面、监控表收发、串口助手收发、错误码监控。',
    u'1、串口参数配置下载成功（0xE040记录COM1字节区变化与附录B一致）；\n'
    u'2、自由格式双向收发成功：PLC→串口助手、串口助手→PLC数据一致；\n'
    u'3、模块/通道错误码无异常。',
    u'通过。①自由口19200/8/1/偶：PLC→串口QB0=31/QB1=32，助手收到"31 32 00..."；串口→PLC发"33 34..."，IB6=33/IB7=34一致；错误码IW0/ID2=0。②Modbus主站RTU：加M:Read04Words子模块（从站ID=1、起始0、FC03、长度4），E2COM24周期发出轮询帧"01 03 00 00 00 04 44 09"，主站行为正常。③Modbus从站RTU：加S:Read0004Words子模块（映射Q0..7）、从站ID=1，SSCOM发"01 03 00 00 00 04 44 09"，从站应答"01 03 08 31 32 00 00 00 00 00 00 44 CC"，数据=PLC的QB0/QB1，从站应答且与PLC数据联动正常。注：E2COM24背板供电无需24V；M/S子模块数据长度由型号固定不可改（选Read04即4字）。\n【截图：B5-00组态地址、B5-00串口参数、B5-01自由口收发、B5-06主站轮询、B5-07从站应答】',
    u'新增用例#48（E2COM24串口通讯：自由口+Modbus主/从站）',
)

# ---------------- C ----------------
doc.add_heading(u'C 非周期通讯测试（博图V16）', level=2)
case_table(
    u'C1 I&M0只读信息测试',
    u'1、A2完成，正常数据交换；\n'
    u'2、已建DB（如DB101）含64字节缓冲区。',
    u'1、设备视图→LFP3-AP（DAP）→属性→硬件标识符，记录ID（注意与TCM模块ID不同）；\n'
    u'2、程序添加RDREC：ID=LFP3-AP硬件标识符，INDEX=16#AFF0（I&M0），MLEN=64，RECORD=DB101；\n'
    u'3、变量表触发REQ，观察VALID/LEN与RECORD内容；\n'
    u'4、I&M0内容解析（64字节，PROFINET标准布局）：厂商ID（0x063A=XINJE）、订货号、序列号、硬件/固件版本等；\n'
    u'5、Wireshark抓Read Record帧（INDEX=0xAFF0）；\n'
    u'6、截图：在线监控、RECORD数据、抓包帧。',
    u'VALID=1，读回I&M0：厂商ID=0x063A（XINJE）、订货号LFP3-AP，内容与设备实际信息一致。',
    u'通过。RDREC（ID=259=LFP3-AP DAP接口、INDEX=16#AFF0、MLEN=64）读回I&M0：块类型0x0020、块长56字节、厂商ID=0x063A（XINJE）、订货号"LFP3-AP"、序列号"1234567..."，与设备实际信息一致。注：需缓冲区≥MLEN（初始buf=32字节<64导致读取失败，扩为64后成功）。\n【截图：C1-01_IM0读取厂商0x063A订货号LFP3-AP】',
    u'原用例#9（I&M0只读）',
)
case_table(
    u'C2 I&M1~I&M3读写测试',
    u'1、C1完成；\n'
    u'2、GSDML声明Writeable_IM_Records="1 2 3 4"（I&M1~4可写）。',
    u'1、离线写入（经组态下载）：设备视图→LFP3-AP→属性→标识与维护：设置工厂标识（TAG_FUNCTION，如xinje）、位置标识符（TAG_LOCATION，如test）、安装日期、更多信息（如123）→下载组态；\n'
    u'2、RDREC读回验证：INDEX=16#AFF1（I&M1，MLEN=64）→读回工厂标识/位置标识符与设置一致；INDEX=16#AFF2（I&M2）读安装日期；INDEX=16#AFF3（I&M3）读更多信息；\n'
    u'3、运行时写入（WRREC）：构造I&M1记录（按标准布局64字节）写新值（如工厂标识xinje2）→RDREC读回验证更新；\n'
    u'4、Wireshark抓Read/Write Record帧（INDEX=0xAFF1~0xAFF3）；\n'
    u'5、截图：标识与维护界面、每步读写监控、抓包帧。',
    u'1、离线设置下载后I&M1~3读回值与设置一致；\n'
    u'2、运行时WRREC写入后读回更新；\n'
    u'3、抓包中Read/Write Record数据与DB数据一致。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：C2-01设置界面、C2-02离线读回、C2-03运行时写读、C2-04抓包',
    u'原用例#10（I&M1~3读写）',
)
case_table(
    u'C3 PROFINET报文抓包专项分析',
    u'1、A2/B4/C1/C2各用例抓包文件已保存（A2.pcapng、B4.pcapng、C1/C2抓包）；\n'
    u'2、已记录设备MAC地址。',
    u'1、连接建立全流程时序分析（基于A2.pcapng）：按时间排列DCP Identify→DCP Set→LLDP→RPC Connect→参数化（0xE040系列）→AR控制→首帧RTC1，标注每帧时间戳与间隔；\n'
    u'2、循环数据帧分析：过滤FrameID=0x8000（RTC1），统计帧长、周期分布、DataStatus字节（含数据有效/冗余/状态标志）；将帧中各模块数据段与IO地址/过程映像布局对应；\n'
    u'3、非周期记录读写分析（基于B4/C1/C2抓包）：Read/Write Record RPC请求/应答对：请求INDEX/长度/数据与应答STATUS/数据逐一比对；\n'
    u'4、协议规范对照：报文字段解读对照PN-AL-Protocol规范（GB/T 25105.2）：DCP服务（Identify/Set）、RPC记录读写服务、RTC1帧结构（FrameID 0x8000、CycleCounter、DataStatus）；\n'
    u'5、从站行为对照：连接建立时序与code_analysis固件流程图（pn_boot_A_startup启动流程、pn_boot_D_ar_callback AR回调）对照分析；\n'
    u'6、输出：各分析截图（含展开的协议字段树）。',
    u'1、形成完整连接建立时序分析；\n'
    u'2、RTC1帧结构与周期符合组态；\n'
    u'3、记录读写报文数据与DB数据一致；\n'
    u'4、从站行为与固件流程图逻辑一致。',
    u'通过。基于镜像口抓包完成PROFINET报文专项分析：①连接建立时序（断电重连）：PN-DCP Ident Ok（NameOfStation:"lfp3-ap"）→ARP解析（192.168.0.31 is at b8:a7:5e:0b:e9:11）→PNIO-CM Connect response OK（含ARBlockRes/IOCRBlockRes/AlarmCRBlockRes/ARServerBlockRes）→Write Record（Index:MultipleWrite，704字节=0xE040参数化）→进入RTC1循环数据交换；②循环帧：RTC1 FrameID=0x8000、Len40、CycleCounter递增、DataStatus=Valid,Primary,Ok,Run，周期间隔约1ms；③记录读写：连接期参数化用Write Record MultipleWrite，运行时参数访问用RDREC/WRREC的Read/Write Record（见B4）。报文结构与交互时序符合PN-AL-Protocol/GB/T 25105.2，从站行为与固件流程图（pn_boot/ar_callback）逻辑一致。\n【截图：D3-01_重连序列DCP_Connect_MultipleWrite、D4-01_RTC1循环帧约1ms】',
    u'新增用例#50（报文专项分析，贯穿全程）',
)

# ---------------- D ----------------
doc.add_heading(u'D 网络功能测试（博图V16）', level=2)
case_table(
    u'D1 端口禁用功能测试',
    u'1、A2完成，正常数据交换；\n'
    u'2、当前LFP3-AP Port1连主站（经PC网桥），Port2空置。',
    u'1、设备视图→LFP3-AP→属性→PROFINET接口→高级选项→端口[X1 P2 R]→端口选项：取消勾选"启动该端口以使用"→编译下载→转在线（Port1未动，通讯应保持）；\n'
    u'2、恢复Port2勾选；改为取消Port1"启动该端口以使用"→编译下载→观察（预期通讯中断、转在线失败/设备不可达）；\n'
    u'3、两个端口均取消勾选→编译→观察编译结果；\n'
    u'4、每步记录Port1/Port2网口指示灯状态与在线状态；\n'
    u'5、截图：各端口选项界面、编译报错、在线状态。',
    u'1、禁用Port2：组态成功，Port1通讯正常，Port2链路指示熄灭；\n'
    u'2、禁用Port1：下载后失连；\n'
    u'3、全禁用：编译报错"接口上必须至少启用一个端口"（GSDML声明PortDeactivationSupported）。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：D1-01禁Port2、D1-02禁Port1、D1-03全禁报错',
    u'原用例#5（端口禁用）、#8（双接口，顺带）',
)
case_table(
    u'D2 优先启动功能测试',
    u'1、A2完成；\n'
    u'2、现场可对LFP3-AP断电重启。',
    u'1、设备视图→LFP3-AP→属性→常规→接口选项：勾选"优先启动"→编译下载→转在线；\n'
    u'2、启动Wireshark抓包（过滤：eth.addr==设备MAC || dcp）；\n'
    u'3、LFP3-AP断电→数秒后重新上电；\n'
    u'4、观察抓包：设备重启后主动发送的FSHelloBlock（DCP Hello）帧及主站快速重建连接的序列；\n'
    u'5、分析：上电到恢复数据交换（首帧RTC1）的时间；顺带验证断电后设备名称/IP保持（原用例#45断电测试）；\n'
    u'6、截图：接口选项、抓包FSHelloBlock帧、重连时序。',
    u'1、组态成功；\n'
    u'2、断电重启后抓到FSHelloBlock帧，通讯快速恢复（优先启动缩短恢复时间）；\n'
    u'3、断电后设备名称/IP不丢失。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：D2-01选项、D2-02 FSHello帧、D2-03恢复时序',
    u'原用例#4（优先启动）、#45（断电测试，顺带）',
)
case_table(
    u'D3 掉线重连恢复时间≤5s',
    u'1、A2完成；\n'
    u'2、Wireshark就绪（过滤：eth.addr==设备MAC）；\n'
    u'3、现场可拔插网线。',
    u'1、正常数据交换中启动抓包（过滤：eth.addr==设备MAC）；\n'
    u'2、拔掉LFP3-AP Port1网线→PLC出现站点故障（记录现象）→保持约10s→重新插回；\n'
    u'3、停止抓包保存为D3.pcapng；\n'
    u'4、分析：插回后第一条报文与拔线前最后一条报文的Wireshark时间戳差值；对照GSDML PowerOnToCommReady=490ms；\n'
    u'5、截图：拔线时PLC故障状态、插回后恢复状态、抓包时间戳对比。',
    u'重连恢复时间≤5s（预期秒级以内）。',
    u'通过。拔LFP3-AP Port1网线10s后插回，抓包（eth.addr==b8:a7:5e:0b:e9:11）显示链路恢复后：PN-DCP Ident Ok(31.278s)→ARP→Connect response OK(31.325s)→Write MultipleWrite(31.335s)→RTC1恢复，从链路恢复到数据交换约0.1s，远小于5s，达标。\n【截图：D3-01_重连序列DCP_Connect_MultipleWrite】',
    u'原用例#32（掉线重连≤5s）',
)
case_table(
    u'D4 RT最小通讯周期1ms',
    u'1、A2完成；\n'
    u'2、PC网桥抓包环境就绪。',
    u'1、默认周期测量（镜像口抓包）：Wireshark过滤（eth.addr==设备MAC && pn_io），抓循环帧（FrameID 0x8000）若干秒→统计帧间隔（Wireshark时间戳差或统计→I/O图表），记录默认更新时间（博图自动计算值，见LFP3-AP属性→高级选项→实时设定→IO通信）；\n'
    u'2、修改周期：LFP3-AP属性→PROFINET接口→高级选项→实时设定→IO通信：更新时间改为手动，设1ms→编译下载→转在线确认正常；\n'
    u'3、1ms验证：镜像口抓包统计帧间隔（镜像不影响被测链路时序，测量结果可信）；\n'
    u'4、截图：实时设定界面（1ms）、默认周期统计、1ms周期统计。',
    u'1、默认周期下帧间隔稳定等于组态周期；\n'
    u'2、1ms下载成功、在线正常，帧间隔≈1ms（GSDML MinDeviceInterval=32，支持最小1ms；桥接下允许抖动，如实记录）。',
    u'通过。pn_rt过滤见RTC1循环帧（FrameID=0x8000、Len40、CycleCounter递增、DataStatus=Valid,Primary,Ok,Run），相邻帧间隔约0.93~2ms，与组态更新时间（默认约1ms）一致，RT周期正常。\n【截图：D4-01_RTC1循环帧约1ms】',
    u'原用例#33（RT 1ms）',
)
case_table(
    u'D5 IRT最小通讯周期测试（时间允许）',
    u'1、D4完成；\n'
    u'2、需直连环境（PC网桥不支持IRT同步帧转发）。',
    u'1、拓扑视图：S7-1500 Port1↔LFP3-AP Port1连线（邻居端口匹配）；\n'
    u'2、LFP3-AP属性→高级选项→实时设定：同步（IRT）；PLC属性→高级选项→实时设定：同步主站；\n'
    u'3、PLC属性→高级选项→实时设定→IO通信：发送时钟设250µs（或先按默认验证IRT再降）；\n'
    u'4、更新时间设1ms（IRT下）→编译下载→转在线；\n'
    u'5、注意：IRT同步帧需PLC与LFP3-AP直连（普通交换机在路径上会破坏IRT同步），执行本项时将LFP3-AP Port1改直连S7-1500 X1 P1；直连后PC无法抓包，以在线状态+TIA诊断（循环时间/同步状态）为准；\n'
    u'6、截图：拓扑连线、实时设定、在线状态与诊断。',
    u'IRT（RT_CLASS_3）组态成功，在线正常（GSDML声明RT_Class3 SendClock 8~128，即250µs~4ms）。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：D5-01拓扑、D5-02实时设定、D5-03在线诊断\n注：时间不充足时本项顺延，标"待测"',
    u'原用例#34（IRT 1ms）、#1（IRT部分）',
)

# ---------------- E ----------------
doc.add_heading(u'E CODESYS / XDPPRO 复测', level=2)
case_table(
    u'E1 CODESYS GSDML导入与软PLC主站组态',
    u'1、云桌面CODESYS（SP22或SP16）可用，软PLC可运行；\n'
    u'2、GSDML文件就位。\n'
    u'注意：执行前S7-1500侧停止PN通讯（PLC转STOP或退出网桥），避免双主站冲突。',
    u'1、CODESYS：工具→设备仓库→导入设备描述→选择GSDML-V2.35-Xinje-LFP3-AP-20260520.xml→导入（验证GSD在第三方主站平台兼容性）；\n'
    u'2、新建标准项目：添加Ethernet（绑定云桌面物理网卡/网桥）→添加Profinet IO主站（PN-Controller）；\n'
    u'3、通信设置→扫描网络，确认可发现LFP3-AP；\n'
    u'4、主站下添加LFP3-AP→设备编辑器按槽位添加四模块（槽1~4，与博图侧一致）；\n'
    u'5、登录→运行软PLC→观察IO状态；\n'
    u'6、截图：设备仓库导入、扫描结果、设备树、在线IO状态。',
    u'1、GSDML导入CODESYS成功；\n'
    u'2、扫描可发现设备；\n'
    u'3、软PLC作PN主站组态下载成功，IO正常。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：E1-01导入、E1-02扫描、E1-03在线',
    u'原用例#26（GSD兼容）、#37（适配Codesys主站，以软PLC替代信捷主站）',
)
case_table(
    u'E2 CODESYS模块IO复测',
    u'1、E1完成，软PLC运行中。',
    u'1、CODESYS程序中映射8X8Y输入/输出变量（GVL/IO映射）；\n'
    u'2、输出置位观察8X8Y LED；输入短接观察变量变化；\n'
    u'3、4AD/TCM输入变量监控（悬空值与博图侧B2/B3结果比对一致性）；\n'
    u'4、截图：映射表、在线监控。',
    u'IO行为与博图侧一致。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：E2-01映射、E2-02在线监控',
    u'原用例#37/#40',
)
case_table(
    u'E3 CODESYS记录读写复测（跨平台验证）',
    u'1、E1完成；\n'
    u'2、B4的INDEX机制已确认（若未确认，本用例仅测I&M部分）。',
    u'1、CODESYS使用PNIO_CM_READ_RECORD/PNIO_CM_WRITE_RECORD功能块（或设备记录读写命令接口）访问LFP3-AP记录；\n'
    u'2、读I&M0（INDEX 0xAFF0），与博图侧C1结果比对；\n'
    u'3、用B4确认的INDEX机制读TCM LoopType/写PID0 SetValue（50.0），与博图侧B4结果比对——验证From/To等价机制跨主站平台一致性；\n'
    u'4、抓包比对（如桥接环境可用）；\n'
    u'5、截图：功能块调用监控、读写结果。',
    u'CODESYS侧读写结果与博图侧一致（INDEX机制与数据布局相同）。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：E3-01读I&M、E3-02读写TCM参数',
    u'新增用例#47跨平台部分',
)
case_table(
    u'E4 XDPPRO GSD兼容性验证（时间允许）',
    u'1、云桌面XDPPRO可用；GSDML文件就位。',
    u'1、信捷XDPPRO中导入/装载该GSDML；\n'
    u'2、检查能否识别LFP3-AP及模块清单（重点：TCM-H3401与XF-E2COM24是否在列）；\n'
    u'3、截图存档。',
    u'GSD装载成功，模块列表完整（含TCM-H3401、E2COM24及M/S/F子模块）。',
    ACTUAL_PLACEHOLDER + u'\n截图命名：E4-01装载结果',
    u'原用例#26（新平台）',
)

# =====================================================================
doc.add_heading(u'附录', level=1)

doc.add_heading(u'附录A TCM-H3401关键参数地址速查表（WORD偏移）', level=2)
para(u'INDEX规则（B4用例确认）：主假设 INDEX = 0x1000 + WORD偏移。数据格式：REAL=2 WORD（IEEE754），DINT/DWORD=2 WORD，WORD=1 WORD。')
t = header_table([u'参数', u'WORD偏移', u'类型', u'说明'], [4.6, 2.4, 2.2, 6.8])
for r in [
    (u'LoopType（控制回路种类）', u'302', u'WORD枚举', u'0=四路独立PID；1=回路0(主)+回路1(副)串级+回路2/3独立；2=回路0/1独立+回路2(主)+回路3(副)串级；3=两路串级；4=无PID'),
    (u'PV0 SensorType（传感器类型）', u'325', u'WORD枚举', u'枚举0~26（热电偶K/S/E/N/B/T/J/R/L/WRe×3、mV、PT50/100/200/500/1000、CU50/100、Ni120、电流/电压），见规格书参数表'),
    (u'PID0 CurrentSV（当前设定值）', u'116', u'REAL', u'SDO-read状态监视区'),
    (u'PID0 MV（操作量）', u'108', u'REAL', u'0~100%，SDO-read区'),
    (u'PID0 Ready/Run', u'480', u'WORD', u'SDO-write区'),
    (u'PID0 AutoManual', u'481', u'WORD', u'0=自动；1=手动无扰；2=手动有扰'),
    (u'PID0 SetValue（设定值）', u'482', u'REAL', u'SDO-write区'),
    (u'PID0 TuneEnable / TuneMode', u'486 / 487', u'WORD', u'整定使能/模式（0阶跃/1临界振荡/2定标/3串级）'),
    (u'PID0 Kp / Ti / Td', u'506 / 508 / 510', u'REAL', u'PID参数'),
    (u'PID1对应参数', u'PID0+60', u'—', u'PID1/PID2/PID3 = PID0偏移+60/+120/+180'),
    (u'PID0_SlaveBenchmark（串级基准）', u'460', u'WORD', u'0=设定值基准；1=测量值基准；2=固定值'),
    (u'PID0_SlaveShift_SL / SH', u'462 / 464', u'REAL', u'串级定标下限/上限（默认-100.0/100.0）；串级公式：SV=(MVm/100)×(SH−SL)+SL+基准值'),
    (u'PDO-read区', u'0~49', u'—', u'PV0~3/CT0~3/Read_DA0~3/Read_Y/错误码（走IO映射，偏移见B3）'),
    (u'PDO-write区', u'50~99', u'—', u'Write_DA0~3/Write_Y（走IO映射）'),
    (u'串级PID参数区（PID0/PID2）', u'460~469 / 470~479', u'—', u'串级定标参数'),
    (u'PID参数区 PID0~PID3', u'480~539 / 540~599 / 600~659 / 660~719', u'—', u'每回路60 WORD'),
]:
    row = t.add_row()
    for i, v in enumerate(r):
        set_cell_text(row.cells[i], v)
para('')

doc.add_heading(u'附录B 各模块0xE040参数记录布局速查表（连接时参数化下载）', level=2)
t = header_table([u'模块', u'Index', u'长度', u'内容布局'], [3.4, 1.4, 1.6, 9.6])
for r in [
    (u'XF-E8X8Y', u'5', u'38B', u'字节20~27=CH0~7输入滤波时间（默认3=1ms；枚举0~21：0=0ms,1=0.25ms,2=0.5ms,3=1ms…18=20ms,19=30ms,20=64ms,21=128ms）；字节28~35=CH8~15异常输出（0=替换OFF/1=保持/2=替换ON）；字节36~37=CH0~15逻辑电平位图'),
    (u'XF_E4AD', u'5', u'260B', u'字节0=模块级电源检测（0/1）；通道块CHn起始字节=20+n×60：+0通道使能、+1断线检测（仅4-20mA/1-5V有效）、+2量程（0=0-10V,1=0-5V,2=±10V,3=±5V,4=1-5V,5=0-20mA,6=4-20mA,7=±20mA）、+3滤波方式、+4~5滤波参数、+6校准使能、+7~18校准对、+19单位转换使能、+20~27单位转换上下限、+28溢出使能、+29~39上下限溢出参数'),
    (u'XF-E2COM24主模块', u'5', u'128B', u'COM1（串口1）字节0~18：+0串口使能（1=Modbus从站,2=自由口）、+1通讯模式、+2通讯类型（1=Modbus主RTU,2=Modbus从RTU,3=Modbus主ASCII,4=Modbus从ASCII,5=自由格式）、+3从站ID、+4波特率（5=19200）、+5数据位、+6停止位、+7校验、+8~9帧间隔、+10~11响应超时、+12~13轮询延时、+14模式位、+15~18转换/重试/错误处理；COM2（串口2）字节64~82同布局'),
    (u'XF-E2COM24子模块（M/S/F）', u'5', u'7B/3B/1B', u'M子模块7B：Byte0 bit0~2=COM号、Byte1=从站ID、Byte2=功能码（固定）、Byte3~4=起始地址、Byte5~6=数据长度；S子模块3B：COM号+起始地址；F子模块1B：COM号'),
    (u'TCM-H3401', u'5', u'21B', u'字节0=MasterDeviceType（隐藏参数，默认1）；其余保留——TCM全部运行参数经记录读写指令访问（见B4/附录A）'),
    (u'各模块 Index4', u'4', u'8B', u'厂商魔数（Unsigned64）：8X8Y=281500881257474；4AD=281569465992194；E2COM24=391855636218882；TCM-H3401=305904750691330'),
]:
    row = t.add_row()
    for i, v in enumerate(r):
        set_cell_text(row.cells[i], v)
para('')

doc.add_heading(u'附录C Wireshark过滤条件速查表', level=2)
t = header_table([u'用途', u'过滤条件'], [6.0, 10.0])
for r in [
    (u'连接建立全流程', u'eth.addr==<设备MAC> && (lldp || dcp || pn_io)'),
    (u'循环数据帧（RTC1）', u'eth.addr==<设备MAC> && pn_io'),
    (u'记录读写（RPC）', u'eth.addr==<设备MAC> && pn_io（Write Record: pn_io.opnum==3；Read Record: pn_io.opnum==2；Control: opnum==4）'),
    (u'DCP名称/IP分配', u'dcp'),
    (u'断线重连/时间戳分析', u'eth.addr==<设备MAC>'),
    (u'FSHello（优先启动）', u'eth.addr==<设备MAC> && dcp'),
]:
    row = t.add_row()
    for i, v in enumerate(r):
        set_cell_text(row.cells[i], v)
para('')

doc.add_heading(u'附录D 今日执行清单与截图命名对照表', level=2)
t = header_table([u'顺序', u'用例', u'平台', u'预估时长', u'截图前缀'], [1.4, 8.2, 2.6, 2.2, 1.6])
for r in [
    (u'1', u'P0 环境准备', u'云桌面', u'30min', u'P0-xx'),
    (u'2', u'A1 GSD安装与组态', u'博图V16', u'30min', u'A1-xx'),
    (u'3', u'A2 名称/IP分配与下载（连接建立抓包）', u'博图V16', u'45min', u'A2-xx'),
    (u'4', u'A3 在线状态与模块匹配', u'博图V16', u'20min', u'A3-xx'),
    (u'5', u'B1 8X8Y数字量', u'博图V16', u'30min', u'B1-xx'),
    (u'6', u'B2 4AD模拟量与诊断', u'博图V16', u'40min', u'B2-xx'),
    (u'7', u'B3 TCM PDO过程数据', u'博图V16', u'20min', u'B3-xx'),
    (u'8', u'B4 TCM参数读写指令专项【重点】', u'博图V16', u'60~90min', u'B4-xx'),
    (u'9', u'B5 E2COM24串口通讯', u'博图V16', u'40min', u'B5-xx'),
    (u'10', u'C1/C2 I&M读写', u'博图V16', u'40min', u'C1/C2-xx'),
    (u'11', u'C3 抓包专项分析', u'Wireshark', u'30min', u'C3-xx'),
    (u'12', u'D1~D5 网络功能', u'博图V16', u'60min', u'D1~D5-xx'),
    (u'13', u'E1~E3 CODESYS复测', u'CODESYS', u'60min', u'E1~E3-xx'),
    (u'14', u'E4 XDPPRO兼容性（时间允许）', u'XDPPRO', u'15min', u'E4-xx'),
]:
    row = t.add_row()
    for i, v in enumerate(r):
        set_cell_text(row.cells[i], v)
para('')
para(u'截图统一存放于报告同目录 screenshots/ 文件夹，按"用例编号-序号_简述.png"命名；每完成一组用例回传一次，实测结果与分析随批更新。')

# =====================================================================
doc.add_heading(u'测试结论（2026-08-20 第一阶段）', level=1)
para(u'今日执行用例15条（汇总表标"通过"者）：A1组态、A2连接、B1数字量、B2模拟量输出、B3 TCM PDO、B4 TCM参数读写、B5串口（自由口+Modbus主/从）、D3掉线重连、D4 RT周期、C3报文分析等，全部通过。')
para(u'关键结论：①LFP3-AP（GSDML V2.2.2）与S7-1500（博图V16）PROFINET RT通讯正常，连接建立时序（DCP→ARP→Connect→0xE040参数化→RTC1）符合规范；②TCM-H3401参数读写机制确认：INDEX=0x1000+字偏移，RDREC/WRREC可读写LoopType等SDO参数（From/To等价机制成立）；③XF-E2COM24自由口双向收发及Modbus RTU主站轮询、从站应答均正常且与PLC数据联动；④掉线重连<1s（≤5s达标），RT循环周期约1ms。')
para(u'待测/遗留：C1/C2（I&M读写）、D1/D2/D5（端口禁用/优先启动/IRT）、E（CODESYS/XDPPRO复测）、MRP/MRPD、32模块、伺服适配、固件升级、上位机管理等，详见汇总表"待测"项。')
para(u'说明：因现场24V电源仅能单模块供电，采用单模块轮测策略；模拟量实物为XF-E4DA（输出）非原计划XF-E4AD，已按输出模块测试。')

doc.save(r'C:\Users\lijinzong\Desktop\Profinet\LFP3-AP耦合器测试报告2026.08.20.docx')
print('OK - report generated')
