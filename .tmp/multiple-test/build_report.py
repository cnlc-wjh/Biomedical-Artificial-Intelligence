from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[2]
OUTDIR = ROOT / "作业报告"
OUTDIR.mkdir(exist_ok=True)
OUT = OUTDIR / "多重检验校正_初学者学习报告_提交版.docx"
FIGDIR = ROOT / ".tmp" / "multiple-test" / "figures"
FIGDIR.mkdir(parents=True, exist_ok=True)
RESULTS_CHART = FIGDIR / "multiple_testing_discoveries.png"
FLOW_CHART = FIGDIR / "analysis_flow.png"

NAVY = "15324B"
TEAL = "147D83"
PALE = "EAF3F4"
PALE_BLUE = "EEF3F8"
GRAY = "D9E0E5"
INK = "202A33"
MUTED = "5D6B75"
WHITE = "FFFFFF"
BODY_FONT = "Microsoft YaHei"


def image_font(size, bold=False):
    path = Path("C:/Windows/Fonts") / ("msyhbd.ttc" if bold else "msyh.ttc")
    return ImageFont.truetype(str(path), size=size)


def make_discovery_chart(path):
    w, h = 1800, 720
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    d.text((90, 42), "同一组数据，三种判定口径", font=image_font(42, True), fill="#202A33")
    d.text((92, 103), "训练样本中的显著探针数（共检验 7,129 个探针）", font=image_font(25), fill="#5D6B75")
    left, right, top, base = 205, 1680, 185, 590
    maxv = 1800
    for val in (0, 500, 1000, 1500):
        y = base - (base-top) * val / maxv
        d.line((left, y, right, y), fill="#D9E0E5", width=2)
        d.text((110, y-16), f"{val:,}", font=image_font(21), fill="#5D6B75")
    bars = [
        ("原始 p ≤ 0.05", 1694, "#D97456", "未校正"),
        ("Bonferroni", 26, "#3976A8", "控制整组至少一次误报"),
        ("BH adjusted p", 535, "#16858B", "控制发现集合的长期平均假发现比例"),
    ]
    centers = [480, 940, 1400]
    bw = 210
    for (label, value, color, sub), cx in zip(bars, centers):
        bh = max(5, (base-top) * value / maxv)
        x0, x1 = cx-bw//2, cx+bw//2
        d.rounded_rectangle((x0, base-bh, x1, base), radius=14, fill=color)
        d.text((cx, base-bh-48), f"{value:,}", font=image_font(31, True), fill="#202A33", anchor="mm")
        d.text((cx, base+39), label, font=image_font(23, True), fill="#202A33", anchor="mm")
        d.text((cx, base+82), sub, font=image_font(17), fill="#5D6B75", anchor="mm")
    d.line((left, base, right, base), fill="#75838D", width=3)
    im.save(path)


def make_flow_chart(path):
    w, h = 1800, 360
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    boxes = [
        ("发现", "train 38 份样本\n完整检验 7,129 项", "#EAF3F4"),
        ("锁定", "按预设规则\n固定 10 个候选", "#EEF3F8"),
        ("检查", "test 34 份样本\n只检查锁定项目", "#EAF3F4"),
        ("结果", "7 个通过候选内\nBonferroni 与方向规则", "#EEF3F8"),
        ("下一步", "独立队列与\n生物学验证", "#F8EFEA"),
    ]
    x0, gap, bw, y0, bh = 35, 35, 320, 84, 205
    for i, (title, text, fill) in enumerate(boxes):
        x = x0 + i*(bw+gap)
        d.rounded_rectangle((x, y0, x+bw, y0+bh), radius=18, fill=fill, outline="#CCD7DE", width=3)
        d.text((x+bw//2, y0+43), title, font=image_font(29, True), fill="#15324B", anchor="mm")
        for j, line in enumerate(text.split("\n")):
            d.text((x+bw//2, y0+104+j*39), line, font=image_font(19), fill="#263746", anchor="mm")
        if i < len(boxes)-1:
            ax = x+bw+4
            cy = y0+bh//2
            d.line((ax, cy, ax+gap-10, cy), fill="#147D83", width=5)
            d.polygon([(ax+gap-10,cy-11),(ax+gap,cy),(ax+gap-10,cy+11)], fill="#147D83")
    im.save(path)


make_discovery_chart(RESULTS_CHART)
make_flow_chart(FLOW_CHART)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn("w:" + m))
        if node is None:
            node = OxmlElement("w:" + m)
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color=GRAY, size="5"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_run_font(run, size=None, bold=None, color=INK, font=BODY_FONT):
    run.font.name = font
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), font)
    run._element.rPr.rFonts.set(qn("w:ascii"), font)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), font)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def style_paragraph(p, after=5, before=0, line=1.18):
    pf = p.paragraph_format
    pf.space_after = Pt(after)
    pf.space_before = Pt(before)
    pf.line_spacing = line
    pf.widow_control = True
    return p


def para(doc, text="", bold_lead=None, size=10.4, color=INK, after=5, italic=False):
    p = doc.add_paragraph()
    style_paragraph(p, after=after)
    if bold_lead and text.startswith(bold_lead):
        set_run_font(p.add_run(bold_lead), size=size, bold=True, color=NAVY)
        set_run_font(p.add_run(text[len(bold_lead):]), size=size, color=color)
    else:
        set_run_font(p.add_run(text), size=size, color=color)
    if italic:
        for r in p.runs:
            r.italic = True
    return p


current_num_id = None


def heading(doc, text, level=1):
    global current_num_id
    current_num_id = None
    p = doc.add_paragraph(style=f"Heading {level}")
    style_paragraph(p, after=7, before=8 if level == 1 else 5, line=1.05)
    r = p.add_run(text)
    set_run_font(r, size=16 if level == 1 else 12.5, bold=True, color=INK)
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    style_paragraph(p, after=3, line=1.12)
    set_run_font(p.add_run(text), size=10.1, color=INK)
    return p


def numbered(doc, text):
    global current_num_id
    p = doc.add_paragraph(style="List Number")
    if current_num_id is None:
        numbering = doc.part.numbering_part.element
        abstract = next(a for a in numbering.findall(qn("w:abstractNum"))
                        if any(lvl.find(qn("w:numFmt")) is not None and
                               lvl.find(qn("w:numFmt")).get(qn("w:val")) == "decimal"
                               for lvl in a.findall(qn("w:lvl"))))
        current_num_id = max(int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))) + 1
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(current_num_id))
        abstract_ref = OxmlElement("w:abstractNumId")
        abstract_ref.set(qn("w:val"), abstract.get(qn("w:abstractNumId")))
        num.append(abstract_ref)
        numbering.append(num)
    p_pr = p._p.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        p_pr.append(num_pr)
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numid = OxmlElement("w:numId")
    numid.set(qn("w:val"), str(current_num_id))
    num_pr.extend([ilvl, numid])
    style_paragraph(p, after=4, line=1.12)
    set_run_font(p.add_run(text), size=10.1, color=INK)
    return p


def formula(doc, text, note=None):
    p = doc.add_paragraph()
    style_paragraph(p, after=4)
    p.paragraph_format.left_indent = Inches(0.22)
    p.paragraph_format.right_indent = Inches(0.18)
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), PALE)
    p_pr.append(shd)
    set_run_font(p.add_run(text), size=10.1, bold=True, color=NAVY, font="Consolas")
    if note:
        p2 = doc.add_paragraph()
        style_paragraph(p2, after=5)
        p2.paragraph_format.left_indent = Inches(0.22)
        set_run_font(p2.add_run(note), size=9.2, color=MUTED)


def code(doc, lines, caption=None):
    if caption:
        p = doc.add_paragraph()
        style_paragraph(p, before=4, after=3)
        set_run_font(p.add_run(caption), size=9.4, bold=True, color=TEAL)
    for line in lines.splitlines():
        p = doc.add_paragraph()
        style_paragraph(p, after=0, line=1.02)
        p.paragraph_format.left_indent = Inches(0.16)
        p.paragraph_format.right_indent = Inches(0.10)
        p_pr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "F3F6F8")
        p_pr.append(shd)
        set_run_font(p.add_run(line if line else " "), size=8.2, color="263746", font="Consolas")


def make_table(doc, headers, rows, widths=None, font_size=8.8, center_cols=(), cell_padding=90):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    set_repeat_table_header(table.rows[0])
    for j, text in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell, top=cell_padding, bottom=cell_padding)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style_paragraph(p, after=0, line=1.0)
        set_run_font(p.add_run(str(text)), size=font_size, bold=True, color=WHITE)
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j, value in enumerate(row):
            cell = cells[j]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell, top=cell_padding, bottom=cell_padding)
            if i % 2 == 1:
                set_cell_shading(cell, "F5F8FA")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in center_cols else WD_ALIGN_PARAGRAPH.LEFT
            style_paragraph(p, after=0, line=1.05)
            set_run_font(p.add_run(str(value)), size=font_size, color=INK)
    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(1)
    spacer.paragraph_format.space_before = Pt(0)
    return table


def add_page_field(p):
    r = p.add_run()
    set_run_font(r, size=8, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r._r.extend([begin, instr, separate, text, end])


def page_break(doc):
    doc.add_page_break()


doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.68)
sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.78)
sec.right_margin = Inches(0.78)
sec.header_distance = Inches(0.30)
sec.footer_distance = Inches(0.30)

styles = doc.styles
for name in ("Normal", "Title", "Subtitle", "Heading 1", "Heading 2", "List Bullet", "List Bullet 2", "List Number"):
    st = styles[name]
    st.font.name = BODY_FONT
    st._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    st._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
    st._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
styles["Normal"].font.size = Pt(10.4)
styles["Normal"].font.color.rgb = RGBColor.from_string(INK)

header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
style_paragraph(header, after=0)
set_run_font(header.add_run("课程作业报告  |  多重检验校正"), size=8, color=MUTED)
footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_paragraph(footer, after=0)
set_run_font(footer.add_run("学习版  ·  依据课程数据与分析协议整理     第 "), size=8, color=MUTED)
add_page_field(footer)
set_run_font(footer.add_run(" 页"), size=8, color=MUTED)

# Cover
p = doc.add_paragraph()
style_paragraph(p, after=4, before=18)
set_run_font(p.add_run("统计学课程作业报告"), size=12, bold=True, color=TEAL)
p = doc.add_paragraph(style="Title")
style_paragraph(p, after=3, line=1.0)
set_run_font(p.add_run("多重检验校正"), size=30, bold=True, color=INK)
p = doc.add_paragraph()
style_paragraph(p, after=12)
set_run_font(p.add_run("怎样从显著结果走向可信发现"), size=17, color=NAVY)
chartp = doc.add_paragraph()
chartp.alignment = WD_ALIGN_PARAGRAPH.CENTER
chartp.paragraph_format.space_after = Pt(3)
chartp.add_run().add_picture(str(RESULTS_CHART), width=Inches(6.85))
cp = doc.add_paragraph()
cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_paragraph(cp, after=8)
set_run_font(cp.add_run("图 1  真实数据的发现数因错误控制目标而改变。"), size=8.4, color=MUTED)
para(doc, "研究对象：Golub 白血病表达数据中的 AML 与 ALL 比较", size=10.3, color=NAVY, after=3)
para(doc, "阅读对象：刚开始学习统计推断与 Python 的学生", size=10.0, color=MUTED, after=2)
para(doc, "核心结论：当一次研究做了成千上万次检验时，不能只凭原始 p 值判断发现是否可靠。校正能帮助控制某类错误，但不能替代效应大小、研究设计和独立验证。", size=10.4, after=4)

# Page 2
page_break(doc)
heading(doc, "先看结论", 1)
para(doc, "本作业比较 AML（急性髓系白血病）与 ALL（急性淋巴细胞白血病）的表达差异。训练部分有 38 个样本，包含 7,129 个微阵列探针。若每个探针都按原始 p≤0.05 判断，得到 1,694 个显著探针；经过 Bonferroni 校正后为 26 个，经过 Benjamini–Hochberg（BH）校正后为 535 个。不同数量来自不同的错误控制目标，不能把任意一个数量直接当成真实基因数。")
make_table(doc,
    ["本报告回答的问题", "先记住的答案"],
    [
        ["为什么检验很多次会有问题？", "每次检验都留有随机误报机会；把很多次结果合并解读，整体误报风险会增加。"],
        ["Bonferroni 和 BH 各控制什么？", "Bonferroni 主要控制整组至少一次误报的风险（FWER）；BH 控制发现集合中假发现比例的长期平均（FDR）。"],
        ["535 个 BH 显著探针意味着什么？", "它们是在本分析和假设条件下值得进一步检查的统计候选，不等同于 535 个已证实疾病基因。"],
        ["检验后还要看什么？", "效应方向与大小、置信区间、探针注释、数据质量和独立研究中的重复结果。"],
    ], widths=[2.05, 4.55], font_size=9.0)
heading(doc, "学习路线", 2)
numbered(doc, "先看懂 p 值、显著性阈值和多重检验，再理解 FWER 与 FDR。")
numbered(doc, "用同一组简单 p 值手算 Bonferroni 与 BH，弄清两种方法的差别。")
numbered(doc, "查看真实数据结果，并把统计数字翻译成有限度的结论。")
numbered(doc, "最后阅读代码解释，知道每一段程序输入什么、输出什么。")

# Page 3
page_break(doc)
heading(doc, "1  为什么做很多次检验会误报", 1)
para(doc, "可以把一次统计检验想成在问一个具体问题，例如：“AML 与 ALL 在这个探针上的平均表达量相同吗？”零假设 H0 说“平均值相同”。若 H0 真的成立，检验仍有小概率因为样本随机波动而错误地报出差异，这叫假阳性或第一类错误。")
para(doc, "显著性阈值 α=0.05 的直观含义是：在模型和零假设成立时，若反复做同一种单项检验，错误拒绝零假设的比例约为 5%。它不是“研究结论有 95% 概率是真的”，也不是“效应有多大”。")
heading(doc, "用一千个完全没差别的变量来想", 2)
formula(doc, "预期假阳性数 = 真零假设数 × 单项错误率 = 1,000 × 0.05 = 50", "这说的是很多次实验的平均数量，不是每一次恰好出现 50 个。期望数量的计算用到期望的线性性质，本身不要求各检验相互独立；但把“一次也不误报的概率”算成固定乘积时，需要独立性假设。")
para(doc, "如果一千个零假设检验彼此独立，至少出现一个假阳性的概率为：")
formula(doc, "FWER = 1 − (1 − 0.05)^1,000 ≈ 1", "实际值极接近 1。此处用独立性公式只作直觉说明；真实探针之间可能相关，不能直接把这个独立公式照搬到真实数据。")
para(doc, "所以问题不在于单个 p 值突然失去定义，而在于不能把大量未经校正的单项 p 值集合，直接当成一组可靠发现。若再尝试多个亚组、模型和结局，实际“试过多少种办法”也会影响可信度。", bold_lead="所以问题不在于")
heading(doc, "几个容易混淆的词", 2)
make_table(doc,["词语", "通俗解释", "不要误读成"], [
    ["零假设 H0", "本次检验暂时采用的基准说法，例如两组平均表达相同。", "“现实中绝对没有任何作用”的证明。"],
    ["α = 0.05", "单项检验预先选定的误报容许水平。", "结果为真的概率是 95%。"],
    ["多重检验家族", "这次研究中一起考虑的一组相关问题。", "只包含最后写进报告的显著结果。"],
], widths=[1.3, 3.5, 1.8], font_size=8.5)

# Page 4
page_break(doc)
heading(doc, "2  先分清 FWER 和 FDR", 1)
para(doc, "假设研究者把通过门槛的项目都放进候选名单。记 R 为名单总数，V 为其中的假发现，S 为真实信号，于是 R=V+S。模拟实验里我们事先知道哪些变量真的有差异，因此能够逐项统计 V 和 S；真实白血病数据没有单项真值标签，不能直接数出真实的假发现。")
make_table(doc,["指标", "它问的问题", "一句话记忆"], [
    ["FDP", "这一轮已经报出的名单里，假发现占了多少？", "本次名单的错误比例。"],
    ["FDR", "把研究重复很多次，每次先算 FDP 再取平均是多少？", "长期平均的名单错误比例。"],
    ["FWER", "重复研究时，至少出现一个假发现的概率是多少？", "整组里有没有错一个。"],
    ["Power 检出力", "所有真实信号里，有多少比例被方法找出来？", "找回真信号的能力。"],
], widths=[1.2, 3.45, 1.95], font_size=8.6)
formula(doc, "FDP = V / max(R, 1)      FDR = E(FDP)      FWER = P(V > 0)", "如果本次没有任何发现，就约定 FDP=0，避免分母为零。FDR 是 FDP 的期望，不是“每个发现正确的概率”。")
heading(doc, "全零时，两个错误指标会碰巧相同", 2)
para(doc, "全零实验指每个变量都确知没有真实差异。有发现时，所有发现都是假的，FDP=1；没有发现时，FDP=0。因此，在这个特殊情形下，重复实验的 FDR 与 FWER 相同。只要存在真实信号，两者就可以差得很远。")
para(doc, "例如一千次模拟里，BH 的 FWER 估计约为 0.754，而 FDR 估计约为 0.046（平均约 4.6%）。这并不矛盾：很多次实验可能至少错中一个，但每次名单里假发现的平均占比仍很低。")
formula(doc, "FDR ≠ 平均 V ÷ 平均 R", "每次先算 FDP=V/max(R,1)，再对 1,000 次 FDP 求平均。一般不能用平均 V 除以平均 R 代替。")

# Page 5
page_break(doc)
heading(doc, "3  两种校正方法怎样工作", 1)
heading(doc, "Bonferroni：把单项门槛调严", 2)
para(doc, "若一组有 m 次检验，Bonferroni 把每项 p 值乘以 m（上限为 1），等价于要求 raw p≤α/m。它主要控制整组至少一次假发现的风险；检验多时门槛严格，真信号也容易漏掉。")
formula(doc, "Bonferroni adjusted p = min(m × raw p, 1)", "先按研究计划定义完整家族，不能看完 p 值后再缩小家族。")
heading(doc, "BH：按排序后的 p 值决定发现集合", 2)
para(doc, "BH 先把 m 个 p 值从小到大排序，比较第 i 项 p(i) 与 i×q/m（常取 q=0.05）。找到最大满足位置 k，把前 k 项列为发现。BH 主要控制 FDR，通常比 Bonferroni 保留更多候选。")
formula(doc, "在排序位置 i：p(i) ≤ i × q / m", "BH 的保证需要有效 p 值及独立或适当正相关等条件；任意相关结构下不作无条件保证。")
heading(doc, "手算小例子", 2)
para(doc, "假设这组里只有 5 个 p 值：0.001、0.009、0.020、0.040、0.200，阈值 q=0.05。排序已经从小到大。")
make_table(doc,["i", "p(i)", "Bonferroni", "BH 门槛", "满足？"], [
    ["1", "0.001", "0.005", "0.010", "是"],
    ["2", "0.009", "0.045", "0.020", "是"],
    ["3", "0.020", "0.100", "0.030", "是"],
    ["4", "0.040", "0.200", "0.040", "是"],
    ["5", "0.200", "1.000", "0.050", "否"],
], widths=[0.55, 1.0, 1.65, 1.7, 1.4], font_size=8.2, center_cols=(0,1,2,3,4), cell_padding=50)
para(doc, "BH 最大满足位置为 k=4，所以 BH 报出前 4 个。Bonferroni 只报前 2 个。BH 放宽一些，通常多保留候选；两种方法回答的是不同的错误控制目标，结果不同不表示谁算错。")

# Page 6
page_break(doc)
heading(doc, "4  这次真实数据分析做了什么", 1)
para(doc, "课程数据来自 Golub 白血病表达数据。数据表共有 72 份样本和 7,129 个数值探针；原课程把前 38 份作为 train（ALL 27 份、AML 11 份），后 34 份作为 test（ALL 20 份、AML 14 份）。本练习的问题是两种白血病亚型之间的表达差异，不是白血病与健康人的差异。")
make_table(doc,["分析步骤", "本作业事先固定的选择"], [
    ["研究问题", "比较 AML 与 ALL 的平均探针表达。"],
    ["发现样本", "train：ALL 27 份、AML 11 份。"],
    ["完整检验家族", "全部 7,129 个探针；每个探针各做一次双侧检验。"],
    ["单项方法", "Welch t 检验；允许两组的样本方差不同。"],
    ["效应方向", "AML 平均值减 ALL 平均值，正数表示 AML 较高。"],
    ["阈值与校正", "原始、Bonferroni、BH adjusted p 均以 0.05 作判断。"],
], widths=[1.55, 5.05], font_size=8.7)
heading(doc, "为什么要把协议放在看结果之前", 2)
para(doc, "如果看过所有 p 值后才决定要测哪些探针、选哪个模型或报告哪个亚组，研究者其实尝试的机会比报告里看起来的更多。多重检验家族应尽量对应研究问题和预先计划。先挑显著项目再对挑出的项目校正，会忘记最初的挑选过程。")
heading(doc, "真实数据的主要结果", 2)
make_table(doc,["判定方式", "显著探针数", "怎么理解"], [
    ["raw p ≤ 0.05", "1,694", "未作多重检验校正；数字较多。"],
    ["Bonferroni adjusted p ≤ 0.05", "26", "以整组至少一次误报风险为主要控制目标。"],
    ["BH adjusted p ≤ 0.05", "535", "以发现集合的假发现比例长期平均为主要控制目标。"],
], widths=[2.05, 1.1, 3.45], font_size=8.8, center_cols=(1,))
formula(doc, "Bonferroni 原始 p 阈值 = 0.05 / 7,129 ≈ 0.00000701", "BH 在第 535 个排序位置达到其规则门槛：该位置 raw p≈0.0037466，BH 门槛约 0.0037523。")
para(doc, "解读边界：真实探针可能彼此相关。本课按 BH 方法计算并报告 535 个候选，但没有全面检验相关结构是否满足 BH 的理论保证条件，因此不要把这个数字理解成任何相关结构下都严格保证 FDR≤0.05。")

# Page 7
page_break(doc)
heading(doc, "5  显著结果还要配上效应大小", 1)
para(doc, "一个小 p 值说明观察到的数据与某个零假设不太相容；它不会告诉我们两组相差多少，也不会说明差异对生物学或临床有多重要。因此报告还应写出组均值、差值方向、置信区间和样本量。")
heading(doc, "例子：U22376_cds2_s_at 探针", 2)
make_table(doc,["指标", "训练样本结果", "初学者读法"], [
    ["AML 平均表达", "702.36", "AML 组的平均测量值。"],
    ["ALL 平均表达", "3,863.30", "ALL 组的平均测量值。"],
    ["差值 AML−ALL", "−3,160.93", "负号表示 AML 平均测量值更低。"],
    ["逐项 95% 置信区间", "[−3,977.85, −2,344.02]", "在本模型下，这项差值估计的不确定范围；不是整个 7,129 项的同时区间。"],
    ["Cohen d", "−1.8484", "以组内标准差作单位的描述性效应大小。"],
    ["raw p", "约 0.00000000839", "单项原始 p 值很小。"],
    ["完整家族 adjusted p", "Bonferroni 和 BH 均约 0.0000598", "在全部 7,129 项的家族下仍通过 0.05 阈值。"],
], widths=[1.55, 2.1, 2.95], font_size=8.1)
heading(doc, "正确与不正确的解读", 2)
para(doc, "可以说：在这 38 份发现样本和当前分析方案下，这个探针的 AML 平均表达测量值低于 ALL，差异在两种校正后仍显著。")
para(doc, "不能只凭这个表就说：这个探针代表一个确定的致病基因、会造成白血病、已经能用于诊断，或有某种倍数变化。原始表达数值不是 log 尺度，均值差不能直接称为 fold change；一个微阵列探针也不保证等于一个独立基因。")
para(doc, "逐项 95% 区间通常只针对某一个事先指定的比较提供覆盖解释。这里的探针是从 7,129 项里筛出来的，逐项区间没有自动处理多重搜索和筛选，因此候选效应有被高估的可能。", bold_lead="逐项 95% 区间")

# Page 8
page_break(doc)
heading(doc, "6  为什么还要做已知真值模拟", 1)
para(doc, "真实数据里不知道每个探针的真实状态，因此无法直接数出有多少假发现。模拟实验让我们自己设定答案：先明确哪些变量真的有差异，再检查方法报出的名单。模拟变量是电脑生成的教学数据，不是患者样本，也没有混入真实 CSV。")
para(doc, "课程模拟每次生成 1,000 个变量；ALL 组 27 个观测、AML 组 11 个观测；噪声标准差为 1；共重复 1,000 次并固定随机种子 42。每种方法使用同一批 p 值，公平比较。模拟的变量独立且正态、两组方差相同，所以用双侧等方差 t 检验；真实表达数据则使用 Welch 检验。")
heading(doc, "A：全部变量都没有真实差异", 2)
make_table(doc,["方法", "平均发现数 R", "FWER 估计", "FDR 估计", "本实验中发现全是假发现"], [
    ["raw", "49.944", "1.000", "1.000", "是"],
    ["Bonferroni", "0.051", "0.051", "0.051", "是"],
    ["BH", "0.052", "0.051", "0.051", "是"],
], widths=[1.2, 1.2, 1.15, 1.15, 1.9], font_size=8.4, center_cols=(1,2,3,4))
para(doc, "原始方法平均约报出 50 个显著结果，与 1,000×0.05 的粗略预期一致。Bonferroni 和 BH 在全零模拟中的错误比例约 5%。模拟只重复了 1,000 次，0.051 是这批重复的估计，不是理论真值必须精确等于 0.051。")
heading(doc, "B：前 100 个变量有真实信号", 2)
make_table(doc,["真实效应 δ", "方法", "平均发现 R", "平均假发现 V", "FDR 估计", "平均检出力"], [
    ["1.0", "raw", "122.523", "44.799", "0.364", "0.777"],
    ["1.0", "Bonferroni", "6.001", "0.044", "0.007", "0.060"],
    ["1.0", "BH", "30.094", "1.445", "0.046", "0.286"],
    ["0.5", "raw", "72.361", "44.799", "0.619", "0.276"],
    ["0.5", "Bonferroni", "0.270", "0.044", "0.039", "0.002"],
    ["0.5", "BH", "0.407", "0.072", "0.045", "0.003"],
], widths=[0.85, 1.2, 1.05, 1.15, 1.0, 1.35], font_size=7.7, center_cols=(0,1,2,3,4,5))
para(doc, "信号从 δ=1.0 变为 δ=0.5 后，校正方法的平均检出力明显降低。真信号依然存在，只是较难从随机噪声中识别。Bonferroni 的错误约束更严，常常漏掉一些真信号；BH 通常保留更多候选，但要接受 FDR 目标所代表的集合层面错误容许。")

# Page 9
page_break(doc)
heading(doc, "7  预先锁定候选，再查看留出样本", 1)
heading(doc, "为什么不能只校正后来挑出的前 20 个", 2)
para(doc, "研究者查看同一批 7,129 个 p 值后挑出最小的 20 个，再把家族大小改写成 20，会忽略“先从 7,129 个里选中这 20 个”的机会。被挑出来的 p 值天生更小。只有当这 20 个探针在看到当前数据之前，已经依据独立信息预先指定，且研究确实只检验它们时，20 才可能是合适的家族大小。")
heading(doc, "课程使用的候选锁定规则", 2)
numbered(doc, "只在 train 的完整 7,129 个探针中执行发现分析。")
numbered(doc, "从 BH adjusted p≤0.05 的项目中按 raw p 从小到大排序；同值时按探针名称排序。")
numbered(doc, "取前 10 个候选（如果不足 10 个则全部保留），保存候选名和发现方向。")
numbered(doc, "锁定后才查看 test。对锁定候选做 Bonferroni 校正，并要求验证方向与发现方向一致。")
make_table(doc,["留出部分", "样本量", "检查结果"], [
    ["test 中的 ALL", "20", "只检验已经锁定的候选。"],
    ["test 中的 AML", "14", "验证家族最多 10 个，故 Bonferroni raw p 门槛为 0.05/10=0.005。"],
    ["预先锁定候选", "10", "其中 7 个方向相同且通过候选家族 Bonferroni；另 3 个未通过显著性规则。"],
], widths=[1.6, 1.2, 3.8], font_size=8.5, center_cols=(1,))
heading(doc, "这一步增加了什么证据，还缺什么", 2)
para(doc, "在留出样本中方向一致并通过预先规定的阈值，会增加这些候选可重复的证据。3 个未通过阈值的候选并不能据此被证明完全没有差异。由于 test 是同一历史数据集预留的部分，而非新招募的独立外部队列，这次练习不能称为外部临床验证，也不能证明因果机制或诊断价值。")
formula(doc, "可接受的结论：10 个预定候选中有 7 个通过留出检查，结果仍属探索性证据。")

# Page 10
page_break(doc)
heading(doc, "8  Python 代码逐段读", 1)
para(doc, "你不需要一次读懂整本 Notebook。先把代码看成一串明确的小步骤：读入表格、分组、对每个探针做检验、校正 p 值、数出通过门槛的探针。下面的小片段对应课程 Notebook 的主要逻辑。")
heading(doc, "第 1 步：载入需要的工具", 2)
code(doc, "import numpy as np\nimport pandas as pd\nfrom scipy import stats")
para(doc, "NumPy 负责数组和数值运算；pandas 负责读取和处理表格；SciPy 提供统计检验函数。import 只是把这些已经安装的工具载入当前程序。")
heading(doc, "第 2 步：读 CSV，挑出表达值", 2)
code(doc, 'df = pd.read_csv("data/leukemia.csv")\nmeta = ["sample_id", "original_split", "diagnosis"]\nprobe_columns = [c for c in df.columns if c not in meta]')
para(doc, "CSV 可以理解为普通的逗号分隔表格。`df` 是整张数据表；`meta` 列保存样本编号、原始分组和诊断类型；剩下的列才是 7,129 个探针的数值。代码通过列名排除元信息，不把诊断标签错误地当成探针。")
heading(doc, "第 3 步：把 AML 和 ALL 的数据分开", 2)
code(doc, 'train = df[df["original_split"] == "train"]\naml = train.loc[train["diagnosis"] == "AML", probe_columns]\nall_ = train.loc[train["diagnosis"] == "ALL", probe_columns]')
para(doc, "方括号内的条件是在筛选行。先只拿 train，再按 diagnosis 分组。`aml` 与 `all_` 每一行是一个样本、每一列是同一个探针；下划线只是为了避免变量名与 Python 关键字 all 冲突。")

# Page 11
page_break(doc)
heading(doc, "8  Python 代码逐段读（续）", 1)
heading(doc, "第 4 步：一次比较全部探针", 2)
code(doc, 'test = stats.ttest_ind(\n    aml.to_numpy(), all_.to_numpy(),\n    axis=0, equal_var=False\n)\np = test.pvalue')
para(doc, "`ttest_ind` 比较两组独立样本的均值。`axis=0` 表示沿着样本行计算，所以每个探针列会得到一个 p 值；`equal_var=False` 采用 Welch t 检验，不要求两组的方差相等。`p` 是长度为 7,129 的数组，每个位置对应一个探针。")
heading(doc, "第 5 步：Bonferroni 调整", 2)
code(doc, 'm = len(p)\nbonferroni_p = np.minimum(m * p, 1.0)')
para(doc, "`len(p)` 得到检验次数 m。把每个原始 p 值乘以 7,129，超过 1 的改成 1。随后数一数 adjusted p≤0.05 的有多少项。`np.minimum` 是逐个数值取较小值。")
heading(doc, "第 6 步：BH 排序规则的精简写法", 2)
code(doc, 'order = np.argsort(p)\np_sorted = p[order]\nranks = np.arange(1, m + 1)\npassed = np.where(p_sorted <= 0.05 * ranks / m)[0]\nk = passed[-1] + 1 if len(passed) else 0\ncutoff = p_sorted[k - 1] if k else -1\ndiscoveries = p <= cutoff')
para(doc, "`argsort` 给出从小到大排列时每个 p 值的原位置；`ranks` 是 1 到 m。程序找出所有满足 BH 排序门槛的位置，取最大的一个 k，再把排序前 k 个列为发现。这段是帮助理解规则的简写；课程 Notebook 还计算每个探针对应的 BH adjusted p，便于直接用 adjusted p≤0.05 作判断。")

# Page 12
page_break(doc)
heading(doc, "8  Python 代码逐段读（续）", 1)
heading(doc, "第 7 步：数显著项目", 2)
code(doc, 'raw_count = int((p <= 0.05).sum())\nbon_count = int((bonferroni_p <= 0.05).sum())\nbh_count = int((bh_adjusted_p <= 0.05).sum())')
para(doc, "`p <= 0.05` 会逐项产生 True 或 False；`.sum()` 把 True 当作 1、False 当作 0 来数数量；`int` 把结果转成普通整数。`bh_adjusted_p` 在课程 Notebook 中由校正函数算出。")
heading(doc, "模拟程序中这几个词的含义", 2)
make_table(doc,["代码或符号", "简单解释"], [
    ["rng = np.random.default_rng(42)", "创建可重复的随机数生成器。相同设置与软件环境下，便于复现实验。"],
    ["delta", "模拟中给真信号增加的均值差；δ 越小，信号相对噪声越难识别。"],
    ["truth", "电脑事先知道哪些变量是真信号；真实 CSV 没有这样的真值标签。"],
    ["R、V、S", "分别是总发现数、假发现数、真发现数；每次 R=V+S。"],
    ["FDP 与 any_false", "每次的错误比例 V/max(R,1)，以及本次是否至少错一次。"],
    ["power", "本次找回的真信号比例；全零时没有真信号，power 应记为 N/A。"],
], widths=[2.1, 4.5], font_size=8.5)
heading(doc, "初学者运行建议", 2)
bullet(doc, "先安装 requirements.txt 中列出的库，再在 VS Code 中打开 Jupyter Notebook。")
bullet(doc, "按 Notebook 顺序运行，先看预测和分析协议，再运行发现分析。")
bullet(doc, "不要根据 test 结果换候选；需要重跑时，用新的运行编号并记录种子与软件版本。")
bullet(doc, "出错时先看 Python 报错最后一行，并确认当前工作目录和数据文件路径。")
heading(doc, "路径报错时先检查什么", 2)
para(doc, "课程 Notebook 默认会从上级课程仓库查找数据。如果你直接在当前作业文件夹启动它，原来的路径设置可能找不到文件；这是文件夹位置问题，不是统计计算错了。先确认 `data/leukemia.csv` 是否在当前项目目录，再把 Notebook 顶部路径设置指向这个目录。")
code(doc, 'from pathlib import Path\nPROJECT = Path.cwd().resolve()\nif not (PROJECT / "data/leukemia.csv").exists():\n    raise FileNotFoundError("请从作业项目根目录启动 Notebook")\nREPO = PROJECT\nROOT = PROJECT')

# Page 13
page_break(doc)
heading(doc, "9  可以直接参考的报告结论", 1)
para(doc, "在 Golub 白血病表达数据的 38 份发现样本中，我们比较了 AML 与 ALL 的平均探针表达，对全部 7,129 个探针分别进行双侧 Welch t 检验。原始 p≤0.05 的探针为 1,694 个；Bonferroni adjusted p≤0.05 的探针为 26 个；BH adjusted p≤0.05 的探针为 535 个。两种校正方法控制不同层面的错误，因此发现数量不同。")
para(doc, "按预先规定的排序规则，从发现样本中锁定 10 个 BH 显著候选。在历史留出样本中，有 7 个候选方向与发现样本一致，并通过候选家族 Bonferroni 校正。该留出部分属于同一数据来源的历史样本，不是新开展的外部队列。结果支持把部分探针作为后续探索候选，但不能证明疾病因果机制、独立基因数量、患者诊断效能或临床应用价值。进一步研究仍需核对探针注释、检查样本和批次质量，并在独立研究中验证。")
heading(doc, "交作业前的五项自查", 2)
make_table(doc,["检查项", "完成了吗？"], [
    ["写明实际测试的家族大小、样本分组和单项检验方法。", "□"],
    ["区分原始 p、adjusted p、FDP、FDR 与 FWER。", "□"],
    ["没有把 BH adjusted p 写成 Storey q-value 或单项为真的概率。", "□"],
    ["报告效应方向、大小和区间，并说明区间是逐项还是同时覆盖。", "□"],
    ["结论说明留出样本与外部验证的区别，也没有越界声称因果或临床价值。", "□"],
], widths=[5.65, 0.95], font_size=8.8, center_cols=(1,))
heading(doc, "术语小抄", 2)
para(doc, "p-value：在零假设和模型成立时，观察到当前或更极端数据的概率。 adjusted p-value：在某个指定的多重检验方法和完整家族下调整后的显著性量。 q-value：FDR 框架中的相关量，具体定义与软件实现有关，本课没有计算 Storey q-value。检验家族：需要一起考虑错误控制的一组假设。效应量：差异大小的描述，不等同于显著性。", size=9.3)

# Page 14
page_break(doc)
heading(doc, "参考资料与数据来源", 1)
para(doc, "本报告中的分析设置、教学模拟参数和结果以课程项目的学生 Notebook、工作单、实验设计与数据说明为准。外部文献用于核对统计方法和原始数据引用。")
numbered(doc, "Benjamini Y, Hochberg Y. Controlling the False Discovery Rate: A Practical and Powerful Approach to Multiple Testing. Journal of the Royal Statistical Society: Series B. 1995;57(1):289–300. DOI: 10.1111/j.2517-6161.1995.tb02031.x")
para(doc, "论文说明了 FDR 这一集合层面的错误率，并提出 Benjamini–Hochberg 方法。作者期刊页面：https://rss.onlinelibrary.wiley.com/doi/abs/10.1111/j.2517-6161.1995.tb02031.x", size=9.1, color=MUTED)
numbered(doc, "Storey JD. A direct approach to false discovery rates. Journal of the Royal Statistical Society: Series B. 2002;64(3):479–498. DOI: 10.1111/1467-9868.00346")
para(doc, "用于区分 FDR 与 q-value 相关定义的背景阅读。作者期刊页面：https://academic.oup.com/jrsssb/article/64/3/479/7098513", size=9.1, color=MUTED)
numbered(doc, "Golub TR, et al. Molecular Classification of Cancer: Class Discovery and Class Prediction by Gene Expression Monitoring. Science. 1999;286(5439):531–537. DOI: 10.1126/science.286.5439.531")
para(doc, "数据集页面将这篇研究列为 leukemia 数据集引用来源：https://pubmed.ncbi.nlm.nih.gov/10521349/", size=9.1, color=MUTED)
numbered(doc, "OpenML dataset 1104: leukemia, version 1. 页面记录 72 个样本、7,129 个数值特征，并提供原研究引用信息。https://www.openml.org/search?id=1104&sort=runs&type=data")
numbered(doc, "SciPy documentation. scipy.stats.false_discovery_control：介绍 BH 与 BY 多重检验校正函数。https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.false_discovery_control.html")
heading(doc, "写作边界", 2)
para(doc, "本报告面向初学者解释课程作业，不构成临床诊断或生物医学结论。统计显著只是在设定模型、数据质量和分析协议下的一类证据；最终科学结论仍依赖领域知识和预先设计的后续研究。")

doc.core_properties.title = "多重检验校正 怎样从显著结果走向可信发现"
doc.core_properties.subject = "课程作业报告与初学者统计和代码说明"
doc.core_properties.author = "课程作业学习报告"
doc.core_properties.keywords = "多重检验, Bonferroni, BH, FDR, 白血病表达数据"
doc.save(OUT)
print(OUT)
