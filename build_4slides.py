"""Dựng CTD_Pitchbook_4slides.pptx — layout kiểu research-slide, palette lấy từ CTD_Pitchbook."""
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_PATTERN, MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree

# ---- palette (trích từ CTD_Pitchbook_2trang.pptx) ----
RED = RGBColor(0xC8, 0x10, 0x2E)
DRED = RGBColor(0x8A, 0x1C, 0x2B)
ROSE = RGBColor(0xE0, 0x7A, 0x86)
PINK = RGBColor(0xF5, 0xC3, 0xCA)
INK = RGBColor(0x2B, 0x2B, 0x2E)
CHAR = RGBColor(0x3A, 0x3A, 0x3E)
GRAY = RGBColor(0x8C, 0x8C, 0x91)
LINE = RGBColor(0xE3, 0xE0, 0xE0)
BOX = RGBColor(0xF6, 0xF3, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


# ---------------- helpers ----------------
def add_runs(p, text, size, color=INK, bold=False, italic=False):
    """**bold** = đậm, {{x}} = đậm đỏ."""
    for part in re.split(r"(\*\*.+?\*\*|\{\{.+?\}\})", text):
        if not part:
            continue
        b, c = bold, color
        if part.startswith("**"):
            part, b = part[2:-2], True
        elif part.startswith("{{"):
            part, b, c = part[2:-2], True, RED
        r = p.add_run()
        r.text = part
        f = r.font
        f.name, f.size, f.bold, f.italic = FONT, Pt(size), b, italic
        f.color.rgb = c


def tb(slide, x, y, w, h, text, size=9, color=INK, bold=False, italic=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=0, line=1.0):
    s = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = text if isinstance(text, list) else [text]
    for i, t in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(spacing)
        p.line_spacing = line
        add_runs(p, t, size, color, bold, italic)
    return s


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, dash=False):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(lw)
        if dash:
            s.line.dash_style = MSO_LINE.DASH
    s.shadow.inherit = False
    return s


def hline(slide, x, y, w, color=GRAY, lw=0.75):
    c = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c.line.color.rgb = color; c.line.width = Pt(lw)
    return c


def vline(slide, x, y, h, color=GRAY, lw=1.0, dash=False):
    c = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x), Inches(y + h))
    c.line.color.rgb = color; c.line.width = Pt(lw)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    return c


def header(slide, chapter, title, page):
    rect(slide, 0.3, 0.22, 0.07, 0.72, fill=RED)
    tb(slide, 0.5, 0.2, 9, 0.28, chapter, 13, GRAY, bold=True)
    tb(slide, 0.5, 0.45, 11.5, 0.55, title, 26, INK, bold=True)
    tb(slide, 10.6, 0.24, 2.43, 0.2, "GO FINANCE 2026 · PITCHBOOK", 8, RED, bold=True, align=PP_ALIGN.RIGHT)
    tb(slide, 10.6, 0.42, 2.43, 0.2, "Số liệu tính đến 1/10/2026", 8, GRAY, align=PP_ALIGN.RIGHT)
    # footer
    tb(slide, 0.3, 7.2, 11.6, 0.2,
       "(F)/(E) = dự phóng của nhóm, không phải số công bố. Không phải khuyến nghị đầu tư cá nhân.", 7, GRAY, italic=True)
    tb(slide, 12.4, 7.2, 0.63, 0.2, f"|  {page}", 8, GRAY, align=PP_ALIGN.RIGHT)


def summary_box(slide, text, y=1.07, h=0.8):
    rect(slide, 0.3, y, 12.73, h, fill=BOX)
    hline(slide, 0.3, y + h, 12.73, color=RED, lw=1.25)
    tb(slide, 0.42, y + 0.07, 12.5, h - 0.1, text, 10, INK, anchor=MSO_ANCHOR.MIDDLE, line=1.05)


def panel_title(slide, x, y, w, text, num=None):
    t = f"[{num}] {text}" if num else text
    tb(slide, x, y, w, 0.25, t, 11, INK, bold=True)
    hline(slide, x, y + 0.27, w, color=GRAY, lw=0.75)


def source(slide, x, y, w, text):
    tb(slide, x, y, w, 0.15, text, 6.5, GRAY, italic=True)


# ---------------- chart helpers ----------------
def _font(obj, size=8, color=CHAR, bold=False):
    obj.font.size = Pt(size); obj.font.name = FONT; obj.font.bold = bold; obj.font.color.rgb = color


def fill_series(s, color):
    s.format.fill.solid(); s.format.fill.fore_color.rgb = color
    s.format.line.fill.background()


def hatch_point(pt, color=GRAY):
    pt.format.fill.patterned()
    pt.format.fill.pattern = MSO_PATTERN.WIDE_UPWARD_DIAGONAL
    pt.format.fill.fore_color.rgb = color
    pt.format.fill.back_color.rgb = LINE
    pt.format.line.color.rgb = color
    pt.format.line.width = Pt(0.75)
    pt.format.line.dash_style = MSO_LINE.DASH


def color_point(pt, color):
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = color
    pt.format.line.fill.background()


def base_chart(slide, ctype, data, x, y, w, h, legend=True, labels=True, numfmt='#,##0',
               gap=60, overlap=None, vmax=None, vmin=None, show_val_axis=False, lsize=7.5, legend_pos=XL_LEGEND_POSITION.TOP):
    gf = slide.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), data)
    ch = gf.chart
    ch.font.size = Pt(8); ch.font.name = FONT; ch.font.color.rgb = CHAR
    ch.has_title = False
    ch.has_legend = legend
    if legend:
        ch.legend.position = legend_pos
        ch.legend.include_in_layout = False
        _font(ch.legend, 8)
    plot = ch.plots[0]
    if hasattr(plot, "gap_width"):
        plot.gap_width = gap
    if overlap is not None and hasattr(plot, "overlap"):
        plot.overlap = overlap
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = numfmt; dl.number_format_is_linked = False
        _font(dl, lsize)
    ca, va = ch.category_axis, ch.value_axis
    ca.format.line.color.rgb = GRAY
    ca.has_major_gridlines = False
    ca.tick_labels.font.size = Pt(8)
    ca.major_tick_mark = 2  # none
    va.has_major_gridlines = False
    va.visible = show_val_axis
    if show_val_axis:
        va.tick_labels.font.size = Pt(7.5)
        va.format.line.fill.background()
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = LINE
    if vmax is not None:
        va.maximum_scale = vmax
    if vmin is not None:
        va.minimum_scale = vmin
    return ch


def set_plot_layout(chart, x, y, w, h):
    """inner plot area theo tỉ lệ để overlay đường tham chiếu chính xác."""
    pa = chart._chartSpace.chart.plotArea
    old = pa.find(qn("c:layout"))
    if old is not None:
        pa.remove(old)
    lay = etree.SubElement(pa, qn("c:layout"))
    pa.insert(0, lay)
    ml = etree.SubElement(lay, qn("c:manualLayout"))
    for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                     ("x", x), ("y", y), ("w", w), ("h", h)):
        e = etree.SubElement(ml, qn("c:" + tag)); e.set("val", str(val))


def point_labels_off(series, idxs):
    for i in idxs:
        dl = series.points[i].data_label
        dl.has_text_frame = True
        dl.text_frame.text = ""


# ---------------- table helper ----------------
def cell_border(cell, bottom=LINE, w=0.5, top=None):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)

    def ln(tag, color, width):
        e = etree.SubElement(tcPr, qn(tag))
        if color is None:
            e.set("w", "0"); etree.SubElement(e, qn("a:noFill"))
        else:
            e.set("w", str(int(width * 12700)))
            sf = etree.SubElement(e, qn("a:solidFill"))
            c = etree.SubElement(sf, qn("a:srgbClr")); c.set("val", str(color))
    ln("a:lnL", None, 0); ln("a:lnR", None, 0)
    ln("a:lnT", top, w); ln("a:lnB", bottom, w)
    # schema order: borders must precede fill
    for e in list(tcPr):
        if e.tag in (qn("a:solidFill"), qn("a:noFill")):
            tcPr.remove(e); tcPr.append(e)


def table(slide, x, y, w, rows, colw, rowh=0.22, size=8, header_fill=INK, header_color=WHITE,
          align=None, bold_rows=(), fills=None, first_col_bold=False, header=True):
    nr, nc = len(rows), len(rows[0])
    gs = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(rowh * nr))
    t = gs.table
    tblPr = t._tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
    tot = sum(colw)
    for i, cw in enumerate(colw):
        t.columns[i].width = Inches(w * cw / tot)
    for r in range(nr):
        t.rows[r].height = Inches(rowh)
        for c in range(nc):
            cell = t.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.05)
            cell.margin_top = cell.margin_bottom = Inches(0.015)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (align[c] if align else PP_ALIGN.LEFT)
            is_h = header and r == 0
            col = header_color if is_h else INK
            add_runs(p, str(rows[r][c]), size, col,
                     bold=is_h or r in bold_rows or (first_col_bold and c == 0))
            f = fills.get((r, c)) if fills else None
            if is_h:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
            elif f is not None:
                cell.fill.solid(); cell.fill.fore_color.rgb = f
            else:
                cell.fill.background()
            cell_border(cell, bottom=(None if is_h else LINE))
    return t


L, R, C = PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER

# =====================================================================
# SLIDE 1 — NGÀNH
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, "CHƯƠNG 1 – ĐẦU TƯ CÔNG & VLXD", "Hai dòng cầu đến cùng lúc, vật liệu thiếu cung", 1)
summary_box(s,
    "FDI và đầu tư công cùng ở **đầu chu kỳ**: [1] FDI giải ngân kỷ lục **27.6 tỷ USD** (2025), vốn đăng ký 8T/2026 "
    "(**40.6 tỷ USD**) đã vượt cả năm 2025; [2] đầu tư công mới giải ngân **55.5%** kế hoạch 2026, kế hoạch 2027-30 "
    "bình quân ~1,800 nghìn tỷ/năm (**+76%**); [3] vật liệu thiếu cung, giá đầu vào tăng **20-32%**. "
    "Hưởng lợi rõ nhất: {{CTD}} (nhà thầu FDI, dự án phức tạp) và {{HPG}} (thép đầu ngành).")

CW, XS = 4.15, [0.3, 4.58, 8.88]
TY, BY = 2.0, 4.65

# T1 — FDI
panel_title(s, XS[0], TY, CW, "FDI giải ngân kỷ lục, còn tăng ~10%/năm (tỷ USD)")
cd = CategoryChartData()
cd.categories = ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026F", "2027F"]
cd.add_series("Vốn đăng ký", [35.5, 38.0, 28.5, 31.2, 27.7, 36.6, 38.2, 38.4, None, None])
cd.add_series("Giải ngân", [19.0, 20.4, 20.0, 19.7, 22.4, 23.2, 25.4, 27.6, 30.9, 33.4])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, XS[0], TY + 0.32, CW, 2.05, numfmt='0.0', gap=40, overlap=-5, lsize=6.5)
fill_series(ch.plots[0].series[0], PINK); fill_series(ch.plots[0].series[1], RED)
for i in (8, 9):
    hatch_point(ch.plots[0].series[1].points[i], CHAR)
source(s, XS[0], TY + 2.4, CW, "Nguồn: TCTK, Bộ KH&ĐT; 2026F-27F ước tính nhóm. CAGR giải ngân 25-27F: +10%.")

# T2 — Đầu tư công
panel_title(s, XS[1], TY, CW, "Đầu tư công 2027-30F: +76% (nghìn tỷ đồng)")
cd = CategoryChartData()
cd.categories = ["2024", "2025", "2026*", "2026F", "2027-30F"]
cd.add_series("Kế hoạch", [683, 906, 1023, 1023, 1800])
cd.add_series("Giải ngân", [549, 755, 568, 856, None])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, XS[1], TY + 0.32, CW, 2.05, gap=50, overlap=-5, lsize=7)
fill_series(ch.plots[0].series[0], PINK); fill_series(ch.plots[0].series[1], RED)
hatch_point(ch.plots[0].series[1].points[3], CHAR)
color_point(ch.plots[0].series[0].points[4], CHAR)
source(s, XS[1], TY + 2.4, CW, "Nguồn: Bộ Tài chính (1/10/2026). *Đến 24/9/2026 (55.5% KH); 2026F = kịch bản cơ sở 83.7%.")

# T3 — Giá vật liệu
panel_title(s, XS[2], TY, CW, "Giá đầu vào xây dựng tăng 20-32% (%)")
cd = CategoryChartData()
items = [("Nhựa đường (Q1/26)", 32.0), ("HRC (đỉnh 5/26)", 28.4), ("Cát, đá (Q1/26)", 23.3),
         ("Dầu diesel (8T/26)", 22.2), ("Thép CB300 D10", 20.1), ("Xi măng (Q1/26)", 7.0)]
cd.categories = [i[0] for i in items][::-1]
cd.add_series("Tăng giá", [i[1] for i in items][::-1])
ch = base_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, cd, XS[2], TY + 0.32, CW, 2.05, legend=False, numfmt='"+"0.0"%"', gap=35, vmax=40, vmin=0)
fill_series(ch.plots[0].series[0], ROSE)
cols = [PINK, ROSE, ROSE, RED, RED, DRED]
for i, c in enumerate(cols):
    color_point(ch.plots[0].series[0].points[i], c)
source(s, XS[2], TY + 2.4, CW, "Nguồn: Bộ Xây dựng, VSA, TCTK. Vật liệu ~70% giá vốn nhà thầu; hợp đồng giá cố định hết hạn 2027.")

# B1 — Tỷ lệ đăng ký / giải ngân
panel_title(s, XS[0], BY, CW, "Vốn đăng ký (t−1) luôn vượt giải ngân (t) (lần)")
cd = CategoryChartData()
cd.categories = ["2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026F"]
cd.add_series("Đăng ký/giải ngân", [1.74, 1.90, 1.45, 1.39, 1.19, 1.44, 1.38, 1.24])
ch = base_chart(s, XL_CHART_TYPE.LINE_MARKERS, cd, XS[0], BY + 0.32, CW, 1.95, legend=False, numfmt='0.00"x"', vmin=0.8, vmax=2.1)
ser = ch.plots[0].series[0]
ser.format.line.color.rgb = RED; ser.format.line.width = Pt(2.25); ser.smooth = False
ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 7
ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = WHITE
ser.marker.format.line.color.rgb = RED
ch.plots[0].data_labels.position = XL_LABEL_POSITION.ABOVE
source(s, XS[0], BY + 2.3, CW, "Nguồn: tính từ TCTK/Bộ KH&ĐT. Cầu 2026-27 phần lớn đã nằm trong vốn đăng ký.")

# B2 — Tiến độ
panel_title(s, XS[1], BY, CW, "Dự án trọng điểm: còn nhiều khối lượng (%)")
cd = CategoryChartData()
prog = [("Long Thành GĐ1", 76), ("Vành đai 4 Hà Nội", 79.4), ("Vành đai 3 TP.HCM", 75),
        ("Cao tốc toàn quốc", 66.9), ("ĐSTĐC Bắc-Nam", 0)]
cd.categories = [p[0] for p in prog][::-1]
cd.add_series("Hoàn thành/giải ngân", [p[1] for p in prog][::-1])
cd.add_series("Còn lại", [100 - p[1] for p in prog][::-1])
ch = base_chart(s, XL_CHART_TYPE.BAR_STACKED_100, cd, XS[1], BY + 0.32, CW, 1.95, legend=False, numfmt='0.0"%"', gap=40, overlap=100)
fill_series(ch.plots[0].series[0], RED)
ch.plots[0].series[0].data_labels.font.color.rgb = WHITE
ch.plots[0].series[0].data_labels.font.size = Pt(8)
ch.plots[0].series[0].data_labels.number_format = '0.0"%"'; ch.plots[0].series[0].data_labels.number_format_is_linked = False
ch.plots[0].series[0].data_labels.show_value = True
sr = ch.plots[0].series[1]
sr.format.fill.patterned(); sr.format.fill.pattern = MSO_PATTERN.WIDE_UPWARD_DIAGONAL
sr.format.fill.fore_color.rgb = GRAY; sr.format.fill.back_color.rgb = WHITE
sr.format.line.color.rgb = GRAY; sr.format.line.dash_style = MSO_LINE.DASH
sr.data_labels.show_value = False
ch.plots[0].series[0].data_labels.show_value = True
point_labels_off(ch.plots[0].series[0], [0])  # ĐSTĐC 0%
source(s, XS[1], BY + 2.3, CW, "Nguồn: Bộ Xây dựng, Bộ Tài chính (4-9/2026). Vành đai 4: tỉ lệ giải ngân. Phần gạch = khối lượng còn lại.")

# B3 — Mega project
panel_title(s, XS[2], BY, CW, "Chuyển pha mega project nên chu kỳ không đứt đoạn")
rows = [["Dự án", "Quy mô", "Mốc chính"],
        ["Long Thành GĐ1", "~16 tỷ USD", "Khai thác kỹ thuật cuối 2026"],
        ["Gia Bình GĐ1 (Bắc Ninh)", "141,000 tỷ", "Vận hành đầu 2027 (APEC)"],
        ["Metro 2 TP.HCM", "47,890 tỷ", "Khởi công 1/2026, 5-7 năm"],
        ["Cảng Cần Giờ", "4.96 tỷ USD", "GĐ1 trước 2030"],
        ["ĐSTĐC Bắc-Nam", "67.3 tỷ USD", "Khởi công 12/2027 → 2035"]]
table(s, XS[2], BY + 0.38, CW, rows, [1.55, 1.0, 1.75], rowh=0.27, size=8, align=[L, R, L], first_col_bold=True,
      header_fill=RED)
rect(s, XS[2], BY + 2.08, CW, 0.3, fill=None, line=RED, dash=True)
tb(s, XS[2] + 0.08, BY + 2.08, CW - 0.16, 0.3,
   "Phần giá trị cao nhất là **phần phức tạp**: ngầm, nhà ga, nhà để xe, siêu công trình ven biển → vùng mục tiêu của {{CTD}}.",
   7.5, INK, anchor=MSO_ANCHOR.MIDDLE)
source(s, XS[2], BY + 2.42, CW, "Nguồn: Cổng TTĐT Chính phủ, Bộ Xây dựng, UBND địa phương.")

# =====================================================================
# SLIDE 2 — SÀNG LỌC MÃ
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, "CHƯƠNG 1 – ĐẦU TƯ CÔNG & VLXD", "Sàng lọc mã: CTD là điểm sáng của chu kỳ", 2)
summary_box(s,
    "Chúng tôi chọn mã theo ba tiêu chí: [1] **hưởng lợi trực tiếp** từ hai dòng cầu FDI và đầu tư công, [2] **vị thế dẫn đầu** trong phân khúc, "
    "[3] câu chuyện đầu tư rõ ràng. Hai mã đạt chuẩn là CTD (nhà thầu) và HPG (thép). Trong 2026-27, "
    "{{CTD là điểm sáng nhất}}: FDI ~50% doanh thu, backlog 70,000 tỷ (2.04x), tiền ròng 1,880 tỷ và định giá chỉ bằng ~một nửa nhóm khu vực.")

# Valuation charts
panel_title(s, 0.3, TY, 6.1, "Định giá CTD thấp hơn nhiều so với nhóm khu vực…")
cd = CategoryChartData()
cd.categories = ["CTD", "Trung vị VN", "Mục tiêu dùng", "Trung vị khu vực"]
cd.add_series("P/E forward (x)", [7.4, 7.3, 9.5, 14.5])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, 0.3, TY + 0.58, 2.95, 1.95, legend=False, numfmt='0.0"x"', gap=45, vmin=0, vmax=18)
for i, c in enumerate([RED, PINK, ROSE, CHAR]):
    color_point(ch.plots[0].series[0].points[i], c)
tb(s, 0.3, TY + 0.34, 2.95, 0.2, "P/E forward (x)", 8.5, CHAR, bold=True, align=C)
cd = CategoryChartData()
cd.categories = ["CTD", "Trung vị VN", "Mục tiêu dùng", "Trung vị khu vực"]
cd.add_series("P/B (x)", [0.69, 0.55, 0.87, 1.61])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, 3.4, TY + 0.58, 2.95, 1.95, legend=False, numfmt='0.00"x"', gap=45, vmin=0, vmax=2.0)
for i, c in enumerate([RED, PINK, ROSE, CHAR]):
    color_point(ch.plots[0].series[0].points[i], c)
tb(s, 3.4, TY + 0.34, 2.95, 0.2, "P/B (x)", 8.5, CHAR, bold=True, align=C)
source(s, 0.3, TY + 2.58, 6.1, "Nguồn: stockanalysis.com, Simplize (17/7-2/10/2026); chưa kiểm chứng Bloomberg/FiinPro. Mục tiêu dùng = 70% Việt Nam + 30% khu vực.")

# CTD vs HPG matrix
panel_title(s, 6.7, TY, 6.33, "…và câu chuyện 2026-27 thuộc về CTD")
G = RGBColor(0xFB, 0xE6, 0xE9)
rows = [["Tiêu chí", "CTD · Nhà thầu", "HPG · Thép đầu ngành"],
        ["Vị thế", "Nhà thầu FDI và dự án phức tạp", "Thị phần thép xây dựng ~36-38%; HRC ~8.6 triệu tấn/năm"],
        ["Cầu FDI", "FDI ~50% doanh thu; 38 dự án LEED", "Gián tiếp qua cầu xây dựng công nghiệp"],
        ["Cầu đầu tư công", "Backlog 70,000 tỷ (2.04x doanh thu)", "Hưởng lợi cầu thép từ mega project"],
        ["Giá vật liệu", "Rủi ro biên: vật liệu ~70% giá vốn", "Hưởng lợi giá thép"],
        ["Phân tích trong pitchbook", "{{MUA · mục tiêu 83,500 (+41%)}}", "Giữ trong danh mục, không phân tích chi tiết"]]
fills = {(r, 1): G for r in range(1, 6)}
table(s, 6.7, TY + 0.38, 6.33, rows, [1.25, 2.4, 2.7], rowh=0.34, size=8, first_col_bold=True, fills=fills, header_fill=INK)
rect(s, 6.7, TY + 0.34, 6.33, 0.04, fill=None)
source(s, 6.7, TY + 2.5, 6.33, "Nguồn: CTD, BSC, KBSV, mô hình nhóm.")

# Bottom-left: revenue mix
panel_title(s, 0.3, BY + 0.1, 6.1, "Doanh thu quy đổi: FDI ~20,000 tỷ năm FY27E (tỷ đồng)")
cd = CategoryChartData()
cd.categories = ["FY26A", "FY27E", "FY28E", "FY29E"]
cd.add_series("FDI", [17170, 20000, 22400, 24650])
cd.add_series("Dân sinh", [15453, 18000, 20160, 22185])
cd.add_series("Khách sạn - du lịch", [1717, 2000, 2240, 2465])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_STACKED, cd, 0.3, BY + 0.7, 6.1, 1.65, numfmt='#,##0', gap=55, overlap=100, lsize=7.5, legend_pos=XL_LEGEND_POSITION.BOTTOM)
fill_series(ch.plots[0].series[0], CHAR); fill_series(ch.plots[0].series[1], ROSE); fill_series(ch.plots[0].series[2], PINK)
ch.plots[0].series[0].data_labels.font.color.rgb = WHITE
ch.plots[0].series[0].data_labels.font.size = Pt(7.5)
ch.plots[0].series[0].data_labels.number_format = '#,##0'; ch.plots[0].series[0].data_labels.number_format_is_linked = False
ch.plots[0].series[0].data_labels.show_value = True
ch.plots[0].series[2].data_labels.show_value = False
ch.plots[0].series[1].data_labels.font.size = Pt(7.5)
ch.plots[0].series[1].data_labels.number_format = '#,##0'; ch.plots[0].series[1].data_labels.number_format_is_linked = False
ch.plots[0].series[1].data_labels.show_value = True
tb(s, 0.3, BY + 0.43, 6.1, 0.22, "Tổng 34,340 → 49,300 · CAGR FY26-29: +12.8%", 8, WHITE, bold=True, align=C, anchor=MSO_ANCHOR.MIDDLE)
s.shapes[-1].fill.solid(); s.shapes[-1].fill.fore_color.rgb = CHAR
source(s, 0.3, BY + 2.38, 6.1, "Nguồn: Coteccons, mô hình nhóm; giữ cơ cấu hiện tại cho FY26A-FY29E.")

# Bottom-right: highlights
panel_title(s, 6.7, BY + 0.1, 6.33, "CTD nổi bật trên ba lợi thế cạnh tranh")
cards = [
    ("FDI & công nghiệp", [("~50%", "doanh thu từ FDI"), (">75%", "doanh thu từ khách cũ (repeat sales)")]),
    ("ESG chuẩn quốc tế", [("38", "dự án đạt chuẩn LEED"), ("76%", "điểm ESG HOSE, vượt VN100")]),
    ("Sức khoẻ tài chính", [("2.04x", "backlog / doanh thu"), ("1,880 tỷ", "tiền ròng FY26A")]),
]
cy = BY + 0.47
for name, chips in cards:
    rect(s, 6.7, cy, 6.33, 0.62, fill=WHITE, line=CHAR, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.12
    tb(s, 6.85, cy, 1.8, 0.62, name, 9.5, INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    for k, (big, small) in enumerate(chips):
        x = 8.65 + k * 2.2
        r = rect(s, x, cy + 0.07, 2.1, 0.48, fill=(RED if k == 0 else PINK))
        tb(s, x + 0.07, cy + 0.07, 0.8, 0.48, big, 12, WHITE if k == 0 else DRED, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, x + 0.82, cy + 0.07, 1.25, 0.48, small, 7.5, WHITE if k == 0 else INK, anchor=MSO_ANCHOR.MIDDLE)
    cy += 0.7
source(s, 6.7, BY + 2.57, 6.33, "Nguồn: Coteccons (ĐHCĐ, báo cáo thường niên), HOSE, mô hình nhóm.")

# =====================================================================
# SLIDE 3 — ĐỘNG LỰC CTD
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, "CÔNG TY CỔ PHẦN XÂY DỰNG COTECCONS (HOSE: CTD)", "Backlog bảo chứng doanh thu, mix phức tạp kéo biên lên", 3)
summary_box(s,
    "Cơ hội đến từ ba lớp: [1] **backlog ~70,000 tỷ** phủ 2.04x doanh thu, trong đó ~30,000 tỷ gối đầu sang FY28+; [2] chuyển trọng tâm sang "
    "**dự án phức tạp** (biên 5-6.5% so với 3-4% của nhà ở) giúp biên gộp lên 4.45%; [3] LNST FY27E **894 tỷ** (+13%). "
    "Điểm cần theo dõi: dòng tiền hoạt động còn phụ thuộc **ứng trước khách hàng** (9,029 tỷ).")

# [1] Backlog
panel_title(s, XS[0], TY, CW, "Backlog phủ 2.04x doanh thu (tỷ đồng)", 1)
cd = CategoryChartData()
cd.categories = ["FY25A\n1.42x", "FY26A\n2.04x", "FY27E\n1.88x", "FY28E\n1.75x", "FY29E\n1.60x"]
cd.add_series("Backlog", [35353, 70000, 75000, 78200, 78900])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, XS[0], TY + 0.32, CW, 2.05, legend=False, numfmt='#,##0', gap=45)
ser = ch.plots[0].series[0]
color_point(ser.points[0], RED); color_point(ser.points[1], CHAR)
for i in (2, 3, 4):
    hatch_point(ser.points[i], CHAR)
source(s, XS[0], TY + 2.4, CW, "Nguồn: CTD, mô hình nhóm (ký mới 45-50 nghìn tỷ/năm). Độ phủ vẫn trên ngưỡng cảnh báo 1.5x.")

# [2] Mix -> biên
panel_title(s, XS[1], TY, CW, "Tỉ trọng dự án phức tạp quyết định biên gộp (%)", 2)
xd = XyChartData()
a = xd.add_series("Biên gộp = 3.5% + 2.25 × tỉ trọng")
for xv in range(0, 61, 10):
    a.add_data_point(xv, round(3.5 + 2.25 * xv / 100, 3))
b = xd.add_series("Mốc")
for xv, yv in ((24, 4.0), (42, 4.45), (50, 4.6)):
    b.add_data_point(xv, yv)
ch = base_chart(s, XL_CHART_TYPE.XY_SCATTER_LINES, xd, XS[1], TY + 0.32, CW, 2.05, legend=False, labels=False,
                show_val_axis=True, vmin=3.0, vmax=5.2)
ch.value_axis.tick_labels.number_format = '0.0"%"'; ch.value_axis.tick_labels.number_format_is_linked = False
ch.category_axis.minimum_scale = 0; ch.category_axis.maximum_scale = 60
ch.category_axis.tick_labels.number_format = '0"%"'; ch.category_axis.tick_labels.number_format_is_linked = False
s0, s1 = ch.plots[0].series
s0.format.line.color.rgb = ROSE; s0.format.line.width = Pt(2.5); s0.marker.style = XL_MARKER_STYLE.NONE; s0.smooth = False
s1.format.line.fill.background(); s1.marker.style = XL_MARKER_STYLE.CIRCLE; s1.marker.size = 9
s1.marker.format.fill.solid(); s1.marker.format.fill.fore_color.rgb = RED; s1.marker.format.line.color.rgb = WHITE
for i, t in enumerate(["Hiện ~24%: ~4.0%", "Cần ~42%: 4.45%", "FDI ~50%: ~4.6%"]):
    dl = s1.points[i].data_label
    dl.position = XL_LABEL_POSITION.BELOW if i == 0 else (XL_LABEL_POSITION.BELOW if i == 1 else XL_LABEL_POSITION.ABOVE)
    dl.text_frame.text = t
    dl.text_frame.paragraphs[0].runs[0].font.size = Pt(8)
    dl.text_frame.paragraphs[0].runs[0].font.bold = True
    dl.text_frame.paragraphs[0].runs[0].font.color.rgb = DRED
source(s, XS[1], TY + 2.4, CW, "Nguồn: Kirin: nhà ở 3-4%, công nghiệp và hạ tầng 5-6.5%; tính nhóm. Trục x = tỉ trọng dự án phức tạp trong doanh thu.")

# [3] Cluster table
panel_title(s, XS[2], TY, CW, "Vùng mục tiêu: các hạng mục kỹ thuật phức tạp", 3)
rows = [["Cụm dự án", "Cấu phần phức tạp"],
        ["Đường sắt đô thị, kết cấu ngầm", "Tường vây sâu, cọc barrette, nhà ga ngầm, xử lý nền yếu"],
        ["Hàng không, ga trung chuyển", "Nhà để xe >200,000 m² sàn, thép khẩu độ lớn, MEP phức tạp"],
        ["Siêu công trình, hạ tầng đặc thù", "Hội nghị quốc tế, cọc khoan nhồi cảng nước sâu, depot kỹ thuật"]]
table(s, XS[2], TY + 0.38, CW, rows, [1.35, 2.5], rowh=0.5, size=8, first_col_bold=True, header_fill=RED)
source(s, XS[2], TY + 2.5, CW, "Nguồn: Bộ Xây dựng, UBND địa phương; tổng hợp nhóm. Biên mục tiêu 5-6.5%.")

# [4] LNST scenarios
panel_title(s, XS[0], BY, CW, "LNST ba kịch bản (tỷ đồng); kế hoạch FY27: 1,050", 4)
cd = CategoryChartData()
cd.categories = ["FY27E", "FY28E", "FY29E"]
cd.add_series("Thấp", [666, 663, 704])
cd.add_series("Cơ sở", [894, 1010, 1129])
cd.add_series("Cao", [1046, 1302, 1565])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, XS[0], BY + 0.32, CW, 1.95, numfmt='#,##0', gap=50, overlap=-5)
fill_series(ch.plots[0].series[0], PINK); fill_series(ch.plots[0].series[1], CHAR); fill_series(ch.plots[0].series[2], RED)
source(s, XS[0], BY + 2.3, CW, "Nguồn: mô hình nhóm. Biên gộp thấp 4.1-4.2%, cơ sở 4.45%, cao 4.6-5.0%.")

# [5] OCF bridge
panel_title(s, XS[1], BY, CW, "OCF FY26 dương nhờ ứng trước (tỷ đồng)", 5)
cd = CategoryChartData()
cd.categories = ["LN trước\nvốn LĐ", "Δ Phải\nthu", "Δ Tồn\nkho", "Δ Phải trả\n+ ứng trước", "Lãi, thuế,\nkhác", "OCF\nFY26"]
cd.add_series("OCF", [1210, -1475, -3884, 5465, -501, 816])
ch = base_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, XS[1], BY + 0.32, CW, 1.95, legend=False, numfmt='#,##0', gap=40)
ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
ser = ch.plots[0].series[0]
ser.invert_if_negative = False
for i, c in enumerate([CHAR, RED, RED, ROSE, RED, CHAR]):
    color_point(ser.points[i], c)
source(s, XS[1], BY + 2.3, CW, "Nguồn: BCTC CTD (Vietcap), nhóm cộng 4 quý. Phải thu 15,381 tỷ (DSO ~176 ngày), tồn kho 9,906 tỷ.")

# [6] LNST bridge (waterfall)
panel_title(s, XS[2], BY, CW, "Cầu nối LNST FY26 → FY27E (tỷ đồng)", 6)
cats = ["LNST\nFY26", "Khối\nlượng", "Biên\ngộp", "SG&A", "Khác", "Tài chính\nròng", "Không\nlặp lại", "LNST\nFY27E"]
steps = [788, 184, 74, -48, 2, 1, -107]
base, up, down, tot = [], [], [], []
run = 0
for i, v in enumerate(steps):
    if i == 0:
        base.append(0); up.append(0); down.append(0); tot.append(v); run = v
    elif v >= 0:
        base.append(run); up.append(v); down.append(0); tot.append(0); run += v
    else:
        run += v; base.append(run); up.append(0); down.append(-v); tot.append(0)
base.append(0); up.append(0); down.append(0); tot.append(894)
cd = CategoryChartData()
cd.categories = cats
cd.add_series("base", base); cd.add_series("Tăng", up); cd.add_series("Giảm", down); cd.add_series("Tổng", tot)
ch = base_chart(s, XL_CHART_TYPE.COLUMN_STACKED, cd, XS[2], BY + 0.32, CW, 1.95, legend=False, labels=False, gap=35, overlap=100,
                vmin=600, vmax=1100)
sr = ch.plots[0].series
sr[0].format.fill.background(); sr[0].format.line.fill.background()
fill_series(sr[1], ROSE); fill_series(sr[2], RED); fill_series(sr[3], CHAR)
for ser, fmt in ((sr[1], '"+"#,##0;;;'), (sr[2], '"−"#,##0;;;'), (sr[3], '#,##0;;;')):
    ser.data_labels.show_value = True
    ser.data_labels.number_format = fmt; ser.data_labels.number_format_is_linked = False
    ser.data_labels.font.size = Pt(7.5); ser.data_labels.font.name = FONT
    ser.data_labels.position = XL_LABEL_POSITION.INSIDE_END
sr[3].data_labels.font.color.rgb = WHITE
sr[2].data_labels.font.color.rgb = WHITE
source(s, XS[2], BY + 2.3, CW, "Nguồn: mô hình nhóm (biên gộp 4.45%). Trục bắt đầu từ 600 để thấy rõ các bước.")

# =====================================================================
# SLIDE 4 — TÓM TẮT / ĐỊNH GIÁ / RỦI RO
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, "CÔNG TY CỔ PHẦN XÂY DỰNG COTECCONS (HOSE: CTD)", "MUA CTD: mục tiêu 83,500 đồng, upside +41%", 4)

# Football field
FX, FY, FW, FH = 0.3, 1.1, 5.0, 2.85
tb(s, FX, FY, FW, 0.22, "Football field: giá mục tiêu nằm giữa các vùng định giá (nghìn đồng/cp)", 9, INK, bold=True)
hline(s, FX, FY + 0.25, FW, color=GRAY)
ff = [("P/E 8-12x × EPS FY27E", 64.0, 96.0), ("P/B 0.8-1.2x × BVPS 6/2027E", 74.7, 112.1),
      ("P/E 8-12x × EPS FY28E, chiết khấu 14%", 63.4, 95.1), ("Nhóm so sánh (phân vị 25-75)", 67.0, 82.4),
      ("DCF: nhân EBITDA - Gordon", 82.4, 100.1), ("DCF stress: 50% ứng trước là nợ", 42.0, 59.7),
      ("Kịch bản thấp - cao", 64.3, 106.0)]
ff_r = ff[::-1]
cd = CategoryChartData()
cd.categories = [f[0] for f in ff_r]
cd.add_series("base", [f[1] for f in ff_r])
cd.add_series("range", [round(f[2] - f[1], 1) for f in ff_r])
CY0, CH0 = FY + 0.55, FH - 0.55
ch = base_chart(s, XL_CHART_TYPE.BAR_STACKED, cd, FX, CY0, FW, CH0, legend=False, labels=False, gap=45, overlap=100, vmin=30, vmax=120)
ch.value_axis.visible = True
ch.value_axis.tick_labels.font.size = Pt(7.5)
ch.value_axis.major_unit = 10
ch.value_axis.format.line.fill.background()
ch.value_axis.has_major_gridlines = True
ch.value_axis.major_gridlines.format.line.color.rgb = LINE
ch.category_axis.tick_labels.font.size = Pt(7.5)
LX, LY, LW, LH = 0.40, 0.02, 0.57, 0.86
set_plot_layout(ch, LX, LY, LW, LH)
sr = ch.plots[0].series
sr[0].format.fill.background(); sr[0].format.line.fill.background()
fill_series(sr[1], ROSE)
rng_colors = {"DCF: nhân EBITDA - Gordon": CHAR, "DCF stress: 50% ứng trước là nợ": DRED,
              "Kịch bản thấp - cao": RED, "Nhóm so sánh (phân vị 25-75)": PINK, "P/E 8-12x × EPS FY28E, chiết khấu 14%": PINK}
for i, f in enumerate(ff_r):
    color_point(sr[1].points[i], rng_colors.get(f[0], ROSE))
    dl = sr[1].points[i].data_label
    dl.has_text_frame = True
    dl.text_frame.text = f"{f[1]:.1f} – {f[2]:.1f}"
    r0 = dl.text_frame.paragraphs[0].runs[0]; r0.font.size = Pt(7); r0.font.bold = True
    r0.font.color.rgb = WHITE if rng_colors.get(f[0]) in (CHAR, DRED, RED, None) or f[0].startswith("P/B") or f[0].startswith("P/E 8-12x × EPS FY27E") else INK
    dl.position = XL_LABEL_POSITION.CENTER
# overlay price/target lines
px0 = FX + FW * LX; pw = FW * LW
py0 = CY0 + CH0 * LY; ph = CH0 * LH
for val, lab, col in ((59.3, "Giá 59.3", GRAY), (83.5, "Mục tiêu 83.5", INK)):
    xx = px0 + pw * (val - 30) / 90
    vline(s, xx, py0 - 0.05, ph + 0.05, color=col, lw=1.25, dash=(val == 59.3))
    lbl = rect(s, xx - 0.4, FY + 0.3, 0.8, 0.18, fill=col)
    tb(s, xx - 0.4, FY + 0.3, 0.8, 0.18, lab, 7, WHITE, bold=True, align=C, anchor=MSO_ANCHOR.MIDDLE)
source(s, FX, FY + FH + 0.08, FW, "Nguồn: mô hình nhóm; đồng thuận 9 báo cáo: trung bình 84.4 (tổng hợp thứ cấp).")

# SUMMARY
SX, SW = 5.55, 3.25
tb(s, SX, FY, SW, 0.22, "TÓM TẮT KHUYẾN NGHỊ", 9, INK, bold=True)
hline(s, SX, FY + 0.25, SW, color=GRAY)
rows = [["Khuyến nghị", "{{MUA}}"], ["Giá hiện tại (đồng)", "59,300"], ["Giá mục tiêu (đồng)", "83,500"],
        ["Upside", "{{+41%}}"], ["P/E FY27E", "7.4x"], ["P/B 6/2027E", "0.63x"],
        ["EV/EBIT FY26", "5.8x (11.2x nếu 50% nợ)"], ["WACC / g", "10.9% / 3.0%"],
        ["Backlog", "~70,000 tỷ (2.04x)"], ["Phương pháp", "30% P/E + 30% P/B + 40% DCF"]]
table(s, SX, FY + 0.3, SW, rows, [1.3, 1.95], rowh=0.245, size=8, align=[L, R], header=False, first_col_bold=True)
source(s, SX, FY + FH + 0.08, SW, "Stress ứng trước 25% nợ → 75,600 (+27%); 50% → 67,500 (+14%).")

# FINANCIALS
QX, QW = 9.05, 3.98
tb(s, QX, FY, QW, 0.22, "TÀI CHÍNH (tỷ đồng)", 9, INK, bold=True)
hline(s, QX, FY + 0.25, QW, color=GRAY)
rows = [["", "FY25A", "FY26A", "FY27E", "FY28E", "FY29E"],
        ["Doanh thu", "24,885", "34,340", "40,000", "44,800", "49,300"],
        ["Tăng trưởng", "18.2%", "38.0%", "16.5%", "12.0%", "10.0%"],
        ["Biên gộp", "3.28%", "4.21%", "4.45%", "4.45%", "4.45%"],
        ["EBIT", "533", "826", "1,100", "1,247", "1,391"],
        ["LNST", "456", "788", "894", "1,010", "1,129"],
        ["EPS (đồng)", "4,079", "7,045", "7,998", "9,032", "10,097"],
        ["OCF", "−1,154", "816", "236", "466", "626"],
        ["Tiền ròng", "1,315", "1,880", "1,915", "2,068", "2,381"]]
table(s, QX, FY + 0.3, QW, rows, [1.15, 0.75, 0.75, 0.75, 0.75, 0.75], rowh=0.255, size=8, align=[L, R, R, R, R, R],
      first_col_bold=True, bold_rows=(5,), header_fill=INK)
source(s, QX, FY + FH + 0.08, QW, "Nguồn: CTD (BCTC, Vietcap), mô hình nhóm CTD_Financial_Model_v4_analyst.")

# THESES
TYY = 4.2
rect(s, 0.3, TYY, 12.73, 1.55, fill=BOX)
tb(s, 0.42, TYY + 0.05, 4, 0.2, "LUẬN ĐIỂM ĐẦU TƯ", 9.5, RED, bold=True)
cols_t = [
    ("FDI và dự án phức tạp",
     "FDI ~50% doanh thu, repeat sales >75%, dẫn đầu LEED (38 dự án). Giải ngân FDI dự phóng tăng từ 27.6 lên 33.4 tỷ USD (2025-27F). "
     "Chuyển sang hạng mục ngầm, thép khẩu độ lớn, MEP nâng biên lên **4.45%**."),
    ("Backlog bảo chứng, dòng tiền cần theo dõi",
     "Backlog 70,000 tỷ phủ **2.04x** doanh thu, ~30,000 tỷ gối đầu FY28+. OCF dương nhờ ứng trước 9,029 tỷ; "
     "OCF FY27E chỉ 236 tỷ, tiền ròng 1,915 tỷ. Phải thu (DSO ~176 ngày) và tồn kho 9,906 tỷ là điểm cần theo dõi."),
    ("Định giá đã chiết khấu thấp",
     "P/E FY27E **7.4x**, P/B 0.63x, EV/EBIT 5.8x so với nhóm khu vực 14.5x và 1.61x. Giá mục tiêu **83,500** "
     "(+41%) = 30% P/E (75,980) + 30% P/B (81,267) + 40% DCF (91,253). Kể cả stress 25% ứng trước là nợ vẫn cho +27%; 50% → +14%."),
]
for k, (t, body) in enumerate(cols_t):
    x = 0.42 + k * 4.2
    tb(s, x, TYY + 0.3, 4.0, 0.2, t, 9, INK, bold=True)
    tb(s, x, TYY + 0.54, 3.95, 1.0, body, 8.5, INK, line=1.05)
    if k:
        vline(s, x - 0.1, TYY + 0.32, 1.15, color=LINE, lw=1)

# RISKS
tb(s, 0.3, 5.82, 4, 0.2, "RỦI RO CHÍNH VÀ ĐIỀU KIỆN XEM XÉT LẠI KHUYẾN NGHỊ", 9, RED, bold=True)
rows = [["Rủi ro", "Tín hiệu theo dõi", "Ngưỡng xem xét lại MUA", "Mức độ"],
        ["BĐS, nhà ở cao tầng (65% backlog)", "Phải thu 15,381 tỷ (40% tài sản), dự phòng 1,374 tỷ", "Trích lập thêm dự phòng >500 tỷ", "Cao"],
        ["Giá vật liệu và biên gộp", "Vật liệu +20-32%; hợp đồng giá cố định hết hạn 2027", "Biên gộp <3.8% hai quý (0.1 điểm ≈ 2.2% giá mục tiêu)", "Cao"],
        ["Dòng tiền phụ thuộc ứng trước", "OCF FY25 −1,154 tỷ; ứng trước 9,029 tỷ", "OCF âm hai quý, ứng trước giảm; độ phủ <1.5x", "Cao"],
        ["Lãi suất, quản trị", "Vay 5,943 tỷ, lãi vay +146%; CFO rời 26/4/2026", "Thu nhập tài chính ±100 tỷ = ±4.6% giá mục tiêu", "Vừa"]]
fills = {}
for r in range(1, 5):
    fills[(r, 3)] = PINK if rows[r][3] == "Cao" else BOX
table(s, 0.3, 6.04, 12.73, rows, [2.4, 4.0, 4.3, 0.7], rowh=0.2, size=7.5, align=[L, L, L, C], first_col_bold=True,
      fills=fills, header_fill=INK)

for sl in prs.slides:
    for sh in sl.shapes:
        spPr = sh._element.find(qn("p:spPr"))
        if spPr is not None and spPr.find(qn("a:effectLst")) is None and sh.shape_type != 3:
            e = etree.SubElement(spPr, qn("a:effectLst"))
        st = sh._element.find(qn("p:style"))
        if st is not None:
            er = st.find(qn("a:effectRef"))
            if er is not None:
                er.set("idx", "0")
prs.save("/home/user/slide-gf/CTD_Pitchbook_4slides.pptx")
print("saved")
