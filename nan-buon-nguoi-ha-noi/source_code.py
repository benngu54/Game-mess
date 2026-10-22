"""Tạo bài thuyết trình "Nạn buôn người ở Hà Nội" (16:9) bằng python-pptx.

Chạy:  python source_code.py
Kết quả: Nan_buon_nguoi_o_Ha_Noi.pptx, Loi_thuyet_trinh.docx, Nguon_anh.txt,
         Tai_lieu_tham_khao.txt (ảnh đã xử lý nằm trong build/img).

Yêu cầu: python-pptx >= 1.0, Pillow, lxml, python-docx.
Ảnh gốc nằm trong assets/photos (Open Images Dataset, ảnh Flickr giấy phép CC BY 2.0).
"""
import copy
from pathlib import Path

from docx import Document
from docx.shared import Pt as DocPt
from lxml import etree
from PIL import Image, ImageEnhance
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
PHOTOS = ROOT / "assets" / "photos"
BUILD = ROOT / "build" / "img"
OUT_PPTX = ROOT / "Nan_buon_nguoi_o_Ha_Noi.pptx"

W, H = 13.333, 7.5
TOTAL = 15

NAVY = "0B1020"
CHARCOAL = "15171D"
CRIMSON = "C83B45"
IVORY = "F4F0E8"
STEEL = "8993A4"
GOLD = "C6A56B"
INK = "2A2F3A"      # chữ phụ trên nền sáng
INK_SOFT = "5B6475"  # chú thích trên nền sáng
CARD_LIGHT = "EAE4D8"

HEAD = "Be Vietnam Pro ExtraBold"
HEAD2 = "Be Vietnam Pro"
BODY = "Arial"

# Ảnh: Open Images Dataset V4 (Google), bản gốc trên Flickr, giấy phép CC BY 2.0.
PHOTO_META = {
    "5507a99062b90ade": ("Tristan Schmurr", "A night rider in Hanoi", "https://www.flickr.com/photos/kewl/11610995284", 1,
                         "Người đi xe máy trên phố Hà Nội về đêm, hiệu ứng chuyển động"),
    "983f8b19a77cb92a": ("Brandon Giesbrecht", "Texting", "https://www.flickr.com/photos/naturegeak/6644669467", 2,
                         "Đôi tay đang nhắn tin trên điện thoại (không lộ mặt)"),
    "1d8a17c5368ba73c": ("Ethan McCandless", "Silhouette", "https://www.flickr.com/photos/mickyshoots/16618905772", 3,
                         "Bóng người nhìn nghiêng lúc hoàng hôn"),
    "862983007c81c0df": ("Leo Chuối", "@HaNoi", "https://www.flickr.com/photos/leo_nghinphu/16634953125", 4,
                         "Toàn cảnh đô thị Hà Nội lúc chiều tối (đã cắt bỏ phần chữ ký ở góc ảnh)"),
    "e25396eca86469fa": ("David McKelvey", "Ga Long Bien, Urban Discovery Tour, Hanoi",
                         "https://www.flickr.com/photos/dgmckelvey/7060733973", 4, "Ga Long Biên, Hà Nội"),
    "7423b6459cb1d8ab": ("Andrea Schaffer", "Hanoi", "https://www.flickr.com/photos/aschaf/5770571286", 4,
                         "Khu vực quảng trường Đông Kinh Nghĩa Thục, hồ Hoàn Kiếm"),
    "17378ce079b06369": ("Jorge Láscar", "Roundabout at the north end of Hoan Kiem Lake",
                         "https://www.flickr.com/photos/jlascar/4550346257", 5,
                         "Vòng xoay phía bắc hồ Hoàn Kiếm lúc chạng vạng"),
    "b9429d082799d857": ("Ryo FUKAsawa", "She seems to wear a watch only to job interviews.",
                         "https://www.flickr.com/photos/fukapon/8849909090", 6,
                         "Đôi tay đan vào nhau trên túi xách khi chờ phỏng vấn xin việc"),
    "c8e56c2df21206d0": ("Joey deVilla", "Close-up of the Alienware Laptop Keyboard",
                         "https://www.flickr.com/photos/accordionguy/3993036711", 7, "Bàn phím máy tính trong bóng tối"),
    "41bfcdc98e8e609f": ("Daniel Ramirez", "Workforce 2011 Job Fair @ Blaisdell Center",
                         "https://www.flickr.com/photos/danramarch/5736207332", 8, "Đông người tìm việc tại một hội chợ việc làm"),
    "0e5836c8bf70dc5b": ("Smitha Murthy", "Hanoi Lake", "https://www.flickr.com/photos/soulmuser/3382731868", 9,
                         "Một người ngồi một mình bên hồ Hoàn Kiếm"),
    "3846c6814a2f0c2c": ("Colin Mutchler", "hanoi streets", "https://www.flickr.com/photos/activefree/85078103", 10,
                         "Đường phố Hà Nội về đêm sau mưa"),
    "4ca5243643dc1818": ("Hong Phuc FOSSASIA", "GNOME Asia FOSSASIA Volunteer Meetup Vietnam",
                         "https://www.flickr.com/photos/fossasia/15957291412/", 11, "Buổi sinh hoạt, thảo luận của các bạn trẻ tình nguyện viên"),
    "1ad0685eda78ec4d": ("Linus Andersson", "On the phone", "https://www.flickr.com/photos/soulbrother73/18314379995", 12,
                         "Một người đang gọi điện thoại trên phố, ảnh đen trắng"),
    "869e098ffd768989": ("Nam-ho Park", "Hanoi", "https://www.flickr.com/photos/namho/2824834454", 13,
                         "Khung cửa mở ra khoảng sân sáng trong một di tích ở Hà Nội"),
    "af8683a47a215a63": ("dbgg1979", "Sunrise Photo walk at Celestial Gardens",
                         "https://www.flickr.com/photos/dbgg1979/4519512183", 14, "Nhóm người đứng ngắm bình minh"),
}
LICENSE = "CC BY 2.0"
LICENSE_URL = "https://creativecommons.org/licenses/by/2.0/"


# ---------------------------------------------------------------- ảnh

def prep(pid, w, h, focus=(0.5, 0.5), dark=0.0, sat=1.0, tint=0.0, grads=(), tag=""):
    """Cắt ảnh đúng tỷ lệ khung, chỉnh màu và phủ gradient; trả về đường dẫn ảnh đã xử lý."""
    im = Image.open(PHOTOS / f"{pid}.jpg").convert("RGB")
    iw, ih = im.size
    ratio = w / h
    if iw / ih > ratio:
        cw, ch = round(ih * ratio), ih
    else:
        cw, ch = iw, round(iw / ratio)
    left = min(max(round(focus[0] * iw - cw / 2), 0), iw - cw)
    top = min(max(round(focus[1] * ih - ch / 2), 0), ih - ch)
    im = im.crop((left, top, left + cw, top + ch))
    target = (round(w * 160), round(h * 160))
    if im.width > target[0]:
        im = im.resize(target, Image.LANCZOS)

    if sat != 1.0:
        im = ImageEnhance.Color(im).enhance(sat)
    if dark:
        im = ImageEnhance.Brightness(im).enhance(1 - dark)
    navy = Image.new("RGB", im.size, tuple(bytes.fromhex(NAVY)))
    if tint:
        im = Image.blend(im, navy, tint)
    for side, strength, extent in grads:
        im = Image.composite(navy, im, _ramp(im.size, side, strength, extent))

    BUILD.mkdir(parents=True, exist_ok=True)
    path = BUILD / f"{pid}{tag}.jpg"
    im.save(path, quality=90, optimize=True)
    return path


def _ramp(size, side, strength, extent):
    """Mặt nạ tuyến tính: đậm `strength` ở cạnh `side`, nhạt dần tới 0 sau `extent` phần khung."""
    w, h = size
    horizontal = side in ("left", "right")
    length = w if horizontal else h
    span = max(int(length * extent), 1)
    vals = []
    for i in range(length):
        t = max(0.0, 1 - i / span)
        vals.append(int(255 * strength * t * t * (3 - 2 * t)))
    if side in ("right", "bottom"):
        vals.reverse()
    strip = Image.new("L", (length, 1) if horizontal else (1, length))
    strip.putdata(vals)
    return strip.resize((w, h))


def photo(shapes, pid, x, y, w, h, name=None, **kw):
    pic = shapes.add_picture(str(prep(pid, w, h, **kw)), Inches(x), Inches(y), Inches(w), Inches(h))
    pic.name = name or f"Ảnh – {PHOTO_META[pid][1]}"
    return pic


def credit(slide, pids, x, y, w, color=STEEL, align=PP_ALIGN.LEFT, prefix="Ảnh"):
    authors = "; ".join(PHOTO_META[p][0] for p in pids)
    return text(slide.shapes, x, y, w, 0.3, f"{prefix}: {authors} (Flickr) · {LICENSE}",
                size=10, color=color, align=align, name="Ghi công ảnh")


# ---------------------------------------------------------------- hình & chữ

def _alpha(fill_parent, alpha):
    clr = fill_parent.find(".//" + qn("a:srgbClr"))
    etree.SubElement(clr, qn("a:alpha")).set("val", str(int(alpha * 100000)))


def rect(shapes, x, y, w, h, fill, line=None, line_w=1.0, alpha=None, shape=MSO_SHAPE.RECTANGLE, name=None):
    s = shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(fill)
        if alpha is not None:
            _alpha(s._element.spPr, alpha)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = RGBColor.from_string(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    s._element.remove(s._element.find(qn("p:style")))
    if name:
        s.name = name
    return s


def line(shapes, x1, y1, x2, y2, color=STEEL, width=1.0, arrow=False):
    c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = RGBColor.from_string(color)
    c.line.width = Pt(width)
    if arrow:
        ln = c.line._get_or_add_ln()
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return c


def _run(p, text, size, color, font, bold=False, italic=False, spacing=None):
    r = p.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.name = font
    f.bold = bold
    f.italic = italic
    f.color.rgb = RGBColor.from_string(color)
    if spacing:
        r._r.get_or_add_rPr().set("spc", str(spacing))
    return r


def text(shapes, x, y, w, h, paras, size=16, color=IVORY, font=BODY, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=None, space_after=0,
         spacing=None, bullet=None, margin=0.0, name=None):
    """paras: chuỗi (xuống dòng = đoạn mới) hoặc list đoạn; mỗi đoạn là chuỗi, list run (text, opts)
    hoặc dict {'runs'|'text', ...override}."""
    tb = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    if isinstance(paras, str):
        paras = paras.split("\n")
    base = dict(size=size, color=color, font=font, bold=bold, italic=italic, spacing=spacing)
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        opts = dict(base, align=align, line_spacing=line_spacing, space_after=space_after, bullet=bullet)
        if isinstance(para, dict):
            opts.update({k: v for k, v in para.items() if k not in ("runs", "text")})
            runs = para.get("runs") or [(para["text"], {})]
        elif isinstance(para, str):
            runs = [(para, {})]
        else:
            runs = para
        p.alignment = opts["align"]
        if opts["line_spacing"]:
            p.line_spacing = opts["line_spacing"]
        if opts["space_after"]:
            p.space_after = Pt(opts["space_after"])
        if opts["bullet"]:
            _bullet(p, opts["bullet"], opts["size"])
        for t, ro in runs:
            r = {k: opts[k] for k in base}
            r.update(ro)
            _run(p, t, **r)
    if name:
        tb.name = name
    return tb


def _bullet(p, color, size):
    pPr = p._p.get_or_add_pPr()
    indent = int(Pt(size) * 1.1)
    pPr.set("marL", str(indent))
    pPr.set("indent", str(-indent))
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(color)


def chrome(slide, label, x=0.6, dark=True, num_color=CRIMSON):
    """Mã chương kiểu hồ sơ điều tra ở góc trên + số trang ở góc dưới (motif chung của bài)."""
    muted = STEEL if dark else INK_SOFT
    n = slide_no(slide)
    k = text(slide.shapes, x, 0.5, 6.5, 0.3,
             [[(f"{n:02d}", dict(color=num_color, bold=True)), ("  —  " + label, dict(color=muted, bold=True))]],
             size=11, spacing=300, name="!!kicker")
    text(slide.shapes, W - 1.55, 7.02, 0.95, 0.3, f"{n:02d} / {TOTAL}", size=10, color=muted,
         align=PP_ALIGN.RIGHT, name="!!pagenum")
    return k


def slide_no(slide):
    return list(slide.part.package.presentation_part.presentation.slides).index(slide) + 1


def title(slide, x, y, w, h, t, size=36, color=IVORY, **kw):
    return text(slide.shapes, x, y, w, h, t, size=size, color=color, font=HEAD, line_spacing=1.05,
                name="!!title", **kw)


def group(slide, name):
    g = slide.shapes.add_group_shape()
    g.name = name
    return g


# ---------------------------------------------------------------- transition & animation (OOXML)

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


def transition(slide, kind="fade", ms=700):
    """Chèn transition có thời lượng (p14:dur); Morph dùng p159 với fallback Fade cho bản Office cũ."""
    effects = {
        "fade": '<p:fade/>',
        "push": '<p:push dir="u"/>',
        "morph": '<p159:morph option="byObject"/>',
    }
    requires = "p159" if kind == "morph" else "p14"
    xml = (
        f'<mc:AlternateContent xmlns:mc="{NS_MC}">'
        f'<mc:Choice xmlns:p14="{NS_P14}" xmlns:p159="{NS_P159}" Requires="{requires}">'
        f'<p:transition xmlns:p="{NS_P}" spd="slow" p14:dur="{ms}">{effects[kind]}</p:transition>'
        f'</mc:Choice><mc:Fallback>'
        f'<p:transition xmlns:p="{NS_P}" spd="med">{effects["fade" if kind == "morph" else kind]}</p:transition>'
        f'</mc:Fallback></mc:AlternateContent>'
    )
    el = etree.fromstring(xml)
    sld = slide._element
    anchor = sld.find(qn("p:clrMapOvr"))
    anchor.addnext(el)
    return el


def animate(slide, steps):
    """steps: list (shape, 'auto' | 'click' | 'with', effect). Hiệu ứng: 'fade' hoặc 'wipe'.

    Bước 'auto' chạy ngay sau khi vào slide (dùng cho tiêu đề); 'click' chờ người thuyết trình bấm;
    'with' chạy cùng lúc với bước liền trước.
    """
    ids = iter(range(3, 10000))
    clicks = []
    for shape, trigger, effect in steps:
        if trigger == "with" and clicks:
            clicks[-1][1].append((shape, effect))
        else:
            clicks.append((trigger, [(shape, effect)]))

    pars = []
    for trigger, items in clicks:
        auto = trigger == "auto"
        effects = []
        for i, (shape, effect) in enumerate(items):
            node = ("afterEffect" if auto else "clickEffect") if i == 0 else "withEffect"
            effects.append(_effect(next(ids), next(ids), next(ids), shape.shape_id, node, effect))
        cond = '<p:cond delay="indefinite"/>' + ('<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' if auto else "")
        pars.append(
            f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>'
        )
    builds = "".join(
        f'<p:bldP spid="{s.shape_id}" grpId="0" animBg="1"/>'
        for _, items in clicks for s, _ in items if s._element.tag == qn("p:sp")
    )
    xml = (
        f'<p:timing xmlns:p="{NS_P}"><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f'{"".join(pars)}'
        f'</p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        + (f'<p:bldLst>{builds}</p:bldLst>' if builds else "")
        + '</p:timing>'
    )
    slide._element.append(etree.fromstring(xml))


def _effect(id1, id2, id3, spid, node, effect):
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    preset, subtype, flt = {"fade": (10, 0, "fade"), "wipe": (22, 8, "wipe(left)")}[effect]
    return (
        f'<p:par><p:cTn id="{id1}" presetID="{preset}" presetClass="entr" presetSubtype="{subtype}" '
        f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{id2}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'</p:cTn>{tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        f'<p:to><p:strVal val="visible"/></p:to></p:set>'
        f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{id3}" dur="600"/>{tgt}</p:cBhvr>'
        f'</p:animEffect></p:childTnLst></p:cTn></p:par>'
    )


# ---------------------------------------------------------------- slides

def s01_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "5507a99062b90ade", 0, 0, W, H, focus=(0.6, 0.5), dark=0.25, sat=0.85,
          grads=[("left", 0.92, 0.75), ("bottom", 0.6, 0.45)])
    k = text(s.shapes, 0.6, 0.55, 8, 0.3, "SHADOWS OF THE CITY  ·  PHÓNG SỰ ẢNH", size=11, color=GOLD,
             bold=True, spacing=400, name="!!kicker")
    t = text(s.shapes, 0.6, 1.95, 9.5, 2.4,
             [[("NẠN BUÔN NGƯỜI", {})], [("Ở HÀ NỘI", dict(color=CRIMSON))]],
             size=66, font=HEAD, line_spacing=1.0, name="!!title")
    sub = text(s.shapes, 0.6, 4.3, 9, 0.5, "Thực trạng  –  Nguyên nhân  –  Hậu quả  –  Giải pháp",
               size=20, font=HEAD2, color=IVORY)
    info = text(s.shapes, 0.6, 5.55, 7.5, 0.95,
                ["Học sinh thực hiện:  ……………………………………………",
                 "Lớp:  ……………          Môn học:  ……………………………"],
                size=14, color=STEEL, space_after=6)
    credit(s, ["5507a99062b90ade"], 7.3, 7.02, 4.3, align=PP_ALIGN.RIGHT, prefix="Ảnh nền")
    transition(s, "fade", 1000)
    animate(s, [(k, "auto", "fade"), (t, "with", "fade"), (sub, "click", "fade"), (info, "with", "fade")])
    return s


def s02_question(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "983f8b19a77cb92a", 0, 0, 7.0, H, focus=(0.55, 0.5), dark=0.35, sat=0.7,
          grads=[("right", 1.0, 0.35), ("bottom", 0.5, 0.3)])
    k = chrome(s, "MỞ ĐẦU", x=7.6)
    lbl = text(s.shapes, 7.6, 1.15, 5, 0.3, "TÌNH HUỐNG GIẢ ĐỊNH", size=12, color=CRIMSON, bold=True, spacing=200)
    msg = group(s, "Tin nhắn giả định")
    rect(msg.shapes, 7.6, 1.6, 5.15, 1.95, CHARCOAL, line=STEEL, line_w=0.75, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    msg.shapes[0].adjustments[0] = 0.08
    text(msg.shapes, 7.85, 1.75, 4.7, 0.3, "Tin nhắn từ tài khoản lạ  ·  22:47", size=10, color=STEEL)
    text(msg.shapes, 7.85, 2.05, 4.7, 1.7,
         "“Chào em! Bên anh đang tuyển gấp nhân viên, việc nhẹ, lương 15–20 triệu/tháng, bao ăn ở. "
         "Không cần kinh nghiệm, không cần hợp đồng. Đi ngay tuần này, nhớ đừng kể với ai nhé!”",
         size=15, color=IVORY, line_spacing=1.1)
    q = title(s, 7.6, 4.15, 5.15, 1.4, "Bạn sẽ trả lời tin nhắn này chứ?", size=32)
    q.name = "Câu hỏi lớn"
    hint = text(s.shapes, 7.6, 5.7, 5.15, 0.8,
                "Hãy giữ câu trả lời của mình — cuối bài, chúng ta sẽ quay lại tin nhắn này.",
                size=15, color=STEEL, line_spacing=1.1)
    text(s.shapes, 7.6, 6.75, 4.3, 0.45,
         "* Tình huống minh họa do nhóm thực hiện xây dựng, không phải vụ việc có thật.",
         size=10, color=STEEL, italic=True)
    credit(s, ["983f8b19a77cb92a"], 0.45, 7.02, 6.2, prefix="Ảnh minh họa")
    transition(s, "fade")
    animate(s, [(k, "auto", "fade"), (lbl, "with", "fade"), (msg, "click", "fade"), (q, "click", "fade"),
                (hint, "click", "fade")])
    return s


def s03_definition(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "1d8a17c5368ba73c", 0, 0, 4.2, H, focus=(0.3, 0.5), dark=-0.45, sat=1.1,
          grads=[("right", 0.8, 0.22)])
    k = chrome(s, "KHÁI NIỆM", x=4.8)
    t = title(s, 4.8, 0.85, 7.9, 0.8, "Buôn người là gì?", size=38)
    d = text(s.shapes, 4.8, 1.75, 7.95, 1.0,
             [[("Là việc ", {}), ("tuyển mộ, vận chuyển, chuyển giao, chứa chấp hoặc tiếp nhận", dict(bold=True)),
               (" một người bằng thủ đoạn lừa dối, ép buộc hay lợi dụng hoàn cảnh khó khăn của họ — ", {}),
               ("nhằm mục đích bóc lột.", dict(bold=True, color=GOLD))]],
             size=16, line_spacing=1.15)

    panels = [
        ("1 · HÀNH VI", "Làm gì?", "Tuyển mộ, vận chuyển, chuyển giao, chứa chấp, tiếp nhận người"),
        ("2 · THỦ ĐOẠN", "Bằng cách nào?", "Lừa dối, đe dọa, dùng vũ lực, lợi dụng sự phụ thuộc hoặc hoàn cảnh khó khăn"),
        ("3 · MỤC ĐÍCH", "Để làm gì?", "Bóc lột tình dục, cưỡng bức lao động, lấy bộ phận cơ thể hoặc mục đích vô nhân đạo khác"),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (d, "click", "fade")]
    for i, (lab, head, body) in enumerate(panels):
        x = 4.8 + i * 2.75
        g = group(s, f"Yếu tố {i + 1}")
        rect(g.shapes, x, 2.95, 2.45, 2.5, CHARCOAL if i < 2 else "2A1519")
        text(g.shapes, x + 0.2, 3.1, 2.1, 0.3, lab, size=11, color=CRIMSON, bold=True, spacing=150)
        text(g.shapes, x + 0.2, 3.42, 2.1, 0.45, head, size=17, font=HEAD2, bold=True)
        text(g.shapes, x + 0.2, 3.9, 2.1, 1.5, body, size=14, color="D9D4CA", line_spacing=1.1)
        if i:
            text(g.shapes, x - 0.3, 3.9, 0.3, 0.5, "+", size=26, color=GOLD, font=HEAD,
                        align=PP_ALIGN.CENTER)
        steps.append((g, "click", "fade"))

    note = text(s.shapes, 4.8, 5.6, 7.95, 0.45,
                "Với trẻ em (dưới 18 tuổi theo chuẩn của LHQ): chỉ cần có hành vi + mục đích bóc lột là đủ.",
                size=14, color=GOLD)
    diff = text(s.shapes, 4.8, 6.08, 7.95, 0.75,
                [[("KHÔNG PHẢI BUÔN NGƯỜI:  ", dict(color=STEEL, bold=True, size=12, spacing=100)),
                  ("đi làm xa tự nguyện, có hợp đồng rõ ràng, tự giữ giấy tờ tùy thân, "
                   "được tự do nghỉ việc và liên lạc với gia đình.", {})]],
                size=14, line_spacing=1.1)
    text(s.shapes, 4.8, 7.02, 6.9, 0.3,
         "Nguồn: Nghị định thư Palermo (LHQ, 2000), Điều 3; Luật Phòng, chống mua bán người 2024.",
         size=10, color=STEEL)
    credit(s, ["1d8a17c5368ba73c"], 0.3, 7.02, 3.7, prefix="Ảnh minh họa")
    steps += [(note, "click", "fade"), (diff, "click", "fade")]
    transition(s, "fade")
    animate(s, steps)
    return s


def s04_hanoi(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    tiles = group(s, "Ảnh Hà Nội")
    photo(tiles.shapes, "862983007c81c0df", 0, 0, 7.4, 3.6, focus=(0.5, 1.0), dark=0.1, sat=0.9,
          grads=[("bottom", 0.7, 0.35)])
    photo(tiles.shapes, "e25396eca86469fa", 0, 3.7, 3.65, 3.8, focus=(0.38, 0.5), dark=0.15, sat=0.8,
          grads=[("bottom", 0.85, 0.4)])
    photo(tiles.shapes, "7423b6459cb1d8ab", 3.75, 3.7, 3.65, 3.8, focus=(0.4, 0.5), dark=0.15, sat=0.8,
          grads=[("bottom", 0.85, 0.4)])
    for x, y, cap in ((0.35, 3.1, "Toàn cảnh Hà Nội lúc chiều tối"), (0.35, 6.95, "Ga Long Biên"),
                      (4.1, 6.95, "Phố cổ, quanh hồ Hoàn Kiếm")):
        text(tiles.shapes, x, y, 3.3, 0.3, cap, size=11, color=IVORY, bold=True)

    k = chrome(s, "BỐI CẢNH", x=7.95)
    t = title(s, 7.95, 0.85, 4.85, 1.3, "Vì sao Hà Nội cần quan tâm?", size=32)
    rows = [
        ("Trung tâm việc làm, học tập", "Thu hút rất nhiều người trẻ từ nơi khác đến tìm cơ hội."),
        ("Đầu mối giao thông lớn", "Nhà ga, bến xe, sân bay: di chuyển nhanh — kể cả với kẻ xấu."),
        ("Môi trường mạng dày đặc", "Tin tuyển dụng, hội nhóm việc làm lan truyền rất nhanh."),
        ("Người mới lập nghiệp dễ cả tin", "Thiếu kinh nghiệm, thiếu người thân bên cạnh để hỏi ý kiến."),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (tiles, "with", "fade")]
    for i, (head, body) in enumerate(rows):
        y = 2.3 + i * 0.95
        g = group(s, f"Yếu tố Hà Nội {i + 1}")
        text(g.shapes, 7.95, y, 0.6, 0.4, f"0{i + 1}", size=18, font=HEAD, color=CRIMSON)
        text(g.shapes, 8.6, y, 4.2, 0.85,
             [{"text": head, "size": 15, "bold": True, "font": HEAD2}, {"text": body, "color": STEEL}],
             size=14, line_spacing=1.05)
        steps.append((g, "click", "fade"))
    note = text(s.shapes, 7.95, 6.15, 4.85, 0.75,
                "Lưu ý: chưa có dữ liệu cho thấy Hà Nội có tỷ lệ mua bán người cao nhất cả nước; "
                "điều đáng lo là mức độ tiếp xúc với nguy cơ.", size=12, color=GOLD, italic=True)
    credit(s, ["862983007c81c0df", "e25396eca86469fa", "7423b6459cb1d8ab"], 7.95, 6.95, 3.65)
    steps.append((note, "click", "fade"))
    transition(s, "push", 800)
    animate(s, steps)
    return s


def s05_data(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "17378ce079b06369", 0, 0, 3.0, H, focus=(0.55, 0.5), dark=0.45, sat=0.7, tint=0.25,
          grads=[("right", 1.0, 0.45)])
    k = chrome(s, "THỰC TRẠNG", x=3.6)
    t = title(s, 3.6, 0.85, 9.1, 0.75, "Con số phía sau những cạm bẫy", size=34)
    sub = text(s.shapes, 3.6, 1.6, 9.15, 0.45,
               "Chưa có số liệu công khai riêng của Hà Nội 2024–2025 → dùng số liệu toàn quốc (Bộ Công an).",
               size=14, color=STEEL)

    big = group(s, "Số liệu nổi bật 2024")
    for i, (num, label, col) in enumerate((("163", "vụ mua bán người được điều tra, xử lý", IVORY),
                                           ("500", "nạn nhân được xác định", CRIMSON))):
        y = 2.3 + i * 1.6
        text(big.shapes, 3.6, y, 3.7, 0.95, num, size=60, font=HEAD, color=col, line_spacing=0.9)
        text(big.shapes, 3.6, y + 0.98, 3.7, 0.4, label, size=14, color=STEEL)

    data = CategoryChartData()
    data.categories = ["Vụ án", "Đối tượng", "Nạn nhân"]
    data.add_series("Tổng số điều tra, xử lý", (163, 455, 500))
    data.add_series("Trong đó: khởi tố mới", (91, 237, 336))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.45), Inches(2.45), Inches(5.3), Inches(3.05), data)
    gf.name = "Biểu đồ số liệu 2024"
    chart = gf.chart
    chart.font.name = BODY
    chart.font.size = Pt(11)
    chart.font.color.rgb = RGBColor.from_string(STEEL)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.color.rgb = RGBColor.from_string(IVORY)
    plot = chart.plots[0]
    plot.gap_width = 70
    plot.overlap = -10
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    labels.font.size = Pt(12)
    labels.font.bold = True
    labels.font.color.rgb = RGBColor.from_string(IVORY)
    for series, col in zip(plot.series, (STEEL, CRIMSON)):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = RGBColor.from_string(col)
    va = chart.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.maximum_scale = 580
    ca = chart.category_axis
    ca.format.line.color.rgb = RGBColor.from_string("3A4152")
    ca.tick_labels.font.color.rgb = RGBColor.from_string(IVORY)
    ca.tick_labels.font.size = Pt(12)
    cap = text(s.shapes, 7.55, 2.15, 5.2, 0.3, "NĂM 2024  ·  TOÀN QUỐC  ·  NGUỒN: BỘ CÔNG AN", size=10,
               color=GOLD, bold=True, spacing=150)

    calls = group(s, "Các chỉ số khác")
    for i, (num, label) in enumerate((("≈ 65%", "nạn nhân là nữ giới (quý I/2025)"),
                                      ("> 44%", "nạn nhân dưới 16 tuổi (quý I/2025)"),
                                      ("100+", "nạn nhân là nam giới, chỉ riêng quý III/2024"))):
        x = 3.6 + i * 3.1
        text(calls.shapes, x, 5.65, 2.9, 0.55, num, size=28, font=HEAD, color=GOLD)
        text(calls.shapes, x, 6.2, 2.9, 0.55, label, size=13, color=IVORY, line_spacing=1.05)
    text(s.shapes, 3.6, 6.85, 8.2, 0.45,
         "Nguồn: Bộ Công an, dẫn theo Vietnamnet, An ninh Thủ đô (2025), Báo Lào Cai (2025), PLO (2024). "
         "Năm 2024: 163 vụ/455 đối tượng/500 nạn nhân; khởi tố mới 91 vụ/237 đối tượng/336 nạn nhân.",
         size=10, color=STEEL)
    credit(s, ["17378ce079b06369"], 0.3, 7.02, 2.6)
    transition(s, "fade")
    animate(s, [(k, "auto", "fade"), (t, "with", "fade"), (sub, "with", "fade"), (big, "click", "fade"),
                (gf, "click", "wipe"), (cap, "with", "fade"), (calls, "click", "fade")])
    return s


def s06_tactics(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, CHARCOAL)
    photo(s.shapes, "b9429d082799d857", 7.2, 0, W - 7.2, H, focus=(0.62, 0.55), dark=0.1, sat=0.55,
          grads=[("left", 1.0, 0.4)])
    k = chrome(s, "THỦ ĐOẠN")
    t = title(s, 0.6, 0.85, 6.4, 1.3, "Cạm bẫy thường bắt đầu bằng một lời mời", size=30)
    rows = [
        ("Việc nhẹ, lương cao", "Mức lương hấp dẫn bất thường, không cần kinh nghiệm"),
        ("Tuyển dụng trực tuyến mập mờ", "Không địa chỉ, không hợp đồng, chỉ nhắn tin qua mạng"),
        ("Lợi dụng tình cảm, lòng tin", "Làm quen, hẹn hò qua mạng; nhờ người quen dẫn dắt"),
        ("Đưa đi làm xa, xuất cảnh trái phép", "Có vụ dụ sang Campuchia làm cho cơ sở lừa đảo"),
        ("Khống chế, đe dọa, giữ giấy tờ", "Ép làm việc, bắt nộp tiền chuộc mới cho về"),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade")]
    for i, (head, body) in enumerate(rows):
        y = 2.3 + i * 0.82
        g = group(s, f"Thủ đoạn {i + 1}")
        text(g.shapes, 0.6, y + 0.03, 0.5, 0.35, f"0{i + 1}", size=14, color=CRIMSON, bold=True)
        text(g.shapes, 1.15, y, 5.8, 0.8,
             [{"text": head, "size": 16, "bold": True, "font": HEAD2}, {"text": body, "color": STEEL}],
             size=14, line_spacing=1.05)
        steps.append((g, "click", "fade"))
    note = text(s.shapes, 0.6, 6.42, 6.3, 0.55,
                "Các thủ đoạn trên đã được Bộ Công an và báo chí chính thống ghi nhận (nguồn ở slide 15). "
                "Slide sau: dấu hiệu cảnh báo chung để tự nhận diện.", size=12, color=GOLD)
    credit(s, ["b9429d082799d857"], 7.45, 7.02, 4.2, prefix="Ảnh minh họa")
    steps.append((note, "click", "fade"))
    transition(s, "fade")
    animate(s, steps)
    return s


def s07_social(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "c8e56c2df21206d0", 0, 0, W, H, focus=(0.5, 0.5), dark=0.62, sat=0.6, tint=0.35,
          grads=[("left", 0.6, 0.45)])
    k = chrome(s, "MẠNG XÃ HỘI & HỌC SINH")
    t = title(s, 0.6, 0.85, 8, 0.8, "Khi lời mời đến qua màn hình", size=34)

    phone = group(s, "Mô phỏng tin nhắn")
    text(phone.shapes, 0.85, 1.75, 3.6, 0.3, "MÔ PHỎNG · NỘI DUNG GIẢ ĐỊNH", size=10, color=CRIMSON, bold=True,
         spacing=150)
    body = rect(phone.shapes, 0.85, 2.1, 3.5, 4.8, CHARCOAL, line=STEEL, line_w=1.25,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    body.adjustments[0] = 0.1
    text(phone.shapes, 1.1, 2.3, 3.0, 0.3, "Tuyển dụng Online 24h  ·  đang hoạt động", size=10, color=STEEL)
    line(phone.shapes, 1.05, 2.68, 4.15, 2.68, color="2E3442", width=0.75)
    bubbles = [
        (False, 1.05, 2.85, 0.95, "Em ơi, bên chị cần 5 bạn trực máy, lương 20 triệu, làm online thôi."),
        (True, 1.55, 3.95, 0.7, "Công ty ở đâu ạ? Có hợp đồng không chị?"),
        (False, 1.05, 4.8, 1.05, "Làm ở tỉnh khác, chị lo vé xe. Gửi chị ảnh CCCD để giữ chỗ nhé, đừng nói với bố mẹ!"),
    ]
    for mine, x, y, h, msg in bubbles:
        b = rect(phone.shapes, x, y, 2.6, h, "5A2229" if mine else "262B36", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        b.adjustments[0] = 0.18
        text(phone.shapes, x + 0.12, y + 0.07, 2.36, h - 0.1, msg, size=13, color=IVORY, line_spacing=1.05)
    bar = rect(phone.shapes, 1.05, 6.25, 3.1, 0.42, "0F1422", line="2E3442", line_w=0.75,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    bar.adjustments[0] = 0.5
    text(phone.shapes, 1.25, 6.31, 2.8, 0.3, "Nhập tin nhắn…", size=11, color=STEEL)

    flags = [
        "Lương cao bất thường so với công việc đơn giản",
        "Không có hợp đồng; địa chỉ công ty mơ hồ, không kiểm chứng được",
        "Yêu cầu giữ bí mật, “đừng nói với bố mẹ”, phải đi ngay",
        "Đòi gửi ảnh CCCD, nộp phí hoặc đặt cọc trước khi làm",
        "Tài khoản mới lập, kéo sang Telegram để “nói chuyện riêng”",
        "Hứa bao vé xe, bao ăn ở, đưa sang tỉnh khác hoặc ra nước ngoài",
    ]
    head = text(s.shapes, 5.2, 1.85, 7.5, 0.45, "6 dấu hiệu cảnh báo — thấy là dừng lại", size=20, font=HEAD2,
                bold=True, color=GOLD)
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (phone, "click", "fade"), (head, "click", "fade")]
    for i, f in enumerate(flags):
        y = 2.5 + i * 0.64
        g = group(s, f"Dấu hiệu {i + 1}")
        rect(g.shapes, 5.2, y + 0.06, 0.26, 0.26, None, line=CRIMSON, line_w=1.5)
        text(g.shapes, 5.7, y, 7.0, 0.62, f, size=16, line_spacing=1.0)
        steps.append((g, "click" if i % 2 == 0 else "with", "fade"))
    plat = text(s.shapes, 5.2, 6.5, 7.5, 0.45,
                "Hay bị lợi dụng: Facebook · TikTok · Zalo · Telegram · nhóm “việc làm thêm”",
                size=14, color=STEEL)
    credit(s, ["c8e56c2df21206d0"], 0.85, 7.02, 4.5, prefix="Ảnh nền")
    steps.append((plat, "click", "fade"))
    transition(s, "fade")
    animate(s, steps)
    return s


def s08_causes(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "41bfcdc98e8e609f", 0, 0, W, 3.0, focus=(0.5, 0.6), dark=0.45, sat=0.5, tint=0.3,
          grads=[("bottom", 1.0, 0.55), ("left", 0.6, 0.6)])
    k = chrome(s, "NGUYÊN NHÂN")
    t = title(s, 0.6, 0.85, 9, 0.8, "Vì sao nhiều người rơi vào bẫy?", size=34)
    sub = text(s.shapes, 0.6, 1.65, 9, 0.45,
               "Không có một nguyên nhân duy nhất — các yếu tố cộng hưởng với nhau.", size=15, color="D9D4CA")
    causes = [
        ("Kinh tế & nhu cầu việc làm", "Thu nhập thấp, áp lực tài chính, mong muốn đổi đời nhanh."),
        ("Thiếu kỹ năng tự bảo vệ", "Ít hiểu biết pháp luật, chưa biết kiểm chứng tin tuyển dụng."),
        ("Công nghệ & lòng tin bị lợi dụng", "Tài khoản ảo, quan hệ quen biết được dựng lên qua mạng."),
        ("Khó phát hiện, báo tin, hỗ trợ", "Nạn nhân sợ hãi, mặc cảm; thủ đoạn kín đáo, xuyên biên giới."),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (sub, "with", "fade")]
    centers = []
    for i, (head, body) in enumerate(causes):
        x = 0.6 + i * 3.09
        centers.append(x + 1.43)
        g = group(s, f"Nguyên nhân {i + 1}")
        rect(g.shapes, x, 3.0, 2.86, 2.2, CHARCOAL)
        text(g.shapes, x + 0.2, 3.13, 2.5, 0.35, f"0{i + 1}", size=16, font=HEAD, color=CRIMSON)
        text(g.shapes, x + 0.2, 3.5, 2.5, 1.65,
             [{"text": head, "size": 16, "bold": True, "font": HEAD2, "space_after": 4},
              {"text": body, "color": STEEL}], size=14, line_spacing=1.05)
        steps.append((g, "click" if i == 0 else "with", "fade"))
    link = group(s, "Sơ đồ liên kết")
    for cx in centers:
        line(link.shapes, cx, 5.2, cx, 5.45, color=STEEL, width=1.25)
    line(link.shapes, centers[0], 5.45, centers[-1], 5.45, color=STEEL, width=1.25)
    line(link.shapes, W / 2, 5.45, W / 2, 5.78, color=CRIMSON, width=1.5, arrow=True)
    rect(link.shapes, 0.6, 5.8, 12.13, 0.65, CRIMSON)
    text(link.shapes, 0.6, 5.8, 12.13, 0.65, "Kẻ phạm tội nhắm vào chính những khoảng trống này để dụ dỗ, khống chế",
         size=16, bold=True, font=HEAD2, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    blame = text(s.shapes, 0.6, 6.6, 9, 0.4, "Trách nhiệm thuộc về kẻ phạm tội — nạn nhân không có lỗi.",
                 size=15, color=GOLD, bold=True)
    credit(s, ["41bfcdc98e8e609f"], 8.2, 2.58, 4.55, align=PP_ALIGN.RIGHT, prefix="Ảnh minh họa")
    steps += [(link, "click", "wipe"), (blame, "click", "fade")]
    transition(s, "push", 800)
    animate(s, steps)
    return s


def s09_consequences(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "0e5836c8bf70dc5b", 0, 0, 7.6, H, focus=(0.45, 0.55), dark=0.3, sat=0.6, tint=0.15,
          grads=[("right", 1.0, 0.3), ("bottom", 0.75, 0.4)])
    k = chrome(s, "HẬU QUẢ", x=8.1)
    t = title(s, 8.1, 0.85, 4.7, 1.3, "Những vết thương không nhìn thấy", size=30)
    groups = [
        ("ĐỐI VỚI NẠN NHÂN", "Tổn hại sức khỏe, sang chấn tâm lý; mất tự do, nhân phẩm; dang dở việc học và tương lai."),
        ("ĐỐI VỚI GIA ĐÌNH", "Lo âu, đau đớn khi mất liên lạc; gánh nặng nợ nần, tiền chuộc; khó khăn khi đón người thân trở về."),
        ("ĐỐI VỚI XÃ HỘI", "Mất an ninh trật tự, tiếp tay cho tội phạm có tổ chức; làm xói mòn niềm tin trong cộng đồng."),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade")]
    for i, (lab, body) in enumerate(groups):
        y = 2.35 + i * 1.45
        g = group(s, lab.title())
        text(g.shapes, 8.1, y, 4.7, 0.3, lab, size=12, color=CRIMSON, bold=True, spacing=150)
        text(g.shapes, 8.1, y + 0.33, 4.7, 1.05, body, size=15, line_spacing=1.1)
        steps.append((g, "click", "fade"))
    quote = text(s.shapes, 0.6, 5.75, 6.3, 1.0,
                 "Ai cũng cần được lắng nghe và hỗ trợ — không phải bị phán xét.",
                 size=21, font=HEAD2, italic=True, line_spacing=1.1)
    credit(s, ["0e5836c8bf70dc5b"], 0.6, 7.02, 6.2)
    steps.append((quote, "click", "fade"))
    transition(s, "fade")
    animate(s, steps)
    return s


def s10_case(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, CHARCOAL)
    photo(s.shapes, "3846c6814a2f0c2c", 8.3, 0, W - 8.3, H, focus=(0.6, 0.5), dark=0.2, sat=0.8,
          grads=[("left", 1.0, 0.35), ("bottom", 0.8, 0.2)])
    k = chrome(s, "TÌNH HUỐNG")
    t = title(s, 0.6, 0.85, 7.6, 0.8, "Lời mời “quá tốt để là thật”", size=32)
    case = group(s, "Tình huống giả định")
    rect(case.shapes, 0.6, 1.75, 7.35, 1.4, NAVY, line=CRIMSON, line_w=1.0)
    text(case.shapes, 0.8, 1.85, 7.0, 0.3, "TÌNH HUỐNG GIẢ ĐỊNH", size=11, color=CRIMSON, bold=True, spacing=150)
    text(case.shapes, 0.8, 2.15, 7.0, 1.0,
         "Một bạn 18 tuổi ở Hà Nội được “người quen trên mạng” mời làm phục vụ ở tỉnh khác, lương 25 triệu/tháng. "
         "Điều kiện: không ký hợp đồng, giữ bí mật với gia đình và lên xe ngay tối nay.",
         size=15, line_spacing=1.08)
    cells = [
        ("DẤU HIỆU NGUY HIỂM", CRIMSON, "Lương cao bất thường, không hợp đồng, bắt giữ bí mật, ép đi gấp."),
        ("LỰA CHỌN AN TOÀN", GOLD, "Từ chối hoặc hoãn lại; không gửi giấy tờ; không lên xe cùng người lạ."),
        ("AI CÓ THỂ HỖ TRỢ", "B9C2D3", "Bố mẹ, thầy cô, công an phường/xã; gọi 111 hoặc 113 khi cần."),
        ("CÁCH PHÒNG TRÁNH", GOLD, "Kiểm tra công ty, đọc kỹ hợp đồng, báo gia đình địa chỉ nơi làm việc."),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (case, "click", "fade")]
    for i, (lab, col, body) in enumerate(cells):
        x = 0.6 + (i % 2) * 3.75
        y = 3.4 + (i // 2) * 1.6
        g = group(s, lab.title())
        rect(g.shapes, x, y, 3.6, 1.42, NAVY)
        text(g.shapes, x + 0.2, y + 0.15, 3.25, 0.3, lab, size=11, color=col, bold=True, spacing=120)
        text(g.shapes, x + 0.2, y + 0.48, 3.25, 0.9, body, size=15, line_spacing=1.08)
        steps.append((g, "click", "fade"))
    credit(s, ["3846c6814a2f0c2c"], 8.55, 7.02, 3.3)
    transition(s, "fade")
    animate(s, steps)
    return s


def s11_solutions(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, IVORY)
    pics = group(s, "Ảnh cộng đồng")
    photo(pics.shapes, "4ca5243643dc1818", 7.4, 0, W - 7.4, 3.55, focus=(0.55, 0.45), sat=0.95)
    k = chrome(s, "GIẢI PHÁP", dark=False)
    t = title(s, 0.6, 0.85, 6.5, 1.3, "Phòng chống buôn người là việc của tất cả chúng ta", size=30, color=NAVY)
    sub = text(s.shapes, 0.6, 2.3, 6.5, 0.95,
               "Bốn lớp bảo vệ cùng hoạt động: khi một lớp yếu đi, các lớp còn lại giúp phát hiện và ngăn chặn kịp thời.",
               size=15, color=INK, line_spacing=1.1)
    cols = [
        ("Cá nhân", ["Kiểm chứng mọi lời mời việc làm", "Bảo mật thông tin, giấy tờ cá nhân",
                     "Biết nói “không” và biết cầu cứu"]),
        ("Gia đình", ["Lắng nghe, trò chuyện thường xuyên", "Biết con làm gì, ở đâu, với ai",
                      "Không chê trách khi con kể điều bất thường"]),
        ("Trường học & cộng đồng", ["Giáo dục pháp luật, kỹ năng số", "Ngoại khóa, câu lạc bộ về an toàn",
                                    "Tổ dân phố chú ý dấu hiệu lạ"]),
        ("Cơ quan chức năng", ["Kiểm soát dịch vụ việc làm, tin tuyển dụng", "Điều tra, xử lý nghiêm các đường dây",
                               "Tiếp nhận, bảo vệ, hỗ trợ nạn nhân"]),
    ]
    steps = [(k, "auto", "fade"), (t, "with", "fade"), (pics, "with", "fade"), (sub, "click", "fade")]
    for i, (head, items) in enumerate(cols):
        x = 0.6 + i * 3.08
        g = group(s, f"Nhóm giải pháp {i + 1}")
        rect(g.shapes, x, 3.85, 2.9, 3.0, CARD_LIGHT)
        text(g.shapes, x + 0.2, 3.95, 2.5, 0.5, f"0{i + 1}", size=24, font=HEAD, color=CRIMSON)
        text(g.shapes, x + 0.2, 4.5, 2.6, 0.4, head, size=15, font=HEAD2, bold=True, color=NAVY)
        text(g.shapes, x + 0.2, 4.95, 2.6, 1.85, items, size=14, color=INK, bullet=CRIMSON, space_after=4,
             line_spacing=1.0)
        steps.append((g, "click", "fade"))
    credit(s, ["4ca5243643dc1818"], 7.4, 3.58, 5.5, color=INK_SOFT, align=PP_ALIGN.RIGHT, prefix="Ảnh minh họa")
    transition(s, "push", 800)
    animate(s, steps)
    return s


def s12_action(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, IVORY)
    photo(s.shapes, "1ad0685eda78ec4d", 0, 0, 4.0, H, focus=(0.27, 0.5), dark=0.05)
    k = chrome(s, "KHI NGHI NGỜ", x=4.6, dark=False)
    t = title(s, 4.6, 0.85, 8.1, 0.8, "Nghi ngờ? Làm ngay 5 bước", size=34, color=NAVY)
    steps_data = [
        ("Dừng lại", "Không vội làm theo yêu cầu; không chuyển tiền, không gửi giấy tờ."),
        ("Chia sẻ", "Kể ngay với bố mẹ, thầy cô hoặc người lớn mà bạn tin tưởng."),
        ("Lưu bằng chứng", "Chụp lại tin nhắn, tài khoản, số điện thoại — nếu việc đó an toàn."),
        ("Báo cơ quan chức năng", "Công an phường/xã gần nhất, hoặc gọi Tổng đài 111."),
        ("Ưu tiên an toàn", "Đang gặp nguy hiểm: đến nơi đông người và gọi ngay 113."),
    ]
    anim = [(k, "auto", "fade"), (t, "with", "fade")]
    line(s.shapes, 4.85, 2.2, 4.85, 6.05, color="C9C2B4", width=1.5)
    for i, (head, body) in enumerate(steps_data):
        y = 1.95 + i * 0.97
        g = group(s, f"Bước {i + 1}")
        rect(g.shapes, 4.6, y, 0.5, 0.5, CRIMSON, shape=MSO_SHAPE.OVAL)
        text(g.shapes, 4.6, y, 0.5, 0.5, str(i + 1), size=16, font=HEAD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        text(g.shapes, 5.3, y - 0.02, 4.35, 0.95,
             [{"text": head, "size": 16, "bold": True, "font": HEAD2, "color": NAVY}, {"text": body, "color": INK}],
             size=14, line_spacing=1.05)
        anim.append((g, "click", "fade"))
    hot = group(s, "Số điện thoại hỗ trợ")
    rect(hot.shapes, 9.9, 1.95, 2.85, 2.35, NAVY)
    text(hot.shapes, 10.1, 2.05, 2.5, 0.9, "111", size=54, font=HEAD, color=IVORY)
    text(hot.shapes, 10.1, 2.95, 2.5, 0.8, "Tổng đài quốc gia về phòng, chống mua bán người", size=13,
         color=IVORY, line_spacing=1.05)
    text(hot.shapes, 10.1, 3.82, 2.5, 0.3, "MIỄN PHÍ · 24/7", size=11, color=GOLD, bold=True, spacing=150)
    rect(hot.shapes, 9.9, 4.45, 2.85, 1.9, CRIMSON)
    text(hot.shapes, 10.1, 4.55, 2.5, 0.9, "113", size=54, font=HEAD, color=IVORY)
    text(hot.shapes, 10.1, 5.45, 2.5, 0.8, "Cảnh sát khẩn cấp — khi có nguy hiểm trực tiếp", size=13,
         color=IVORY, line_spacing=1.05)
    text(s.shapes, 4.6, 6.85, 7.1, 0.45,
         "Nguồn: Tổng đài 111 theo Nghị định 162/2025/NĐ-CP, khai trương 19/9/2025 (VnEconomy, 2025); "
         "113 là số gọi khẩn cấp của lực lượng Công an.", size=10, color=INK_SOFT)
    credit(s, ["1ad0685eda78ec4d"], 0.25, 7.02, 3.6, color=IVORY, prefix="Ảnh minh họa")
    anim.append((hot, "click", "fade"))
    transition(s, "morph", 1000)
    animate(s, anim)
    return s


def s13_law(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "869e098ffd768989", 0, 0, 5.2, H, focus=(0.5, 0.5), dark=0.05, sat=0.9,
          grads=[("right", 0.85, 0.25)])
    k = chrome(s, "PHÁP LUẬT", x=5.8)
    t = title(s, 5.8, 0.85, 6.95, 1.3, "Pháp luật bảo vệ nạn nhân", size=34)
    events = [
        ("28/11/2024", "Quốc hội thông qua Luật Phòng, chống mua bán người số 53/2024/QH15 (8 chương, 63 điều)."),
        ("29/6/2025", "Nghị định 162/2025/NĐ-CP hướng dẫn thi hành; Tổng đài 111 tiếp nhận tin báo, tố giác."),
        ("01/7/2025", "Luật có hiệu lực, thay thế Luật Phòng, chống mua bán người năm 2011."),
        ("07/01/2026", "Quyết định 33/QĐ-TTg: Chương trình phòng, chống mua bán người 2026–2030."),
    ]
    anim = [(k, "auto", "fade"), (t, "with", "fade")]
    for i, (date, body) in enumerate(events):
        y = 2.3 + i * 0.75
        g = group(s, f"Mốc {date}")
        text(g.shapes, 5.8, y, 1.4, 0.35, date, size=14, color=GOLD, bold=True)
        text(g.shapes, 7.25, y, 5.5, 0.7, body, size=14, line_spacing=1.05)
        anim.append((g, "click", "fade"))
    core = group(s, "Điểm cốt lõi")
    rect(core.shapes, 5.8, 5.4, 6.95, 1.4, CHARCOAL)
    text(core.shapes, 6.0, 5.48, 6.6, 0.3, "ĐIỂM CỐT LÕI", size=11, color=CRIMSON, bold=True, spacing=150)
    text(core.shapes, 6.0, 5.78, 6.6, 1.0,
         ["Lấy nạn nhân làm trung tâm; mở rộng những người được bảo vệ, hỗ trợ.",
          "Nghiêm cấm cản trở giải cứu, hỗ trợ nạn nhân; cấm bao che tội phạm.",
          "Hành vi mua bán người bị xử lý hình sự theo Điều 150, 151 Bộ luật Hình sự."],
         size=13, bullet=CRIMSON, line_spacing=1.0, space_after=2)
    text(s.shapes, 5.8, 6.9, 6.0, 0.45,
         "Nguồn: Luật 53/2024/QH15 (VBHN 138/VBHN-VPQH, 2025); NĐ 162/2025/NĐ-CP; QĐ 33/QĐ-TTg (2026).",
         size=10, color=STEEL)
    credit(s, ["869e098ffd768989"], 0.3, 7.02, 4.6, color=IVORY, prefix="Ảnh minh họa")
    anim.append((core, "click", "fade"))
    transition(s, "fade")
    animate(s, anim)
    return s


def s14_closing(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, NAVY)
    photo(s.shapes, "af8683a47a215a63", 0, 0, W, H, focus=(0.5, 0.55), dark=0.2, sat=1.0,
          grads=[("left", 0.85, 0.6), ("bottom", 0.85, 0.55)])
    k = chrome(s, "THÔNG ĐIỆP", num_color=GOLD)
    t = text(s.shapes, 0.6, 1.15, 8.5, 2.2, [[("Không ai là", {})], [("món hàng.", dict(color=GOLD))]],
             size=60, font=HEAD, line_spacing=0.95, name="!!title")
    sub = text(s.shapes, 0.6, 3.4, 9.5, 0.5,
               "Tỉnh táo trước mỗi lời mời — và lên tiếng khi thấy điều bất thường.", size=19, font=HEAD2)
    acts = ["Kiểm chứng mọi lời mời việc làm trước khi trả lời",
            "Kể với người lớn tin cậy khi thấy điều bất thường",
            "Lưu số 111 và chia sẻ bài học này với bạn bè"]
    anim = [(k, "auto", "fade"), (t, "with", "fade"), (sub, "click", "fade")]
    for i, a in enumerate(acts):
        x = 0.6 + i * 4.1
        g = group(s, f"Hành động {i + 1}")
        rect(g.shapes, x, 4.6, 3.8, 1.3, NAVY, alpha=0.62)
        text(g.shapes, x + 0.25, 4.72, 0.6, 0.6, str(i + 1), size=28, font=HEAD, color=GOLD)
        text(g.shapes, x + 0.85, 4.82, 2.8, 1.0, a, size=16, line_spacing=1.08)
        anim.append((g, "click", "fade"))
    thanks = text(s.shapes, 0.6, 6.55, 7, 0.4, "Cảm ơn thầy cô và các bạn đã lắng nghe.", size=14, color=STEEL)
    credit(s, ["af8683a47a215a63"], 6.0, 7.02, 5.6, align=PP_ALIGN.RIGHT, prefix="Ảnh minh họa")
    anim.append((thanks, "click", "fade"))
    transition(s, "push", 800)
    animate(s, anim)
    return s


REFERENCES = [
    ("Pháp luật, chính sách & tổ chức quốc tế", [
        "Văn phòng Quốc hội (2025). Văn bản hợp nhất 138/VBHN-VPQH: Luật Phòng, chống mua bán người (Luật 53/2024/QH15). "
        "https://datafiles.chinhphu.vn/cpp/files/vbpq/2025/9/138-vbhn-vpqh.pdf",
        "Chính phủ (2025). Nghị định 162/2025/NĐ-CP quy định chi tiết Luật Phòng, chống mua bán người. "
        "https://vbpl.vn/TW/Lists/vbpq/Attachments/179697/162_2025_ND-CP_663357.doc",
        "Báo Văn hóa (2026). 9 mục tiêu, 8 nhiệm vụ phòng, chống mua bán người (QĐ 33/QĐ-TTg). "
        "https://baovanhoa.vn/doi-song/9-muc-tieu-8-nhiem-vu-phong-chong-mua-ban-nguoi-195997.html",
        "UNODC. Human trafficking: act, means, purpose (Nghị định thư Palermo, Điều 3). "
        "https://unodc.org/conig/en/stories/human-trafficking.html",
        "Vietnamnet. 30/7 là Ngày toàn dân phòng, chống mua bán người (QĐ 793/QĐ-TTg). "
        "https://vietnamnet.vn/307-la-ngay-toan-dan-phong-chong-mua-ban-nguoi-315614.html",
        "VnEconomy (2025). A national telephone hotline for human trafficking prevention launched. "
        "https://en.vneconomy.vn/a-national-telephone-hotline-for-human-trafficking-prevention-launched.htm",
    ]),
    ("Số liệu & báo chí", [
        "Vietnamnet (2025). Phòng ngừa nạn mua bán người ngay từ tuyến đầu (số liệu 2024 của Bộ Công an). "
        "https://vietnamnet.vn/phong-ngua-nan-mua-ban-nguoi-ngay-tu-tuyen-dau-2479028.html",
        "An ninh Thủ đô (2025). Bộ Công an điều tra, xử lý 163 vụ mua bán người. "
        "https://anninhthudo.vn/bo-cong-an-dieu-tra-xu-ly-163-vu-mua-ban-nguoi-post603866.antd",
        "Báo Lào Cai (2025). Việt Nam siết chặt trận tuyến phòng chống tội phạm mua bán người. "
        "https://baolaocai.vn/viet-nam-siet-chat-tran-tuyen-phong-chong-toi-pham-mua-ban-nguoi-post649993.html",
        "PLO (2024). Quý III-2024, có hơn 100 nam giới là nạn nhân của tội phạm mua bán người. "
        "https://plo.vn/quy-iii-2024-co-hon-100-nam-gioi-la-nan-nhan-cua-toi-pham-mua-ban-nguoi-post816519.html",
        "Bộ Công an (2025). Phát hiện, điều tra 50 vụ mua bán người trong đợt cao điểm. "
        "https://www.mps.gov.vn/bai-viet/luc-luong-cong-an-phat-hien-dieu-tra-50-vu-mua-ban-nguoi-trong-dot-cao-diem-"
        "tan-cong-tran-ap-toi-pham-mua-ban-nguoi-tren-pham-vi-toan-quoc-1762498361",
        "Tuổi Trẻ (2025). Công an TP.HCM triệt phá đường dây bán người cho “trung tâm lừa đảo trực tuyến” ở Campuchia. "
        "https://tuoitre.vn/cong-an-tp-hcm-triet-pha-duong-day-ban-nguoi-cho-trung-tam-lua-dao-truc-tuyen-o-campuchia-20250107115109926.htm",
        "SGGP. Nhiều người bị lừa “việc nhẹ, lương cao”, Bộ Công an phát cảnh báo. "
        "https://www.sggp.org.vn/nhieu-nguoi-bi-lua-viec-nhe-luong-cao-bo-cong-an-phat-canh-bao-post642677.html",
    ]),
]


def s15_references(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, CHARCOAL)
    k = chrome(s, "TÀI LIỆU THAM KHẢO")
    t = title(s, 0.6, 0.85, 8, 0.7, "Tài liệu tham khảo", size=30)
    n = 1
    for col, (head, items) in enumerate(REFERENCES):
        x, w = (0.6, 5.9) if col == 0 else (6.85, 5.9)
        paras = [{"text": head.upper(), "size": 11, "bold": True, "color": GOLD, "spacing": 150, "space_after": 6}]
        for it in items:
            label, url = it.rsplit(" https://", 1)
            paras.append({"runs": [(f"[{n}] ", dict(color=CRIMSON, bold=True)), (label + " ", {}),
                                   ("https://" + url, dict(color=STEEL))], "space_after": 5})
            n += 1
        text(s.shapes, x, 1.7, w, 4.6, paras, size=10, color="D9D4CA", line_spacing=1.05)
    text(s.shapes, 0.6, 5.35, 5.9, 1.5,
                  [[("ẢNH: ", dict(color=GOLD, bold=True)),
                    ("Open Images Dataset V4 (Google), ảnh gốc trên Flickr, giấy phép CC BY 2.0 — "
                     + ", ".join(sorted({m[0] for m in PHOTO_META.values()}))
                     + ". Chi tiết từng ảnh: Nguon_anh.txt. Font: Be Vietnam Pro (SIL OFL).", {})]],
                  size=10, color="D9D4CA", line_spacing=1.05)
    transition(s, "fade")
    animate(s, [(k, "auto", "fade"), (t, "with", "fade")])
    return s


# ---------------------------------------------------------------- lời thuyết trình

NOTES = [
    # 1
    "(~30 giây) Xin chào thầy cô và các bạn. Bức ảnh phía sau là Hà Nội về đêm — rực rỡ, chuyển động không ngừng. "
    "Nhưng trong chính nhịp sống ấy vẫn có những cái bẫy ít ai nhìn thấy. Hôm nay nhóm mình trình bày về nạn buôn "
    "người ở Hà Nội: nó đang diễn ra thế nào, vì sao xảy ra, gây hậu quả gì, và quan trọng nhất — chúng ta có thể làm gì.",
    # 2
    "(~45 giây) Hãy tưởng tượng tối nay bạn nhận được tin nhắn này. Lương 15 đến 20 triệu, việc nhẹ, bao ăn ở, "
    "không cần kinh nghiệm. Nghe rất hấp dẫn, đúng không? Mình muốn mỗi bạn tự trả lời thầm trong đầu: mình có nhắn "
    "lại không? Đây là tình huống do nhóm tự xây dựng, không phải vụ việc có thật — nhưng nó được mô phỏng theo đúng "
    "những dấu hiệu mà cơ quan chức năng đã cảnh báo. Cuối bài, chúng ta sẽ quay lại tin nhắn này.",
    # 3
    "(~50 giây) Vậy buôn người là gì? Theo Nghị định thư Palermo của Liên Hợp Quốc và Luật Phòng, chống mua bán người "
    "năm 2024, có thể hiểu qua ba yếu tố. Thứ nhất là hành vi — tuyển mộ, vận chuyển, chứa chấp một người. Thứ hai là "
    "thủ đoạn — lừa dối, đe dọa, lợi dụng hoàn cảnh khó khăn. Thứ ba là mục đích — bóc lột. Riêng với trẻ em, chỉ cần "
    "hành vi và mục đích bóc lột là đủ. Ở Việt Nam, mua bán người dưới 16 tuổi còn bị xử lý theo một điều luật riêng "
    "với hình phạt nặng hơn. Và cần phân biệt: đi làm xa tự nguyện, có hợp đồng, tự giữ giấy tờ, được nghỉ việc khi "
    "muốn thì không phải là buôn người.",
    # 4
    "(~45 giây) Vì sao Hà Nội cần đặc biệt quan tâm? Không phải vì Hà Nội 'nhiều nhất' — nhóm chưa tìm thấy dữ liệu "
    "nào chứng minh điều đó. Lý do là mức độ tiếp xúc với nguy cơ: đây là trung tâm việc làm, học tập, là đầu mối "
    "giao thông với nhà ga, bến xe, sân bay; người trẻ ở đây dùng mạng xã hội rất nhiều; và nhiều bạn mới lên thành "
    "phố, chưa có kinh nghiệm, chưa có người thân bên cạnh để hỏi ý kiến.",
    # 5
    "(~60 giây) Về số liệu: nhóm đã tìm thống kê riêng của Hà Nội nhưng chưa thấy số liệu công khai đầy đủ của "
    "Công an thành phố cho 2024–2025, nên dùng số liệu toàn quốc của Bộ Công an làm bối cảnh. Năm 2024, cả nước điều "
    "tra, xử lý 163 vụ, 455 đối tượng và xác định 500 nạn nhân; trong đó khởi tố mới 91 vụ với 336 nạn nhân. "
    "Biểu đồ so sánh tổng số với số khởi tố mới. Quý I/2025, khoảng 65% nạn nhân là nữ và hơn 44% dưới 16 tuổi. "
    "Nhưng nam giới cũng không an toàn: riêng quý III/2024 đã có hơn 100 nạn nhân là nam. Lưu ý, các con số này đến "
    "từ những đợt báo cáo khác nhau nên không cộng dồn với nhau.",
    # 6
    "(~50 giây) Kẻ xấu thường không bắt cóc ngay mà bắt đầu bằng một lời mời. Năm thủ đoạn trên slide đều đã được "
    "Bộ Công an và báo chí chính thống ghi nhận: việc nhẹ lương cao; tuyển dụng trực tuyến mập mờ; lợi dụng tình cảm, "
    "hẹn hò qua mạng; đưa đi làm xa hoặc xuất cảnh trái phép — đã có những vụ dụ người sang Campuchia làm cho các cơ "
    "sở lừa đảo trực tuyến; và cuối cùng là khống chế, giữ giấy tờ, đòi tiền chuộc.",
    # 7
    "(~60 giây) Với học sinh chúng ta, cánh cửa nguy hiểm nhất là chiếc điện thoại. Đây là đoạn chat mô phỏng — "
    "nội dung hoàn toàn giả định. Bên phải là sáu dấu hiệu cảnh báo chung giúp các bạn tự nhận diện: lương cao bất "
    "thường, không hợp đồng, bắt giữ bí mật, đòi ảnh căn cước hay tiền cọc, tài khoản mới lập rồi kéo sang Telegram, "
    "và hứa bao vé xe đưa đi xa. Chỉ cần thấy một dấu hiệu, hãy dừng lại và hỏi người lớn.",
    # 8
    "(~50 giây) Vì sao vẫn có người rơi vào bẫy? Không có một nguyên nhân duy nhất. Áp lực kinh tế, thiếu kỹ năng tự "
    "bảo vệ, công nghệ giúp kẻ xấu dễ tạo lòng tin, và việc phát hiện, báo tin còn khó khăn — các yếu tố này cộng hưởng "
    "với nhau. Kẻ phạm tội nhắm đúng vào những khoảng trống đó. Điều mình muốn nhấn mạnh: lỗi thuộc về kẻ phạm tội, "
    "không phải nạn nhân.",
    # 9
    "(~45 giây) Hậu quả không chỉ là những gì nhìn thấy được. Nạn nhân có thể bị tổn hại sức khỏe, sang chấn tâm lý, "
    "mất tự do và tương lai. Gia đình sống trong lo âu, có khi gánh cả nợ nần, tiền chuộc. Xã hội mất an toàn và mất "
    "niềm tin. Vì vậy khi một người trở về, điều họ cần là được lắng nghe và hỗ trợ, không phải bị phán xét.",
    # 10
    "(~50 giây) Quay lại một tình huống giả định cụ thể: một bạn 18 tuổi được người quen trên mạng mời làm ở tỉnh "
    "khác, lương 25 triệu, không hợp đồng, phải giữ bí mật và đi ngay tối nay. Dấu hiệu nguy hiểm đã quá rõ. Lựa chọn "
    "an toàn là từ chối hoặc hoãn lại, không gửi giấy tờ, không lên xe. Người có thể giúp là bố mẹ, thầy cô, công an "
    "phường, hoặc gọi 111. Và cách phòng tránh lâu dài là luôn kiểm tra công ty, đọc kỹ hợp đồng, báo gia đình nơi làm việc.",
    # 11
    "(~50 giây) Giải pháp cần bốn lớp bảo vệ cùng hoạt động. Bản thân mỗi người kiểm chứng thông tin và bảo mật giấy "
    "tờ. Gia đình lắng nghe, không chê trách. Nhà trường và cộng đồng lồng ghép giáo dục pháp luật, kỹ năng số và để "
    "ý những dấu hiệu lạ. Cơ quan chức năng kiểm soát dịch vụ việc làm, xử lý nghiêm các đường dây và hỗ trợ nạn nhân. "
    "Khi một lớp yếu đi, các lớp còn lại sẽ giúp phát hiện kịp thời.",
    # 12
    "(~50 giây) Nếu nghi ngờ, hãy nhớ năm bước: dừng lại, chia sẻ, lưu bằng chứng nếu an toàn, báo cơ quan chức năng, "
    "và luôn ưu tiên an toàn của bản thân. Hai số cần nhớ: 111 là Tổng đài quốc gia về phòng, chống mua bán người, "
    "miễn phí, hoạt động 24/7, khai trương tháng 9/2025 theo Nghị định 162; còn 113 là số cảnh sát khẩn cấp khi đang "
    "gặp nguy hiểm trực tiếp.",
    # 13
    "(~50 giây) Về pháp luật: ngày 28/11/2024, Quốc hội thông qua Luật Phòng, chống mua bán người số 53, có hiệu lực "
    "từ 01/7/2025, thay thế luật năm 2011; Nghị định 162/2025 hướng dẫn thi hành. Điểm cốt lõi là lấy nạn nhân làm "
    "trung tâm và mở rộng những người được bảo vệ. Người phạm tội bị xử lý hình sự theo Điều 150 và 151 Bộ luật Hình "
    "sự. Đầu năm 2026, Thủ tướng ban hành Quyết định 33 về chương trình phòng, chống mua bán người giai đoạn "
    "2026–2030, và Hà Nội cũng đã có kế hoạch triển khai trên địa bàn.",
    # 14
    "(~40 giây) Quay lại tin nhắn ở đầu bài: bây giờ, bạn sẽ trả lời thế nào? Thông điệp của nhóm rất đơn giản: không "
    "ai là món hàng. Ba việc chúng ta có thể làm ngay từ hôm nay: kiểm chứng mọi lời mời việc làm, kể với người lớn "
    "khi thấy bất thường, và lưu số 111 rồi chia sẻ bài học này với bạn bè. Ngày 30/7 hằng năm là Ngày toàn dân "
    "phòng, chống mua bán người — nhưng tỉnh táo thì cần mỗi ngày. Cảm ơn thầy cô và các bạn đã lắng nghe.",
    # 15
    "(Không cần trình bày chi tiết) Đây là danh sách nguồn tham khảo về văn bản pháp luật, số liệu và hình ảnh. "
    "Thông tin chi tiết từng ảnh và giấy phép nằm trong file Nguon_anh.txt. Nhóm sẵn sàng trả lời câu hỏi.",
]


def write_docx(path):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = DocPt(12)
    doc.add_heading("Lời thuyết trình — Nạn buôn người ở Hà Nội", level=1)
    doc.add_paragraph("Tổng thời lượng dự kiến: khoảng 10–11 phút. Thời gian gợi ý ghi ở đầu mỗi slide. "
                      "Lời thuyết trình cũng đã được nhúng vào phần Speaker Notes của file PowerPoint.")
    for i, note in enumerate(NOTES, 1):
        doc.add_heading(f"Slide {i}", level=2)
        doc.add_paragraph(note)
    doc.save(path)


def write_photo_credits(path):
    lines = ["NGUỒN ẢNH — Bài thuyết trình “Nạn buôn người ở Hà Nội”",
             "",
             "Tất cả ảnh được tải từ Open Images Dataset V4 (Google) — bản sao lưu trữ trên Amazon S3:",
             "  https://storage.googleapis.com/openimages/web/index.html",
             "Ảnh gốc được đăng trên Flickr; mỗi ảnh mang giấy phép Creative Commons Attribution 2.0 (CC BY 2.0):",
             f"  {LICENSE_URL}",
             "Yêu cầu ghi công: ghi tên tác giả + giấy phép (đã đặt trên từng slide và ở slide 15).",
             "Các ảnh có người chỉ mang tính MINH HỌA, không liên quan tới bất kỳ vụ việc mua bán người nào.",
             "Ảnh đã được cắt khung, chỉnh màu và phủ lớp tối để tăng độ tương phản (bản gốc: assets/photos/).",
             ""]
    for pid, (author, title_, url, slide, desc) in sorted(PHOTO_META.items(), key=lambda kv: kv[1][3]):
        lines += [f"Slide {slide:02d} — {desc}",
                  f"  Tiêu đề gốc : {title_}",
                  f"  Tác giả     : {author}",
                  f"  Trang nguồn : {url}",
                  f"  Giấy phép   : {LICENSE} ({LICENSE_URL})",
                  f"  Open Images : ImageID {pid} — https://open-images-dataset.s3.amazonaws.com/train/{pid}.jpg",
                  f"  File        : assets/photos/{pid}.jpg",
                  ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def write_references(path):
    lines = ["TÀI LIỆU THAM KHẢO — Bài thuyết trình “Nạn buôn người ở Hà Nội”", ""]
    n = 1
    for head, items in REFERENCES:
        lines += [head.upper(), ""]
        for it in items:
            lines.append(f"[{n}] {it}")
            n += 1
        lines.append("")
    lines += [
        "GHI CHÚ VỀ DỮ LIỆU",
        "- Số liệu năm 2024 và quý I/2025 là số liệu TOÀN QUỐC do Bộ Công an công bố, được báo chí dẫn lại.",
        "- Chưa tìm thấy thống kê công khai riêng của Công an TP. Hà Nội cho giai đoạn 2024–2025.",
        "- Các con số thuộc những đợt báo cáo khác nhau (năm, quý, đợt cao điểm) nên không cộng dồn.",
        "- Tổng đài 111: Tổng đài quốc gia về phòng, chống mua bán người (dùng chung số với Tổng đài quốc gia",
        "  bảo vệ trẻ em), theo Nghị định 162/2025/NĐ-CP; miễn phí, 24/7.",
        "- Luật 53/2024/QH15 được sửa đổi, bổ sung bởi Luật 86/2025/QH15 (sửa đổi Bộ luật Hình sự), cùng có",
        "  hiệu lực từ 01/7/2025 — theo VBHN 138/VBHN-VPQH.",
        "",
        "FONT",
        "- Be Vietnam Pro — SIL Open Font License, https://fonts.google.com/specimen/Be+Vietnam+Pro",
        "  (file font kèm theo trong assets/fonts/).",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    props = prs.core_properties
    props.title = "Nạn buôn người ở Hà Nội — Thực trạng, nguyên nhân, hậu quả và giải pháp"
    props.subject = "Bài thuyết trình cho học sinh THPT"
    props.author = ""

    makers = [s01_cover, s02_question, s03_definition, s04_hanoi, s05_data, s06_tactics, s07_social,
              s08_causes, s09_consequences, s10_case, s11_solutions, s12_action, s13_law, s14_closing,
              s15_references]
    for make, note in zip(makers, NOTES):
        slide = make(prs)
        slide.notes_slide.notes_text_frame.text = note

    prs.save(OUT_PPTX)
    write_docx(ROOT / "Loi_thuyet_trinh.docx")
    write_photo_credits(ROOT / "Nguon_anh.txt")
    write_references(ROOT / "Tai_lieu_tham_khao.txt")
    print(f"Đã tạo {OUT_PPTX.name} ({len(prs.slides)} slide)")


if __name__ == "__main__":
    build()
