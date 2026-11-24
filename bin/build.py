#!/usr/bin/env python3
"""Build the History of Economic Thought static site from course material kept elsewhere.

Collects:
  * lecture decks (.pptx)        -> lectures/<slug>.html + lectures/<slug>/ images
  * handouts (.docx, .rtf)       -> handouts/<slug>.html (via pandoc)
  * spreadsheets (.xlsx)         -> spreadsheets/
  * quizzes and the final exam   -> quizzes/<slug>.html (via pandoc)
and writes index.html, organised around the course schedule.

Usage: python3 bin/build.py [--src DIR] [--out DIR]
Needs: python-pptx (pip install python-pptx) and pandoc.
"""

import argparse
import hashlib
import html
import re
import shutil
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

DEFAULT_SRC = Path.home() / "Documents/CoursesTaught/HistoryOfThought"

# ---------------------------------------------------------------- course data
#
# Everything about *what* goes on the site lives in these tables.
#
# Deliberately left out of the source folder:
#   BaisdenGrade.docx, ClassList.xlsx    student records
#   QuizTemplate.docx                    blank template
#   Quiz1.docx                           an earlier draft of Final.docx (it is headed "Final")
#   Syllabus-02.docx                     2010 admin details; its schedule is reproduced in TOPICS
#   Aristotle's Economic Thought.docx    empty apart from its title
#   ClassicalConsensus/ClassicalConsensus.pptx   older version of ClassicalConsensus-03
#   Keynes/cbm.pptx, kandh.pptx, circflow.pptx   Roger Garrison's slides (Time and Money)
#   Marx/SocialClass.html, SocialClass.xls       a saved Wikipedia page
#   *.pdf, AdamSmithTOC.txt              exports of files already converted
#   loose .jpg/.png files                already embedded in the decks

# Lecture decks: slug -> source path relative to SRC.
DECKS = {
    "adam-smith": "AdamSmith/AdamSmith.pptx",
    "classical-consensus": "ClassicalConsensus/ClassicalConsensus-03.pptx",
    "says-law": "SaysLaw/SaysLaw.pptx",
    "bentham-comte": "Bentham/Bentham.pptx",
    "marx": "Marx/Marx.pptx",
    "engels": "Marx/Engels.pptx",
    "menger": "Menger/Menger.pptx",
    "keynes": "Keynes/Keynes.pptx",
    "science-and-economics": "ScienceAndEconomics9.pptx",
}

# Handouts: slug -> (source path, title).
HANDOUTS = {
    "wealth-of-nations-contents": ("AdamSmith/AdamSmithTOC.docx",
                                   "The Wealth of Nations: chapter outline"),
    "surplus-value-and-profit": ("Marx/SurplusValueAndProfit.docx",
                                 "Surplus value and the rate of profit"),
    "origin-of-money": ("Menger/OriginOfMoney.docx", "What makes a good money?"),
    "rats-and-traps": ("Menger/Rich.docx", "Rich and Helena: an exchange problem"),
    "menger-on-communism": ("Menger/Communism.rtf", "Menger on goods held in common"),
    "notes-on-hayek": ("NotesOnHayek.docx", "Notes on Hayek"),
}

# Spreadsheets: slug -> (source path, title, description).
SPREADSHEETS = {
    "deer-beaver": ("AdamSmith/DeerBeaver.xlsx", "Deer and beaver",
                    "Smith's hunting economy: a production possibilities frontier "
                    "and an indifference curve for deer and beaver."),
    "menger-value-chart": ("Menger/MengerValueChart.xlsx", "Menger's value table",
                           "Menger's table of declining marginal importance for ten "
                           "classes of want."),
    "horses-cows-1": ("Menger/HorsesCows1.xlsx", "Horses and cows (1)",
                      "The declining value of each additional horse and cow, for Menger's "
                      "isolated exchange between two farmers."),
    "horses-cows-2": ("Menger/HorsesCows2.xlsx", "Horses and cows (2)",
                      "A second version of the horse and cow value table."),
}

# Quizzes and the final, in order. Titles come from each document's first line.
QUIZZES = ["Quiz2.docx", "Quiz3.docx", "Quiz4.docx", "Final.docx"]

# The course schedule (from the Fall 2010 syllabus). Each topic lists the
# decks, handouts and spreadsheets that go with it; topics with none are
# still shown so the shape of the course is clear.
TOPICS = [
    {"title": "Aristotle: economics as household management"},
    {"title": "Hebrew and medieval Christian views on economics"},
    {"title": "Bacon on society as a productive enterprise; the rise of mercantilism"},
    {"title": "Smith responds to the mercantilists: economics as the science of national wealth",
     "reading": "Barber, ch. I",
     "decks": ["adam-smith"], "handouts": ["wealth-of-nations-contents"],
     "sheets": ["deer-beaver"]},
    {"title": "The classical consensus: Malthus, Ricardo and Mill",
     "reading": "Barber, chs. II–IV",
     "decks": ["classical-consensus", "says-law"]},
    {"title": "Bentham, Comte and the science of society",
     "decks": ["bentham-comte"]},
    {"title": "Responses to classical economics: Marx and Veblen",
     "reading": "Barber, Part II intro and ch. V",
     "decks": ["marx", "engels"], "handouts": ["surplus-value-and-profit"]},
    {"title": "The marginal revolution: Menger, Jevons, Walras and economics as the science of choice",
     "reading": "Barber, ch. VII",
     "decks": ["menger"],
     "handouts": ["origin-of-money", "rats-and-traps", "menger-on-communism"],
     "sheets": ["menger-value-chart", "horses-cows-1", "horses-cows-2"]},
    {"title": "Marshall and the neoclassical synthesis",
     "reading": "Barber, ch. VI"},
    {"title": "Keynes: the economist as social therapist",
     "reading": "Barber, Part IV intro and ch. VIII",
     "decks": ["keynes"]},
    {"title": "Mises and Hayek: economic activity as a discovery process",
     "reading": "Hayek, “Economics and Knowledge”; Mises, “A First Analysis of the Category of Action”",
     "decks": ["science-and-economics"], "handouts": ["notes-on-hayek"]},
    {"title": "Post-war neoclassical economics: economics as mathematical modelling",
     "reading": "Friedman, “The Methodology of Positive Economics”"},
    {"title": "Heterodox approaches: Austrian, critical realist, evolutionary and post-Keynesian"},
]

TEXTBOOK = "William J. Barber, <cite>The History of Economic Thought</cite>"

# ---------------------------------------------------------------- site chrome

SITE_TITLE = "History of Economic Thought"
FONTS_URL = ("https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,600"
             "&family=EB+Garamond:ital,wght@0,400;0,600;1,400"
             "&family=IBM+Plex+Sans:wght@500;600&display=swap")
BACKDROP_CREDIT = ("Background: Ambrogio Lorenzetti, <cite>Effects of Good Government in the City</cite> "
                   "(1338–39), Palazzo Pubblico, Siena. Public domain, via "
                   '<a href="https://commons.wikimedia.org/wiki/File:Ambrogio_Lorenzetti_-_Effects_of_Good_Government_in_the_city_-_Google_Art_Project.jpg">'
                   "Wikimedia Commons</a>.")
FOOTER = "From Aristotle to Hayek: lectures, handouts and exams from a course on the history of economic ideas."
N_CHART_SERIES = 4  # chart series colours are --series-0 .. --series-3 in style.css

WEB_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "svg", "webp"}
CONVERT_WITH_SIPS = {"tif", "tiff", "bmp"}

# Set in main() from style.css's hash, so browsers refetch it after changes.
CSS_VERSION = ""

warnings = []


def warn(msg):
    warnings.append(msg)


def esc(text):
    return html.escape(text, quote=True)


def page(title, body, root="", body_class="", hero=None):
    """Wrap body in the site chrome.

    root:  relative path from this page to the site root ("" or "../").
    hero:  optional (kicker, heading, subtitle) for the page header.
    """
    hero_html = ""
    if hero:
        kicker, heading, sub = hero
        hero_html = f"""<header class="hero">
  {f'<p class="kicker">{kicker}</p>' if kicker else ''}
  <h1>{heading}</h1>
  {f'<p class="sub">{sub}</p>' if sub else ''}
</header>"""
    doc_title = SITE_TITLE if title == SITE_TITLE else f"{title} · {SITE_TITLE}"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{esc(doc_title)}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="{FONTS_URL}">
    <link rel="stylesheet" href="{root}style.css?v={CSS_VERSION}">
</head>
<body class="{body_class}">
<nav class="site-nav">
  <a class="brand" href="{root}index.html">{SITE_TITLE}</a>
  <span class="links">
    <a href="{root}index.html#schedule">Schedule</a>
    <a href="{root}index.html#lectures">Lectures</a>
    <a href="{root}index.html#handouts">Handouts</a>
    <a href="{root}index.html#quizzes">Exams</a>
  </span>
</nav>
<div class="sheet">
{hero_html}
<main>
{body}
</main>
</div>
<footer class="site-footer">
  <p>{FOOTER}</p>
  <p class="credit">{BACKDROP_CREDIT}</p>
</footer>
</body>
</html>
"""


def slugify(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


# ---------------------------------------------------------------- slides

def run_html(run):
    text = esc(run.text).replace("\x0b", "<br>")
    if not text:
        return ""
    font = run.font
    if font.bold:
        text = f"<strong>{text}</strong>"
    if font.italic:
        text = f"<em>{text}</em>"
    addr = run.hyperlink.address if run.hyperlink else None
    if addr:
        text = f'<a href="{esc(addr)}">{text}</a>'
    return text


def text_frame_html(tf):
    """Render a text frame as nested lists keyed on paragraph level."""
    paras = [(p.level, "".join(run_html(r) for r in p.runs)) for p in tf.paragraphs]
    paras = [(lvl, t) for lvl, t in paras if t.strip()]
    if not paras:
        return ""
    out, depth = [], -1
    for lvl, text in paras:
        while depth < lvl:
            out.append("<ul>")
            depth += 1
        while depth > lvl:
            out.append("</li></ul>")
            depth -= 1
        if out[-1] != "<ul>":
            out.append("</li>")
        out.append(f"<li>{text}")
    out.append("</li>")
    while depth >= 0:
        out.append("</ul>")
        depth -= 1
    return "\n".join(out)


def table_html(tbl):
    rows = []
    for r_i, row in enumerate(tbl.rows):
        tag = "th" if r_i == 0 else "td"
        cells = "".join(
            f"<{tag}>{esc(cell.text).replace(chr(10), '<br>')}</{tag}>"
            for cell in row.cells)
        rows.append(f"<tr>{cells}</tr>")
    return "<table>\n" + "\n".join(rows) + "\n</table>"


C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def series_xy(series, index):
    """Return [(x, y)] for a chart series. Category charts use the point index as x."""
    ys = list(series.values)
    xval = series._element.find(f"{C_NS}xVal")  # XY scatter: explicit x values
    if xval is None:
        return [(i, y) for i, y in enumerate(ys)]
    xs = {int(pt.get("idx")): float(pt.find(f"{C_NS}v").text)
          for pt in xval.iter(f"{C_NS}pt")}
    return [(xs.get(i, i), y) for i, y in enumerate(ys)]


def chart_html(chart):
    """Draw a line or XY chart as inline SVG; colours come from style.css."""
    try:
        series = [(s.name, series_xy(s, k)) for k, s in enumerate(chart.plots[0].series)]
    except Exception as e:  # noqa: BLE001
        warn(f"chart could not be read: {e}")
        return ""
    pts = [(x, y) for _, ps in series for x, y in ps if y is not None]
    if not pts:
        return ""
    xlo, xhi = min(p[0] for p in pts), max(p[0] for p in pts)
    ylo, yhi = min([p[1] for p in pts] + [0]), max(p[1] for p in pts)
    xhi = xhi if xhi != xlo else xlo + 1
    yhi = yhi if yhi != ylo else ylo + 1
    W, H, L, R, T, B = 640, 320, 48, 16, 16, 32

    def sx(v):
        return L + (W - L - R) * (v - xlo) / (xhi - xlo)

    def sy(v):
        return T + (H - T - B) * (1 - (v - ylo) / (yhi - ylo))

    parts = [f'<svg class="chart" viewBox="0 0 {W} {H}" role="img" xmlns="http://www.w3.org/2000/svg">']
    for v in (ylo, (ylo + yhi) / 2, yhi):
        parts.append(f'<line x1="{L}" x2="{W - R}" y1="{sy(v):.1f}" y2="{sy(v):.1f}" class="grid"/>'
                     f'<text x="{L - 6}" y="{sy(v) + 4:.1f}" text-anchor="end">{v:,.4g}</text>')
    for v in (xlo, (xlo + xhi) / 2, xhi):
        parts.append(f'<text x="{sx(v):.1f}" y="{H - B + 18}" text-anchor="middle">{v:,.4g}</text>')
    for k, (_, ps) in enumerate(series):
        line = " ".join(f"{sx(x):.1f},{sy(y):.1f}" for x, y in ps if y is not None)
        parts.append(f'<polyline class="series-{k % N_CHART_SERIES}" points="{line}"/>')
    parts.append("</svg>")
    legend = '<p class="legend">' + " ".join(
        f'<span class="series-{k % N_CHART_SERIES}">&#9644; {esc(nm)}</span>'
        for k, (nm, _) in enumerate(series) if nm) + "</p>"
    return "\n".join(parts) + legend


class Deck:
    def __init__(self, src, slug, outdir):
        self.src, self.slug = src, slug
        self.imgdir = outdir / slug
        self.n_images = 0

    def image_html(self, shape):
        img = shape.image
        ext = img.ext.lower()
        self.n_images += 1
        name = f"img{self.n_images:03d}"
        self.imgdir.mkdir(parents=True, exist_ok=True)
        if ext in CONVERT_WITH_SIPS:
            raw = self.imgdir / f"{name}.{ext}"
            raw.write_bytes(img.blob)
            dest = self.imgdir / f"{name}.png"
            subprocess.run(["sips", "-s", "format", "png", str(raw), "--out", str(dest)],
                           check=True, capture_output=True)
            raw.unlink()
        elif ext in WEB_IMAGE_EXTS:
            dest = self.imgdir / f"{name}.{ext}"
            dest.write_bytes(img.blob)
        else:
            warn(f"{self.src.name}: skipped .{ext} image (not viewable in browsers)")
            return ""
        alt = esc(getattr(shape, "name", "") or "")
        return f'<img src="{self.slug}/{dest.name}" alt="{alt}" loading="lazy">'

    def shape_html(self, shape, title_id):
        if shape.shape_id == title_id:
            return ""
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return "\n".join(filter(None, (self.shape_html(s, title_id) for s in shape.shapes)))
        parts = []
        if hasattr(shape, "image"):
            try:
                parts.append(self.image_html(shape))
            except ValueError:
                pass  # placeholder with no picture in it
        if getattr(shape, "has_table", False) and shape.has_table:
            parts.append(table_html(shape.table))
        if getattr(shape, "has_chart", False) and shape.has_chart:
            parts.append(chart_html(shape.chart))
        if shape.has_text_frame:
            parts.append(text_frame_html(shape.text_frame))
        return "\n".join(filter(None, parts))

    def build(self):
        """Return (deck title, [(slide num, slide title)], slides html)."""
        prs = Presentation(self.src)
        if self.imgdir.exists():
            shutil.rmtree(self.imgdir)
        deck_title, toc, sections = None, [], []
        for num, slide in enumerate(prs.slides, 1):
            title_shape = slide.shapes.title
            if title_shape is not None and not title_shape.has_text_frame:
                title_shape = None
            title = title_shape.text_frame.text.strip() if title_shape is not None else ""
            title_id = title_shape.shape_id if title_shape is not None and title else None
            body = "\n".join(filter(None, (self.shape_html(s, title_id) for s in slide.shapes)))
            title = re.sub(r"\s+", " ", title) or f"Slide {num}"
            if deck_title is None:
                deck_title = title
            toc.append((num, title))
            notes = ""
            if slide.has_notes_slide:
                n = slide.notes_slide.notes_text_frame
                if n is not None and n.text.strip():
                    notes = f'<aside class="notes">{text_frame_html(n)}</aside>'
            sections.append(
                f'<section class="slide" id="s{num}">\n'
                f'<h2><a class="num" href="#s{num}">{num}</a> {esc(title)}</h2>\n'
                f'{body}\n{notes}\n</section>')
        return deck_title, toc, "\n".join(sections)


# ---------------------------------------------------------------- documents

def pandoc_html(path, media_dir, link_prefix):
    """Convert a .docx/.rtf to an HTML fragment, extracting images to media_dir."""
    if path.suffix.lower() == ".rtf":  # pandoc 2.x cannot read RTF; macOS textutil can
        res = subprocess.run(["textutil", "-convert", "html", "-stdout", str(path)],
                             capture_output=True, text=True, check=True)
        res = subprocess.run(["pandoc", "-f", "html", "-t", "html5"], input=res.stdout,
                             capture_output=True, text=True, check=True)
        # Text pasted from a PDF keeps its line-end hyphens: "sat- isfaction".
        return re.sub(r"(?<=[a-z])- (?=[a-z])", "", res.stdout)
    res = subprocess.run(
        ["pandoc", str(path), "-t", "html5", f"--extract-media={media_dir}"],
        capture_output=True, text=True, check=True)
    return res.stdout.replace(f'src="{media_dir}/', f'src="{link_prefix}/')


def first_line(path):
    res = subprocess.run(["pandoc", str(path), "-t", "plain"],
                         capture_output=True, text=True, check=True)
    return next((ln.strip() for ln in res.stdout.splitlines() if ln.strip()), path.stem)


def fresh_dir(path):
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True)
    return path


# ---------------------------------------------------------------- topic lookups

def topic_of(kind, slug):
    """Return (index, topic) for the schedule topic that lists this item."""
    for i, t in enumerate(TOPICS):
        if slug in t.get(kind, []):
            return i, t
    return None, None


def topic_number(i):
    return f"Week {i + 1}"


def related_links(topic, root, built, exclude=None):
    """Chips linking to a topic's other decks, handouts and spreadsheets."""
    chips = []
    for kind, cls in (("decks", "deck"), ("handouts", "handout"), ("sheets", "spreadsheet")):
        for slug in topic.get(kind, []):
            if slug == exclude or slug not in built[kind]:
                continue
            href, title = built[kind][slug]
            dl = " download" if kind == "sheets" else ""
            chips.append(f'<a class="chip {cls}" href="{root}{href}"{dl}>{esc(title)}</a>')
    return chips


# ---------------------------------------------------------------- builders

def build_lectures(src, out):
    outdir = out / "lectures"
    outdir.mkdir(parents=True, exist_ok=True)
    built = {}
    for slug, rel in DECKS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing deck: {path}")
            continue
        deck = Deck(path, slug, outdir)
        title, toc, body = deck.build()
        built[slug] = (title, toc, body)
        print(f"  {rel}: {len(toc)} slides, {deck.n_images} images")
    return built


def write_lectures(out, decks, built):
    outdir = out / "lectures"
    order = list(decks)
    for i, slug in enumerate(order):
        title, toc, body = decks[slug]
        prev = order[i - 1] if i > 0 else None
        nxt = order[i + 1] if i + 1 < len(order) else None
        ti, topic = topic_of("decks", slug)
        chips = related_links(topic, "../", built, exclude=slug) if topic else []
        related = (f'<p class="related"><span>With this lecture:</span> {" ".join(chips)}</p>'
                   if chips else "")
        contents = "\n".join(f'<li><a href="#s{n}">{esc(t)}</a></li>' for n, t in toc)
        content = f"""{pager(prev, nxt, decks)}
{related}
<details class="contents">
<summary>Contents &middot; {len(toc)} slides</summary>
<ol>
{contents}
</ol>
</details>
{body}
{pager(prev, nxt, decks)}"""
        kicker = esc(topic_number(ti)) if topic else "Lecture"
        sub = esc(topic["title"]) if topic else f"{len(toc)} slides"
        (outdir / f"{slug}.html").write_text(
            page(title, content, root="../", body_class="lecture",
                 hero=(kicker, esc(title), sub)),
            encoding="utf-8")


def pager(prev, nxt, decks):
    def link(slug, cls, arrow):
        if not slug:
            return "<span></span>"
        return (f'<a class="{cls}" href="{slug}.html"><span class="dir">{arrow}</span>'
                f'<span class="t">{esc(decks[slug][0])}</span></a>')
    return (f'<nav class="pager">{link(prev, "prev", "&larr; Previous lecture")}'
            f'{link(nxt, "next", "Next lecture &rarr;")}</nav>')


def build_handouts(src, out, built):
    outdir = fresh_dir(out / "handouts")
    for slug, (rel, title) in HANDOUTS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing handout: {path}")
            continue
        body = pandoc_html(path, outdir / slug, slug)
        ti, topic = topic_of("handouts", slug)
        chips = related_links(topic, "../", built, exclude=slug) if topic else []
        related = (f'<p class="related"><span>Goes with:</span> {" ".join(chips)}</p>'
                   if chips else "")
        (outdir / f"{slug}.html").write_text(
            page(title, f'{related}\n<article class="paper">\n{body}\n</article>', root="../",
                 body_class="handout",
                 hero=(esc(topic_number(ti)) if topic else "Handout", esc(title),
                       esc(topic["title"]) if topic else "")),
            encoding="utf-8")
    print(f"  {len(list(outdir.glob('*.html')))} handouts")


def copy_spreadsheets(src, out):
    outdir = fresh_dir(out / "spreadsheets")
    built = {}
    for slug, (rel, title, _) in SPREADSHEETS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing spreadsheet: {path}")
            continue
        shutil.copy2(path, outdir / path.name)
        built[slug] = (f"spreadsheets/{path.name}", title)
    print(f"  {len(built)} spreadsheets")
    return built


def build_quizzes(src, out):
    outdir = fresh_dir(out / "quizzes")
    entries = []
    for fname in QUIZZES:
        path = src / fname
        if not path.exists():
            warn(f"missing quiz: {path}")
            continue
        slug = slugify(path.stem)
        # Headings read "History of Economic Thought: Quiz 2"; keep the part after the colon.
        title = first_line(path).split(":", 1)[-1].strip()
        body = pandoc_html(path, outdir / slug, slug)
        (outdir / f"{slug}.html").write_text(
            page(title, f'<article class="paper">\n{body}\n</article>', root="../",
                 body_class="quiz", hero=("Quizzes and exams", esc(title), "")),
            encoding="utf-8")
        entries.append((f"quizzes/{slug}.html", title))
    print(f"  {len(entries)} quizzes/finals")
    return entries


# ---------------------------------------------------------------- index

def build_index(out, decks, built, quizzes):
    rows = []
    for i, t in enumerate(TOPICS):
        chips = related_links(t, "", built)
        reading = f'<span class="reading">Reading: {esc(t["reading"])}</span>' if t.get("reading") else ""
        rows.append(
            f'<li class="topic{"" if chips else " bare"}">'
            f'<span class="week">{i + 1}</span>'
            f'<div><span class="title">{esc(t["title"])}</span>{reading}'
            + (f'<span class="chips">{"".join(chips)}</span>' if chips else "")
            + "</div></li>")

    lec = "\n".join(
        f'<li><a class="card" href="lectures/{slug}.html">'
        f'<span class="label">{esc(topic_number(topic_of("decks", slug)[0]))}</span>'
        f'<span class="title">{esc(title)}</span>'
        f'<span class="meta">{len(toc)} slides</span></a></li>'
        for slug, (title, toc, _) in decks.items())

    hand = "\n".join(
        f'<li><a href="{built["handouts"][slug][0]}">{esc(built["handouts"][slug][1])}</a></li>'
        for slug in HANDOUTS if slug in built["handouts"])
    sheets = "\n".join(
        f'<li><a href="{built["sheets"][slug][0]}" download>{esc(built["sheets"][slug][1])}</a>'
        f' <span class="desc">{esc(SPREADSHEETS[slug][2])}</span></li>'
        for slug in SPREADSHEETS if slug in built["sheets"])
    qz = "".join(f'<a class="chip" href="{h}">{esc(t)}</a>' for h, t in quizzes)

    body = f"""<section id="schedule">
<h2>Schedule</h2>
<p class="lede">The course runs from Aristotle to the present. Textbook: {TEXTBOOK}.</p>
<ol class="schedule">
{chr(10).join(rows)}
</ol>
</section>
<section id="lectures">
<h2>Lectures</h2>
<ul class="cards">
{lec}
</ul>
</section>
<section id="handouts">
<h2>Handouts and spreadsheets</h2>
<ul class="plain-list">
{hand}
</ul>
<h3>Spreadsheets</h3>
<ul class="plain-list">
{sheets}
</ul>
</section>
<section id="quizzes">
<h2>Quizzes and exams</h2>
<p class="chips">{qz}</p>
</section>"""
    (out / "index.html").write_text(
        page(SITE_TITLE, body, body_class="home",
             hero=("A course in the", SITE_TITLE,
                   "How thinkers from Aristotle to Hayek have understood economic life "
                   "and its place in human life as a whole.")),
        encoding="utf-8")


def main():
    global CSS_VERSION
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--src", type=Path, default=DEFAULT_SRC)
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent)
    args = ap.parse_args()
    if not args.src.is_dir():
        sys.exit(f"source directory not found: {args.src}")
    CSS_VERSION = hashlib.sha1((args.out / "style.css").read_bytes()).hexdigest()[:8]

    print("Lectures:")
    decks = build_lectures(args.src, args.out)
    print("Spreadsheets:")
    built = {
        "decks": {s: (f"lectures/{s}.html", d[0]) for s, d in decks.items()},
        "handouts": {s: (f"handouts/{s}.html", t) for s, (r, t) in HANDOUTS.items()
                     if (args.src / r).exists()},
        "sheets": copy_spreadsheets(args.src, args.out),
    }
    write_lectures(args.out, decks, built)
    print("Handouts:")
    build_handouts(args.src, args.out, built)
    print("Quizzes:")
    quizzes = build_quizzes(args.src, args.out)
    build_index(args.out, decks, built, quizzes)
    for w in warnings:
        print("warning:", w, file=sys.stderr)


if __name__ == "__main__":
    main()
