"""Build the Amharic tutorial v2 (HTML -> PDF via Chromium).

Two paths – A: private residence (P1–P4) · B: commercial project (01–61) – plus Google Sheets use.
usage: python build_tutorial.py   (expects the finalized workbooks in /tmp or ../output and screenshots in img2/)
"""
import os
import re
import subprocess
import sys

import pdfplumber
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "build"))
from facts import grab, etb, m, pct, C_NAMES, P_NAMES  # noqa: E402
from glossary import GLOSSARY  # noqa: E402

OUT = os.path.join(HERE, "..", "output")
IMG = os.path.join(HERE, "img2")
SRC = os.environ.get("RAPHA_SRC", "/tmp/claude-0/v2/out")

CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
body { margin: 0; font-family: 'Noto Sans Ethiopic', 'Liberation Sans', sans-serif; color: #1B2430; font-size: 10.4pt; line-height: 1.62; }
.page { padding: 4mm 16mm 2mm 16mm; page-break-after: always; }
.cover { height: 263mm; background: #0F2A44; color: #fff; padding: 30mm 18mm; position: relative; page-break-after: always; margin: -16mm 0 0 0; }
.cover h1 { font-family: 'Noto Serif Ethiopic'; font-size: 30pt; line-height: 1.3; margin: 18mm 0 4mm; }
.cover h2 { font-size: 15pt; font-weight: 400; color: #F3D7AE; margin: 0 0 10mm; }
.cover .rule { width: 60mm; height: 3px; background: #D98E2B; margin: 8mm 0; }
.cover .meta { position: absolute; bottom: 24mm; left: 18mm; right: 18mm; font-size: 9.5pt; color: #C9D3DF; }
.cover .badge { display: inline-block; border: 1px solid #D98E2B; color: #F3D7AE; padding: 2px 10px; border-radius: 14px; font-size: 9pt; margin-right: 6px; }
h1.ch { font-family: 'Noto Serif Ethiopic'; color: #0F2A44; font-size: 21pt; margin: 4mm 0 1mm; line-height: 1.3; }
.chno { color: #D98E2B; font-weight: 700; font-size: 10pt; letter-spacing: 1px; margin-top: 6mm; }
.chrule { height: 3px; width: 100%; background: linear-gradient(90deg, #D98E2B 0 30%, #D9DEE4 30% 100%); margin-bottom: 5mm; }
h2 { color: #0F2A44; font-size: 13pt; margin: 6mm 0 2mm; font-family: 'Noto Serif Ethiopic'; }
h3 { color: #1F5FAD; font-size: 11pt; margin: 4mm 0 1mm; }
p { margin: 1.5mm 0 2.5mm; text-align: justify; }
.en { font-family: 'Liberation Sans', sans-serif; font-size: 9pt; color: #5B6675; }
table { border-collapse: collapse; width: 100%; margin: 2mm 0 4mm; font-size: 9.2pt; }
th { background: #0F2A44; color: #fff; text-align: left; padding: 4px 6px; font-weight: 600; }
td { border-bottom: 1px solid #D9DEE4; padding: 4px 6px; vertical-align: top; }
tr:nth-child(even) td { background: #F4F6F8; }
td.num, th.num { text-align: right; white-space: nowrap; }
.box { border-left: 4px solid #D98E2B; background: #FDF1DC; padding: 3mm 4mm; margin: 3mm 0; border-radius: 2px; }
.box.blue { border-color: #1F5FAD; background: #EAF1FA; }
.box.green { border-color: #2E7D4F; background: #E3F2E8; }
.box.red { border-color: #B42318; background: #FDE7E5; }
.box b.t { display: block; color: #0F2A44; margin-bottom: 1mm; }
figure { margin: 3mm 0 4mm; page-break-inside: avoid; break-inside: avoid; display: inline-block; vertical-align: top; }
.shot { position: relative; display: block; break-inside: avoid; page-break-inside: avoid; border: 1px solid #D9DEE4; box-shadow: 0 1px 4px rgba(15,42,68,.12); }
.shot img { display: block; width: 100%; break-inside: avoid; }
.co { position: absolute; width: 20px; height: 20px; margin: -10px 0 0 -10px; border-radius: 50%; background: #D98E2B; color: #fff;
      font: 700 10.5px/20px 'Liberation Sans'; text-align: center; border: 2px solid #fff; box-shadow: 0 0 0 1.5px #D98E2B; }
figcaption { font-size: 8.8pt; color: #5B6675; margin-top: 1.5mm; }
ol.legend { margin: 1.5mm 0 0 0; padding-left: 0; list-style: none; font-size: 9pt; columns: 2; column-gap: 8mm; }
ol.legend li { margin-bottom: 1mm; break-inside: avoid; }
ol.legend li span { display: inline-block; width: 16px; height: 16px; border-radius: 50%; background: #D98E2B; color: #fff;
      font: 700 9px/16px 'Liberation Sans'; text-align: center; margin-right: 5px; }
.two { display: flex; gap: 6mm; }
.two > div { flex: 1; }
.pathA { border-top: 5px solid #2A9D8F; background: #F4F6F8; padding: 3mm 4mm; }
.pathB { border-top: 5px solid #1F5FAD; background: #F4F6F8; padding: 3mm 4mm; }
.flowwrap { display: inline-block; width: 100%; }
.flow { display: flex; align-items: center; justify-content: space-between; margin: 5mm 0; }
.flow .n { flex: 1; background: #0F2A44; color: #fff; border-radius: 6px; padding: 4mm 2mm; text-align: center; font-size: 9.5pt; min-height: 30mm; }
.flow .n b { display: block; color: #F3D7AE; font-size: 11pt; margin-bottom: 1mm; }
.flow .n.amber { background: #D98E2B; } .flow .n.amber b { color: #fff; }
.flow .a { width: 8mm; text-align: center; color: #D98E2B; font-size: 18pt; }
.toc td { border: none; padding: 2px 4px; font-size: 10.5pt; }
.toc td.p { text-align: right; color: #D98E2B; font-weight: 700; width: 12mm; }
.toc tr:nth-child(even) td { background: none; }
.kpi { display: flex; gap: 3mm; margin: 3mm 0; }
.kpi div { flex: 1; background: #F4F6F8; border-top: 3px solid #1F5FAD; padding: 2mm 3mm; font-size: 8.5pt; color: #5B6675; }
.kpi div b { display: block; font-size: 13pt; color: #0F2A44; }
.marker { font-size: 1px; color: #fff; }
.small { font-size: 8.8pt; color: #5B6675; }
.gl td { font-size: 8.6pt; padding: 2px 5px; }
.formula { font-family: 'Liberation Mono', monospace; background: #F4F6F8; padding: 2mm 3mm; font-size: 9pt; border-radius: 2px; display: block; margin: 2mm 0; }
"""


def crop(name, t=0.0, b=1.0, l=0.0, r=1.0):
    im = Image.open(os.path.join(IMG, name))
    W, H = im.size
    out = f"c_{name[:-4]}_{int(t*100)}_{int(b*100)}_{int(l*100)}_{int(r*100)}.png"
    im.crop((int(W * l), int(H * t), int(W * r), int(H * b))).save(os.path.join(IMG, out))
    return out


def fig(name, cap, callouts=(), legend=(), width=100, crop_=None):
    if crop_:
        name = crop(name, *crop_)
    path = os.path.join(IMG, name)
    cos = "".join(f'<span class="co" style="left:{x}%;top:{y}%">{n}</span>' for n, x, y in callouts)
    leg = ""
    if legend:
        leg = '<ol class="legend">' + "".join(f"<li><span>{i + 1}</span>{t}</li>" for i, t in enumerate(legend)) + "</ol>"
    return (f'<figure style="width:{width}%"><div class="shot"><img src="file://{path}">{cos}</div>'
            f"<figcaption>{cap}</figcaption>{leg}</figure>")


def chapter(no, title, en, marker, part=None):
    p = f'<div class="part">{part}</div>' if part else ""
    return (f'{p}<div class="chno">ምዕራፍ {no}<span class="marker">{marker}</span></div><h1 class="ch">{title}</h1>'
            f'<div class="en">{en}</div><div class="chrule"></div>')


def load_facts():
    C, wb = grab(f"{SRC}/main.xlsx", C_NAMES + P_NAMES)
    H, _ = grab(f"{SRC}/s2.xlsx", C_NAMES)
    P = {k: C[k] for k in P_NAMES}
    ws = wb["P2_Calc"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "Month after completion":
            P["AM1"] = [ws.cell(r + 1, c).value for c in range(2, 10)]
            P["AM2"] = [ws.cell(r + 2, c).value for c in range(2, 10)]
            break
    return C, H, P


EXTRA_CSS = """
.part { background:#0F2A44; color:#F3D7AE; display:inline-block; padding:1mm 4mm; border-radius:3px; font-size:9pt; margin-top:4mm; }
.partA { border-left:6px solid #2A9D8F; } .partB { border-left:6px solid #1F5FAD; }
.step { display:flex; gap:3mm; margin:2mm 0; } .step .n { flex:0 0 7mm; height:7mm; border-radius:50%; background:#D98E2B; color:#fff;
  font:700 10pt/7mm 'Liberation Sans'; text-align:center; } .step .t { flex:1; }
"""


def build(C, H, P, toc_pages=None):
    tp = toc_pages or {}
    S = []
    # ================================================================ cover
    S.append(f"""<div class="cover">
<div class="badge">የአጠቃቀም መመሪያ</div><div class="badge">ስሪት 2.0</div><div class="badge">Excel · Google Sheets</div>
<h1>የራፋ የግንባታ ፕሮጀክት<br>ትርፋማነት ሞዴል</h1>
<h2>Rapha Development Profitability Model — ደረጃ በደረጃ የአማርኛ መመሪያ</h2>
<div class="rule"></div>
<p style="font-size:12pt;color:#E6ECF2;max-width:150mm;text-align:left">አንድ ፋይል፣ ሁለት መሣሪያዎች፡ <b>ሀ — የግል መኖሪያ ቤት</b> የወጪ ግምትና የቤት ብድር ዕቅድ፤
<b>ለ — የንግድ ሕንፃ</b> ሙሉ ትርፋማነት፣ የኪራይ መዝገብ፣ የሆቴል ኪራይ፣ የIFRS መግለጫዎች፣ የባንክ ምጣኔዎችና የቀጥታ የስጋት ትንተና።
በExcel እና በGoogle Sheets ይሠራል።</p>
<div class="meta">ሦስት ምሳሌዎችን ይዟል፡ የG+1 የቤተሰብ ቤት · ምሳሌ 1 — B+G+8 ቅይጥ ሕንፃ · ምሳሌ 2 — B+G+10 የንግድ ማዕከልና ሆቴል (መሬት በአጋርነት)<br>
ፋይሎች፡ Rapha_Development_Profitability_Model_v2.xlsx · Rapha_Model_Blank.xlsx<br>
ይህ መሣሪያ ለዕቅድ ብቻ ነው፤ የኢንቨስትመንት፣ የግብር ወይም የሕግ ምክር አይደለም።</div></div>""")
    # ================================================================ TOC
    toc = [("1", "መግቢያ — በስሪት 2 ምን ተቀየረ?", "CH01"), ("2", "መጀመሪያ ገጽ፡ የግል ቤት ወይስ የንግድ ፕሮጀክት?", "CH02"),
           ("3", "ከመጀመርዎ በፊት — Excel እና Google Sheets", "CH03"),
           ("4", "ክፍል ሀ · የቤትዎ ግብዓቶች (P1)", "CH04"), ("5", "ክፍል ሀ · ስሌቶች፣ የገንዘብ ምንጭና የብድር ሰንጠረዥ (P2)", "CH05"),
           ("6", "ክፍል ሀ · የቤት ግምት ገጽና ማረጋገጫዎች (P3–P4)", "CH06"),
           ("7", "ክፍል ለ · የፕሮጀክቱ ግብዓቶችና ሁኔታዎች (10–12)", "CH07"), ("8", "ክፍል ለ · የቦታ ዝርዝርና የኪራይ መዝገብ (13)", "CH08"),
           ("9", "ክፍል ለ · ወጪ፣ ወርሃዊ ግንባታና ብድር (20–30)", "CH09"), ("10", "ክፍል ለ · የፋይናንስ መግለጫዎች (40–43)", "CH10"),
           ("11", "ክፍል ለ · ተመላሽና ውሳኔ (50–51)", "CH11"), ("12", "ክፍል ለ · የቀጥታ ስጋትና የሚዛን ነጥብ (52)", "CH12"),
           ("13", "ክፍል ለ · ዳሽቦርድና ቼኮች (01, 60)", "CH13"), ("14", "ምሳሌ 1፡ B+G+8 ቅይጥ ሕንፃ", "CH14"),
           ("15", "ምሳሌ 2፡ የንግድ ማዕከልና ሆቴል", "CH15"), ("16", "የራስዎን ፕሮጀክት ማስገባት — የማረጋገጫ ዝርዝር", "CH16"),
           ("17", "መዝገበ ቃላትና ማሳሰቢያ", "CH17")]
    rows = "".join(f'<tr><td style="width:14mm;color:#0F2A44;font-weight:700">{n}</td><td>{t}</td><td class="p">{tp.get(mk, "")}</td></tr>'
                   for n, t, mk in toc)
    S.append(f"""<div class="page"><h1 class="ch" style="margin-top:8mm">ማውጫ</h1><div class="chrule"></div>
<table class="toc">{rows}</table>
<div class="two"><div class="pathA"><b style="color:#2A9D8F">የራስዎን ቤት ብቻ ለመገመት</b><br>ምዕራፍ 1–6ን ያንብቡ (መንገድ ሀ)። ቀሪው አያስፈልግዎትም።</div>
<div class="pathB"><b style="color:#1F5FAD">ለሽያጭ ወይም ለኪራይ ለሚገነቡ</b><br>ምዕራፍ 1–3ን ከዚያ 7–16ን ይከተሉ (መንገድ ለ)።</div></div>
<div class="box blue"><b class="t">ምስሎቹን እንዴት ማንበብ?</b>በምስሎቹ ላይ ያሉት <b style="color:#D98E2B">ብርቱካናማ ቁጥሮች</b> ከምስሉ በታች ካለው ማብራሪያ ጋር ይዛመዳሉ።
በመመሪያው ውስጥ የተጠቀሱት ቁጥሮች ሁሉ ከሞዴሉ ፋይል በቀጥታ የተወሰዱ ናቸው።</div>
<div class="box"><b class="t">ቃላት</b>እንደ IRR፣ NPV፣ DSCR፣ IFRS ያሉ የእንግሊዝኛ ቃላት በመጀመሪያ ሲጠቀሱ ተብራርተዋል፤ ሙሉ ዝርዝሩ በምዕራፍ 17 እና በሞዴሉ ገጽ 61 ይገኛል።</div>
</div>""")
    # ================================================================ CH1
    S.append(f"""<div class="page">{chapter(1, "መግቢያ — በስሪት 2 ምን ተቀየረ?", "Introduction – what is new in version 2", "CH01")}
<p>ቤት ወይም ሕንፃ ከመገንባትዎ በፊት የሚጠይቁት ጥያቄ እንደ ፕሮጀክቱ ዓይነት ይለያያል። ለራስዎ ቤት፡ <b>«ስንት ያስወጣኛል? በየወሩ ስንት ገንዘብ ያስፈልጋል?
ብድሩን እችለዋለሁ?»</b>። ለንግድ ሕንፃ፡ <b>«ያተርፋል? ባንኩ ያበድረኛል? ዋጋ ቢወርድ ወይም ግንባታው ቢዘገይ ምን ይሆናል?»</b>።
ስሪት 2 እነዚህን ሁለት ጥያቄዎች <b>በሁለት የተለያዩ ክፍሎች</b> ይመልሳል — በአንድ ፋይል ውስጥ፣ ግን ግብዓቶቻቸውና ስሌቶቻቸው ተለያይተው።</p>
<h2>በስሪት 2 የተጨመሩና የተሻሻሉ</h2>
<table><tr><th style="width:42%">ለውጥ</th><th>ጥቅሙ</th></tr>
<tr><td><b>ሁለት መንገዶች</b> — በመጀመሪያ ገጽ «የግል ቤት» ወይም «የንግድ ፕሮጀክት» ይመርጣሉ</td><td>የግል ቤት ገንቢ የንግድ ግብዓቶችን አያይም፤ የንግድ ሞዴሉ ደግሞ የቤተሰብ ገቢ የመሳሰሉ የግል መረጃዎችን አይጠይቅም</td></tr>
<tr><td><b>ለግል ቤት የተለየ ክፍል</b> (P1–P4)፡ የራሱ የወጪ መረጃ ቋት፣ ወርሃዊ ወጪ፣ የቁጠባ→ቤተሰብ→ብድር ቅደም ተከተልና በወር የተዘረዘረ የ240 ወራት የብድር ሰንጠረዥ</td><td>ቀላል፣ ግልጽና በአንድ ገጽ የሚታተም የቤት ግምት</td></tr>
<tr><td><b>የቦታ ዝርዝርና የኪራይ መዝገብ</b> (ገጽ 13)፡ እስከ 12 ቦታዎች በዓይነት — ሱቅ፣ ቢሮ፣ መኖሪያ፣ ማሳያ፣ ምግብ ቤት፣ መጋዘን፣ ሆቴል፣ መኪና ማቆሚያ፣ የባለቤት ቢሮ፣ ያለ ኪራይ የሚሰጥ ቦታ</td><td>እያንዳንዱ ቦታ የራሱ ዋጋ፣ ኪራይና የተያዘ መጠን አለው</td></tr>
<tr><td><b>የሆቴል ኦፕሬተር ኪራይ</b>፣ አራት ስም ያላቸው ሌሎች ገቢዎች፣ የልዩ ባለሙያ አማካሪ ክፍያና ከመክፈቻ በፊት ወጪ</td><td>ለትልልቅ የንግድ ማዕከላት የሚያስፈልጉ ዝርዝሮች</td></tr>
<tr><td><b>ሁሉም ምርጫዎች ዝርዝር (dropdown)</b> ናቸው፤ ለእያንዳንዱ ግብዓት «እርግጠኝነት» (ከፍተኛ/መካከለኛ/ዝቅተኛ) እና «ምንጭ» ዓምድ</td><td>የመጻፍ ስህተት ይቀንሳል፤ ደካማ ግምቶች ይታወቃሉ</td></tr>
<tr><td><b>የቀጥታ የስጋት ትንተና</b> — Data Table የለም፤ ቶርኔዶ፣ የሚዛን ነጥቦችና «ቢሆንስ» ማስተካከያዎች ሁሉ በቀመር ይሰላሉ</td><td>በGoogle Sheets እና በLibreOffice ጭምር በትክክል ይሠራል</td></tr>
<tr><td><b>ጠቅ ሲያደርጉ ወደ ገጹ የሚወስዱ አገናኞች</b> — በመጀመሪያ ገጽና በእያንዳንዱ ገጽ ራስጌ «◂ Start» እና «Dashboard ▸»</td><td>በExcel እና በGoogle Sheets በቀጥታ ወደሚፈልጉት ገጽ ይዘልላሉ</td></tr>
<tr><td><b>ራሱን የሚጽፍ ማጠቃለያ</b> — በዳሽቦርዱ ላይ «ዋና ዋና ግኝቶች»</td><td>ቁጥሮቹን በቃላት ያብራራል፤ ግብዓት ሲቀየር ራሱ ይቀየራል</td></tr></table>
<h2>የፋይሉ አወቃቀር</h2>
<table><tr><th>ክፍል</th><th>ገጾች</th><th>ሥራቸው</th></tr>
<tr><td>መነሻ</td><td>00 Start</td><td>መንገዱን መምረጥ፣ የሁኔታ ሳጥኖች፣ ወደ ሁሉም ገጾች አገናኝ</td></tr>
<tr class="partA"><td>ሀ · የግል ቤት</td><td>P1 Inputs · P2 Calc · P3 Estimate · P4 Checks</td><td>ግብዓት → ስሌትና ብድር → የአንድ ገጽ ግምት → ማረጋገጫዎች</td></tr>
<tr class="partB"><td>ለ · ሪፖርት</td><td>01 Dashboard</td><td>ውሳኔ፣ ስድስት ቁልፍ ቁጥሮች፣ ገበታዎች፣ ማጠቃለያ</td></tr>
<tr class="partB"><td>ለ · ግብዓት</td><td>10 Inputs · 11 Scenarios · 12 TaxLaw · 13 Spaces</td><td>እርስዎ የሚሞሉት ቦታ</td></tr>
<tr class="partB"><td>ለ · ሞተር</td><td>20–30</td><td>ጊዜ፣ ስፋት፣ ወጪ፣ ወርሃዊ ግንባታ፣ ገቢ፣ ወጪ፣ ንብረት፣ ብድር፣ ግብር፣ ተ.እ.ታ</td></tr>
<tr class="partB"><td>ለ · መግለጫና ትንተና</td><td>40–43 · 50–52 · 60</td><td>IFRS መግለጫዎች፣ ተመላሽ፣ ስጋት፣ 25 ማረጋገጫዎች</td></tr>
<tr><td>የጋራ</td><td>61 Glossary</td><td>{len(GLOSSARY)} የእንግሊዝኛ–አማርኛ ቃላት</td></tr></table>
</div>""")
    # ================================================================ CH2
    S.append(f"""<div class="page">{chapter(2, "መጀመሪያ ገጽ፡ የግል ቤት ወይስ የንግድ ፕሮጀክት?", "The Start sheet – choose your track", "CH02")}
<p>ፋይሉን ሲከፍቱ <b>00 Start</b> ገጽ ይታያል። የሚያደርጉት የመጀመሪያው ነገር <b>ደረጃ 1</b> ላይ ካለው ዝርዝር አንዱን መምረጥ ነው፡</p>
<div class="two"><div class="pathA"><h3 style="color:#2A9D8F;margin-top:0">ሀ — Private residence (የግል መኖሪያ ቤት)</h3>
<p>ለራስዎ ወይም ለቤተሰብዎ ለሚሠሩት G+0 ቪላ፣ G+1፣ G+2 ወይም G+3 ቤት። የሚያገኙት፡ የግንባታ ወጪ ግምትና ክልል፣ ወርሃዊ የገንዘብ ፍላጎት፣
የኮንትራክተር ክፍያ ደረጃዎች፣ የቁጠባ–ቤተሰብ–ብድር ዕቅድ፣ ወርሃዊ የብድር ክፍያና የመክፈል አቅም። ትርፍ፣ ግብርና IFRS አይመለከቱዎትም።</p></div>
<div class="pathB"><h3 style="color:#1F5FAD;margin-top:0">ለ — Commercial project (የንግድ ፕሮጀክት)</h3>
<p>ለሽያጭ፣ ለኪራይ ወይም ለሁለቱም ለሚገነባ ሕንፃ። የፕሮጀክቱን ዓይነት (1 ለሽያጭ · 2 ለኪራይ/ለመያዝ · 3 ቅይጥ)፣ መረጃውን (የእኔ ፕሮጀክት፣
ምሳሌ 1፣ ምሳሌ 2) እና ሁኔታውን (መሠረታዊ፣ ብሩህ፣ አስቸጋሪ፣ ባንክ) ይመርጣሉ።</p></div></div>
{fig("P_00_Start_1.png", "ምስል 2.1 — «Private residence» ሲመረጥ፡ ካርድ ሀ ንቁ ነው፤ ካርድ ለ በግራጫ ይጠፋል።",
     [(1, 51, 18.6), (2, 37, 30.6), (3, 41, 40), (4, 9, 54.5), (5, 75, 34), (6, 7, 80)],
     ["ደረጃ 1፡ መንገዱን ይምረጡ", "የትኛው መረጃ — የእኔ ቤት ወይም የናሙና ቤት", "የሁኔታ ሳጥን፡ ማረጋገጫዎች፣ ጠቅላላ ወጪ፣ ወርሃዊ ክፍያ፣ የመክፈል አቅም",
      "ጠቅ ሲያደርጉ ወደ ገጹ የሚወስዱ አገናኞች", "ያልተመረጠው ካርድ በግራጫ ይታያል", "የሁሉም ገጾች ዝርዝር — እያንዳንዱ አገናኝ ነው"],
     width=100, crop_=(0, 0.63, 0.22, 0.78))}
</div>""")
    S.append(f"""<div class="page">
{fig("C_00_Start_1.png", "ምስል 2.2 — «Commercial project» ሲመረጥ፡ ካርድ ለ ንቁ ነው። ምሳሌ 1 በዓይነት 3 (ቅይጥ) ተጭኗል።",
     [(1, 84, 30.6), (2, 84, 34.6), (3, 84, 38.6), (4, 93, 50), (5, 54, 66)],
     ["የፕሮጀክት ዓይነት፡ 1 ለሽያጭ · 2 ለኪራይ/ለመያዝ · 3 ቅይጥ", "መረጃ፡ የእኔ ፕሮጀክት · ምሳሌ 1 · ምሳሌ 2",
      "ሁኔታ፡ መሠረታዊ · ብሩህ · አስቸጋሪ · የባንክ", "ማረጋገጫዎች፣ ውሳኔ፣ TDC፣ የባለቤት IRR፣ ዝቅተኛ DSCR", "ወደ ዋና ዋና ገጾች መሄጃ"],
     width=100, crop_=(0, 0.63, 0.22, 0.78))}
<h2>አገናኞቹ እንዴት ይሠራሉ?</h2>
<p>በመጀመሪያ ገጽ ላይ ያሉት ሰማያዊ የተሰመረባቸው ጽሑፎች ሁሉ <b>ቋሚ አገናኞች (hyperlinks)</b> ናቸው — አንዱን ጠቅ ሲያደርጉ በቀጥታ ወደዚያ ገጽ ይወስድዎታል።
በእያንዳንዱ ገጽ ቀኝ ላይኛው ጥግ <b>«◂ Start»</b> ወደ መጀመሪያ ገጽ ይመልስዎታል፤ ከሥሩ ያለው <b>«Dashboard ▸»</b> (ለንግድ ገጾች) ወይም
<b>«Estimate ▸»</b> (ለግል ቤት ገጾች) ወደ ዋናው ሪፖርት ይወስድዎታል። በስሪት 1 እነዚህ አገናኞች በቀመር (HYPERLINK) ስለነበሩ በGoogle Sheets አይሠሩም ነበር፤
አሁን ቀመር ስላልሆኑ በExcel፣ በLibreOffice እና በGoogle Sheets እኩል ይሠራሉ።</p>
<h2>ምርጫው በሁሉም ገጾች ይታያል</h2>
<p>ደረጃ 1ን ሲቀይሩ፡ (1) ያልተመረጠው ካርድና ቁልፉ በግራጫ ይጠፋሉ፤ (2) ከደረጃ 1 በታች ያሉት <b>«▶ A · Open my house estimate»</b> እና
<b>«▶ B · Open the Dashboard»</b> ቁልፎች በቀጥታ ወደ ዋናው ሪፖርት ይወስዱዎታል፤ (3) ያልተመረጠው ክፍል ገጾች ራስጌ <b style="color:#D98E2B">ብርቱካናማ</b>
ሆኖ «⚠ … is selected on Start – this sheet belongs to the … track (not in use)» ይላል። ስለዚህ በስህተት የሌላውን ክፍል ገጽ እያነበቡ መሆኑን ወዲያውኑ ያውቃሉ።</p>
<div class="box"><b class="t">ጠቃሚ ምክር</b>ሁለቱም ክፍሎች ሁልጊዜ ይሰላሉ — መንገድ መምረጥ የሚወስነው የትኛው ካርድ ንቁ ሆኖ እንደሚታይ ብቻ ነው። ስለዚህ የቤት ግምትዎንና የንግድ ፕሮጀክትዎን
በአንድ ፋይል ውስጥ አብረው ማስቀመጥ ይችላሉ።</div>
</div>""")
    # ================================================================ CH3
    S.append(f"""<div class="page">{chapter(3, "ከመጀመርዎ በፊት — Excel እና Google Sheets", "Before you start – Excel and Google Sheets", "CH03")}
<h2>በExcel</h2>
<p><b>Microsoft Excel 2016 ወይም ከዚያ በኋላ</b> (Microsoft 365 ይመረጣል)። ፋይሉ ማክሮ የለውም። <b>Formulas › Calculation Options › Automatic</b>
መመረጡን ያረጋግጡ። ሞዴሉ ዙሪያ-ጥገኛ ቀመር (circularity) ስለሌለው «iterative calculation» አያስፈልግም።</p>
<h2>በGoogle Sheets</h2>
<div class="step"><div class="n">1</div><div class="t">ፋይሉን (.xlsx) ወደ Google Drive ይጫኑ (New › File upload)።</div></div>
<div class="step"><div class="n">2</div><div class="t">ፋይሉን በቀኝ ጠቅ ያድርጉና <b>Open with › Google Sheets</b> ይምረጡ። ቋሚ ቅጂ ለማድረግ <b>File › Save as Google Sheets</b>።</div></div>
<div class="step"><div class="n">3</div><div class="t">ሁሉም ቀመሮች እንደገና ይሰላሉ (ጥቂት ሰከንዶች)። በመጀመሪያ ገጽ ያሉትን ዝርዝሮችና አገናኞች ይሞክሩ፤ ወደ ገጾቹ መዝለል አለባቸው።</div></div>
<div class="step"><div class="n">4</div><div class="t"><b>File › Settings › Calculation</b>፡ «Iterative calculation» ጠፍቶ (Off) ይተዉት።</div></div>
<div class="box green"><b class="t">ለGoogle Sheets የተደረጉ ማስተካከያዎች</b>ቋሚ አገናኞች (ቀመር ያልሆኑ)፤ ዝርዝሮቹ ከተደበቀው «Lists» ገጽ ክልል ይነበባሉ፤
የቀለም ማስጠንቀቂያዎች (conditional formatting) የሚያነቡት በዚያው ገጽ ያሉ ሕዋሶችን ብቻ ነው፤ Data Table የለም — የስጋት ትንተናው በቀመር ይሰላል፤
የተሰየሙ ክልሎች (named ranges) ከተግባር ስሞች ጋር አይጋጩም። በሙከራ ፋይል በGoogle Sheets ላይ ተፈትኗል።</div>
<h2>የቀለም ኮድ — ዋናው ሕግ</h2>
<table><tr><th>መልክ</th><th>ትርጉም</th><th>ምን ያደርጋሉ?</th></tr>
<tr><td style="color:#1F4E9C;background:#FFF8E6;border:1px solid #D9DEE4">ሰማያዊ ፊደል፣ ቀላል ቢጫ ጀርባ</td><td>ግብዓት</td><td>እዚህ ብቻ ይጻፉ ወይም ከዝርዝሩ ይምረጡ</td></tr>
<tr><td style="color:#1B2430">ጥቁር ፊደል</td><td>ቀመር</td><td>በፍጹም አይጻፉበት</td></tr>
<tr><td style="color:#1E7B34">አረንጓዴ ፊደል</td><td>ከሌላ ገጽ የመጣ ማገናኛ</td><td>አይጻፉበት</td></tr>
<tr><td style="background:#FFF2B3;font-weight:700">ቢጫ ጀርባ፣ ደማቅ</td><td>ቁልፍ ውጤት ወይም ማረጋገጫ</td><td>ያንብቡት</td></tr></table>
<div class="box red"><b class="t">ጥቁር ሕዋስ ላይ አይጻፉ!</b>በቀመር ላይ ቁጥር ከጻፉ ሞዴሉ ይሰበራል (ለምሳሌ የሀብትና ዕዳ መግለጫ ሚዛኑን ያጣል — ማረጋገጫ C01 ይወድቃል)።
በስህተት ከጻፉ Ctrl+Z ይጫኑ፣ ወይም ከዋናው ፋይል አዲስ ቅጂ ይውሰዱ።</div>
<h2>መረጃዎን የት ያስገባሉ? — የዓምዶቹ ሕግ</h2>
<p>እያንዳንዱ የግብዓት ገጽ የመረጃ ዓምዶች አሉት። <b>የግል ቤት (P1)</b>፡ E = የእኔ ቤት፣ F = ናሙና፣ <b>G = ACTIVE</b>።
<b>የንግድ (10 Inputs)</b>፡ E = የእኔ ፕሮጀክት፣ F = ምሳሌ 1፣ G = ምሳሌ 2፣ <b>H = ACTIVE</b>። ሞዴሉ የሚጠቀመው ACTIVE ዓምዱን ብቻ ነው፤
እሱም በመጀመሪያ ገጽ «Which data?» የመረጡትን ዓምድ ይወስዳል። ስለዚህ፡ (1) «My house» ወይም «My project» ይምረጡ፣ (2) በዓምድ E ይጻፉ፣
(3) ምሳሌዎቹን ለንጽጽር እንዳሉ ይተዉዋቸው። ባዶ ፋይሉ (Rapha_Model_Blank.xlsx) ምሳሌዎቹ የተሰረዙበት ተመሳሳይ ሞዴል ነው።</p>
</div>""")
    # ================================================================ CH4 – P1
    S.append(f"""<div class="page">{chapter(4, "የቤትዎ ግብዓቶች (P1 Inputs)", "Private track – inputs", "CH04", "ክፍል ሀ · የግል መኖሪያ ቤት")}
<p>የግል ቤት ክፍል የራሱ የግብዓት ገጽ አለው። ሰባት ብሎኮች አሉት፤ ሁሉም ምርጫዎች ዝርዝር (dropdown) ናቸው፣ ቁጥሮቹም በትክክለኛ ክልል ውስጥ መሆናቸው ይፈተሻል።</p>
{fig("P_P1_Inputs_1.png", "ምስል 4.1 — P1 Inputs፡ ዓምድ E (አረንጓዴ ርዕስ) የእርስዎ፣ F ናሙናው፣ G ሞዴሉ የሚጠቀመው።", width=100, crop_=(0.07, 0.75, 0, 0.82))}
<table><tr><th style="width:30%">ብሎክ</th><th>ምን ይጽፋሉ</th></tr>
<tr><td>1. ቤትዎ</td><td>ስም፣ አካባቢ፣ የቦታ ዓይነት (አዲስ አበባ / የክልል ከተማ)፣ የቦታ ስፋት፣ የመጀመሪያ ቀን፣ የግንባታ ወራት፣ የሕንፃ ዓይነት (G+0…G+3)፣ የማጠናቀቂያ ደረጃ፣ ምድር ቤት</td></tr>
<tr><td>2. የወለል ስፋት</td><td>ለእያንዳንዱ ወለል ጠቅላላ ስፋት (ግድግዳና ደረጃ ጨምሮ)። ሞዴሉ የሕንፃ ዓይነቱ ከፎቆቹ ጋር መስማማቱን ይፈትሻል</td></tr>
<tr><td>3. ተጨማሪ ሥራዎችና ክፍያዎች</td><td>አጥር፣ የፍሳሽ ጉድጓድ፣ የውሃ ታንከር፣ ንጣፍ፣ መጠባበቂያ፣ ዲዛይን፣ ቁጥጥር፣ ፈቃድ፣ የቤት ዕቃ፣ የዋጋ ጭማሪ፣ የወጪ ስርጭት፣ የኮንትራክተር ማስያዣ፣ የግምት ክልል</td></tr>
<tr><td>4. መሬት</td><td>መሬቱ የእርስዎ ከሆነ «Already owned»፤ በግንባታ ወቅት የሊዝ/የግዢ ክፍያ ካለ መጠኑን</td></tr>
<tr><td>5. የገንዘብ ምንጭ</td><td>ቁጠባ፣ የቤተሰብ/ዳያስፖራ መዋጮ፣ ባንኩ የሚሰጠው ከፍተኛ ብድር፣ ወለድ፣ የመክፈያ ዓመታት፣ የእፎይታ ወራት፣ በግንባታ ወቅት ወለዱ ይከፈላል ወይስ በብድሩ ላይ ይጨመራል፣ የብድር ማመቻቻ ክፍያ</td></tr>
<tr><td>6. የመክፈል አቅም</td><td>የቤተሰቡ ወርሃዊ የተጣራ ገቢ፣ ሌሎች ወርሃዊ የብድር ክፍያዎች፣ ባንኩ የሚፈቅደው ከፍተኛ ድርሻ (በተለምዶ 30–40%)</td></tr>
<tr><td>7. አማራጭ፡ ማከራየት</td><td>ከፊሉን ቢያከራዩ፡ ስፋት፣ ወርሃዊ ኪራይ፣ ጭማሪ፣ ባዶ የሚሆንባቸው ወራት፣ ወጪና የኪራይ ገቢ ግብር</td></tr></table>
<div class="box blue"><b class="t">የወጪ መረጃ ቋት (ብሎክ 8)</b>ከገጹ ግርጌ ለ8 የሥራ ክፍሎች በካሬ ሜትር ዋጋ አለ (Basic / Standard / Premium)። የመረጡት የማጠናቀቂያ ደረጃ
ዓምዱን ይወስናል። በአካባቢዎ ካለው ወቅታዊ ዋጋ ጋር ያስተካክሉት — ናሙናዎቹ «Sample value — not market advice» ናቸው።</div>
</div>""")
    # ================================================================ CH5 – P2
    a1, a2 = P["AM1"], P["AM2"]
    S.append(f"""<div class="page">{chapter(5, "ስሌቶች፣ የገንዘብ ምንጭና የብድር ሰንጠረዥ (P2 Calc)", "Private track – calculations, funding and the loan schedule", "CH05", "ክፍል ሀ · የግል መኖሪያ ቤት")}
<p>P2 ገጽ ከላይ ወደ ታች ይነበባል፡ የሕንፃ ማጠቃለያ → የዛሬ ዋጋ ወጪ → ወር በወር ወጪ → ማን ይከፍላል → በግንባታ ወቅት ብድር → ድምሮች → የኮንትራክተር ደረጃዎች → የብድር ሰንጠረዥ።</p>
{fig("P_P2_Calc_1.png", "ምስል 5.1 — P2፡ የሕንፃ ወጪ በዛሬ ዋጋ፣ ወር በወር ወጪና የገንዘብ ምንጭ (ናሙና ቤት፣ 14 ወራት)።",
     [(1, 26, 26), (2, 26, 50), (3, 26, 75), (4, 26, 89)],
     ["የዛሬ ዋጋ፡ የመረጃ ቋት ዋጋ × ስፋት × የቦታ ማስተካከያ + ተጨማሪ ሥራዎች", "የS-ኩርባ ስርጭትና የዋጋ ጭማሪ ማባዣ",
      "ቁጠባ → ቤተሰብ → ብድር → ክፍተት (በዚህ ቅደም ተከተል)", "በግንባታ ወቅት የብድሩ ቀሪ ሂሳብና ወለድ"], width=100, crop_=(0.12, 0.98, 0, 1))}
<h2>ወጪው እንዴት ይሰላል?</h2>
<p>ናሙና ቤቱ {etb(P["P_GFA"])} m² ነው። በዛሬ ዋጋ የሕንፃው ወጪ <b>{etb(P["P_BUILD_BASE"])} ብር</b> ነው። ሥራው በ{P["P_MONTHS"]} ወራት በS-ኩርባ ይሰራጫል፤
በየወሩ ዋጋ በዓመት {pct(P["P_ESC"], 0)} ስለሚጨምር የሕንፃ ሥራው <b>{etb(P["P_BUILD_ESC"])} ብር</b> ይሆናል። መጠባበቂያ ({etb(P["P_CONT_TOT"])})፣ ዲዛይን፣
ቁጥጥር፣ ፈቃድና የቤት ዕቃ ሲጨመሩ የግንባታ ወጪው <b>{etb(P["P_CONSTR"])} ብር</b> ({etb(P["P_CC_M2"])} ብር/m²) ይሆናል።</p>
</div>""")
    S.append(f"""<div class="page">
<h2>ማን ይከፍላል? — ቁጠባ → ቤተሰብ → ብድር</h2>
<p>በየወሩ የሚያስፈልገው ገንዘብ በመጀመሪያ ከቁጠባዎ፣ ሲያልቅ ከቤተሰብ መዋጮ፣ ከዚያም ከቤት ብድር ይወሰዳል። ሦስቱም ካልበቁ ቀሪው «ክፍተት» (funding gap) ይባላል።
በናሙናው ቁጠባ {etb(P["P_SAV_USED"])}፣ የቤተሰብ መዋጮ {etb(P["P_FAM_USED"])} እና ብድር {etb(P["P_LOAN_DRAWN"])} ብር ጥቅም ላይ ውለዋል፤ ክፍተት የለም።
የኮንትራክተሩ ማስያዣ ({etb(P["P_RET_TOT"])} ብር) በርክክብ ወር ይቀመጣል፣ ከጉድለት ማረሚያ ጊዜ በኋላ ይከፈላል።</p>
<h2>በግንባታ ወቅት ወለድ</h2>
<p>ብድር ከወጣበት ወር ጀምሮ ወለድ ይታሰባል፡ <span class="formula">ወለድ = (የወሩ መነሻ ቀሪ + የወሩ ብድር) × {pct(P["P_RATE"], 0)} ÷ 12</span>
«Capitalised» ከመረጡ ወለዱ በብድሩ ላይ ይጨመራል (ናሙናው እንዲህ ነው፤ በግንባታ ወቅት በጠቅላላ {etb(P["P_IDC_TOT"])} ብር)፤
«Paid monthly» ከመረጡ ከገቢዎ ይከፍሉታል — ሞዴሉ ከፍተኛውን ወርሃዊ መጠን ያሳያል። በግንባታው መጨረሻ የብድሩ ቀሪ <b>{etb(P["P_B0"])} ብር</b> ነው።</p>
{fig("P_P2_Calc_2.png", "ምስል 5.2 — ድምሮች፣ የወጪ ክፍፍል፣ የኮንትራክተር ደረጃዎችና በወር የተዘረዘረው የብድር ሰንጠረዥ መጀመሪያ።", width=82, crop_=(0, 1, 0, 0.6))}
<h2>የብድር ሰንጠረዡ (በወር፣ ወደ ታች)</h2>
<p>ከግንባታው በኋላ ለእያንዳንዱ ወር አንድ ረድፍ አለ (እስከ 240 ወራት)። የእፎይታ ወራት ካሉ ወለድ ብቻ ይከፈላል፤ ከዚያ እኩል ወርሃዊ ክፍያ (annuity):
<span class="formula">ወርሃዊ ክፍያ = PMT({pct(P["P_RATE"], 0)}/12, {P["P_NPAY"]}, −{etb(P["P_B0"])}) = {etb(P["P_INSTAL"])} ብር</span>
የመጀመሪያው ወር፡ መነሻ {etb(a1[3])} · ወለድ {etb(a1[4])} · ክፍያ {etb(a1[5])} · ዋና ገንዘብ {etb(a1[6])} · ቀሪ {etb(a1[7])}።
ሁለተኛው ወር ወለዱ ትንሽ ቀንሶ ({etb(a2[4])}) ዋናው ገንዘብ ይጨምራል ({etb(a2[6])})። ብድሩ በ{P["P_PAYOFF"]:%d/%m/%Y} ይጠናቀቃል።</p>
</div>""")
    # ================================================================ CH6 – P3/P4
    S.append(f"""<div class="page">{chapter(6, "የቤት ግምት ገጽና ማረጋገጫዎች (P3–P4)", "Private track – the one-page estimate and checks", "CH06", "ክፍል ሀ · የግል መኖሪያ ቤት")}
<div class="two"><div style="flex:1.25">{fig("P_P3_Estimate_1.png", "ምስል 6.1 — P3 Estimate፡ በአንድ A4 ገጽ የሚታተም የቤት ግምት።", width=100)}</div>
<div><h3 style="margin-top:2mm">አራቱ ካርዶች</h3>
<p><b>የግንባታ ወጪ</b> {m(P["P_CONSTR"])} ሚሊዮን (ክልል {m(P["P_LOW"])}–{m(P["P_HIGH"])} ሚ.)፤ <b>ጠቅላላ ወጪ</b> {m(P["P_TOTAL"])} ሚ.
(የብድር ክፍያና በግንባታ ወቅት ወለድ ጨምሮ)፤ <b>ወርሃዊ ክፍያ</b> {etb(P["P_INSTAL"])} ብር — የገቢውን {pct(P["P_DTI"], 0)}፤ <b>ክፍተት</b> የለም።</p>
<h3>የወጪ ክፍፍል</h3><p>አሥር ክፍሎች ከመሠረት እስከ ብድር ወጪ፤ ድምራቸው ከጠቅላላው ወጪ ጋር እኩል ነው (ማረጋገጫ P02)።</p>
<h3>ለኮንትራክተሩ ክፍያ</h3><p>ስድስት ደረጃዎች፣ የተጠራቀመ መቶኛና በግምት በስንተኛው ወር እንደሚደርሱ። ውል ሲደራደሩ ይጠቅማል።</p>
<h3>የመክፈል አቅም</h3><p>(ወርሃዊ ክፍያ + ሌሎች ክፍያዎች) ÷ ገቢ። ከ{pct(P["P_AFF_MAX"], 0)} በታች «Affordable» (አረንጓዴ)፤ እስከ 25% በላይ «Stretch»፤ ከዚያ በላይ «Not affordable» (ቀይ)።
ገቢዎ ሊሸከመው የሚችለው ከፍተኛ ብድርም ይታያል ({m(P["P_LOAN_AFF"])} ሚ.)።</p>
<h3>ማከራየት</h3><p>ናሙናው 60 m² ቢያከራይ፡ የመጀመሪያው ዓመት ኪራይ {etb(P["P_RENT_Y1"])} ብር፣ ከወጪና ግብር በኋላ {etb(P["P_RENT_NET"])}፤
የተከራየው ክፍል ወጪ በ{P["P_RENT_PB"]:.1f} ዓመታት ይመለሳል፤ ኪራዩ ወርሃዊ ክፍያውን {pct(P["P_RENT_COVER"], 0)} ይሸፍናል።</p></div></div>
</div>""")
    S.append(f"""<div class="page">
<h2>P4 Checks — የግል ቤት ማረጋገጫዎች</h2>
{fig("P_P4_Checks_1.png", f"ምስል 6.2 — {P['PCHK_PASS']} / {P['PCHK_TOTAL']} ማረጋገጫዎች PASS። ስህተቶች (Error) ሁልጊዜ PASS መሆን አለባቸው፤ ማስጠንቀቂያዎች (Warning) ምክር ናቸው።",
     width=90, crop_=(0, 1, 0, 0.62))}
<table><tr><th>ማረጋገጫ</th><th>ምን ይፈትሻል</th><th>ቢወድቅ</th></tr>
<tr><td>P01–P04</td><td>ገንዘብ ገቢ = ወጪ በየወሩ፤ ክፍፍሉ ከጠቅላላው ጋር እኩል፤ የኮንትራክተር ደረጃዎች</td><td>ቀመር ተጽፎበታል — ከዋናው ፋይል ይቅዱ</td></tr>
<tr><td>P05–P06</td><td>ብድሩ በጊዜው ተከፍሎ ያልቃል</td><td>የመክፈያ ጊዜውን ያስተካክሉ</td></tr>
<tr><td>P07–P08</td><td>ምርጫዎቹ ትክክል ናቸው፤ በገጾቹ ላይ የቀመር ስህተት የለም</td><td>ዝርዝሩን እንደገና ይምረጡ</td></tr>
<tr><td>P09 (W)</td><td>የሕንፃ ዓይነቱ ከፎቆቹ ጋር ይስማማል</td><td>በP1 ዓይነቱን ወይም ስፋቱን ያስተካክሉ</td></tr>
<tr><td>P10 (W)</td><td>የገንዘብ ክፍተት የለም</td><td>ቁጠባ፣ መዋጮ ወይም ብድር ይጨምሩ፤ ወጪውን ይቀንሱ</td></tr>
<tr><td>P11 (W)</td><td>ወርሃዊ ክፍያው በአቅምዎ ውስጥ ነው</td><td>ብድሩን ይቀንሱ ወይም የመክፈያ ጊዜውን ያራዝሙ</td></tr>
<tr><td>P12 (W)</td><td>የሚከራየው ስፋት ከቤቱ አይበልጥም</td><td>ስፋቱን ያስተካክሉ</td></tr></table>
<div class="box green"><b class="t">የግል ቤት — አምስት ደረጃዎች</b>1) Start፡ «Private residence» + «My house»። 2) P1 ዓምድ E ይሙሉ። 3) P4 ሁሉም PASS።
4) P3 ያንብቡና ያትሙ። 5) የቤት ዕቃ፣ ማጠናቀቂያ ወይም የብድር ጊዜ ቀይረው ውጤቱን ያወዳድሩ።</div>
</div>""")
    # ================================================================ CH7 – inputs
    S.append(f"""<div class="page">{chapter(7, "የፕሮጀክቱ ግብዓቶችና ሁኔታዎች (10, 11, 12)", "Commercial track – inputs, scenarios and the tax register", "CH07", "ክፍል ለ · የንግድ ፕሮጀክት")}
<p>የንግድ ሞዴሉ ግብዓቶች በ11 ብሎኮች ተደራጅተዋል። የግል መረጃ (የቤተሰብ ገቢ፣ ቁጠባ፣ የቤት ብድር) የለበትም — እነዚያ በግል ቤት ክፍል ብቻ ናቸው።</p>
{fig("C_10_Inputs_1.png", "ምስል 7.1 — 10 Inputs፡ E የእኔ ፕሮጀክት · F ምሳሌ 1 · G ምሳሌ 2 · H ACTIVE · K እርግጠኝነት · L ምንጭ/ቀን።",
     [(1, 43, 9.5), (2, 60, 9.5), (3, 87, 9.5), (4, 94, 9.5)],
     ["የእርስዎ ዓምድ (E)", "ACTIVE (H) — ሞዴሉ የሚጠቀመው", "እርግጠኝነት — High / Medium / Low", "ምንጭና ቀን (ለምሳሌ «BOQ ጥቅምት 2026»)"],
     width=100, crop_=(0.05, 0.62, 0, 1))}
<table><tr><th style="width:30%">ብሎክ</th><th>ይዘት</th></tr>
<tr><td>1–2 ፕሮጀክትና ፎቆች</td><td>ስም፣ ቦታ፣ FAR፣ ቀን፣ ወራት፣ ማጠናቀቂያ፤ ለ16 ወለሎች ጠቅላላ ስፋት (የተጣራ ስፋት በገጽ 13 ነው)</td></tr>
<tr><td>3 የግንባታ ወጪ</td><td>ለ9 የሥራ ክፍሎች «Rate x area» ወይም «Lump sum (BOQ)»፣ ተጨማሪ ሥራዎች፣ መጠባበቂያ፣ ዲዛይን፣ ቁጥጥር፣ <b>የልዩ ባለሙያ አማካሪ</b>፣ ፈቃድ፣ FF&amp;E፣ <b>ከመክፈቻ በፊት ወጪ</b>፣ የዋጋ ጭማሪ፣ S-ኩርባ፣ ማስያዣ</td></tr>
<tr><td>4 መሬት</td><td>ሊዝ (በጥሬ ገንዘብ)፣ የራስ መሬት፣ ወይም <b>በአጋርነት የሚገባ መሬት (NCI)</b> — ለምሳሌ ከከተማ አስተዳደር ጋር በPPP</td></tr>
<tr><td>5 ሽያጭ</td><td>የዋጋ ጭማሪ፣ ቅድመ-ሽያጭ በዓመት፣ ከግንባታ በኋላ የሽያጭ ፍጥነት፣ የክፍያ ዕቅድ፣ ኮሚሽን፣ IFRS 15 ፖሊሲ</td></tr>
<tr><td>6 ኪራይ</td><td>የመሙላት ፍጥነት (O1/O2)፣ የኪራይ ጭማሪ፣ የቁልፍ ገንዘብ፣ የአገልግሎት ክፍያ፣ የመኖሪያ ኪራይ ከተ.እ.ታ ነፃ መሆኑ</td></tr>
<tr><td>7 ሆቴል</td><td>ክፍሎች፣ አማካይ የቀን ዋጋ (ADR)፣ ይዞታ፣ የምግብና ሌላ ገቢ፣ ከገቢ የሚሰላ ኪራይ፣ ዝቅተኛ ዋስትና ኪራይ</td></tr>
<tr><td>8 ሌሎች ገቢዎች</td><td>አራት ስም ያላቸው ገቢዎች (ማስታወቂያ፣ የቴሌኮም ጣሪያ፣ ኤግዚቢሽን፣ የጎብኚ መኪና ማቆሚያ…) እና ጭማሪያቸው</td></tr>
<tr><td>9–11</td><td>የሥራ ማስኬጃ ወጪ፤ ፋይናንስ (የባለቤት ድርሻ፣ የብድር ተጨማሪ ለክፍያዎች፣ ወለድ፣ የእፎይታ ጊዜ፣ ዘመን፣ DSRA፣ cash sweep፣ DSCR ግዴታ)፤ ግብር፣ ሂሳብና ግምገማ</td></tr></table>
</div>""")
    S.append(f"""<div class="page">
<h2>እርግጠኝነትና ምንጭ</h2>
<p>ለእያንዳንዱ ግብዓት በዓምድ K «High / Medium / Low» ይምረጡ፤ በዓምድ L ምንጩን ይጻፉ። «Low» የተባሉት ቁጥር ተቆጥሮ በዳሽቦርዱ ማስጠንቀቂያና ማጠቃለያ ላይ ይታያል
(በምሳሌዎቹ {C["N_LOWCONF"]} ግብዓቶች — ለምሳሌ የሆቴል ADR እና የመውጫ ካፕ ሬት — «Low» ተደርገዋል)። ከባንክ ጋር ከመነጋገርዎ በፊት እነዚህን በገበያ ጥናት ያጠናክሩ።</p>
{fig("H_10_Inputs_2.png", "ምስል 7.2 — ብሎክ 6–8፡ የመሙላት ፍጥነት፣ የሆቴል ኦፕሬተር ኪራይና አራቱ ሌሎች ገቢዎች (ምሳሌ 2)።", width=100, crop_=(0.39, 0.75, 0, 1))}
<h2>ሁኔታዎች (11 Scenarios)</h2>
<p>አራት ሁኔታዎች አሉ — <b>Base</b> (እንደ ግብዓቱ)፣ <b>Upside</b>፣ <b>Downside</b> እና <b>Lender case</b>። እያንዳንዱ 15 ነጂዎችን ያስተካክላል
(ዋጋ፣ ኪራይ፣ ይዞታ፣ ወጪ፣ መዘግየት፣ ወለድ፣ ቅድመ-ሽያጭ፣ ካፕ ሬት…)። በመጀመሪያ ገጽ «Scenario» ሲቀይሩ ሙሉ ሞዴሉ ይሰላል፤ ንቁው ዓምድ በብርቱካናማ ይደምቃል።</p>
<h2>የግብር ሕግ መዝገብ (12 TaxLaw)</h2>
<p>CIT 30%፣ የሕንፃ የግብር እርጅና 5%፣ ኪሳራ 5 ዓመት፣ ተ.እ.ታ 15%፣ ያልተከፋፈለ ትርፍ ግብር 10%፣ የትርፍ ድርሻ ግብር 10%፣ ሕጋዊ መጠባበቂያ 5%፣ የንብረት ግብር፣ የሊዝ ሁኔታዎች —
ከአዋጁ ምንጭ ጋር። የኪራይ ተ.እ.ታ ድርሻ አሁን ከኪራይ መዝገቡ ይሰላል (የመኖሪያ ኪራይ ነፃ ከሆነ ይቀነሳል)።
<b>ከመጠቀምዎ በፊት ሁሉንም መጣኔዎች ያረጋግጡ።</b></p>
</div>""")
    # ================================================================ CH8 – spaces
    S.append(f"""<div class="page">{chapter(8, "የቦታ ዝርዝርና የኪራይ መዝገብ (13 Spaces)", "Commercial track – space schedule and rent roll", "CH08", "ክፍል ለ · የንግድ ፕሮጀክት")}
<p>የሕንፃው የሚሸጥ ወይም የሚከራይ እያንዳንዱ ቦታ በአንድ ረድፍ ይገባል — እስከ 12 ቦታዎች። ሦስት የመረጃ ብሎኮች (የእኔ ፕሮጀክት፣ ምሳሌ 1፣ ምሳሌ 2) እና
ሞዴሉ የሚጠቀመው <b>ACTIVE</b> ብሎክ አሉ።</p>
{fig("H_13_Spaces_1.png", "ምስል 8.1 — 13 Spaces፡ ምሳሌ 2 — ሱቆች፣ ማሳያዎች፣ ምግብ ቤት፣ ቢሮዎች፣ ሆቴል፣ የባለቤት ቢሮ፣ የከተማ አገልግሎት፣ መኪና ማቆሚያና መጋዘን።",
     width=100, crop_=(0.52, 1, 0, 0.52))}
<table><tr><th>ዓምድ</th><th>ትርጉም</th></tr>
<tr><td>Type (ዝርዝር)</td><td>Retail · Office · Residential · Showroom/wholesale · Food &amp; beverage · Warehouse · <b>Hotel (operator lease)</b> · Parking · <b>Owner-occupied</b> · <b>Rent-free / public service</b> · Other</td></tr>
<tr><td>Net area m²</td><td>የሚከራይ/የሚሸጥ የተጣራ ስፋት። ድምሩ ከጠቅላላው የወለል ስፋት መብለጥ የለበትም (C11)፤ ጥምርታው 50–95% መሆን አለበት (C22)</td></tr>
<tr><td>Sold % (Mixed)</td><td>በዓይነት 3 (ቅይጥ) የሚሸጠው ድርሻ። ዓይነት 1 ሁሉንም የሚሸጥ ቦታ ይሸጣል፤ ዓይነት 2 ምንም አይሸጥም</td></tr>
<tr><td>Sale price · Rent · Occupancy</td><td>የዛሬ የሽያጭ ዋጋ በm²፣ ወርሃዊ ኪራይ በm²፣ የተረጋጋ የተያዘ መጠን</td></tr></table>
<div class="box"><b class="t">ልዩ ዓይነቶች</b><b>ሆቴል</b> አይሸጥም፤ ገቢው ከብሎክ 7 ይሰላል። <b>የባለቤት ቢሮና ያለ ኪራይ ቦታ</b> አይሸጡም አይከራዩም — እንደ IAS 16 ንብረት ይያዛሉ።
የተቀረው ወጪ በተጣራ ስፋት ይከፋፈላል፡ የሚሸጠው → ክምችት (IAS 2)፣ የሚከራየው → የኢንቨስትመንት ንብረት (IAS 40)፣ የባለቤቱ → IAS 16።</div>
</div>""")
    S.append(f"""<div class="page">
<h2>ኪራዩ እንዴት ይሰላል?</h2>
<p>ለእያንዳንዱ ቦታ፡ <span class="formula">ሙሉ ኪራይ በዓመት = የሚከራይ m² × ወርሃዊ ኪራይ × 12</span> ከዚያ ሞዴሉ የኪራይ-ክብደት ያለውን
የተረጋጋ ይዞታ ያሰላል (ምሳሌ 2፡ {pct(H["OCC_STAB"])})። በመጀመሪያዎቹ ሁለት ዓመታት ተከራዮች ቀስ በቀስ ስለሚገቡ ይዞታው በ«lease-up» መጠን ይባዛል
(ምሳሌ 2፡ O1 70%፣ O2 90%፣ ከO3 ጀምሮ 100%)። ኪራዩ በዓመት 8% ያድጋል።</p>
<h2>የሆቴል ኦፕሬተር ኪራይ</h2>
<p>ሆቴሉን ኦፕሬተር ያስተዳድረዋል፤ ባለቤቱ ከሆቴሉ ጠቅላላ ገቢ መቶኛ ወይም ዝቅተኛ ዋስትና ኪራይ — ከፍ ያለውን — ያገኛል፡
<span class="formula">የክፍል ገቢ = ክፍሎች × 365 × ይዞታ × ADR ;  ጠቅላላ = የክፍል ገቢ × (1 + የምግብና ሌላ %) ;  ኪራይ = MAX(18% × ጠቅላላ, ዝቅተኛ ዋስትና)</span>
በምሳሌ 2 (85 ክፍሎች፣ ADR 12,000 ብር፣ 65% ይዞታ) በዛሬ ዋጋ የተረጋጋው የሆቴል ኪራይ <b>{m(H["HOTEL_RENT_TODAY"])} ሚሊዮን ብር</b> ነው —
ከቦታዎቹ ኪራይ ({m(H["GPR_TODAY"])} ሚ.) ጋር ሲነጻጸር ትልቅ ድርሻ። የሆቴል ወጪዎችን ኦፕሬተሩ ስለሚሸፍን የንብረት ወጪ % የሚሰላው በቦታዎቹ ኪራይ ላይ ብቻ ነው።</p>
{fig("H_24_Revenue_1.png", "ምስል 8.2 — 24 Revenue፡ የቦታዎች ኪራይ፣ የሆቴል ብሎክና አራቱ ሌሎች ገቢዎች (ምሳሌ 2)።", width=100, crop_=(0.36, 1, 0, 1))}
</div>""")
    # ================================================================ CH9 – engine
    S.append(f"""<div class="page">{chapter(9, "ወጪ፣ ወርሃዊ ግንባታና ብድር (20–30)", "Commercial track – the engine", "CH09", "ክፍል ለ · የንግድ ፕሮጀክት")}
<h2>የልማት ወጪ (22 Capex)</h2>
<p>የሕንፃ ወጪ በ9 ክፍሎች፡ የመረጃ ቋት ዋጋ × የቦታ ማስተካከያ × የፎቅ ማስተካከያ × GFA፣ ወይም BOQ ጠቅላላ ዋጋ። በምሳሌ 1 የዛሬ ዋጋ የሕንፃ ወጪ
{m(C["HARD_BASE"])} ሚ. ነው፤ ከጭማሪ፣ መጠባበቂያና ክፍያዎች በኋላ የግንባታ ወጪ {m(C["CONSTR_COST"])} ሚ.፣ መሬት {m(C["LAND_TOT"])} ሚ.፣
የብድር ወጪ (IDC ጨምሮ) {m(C["FIN_CAP"])} ሚ. — <b>ጠቅላላ የልማት ወጪ (TDC) {m(C["TDC"])} ሚሊዮን ብር</b>።</p>
<h2>ወርሃዊ ግንባታና የገንዘብ ምንጭ (23 Construction)</h2>
<p>እስከ 48 ወራት፡ S-ኩርባ፣ የዋጋ ጭማሪ፣ ወርሃዊ ወጪ፣ የልዩ ባለሙያ አማካሪ፣ ቅድመ-ሽያጭ ገቢና የገንዘብ ምንጭ ቅደም ተከተል፡
<b>(1) የቅድመ-ሽያጭ ገንዘብ → (2) የባለቤት ካፒታል → (3) የባንክ ብድር → (4) ተጨማሪ ካፒታል</b>። ከመክፈቻ በፊት ወጪ በርክክብ ወር ይወጣል፤ እንደ ወጪ ይመዘገባል ግን
እንደ ግንባታ ወጪ በገንዘብ ምንጮቹ ይሸፈናል። «Facility headroom» ግብዓት ብድሩ የባንክ ክፍያዎችንም እንዲሸፍን ያደርጋል — ስለዚህ ክፍተት አይቀርም።</p>
{fig("C_23_Construction_1.png", "ምስል 9.1 — 23 Construction (ምሳሌ 1)፡ ወርሃዊ ወጪ፣ የገንዘብ ፍላጎትና ቅድመ-ሽያጭ።", width=100, crop_=(0.06, 0.72, 0, 1))}
<h2>ብድር (27 Loan)</h2>
<p>ብድሩ በወር ለ288 ወራት ይሰላል። በግንባታ ወቅት ወለዱ (IDC) በብድሩ ላይ ይጨመራል — <span class="formula">IDC = (መነሻ ቀሪ + የወሩ ብድር) × ወለድ ÷ 12</span>
ዙሪያ-ጥገኛ ቀመር የለም። ከዚያ የእፎይታ ጊዜ (ወለድ-ብቻ ወይም ወደ ብድሩ የሚጨመር)፣ እኩል ክፍያ፣ እኩል ዋና ገንዘብ ወይም balloon፣ በወር/በሩብ/በዓመት፣
ከሽያጭ ገቢ ቅድመ ክፍያ (cash sweep) እና DSRA። ምሳሌ 1 ብድሩን በ{C["LN_FINAL_DATE"]:%m/%Y} ይጨርሳል፤ ምሳሌ 2 ደግሞ በ{H["LN_FINAL_DATE"]:%m/%Y}።</p>
</div>""")
    S.append(f"""<div class="page">
{fig("C_27_Loan_1.png", "ምስል 9.2 — 27 Loan፡ ዓመታዊ ማጠቃለያና የባንክ መረጃ (ምሳሌ 1)።", width=100)}
<table><tr><th>ገጽ</th><th>ምን ያሰላል</th></tr>
<tr><td>20 Timeline</td><td>C1–C4 የግንባታ ዓመታት፣ O1–O10 የሥራ ዓመታት፣ የኢትዮጵያ ዓመት፣ የቅናሽ ማባዣዎች፣ የወር ጠቋሚዎች</td></tr>
<tr><td>21 Area</td><td>የፎቅ ስፋቶች፣ FAR፣ የተጣራ/ጠቅላላ ጥምርታ፣ የወጪ ድልድል (IAS 2/40/16)፣ የኪራይ ተ.እ.ታ ድርሻ</td></tr>
<tr><td>24 Revenue</td><td>ቅድመ-ሽያጭና ሽያጭ (IFRS 15)፣ ኪራይ፣ ሆቴል፣ የቁልፍ ገንዘብ፣ ሌሎች ገቢዎች፣ NOI</td></tr>
<tr><td>25 Opex</td><td>የንብረት ወጪ፣ አስተዳደር፣ መድን፣ የንብረት ግብር፣ ጥገና፣ ማስታወቂያ፣ አጠቃላይ ወጪ፣ ከመክፈቻ በፊት ወጪ</td></tr>
<tr><td>26 Assets</td><td>የእርጅና ቅናሽ በክፍል (መዋቅር 50፣ ሲስተሞች 20፣ ዕቃ 7 ዓመት)፣ IAS 36 ፈተና</td></tr>
<tr><td>28 Tax</td><td>CIT፣ የግብር እርጅና፣ ኪሳራ በዓመት (FIFO፣ 5 ዓመት)፣ የዘገየ ግብር፣ ያልተከፋፈለ ትርፍ ግብር</td></tr>
<tr><td>29 VAT · 30 WorkingCap</td><td>የተ.እ.ታ ሂሳብና ተመላሽ፤ ተሰብሳቢ፣ ተከፋይ፣ ማስያዣ፣ የውል ዕዳ</td></tr></table>
<div class="box"><b class="t">እነዚህን ገጾች መቀየር አያስፈልግዎትም</b>ሞተሩ ከግብዓቶቹ ሁሉንም ያሰላል። ለማየትና ለመፈተሽ ብቻ ይክፈቷቸው — እያንዳንዱ ረድፍ ስም፣ መለኪያና ጠቅላላ ዓምድ አለው።</div>
</div>""")
    # ================================================================ CH10 – statements
    S.append(f"""<div class="page">{chapter(10, "የፋይናንስ መግለጫዎች (40–43)", "Commercial track – IFRS statements and lender ratios", "CH10", "ክፍል ለ · የንግድ ፕሮጀክት")}
<p>ሦስቱ መግለጫዎች ለ14 ጊዜያት (C1–C4፣ O1–O10) በሺህ ብር ይቀርባሉ። ሀብትና ዕዳ በየዓመቱ ሚዛናዊ ነው (C01)፣ የገንዘብ ፍሰቱ የመጨረሻ ቀሪ ከሀብትና ዕዳው ጥሬ ገንዘብ ጋር እኩል ነው (C02)።</p>
{fig("C_40_PL_1.png", "ምስል 10.1 — 40 PL፡ ገቢ፣ የሽያጭ ወጪ፣ የሥራ ማስኬጃ ወጪ (ከመክፈቻ በፊት ወጪ ጨምሮ)፣ EBITDA፣ ወለድ፣ ግብርና የተጣራ ትርፍ (ምሳሌ 1)።", width=100)}
{fig("C_42_CF_1.png", "ምስል 10.2 — 42 CF (IAS 7)፡ ከሥራ፣ ከኢንቨስትመንትና ከፋይናንስ፣ የትርፍ ድርሻና የካፒታል ተመላሽ (በDSCR ፈተና)፣ DSRA።", width=100)}
</div>""")
    S.append(f"""<div class="page">
{fig("C_41_BS_1.png", "ምስል 10.3 — 41 BS፡ የኢንቨስትመንት ንብረት፣ PPE፣ ክምችት፣ ጥሬ ገንዘብ፣ ብድር፣ የሊዝ ዕዳ፣ የዘገየ ግብር፣ የባለቤት ሀብትና NCI።", width=100)}
{fig("C_43_Ratios_1.png", "ምስል 10.4 — 43 Ratios፡ CFADS፣ DSCR፣ LLCR፣ ICR፣ gearing፣ net debt/EBITDA።", width=100)}
<p><b>DSCR</b> (የዕዳ ክፍያ ሽፋን) = ለዕዳ ክፍያ ያለ ገንዘብ ÷ የዕዳ ክፍያ። ምሳሌ 1 ዝቅተኛው {C["MIN_DSCR"]:.2f}x ነው (ግዴታው 1.30x)።
ከግዴታው በታች ከወረደ የትርፍ ድርሻ ይቆለፋል። <b>LLCR</b> በብድሩ ቀሪ ዘመን የሚገኘው ገንዘብ የአሁኑ ዋጋ ÷ የብድሩ ቀሪ ({C["LLCR_O1"]:.2f}x በO1)።</p>
</div>""")
    # ================================================================ CH11 – returns
    S.append(f"""<div class="page">{chapter(11, "ተመላሽና ውሳኔ (50–51)", "Commercial track – returns, payback and the verdict", "CH11", "ክፍል ለ · የንግድ ፕሮጀክት")}
{fig("C_50_Returns_1.png", "ምስል 11.1 — 50 Returns፡ የመጨረሻ ዋጋ፣ የፕሮጀክትና የባለቤት የገንዘብ ፍሰት፣ NPV፣ XIRR፣ ብዜት፣ GDV፣ ውሳኔ (ምሳሌ 1)።", width=100)}
<table><tr><th>መለኪያ</th><th>ትርጉም</th><th class="num">ምሳሌ 1</th><th class="num">ምሳሌ 2</th></tr>
<tr><td>የባለቤት IRR (XIRR)</td><td>ባለቤቱ ያስገባው ገንዘብ በዓመት የሚያስገኘው ተመላሽ፣ በቀን የተሰላ</td><td class="num">{pct(C["EIRR"])}</td><td class="num">{pct(H["EIRR"])}</td></tr>
<tr><td>የፕሮጀክት IRR</td><td>ብድርን ሳይቆጥር የፕሮጀክቱ ተመላሽ</td><td class="num">{pct(C["PIRR"])}</td><td class="num">{pct(H["PIRR"])}</td></tr>
<tr><td>NPV @ 18%</td><td>የወደፊት ገንዘብ የአሁን ዋጋ ሲቀነስ ወጪ</td><td class="num">{m(C["PNPV"])} ሚ.</td><td class="num">{m(H["PNPV"])} ሚ.</td></tr>
<tr><td>GDV ÷ TDC</td><td>የተጠናቀቀው ሕንፃ ዋጋ ÷ ጠቅላላ ወጪ</td><td class="num">{C["GDV"]/C["TDC"]:.2f}x</td><td class="num">{H["GDV"]/H["TDC"]:.2f}x</td></tr>
<tr><td>Profit on cost</td><td>(GDV − TDC) ÷ TDC</td><td class="num">{pct(C["POC"])}</td><td class="num">{pct(H["POC"])}</td></tr>
<tr><td>Yield on cost</td><td>የተረጋጋ NOI ÷ የተያዘው ንብረት ወጪ</td><td class="num">{pct(C["YOC"])}</td><td class="num">{pct(H["YOC"])}</td></tr>
<tr><td>የመመለሻ ጊዜ (ባለቤት)</td><td>ከግንባታ መጀመሪያ ጀምሮ በዓመት</td><td class="num">{C["PB_E"]:.1f}</td><td class="num">{H["PB_E"]:.1f}</td></tr></table>
<div class="box green"><b class="t">ውሳኔው እንዴት ይወሰናል?</b><b>VIABLE (አዋጭ)</b>፡ የባለቤት IRR ≥ ዝቅተኛ ተፈላጊ (20%)፣ profit on cost &gt; 0 እና DSCR ግዴታውን ያሟላል።
<b>MARGINAL</b>፡ IRR ከተፈላጊው በ3 ነጥብ ውስጥ ሲሆንና GDV ≥ TDC። <b>NOT VIABLE</b>፡ ሌላው ሁሉ። ሁለቱም ምሳሌዎች በመሠረታዊ ሁኔታ VIABLE ናቸው።</div>
</div>""")
    # ================================================================ CH12 – sensitivity
    S.append(f"""<div class="page">{chapter(12, "የቀጥታ ስጋትና የሚዛን ነጥብ (52 Sensitivity)", "Commercial track – live sensitivity and break-even", "CH12", "ክፍል ለ · የንግድ ፕሮጀክት")}
<p>በስሪት 2 ገጽ 52 Data Table አይጠቀምም — ሁሉም በቀመር ስለሚሰላ በExcel፣ በLibreOffice እና በGoogle Sheets ወዲያውኑ ይታደሳል። አራት ክፍሎች አሉት።</p>
{fig("C_52_Sensitivity_1.png", "ምስል 12.1 — 52 Sensitivity (ምሳሌ 1)፡ ማስተካከያዎች፣ ውጤቶች፣ የአሁን ዋጋዎች፣ ቶርኔዶና የሚዛን ነጥቦች።",
     [(1, 2, 14), (2, 2, 27), (3, 2, 41), (4, 2, 64), (5, 2, 90)],
     ["«ቢሆንስ» ማስተካከያዎች (100% = እንደ ግብዓቱ)", "ውጤቶች — ማስተካከያ ሲቀየር ወዲያውኑ ይቀየራሉ", "የቶርኔዶና የሚዛን ነጥብ ረዳት የአሁን ዋጋዎች",
      "ቶርኔዶ፡ ለ±20% የዋጋ ለውጥ (ሚሊዮን ብር)", "የሚዛን ነጥቦች ከ«headroom» ጋር"], width=100, crop_=(0.1, 1, 0, 0.62))}
<h2>1. «ቢሆንስ» ማስተካከያዎች</h2>
<p>ሰባት ሕዋሶች፡ የሽያጭ ዋጋ፣ የግንባታ ወጪ፣ ኪራይና የሆቴል ዋጋ፣ ይዞታ፣ ወለድ፣ የሥራ ማስኬጃ ወጪ (በ%) እና ተጨማሪ መዘግየት (በወር)። ለምሳሌ የግንባታ ወጪን 110% ቢያደርጉ
ሙሉ ሞዴሉ — ብድሩ፣ ግብሩ፣ መግለጫዎቹ፣ IRR — ይሰላል። <b>ሲጨርሱ ወደ 100% / 0 ይመልሱ</b>፤ ካልመለሱ ማረጋገጫ C17 እና ዳሽቦርዱ ያስጠነቅቁዎታል።</p>
</div>""")
    S.append(f"""<div class="page">
<h2>2. ቶርኔዶ — የትኛው ግምት ይበልጥ ያሰጋል?</h2>
<p>ለእያንዳንዱ ነጂ ±20% ለውጥ በፕሮጀክቱ ዋጋ (NPV) ላይ የሚያመጣው ለውጥ ይሰላል — ከግብር በፊት፣ በመጀመሪያ ደረጃ ግምት፡</p>
<table><tr><th>ነጂ</th><th>ቀመር (ሚሊዮን ብር)</th><th class="num">ምሳሌ 1</th><th class="num">ምሳሌ 2</th></tr>
<tr><td>የሽያጭ ዋጋ</td><td>±20% × PV(የሽያጭ ገቢ)</td><td class="num">±{m(0.2*C["PV_SALESNET"])}</td><td class="num">±{m(0.2*H["PV_SALESNET"])}</td></tr>
<tr><td>የቦታዎች ኪራይ</td><td>±20% × PV(ኪራይ + የመጨረሻ ዋጋ ድርሻ)</td><td class="num">±{m(0.2*C["PV_SPRENT"])}</td><td class="num">±{m(0.2*H["PV_SPRENT"])}</td></tr>
<tr><td>የሆቴል ገቢ</td><td>±20% × PV(የሆቴል ኪራይ)</td><td class="num">±{m(0.2*C["PV_HOTRENT"])}</td><td class="num">±{m(0.2*H["PV_HOTRENT"])}</td></tr>
<tr><td>የግንባታ ወጪ</td><td>∓20% × PV(የግንባታ ወጪ)</td><td class="num">∓{m(0.2*C["PV_CAPX"])}</td><td class="num">∓{m(0.2*H["PV_CAPX"])}</td></tr>
<tr><td>6 ወር መዘግየት</td><td>የገቢዎች መዘግየትና የዋጋ ጭማሪ</td><td class="num">−{m(C["DELAY6_PV"])}</td><td class="num">−{m(H["DELAY6_PV"])}</td></tr></table>
<p>ትልቁ ስጋት — ምሳሌ 1፡ <b>{C["TOP_RISK"]}</b>፤ ምሳሌ 2፡ <b>{H["TOP_RISK"]}</b>። ቶርኔዶው በትልቁ ለውጥ ተደርድሮ በዳሽቦርዱ ገበታ ላይ ይታያል።</p>
<h2>3. የሚዛን ነጥቦች (break-even)</h2>
<table><tr><th>የሚዛን ነጥብ</th><th class="num">ምሳሌ 1</th><th class="num">ምሳሌ 2</th></tr>
<tr><td>NPV = 0 የሚሆንበት የሽያጭ ዋጋ (ብር/m²)</td><td class="num">{etb(C["BE_PRICE_NPV"])} (አሁን {etb(C["BLEND_PRICE"])})</td><td class="num">— (ሽያጭ የለም)</td></tr>
<tr><td>NPV = 0 የሚሆንበት የኪራይ ደረጃ (ከግብዓቱ %)</td><td class="num">{pct(C["BE_RENT_NPV"], 0)}</td><td class="num">{pct(H["BE_RENT_NPV"], 0)}</td></tr>
<tr><td>ዝቅተኛ DSCR = ግዴታ የሚሆንበት ይዞታ</td><td class="num">{pct(C["BE_OCC"], 0)}</td><td class="num">{pct(H["BE_OCC"], 0)}</td></tr>
<tr><td>IRR ከተፈላጊው በታች ከመውረዱ በፊት የወጪ ጭማሪ</td><td class="num">{pct(C["BE_COST"], 0)}</td><td class="num">{pct(H["BE_COST"], 0)}</td></tr>
<tr><td>ዝቅተኛ DSCR = ግዴታ የሚሆንበት ወለድ</td><td class="num">{pct(C["BE_RATE"])}</td><td class="num">{pct(H["BE_RATE"])}</td></tr>
<tr><td>IRR ከተፈላጊው በታች ከመውረዱ በፊት መዘግየት (ወራት)</td><td class="num">{C["BE_DELAY"]:.0f}</td><td class="num">{H["BE_DELAY"]:.0f}</td></tr>
<tr><td>NPV = 0 የሚሆንበት የመውጫ ካፕ ሬት</td><td class="num">{"— (NPV በመጨረሻ ዋጋ ላይ አይመሠረትም)" if not C["BE_CAP"] else pct(C["BE_CAP"])}</td><td class="num">{pct(H["BE_CAP"])}</td></tr></table>
<div class="box"><b class="t">ትክክለኛውን ውጤት ለማየት</b>ቶርኔዶውና የሚዛን ነጥቦቹ ፈጣን ግምቶች ናቸው። ትክክለኛውን ለማየት ተገቢውን ማስተካከያ በገጹ ላይኛው ክፍል ይቀይሩ፤
ሁኔታዎችን ለማወዳደር ደግሞ በመጀመሪያ ገጽ «Scenario» ይቀይሩ።</div>
</div>""")
    # ================================================================ CH13 – dashboard & checks
    S.append(f"""<div class="page">{chapter(13, "ዳሽቦርድና ቼኮች (01, 60)", "Commercial track – dashboard, key findings and checks", "CH13", "ክፍል ለ · የንግድ ፕሮጀክት")}
{fig("H_01_Dashboard_1.png", "ምስል 13.1 — 01 Dashboard (ምሳሌ 2)፡ ውሳኔ፣ ስድስት ቁልፍ ቁጥሮች፣ አምስት ገበታዎች፣ ቶርኔዶ፣ የገንዘብ ምንጭና አጠቃቀም፣ ማስጠንቀቂያዎችና ዋና ዋና ግኝቶች።",
     [(1, 63, 7.7), (2, 7, 12.5), (3, 8, 23), (4, 74, 43), (5, 52, 64), (6, 8, 82)],
     ["ውሳኔ — አረንጓዴ/ቢጫ/ቀይ", "ስድስት ካርዶች፡ TDC፣ GDV፣ profit on cost፣ IRR፣ ከፍተኛ ዕዳ፣ ዝቅተኛ DSCR", "ገበታዎች፡ የገንዘብ ፍሰት፣ ብድር፣ DSCR፣ ገቢ፣ ወጪ",
      "ቶርኔዶ (ሚሊዮን ብር)", "ማስጠንቀቂያዎች (✓ / ! / ✗)", "ዋና ዋና ግኝቶች — በሞዴሉ የሚጻፍ ማጠቃለያ"], width=100)}
<h2>ዋና ዋና ግኝቶች</h2>
<p>በዳሽቦርዱ ግርጌ ሰባት ዓረፍተ ነገሮች ከቁጥሮቹ በቀጥታ ይጻፋሉ፡ ውሳኔ፣ ተመላሽ፣ ወጪና የገንዘብ ምንጭ፣ የባንክ እይታ፣ ገቢ፣ ትልቁ ስጋትና የሚዛን ነጥቦች፣
የመሬት አጋርና ዝቅተኛ እርግጠኝነት ያላቸው ግብዓቶች። ለባንክ ወይም ለባለሀብት ማቅረቢያ ዝግጁ ናቸው — ግብዓት ሲቀየር ራሳቸው ይቀየራሉ።</p>
</div>""")
    S.append(f"""<div class="page">
{fig("C_60_Checks_1.png", f"ምስል 13.2 — 60 Checks፡ {C['CHK_PASS']} / {C['CHK_TOTAL']} ማረጋገጫዎች PASS (ምሳሌ 1)።", width=88, crop_=(0, 1, 0, 0.7))}
<table><tr><th>ቡድን</th><th>ማረጋገጫዎች</th></tr>
<tr><td>ሂሳብ</td><td>C01 ሀብት = ዕዳ + ካፒታል · C02 ጥሬ ገንዘብ · C03 ምንጭ = አጠቃቀም · C04 TDC · C05 የወጪ ድልድል · C16 ተ.እ.ታ</td></tr>
<tr><td>ብድርና ግብር</td><td>C06 ብድር በጊዜው ያልቃል · C07 የዋና ገንዘብ ድምር · C12 ኪሳራ አሉታዊ አይሆንም · C10 መሬት እርጅና የለውም</td></tr>
<tr><td>ሽያጭና ቦታ</td><td>C08 ክምችት ያልቃል · C09 የውል ሂሳቦች · C11 የተጣራ ስፋት ≤ ጠቅላላ · C15 ቅድመ-ሽያጭ ≤ 100%</td></tr>
<tr><td>ቁጥጥር</td><td>C13 ምርጫዎች ትክክል · C14 የቀመር ስህተት የለም</td></tr>
<tr><td>ማስጠንቀቂያዎች (W)</td><td>C17 ማስተካከያዎች ገለልተኛ · C18 DSCR · C19 ዝቅተኛ ጥሬ ገንዘብ · C20 ክፍተት · C21 FAR · C22 የተጣራ/ጠቅላላ ጥምርታ · C23 የሆቴል ግብዓቶች · C24 ቦታ ገብቷል · C25 IAS 36</td></tr></table>
<div class="box red"><b class="t">FAIL ካዩ</b>ቁጥሮቹን አይመኑ። በጣም የተለመደው ምክንያት በቀመር ሕዋስ ላይ መጻፍ ነው። ማስጠንቀቂያ (WARN) ስህተት አይደለም — ፕሮጀክቱ
ትኩረት የሚፈልግበት ቦታ ነው (ለምሳሌ በአስቸጋሪ ሁኔታ DSCR ከግዴታው በታች መውረድ)።</div>
</div>""")
    # ================================================================ CH14 – sample 1
    S.append(f"""<div class="page">{chapter(14, "ምሳሌ 1፡ B+G+8 ቅይጥ ሕንፃ", "Worked example – Sample 1, mixed-use (type 3)", "CH14")}
<p>ምሳሌ 1 በአዲስ አበባ ቂርቆስ ክ/ከተማ ያለ 1,000 m² ቦታ ላይ የሚገነባ B+G+8 ሕንፃ ነው፡ ምድር ቤት መኪና ማቆሚያ፣ ምድርና አንደኛ ፎቅ ሱቆች (የሚከራዩ)፣
ከ2ኛ–8ኛ ፎቅ አፓርትመንቶች (የሚሸጡ)። ጠቅላላ ስፋት {etb(C["GFA"])} m²፣ የተጣራ {etb(C["NET_AREA"])} m² — {etb(C["SALE_AREA"])} m² ለሽያጭ፣
{etb(C["LET_AREA"])} m² ለኪራይ። ግንባታው {C["CONSTR_EFF"]} ወራት ነው። 70% አፓርትመንቶች በግንባታ ወቅት በቅድሚያ ይሸጣሉ።</p>
<div class="kpi"><div>TDC<b>{m(C["TDC"])} ሚ.</b></div><div>GDV<b>{m(C["GDV"])} ሚ.</b></div><div>የባለቤት IRR<b>{pct(C["EIRR"])}</b></div>
<div>ዝቅተኛ DSCR<b>{C["MIN_DSCR"]:.2f}x</b></div><div>ውሳኔ<b>{C["VERDICT"]}</b></div></div>
{fig("C_01_Dashboard_1.png", "ምስል 14.1 — የምሳሌ 1 ዳሽቦርድ (ዓይነት 3፣ መሠረታዊ ሁኔታ)።", width=100, crop_=(0, 0.62, 0, 1))}
<h2>ምን እንማራለን?</h2>
<p>• <b>የገንዘብ ምንጭ</b>፡ ከ{m(C["PEAK_FUND"])} ሚ. ፍላጎት ውስጥ የቅድመ-ሽያጭ ገንዘብ {m(C["PS_USED"])} ሚ.፣ ባለቤቱ {m(C["EQ_DRAWN"])} ሚ.፣ ባንኩ {m(C["DEBT_DRAWN"])} ሚ. ሸፍነዋል —
ቅድመ-ሽያጭ የብድር ፍላጎትን በእጅጉ ቀንሷል። • <b>የሽያጭ ገቢ</b> 40% በ«cash sweep» ብድሩን በፍጥነት ይከፍላል። • <b>ትልቁ ስጋት</b> የሽያጭ ዋጋ ነው፤
ዋጋው ከ{etb(C["BLEND_PRICE"])} ወደ {etb(C["BE_PRICE_NPV"])} ብር/m² ቢወርድ NPV ዜሮ ይሆናል። • በ«Downside» ሁኔታ ፕሮጀክቱ NOT VIABLE ይሆናል —
ከባንኩ ጋር ሲደራደሩ ይህን ያስቡበት።</p>
</div>""")
    # ================================================================ CH15 – sample 2
    S.append(f"""<div class="page">{chapter(15, "ምሳሌ 2፡ የንግድ ማዕከልና ሆቴል", "Worked example – Sample 2, trade centre with hotel and land partner (type 2)", "CH15")}
<p>ምሳሌ 2 ትልቅ B+G+10 የንግድ ማዕከል ነው ({etb(H["GFA"])} m²)፡ ሱቆችና ጋለሪያ፣ የጅምላ ማሳያዎች፣ ምግብ ቤት፣ ቢሮዎች፣ ከ7ኛ–10ኛ ፎቅ 85 ክፍሎች ያሉት ሆቴል
(በኦፕሬተር ኪራይ)፣ የባለቤቱ ዋና መሥሪያ ቤት፣ ያለ ኪራይ የሚሰጥ የከተማ አገልግሎት ቦታ፣ የሚከፈልበት መኪና ማቆሚያና መጋዘን። <b>መሬቱ ({m(H["LAND_TOT"])} ሚ.)
ከከተማ አጋር በዓይነት የገባ ነው</b> — አጋሩ የካፒታሉን {pct(H["NCI_PCT"], 0)} እንደ አናሳ ባለድርሻ (NCI) ይይዛል። ዓይነት 2፡ ሁሉም ተይዞ ይከራያል።</p>
<div class="kpi"><div>TDC<b>{m(H["TDC"])} ሚ.</b></div><div>የተረጋጋ NOI<b>{m(H["NOI_STAB"])} ሚ.</b></div><div>የባለቤት IRR<b>{pct(H["EIRR"])}</b></div>
<div>ዝቅተኛ DSCR<b>{H["MIN_DSCR"]:.2f}x</b></div><div>ውሳኔ<b>{H["VERDICT"]}</b></div></div>
<h2>ለምን እንዲህ ተዋቀረ?</h2>
<p>• <b>የመሙላት ጊዜ</b>፡ በO1 ይዞታው ዝቅተኛ ስለሆነ የመጀመሪያው ዓመት ወለድ ወደ ብድሩ ይጨመራል (capitalised grace) እና ዘመኑ 15 ዓመት ነው —
አለበለዚያ በO1 DSCR ከ1.1x በታች ይወርድ ነበር። • <b>የባለቤት ድርሻ 50%</b>፡ ብድሩ {m(H["DEBT_DRAWN"])} ሚ.፣ ከፍተኛ ዕዳ ከIDC ጋር {m(H["PEAK_DEBT"])} ሚ.።
• <b>ሆቴሉ</b> ከጠቅላላ ኪራዩ ትልቅ ድርሻ ያመጣል፤ ስለዚህ የሆቴል ገቢ በቶርኔዶው ከፍተኛ ስጋቶች አንዱ ነው። • <b>ትልቁ ስጋት</b> የግንባታ ወጪ ነው፤ IRR ከተፈላጊው በታች
ከመውረዱ በፊት ወጪው {pct(H["BE_COST"], 0)} ሊጨምር ይችላል። ኪራዩ {pct(1-H["BE_RENT_NPV"], 0)} ቢቀንስ NPV ዜሮ ይሆናል።</p>
{fig("H_21_Area_1.png", "ምስል 15.1 — 21 Area (ምሳሌ 2)፡ የተጣራ/ጠቅላላ ጥምርታ፣ የሆቴል ስፋት፣ የወጪ ድልድል በIAS 40/16 እና የኪራይ ተ.እ.ታ ድርሻ።", width=100, crop_=(0.3, 1, 0, 0.6))}
<div class="box blue"><b class="t">ሙከራ ያድርጉ</b>በመጀመሪያ ገጽ ዓይነቱን ወደ 3 (ቅይጥ) ይቀይሩ — 30% ሱቆች ይሸጣሉ፣ የባለቤት IRR {pct(0.247)} ገደማ ይሆናል።
ከዚያ «Lender case» ይምረጡና DSCR እንዴት እንደሚለወጥ ይመልከቱ።</div>
</div>""")
    # ================================================================ CH16 – own project checklist
    S.append(f"""<div class="page">{chapter(16, "የራስዎን ፕሮጀክት ማስገባት — የማረጋገጫ ዝርዝር", "Entering your own project – checklist", "CH16")}
<div class="two"><div class="pathA"><h3 style="color:#2A9D8F;margin-top:0">ሀ · የግል ቤት</h3>
<p>☐ Start፡ «Private residence» + «My house»<br>☐ P1 ብሎክ 1–2፡ ቤት፣ ቀን፣ ወራት፣ ዓይነት፣ ፎቆች<br>☐ ብሎክ 3፡ ተጨማሪ ሥራዎችና ክፍያዎች<br>
☐ ብሎክ 8፡ የመረጃ ቋቱን በወቅታዊ ዋጋ ያስተካክሉ<br>☐ ብሎክ 4–6፡ መሬት፣ ቁጠባ፣ ቤተሰብ፣ ብድር፣ ገቢ<br>☐ P4፡ ሁሉም PASS<br>☐ P3፡ ያንብቡና ያትሙ</p></div>
<div class="pathB"><h3 style="color:#1F5FAD;margin-top:0">ለ · የንግድ ፕሮጀክት</h3>
<p>☐ Start፡ «Commercial project» + ዓይነት + «My project» + «Base»<br>☐ 10 Inputs ብሎክ 1–4፡ ፕሮጀክት፣ ፎቆች፣ ወጪ፣ መሬት<br>
☐ 13 Spaces፡ እያንዳንዱ ቦታ — ዓይነት፣ ስፋት፣ ዋጋ፣ ኪራይ፣ ይዞታ<br>☐ ብሎክ 5–8፡ ሽያጭ፣ ኪራይ፣ ሆቴል፣ ሌሎች ገቢዎች<br>☐ ብሎክ 9–11፡ ወጪ፣ ፋይናንስ፣ ግብር<br>
☐ ዓምድ K–L፡ እርግጠኝነትና ምንጭ<br>☐ 12 TaxLaw፡ መጣኔዎቹን ያረጋግጡ<br>☐ 60 Checks፡ ምንም FAIL የለም<br>☐ 01 Dashboard + 52 + አራቱ ሁኔታዎች</p></div></div>
<h2>የተለመዱ ጥያቄዎች</h2>
<table><tr><th style="width:40%">ጥያቄ</th><th>መልስ</th></tr>
<tr><td>ከ12 በላይ ቦታዎች አሉኝ</td><td>ተመሳሳይ ዋጋና ኪራይ ያላቸውን ያጣምሩ (ለምሳሌ «ሱቆች F1–F2»)።</td></tr>
<tr><td>ሆቴል የለኝም</td><td>«Hotel» ዓይነት ቦታ አይጨምሩ፤ የብሎክ 7 ክፍሎችን 0 ያድርጉ (C23 ይፈትሻል)።</td></tr>
<tr><td>የመሬት አጋር የለኝም</td><td>ብሎክ 4 «Lease (cash)» ወይም «Already owned»።</td></tr>
<tr><td>ባንኩ የተለየ ወለድ በየዓመቱ ያስከፍላል</td><td>«Variable path by year» ይምረጡና ለO1–O10 ወለዱን ይጻፉ።</td></tr>
<tr><td>ውጤቱ በGoogle Sheets ቀርፋፋ ነው</td><td>ሞዴሉ ከ24,000 በላይ ቀመሮች አሉት፤ ከተቀየረ በኋላ ጥቂት ሰከንዶች ይጠብቁ።</td></tr>
<tr><td>ገጽ በስህተት ተሰበረ</td><td>ከዋናው ፋይል አዲስ ቅጂ ይውሰዱና ዓምድ E (እና ገጽ 13 የላይኛው ብሎክ) ብቻ ይቅዱ።</td></tr></table>
</div>""")
    # ================================================================ CH17 – glossary
    half = (len(GLOSSARY) + 1) // 2
    def gtab(items):
        return '<table class="gl"><tr><th>English</th><th>አማርኛ</th></tr>' + "".join(
            f"<tr><td>{en}</td><td>{am}</td></tr>" for en, am, _ in items) + "</table>"
    S.append(f"""<div class="page">{chapter(17, "መዝገበ ቃላትና ማሳሰቢያ", "Glossary and disclaimer", "CH17")}
<p class="small">ሙሉ ትርጉሞቹ በሞዴሉ ገጽ 61 Glossary ይገኛሉ።</p>
<div class="two"><div>{gtab(GLOSSARY[:half])}</div><div>{gtab(GLOSSARY[half:])}</div></div></div>""")
    S.append("""<div class="page"><h2>ማሳሰቢያ</h2>
<p>• ይህ ሞዴልና መመሪያ <b>የዕቅድ መሣሪያዎች ብቻ</b> ናቸው፤ የኢንቨስትመንት፣ የግብር፣ የሂሳብ ወይም የሕግ ምክር አይደሉም።<br>
• በምሳሌዎቹ ውስጥ ያሉት ዋጋዎች፣ ኪራዮችና ወጪዎች ለማሳያ ብቻ ናቸው — «Sample value — not market advice»።<br>
• የግብር መጣኔዎችና ሕጎች በገጽ 12 ተመዝግበዋል፤ <b>ከመጠቀምዎ በፊት ወቅታዊ መሆናቸውን ማረጋገጥ ግዴታዎ ነው</b>።<br>
• የIFRS አተገባበሩ ቀለል ተደርጎ የቀረበ ነው። ለሕጋዊ የሂሳብ መግለጫ ፈቃድ ያለው ኦዲተር ያማክሩ።<br>
• ውሳኔ ከመወሰንዎ በፊት ከባለሙያ የተዘጋጀ BOQ፣ የገበያ ጥናትና የባንክ ይፋዊ የብድር ውል ያግኙ።<br>
• ቀመሮችን መቀየር ወይም ማረጋገጫዎቹን ማጥፋት ውጤቱን ሊያበላሽ ይችላል። ሁልጊዜ ሁሉም ማረጋገጫዎች PASS መሆናቸውን ያረጋግጡ።</p>
<div class="box red"><b class="t">ኃላፊነት</b>ደራሲዎቹና አከፋፋዮቹ በዚህ ሞዴል ላይ ተመሥርቶ ለሚወሰድ ማንኛውም ውሳኔ ወይም ለሚደርስ ኪሳራ ኃላፊነት አይወስዱም።</div>
<p class="en" style="margin-top:8mm">English summary: This model and guide are planning tools only and do not constitute investment, tax, accounting or legal advice.
All sample figures are illustrative. Tax rates must be verified against current Ethiopian proclamations before use.</p>
<p class="small" style="margin-top:20mm;text-align:center">የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል · ስሪት 2.0 · Rapha Development Profitability Model</p>
</div>""")
    return f'<!doctype html><html lang="am"><head><meta charset="utf-8"><style>{CSS}{EXTRA_CSS}</style></head><body>{"".join(S)}</body></html>'


def toc_from_pdf(pdf):
    pages = {}
    with pdfplumber.open(pdf) as p:
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            for mk in re.findall(r"CH\d\d", t):
                pages.setdefault(mk, i + 1)
    return pages


if __name__ == "__main__":
    C, H, P = load_facts()
    html = os.path.join(HERE, "tutorial.html")
    pdf = os.path.join(OUT, "Rapha_Model_Tutorial_AM.pdf")
    open(html, "w").write(build(C, H, P))
    subprocess.run(["node", os.path.join(HERE, "topdf.js"), html, pdf], check=True)
    toc = toc_from_pdf(pdf)
    open(html, "w").write(build(C, H, P, toc))
    subprocess.run(["node", os.path.join(HERE, "topdf.js"), html, pdf], check=True)
    print("TOC", toc)
