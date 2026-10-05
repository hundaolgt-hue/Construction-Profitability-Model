"""Build the Amharic tutorial (HTML -> PDF via Chromium).

usage: python build_tutorial.py  (expects ../output/*.xlsx and img/ screenshots)
"""
import json
import os
import re
import subprocess
import sys

import pdfplumber

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "build"))
from facts import grab, etb, m, pct, A_NAMES, B_NAMES  # noqa: E402
from glossary import GLOSSARY  # noqa: E402

OUT = os.path.join(HERE, "..", "output")
IMG = os.path.join(HERE, "img")


def find_row(ws, label, after=1, col=2):
    for r in range(after, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(label):
            return r
    raise KeyError(label)


def load_facts():
    A, _ = grab(f"{OUT}/sampleA_fast.xlsx", A_NAMES)
    Ap, _ = grab(f"{OUT}/sampleA_premium.xlsx", A_NAMES)
    Ar, wbr = grab(f"{OUT}/sampleA_rent.xlsx", A_NAMES)
    B, wb = grab(f"{OUT}/Rapha_Development_Profitability_Model_v1.xlsx", B_NAMES)
    # rent-out facts
    ws = wbr["02_Home"]
    r = find_row(ws, "Rent income, first year")
    Ar["RENT1"] = ws.cell(r, 3).value
    Ar["YIELD"] = ws.cell(r + 1, 3).value
    Ar["PAYBACK"] = ws.cell(r + 2, 3).value
    Ar["COVER"] = ws.cell(r + 3, 3).value
    # loan worked example (first scheduled repayment)
    ln = wb["27_Loan"]
    mon0 = find_row(ln, "Timing and rate")
    rev = find_row(ln, "Repayment event flag", mon0)
    ro, rr, rrem, rp, ri = (find_row(ln, s, mon0) for s in ("Opening balance", "Interest rate (p.a.)", "Remaining payments",
                                                         "Scheduled principal", "Interest accrued"))
    rintp = find_row(ln, "Interest paid in cash", mon0)
    for c in range(5, 5 + 288):
        if ln.cell(rev, c).value == 1:
            B["L_T"] = c - 4
            B["L_OPEN"] = ln.cell(ro, c).value
            B["L_RATE"] = ln.cell(rr, c).value
            B["L_REM"] = ln.cell(rrem, c).value
            B["L_PRIN"] = ln.cell(rp, c).value
            B["L_INT_M"] = ln.cell(ri, c).value
            B["L_DATE"] = wb["20_Timeline"].cell(find_row(wb["20_Timeline"], "Month start date"), c).value
            break
    # sensitivity values
    se = wb["52_Sensitivity"]
    r0 = find_row(se, "One-way: Sale price")
    B["EIRR_P90"] = se.cell(r0 + 2 + 3, 4).value     # 0.9 row: label, formula, 0.7, 0.8, 0.9
    B["EIRR_C110"] = se.cell(find_row(se, "One-way: Construction cost") + 2 + 5, 4).value
    rs = find_row(se, "Scenario grid")
    B["SC"] = {k: (se.cell(rs + 2 + k, 4).value, se.cell(rs + 2 + k, 6).value, se.cell(rs + 2 + k, 7).value) for k in (1, 2, 3, 4)}
    return A, Ap, Ar, B


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


def fig(name, cap, callouts=(), legend=(), width=100, crop=None):
    if crop:
        from cropgrid import crop as _crop
        cwd = os.getcwd()
        os.chdir(HERE)
        name = _crop(name.replace(".png", ""), *crop)
        os.chdir(cwd)
    path = os.path.join(IMG, name)
    cos = "".join(f'<span class="co" style="left:{x}%;top:{y}%">{n}</span>' for n, x, y in callouts)
    leg = ""
    if legend:
        leg = '<ol class="legend">' + "".join(f"<li><span>{i + 1}</span>{t}</li>" for i, t in enumerate(legend)) + "</ol>"
    return (f'<figure style="width:{width}%"><div class="shot"><img src="file://{path}">{cos}</div>'
            f"<figcaption>{cap}</figcaption>{leg}</figure>")


def chapter(no, title, en, marker):
    return (f'<div class="chno">ምዕራፍ {no}<span class="marker">{marker}</span></div><h1 class="ch">{title}</h1>'
            f'<div class="en">{en}</div><div class="chrule"></div>')


def build(A, Ap, Ar, B, toc_pages=None):
    tp = toc_pages or {}
    S = []
    # ------------------------------------------------------------------ cover
    S.append(f"""<div class="cover">
<div class="badge">የአጠቃቀም መመሪያ</div><div class="badge">ስሪት 1.0</div>
<h1>የራፋ የግንባታ ፕሮጀክት<br>ትርፋማነት ሞዴል</h1>
<h2>Rapha Development Profitability Model — ደረጃ በደረጃ የአማርኛ መመሪያ</h2>
<div class="rule"></div>
<p style="font-size:12pt;color:#E6ECF2;max-width:150mm;text-align:left">ለራስዎ ቤት የግንባታ ወጪ ግምት ከማውጣት ጀምሮ፣ የንግድ ሕንፃ ትርፋማነትን፣
የብድር ክፍያ ሰንጠረዥን፣ የፋይናንስ መግለጫዎችንና የስጋት ትንተናን እስከ መሥራት — የፋይናንስ ባለሙያ ሳይሆኑ ሞዴሉን እንዴት እንደሚጠቀሙ የሚያሳይ መመሪያ።</p>
<div class="meta">ሁለት ምሳሌዎችን ይዟል፡ ምሳሌ ሀ — G+1 የቤተሰብ መኖሪያ ቤት (አዲስ አበባ) · ምሳሌ ለ — B+G+8 ቅይጥ አገልግሎት ሕንፃ (አዲስ አበባ)<br>
ፋይሎች፡ Rapha_Development_Profitability_Model_v1.xlsx · Rapha_Model_Blank.xlsx<br>
ይህ መሣሪያ ለዕቅድ ብቻ ነው፤ የኢንቨስትመንት፣ የግብር ወይም የሕግ ምክር አይደለም።</div></div>""")
    # ------------------------------------------------------------------ TOC
    toc = [("1", "መግቢያ", "CH01"), ("2", "ሁለት መንገዶች፡ የግል ቤት ወይስ የንግድ ፕሮጀክት?", "CH02"), ("3", "ከመጀመርዎ በፊት", "CH03"),
           ("4", "ደረጃ በደረጃ ግብዓት", "CH04"), ("5", "የግንባታ ወጪና የገንዘብ ፍሰት", "CH05"), ("6", "የብድር ክፍያ ሰንጠረዥ", "CH06"),
           ("7", "የፋይናንስ መግለጫዎች", "CH07"), ("8", "ትርፋማነትና ተመላሽ", "CH08"), ("9", "ስጋትና ትንተና", "CH09"),
           ("10", "ዳሽቦርድ", "CH10"), ("11", "ቼኮች", "CH11"), ("12ሀ", "ምሳሌ ሀ፡ የግል ቤት ወጪ ግምት", "CH12A"),
           ("12ለ", "ምሳሌ ለ፡ የንግድ ፕሮጀክት ትርፋማነት", "CH12B"), ("13", "ተደጋጋሚ ጥያቄዎችና መዝገበ ቃላት", "CH13"),
           ("14", "የኃላፊነት ማስተባበያ", "CH14")]
    rows = "".join(f'<tr><td style="width:14mm;color:#0F2A44;font-weight:700">{n}</td><td>{t}</td><td class="p">{tp.get(mk, "")}</td></tr>'
                   for n, t, mk in toc)
    S.append(f"""<div class="page"><h1 class="ch" style="margin-top:8mm">ማውጫ</h1><div class="chrule"></div>
<table class="toc">{rows}</table>
<div class="box blue"><b class="t">ይህን መመሪያ እንዴት ማንበብ ይቻላል?</b>
የራስዎን ቤት ብቻ መገመት ከፈለጉ፡ ምዕራፍ 1፣ 2 እና 3ን አንብበው በቀጥታ ወደ <b>ምዕራፍ 12ሀ</b> ይሂዱ (መንገድ ሀ)።
ለሽያጭ ወይም ለኪራይ የሚገነቡ ከሆነ፡ ሁሉንም ምዕራፎች በቅደም ተከተል ይከተሉ (መንገድ ለ)።
በምስሎቹ ላይ ያሉት <b style="color:#D98E2B">ብርቱካናማ ቁጥሮች</b> ከምስሉ በታች ካለው ማብራሪያ ጋር ይዛመዳሉ።</div>
<div class="box"><b class="t">ቃላት</b> እንደ IRR፣ NPV፣ DSCR፣ IFRS ያሉ የእንግሊዝኛ ቃላት በመጀመሪያ ሲጠቀሱ በአማርኛ ተብራርተዋል። ሙሉ ዝርዝሩ በምዕራፍ 13 መዝገበ ቃላትና በሞዴሉ ገጽ 61 ይገኛል።</div>
</div>""")
    # ------------------------------------------------------------------ CH1
    S.append(f"""<div class="page">{chapter(1, "መግቢያ", "Introduction – what the model answers", "CH01")}
<p>ቤት ወይም ሕንፃ ከመገንባትዎ በፊት ሁሉም ሰው የሚጠይቀው ጥያቄ አንድ ነው፡ <b>«ፕሮጀክቴ ያተርፋል?»</b> — ለራስዎ ቤት ደግሞ <b>«ስንት ያስወጣኛል፣ እንዴትስ እከፍለዋለሁ?»</b>።
የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል እነዚህን ጥያቄዎች በአንድ የExcel ፋይል ይመልሳል። እርስዎ የሚያውቋቸውን መረጃዎች (የቤቱ ስፋት፣ የማጠናቀቂያ ደረጃ፣
የብድር ሁኔታ፣ የሽያጭ ዋጋ…) ያስገባሉ፤ ሞዴሉ ቀሪውን ሁሉ በቀመር ያሰላል።</p>
<h2>ሞዴሉ የሚመልሳቸው ጥያቄዎች</h2>
<table><tr><th>ጥያቄ</th><th>መልሱ የሚገኝበት</th></tr>
<tr><td>ግንባታው በጠቅላላ ስንት ያስወጣል? በካሬ ሜትር ስንት ነው?</td><td>የቤት ግምት ገጽ (02) እና የወጪ ገጽ (22)</td></tr>
<tr><td>በየወሩ ምን ያህል ገንዘብ ያስፈልጋል? ለኮንትራክተሩ መቼ ይከፈላል?</td><td>ወርሃዊ የግንባታ ገጽ (23) እና የቤት ግምት ገጽ</td></tr>
<tr><td>ብድሩ ወርሃዊ ክፍያ ስንት ነው? መቼ ተከፍሎ ያልቃል?</td><td>የብድር ገጽ (27)</td></tr>
<tr><td>ፕሮጀክቱ ያተርፋል? ስንት በመቶ ተመላሽ ይሰጣል?</td><td>ዳሽቦርድ (01) እና የተመላሽ ገጽ (50)</td></tr>
<tr><td>ባንኩ ብድሩን ይሰጠኛል? ሽፋኑ (DSCR) በቂ ነው?</td><td>የባንክ መረጃ ገጽ (43) እና ዳሽቦርድ</td></tr>
<tr><td>ዋጋ ቢቀንስ፣ ወጪ ቢጨምር ወይም ግንባታ ቢዘገይ ምን ይሆናል?</td><td>የስጋት ትንተና ገጽ (52)</td></tr></table>
<h2>ለማን ነው?</h2>
<p><b>የግል ቤት ገንቢዎች</b> — G+0 ቪላ፣ G+1 ወይም G+2 የቤተሰብ ቤት ለሚሠሩ፤ የወጪ ግምትና የቤት ብድር ዕቅድ ብቻ ለሚፈልጉ።
<b>አልሚዎችና ባለሀብቶች</b> — ለሽያጭ (አፓርትመንት፣ ሱቅ፣ ቢሮ) ወይም ለኪራይ ለሚገነቡ። <b>ኮንትራክተር-አልሚዎች</b> እና
<b>ባንኮች/አበዳሪዎች</b> — ፕሮጀክቱ ብድሩን መመለስ ይችል እንደሆነ ለሚፈትሹ።</p>
<h2>በውስጡ ምን አለ?</h2>
<table><tr><th>ቡድን</th><th>ገጾች</th><th>ሥራቸው</th></tr>
<tr><td>መነሻ</td><td>00 መጀመሪያ</td><td>አራቱ ጥያቄዎች፣ የማረጋገጫ ሁኔታና ወደ ገጾች መሄጃ</td></tr>
<tr><td>ሪፖርት</td><td>01 ዳሽቦርድ፣ 02 የቤት ግምት</td><td>በአንድ ገጽ የሚታተሙ ማጠቃለያዎች</td></tr>
<tr><td>ግብዓት</td><td>10 ግብዓቶች፣ 11 ሁኔታዎች፣ 12 የግብር ሕግ</td><td>እርስዎ የሚሞሉት ብቸኛው ቦታ</td></tr>
<tr><td>ሞተር (Engine)</td><td>20–26</td><td>ጊዜ፣ ስፋት፣ ወጪ፣ ወርሃዊ ግንባታ፣ ገቢ፣ ወጪ፣ ንብረት</td></tr>
<tr><td>ፋይናንስ</td><td>27–30</td><td>ብድር፣ ግብር፣ ተ.እ.ታ፣ የሥራ ማስኬጃ ካፒታል</td></tr>
<tr><td>መግለጫዎች</td><td>40–43</td><td>ትርፍና ኪሳራ፣ ሀብትና ዕዳ፣ የገንዘብ ፍሰት፣ ምጣኔዎች</td></tr>
<tr><td>ትንተና</td><td>50–52</td><td>ተመላሽ፣ የመመለሻ ጊዜ፣ የስጋትና የሚዛን ነጥብ ትንተና</td></tr>
<tr><td>ቁጥጥር</td><td>60 ቼኮች፣ 61 መዝገበ ቃላት</td><td>23 ራስ-ሰር ማረጋገጫዎችና የቃላት ዝርዝር</td></tr></table>
<h2>የሥራ ፍሰት በአንድ ገጽ</h2>
<div class="flowwrap"><div class="flow"><div class="n amber"><b>1. ግብዓቶች</b>ስፋት፣ ዋጋ፣ ብድር፣ ግብር<br><span class="small" style="color:#fff">ገጽ 10–12</span></div><div class="a">➜</div>
<div class="n"><b>2. ሞተር</b>ወጪ፣ ወርሃዊ ግንባታ፣ ገቢ፣ ብድር፣ ግብር<br><span class="small" style="color:#C9D3DF">ገጽ 20–30</span></div><div class="a">➜</div>
<div class="n"><b>3. መግለጫዎች</b>ትርፍና ኪሳራ፣ ሀብትና ዕዳ፣ የገንዘብ ፍሰት<br><span class="small" style="color:#C9D3DF">ገጽ 40–43</span></div><div class="a">➜</div>
<div class="n"><b>4. ዳሽቦርድ</b>ውሳኔ፡ አዋጭ / በድንበር ላይ / አዋጭ አይደለም<br><span class="small" style="color:#C9D3DF">ገጽ 01–02, 50–52</span></div></div></div>
<p class="small">ማረጋገጫዎቹ (ገጽ 60) በሁሉም ደረጃ ስህተት እንዳይኖር ይቆጣጠራሉ። ቁጥር የሚጽፉት በደረጃ 1 ብቻ ነው።</p></div>""")
    # ------------------------------------------------------------------ CH2
    S.append(f"""<div class="page">{chapter(2, "ሁለት መንገዶች፡ የግል ቤት ወይስ የንግድ ፕሮጀክት?", "Two paths – choosing the project mode", "CH02")}
<p>ሞዴሉ ሁለት ዓይነት ተጠቃሚዎችን በአንድ ፋይል ያገለግላል። የመጀመሪያው ጥያቄ — <b>«ምን እየገነቡ ነው?»</b> — የትኛውን መንገድ እንደሚከተሉ ይወስናል።</p>
<table><tr><th>ቁጥር</th><th>ዓይነት (Mode)</th><th>ለማን</th><th>የሚያገኙት</th></tr>
<tr><td>0</td><td>የግል መኖሪያ ቤት (የወጪ ግምት ብቻ)</td><td>ለራሱ ቤት የሚሠራ</td><td>ወጪ፣ ወርሃዊ ገንዘብ፣ የክፍያ ደረጃዎች፣ የቤት ብድር፣ የመክፈል አቅም</td></tr>
<tr><td>1</td><td>ለሽያጭ መገንባት (Build-to-Sell)</td><td>አፓርትመንት/ሱቅ ሸጦ የሚወጣ አልሚ</td><td>ቅድመ-ሽያጭ፣ ሙሉ መግለጫዎች፣ ተመላሽ</td></tr>
<tr><td>2</td><td>ለኪራይ መገንባት (Build-to-Rent)</td><td>ሕንፃውን ይዞ የሚያከራይ</td><td>የኪራይ ገቢ፣ ሙሉ መግለጫዎች፣ የባንክ ምጣኔዎች</td></tr>
<tr><td>3</td><td>ቅይጥ (Mixed)</td><td>ከፊሉን ሸጦ ከፊሉን የሚያከራይ</td><td>ሁሉም — ሽያጭና ኪራይ በአንድ ላይ</td></tr></table>
<div class="two"><div class="pathA"><h3 style="color:#2A9D8F;margin-top:0">መንገድ ሀ — የግል ቤት</h3>
<p>• ጥያቄ 1፡ <b>0</b> ይምረጡ።<br>• በግብዓት ገጹ ላይ <b>ብሎክ 2 (የቤት ግምት መረጃ)</b>፣ <b>ብሎክ 3 (የፎቅ ስፋት)</b> እና <b>ብሎክ 4 (ወጪ)</b>ን ብቻ ይሙሉ።<br>
• ውጤቱን በ<b>የቤት ግምት ገጽ (02)</b> ያንብቡ — በአንድ A4 ገጽ ይታተማል።<br>• ትርፋማነት፣ IFRS፣ ግብርና የባንክ ምጣኔዎች አይታዩዎትም፤
እነዚያ ገጾች «ለግል ፕሮጀክት አያስፈልግም» የሚል ቀይ ማሳሰቢያ ያሳያሉ።<br>• በኋላ ከፊሉን ለማከራየት ካሰቡ «I may rent out part of the house» የሚለውን ያብሩ።</p>
<p class="small">➜ በቀጥታ ወደ ምዕራፍ 12ሀ መሄድ ይችላሉ።</p></div>
<div class="pathB"><h3 style="color:#1F5FAD;margin-top:0">መንገድ ለ — የንግድ ፕሮጀክት</h3>
<p>• ጥያቄ 1፡ <b>1፣ 2 ወይም 3</b> ይምረጡ።<br>• ሁሉንም የግብዓት ብሎኮች (1–11) ይሙሉ።<br>
• ውጤቱን በ<b>ዳሽቦርድ (01)</b>፣ በመግለጫዎች (40–43) እና በተመላሽ ገጽ (50) ያንብቡ።<br>
• ከባንክ ጋር ከመነጋገርዎ በፊት በሁኔታዎችና በስጋት ትንተና (ገጽ 52) ይፈትኑ።</p>
<p class="small">➜ ምዕራፍ 3–11 እና 12ለን ይከተሉ።</p></div></div>
<div class="box"><b class="t">አንድ ሞተር ለሁሉም</b>ሁለቱም መንገዶች ተመሳሳይ የወጪ ሞተር ይጠቀማሉ። ስለዚህ በዓይነት 0 የወጪ ግምት ጀምረው፣ ዓይነቱን ወደ 2 ወይም 3 ሲቀይሩ በዚያው መረጃ
ትርፋማነቱን ወዲያውኑ ያያሉ። ማረጋገጫ C16 የቤት ግምቱ ጠቅላላ ወጪ ከንግድ ሞዴሉ ወጪ ጋር እኩል መሆኑን ይፈትሻል።</div>
{fig("B_00_Start_1.png", "ምስል 2.1 — የመጀመሪያ ገጽ (00 Start)። ምሳሌ ለ በዓይነት 3 ተጭኗል።",
     [(1, 23, 28.4), (2, 23, 34), (3, 23, 39.3), (4, 23, 45), (5, 69, 22.7), (6, 6, 59)],
     ["ጥያቄ 1፡ ምን እየገነቡ ነው? (ዓይነት 0–3)", "ጥያቄ 2፡ የትኛው መረጃ — የእኔ ፕሮጀክት፣ ምሳሌ ሀ ወይም ምሳሌ ለ", "ጥያቄ 3፡ ሁኔታ (መሠረታዊ፣ ብሩህ፣ አስቸጋሪ፣ ባንክ)",
      "ጥያቄ 4፡ መግለጫዎች በብር፣ በሺህ ብር ወይስ በሚሊዮን", "የሁኔታ ሳጥን፡ ማረጋገጫዎች፣ ውሳኔ፣ ጠቅላላ ወጪ", "መሄጃ — በዓይነቱ መሠረት ይቀየራል"], crop=(0.005, 0.6, 0, 1))}
</div>""")
    # ------------------------------------------------------------------ CH3
    S.append(f"""<div class="page">{chapter(3, "ከመጀመርዎ በፊት", "Before you start", "CH03")}
<h2>የሚያስፈልግዎ</h2>
<p><b>Microsoft Excel 2016 ወይም ከዚያ በኋላ</b> (Microsoft 365 ይመረጣል)። ፋይሉ ማክሮ የለውም፤ ስለዚህ «Enable Macros» አያስፈልግም።
LibreOffice ፋይሉን ይከፍታል፣ ነገር ግን የስጋት ሰንጠረዦቹ (Data Tables) በትክክል የሚሠሩት በExcel ነው። ፋይሉን ሲከፍቱ Excel ሁሉንም ቀመሮች እንደገና ያሰላል።</p>
<h2>የቀለም ኮድ — ከሁሉ የሚበልጠው ሕግ</h2>
<table><tr><th>መልክ</th><th>ትርጉም</th><th>ምን ያደርጋሉ?</th></tr>
<tr><td style="color:#1F4E9C;background:#FFF8E6;border:1px solid #D9DEE4">ሰማያዊ ፊደል፣ ቀላል ቢጫ ጀርባ</td><td>ግብዓት</td><td>እዚህ ብቻ ይጻፉ</td></tr>
<tr><td style="color:#1B2430">ጥቁር ፊደል</td><td>ቀመር</td><td>በፍጹም አይጻፉበት</td></tr>
<tr><td style="color:#1E7B34">አረንጓዴ ፊደል</td><td>ከሌላ ገጽ የመጣ ማገናኛ</td><td>አይጻፉበት</td></tr>
<tr><td style="background:#FFF2B3;font-weight:700">ቢጫ ጀርባ፣ ደማቅ</td><td>ቁልፍ ውጤት ወይም ማረጋገጫ</td><td>ያንብቡት</td></tr></table>
<div class="box red"><b class="t">ጥቁር ሕዋስ ላይ አይጻፉ!</b>በቀመር ሕዋስ ላይ ቁጥር ከጻፉ ሞዴሉ ይሰበራል — ብዙውን ጊዜ የሀብትና ዕዳ መግለጫ ሚዛኑን ያጣል (ማረጋገጫ C01 ይወድቃል)።
በስህተት ከጻፉ Ctrl+Z ይጫኑ ወይም ከዋናው ፋይል አዲስ ቅጂ ይውሰዱ።</div>
<h2>ስሌትን ማብራት</h2>
<p>በExcel ውስጥ <b>Formulas › Calculation Options › Automatic</b> መመረጡን ያረጋግጡ። «Automatic except for data tables» ከተመረጠ፣ በገጽ 52 ያሉትን
የስጋት ሰንጠረዦች ለማደስ <b>F9</b> ይጫኑ። ሞዴሉ ዙሪያ-ጥገኛ ቀመር (circularity) የለውም፤ ስለዚህ «Enable iterative calculation» አያስፈልግም።</p>
<h2>መረጃዎን የት ያስገባሉ?</h2>
<p>በገጽ 10 (ግብዓቶች) እያንዳንዱ ግብዓት ሦስት ዓምዶች አሉት፡ <b>E — የእኔ ፕሮጀክት</b>፣ <b>F — ምሳሌ ሀ</b>፣ <b>G — ምሳሌ ለ</b>።
ዓምድ <b>H (ACTIVE)</b> በመጀመሪያ ገጽ ጥያቄ 2 የመረጡትን ዓምድ ይወስዳል፤ ሞዴሉ የሚጠቀመው ዓምድ H ብቻ ነው። ስለዚህ፡</p>
<p>1. ጥያቄ 2ን ወደ «1 – My project» ይቀይሩ። 2. በዓምድ E ቁጥሮችዎን ይጻፉ። 3. ምሳሌዎቹን ለንጽጽር እንዳሉ ይተዉዋቸው።</p>
<div class="box green"><b class="t">ጠቃሚ ልማድ</b>ከመጀመርዎ በፊት ፋይሉን በአዲስ ስም ያስቀምጡ (ለምሳሌ «የእኔ_ቤት_2027.xlsx»)። ባዶ ቅጂ ሲፈልጉ <b>Rapha_Model_Blank.xlsx</b>ን ይጠቀሙ —
ሁሉም ቁጥሮች ገለልተኛ ናቸው፣ ግን ሁሉም ማረጋገጫዎች ያልፋሉ።</div>
<h2>ቁጥሮች እንዴት ይታያሉ?</h2>
<p>አሉታዊ ቁጥሮች በቅንፍ ይታያሉ፤ ለምሳሌ (1,250) ማለት −1,250 ነው። ዜሮ «–» ሆኖ ይታያል። መቶኛ 15.0% ተብሎ ይጻፋል፤ ብዜቶች 1.30x። መግለጫዎቹ
በነባሪ በሺህ ብር (ETB '000) ናቸው — በመጀመሪያ ገጽ ጥያቄ 4 መቀየር ይችላሉ። የጊዜ ዓምዶቹ <b>C1–C4</b> (የግንባታ ዓመታት) እና <b>O1–O10</b> (የሥራ ዓመታት) ናቸው፤
ከሥራቸው የኢትዮጵያ ዘመን አቆጣጠር ዓመትም (EC) ተጽፏል።</p>
</div>""")
    # ------------------------------------------------------------------ CH4
    blocks = [
        ("1. ፕሮጀክት", "ስም፣ ቦታ፣ የቦታ ስፋት፣ FAR (የተፈቀደ የወለል ስፋት ጥምርታ)፣ የመጀመሪያ ቀን፣ የግንባታ ወራት", "ግንባታ 12–36 ወር", "ካርታ፣ የሊዝ ውል፣ የግንባታ ፈቃድ"),
        ("2. የቤት ግምት መረጃ", "የሕንፃ ዓይነት፣ ምድር ቤት፣ የማጠናቀቂያ ደረጃ፣ ገቢ፣ ቁጠባ፣ የቤተሰብ መዋጮ፣ የቤት ብድር፣ ማከራየት", "ክፍያ ≤ 35% የገቢ", "የባንክ ደብተር፣ የቤት ብድር ማመልከቻ"),
        ("3. የፎቅ ስፋት", "የእያንዳንዱ ፎቅ ጠቅላላ ስፋት (m²) እና አገልግሎቱ (መኖሪያ/ሱቅ/ቢሮ/ማቆሚያ)", "–", "የአርክቴክቸር ፕላን"),
        ("3ሐ. ቅልጥፍናና ድልድል", "የተጣራ/ጠቅላላ ጥምርታ፣ በቅይጥ ዓይነት የሚሸጠው ድርሻ", "መኖሪያ 75–85%", "አርክቴክት"),
        ("4. የግንባታ ወጪ", "በክፍል ዋጋ (ብር/m²) ወይም ዝርዝር BOQ፣ አጥር፣ የፍሳሽ ጉድጓድ፣ ታንከር፣ መጠባበቂያ፣ ክፍያዎች፣ የዋጋ ጭማሪ", "መጠባበቂያ 5–15%", "BOQ፣ የኮንትራክተር ዋጋ፣ የራፋ ኢንዴክስ"),
        ("5. መሬት", "የሊዝ ዋጋ፣ ቅድመ ክፍያ፣ የክፍያ ዓመታት፣ ወለድ ወይም በዓይነት የተዋጣ", "ቅድመ ክፍያ ≥ 10%", "የሊዝ ውል"),
        ("6. ሽያጭ", "ዋጋ በካሬ፣ የዋጋ ጭማሪ፣ ቅድመ-ሽያጭ በዓመት፣ የክፍያ ዕቅድ፣ ኮሚሽን፣ የገቢ ዕውቅና ፖሊሲ", "ቅድመ ክፍያ 20–40%", "የገበያ ጥናት፣ ደላሎች"),
        ("7. ኪራይ", "ኪራይ በካሬ በወር፣ የተያዘ መጠን፣ የኪራይ ጭማሪ፣ የቁልፍ ገንዘብ፣ የአገልግሎት ክፍያ", "የተረጋጋ 85–95%", "የአካባቢ ኪራይ ጥናት"),
        ("8. የሥራ ማስኬጃ", "የንብረት ወጪ፣ የአስተዳደር ክፍያ፣ መድን፣ የንብረት ግብር፣ ጥገና", "15–25% የኪራይ", "የሕንፃ አስተዳዳሪዎች"),
        ("9. ፋይናንስ", "የካፒታል ድርሻ፣ ወለድ፣ የእፎይታ ጊዜ፣ የብድር ዘመን፣ የአከፋፈል ዓይነት፣ ክፍያዎች፣ DSRA፣ የDSCR ግዴታ", "ወለድ 13–17%", "የባንክ የብድር ውል (term sheet)"),
        ("10. ግብርና ሂሳብ", "የገቢ ግብር፣ የእርጅና ቅናሽ፣ ኪሳራ ማሸጋገር፣ ተ.እ.ታ፣ የትርፍ ድርሻ፣ ሕጋዊ መጠባበቂያ፣ የአገልግሎት ዘመን", "CIT 30%", "የግብር አማካሪ፣ ገጽ 12"),
        ("11. ግምገማ", "የቅናሽ መጠን፣ የካፒታል ወጪ፣ የመውጫ ካፕ ሬት፣ ዝቅተኛ ተፈላጊ ተመላሽ (hurdle)", "hurdle 18–25%", "ባለሀብቱ፣ የገበያ ግብይቶች"),
    ]
    brows = "".join(f"<tr><td><b>{a}</b></td><td>{b_}</td><td>{c}</td><td>{d}</td></tr>" for a, b_, c, d in blocks)
    S.append(f"""<div class="page">{chapter(4, "ደረጃ በደረጃ ግብዓት", "Step-by-step inputs (sheet 10)", "CH04")}
<p>ሁሉም ግምቶች በገጽ 10 ላይ በ12 ብሎኮች ተደራጅተዋል። እያንዳንዱ ረድፍ የእንግሊዝኛና የአማርኛ ስም፣ መለኪያ (unit)፣ ሦስቱ የመረጃ ዓምዶች፣ ሞዴሉ የሚጠቀመው
ዓምድ H፣ የስም መለያ (range name) እና ማስታወሻ አለው። ዝርዝር ምርጫ ያላቸው ሕዋሶች (ለምሳሌ «Standard») ቀስት ሲጫኑ አማራጮችን ያሳያሉ፤ ከክልል ውጪ ቁጥር
ሲጽፉ Excel ያስጠነቅቃል።</p>
{fig("B_10_Inputs_1.png", "ምስል 4.1 — የግብዓት ገጽ (10 Inputs)፣ የመጀመሪያው ክፍል።",
     [(1, 2.5, 17.5), (2, 46, 17.5), (3, 52.5, 17.5), (4, 59, 17.5), (5, 66, 17.5), (6, 78, 17.5), (7, 1.2, 60.5)],
     ["የግብዓቱ ስም በእንግሊዝኛና በአማርኛ", "የእኔ ፕሮጀክት — ቁጥሮችዎን እዚህ ይጻፉ", "ምሳሌ ሀ (የግል ቤት)", "ምሳሌ ለ (የንግድ ሕንፃ)",
      "ACTIVE — ሞዴሉ የሚጠቀመው ዓምድ", "ማስታወሻና ምንጭ", "ብሎክ 2 — የቤት ግምት መረጃ (ለዓይነት 0)"], crop=(0, 0.43, 0, 1))}
</div><div class="page">
<h2>እያንዳንዱ ብሎክ ምን ይጠይቃል?</h2>
<table><tr><th style="width:24mm">ብሎክ</th><th>ምን ያስገባሉ</th><th style="width:26mm">የተለመደ ክልል</th><th style="width:36mm">መረጃውን ከየት ያገኛሉ</th></tr>{brows}</table>
<div class="box blue"><b class="t">የወጪ መረጃ ቋት (ብሎክ 12)</b>ከግብዓት ገጹ መጨረሻ ላይ ለ9 የግንባታ ክፍሎች በሦስት የማጠናቀቂያ ደረጃዎች (Basic, Standard, Premium) የብር በካሬ
ዋጋዎች አሉ። የማጠናቀቂያ ደረጃ ሲመርጡ ሞዴሉ ተገቢውን ዓምድ ይወስዳል። እነዚህ ናሙና ዋጋዎች ናቸው — የራፋ የግንባታ ዋጋ ኢንዴክስ ሲገኝ በእሱ ይተኩዋቸው።
ዝርዝር BOQ ካለዎት ለክፍሉ «Lump sum (BOQ)» መርጠው ጠቅላላ ዋጋውን ያስገቡ።</div>
<div class="box"><b class="t">ሁኔታዎች (ገጽ 11)</b>ከግብዓቶቹ በተጨማሪ 15 ቁልፍ ነጂዎች — የሽያጭ ዋጋ፣ ኪራይ፣ የተያዘ መጠን፣ የግንባታ ወጪ፣ መዘግየት፣ ወለድ፣ ቅድመ-ሽያጭ፣ የመውጫ ካፕ ሬት… — ለአራቱ ሁኔታዎች
ማስተካከያ አላቸው። ሞዴሉ የሚጠቀመው፡ <span class="formula">ውጤታማ ዋጋ = ግብዓት × (1 + የሁኔታው ማስተካከያ) × የማስተካከያ ሕዋስ (ገጽ 52)</span>
«መሠረታዊ» ዓምድ ዜሮ ነው፤ ማለትም የራስዎን ግብዓት እንዳለ ይጠቀማል።</div>
{fig("B_11_Scenarios_1.png", "ምስል 4.2 — የሁኔታዎች ገጽ (11 Scenarios)። ንቁው ሁኔታ በብርቱካናማ ይደምቃል።",
     [(1, 2, 44), (2, 55, 44), (3, 88, 44)],
     ["ማስተካከያ የሚደረግባቸው ነጂዎች", "የአራቱ ሁኔታዎች ማስተካከያዎች (ሰማያዊ = ሊቀየር ይችላል)", "ACTIVE — ከመጀመሪያ ገጽ ጥያቄ 3 የተመረጠው"], width=100, crop=(0, 0.62, 0, 0.95))}
</div>""")
    # ------------------------------------------------------------------ CH5
    S.append(f"""<div class="page">{chapter(5, "የግንባታ ወጪና የገንዘብ ፍሰት", "Construction cost, S-curve, presales and the funding waterfall", "CH05")}
<h2>የወጪ አሠራር (ገጽ 22)</h2>
<p>ቀጥተኛ የግንባታ ወጪ (hard cost) = ለእያንዳንዱ ክፍል <b>ዋጋ በካሬ × ጠቅላላ ስፋት × የቦታ ማስተካከያ</b>። ከG+3 በላይ ለሆኑ ሕንፃዎች የፍሬም ዋጋ በፎቅ ይጨምራል፤
ምድር ቤት ተጨማሪ የቁፋሮ ወጪ አለው። ከዚያ አጥር፣ የፍሳሽ ጉድጓድ፣ ታንከርና ንጣፍ ይደመራሉ። በምሳሌ ለ ቀጥተኛ ወጪው በበጀት ቀን ዋጋ <b>{m(B['HARD_BASE'])} ሚሊዮን ብር</b>
ሲሆን በዋጋ ጭማሪ <b>{m(B['HARD_ESC_TOT'])} ሚሊዮን</b> ይሆናል።</p>
<p>ቀጥተኛ ያልሆኑ ወጪዎች፡ መጠባበቂያ (contingency) ከተጨመረው ወጪ ላይ፣ የዲዛይንና የቁጥጥር ክፍያ በመቶኛ፣ ፈቃድና የቤት ዕቃዎች በብር። በመጨረሻ መሬትና የብድር ወጪ
(IDC — በግንባታ ጊዜ የሚታሰብ ወለድ) ተደምረው <b>ጠቅላላ የልማት ወጪ (TDC)</b> ይገኛል፡ ምሳሌ ለ — <b>{m(B['TDC'])} ሚሊዮን ብር</b> ({etb(B['TDC_M2G'])} ብር በጠቅላላ ካሬ)።</p>
{fig("B_22_Capex_1.png", "ምስል 5.1 — የወጪ ገጽ (22 Capex)።",
     [(1, 1.5, 32.5), (2, 31, 52.5), (3, 43, 52.5), (4, 85, 52.5), (5, 1.5, 84.5), (6, 30, 97)],
     ["የወጪ መወሰኛዎች፡ የማጠናቀቂያ ደረጃ፣ የቦታና የፎቅ ማስተካከያ", "የመጨረሻው በጀት በክፍል (ቢጫ)", "ከመረጃ ቋቱ የመጣ ዋጋ በካሬ",
      "ዝርዝር BOQ ሲመረጥ የሚገባው ጠቅላላ ዋጋ", "ተጨማሪ የውጭ ሥራዎች", "ጠቅላላ ቀጥተኛ ወጪ በበጀት ቀን ዋጋ"], width=100, crop=(0, 0.64, 0, 1))}
</div><div class="page">
<h2>S-ኩርባና ወርሃዊ ገንዘብ (ገጽ 23)</h2>
<p>ግንባታ በየወሩ እኩል ገንዘብ አይፈልግም — በመጀመሪያና በመጨረሻ ወራት ትንሽ፣ በመሃል ብዙ። ይህ ቅርጽ <b>S-ኩርባ</b> ይባላል። ሞዴሉ ለእያንዳንዱ ወር የሥራ ድርሻ ያሰላል፣
በዚያ ወር የዋጋ ጭማሪ ያባዛል፣ ከዚያ ዲዛይን፣ ቁጥጥር፣ ፈቃድ፣ የሊዝ ክፍያና ተ.እ.ታን ጨምሮ <b>«ጠቅላላ የገንዘብ ፍላጎት»</b>ን ያወጣል። ከፍተኛው ድምር ፍላጎት በምሳሌ ለ
<b>{m(B['PEAK_FUND'])} ሚሊዮን ብር</b> ነው። ከኮንትራክተሩ ክፍያ 5% (retention) ተይዞ ከግንባታ በኋላ ይለቀቃል።</p>
<h2>ቅድመ-ሽያጭ</h2>
<p>በምሳሌ ለ ከአፓርትመንቶቹ {pct(B['PRESOLD'], 0)} ግንባታው ሳያልቅ ይሸጣሉ። ገዢዎች በክፍያ ዕቅዱ መሠረት ይከፍላሉ፡ ውል ሲፈርሙ 30% ቅድመ ክፍያ፣ ግንባታው ሲገፋ 40%፣ በርክክብ 30%።
በIFRS 15 መሠረት ገቢው የሚታወቀው በርክክብ ጊዜ ነው (ፖሊሲውን «Over time» ማድረግ ይቻላል)፤ እስከዚያ የተሰበሰበው ገንዘብ <b>ከደንበኞች የተቀበለ ቅድሚያ (contract liability)</b>
ተብሎ እንደ ዕዳ ይመዘገባል።</p>
<h2>የገንዘብ ምንጭ ቅደም ተከተል (funding waterfall)</h2>
<p>በየወሩ ፍላጎቱ በዚህ ቅደም ተከተል ይሸፈናል፡ <b>(1)</b> ጥቅም ላይ እንዲውል የተፈቀደ የቅድመ-ሽያጭ ገንዘብ (ቀሪው በዝግ ሂሳብ ይቀመጣል) → <b>(2)</b> የባለቤት ካፒታል እስከ ቃል ኪዳኑ
→ <b>(3)</b> የባንክ ብድር እስከ ጣሪያው → <b>(4)</b> ያልተሸፈነው «የገንዘብ ክፍተት» ሆኖ በባለቤቱ ተጨማሪ ይሸፈናል። በምሳሌ ለ፡ ካፒታል {m(B['EQ_DRAWN'])}፣ ቅድመ-ሽያጭ
{m(B['PS_USED'])}፣ ብድር {m(B['DEBT_DRAWN'])} ሚሊዮን ብር፤ ክፍተት {etb(B['FUND_GAP'])}።</p>
{fig("B_23_Construction_1.png", "ምስል 5.2 — ወርሃዊ የግንባታ ገጽ (23 Construction) የመጀመሪያ ገጽ — ወራት 1–12።",
     [(1, 1.2, 31.5), (2, 1.2, 46), (3, 1.2, 70.5), (4, 1.2, 77.5)],
     ['S-ኩርባ፡ ድምርና ወርሃዊ የሥራ ድርሻ፣ የዋጋ ጭማሪ', 'የወጪ ክምችት በክፍል (IDC ጨምሮ)', 'ጠቅላላ የወጪ ክምችት', 'ለግንባታ የሚያስፈልግ ጥሬ ገንዘብ (ተ.እ.ታ፣ ሊዝ፣ ክፍያዎች) → ጠቅላላ ፍላጎት'], width=100, crop=(0, 0.655, 0, 0.75))}
<div class="box"><b class="t">ዙሪያ-ጥገኛ ቀመር የለም</b>የብድር ማመቻቻ ክፍያ የሚሰላው IDCን ባላካተተ የብድር መጠን ላይ ነው፤ IDC ደግሞ በየወሩ በመነሻ ቀሪ ሂሳብና በዚያ ወር በወጣው ብድር ላይ ይሰላል።
ስለዚህ Excel «circular reference» አያሳይም።</div>
</div>""")
    # ------------------------------------------------------------------ CH6
    L_int_q = B["L_OPEN"] * B["L_RATE"] / 4
    pmt_q = B["L_PRIN"] + L_int_q
    S.append(f"""<div class="page">{chapter(6, "የብድር ክፍያ ሰንጠረዥ", "Loan amortization schedule (sheet 27)", "CH06")}
<h2>ቁልፍ ጽንሰ-ሐሳቦች</h2>
<table><tr><th style="width:42mm">ቃል</th><th>ትርጉም</th></tr>
<tr><td>IDC — በግንባታ ጊዜ የሚታሰብ ወለድ</td><td>በግንባታ ወቅት ወለድ አይከፈልም፤ በብድሩ ላይ ይደመራል (ይካፈላል)። ወሩ ወለድ = (መነሻ ቀሪ + የዚያ ወር ብድር) × ዓመታዊ ወለድ ÷ 12።</td></tr>
<tr><td>የእፎይታ ጊዜ (grace)</td><td>ግንባታ ካለቀ በኋላ ዋና ገንዘብ የማይከፈልበት ጊዜ። ወለዱ ይከፈላል (interest-only) ወይም ይደመራል (capitalised)።</td></tr>
<tr><td>የብድር ዘመን (tenor)</td><td>ከግንባታ መጠናቀቅ ጀምሮ እፎይታን ጨምሮ ብድሩ ተከፍሎ የሚያልቅበት ጊዜ (እስከ 20 ዓመት)።</td></tr>
<tr><td>እኩል ክፍያ (annuity)</td><td>በየክፍያው ወለድ + ዋና እኩል ነው። መጀመሪያ ብዙው ወለድ፣ በኋላ ብዙው ዋና ገንዘብ ነው። Excel PMT ቀመር።</td></tr>
<tr><td>እኩል ዋና ገንዘብ</td><td>ዋና ገንዘቡ በእኩል ክፍሎች ይከፈላል፤ ወለዱ ስለሚቀንስ ጠቅላላ ክፍያ በጊዜ ይቀንሳል።</td></tr>
<tr><td>የመጨረሻ ትልቅ ክፍያ (balloon)</td><td>የተወሰነው መቶኛ እስከ መጨረሻው ክፍያ ይቆያል።</td></tr>
<tr><td>DSCR — የዕዳ ክፍያ ሽፋን</td><td>ለዕዳ ክፍያ የሚገኝ ገንዘብ (CFADS) ÷ የዓመቱ የዕዳ ክፍያ። 1.30x ማለት ለእያንዳንዱ 1 ብር ክፍያ 1.30 ብር አለ። ባንኮች ዝቅተኛ ገደብ ያስቀምጣሉ።</td></tr>
<tr><td>ትርፍ ገንዘብ ለብድር (cash sweep)</td><td>ከሽያጭ የሚገኘውን ገንዘብ የተወሰነ መቶኛ ብድሩን ቀድሞ ለመክፈል ማዋል (በምሳሌ ለ 40%)።</td></tr>
<tr><td>DSRA</td><td>የሚቀጥሉትን ወራት የዕዳ ክፍያ የሚሸፍን በባንክ የሚያዝ መጠባበቂያ ሂሳብ።</td></tr></table>
<p>የብድር ገጹ ከላይ <b>ዓመታዊ ማጠቃለያ</b> እና <b>የባንክ ማጠቃለያ</b> (ጠቅላላ ወለድ፣ አማካይ ዘመን፣ የመጨረሻ ክፍያ ቀን፣ ዝቅተኛና አማካይ DSCR፣ LLCR፣ ICR፣ ከፍተኛ የብድር ድርሻ) ያሳያል፤
ከታች ደግሞ ለ288 ወራት (48 የግንባታ + 240 የክፍያ) ወርሃዊ ሰንጠረዥ አለ።</p>
{fig("B_27_Loan_1.png", "ምስል 6.1 — የብድር ገጽ (27 Loan)፡ ዓመታዊ ማጠቃለያና የባንክ ማጠቃለያ።",
     [(1, 1.2, 24), (2, 1.2, 31.3), (3, 1.2, 39.6), (4, 1.2, 44.3), (5, 1.2, 53.6), (6, 1.2, 92)],
     ['የወጣ ብድርና የተደመረ ወለድ (IDC)', 'የተከፈለ ወለድ፣ ዋና ገንዘብና ቅድመ ክፍያ (sweep)', 'የዓመቱ መጨረሻ የብድር ቀሪ', 'የመሬት ሊዝ ክፍያና ዕዳ', 'የብድር ማጠቃለያ፡ ብዛት፣ ወለድ፣ ከፍተኛ ብድር፣ የመጀመሪያ ክፍያ', 'DSCR፣ LLCR፣ ICR እና የብድር ድርሻ'], width=100, crop=(0.01, 0.96, 0, 1))}
</div><div class="page">
<h2>በእጅ የተሰላ ምሳሌ — የመጀመሪያው የሩብ ዓመት ክፍያ (ምሳሌ ለ)</h2>
<p>ምሳሌ ለ የ{B['CONSTR_EFF']} ወር ግንባታ፣ የ12 ወር እፎይታ፣ 10 ዓመት የብድር ዘመንና የሩብ ዓመት እኩል ክፍያ አለው። ስለዚህ የመጀመሪያው ዋና ገንዘብ ክፍያ በወር
{B['L_T']} ({B['L_DATE']:%B %Y}) ይደርሳል። በዚያን ጊዜ፡</p>
<table><tr><th>ደረጃ</th><th>ስሌት</th><th class="num">ውጤት (ብር)</th></tr>
<tr><td>1. የብድሩ ቀሪ (ከቅድመ ክፍያዎች በኋላ)</td><td>ከሞዴሉ</td><td class="num">{etb(B['L_OPEN'])}</td></tr>
<tr><td>2. የሩብ ዓመት የወለድ መጠን</td><td>{pct(B['L_RATE'])} ÷ 4</td><td class="num">{B['L_RATE'] / 4:.4%}</td></tr>
<tr><td>3. የቀሩ የክፍያዎች ብዛት</td><td>(120 − 12) ÷ 3</td><td class="num">{B['L_REM']}</td></tr>
<tr><td>4. እኩል ክፍያ (ወለድ + ዋና)</td><td>PMT({B['L_RATE'] / 4:.4%}, {B['L_REM']}, −{etb(B['L_OPEN'])})</td><td class="num">{etb(pmt_q)}</td></tr>
<tr><td>5. የሩብ ዓመቱ ወለድ</td><td>{etb(B['L_OPEN'])} × {B['L_RATE'] / 4:.4%}</td><td class="num">{etb(L_int_q)}</td></tr>
<tr><td>6. ዋና ገንዘብ ክፍያ</td><td>4 − 5</td><td class="num"><b>{etb(B['L_PRIN'])}</b></td></tr></table>
<p>የሞዴሉ የወር {B['L_T']} «Scheduled principal» ሕዋስ በትክክል {etb(B['L_PRIN'])} ብር ያሳያል። ወለዱ በየወሩ ይከፈላል (በወር ≈ {etb(B['L_INT_M'])} ብር)፣ ዋና ገንዘቡ በየሩብ ዓመቱ።
ቀጣዩ ክፍያ በአዲሱ (ያነሰ) ቀሪ ላይ እንደገና ይሰላል — ስለዚህ ወለድ ቢቀየር ወይም ቅድመ ክፍያ ቢደረግ ሰንጠረዡ ራሱን ያስተካክላል።</p>
<div class="box blue"><b class="t">ለቤት ብድር</b>ምሳሌ ሀ ወርሃዊ ክፍያ ይጠቀማል፡ የ{etb(A['DEBT_DRAWN'])} ብር ብድር (IDC ጨምሮ በግንባታ መጨረሻ {etb(A['PEAK_DEBT'])})፣ {pct(A['RATE_EFF'])} ወለድ፣ {A['TENOR_Y']} ዓመት →
<b>ወርሃዊ ክፍያ {etb(A['LN_INSTAL_M'])} ብር</b>። በExcel፡ =PMT({pct(A['RATE_EFF'])}/12, {int(A['TENOR_Y']) * 12}, −{etb(A['PEAK_DEBT'])})። ጠቅላላ የሚከፈል ወለድ {etb(A['LN_INT_TOT'])} ብር ነው።</div>
<h2>DSCR እንዴት ይነበባል?</h2>
<p>በምሳሌ ለ ዝቅተኛው DSCR <b>{B['MIN_DSCR']:.2f}x</b>፣ አማካዩ {B['AVG_DSCR']:.2f}x ነው — ከባንኩ ገደብ {B['DSCR_COV']:.2f}x በላይ። DSCR ከገደቡ በታች ከወረደ ሞዴሉ ለባለአክሲዮኖች የሚከፈለውን
የትርፍ ድርሻ በዚያ ዓመት ያግዳል (lock-up)፣ ማረጋገጫ C18 ደግሞ ማስጠንቀቂያ ያሳያል። <b>LLCR</b> (በቀሪው የብድር ዘመን የገንዘብ ፍሰት የአሁን ዋጋ ÷ የብድር ቀሪ) በመጀመሪያው የሥራ ዓመት
{B['LLCR_O1']:.2f}x ነው።</p>
</div>""")
    # ------------------------------------------------------------------ CH7
    S.append(f"""<div class="page">{chapter(7, "የፋይናንስ መግለጫዎች", "Reading the P&L, balance sheet and cash flow (sheets 40–42)", "CH07")}
<p>ሦስቱ መግለጫዎች ለእያንዳንዱ ጊዜ (C1–C4፣ O1–O10) ይዘጋጃሉ። በIFRS መሠረት፡ የሚሸጡ ቤቶች <b>ለሽያጭ የተያዘ ንብረት (IAS 2)</b>፣ ለኪራይ የሚቀሩት
<b>ለኪራይ የተያዘ ንብረት (IAS 40)</b>፣ ባለቤቱ የሚጠቀምበት <b>ንብረት (IAS 16)</b> ሆነው ይመዘገባሉ። ወጪው በተጣራ ስፋት መጠን ይከፋፈላል።</p>
<h2>1. የትርፍና ኪሳራ መግለጫ (P&L)</h2>
<p>ከላይ ገቢ (ሽያጭ፣ ኪራይ፣ የቁልፍ ገንዘብ፣ የአገልግሎት ክፍያ)፣ ከዚያ የተሸጡ ቤቶች ወጪ፣ የሥራ ማስኬጃ ወጪ፣ <b>EBITDA</b> (ከወለድ፣ ግብርና እርጅና በፊት ትርፍ)፣
የእርጅና ቅናሽ፣ የወለድ ወጪ፣ ከግብር በፊት ትርፍ፣ ግብሮች፣ በመጨረሻ <b>ከግብር በኋላ ትርፍ</b>። የሽያጭ ገቢ የሚታወቀው በርክክብ ዓመት ስለሆነ በግንባታው የመጨረሻ ዓመት
ትልቅ ገቢና ተመጣጣኝ የወጪ ልቀት ይታያል።</p>
{fig("B_40_PL_1.png", "ምስል 7.1 — የትርፍና ኪሳራ መግለጫ (40 PL)፣ በሺህ ብር።",
     [(1, 1, 21), (2, 1, 34), (3, 1, 41), (4, 1, 61.5), (5, 1, 74.5), (6, 1, 81)],
     ["ገቢ በIFRS 15/16", "የተሸጡ ቤቶች ወጪ (ከክምችት የሚለቀቅ)", "የሥራ ማስኬጃ ወጪዎች", "EBITDA", "ከግብር በፊት ትርፍ", "ከግብር በኋላ ትርፍና ክፍፍል"], crop=(0, 1, 0, 1))}
</div><div class="page">
<h2>2. የሀብትና ዕዳ መግለጫ (Balance sheet)</h2>
<p>በእያንዳንዱ ጊዜ መጨረሻ ያለውን ያሳያል፡ <b>ሀብት</b> (ሕንፃ፣ ክምችት፣ ተሰብሳቢ፣ ጥሬ ገንዘብ) = <b>ዕዳ</b> (ብድር፣ የሊዝ ዕዳ፣ ከደንበኞች ቅድሚያ፣ ግብር) + <b>የባለቤት ሀብት</b>
(ካፒታል፣ ሕጋዊ መጠባበቂያ፣ ያልተከፋፈለ ትርፍ)። የመጨረሻው ረድፍ «Balance check» ሁልጊዜ ዜሮ መሆን አለበት።</p>
{fig("B_41_BS_1.png", "ምስል 7.2 — የሀብትና ዕዳ መግለጫ (41 BS)።",
     [(1, 1, 25.5), (2, 1, 34.5), (3, 1, 54.5), (4, 1, 68), (5, 1, 77.5), (6, 1, 93)],
     ["ቋሚ ሀብት (IAS 40/16)", "ክምችት፣ ተሰብሳቢና ጥሬ ገንዘብ", "የባንክ ብድርና የሊዝ ዕዳ", "ከደንበኞች ቅድሚያ (IFRS 15)", "የባለቤት ሀብት", "የሚዛን ማረጋገጫ — ዜሮ"], width=100, crop=(0, 1, 0, 1))}
</div><div class="page">
<h2>3. የጥሬ ገንዘብ ፍሰት መግለጫ (IAS 7፣ ቀጥተኛ ያልሆነ ዘዴ)</h2>
<p>ከግብር በፊት ትርፍ ተነስቶ ገንዘብ ያልሆኑ ነገሮች (የእርጅና ቅናሽ፣ የወጪ ልቀት) ይመለሳሉ፣ የሥራ ማስኬጃ ካፒታል ለውጦችና የተከፈለ ግብር ይቀነሳሉ → <b>ከሥራ እንቅስቃሴ</b>።
ለኪራይ የሚቀረው ሕንፃ ወጪ <b>ከኢንቨስትመንት</b> ሥር፣ ካፒታል፣ ብድር፣ የብድር ክፍያና የትርፍ ድርሻ <b>ከፋይናንስ</b> ሥር ይመጣሉ። የመጨረሻው ጥሬ ገንዘብ በሀብትና ዕዳ መግለጫው ካለው ጋር
እኩል መሆን አለበት (ማረጋገጫ C02)።</p>
{fig("B_42_CF_1.png", "ምስል 7.3 — የጥሬ ገንዘብ ፍሰት መግለጫ (42 CF)።",
     [(1, 1, 22.5), (2, 1, 49), (3, 1, 55.5), (4, 1, 78), (5, 1, 93.5)],
     ["ከሥራ እንቅስቃሴ", "ከኢንቨስትመንት", "ከፋይናንስ፡ ካፒታል፣ ብድር፣ ክፍያዎች", "የትርፍ ድርሻና የካፒታል ተመላሽ", "የመጨረሻ ጥሬ ገንዘብ"], width=100, crop=(0, 1, 0, 1))}
<div class="box"><b class="t">ለምን መመጣጠን አለባቸው?</b>እያንዳንዱ ብር ከአንድ ቦታ መጥቶ ወደ አንድ ቦታ ይሄዳል። ብድር ሲወጣ ጥሬ ገንዘብ ይጨምራል፣ ዕዳም ይጨምራል። ቤት ሲሸጥ ገቢ ትርፍን፣
ጥሬ ገንዘብ ሀብትን ይጨምራል። ሦስቱ መግለጫዎች በዚህ መንገድ ስለተሳሰሩ አንዱ ቢሳሳት ሚዛኑ ይጠፋል — ስለዚህ ሚዛኑ የሞዴሉ ጤንነት ዋና ምልክት ነው።</div>
</div>""")
    # ------------------------------------------------------------------ CH8
    S.append(f"""<div class="page">{chapter(8, "ትርፋማነትና ተመላሽ", "Profitability and returns – NPV, IRR, MOIC, payback (sheets 50–51)", "CH08")}
<p>የዕለት ተዕለት ምሳሌ እንውሰድ፡ ዛሬ 100 ብር ቢያበድሩ እና በዓመት 120 ብር ቢመለስልዎ ተመላሹ 20% ነው። ነገር ግን ዛሬ ያለ 100 ብር ከዓመት በኋላ ከሚገኝ 100 ብር ይበልጣል —
ምክንያቱም ዛሬ ያለውን ገንዘብ በባንክ ማስቀመጥ ወይም በሌላ ሥራ ማዋል ይቻላል። የፕሮጀክት ትንተና ይህን የጊዜ ዋጋ ግምት ውስጥ ያስገባል።</p>
<table><tr><th style="width:40mm">መለኪያ</th><th>በቀላል አማርኛ</th><th class="num" style="width:26mm">ምሳሌ ለ</th></tr>
<tr><td><b>NPV</b> — የተጣራ የአሁን ዋጋ</td><td>ሁሉም የወደፊት ገቢና ወጪ በቅናሽ መጠን ወደ ዛሬ ዋጋ ተቀይሮ ሲደመር። ከዜሮ በላይ = ከሚጠበቀው ተመላሽ በላይ ያስገኛል።</td><td class="num">{m(B['PNPV'])} ሚ.</td></tr>
<tr><td><b>IRR</b> — የውስጥ ተመላሽ መጠን</td><td>NPV ዜሮ የሚሆንበት ዓመታዊ ተመላሽ — «ገንዘቤ በዓመት ስንት በመቶ ያድጋል?»። ቀኖቹን በትክክል በመጠቀም (XIRR) ይሰላል።</td><td class="num">ፕሮጀክት {pct(B['PIRR'])}<br>ባለቤት {pct(B['EIRR'])}</td></tr>
<tr><td><b>MOIC</b> — የካፒታል ብዜት</td><td>ባለቤቱ ያገኘው ጠቅላላ ገንዘብ ÷ ያዋጣው። 2.0x ማለት ገንዘቡ በእጥፍ ተመለሰ።</td><td class="num">{B['EMOIC']:.2f}x</td></tr>
<tr><td><b>ትርፍ ከወጪ አንጻር</b></td><td>(GDV − TDC) ÷ TDC። GDV = የሚሸጡ ቤቶች ገቢ + የሚቀረው ንብረት ዋጋ (የተረጋጋ NOI ÷ ካፕ ሬት)።</td><td class="num">{pct(B['POC'])}</td></tr>
<tr><td><b>የልማት ትርፍ ህዳግ</b></td><td>(GDV − TDC) ÷ GDV</td><td class="num">{pct(B['DEV_MARGIN'])}</td></tr>
<tr><td><b>ከወጪ አንጻር ዓመታዊ ገቢ</b></td><td>የተረጋጋ NOI ÷ ለኪራይ የተያዘው ወጪ</td><td class="num">{pct(B['YOC'])}</td></tr>
<tr><td><b>የመመለሻ ጊዜ</b></td><td>የተጠራቀመ ገንዘብ ፍሰት ዜሮ የሚያልፍበት ጊዜ (ከግንባታ መጀመሪያ፣ በዓመት ክፍልፋይ)። በቅናሽም ይሰላል።</td><td class="num">ፕሮጀክት {B['PB_P']:.1f} ዓ.<br>ባለቤት {B['PB_E']:.1f} ዓ.</td></tr></table>
<div class="box"><b class="t">የፕሮጀክት እና የባለቤት ተመላሽ ልዩነት</b><b>የፕሮጀክት ገንዘብ ፍሰት</b> ብድርን ሳይጨምር ፕሮጀክቱ ራሱ የሚያስገኘውን ያሳያል (ከቅናሽ መጠን {pct(B['PIRR'] * 0 + 0.18, 0)} ጋር ይነጻጸራል)።
<b>የባለቤት ገንዘብ ፍሰት</b> = − የተዋጣ ካፒታል + የትርፍ ድርሻና የካፒታል ተመላሽ + በ10ኛው ዓመት የባለቤቱ ድርሻ ዋጋ። ውሳኔው (VIABLE) በባለቤት IRR ከዝቅተኛ ተፈላጊ ተመላሽ ({pct(B['HURDLE'], 0)}) ጋር ይወሰናል።</div>
{fig("B_50_Returns_1.png", "ምስል 8.1 — የተመላሽ ገጽ (50 Returns)።",
     [(1, 1, 17.5), (2, 1, 29.5), (3, 1, 42), (4, 1, 61), (5, 1, 76)],
     ["የመጨረሻ ዋጋ (Terminal value) በ10ኛው ዓመት", "የፕሮጀክት ገንዘብ ፍሰት", "የባለቤት ገንዘብ ፍሰት", "NPV፣ IRR፣ MOIC", "GDV፣ ትርፍ ከወጪና ውሳኔ"], width=100, crop=(0, 1, 0, 1))}
</div>""")
    # ------------------------------------------------------------------ CH9
    sc = B["SC"]
    S.append(f"""<div class="page">{chapter(9, "ስጋትና ትንተና", "Scenarios, flex cells, sensitivity, tornado and break-even (sheet 52)", "CH09")}
<p>ባንክ ወይም ባለሀብት ፊት ከመቅረብዎ በፊት ፕሮጀክቱን «ይፈትኑት» (stress-test)፡ ነገሮች እንደታሰበው ባይሄዱስ? ሞዴሉ አራት መሣሪያዎች አሉት።</p>
<h3>1. ሁኔታዎች (Scenarios)</h3>
<p>በመጀመሪያ ገጽ ጥያቄ 3 ላይ ሁኔታውን ይቀይሩ። ሁሉም ውጤቶች ወዲያው ይቀየራሉ። የገጽ 52 «Scenario grid» አራቱንም ጎን ለጎን ያሳያል፡</p>
<table><tr><th>ሁኔታ</th><th class="num">ባለቤት IRR</th><th class="num">ትርፍ ከወጪ</th><th class="num">ዝቅተኛ DSCR</th></tr>
<tr><td>መሠረታዊ (Base)</td><td class="num">{pct(sc[1][0])}</td><td class="num">{pct(sc[1][1])}</td><td class="num">{sc[1][2]:.2f}x</td></tr>
<tr><td>ብሩህ (Upside)</td><td class="num">{pct(sc[2][0])}</td><td class="num">{pct(sc[2][1])}</td><td class="num">{sc[2][2]:.2f}x</td></tr>
<tr><td>አስቸጋሪ (Downside)</td><td class="num">{pct(sc[3][0])}</td><td class="num">{pct(sc[3][1])}</td><td class="num">{sc[3][2]:.2f}x</td></tr>
<tr><td>የባንክ (Lender)</td><td class="num">{pct(sc[4][0])}</td><td class="num">{pct(sc[4][1])}</td><td class="num">{sc[4][2]:.2f}x</td></tr></table>
<h3>2. የማስተካከያ ሕዋሶች (Flex cells)</h3>
<p>በገጽ 52 ላይ ለሽያጭ ዋጋ፣ ለግንባታ ወጪ፣ ለኪራይ፣ ለተያዘ መጠን፣ ለወለድና ለሥራ ማስኬጃ ወጪ 100% የሆኑ ሕዋሶች አሉ። ለምሳሌ የሽያጭ ዋጋን 90% ቢያደርጉ
ባለቤት IRR ከ{pct(B['EIRR'])} ወደ <b>{pct(B['EIRR_P90'])}</b> ይወርዳል። ተጨማሪ የግንባታ መዘግየትም በወር ማስገባት ይቻላል። ከሙከራ በኋላ ወደ 100% መመለስዎን አይርሱ!</p>
<h3>3. የተጽዕኖ ሰንጠረዦችና ቶርኔዶ</h3>
<p>የExcel «Data Tables» እያንዳንዱን ነጂ ከ−30% እስከ +30% ቀይረው ባለቤት IRR፣ ዝቅተኛ DSCR እና ትርፍ ከወጪን ያሳያሉ። ባለሁለት አቅጣጫ ሰንጠረዡ የሽያጭ ዋጋንና
የግንባታ ወጪን አንድ ላይ ይቀይራል — አረንጓዴ ሕዋሶች ከዝቅተኛ ተፈላጊ ተመላሹ በላይ ናቸው። <b>ቶርኔዶ</b> ሰንጠረዡ የትኛው ነጂ ውጤቱን በብዛት እንደሚያናውጥ በቅደም ተከተል ያሳያል።</p>
{fig("B_52_Sensitivity_1.png", "ምስል 9.1 — የስጋት ትንተና ገጽ (52)፡ የማስተካከያ ሕዋሶች፣ ውጤቶችና የመጀመሪያዎቹ ሰንጠረዦች።",
     [(1, 49, 27), (2, 49, 51), (3, 37, 74)],
     ["የማስተካከያ ሕዋሶች (ሰማያዊ)", "በቀጥታ የሚቀየሩ ውጤቶች", "ባለአንድ አቅጣጫ የተጽዕኖ ሰንጠረዥ (Data Table)"], width=80, crop=(0, 0.58, 0, 0.66))}
</div><div class="page">
<h3>4. የሚዛን ነጥብ (Break-even)</h3>
<p>«ምን ያህል ቢበላሽ ነው ፕሮጀክቱ የሚወድቀው?» ሞዴሉ ለእያንዳንዱ ሁለት መልስ ይሰጣል፡ በቀመር የተሰላ ፈጣን ግምት (closed-form) እና በ21 ረድፍ የቀጥታ ፍለጋ ሰንጠረዥ
(grid-refined) የተገኘ ትክክለኛ ዋጋ — goal seek ወይም ማክሮ ሳያስፈልግ።</p>
<table><tr><th>የሚዛን ነጥብ</th><th class="num">ውጤት (ምሳሌ ለ)</th><th>ትርጉም</th></tr>
<tr><td>የሽያጭ ዋጋ (NPV = 0)</td><td class="num">{etb(B['BE_PRICE_NPV'])} ብር/m²</td><td>የአሁኑ ዋጋ {etb(B['BLEND_PRICE'])} — ዋጋው እስከዚህ ቢወርድ ፕሮጀክቱ የቅናሽ መጠኑን ብቻ ያስገኛል</td></tr>
<tr><td>የተያዘ መጠን (DSCR = ግዴታ)</td><td class="num">{pct(B['BE_OCC'])}</td><td>የተረጋጋው የኪራይ መያዝ ከዚህ በታች ቢወርድ የባንክ ግዴታ ይጣሳል</td></tr>
<tr><td>ከፍተኛ የወጪ ጭማሪ (IRR &lt; hurdle)</td><td class="num">{pct(B['BE_COST'])}</td><td>ወጪው ከዚህ በላይ ቢጨምር ባለቤት IRR ከ{pct(B['HURDLE'], 0)} በታች ይወርዳል</td></tr>
<tr><td>የወለድ መጠን (DSCR = ግዴታ)</td><td class="num">{pct(B['BE_RATE'])}</td><td>የአሁኑ {pct(B['RATE_EFF'])}</td></tr>
<tr><td>ከፍተኛ መዘግየት (IRR &lt; hurdle)</td><td class="num">{(f"{B['BE_DELAY']:.1f} ወር") if B['BE_DELAY'] >= 0 else "ከፈተናው ክልል በላይ"}</td><td>ግንባታው ከዚህ በላይ ቢዘገይ ውሳኔው ይቀየራል</td></tr></table>
{fig("B_52_Sensitivity_2.png", "ምስል 9.2 — ቶርኔዶና የሚዛን ነጥብ ሰንጠረዦች።",
     [(1, 1.5, 7), (2, 1.5, 67)],
     ["ቶርኔዶ፡ በእያንዳንዱ ነጂ የባለቤት IRR ለውጥ", "የሚዛን ነጥቦች፡ በቀመርና በፍለጋ ሰንጠረዥ"], width=96, crop=(0.46, 0.84, 0, 0.6))}
<div class="box"><b class="t">ባንክ ፊት ከመቅረብዎ በፊት</b>የባንኩን ሁኔታ (Lender case) ይምረጡ፣ ዝቅተኛው DSCR ከገደቡ በላይ መሆኑን ያረጋግጡ፣ ቶርኔዶውን ያትሙ፣ እና «ዋጋው ቢቀንስ የምናደርገውን»
(ለምሳሌ ተጨማሪ ቅድመ-ሽያጭ፣ ረዘም ያለ የብድር ዘመን) አስቀድመው ያቅዱ።</div>
</div>""")
    # ------------------------------------------------------------------ CH10
    S.append(f"""<div class="page">{chapter(10, "ዳሽቦርድ", "The dashboard – one page for decision makers (sheet 01)", "CH10")}
<p>ዳሽቦርዱ ሁሉንም በአንድ A4 ገጽ (አግድም) ያሳያል። ከላይ ወደ ታች ያንብቡት፡</p>
{fig("B_01_Dashboard_1.png", "ምስል 10.1 — ዳሽቦርድ (ምሳሌ ለ፣ መሠረታዊ ሁኔታ)።",
     [(1, 1.5, 11.5), (2, 57, 11.5), (3, 1.5, 17), (4, 50.5, 17), (5, 83.5, 17), (6, 1.5, 29.7), (7, 1.5, 55.6), (8, 67, 55.6), (9, 1.5, 81), (10, 50.5, 81)],
     ["የሁኔታ መስመር፡ ሁኔታ፣ ማረጋገጫዎች፣ ዓይነት", "ውሳኔ፡ VIABLE / MARGINAL / NOT VIABLE", "ጠቅላላ የልማት ወጪ (TDC)",
      "ባለቤት IRR ከhurdle ጋር", "ዝቅተኛ DSCR ከግዴታው ጋር", "የገንዘብ ፍሰትና ድምሩ", "የገቢ ድብልቅ", "ቶርኔዶ", "የገንዘብ ምንጮችና አጠቃቀም", "ማስጠንቀቂያዎች"], crop=(0, 1, 0, 1))}
<h2>ውሳኔው እንዴት ይወሰናል?</h2>
<table><tr><th>ውሳኔ</th><th>ሁኔታ</th></tr>
<tr><td style="color:#2E7D4F;font-weight:700">VIABLE · አዋጭ</td><td>ባለቤት IRR ≥ ዝቅተኛ ተፈላጊ ተመላሽ <b>እና</b> ትርፍ ከወጪ &gt; 0 <b>እና</b> DSCR ≥ ግዴታ (ብድር ካለ)</td></tr>
<tr><td style="color:#B7791F;font-weight:700">MARGINAL · በድንበር ላይ</td><td>ባለቤት IRR ከhurdle እስከ {pct(0.03, 0)} በታች ነው (ክልሉ በግብዓት ይቀየራል) እና GDV ≥ TDC</td></tr>
<tr><td style="color:#B42318;font-weight:700">NOT VIABLE · አዋጭ አይደለም</td><td>ከላይ ያሉት ካልተሟሉ</td></tr></table>
<p>ስድስቱ የቁልፍ ውጤት ሳጥኖች፡ ጠቅላላ የልማት ወጪ፣ GDV፣ ትርፍ ከወጪ፣ ባለቤት IRR፣ ከፍተኛ ብድር፣ ዝቅተኛ DSCR። በዓይነት 0 እነዚህ ሳጥኖች ወደ ቤት መረጃ ይቀየራሉ፡
ጠቅላላ ወጪ፣ ዋጋ በካሬ፣ ወርሃዊ ክፍያ፣ የገንዘብ ክፍተት፣ የወጣ ብድርና የመክፈል አቅም።</p>
</div>""")
    # ------------------------------------------------------------------ CH11
    checks = [("C01", "የሀብትና ዕዳ መግለጫ በየዓመቱ ይመጣጠናል", "በጥቁር ሕዋስ ላይ ተጽፏል — Ctrl+Z ወይም አዲስ ቅጂ"),
              ("C02", "የገንዘብ ፍሰት መጨረሻ = የሀብትና ዕዳ ጥሬ ገንዘብ", "እንደ C01"),
              ("C03", "የገንዘብ ምንጮች = አጠቃቀሞች (ወርሃዊ)", "እንደ C01"),
              ("C04", "TDC = ወርሃዊ የወጪ ክምችት", "እንደ C01"),
              ("C05", "ወጪው በIAS 2/40/16 ሙሉ በሙሉ ተከፋፍሏል", "የተጣራ ስፋትና የሚሸጠውን ድርሻ ይፈትሹ"),
              ("C06", "ብድሩ በዘመኑ ውስጥ ተከፍሎ ያልቃል", "የብድር ዘመኑን ያራዝሙ ወይም እፎይታውን ያሳጥሩ"),
              ("C07", "የተከፈለ ዋና = የወጣ ብድር + የተደመረ ወለድ", "እንደ C01"),
              ("C08", "ክምችቱ በመጨረሻ ሽያጭ ያልቃል", "የሽያጭ ፍጥነት (ብሎክ 6) ይፈትሹ"),
              ("C09", "ከደንበኞች ቅድሚያ ወደ ዜሮ ይመለሳል", "የክፍያ ዕቅዱን ይፈትሹ"),
              ("C10", "መሬት የእርጅና ቅናሽ አይደረግበትም", "እንደ C01"),
              ("C11", "የፎቅ ስፋት = የአገልግሎት ስፋት", "ለእያንዳንዱ ፎቅ አገልግሎት ይምረጡ"),
              ("C12", "የግብር ኪሳራ አሉታዊ አይሆንም", "እንደ C01"),
              ("C13", "የምርጫ ሕዋሶች ትክክል ናቸው", "በመጀመሪያ ገጽ ከዝርዝሩ ይምረጡ"),
              ("C14", "በየትኛውም ገጽ የቀመር ስህተት የለም (#REF፣ #DIV/0…)", "የስህተቱን ሕዋስ ያግኙ፣ ግብዓቱን ያስተካክሉ"),
              ("C15", "የክፍያ ዕቅድና ቅድመ-ሽያጭ ≤ 100%", "መቶኛዎቹን ይቀንሱ"),
              ("C16", "የቤት ግምት = የንግድ ሞዴሉ ወጪ", "F9 ይጫኑ"),
              ("C17", "የተ.እ.ታ ሂሳብ ይቀጥላል", "እንደ C01"),
              ("C18", "DSCR ≥ ግዴታ (ማስጠንቀቂያ)", "ብድሩን ይቀንሱ፣ ዘመኑን ያራዝሙ"),
              ("C19", "ጥሬ ገንዘብ ከዝቅተኛው በታች አይወርድም (ማስጠንቀቂያ)", "ካፒታል ይጨምሩ ወይም የትርፍ ክፍፍል ይቀንሱ"),
              ("C20", "የገንዘብ ክፍተት = 0 (ማስጠንቀቂያ)", "ተጨማሪ ቁጠባ፣ ብድር ወይም ቅድመ-ሽያጭ"),
              ("C21", "FAR በተፈቀደው ውስጥ ነው (ማስጠንቀቂያ)", "ፎቆችን ይቀንሱ ወይም ፈቃዱን ያረጋግጡ"),
              ("C22", "የሕንፃ ዓይነት ከፎቆቹ ጋር ይስማማል (ማስጠንቀቂያ)", "ብሎክ 2 ወይም ብሎክ 3ን ያስተካክሉ"),
              ("C23", "የንብረት ዋጋ መቀነስ ምልክት የለም (ማስጠንቀቂያ)", "ኪራዩንና የመውጫ ካፕ ሬቱን ይፈትሹ")]
    crow = "".join(f"<tr><td><b>{a}</b></td><td>{b_}</td><td>{c}</td></tr>" for a, b_, c in checks)
    S.append(f"""<div class="page">{chapter(11, "ቼኮች", "Checks – what each one means and how to fix it (sheet 60)", "CH11")}
<p>ሞዴሉ 23 ራስ-ሰር ማረጋገጫዎች አሉት። <b>Error</b> ዓይነቶቹ ሁልጊዜ PASS መሆን አለባቸው — FAIL ካለ ውጤቱን አይጠቀሙ። <b>Warning</b> ዓይነቶቹ የሞዴል ስህተት ሳይሆኑ
የፕሮጀክቱ ስጋት ምልክቶች ናቸው (ለምሳሌ የገንዘብ ክፍተት)። ድምሩ በመጀመሪያ ገጽና በዳሽቦርዱ ላይ ይታያል (ምሳሌ ለ፡ {B['CHK_PASS']} / {B['CHK_TOTAL']})።</p>
<table><tr><th style="width:12mm">ቁጥር</th><th>ምን ይፈትሻል</th><th style="width:62mm">ቢወድቅ ምን ያደርጋሉ</th></tr>{crow}</table>
</div><div class="page">
{fig("B_60_Checks_1.png", "ምስል 11.1 — የማረጋገጫ ገጽ (60 Checks)።",
     [(1, 1.9, 26), (2, 52.5, 26), (3, 85, 26), (4, 1.9, 82)],
     ["የማረጋገጫው ስምና ቁጥር", "ዓይነት፡ Error ወይም Warning", "PASS / WARN / FAIL", "ድምርና አጠቃላይ ሁኔታ"], width=88, crop=(0, 1, 0, 0.8))}
<div class="box green"><b class="t">የተለመዱ ስህተቶችና መፍትሄዎቻቸው</b>
• <b>ሁሉም ቁጥሮች ዜሮ ናቸው</b> — ጥያቄ 2 «My project» ሆኖ ዓምድ E ባዶ ነው፤ ወይም ቁጥሮቹን በF/G ዓምድ ጽፈዋል።<br>
• <b>C06 ወደቀ</b> — የብድር ዘመኑ ለክፍያው አጭር ነው፤ የብድር ዘመኑን ይጨምሩ።<br>
• <b>C20 ማስጠንቀቂያ</b> — ቁጠባ + ቤተሰብ + ብድር ከወጪው ያንሳል፤ የቤት ግምት ገጹ ክፍተቱን በብር ያሳያል።<br>
• <b>C22 ማስጠንቀቂያ</b> — «G+1» መርጠው ሦስት ፎቅ አስገብተዋል።<br>
• <b>የስጋት ሰንጠረዦች አይቀየሩም</b> — F9 ይጫኑ (ስሌት «Automatic except tables» ከሆነ)።</div>
</div>""")
    # ------------------------------------------------------------------ CH12A
    S.append(f"""<div class="page">{chapter("12ሀ", "ምሳሌ ሀ፡ የግል ቤት ወጪ ግምት", "Sample A – G+1 family house, end to end in about 10 minutes", "CH12A")}
<p><b>ሁኔታው፡</b> አንድ ቤተሰብ በአዲስ አበባ ቦሌ ክፍለ ከተማ በ300 m² ቦታ ላይ G+1 መኖሪያ ቤት መሥራት ይፈልጋል (ምድር ወለል 130 m²፣ አንደኛ ፎቅ 120 m²፣ ጠቅላላ 250 m²)።
የማጠናቀቂያ ደረጃ «Standard»፣ ግንባታው 14 ወር። ቦታው የራሳቸው ስለሆነ የመሬት ወጪ የለም። ገንዘቡ፡ ቁጠባ {etb(A['SAVINGS'])}፣ የቤተሰብና ዳያስፖራ መዋጮ {etb(A['FAMILY'])}፣
የቤት ብድር እስከ {etb(A['HOME_LOAN'])} ብር ({pct(A['RATE_EFF'])}፣ {A['TENOR_Y']} ዓመት)። የቤተሰቡ ወርሃዊ ገቢ {etb(A['HH_INCOME'])} ብር ነው።</p>
<h3>ደረጃዎች</h3>
<p><b>1.</b> በመጀመሪያ ገጽ ጥያቄ 1 = «0 – Personal residence»፣ ጥያቄ 2 = «2 – SAMPLE A»። <b>2.</b> በገጽ 10 ብሎክ 2 ያሉትን (የሕንፃ ዓይነት G+1፣ Standard፣ ገቢ፣ ቁጠባ፣ ብድር) ይመልከቱ።
<b>3.</b> ብሎክ 3 የፎቅ ስፋቶችን፣ ብሎክ 4 አጥር (70 ሜ.)፣ የፍሳሽ ጉድጓድ፣ ሁለት ታንከርና 80 m² ንጣፍ ያሳያሉ። <b>4.</b> «Home Estimate» ገጹን ይክፈቱ።</p>
<div class="kpi"><div>የግንባታ ወጪ<b>{etb(A['CONSTR_COST'])}</b>ብር</div><div>ዋጋ በካሬ<b>{etb(A['CC_M2'])}</b>ብር/m²</div>
<div>ወርሃዊ ክፍያ<b>{etb(A['LN_INSTAL_M'])}</b>{pct(A['AFFORD_PCT'], 0)} የገቢ</div><div>የገንዘብ ክፍተት<b>{etb(A['FUND_GAP'])}</b>ብር</div></div>
{fig("A_02_Home_1.png", "ምስል 12.1 — የቤት ግምት ገጽ (ምሳሌ ሀ)። በአንድ A4 ገጽ ይታተማል።",
     [(1, 3, 15.6), (2, 57, 15.6), (3, 3, 24.3), (4, 57, 24.3), (5, 3, 46.7), (6, 57, 46.7), (7, 3, 62.6), (8, 57, 62.6), (9, 3, 83.4)],
     ["የግንባታ ወጪና ከዝቅተኛ እስከ ከፍተኛ ክልል", "ወርሃዊ ክፍያና የገንዘብ ክፍተት", "የወጪ ክፍፍል በክፍል", "የወጪ ክፍፍል ሰንጠረዥ",
      "ለኮንትራክተሩ የክፍያ ደረጃዎችና ወራቸው", "ወርሃዊ የሚያስፈልግ ገንዘብ (S-ኩርባ)", "ገንዘብ፣ ብድርና የመክፈል አቅም", "የቤት ብድር ሰንጠረዥ በዓመት",
      "ከፊሉን የማከራየት አማራጭ"], width=72, crop=(0, 1, 0, 1))}
</div><div class="page">
<h3>ውጤቱን ማንበብ</h3>
<p>የግንባታ ወጪው (መሬትና የብድር ወጪ ሳይጨምር) <b>{etb(A['CONSTR_COST'])} ብር</b> ነው — ከ{m(A['CONSTR_COST'] * (1 + A['EST_LOW']))} እስከ {m(A['CONSTR_COST'] * (1 + A['EST_HIGH']))} ሚሊዮን ብር ክልል
(−10% / +15%)። ይህ በካሬ <b>{etb(A['CC_M2'])} ብር</b> ማለት ነው። የበጀት ቀን ቀጥተኛ ወጪ {etb(A['HARD_BASE'])} ብር ሲሆን በ14 ወራት የዋጋ ጭማሪና 10% መጠባበቂያ
{etb(A['ESC_AMT'] + A['CONT_TOT'])} ብር ይጨምራሉ። የብድር ወጪን (የማመቻቻ ክፍያና IDC {etb(A['IDC_TOT'])}) ጨምሮ ጠቅላላ የፕሮጀክት ወጪ <b>{etb(A['HOME_TOTAL'])} ብር</b> ነው።</p>
<p><b>የክፍያ ደረጃዎች፡</b> ገጹ ለኮንትራክተሩ የሚከፈለውን በስድስት ደረጃዎች (ቅድመ-ዝግጅት፣ መሠረት፣ ፍሬም፣ ግድግዳና ጣሪያ፣ ኤሌክትሪክና ቧንቧ፣ ማጠናቀቂያ) እና ግምታዊ ወራቸውን ያሳያል — ውል ሲደራደሩ ይጠቅማል።
5% የማስያዣ ገንዘብ ({etb(A['RET_TOT'])} ብር) ከግንባታ በኋላ ይለቀቃል።</p>
<p><b>ብድሩ፡</b> ቁጠባውና የቤተሰብ መዋጮው መጀመሪያ ስለሚወጡ ብድሩ የሚወሰደው በግንባታው መጨረሻ ወራት ብቻ ነው — ስለዚህ IDC አነስተኛ ነው። የወጣው ብድር {etb(A['DEBT_DRAWN'])} ብር፣
ወርሃዊ ክፍያው <b>{etb(A['LN_INSTAL_M'])} ብር</b> ሲሆን ከገቢው {pct(A['AFFORD_PCT'])} ነው — ከ35% ገደብ በታች ስለሆነ «Affordable · ይቻላል» ይላል። ብድሩ {A['LN_FINAL_DATE']:%B %Y} ተከፍሎ ያልቃል።</p>
<h2>«የPremium ማጠናቀቂያ ብመርጥስ?»</h2>
<p>በብሎክ 2 «Finish level»ን ወደ <b>Premium</b> ይቀይሩ። የማጠናቀቂያ፣ የውስጥ ሥራና የኤሌክትሪክ ዋጋዎች ይጨምራሉ፡</p>
<table><tr><th></th><th class="num">Standard</th><th class="num">Premium</th><th class="num">ልዩነት</th></tr>
<tr><td>የግንባታ ወጪ</td><td class="num">{etb(A['CONSTR_COST'])}</td><td class="num">{etb(Ap['CONSTR_COST'])}</td><td class="num">+{pct(Ap['CONSTR_COST'] / A['CONSTR_COST'] - 1)}</td></tr>
<tr><td>ዋጋ በካሬ</td><td class="num">{etb(A['CC_M2'])}</td><td class="num">{etb(Ap['CC_M2'])}</td><td class="num">+{etb(Ap['CC_M2'] - A['CC_M2'])}</td></tr>
<tr><td>ወርሃዊ ክፍያ</td><td class="num">{etb(A['LN_INSTAL_M'])}</td><td class="num">{etb(Ap['LN_INSTAL_M'])}</td><td class="num">+{etb(Ap['LN_INSTAL_M'] - A['LN_INSTAL_M'])}</td></tr>
<tr><td>የገንዘብ ክፍተት</td><td class="num">{etb(A['FUND_GAP'])}</td><td class="num">{etb(Ap['FUND_GAP'])}</td><td class="num">ማስጠንቀቂያ C20</td></tr></table>
<p>የቤት ብድሩ ጣሪያ {etb(A['HOME_LOAN'])} ብር ስለሆነ ተጨማሪው ወጪ <b>{etb(Ap['FUND_GAP'])} ብር የገንዘብ ክፍተት</b> ይፈጥራል — የዳሽቦርዱና የቤት ግምቱ ገጽ በቀይ ያሳያሉ።
መፍትሄ፡ ቁጠባ መጨመር፣ ከፍ ያለ ብድር መጠየቅ ወይም Premium የሚሆነውን ክፍል (ለምሳሌ ኩሽናና መታጠቢያ ብቻ) መወሰን።</p>
{fig("Ap_02_Home_1.png", "ምስል 12.2 — ተመሳሳይ ቤት በPremium ማጠናቀቂያ፡ የወጪ ጭማሪና የገንዘብ ክፍተት።", [(1, 77, 62)], ["የገንዘብ ክፍተት ታየ"], width=90, crop=(0, 0.235, 0, 1))}
</div>""")
    # ------------------------------------------------------------------ CH12B
    S.append(f"""<div class="page">{chapter("12ለ", "ምሳሌ ለ፡ የንግድ ፕሮጀክት ትርፋማነት", "Sample B – B+G+8 mixed-use building, end to end, plus three what-if exercises", "CH12B")}
<p><b>ሁኔታው፡</b> በአዲስ አበባ ቂርቆስ ክፍለ ከተማ በ1,000 m² የሊዝ ቦታ ላይ B+G+8 ቅይጥ ሕንፃ (ጠቅላላ {etb(B['GFA'])} m²)። ምድር ቤት — መኪና ማቆሚያ፣ ምድርና አንደኛ ፎቅ — ሱቆች
(ለኪራይ ይቀራሉ፣ {etb(B['LET_AREA'])} m² ተከራይ)፣ ከ2ኛ እስከ 8ኛ ፎቅ — አፓርትመንቶች (ለሽያጭ፣ {etb(B['SALE_AREA'])} m²)። ዓይነት 3 (ቅይጥ)፣ {B['CONSTR_EFF']} ወር ግንባታ፣ የባንክ ብድር
ከ12 ወር እፎይታ ጋር። መሬቱ 150 ሚሊዮን ብር ሊዝ (10% ቅድመ ክፍያ፣ 30 ዓመት)።</p>
<h3>ደረጃዎች</h3>
<p><b>1.</b> ጥያቄ 1 = «3 – Mixed»፣ ጥያቄ 2 = «3 – SAMPLE B»። <b>2.</b> ማረጋገጫዎቹ {B['CHK_PASS']}/{B['CHK_TOTAL']} መሆናቸውን ይመልከቱ። <b>3.</b> ዳሽቦርዱን ያንብቡ፡</p>
<div class="kpi"><div>TDC<b>{m(B['TDC'])} ሚ.</b>ብር</div><div>GDV<b>{m(B['GDV'])} ሚ.</b>ብር</div><div>ትርፍ ከወጪ<b>{pct(B['POC'])}</b></div>
<div>ባለቤት IRR<b>{pct(B['EIRR'])}</b>hurdle {pct(B['HURDLE'], 0)}</div><div>ዝቅተኛ DSCR<b>{B['MIN_DSCR']:.2f}x</b>ግዴታ {B['DSCR_COV']:.2f}x</div></div>
<p><b>4. ገንዘቡ ከየት መጣ?</b> ካፒታል {m(B['EQ_DRAWN'])}፣ ቅድመ-ሽያጭ {m(B['PS_USED'])}፣ ብድር {m(B['DEBT_DRAWN'])} ሚሊዮን ብር። ከፍተኛው ብድር (IDC ጨምሮ) {m(B['PEAK_DEBT'])} ሚሊዮን ነው።
በርክክብ የሚሰበሰበው {m(B['HAND_TOT'])} ሚሊዮን ብርና ከግንባታ በኋላ የሚሸጡ አፓርትመንቶች 40% በሽያጭ ገንዘብ ብድሩን ቀድሞ ለመክፈል ይውላሉ።</p>
<p><b>5. ትርፉ ከየት ይመጣል?</b> የአፓርትመንቶች ሽያጭ {m(B['SALES_NOM'])} ሚሊዮን ብር ገቢ ያስገኛል፤ ሱቆቹ በተረጋጋ ዓመት {m(B['NOI_STAB'])} ሚሊዮን ብር የተጣራ ኪራይ (NOI) ይሰጣሉ፣
ዋጋቸውም በ9% ካፕ ሬት {m(B['RET_VALUE'])} ሚሊዮን ብር ይገመታል። ከወጪ አንጻር ዓመታዊ ገቢ {pct(B['YOC'])}። ውሳኔ፡ <b>{B['VERDICT']}</b>።</p>
<h2>ሦስት «ምን ቢሆን» መልመጃዎች</h2>
<div class="box blue"><b class="t">መልመጃ 1 — የሽያጭ ዋጋ በ10% ቢቀንስ?</b>በገጽ 52 «Sale price» የማስተካከያ ሕዋሱን 90% ያድርጉ (ወይም ባለአንድ አቅጣጫ ሰንጠረዡን 90% ረድፍ ያንብቡ)።<br>
<b>መልስ፡</b> ባለቤት IRR ከ{pct(B['EIRR'])} ወደ {pct(B['EIRR_P90'])} ይወርዳል። የNPV ሚዛን ነጥብ ዋጋ {etb(B['BE_PRICE_NPV'])} ብር/m² ስለሆነ ፕሮጀክቱ አሁንም ከቅናሽ መጠኑ በላይ ያስገኛል።</div>
<div class="box"><b class="t">መልመጃ 2 — አስቸጋሪው ሁኔታ (Downside)</b>በመጀመሪያ ገጽ ጥያቄ 3ን «3 – Downside» ያድርጉ፡ ዋጋ −10%፣ ወጪ +10%፣ 6 ወር መዘግየት፣ ወለድ +2 ነጥብ፣ ቅድመ-ሽያጭ −25%…<br>
<b>መልስ፡</b> ባለቤት IRR {pct(sc[3][0])}፣ ትርፍ ከወጪ {pct(sc[3][1])}፣ ዝቅተኛ DSCR {sc[3][2]:.2f}x — ውሳኔው «NOT VIABLE» ይሆናል፣ C18 ማስጠንቀቂያ ያሳያል። ይህ ፕሮጀክቱ
ለወጪ ጭማሪና ለመዘግየት ስሜታዊ መሆኑን ያሳያል — ቋሚ ዋጋ ያለው የግንባታ ውል ወይም ተጨማሪ ቅድመ-ሽያጭ ያስቡ።</div>
<div class="box green"><b class="t">መልመጃ 3 — ምሳሌ ሀን «ምድር ወለሉን ማከራየት» (የማሻሻያ መንገድ)</b>ጥያቄ 2 = ምሳሌ ሀ፣ ዓይነት 0። በብሎክ 2 «I may rent out part of the house» = Yes፣ 60 m²፣ በወር 25,000 ብር (የዛሬ ዋጋ)።<br>
<b>መልስ፡</b> የመጀመሪያ ዓመት ኪራይ {etb(Ar['RENT1'])} ብር፣ በጠቅላላ ወጪ ላይ ዓመታዊ ገቢ {pct(Ar['YIELD'])}፣ ቀላል የመመለሻ ጊዜ {Ar['PAYBACK']:.1f} ዓመት፣ ኪራዩ የወርሃዊ ክፍያውን
{pct(Ar['COVER'], 0)} ይሸፍናል። ቀጣዩ እርምጃ፡ ዓይነቱን ወደ 3 ቀይረው ሙሉ ትርፋማነትን በተመሳሳይ መረጃ ማየት ይቻላል።</div>
{fig("Ar_02_Home_1.png", "ምስል 12.3 — ምሳሌ ሀ ከፊሉን በማከራየት፡ የኪራይ ገቢ፣ ዓመታዊ ገቢና የመመለሻ ጊዜ።", [(1, 3, 12)], ["የማከራየት አማራጭ ውጤቶች"], width=70, crop=(0.8, 0.94, 0, 0.62))}
</div>""")
    # ------------------------------------------------------------------ CH13
    faqs = [
        ("ሞዴሉን ለመጠቀም የፋይናንስ ባለሙያ መሆን አለብኝ?", "አይ። ሰማያዊ ሕዋሶችን ብቻ ይሞላሉ፤ ቀሪውን ሞዴሉ ያሰላል። ለንግድ ፕሮጀክት ግን ውጤቱን ከባንክዎ ወይም ከአማካሪ ጋር መወያየት ይመከራል።"),
        ("የወጪ ዋጋዎቹ ትክክለኛ ናቸው?", "የመረጃ ቋቱ ዋጋዎች ናሙና ናቸው (Sample value — not market advice)። ከኮንትራክተር፣ ከBOQ ወይም ከራፋ የግንባታ ዋጋ ኢንዴክስ በተገኘ ዋጋ ይተኩዋቸው።"),
        ("ፎቆቼ ከ16 በላይ ቢሆኑስ?", "ሞዴሉ 2 ምድር ቤትና እስከ G+13 ይቀበላል። ለበለጠ፣ ተመሳሳይ ፎቆችን በአንድ ረድፍ ደምረው ያስገቡ (አገልግሎታቸው አንድ እስከሆነ)።"),
        ("ግንባታው ከ48 ወር በላይ ቢሆንስ?", "ሞዴሉ እስከ 48 ወር ያስተናግዳል። ረጅም ፕሮጀክቶችን በደረጃ (phase) ከፍለው ለእያንዳንዱ ፋይል ይጠቀሙ።"),
        ("የግብር መጣኔዎቹ ወቅታዊ ናቸው?", "በገጽ 12 ምንጫቸው ተጽፏል፤ ግን ሁሉም «Verify before use» ተብለዋል። ከመጠቀምዎ በፊት ከገቢዎች ሚኒስቴር ወይም ከግብር አማካሪ ያረጋግጡ።"),
        ("ለምን የቤት ብድር ክፍያዬ ከባንኩ ትንሽ ይለያል?", "ባንኮች የተለያዩ የክፍያ ቀኖች፣ የአገልግሎት ክፍያዎችና መድን ይጨምራሉ። ሞዴሉ ንጹህ PMT ይጠቀማል — ልዩነቱ ትንሽ መሆን አለበት።"),
        ("ውጤቱ «NOT VIABLE» ቢል ምን ላድርግ?", "ቶርኔዶውን ይመልከቱ — ብዙ የሚያናውጠውን ነጂ ይለዩ። ብዙ ጊዜ ዋጋ፣ ወጪ ወይም የቅድመ-ሽያጭ መጠን ናቸው። የሚዛን ነጥቦቹ ምን ያህል መቀየር እንዳለበት ያሳያሉ።"),
        ("Data Tables ለምን አይቀየሩም?", "Excel «Automatic except for data tables» ላይ ከሆነ F9 ይጫኑ። LibreOffice ሰንጠረዦቹን በትክክል አያሰላም — በExcel ይጠቀሙ።"),
        ("ማክሮ ወይም ኢንተርኔት ያስፈልጋል?", "አያስፈልግም። ሁሉም በቀመር ነው፤ ፋይሉ ከመስመር ውጪ ይሠራል።"),
        ("የራሴን ፕሮጀክት ከምሳሌዎቹ እንዴት እለያለሁ?", "ጥያቄ 2ን «My project» ያድርጉና በዓምድ E ብቻ ይጻፉ። ባዶ ፋይሉን (Rapha_Model_Blank.xlsx) መጀመሪያ መጠቀምም ይቻላል።"),
    ]
    fq = "".join(f"<h3>{i + 1}. {q}</h3><p>{a}</p>" for i, (q, a) in enumerate(faqs))
    S.append(f"""<div class="page">{chapter(13, "ተደጋጋሚ ጥያቄዎችና መዝገበ ቃላት", "FAQ and English–Amharic glossary", "CH13")}
<h2>ተደጋጋሚ ጥያቄዎች</h2>{fq}</div>""")
    half = (len(GLOSSARY) + 1) // 2
    def gtable(items):
        return ('<table class="gl"><tr><th>English</th><th>አማርኛ</th><th>ትርጉም</th></tr>' +
                "".join(f"<tr><td><b>{e}</b></td><td style='color:#0F2A44;font-weight:600'>{a}</td><td>{d}</td></tr>" for e, a, d in items) + "</table>")
    S.append(f"""<div class="page"><h2>መዝገበ ቃላት (እንግሊዝኛ – አማርኛ) — {len(GLOSSARY)} ቃላት</h2>
<p class="small">ተመሳሳይ ቃላት በሞዴሉ ገጽ 61 ይገኛሉ።</p>{gtable(GLOSSARY[:half])}</div>
<div class="page">{gtable(GLOSSARY[half:])}</div>""")
    # ------------------------------------------------------------------ CH14
    S.append(f"""<div class="page">{chapter(14, "የኃላፊነት ማስተባበያ", "Disclaimer", "CH14")}
<p>የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴልና ይህ መመሪያ <b>ለዕቅድና ለትምህርት ብቻ</b> የተዘጋጁ ናቸው። የኢንቨስትመንት፣ የግብር፣ የሂሳብ አያያዝ ወይም የሕግ ምክር አይደሉም።</p>
<p>• በምሳሌዎቹ ውስጥ ያሉ ቁጥሮች በሙሉ <b>ናሙና</b> ናቸው (Sample value — not market advice)። የገበያ ዋጋ፣ የግንባታ ወጪ፣ ኪራይና የወለድ መጠን ይለዋወጣሉ።<br>
• የግብር መጣኔዎችና ሕጎች (የገቢ ግብር፣ ተ.እ.ታ፣ የትርፍ ድርሻ ግብር፣ የንብረት ግብር፣ የሊዝ ሁኔታዎች) በገጽ 12 ተመዝግበዋል፤ <b>ከመጠቀምዎ በፊት ወቅታዊ መሆናቸውን ማረጋገጥ ግዴታዎ ነው</b>።<br>
• የIFRS አተገባበሩ ቀለል ተደርጎ የቀረበ ነው (ለምሳሌ ወጪ በተጣራ ስፋት ይከፋፈላል፣ የንብረት ዋጋ መቀነስ ፈተና ብቻ ይደረጋል)። ለሕጋዊ የሂሳብ መግለጫ ፈቃድ ያለው ኦዲተር ያማክሩ።<br>
• የሞዴሉ ውጤቶች ጥራት በገባው መረጃ ጥራት ላይ ይመሠረታል። ውሳኔ ከመወሰንዎ በፊት ከባለሙያ የተዘጋጀ BOQ፣ የገበያ ጥናትና የባንክ ይፋዊ የብድር ውል ያግኙ።<br>
• ሞዴሉን ማሻሻል፣ ቀመሮችን መቀየር ወይም ማረጋገጫዎቹን ማጥፋት ውጤቱን ሊያበላሽ ይችላል። ሁልጊዜ ሁሉም ማረጋገጫዎች PASS መሆናቸውን ያረጋግጡ።</p>
<div class="box red"><b class="t">ማሳሰቢያ</b>ደራሲዎቹና አከፋፋዮቹ በዚህ ሞዴል ላይ ተመሥርቶ ለሚወሰድ ማንኛውም ውሳኔ ወይም ለሚደርስ ኪሳራ ኃላፊነት አይወስዱም።</div>
<p class="en" style="margin-top:8mm">English summary: This model and guide are planning tools only and do not constitute investment, tax, accounting or legal advice.
All sample figures are illustrative. Tax rates must be verified against current Ethiopian proclamations before use.</p>
<p class="small" style="margin-top:20mm;text-align:center">የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል · ስሪት 1.0 · Rapha Development Profitability Model</p>
</div>""")
    return f'<!doctype html><html lang="am"><head><meta charset="utf-8"><style>{CSS}</style></head><body>{"".join(S)}</body></html>'


def toc_from_pdf(pdf):
    pages = {}
    with pdfplumber.open(pdf) as p:
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            for mk in re.findall(r"CH\d\d[AB]?", t):
                pages.setdefault(mk, i + 1)
    return pages


if __name__ == "__main__":
    A, Ap, Ar, B = load_facts()
    html = os.path.join(HERE, "tutorial.html")
    pdf = os.path.join(OUT, "Rapha_Model_Tutorial_AM.pdf")
    open(html, "w").write(build(A, Ap, Ar, B))
    subprocess.run(["node", os.path.join(HERE, "topdf.js"), html, pdf], check=True)
    toc = toc_from_pdf(pdf)
    open(html, "w").write(build(A, Ap, Ar, B, toc))
    subprocess.run(["node", os.path.join(HERE, "topdf.js"), html, pdf], check=True)
    print("TOC", toc)
