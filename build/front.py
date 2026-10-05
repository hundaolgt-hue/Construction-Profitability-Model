"""Front sheets: Lists, Inputs, Scenarios, TaxLaw, Start, Home, Dashboard, Glossary."""
import datetime as dt
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from core import (axis_font, Sheet, pc, CL, NF, ANN_N, MON_C, MON_L, font, fill, F_BODY, F_BOLD, F_UNIT, F_NOTE, F_INPUT, FILL_INPUT, BOX,
                  NAVY, AMBER, BLUE, TEAL, VIOLET, ROSE, CANVAS, CARD, BORDER, INK, SEC, GOOD_FG, GOOD_BG, WARN_FG, WARN_BG,
                  CRIT_FG, CRIT_BG, FILL_KEY, FILL_NAVY, FILL_AMBER, FILL_CANVAS, FILL_CARD, AMFONT, LINK_FG, BOTTOM_AMBER,
                  INPUT_FG, INPUT_BG, PERIOD_LABELS)
from inputs_def import INPUTS, LISTS, RATE_LIB, EL_KEYS, ELEMENTS, ELEMENTS_AM, SAMPLE_NOTE
from glossary import GLOSSARY

TAB_SETUP = "8A94A3"
TAB_REP = "0F2A44"
TAB_IN = "D98E2B"


def thin_box(ws, rng, color=BORDER):
    s = Side(style="thin", color=color)
    for row in ws[rng]:
        for c in row:
            c.border = Border(left=s, right=s, top=s, bottom=s)


def card(ws, r1, c1, r2, c2, bg=CARD, border=BORDER):
    s = Side(style="thin", color=border)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = fill(bg)
            cell.border = Border(left=s if c == c1 else None, right=s if c == c2 else None,
                                 top=s if r == r1 else None, bottom=s if r == r2 else None)


# ---------------------------------------------------------------------- LISTS
def build_lists(b):
    ws = b.wb.create_sheet("Lists")
    ws.sheet_state = "hidden"
    for j, (nm, vals) in enumerate(LISTS.items()):
        col = j + 1
        ws.cell(1, col, nm).font = F_BOLD
        for k, v in enumerate(vals):
            ws.cell(2 + k, col, v)
        b.name(nm, f"Lists!${CL(col)}$2:${CL(col)}${1 + len(vals)}")
        ws.column_dimensions[CL(col)].width = 30
    return ws


# ---------------------------------------------------------------------- INPUTS
def build_inputs(b, blank=False):
    sh = Sheet(b, "10_Inputs", "10 · Inputs – all assumptions", "ግብዓቶች – ሁሉም ግምቶች", TAB_IN, kind="free", n=0, width_label=52)
    ws = sh.ws
    b.INPUT_ROWS = {}
    b.DV = []
    r = 6
    heads = [("B", "Input"), ("C", "ግብዓት"), ("D", "Unit"), ("E", "MY PROJECT\n(your values)"),
             ("F", "SAMPLE A\nG+1 house" if not blank else "SAMPLE A\n(cleared)"),
             ("G", "SAMPLE B\nB+G+8 mixed" if not blank else "SAMPLE B\n(cleared)"),
             ("H", "ACTIVE\n(used by model)"), ("I", "Range name"), ("J", "Notes / source")]
    for col, h in heads:
        c = ws[f"{col}{r}"]
        c.value = h
        c.font = font(9, True, "FFFFFF")
        c.fill = fill(NAVY if col not in ("E",) else BLUE)
        c.alignment = Alignment(horizontal="center" if col in "EFGH" else "left", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 30
    ws["B7"] = ('Blue cells on amber are inputs. Column H picks the data set chosen on the Start sheet '
                '(1 = My project, 2 = Sample A, 3 = Sample B). Sample columns are examples – "Sample value — not market advice".')
    ws["B7"].font = F_NOTE
    r = 9
    for item in INPUTS:
        if item[0] == "sec":
            c = ws.cell(r, 2, item[1])
            c.font = font(11, True, NAVY)
            ws.cell(r, 3, item[2]).font = font(10, False, SEC, name=AMFONT)
            for col in range(2, 11):
                ws.cell(r, col).border = BOTTOM_AMBER
            ws.row_dimensions[r].height = 20
            r += 1
            continue
        key, en, am, unit, fmt, uv, av, bv, note, dv = item
        if blank:
            av, bv = uv, uv
        ws.cell(r, 2, en).font = F_BODY
        ws.cell(r, 3, am).font = font(9, False, SEC, name=AMFONT)
        ws.cell(r, 4, unit).font = F_UNIT
        for col, v in ((5, uv), (6, av), (7, bv)):
            c = ws.cell(r, col, v)
            c.font = F_INPUT
            c.fill = FILL_INPUT
            c.border = BOX
            c.number_format = NF.get(fmt, fmt)
            c.alignment = Alignment(horizontal="right" if fmt not in ("gen", "txt") else "left")
            if col in (6, 7) and not blank and v not in (None, "", 0) and key not in ("PROJ_NAME", "LOCATION"):
                c.comment = Comment(SAMPLE_NOTE, "Rapha")
        ch = f"CHOOSE(DATA_SET,E{r},F{r},G{r})"
        if isinstance(dv, str) and dv.startswith("L_"):
            f = f"IFERROR(MATCH({ch},{dv},0),1)"
            hf = "int"
        elif dv == "yn":
            f = f'IF({ch}="Yes",1,0)'
            hf = "int"
        else:
            f = ch
            hf = fmt
        c = ws.cell(r, 8, "=" + f)
        c.number_format = NF.get(hf, hf)
        c.font = F_BOLD
        c.fill = fill("EEF1F5")
        c.border = BOX
        ws.cell(r, 9, key).font = font(8, False, SEC)
        ws.cell(r, 10, note).font = F_NOTE
        b.name(key, f"'10_Inputs'!$H${r}")
        b.INPUT_ROWS[key] = r
        # validation
        rng = f"E{r}:G{r}"
        if isinstance(dv, str) and dv.startswith("L_"):
            b.DV.append(("list", dv, rng))
        elif dv == "yn":
            b.DV.append(("list", "L_YN", rng))
        elif dv == "use":
            b.DV.append(("list", "L_USE", rng))
        elif isinstance(dv, tuple):
            b.DV.append((dv[0], dv, rng))
        r += 1
    # rate path range
    b.name("RATE_PATH", f"'10_Inputs'!$H${b.INPUT_ROWS['RATE_O1']}:$H${b.INPUT_ROWS['RATE_O10']}")
    # cost library
    r += 1
    c = ws.cell(r, 2, "12. Cost library – ETB per m² GFA, budget-date prices (link-ready to the Rapha Construction Cost Index)")
    c.font = font(11, True, NAVY)
    ws.cell(r, 3, "12. የወጪ መረጃ ቋት").font = font(10, False, SEC, name=AMFONT)
    for col in range(2, 11):
        ws.cell(r, col).border = BOTTOM_AMBER
    r += 1
    for col, h in (("E", "Basic"), ("F", "Standard"), ("G", "Premium")):
        cc = ws[f"{col}{r}"]
        cc.value = h
        cc.font = font(9, True, "FFFFFF")
        cc.fill = fill(NAVY)
        cc.alignment = Alignment(horizontal="center")
    ws.cell(r, 10, "Applies to all data sets. Replace with Rapha Construction Cost Index rates when available.").font = F_NOTE
    r += 1
    lib0 = r
    for j, k in enumerate(EL_KEYS):
        ws.cell(r, 2, ELEMENTS[j]).font = F_BODY
        ws.cell(r, 3, ELEMENTS_AM[j]).font = font(9, False, SEC, name=AMFONT)
        ws.cell(r, 4, "ETB/m²").font = F_UNIT
        for jj, v in enumerate(RATE_LIB[k]):
            cc = ws.cell(r, 5 + jj, v)
            cc.font = F_INPUT
            cc.fill = FILL_INPUT
            cc.border = BOX
            cc.number_format = NF["n0"]
            if not blank:
                cc.comment = Comment(SAMPLE_NOTE, "Rapha")
            b.DV.append(("dec", ("dec", 0, 1e7), f"{CL(5 + jj)}{r}"))
        ws.cell(r, 8, f"=INDEX(E{r}:G{r},FIN_IX)").number_format = NF["n0"]
        ws.cell(r, 9, "RATE_LIB").font = font(8, False, SEC)
        r += 1
    b.name("RATE_LIB", f"'10_Inputs'!$E${lib0}:$G${r - 1}")
    ws.cell(r, 2, "Total library rate (sum of elements)").font = F_BOLD
    for jj in range(3):
        cc = ws.cell(r, 5 + jj, f"=SUM({CL(5 + jj)}{lib0}:{CL(5 + jj)}{r - 1})")
        cc.number_format = NF["n0"]
        cc.font = F_BOLD
    r += 2
    for key, lab, v, note in (("LOCF_ADDIS", "Location factor – Addis Ababa", 1.0, "Base = 1.00"),
                              ("LOCF_REG", "Location factor – regional town", 0.92, "Lower labour cost, higher transport")):
        ws.cell(r, 2, lab).font = F_BODY
        ws.cell(r, 4, "x").font = F_UNIT
        cc = ws.cell(r, 5, v)
        cc.font = F_INPUT
        cc.fill = FILL_INPUT
        cc.border = BOX
        cc.number_format = NF["n2"]
        if not blank:
            cc.comment = Comment(SAMPLE_NOTE, "Rapha")
        ws.cell(r, 9, key).font = font(8, False, SEC)
        ws.cell(r, 10, note).font = F_NOTE
        b.name(key, f"'10_Inputs'!$E${r}")
        b.DV.append(("dec", ("dec", 0.1, 5), f"E{r}"))
        r += 1
    for col, w in (("C", 30), ("D", 10), ("E", 15), ("F", 15), ("G", 15), ("H", 15), ("I", 16), ("J", 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "E9"
    sh.r = r + 1
    return sh


def add_validations(b):
    for typ, arg, rng in b.DV:
        sheet = b.wb["10_Inputs"]
        if typ == "list":
            dv = DataValidation(type="list", formula1=f"={arg}", allow_blank=True)
        else:
            lo, hi = arg[1], arg[2]
            dv = DataValidation(type="decimal" if typ == "dec" else "whole", operator="between",
                                formula1=str(lo), formula2=str(hi), allow_blank=True)
            dv.error = f"Enter a value between {lo:g} and {hi:g}"
            dv.showErrorMessage = True
        sheet.add_data_validation(dv)
        dv.add(rng)


# ---------------------------------------------------------------------- SCENARIOS
SCEN_DRIVERS = [
    ("SC_PRICE", "Sale price", "የሽያጭ ዋጋ", "%", "pct", 0, 0.10, -0.10, -0.15),
    ("SC_RENT", "Rent", "ኪራይ", "%", "pct", 0, 0.10, -0.10, -0.15),
    ("SC_OCC", "Occupancy (points)", "የተያዘ መጠን", "pts", "pct", 0, 0.03, -0.10, -0.15),
    ("SC_COST", "Construction cost", "የግንባታ ወጪ", "%", "pct", 0, -0.05, 0.10, 0.15),
    ("SC_DELAY", "Construction delay", "የግንባታ መዘግየት", "months", "int", 0, 0, 6, 6),
    ("SC_RATE", "Interest rate (points)", "የወለድ መጠን", "pts", "pct", 0, -0.01, 0.02, 0.03),
    ("SC_PRESALE", "Presales achieved", "ቅድመ-ሽያጭ", "%", "pct", 0, 0.10, -0.25, -0.40),
    ("SC_EXITCAP", "Exit cap rate (points)", "የመውጫ ካፕ ሬት", "pts", "pct", 0, -0.005, 0.01, 0.015),
    ("SC_OPEX", "Operating costs", "የሥራ ማስኬጃ ወጪ", "%", "pct", 0, -0.05, 0.10, 0.15),
    ("SC_ESC", "Cost escalation (points)", "የዋጋ ጭማሪ", "pts", "pct", 0, -0.02, 0.05, 0.05),
    ("SC_RESC", "Rent escalation (points)", "የኪራይ ጭማሪ", "pts", "pct", 0, 0.01, -0.02, -0.03),
    ("SC_PESC", "Sale price escalation (points)", "የሽያጭ ዋጋ ጭማሪ", "pts", "pct", 0, 0.02, -0.03, -0.05),
    ("SC_DEFAULT", "Default / cancellations (points)", "የውል ማፍረስ", "pts", "pct", 0, -0.01, 0.03, 0.05),
    ("SC_DISC", "Discount rates (points)", "የቅናሽ መጠን", "pts", "pct", 0, 0, 0.01, 0.02),
    ("SC_ABS", "Sales absorption speed", "የሽያጭ ፍጥነት", "%", "pct", 0, 0.20, -0.30, -0.50),
]


def build_scenarios(b):
    sh = Sheet(b, "11_Scenarios", "11 · Scenarios – Base / Upside / Downside / Lender", "ሁኔታዎች – መሠረታዊ፣ ብሩህ፣ አስቸጋሪ፣ ባንክ",
               TAB_IN, kind="free", n=0, width_label=36)
    ws = sh.ws
    sh.sec("Scenario adjustments (applied on top of the Inputs). Active = CHOOSE(SCEN, Base, Upside, Downside, Lender)",
           "የሁኔታ ማስተካከያዎች")
    hdr = sh.r
    sh.r += 1 + len(SCEN_DRIVERS) + 1

    def _t(_r):
        for j, h in enumerate(["Driver", "ነጂ", "Unit", "Base", "Upside", "Downside", "Lender", "ACTIVE", "Name"]):
            c = ws.cell(hdr, 2 + j, h)
            c.font = font(9, True, "FFFFFF")
            c.fill = fill(NAVY)
            c.alignment = Alignment(horizontal="center" if j > 2 else "left")
        for k, (key, en, am, unit, fmt, *vals) in enumerate(SCEN_DRIVERS):
            rr = hdr + 1 + k
            ws.cell(rr, 2, en).font = F_BODY
            ws.cell(rr, 3, am).font = font(9, False, SEC, name=AMFONT)
            ws.cell(rr, 4, unit).font = F_UNIT
            for j, v in enumerate(vals):
                c = ws.cell(rr, 5 + j, v)
                c.font = F_INPUT
                c.fill = FILL_INPUT
                c.border = BOX
                c.number_format = NF[fmt]
            c = ws.cell(rr, 9, f"=CHOOSE(SCEN,E{rr},F{rr},G{rr},H{rr})")
            c.number_format = NF[fmt]
            c.font = F_BOLD
            c.fill = fill("EEF1F5")
            ws.cell(rr, 10, key).font = font(8, False, SEC)
            b.name(key, f"'11_Scenarios'!$I${rr}")
        ws.conditional_formatting.add(f"E{hdr + 1}:H{hdr + len(SCEN_DRIVERS)}",
                                      FormulaRule(formula=[f"COLUMN(E{hdr + 1})-4=SCEN"], fill=fill("FFE9C7")))
        for col, w in (("C", 20), ("D", 8), ("E", 11), ("F", 11), ("G", 11), ("H", 11), ("I", 12), ("J", 14)):
            ws.column_dimensions[col].width = w
    sh.raw(_t)
    sh.sec("Effective drivers used by the engine (scenario × flex cell)", "ሞዴሉ የሚጠቀማቸው ውጤታማ ነጂዎች")
    eff = [
        ("SCEN", "Active scenario (1–4)", "IF(SCEN_OVR>0,SCEN_OVR,SCEN_SEL)", "int"),
        ("D_PRICE_F", "Sale price factor", "(1+SC_PRICE)*FLX_PRICE", "n2"),
        ("D_RENT_F", "Rent factor", "(1+SC_RENT)*FLX_RENT", "n2"),
        ("D_COST_F", "Construction cost factor", "(1+SC_COST)*FLX_COST", "n2"),
        ("D_OPEX_F", "Opex factor", "(1+SC_OPEX)*FLX_OPEX", "n2"),
        ("D_OCC_ADD", "Occupancy points added", "SC_OCC", "pct"),
        ("D_OCC_F", "Occupancy factor", "FLX_OCC", "n2"),
        ("D_RATE_ADD", "Interest rate points added", "SC_RATE", "pct2"),
        ("D_RATE_F", "Interest rate factor", "FLX_RATE", "n2"),
        ("RATE_EFF", "Effective base interest rate", "MAX(0,(RATE+D_RATE_ADD)*D_RATE_F)", "pct2"),
        ("CONSTR_EFF", "Effective construction months (incl. delay)", "MAX(1,MIN(48,ROUND(CONSTR_M+SC_DELAY+FLX_DELAY,0)))", "int"),
        ("D_PRESALE_F", "Presale factor", "MAX(0,1+SC_PRESALE)", "n2"),
        ("EXIT_CAP_EFF", "Effective exit cap rate", "MAX(0.001,EXIT_CAP+SC_EXITCAP)", "pct2"),
        ("ESC_EFF", "Effective cost escalation", "ESC_PA+SC_ESC", "pct"),
        ("RESC_EFF", "Effective rent escalation", "RENT_ESC+SC_RESC", "pct"),
        ("PESC_EFF", "Effective sale-price escalation", "PRICE_ESC+SC_PESC", "pct"),
        ("DEFAULT_EFF", "Effective default allowance", "MIN(1,MAX(0,DEFAULT_PCT+SC_DEFAULT))", "pct"),
        ("DISC_EFF", "Effective project discount rate", "DISC+SC_DISC", "pct"),
        ("COE_EFF", "Effective cost of equity", "COE+SC_DISC", "pct"),
        ("ABS_F", "Absorption factor", "MAX(0,1+SC_ABS)", "n2"),
    ]
    for key, lab, f, fmt in eff:
        sh.scalar(key, lab, f, fmt=fmt, unit="", name=key)
    return sh


# ---------------------------------------------------------------------- TAX LAW
TAXLAW = [
    ("Corporate income tax", "CIT", "Federal Income Tax Proclamation No. 979/2016, Art. 19(1) – 30% for bodies"),
    ("Tax depreciation – buildings", "TAXDEP_BLD", "Proc. 979/2016, Art. 25 & Income Tax Regulation No. 410/2017 – buildings 5% straight line"),
    ("Tax depreciation – plant & equipment", "TAXDEP_PLANT", "Proc. 979/2016, Art. 25 – other business assets 20% (pool / diminishing value per spec)"),
    ("Loss carry-forward period", "LCF_YEARS", "Proc. 979/2016, Art. 26 – losses carried forward up to 5 years (limits on number of carry-forwards apply)"),
    ("Loss offset cap", "LCF_CAP", "Model assumption – no statutory cap identified; set below 100% if a cap applies"),
    ("VAT – standard rate", "VAT_R", "Value Added Tax Proclamation No. 1341/2024 – 15% standard rate"),
    ("VAT on sale of buildings (share taxable)", "VAT_SALES_PCT", "Proc. 1341/2024 – exemptions for residential dwellings; confirm treatment of new-build sales"),
    ("VAT on rent (share taxable)", "RENT_VAT_PCT", "Proc. 1341/2024 – residential rent exempt; commercial rent taxable"),
    ("Undistributed profits tax", "UPT_R", "Proc. 979/2016, Art. 55 – undistributed / non-reinvested profits treated as dividend (10%)"),
    ("Dividend withholding tax", "DIV_WHT", "Proc. 979/2016, Art. 55 – 10% on dividends"),
    ("Legal reserve transfer", "LR_PCT", "Commercial Code of Ethiopia, Proc. No. 1243/2021, Art. 434 – 5% of profits"),
    ("Legal reserve cap", "LR_CAP", "Commercial Code, Art. 434 – until reserve reaches 10% of capital"),
    ("Property tax", "PROPTAX_PCT", "City property-tax proclamation (Addis Ababa) – rate and base to be confirmed"),
    ("Land lease down payment", "LEASE_DP", "Urban Lands Lease Holding Proclamation No. 721/2011 and city regulations – minimum down payment"),
    ("Land lease instalment period", "LEASE_N", "Proc. 721/2011 – instalment period per lease contract"),
    ("Interest on unpaid lease", "LEASE_INT", "Lease contract / city regulation – usually linked to the NBE / bank lending rate"),
]


def build_taxlaw(b):
    sh = Sheet(b, "12_TaxLaw", "12 · Tax law register", "የግብር ሕግ መዝገብ", TAB_IN, kind="free", n=0, width_label=38)
    ws = sh.ws
    sh.sec("Rates used by the model and their statutory source – YOU MUST CONFIRM CURRENT RATES BEFORE USE", "የግብር መጣኔዎች")
    hdr = sh.r
    sh.r += 1 + len(TAXLAW) + 3

    def _t(_r):
        for j, h in enumerate(["Item", "Rate used", "Statutory source", "Last verified", "Verify before use", "Model name"]):
            c = ws.cell(hdr, 2 + j, h)
            c.font = font(9, True, "FFFFFF")
            c.fill = fill(NAVY)
        for k, (lab, nm, src) in enumerate(TAXLAW):
            rr = hdr + 1 + k
            ws.cell(rr, 2, lab).font = F_BODY
            c = ws.cell(rr, 3, "=" + nm)
            fmt = "int" if nm == "LEASE_N" else ("pct2" if nm == "PROPTAX_PCT" else "pct")
            c.number_format = NF[fmt]
            c.font = Font(name="Segoe UI", size=10, color=LINK_FG, bold=True)
            c4 = ws.cell(rr, 4, src)
            c4.font = F_BODY
            c4.alignment = Alignment(wrap_text=True, vertical="top")
            c5 = ws.cell(rr, 5, dt.datetime(2026, 10, 5))
            c5.number_format = NF["date"]
            c5.font = F_INPUT
            c5.fill = FILL_INPUT
            c6 = ws.cell(rr, 6, "YES – verify")
            c6.font = font(10, True, CRIT_FG)
            c6.fill = fill(CRIT_BG)
            ws.cell(rr, 7, nm).font = font(8, False, SEC)
            ws.row_dimensions[rr].height = 30
        n = hdr + len(TAXLAW) + 2
        ws.cell(n, 2, "Note: 'Last verified' records a desk review on that date, not confirmation against the official gazette. "
                      "Tax law in Ethiopia changes frequently – confirm every rate with the Ministry of Revenue, the city "
                      "administration or a licensed tax adviser before relying on results.").font = font(10, True, CRIT_FG)
        ws.cell(n + 1, 2, "ማሳሰቢያ፡ ሞዴሉን ከመጠቀምዎ በፊት ሁሉንም የግብር መጣኔዎች ከገቢዎች ሚኒስቴር ወይም ፈቃድ ካለው የግብር አማካሪ ያረጋግጡ።").font = font(10, True, CRIT_FG, name=AMFONT)
        for col, w in (("C", 12), ("D", 70), ("E", 14), ("F", 15), ("G", 16)):
            ws.column_dimensions[col].width = w
    sh.raw(_t)
    return sh


# ---------------------------------------------------------------------- GLOSSARY
def build_glossary(b):
    sh = Sheet(b, "61_Glossary", "61 · Glossary (English – Amharic)", "መዝገበ ቃላት (እንግሊዝኛ – አማርኛ)", TAB_SETUP, kind="free", n=0,
               width_label=34)
    ws = sh.ws
    hdr = 6

    def _g(_r):
        for j, h in enumerate(["English term", "የአማርኛ ቃል", "ትርጉም (Meaning)"]):
            c = ws.cell(hdr, 2 + j, h)
            c.font = font(9, True, "FFFFFF")
            c.fill = fill(NAVY)
        for k, (en, am, d) in enumerate(GLOSSARY):
            rr = hdr + 1 + k
            ws.cell(rr, 2, en).font = F_BOLD
            ws.cell(rr, 3, am).font = font(10, True, NAVY, name=AMFONT)
            c = ws.cell(rr, 4, d)
            c.font = font(10, False, INK, name=AMFONT)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if k % 2:
                for col in (2, 3, 4):
                    ws.cell(rr, col).fill = fill(CANVAS)
        ws.column_dimensions["C"].width = 32
        ws.column_dimensions["D"].width = 80
    sh.raw(_g)
    sh.r = hdr + len(GLOSSARY) + 2
    return sh


# ---------------------------------------------------------------------- START
def build_start(b, blank=False):
    sh = Sheet(b, "00_Start", "Rapha Development Profitability Model", "የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል", TAB_SETUP, kind="free",
               n=0, width_label=4)
    ws = sh.ws

    def _s(_r):
        for col, w in (("A", 2), ("B", 4), ("C", 34), ("D", 30), ("E", 4), ("F", 30), ("G", 22), ("H", 4), ("I", 30), ("J", 20), ("K", 3)):
            ws.column_dimensions[col].width = w
        ws["C5"] = "Does my project make money?  ·  ፕሮጀክቴ ያተርፋል?"
        ws["C5"].font = font(13, True, NAVY)
        ws["C6"] = ("One workbook, two uses: a quick cost estimate for your own house (Mode 0) or full profitability, "
                    "IFRS statements and a loan schedule for a commercial project (Modes 1–3).")
        ws["C6"].font = F_NOTE
        # ---- questions card
        card(ws, 8, 2, 19, 7, bg=CARD)
        ws["C8"] = "SET UP  ·  ማዘጋጀት"
        ws["C8"].font = font(11, True, NAVY)
        qs = [(10, "1. What are you building?", "ምን እየገነቡ ነው?", "L_MODE",
               LISTS_DEFAULT["mode_blank"] if blank else LISTS_DEFAULT["mode"]),
              (12, "2. Which data?", "የትኛውን መረጃ?", "L_DATASET",
               LISTS_DEFAULT["data_blank"] if blank else LISTS_DEFAULT["data"]),
              (14, "3. Scenario", "ሁኔታ", "L_SCEN", LISTS_DEFAULT["scen"]),
              (16, "4. Show statements in", "የመግለጫ መለኪያ", "L_UNITS", "ETB '000")]
        for r, en, am, lst, default in qs:
            ws.cell(r, 3, en).font = F_BOLD
            ws.cell(r + 1, 3, am).font = font(10, False, SEC, name=AMFONT)
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
            c = ws.cell(r, 4, default)
            c.font = font(11, True, INPUT_FG)
            c.fill = FILL_INPUT
            c.border = BOX
            for cc in range(4, 7):
                ws.cell(r, cc).fill = FILL_INPUT
                ws.cell(r, cc).border = BOX
            dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(f"D{r}")
        b.name("PROJECT_MODE", "'00_Start'!$G$10")
        b.name("DATA_SET", "'00_Start'!$G$12")
        b.name("SCEN_SEL", "'00_Start'!$G$14")
        b.name("UNIT_SEL", "'00_Start'!$G$16")
        for r, f in ((10, "IFERROR(MATCH(D10,L_MODE,0)-1,0)"), (12, "IFERROR(MATCH(D12,L_DATASET,0),1)"),
                     (14, "IFERROR(MATCH(D14,L_SCEN,0),1)"), (16, "IFERROR(MATCH(D16,L_UNITS,0),2)")):
            c = ws.cell(r, 7, "=" + f)
            c.font = font(9, False, SEC)
            c.alignment = Alignment(horizontal="center")
        ws["G9"] = "code"
        ws["G9"].font = font(8, False, SEC)
        ws["C18"] = ('=IF(AND(DATA_SET=2,PROJECT_MODE<>0),"Tip: Sample A is a personal house – choose mode 0.",'
                     'IF(AND(DATA_SET=3,PROJECT_MODE<>3),"Tip: Sample B is a mixed-use building – choose mode 3.",'
                     'IF(PROJECT_MODE=0,"Home view: open the Home Estimate page.","Commercial view: open the Dashboard.")))')
        ws["C18"].font = font(10, True, AMBER)
        b.name("MODE_NAME", "'00_Start'!$M$10")
        b.name("SCEN_NAME", "'00_Start'!$M$14")
        b.name("UNIT_LBL", "'00_Start'!$M$16")
        ws["M16"] = '=CHOOSE(UNIT_SEL,"ETB","ETB \'000","ETB m")'
        ws["M16"].font = font(8, False, SEC)
        ws["M10"] = '=CHOOSE(PROJECT_MODE+1,"0 Personal residence","1 Build-to-Sell","2 Build-to-Rent","3 Mixed")'
        ws["M14"] = '=CHOOSE(SCEN,"Base","Upside","Downside","Lender")'
        ws["M10"].font = ws["M14"].font = font(8, False, SEC)
        ws["M9"] = "helper labels"
        ws["M9"].font = font(8, False, SEC)
        # ---- status card
        card(ws, 8, 9, 19, 10, bg=CARD)
        ws["I8"] = "STATUS  ·  ሁኔታ"
        ws["I8"].font = font(11, True, NAVY)
        rows = [(10, "Checks passed", '=CHK_PASS&" / "&CHK_TOTAL', None),
                (11, "Model status", "=CHK_STATUS", None),
                (13, "Verdict", "=VERDICT", None),
                (14, "ውሳኔ", "=VERDICT_AM", None),
                (16, "Total development cost (ETB)", "=TDC", "n0"),
                (17, '=IF(PROJECT_MODE=0,"Monthly loan instalment (ETB)","Equity IRR")',
                 "=IF(PROJECT_MODE=0,LN_INSTAL_M,EIRR)", None)]
        for r, lab, f, fmt in rows:
            ws.cell(r, 9, lab).font = F_BODY
            c = ws.cell(r, 10, f)
            c.font = font(11, True, NAVY)
            c.alignment = Alignment(horizontal="right")
            if fmt:
                c.number_format = NF[fmt]
        ws["J17"].number_format = '[<1]0.0%;#,##0'
        ws.conditional_formatting.add("J11", FormulaRule(formula=["CHK_FAIL>0"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("J11", FormulaRule(formula=["CHK_FAIL=0"], fill=fill(GOOD_BG), font=Font(color=GOOD_FG, bold=True)))
        ws.conditional_formatting.add("J13", FormulaRule(formula=['VERDICT="VIABLE"'], fill=fill(GOOD_BG), font=Font(color=GOOD_FG, bold=True)))
        ws.conditional_formatting.add("J13", FormulaRule(formula=['VERDICT="MARGINAL"'], fill=fill(WARN_BG), font=Font(color=WARN_FG, bold=True)))
        ws.conditional_formatting.add("J13", FormulaRule(formula=['VERDICT="NOT VIABLE"'], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        # ---- navigation (switches with mode)
        ws["C21"] = "NAVIGATE  ·  መሄጃ"
        ws["C21"].font = font(11, True, NAVY)
        home_nav = [("02_Home", "Home Estimate (one page) · የቤት ወጪ ግምት"), ("10_Inputs", "Inputs – block 2 home card · ግብዓቶች"),
                    ("22_Capex", "Cost breakdown · የወጪ ዝርዝር"), ("23_Construction", "Monthly cash needed · ወርሃዊ ወጪ"),
                    ("27_Loan", "Home-loan schedule · የብድር ሰንጠረዥ"), ("11_Scenarios", "Scenarios · ሁኔታዎች"),
                    ("60_Checks", "Checks · ማረጋገጫዎች"), ("61_Glossary", "Glossary · መዝገበ ቃላት")]
        com_nav = [("01_Dashboard", "Dashboard · ዳሽቦርድ"), ("10_Inputs", "Inputs · ግብዓቶች"), ("11_Scenarios", "Scenarios · ሁኔታዎች"),
                   ("40_PL", "Profit & loss · ትርፍና ኪሳራ"), ("41_BS", "Balance sheet · ሀብትና ዕዳ"), ("42_CF", "Cash flow · የገንዘብ ፍሰት"),
                   ("27_Loan", "Loan schedule · የብድር ሰንጠረዥ"), ("50_Returns", "Returns · ተመላሽ"),
                   ("52_Sensitivity", "Sensitivity & break-even · የስጋት ትንተና"), ("60_Checks", "Checks · ማረጋገጫዎች")]
        n = max(len(home_nav), len(com_nav))
        for k in range(n):
            r = 22 + k
            h = home_nav[k] if k < len(home_nav) else ("00_Start", "")
            cm_ = com_nav[k] if k < len(com_nav) else ("00_Start", "")
            f = (f'=HYPERLINK(IF(PROJECT_MODE=0,"#\'{h[0]}\'!A1","#\'{cm_[0]}\'!A1"),'
                 f'IF(PROJECT_MODE=0,"{("▸ " + h[1]) if h[1] else ""}","{("▸ " + cm_[1]) if cm_[1] else ""}"))')
            c = ws.cell(r, 3, f)
            c.font = font(10, False, BLUE, underline="single")
        ws["F21"] = "ALL SHEETS"
        ws["F21"].font = font(11, True, NAVY)
        allsheets = ["01_Dashboard", "02_Home", "10_Inputs", "11_Scenarios", "12_TaxLaw", "20_Timeline", "21_Area", "22_Capex",
                     "23_Construction", "24_Revenue", "25_Opex", "26_Assets", "27_Loan", "28_Tax", "29_VAT", "30_WorkingCap",
                     "40_PL", "41_BS", "42_CF", "43_Ratios", "50_Returns", "51_Payback", "52_Sensitivity", "60_Checks",
                     "61_Glossary"]
        for k, s in enumerate(allsheets):
            col = 6 if k < 13 else 7
            r = 22 + (k if k < 13 else k - 13)
            c = ws.cell(r, col, f'=HYPERLINK("#\'{s}\'!A1","{s.replace("_", " ")}")')
            c.font = font(9, False, BLUE, underline="single")
        # ---- how to use
        r0 = 37
        ws.cell(r0, 3, "HOW TO USE  ·  አጠቃቀም").font = font(11, True, NAVY)
        steps = [("1", "Answer question 1. Personal house → 0. Selling, renting or both → 1, 2 or 3.",
                  "ጥያቄ 1ን ይመልሱ፡ የግል ቤት → 0፤ ለሽያጭ፣ ለኪራይ ወይም ለሁለቱም → 1፣ 2 ወይም 3።"),
                 ("2", "Choose 'My project' in question 2 and type your values in column E of sheet 10 Inputs (blue cells only).",
                  "በጥያቄ 2 'የእኔ ፕሮጀክት'ን ይምረጡና በገጽ 10 ዓምድ E ሰማያዊ ሕዋሶች ላይ ብቻ ይጻፉ።"),
                 ("3", "Check the status box: all checks must pass. Then read the Home Estimate (Mode 0) or the Dashboard.",
                  "የማረጋገጫ ሳጥኑን ይመልከቱ፤ ከዚያ የቤት ግምቱን ወይም ዳሽቦርዱን ያንብቡ።"),
                 ("4", "Stress-test: switch scenarios (question 3) and use the flex cells on sheet 52.",
                  "ሁኔታዎችን ይቀያይሩና በገጽ 52 የማስተካከያ ሕዋሶችን ይጠቀሙ።")]
        for k, (nmb, en, am) in enumerate(steps):
            r = r0 + 1 + 2 * k
            ws.cell(r, 2, nmb).font = font(13, True, AMBER)
            ws.cell(r, 3, en).font = F_BODY
            ws.cell(r + 1, 3, am).font = font(10, False, SEC, name=AMFONT)
        r1 = r0 + 10
        ws.cell(r1, 3, "COLOUR CODE  ·  የቀለም ኮድ").font = font(11, True, NAVY)
        leg = [("Input – type here", "ግብዓት – እዚህ ይጻፉ", F_INPUT, FILL_INPUT),
               ("Formula – do not overwrite", "ቀመር – አይጻፉበት", font(10, False, INK), fill("FFFFFF")),
               ("Link from another sheet", "ከሌላ ገጽ ማገናኛ", font(10, False, LINK_FG), fill("FFFFFF")),
               ("Key output / check", "ቁልፍ ውጤት / ማረጋገጫ", font(10, True, INK), FILL_KEY)]
        for k, (en, am, f, fl) in enumerate(leg):
            c = ws.cell(r1 + 1 + k, 3, en)
            c.font = f
            c.fill = fl
            c.border = BOX
            ws.cell(r1 + 1 + k, 4, am).font = font(10, False, SEC, name=AMFONT)
        r2 = r1 + 6
        ws.cell(r2, 3, "METHOD NOTES").font = font(11, True, NAVY)
        notes = [
            "No circularity: interest during construction (IDC) is built month by month on the opening balance plus draws; "
            "the arrangement fee is charged on a facility base that excludes IDC; undistributed-profits tax uses a closed form.",
            "Timeline: construction is modelled monthly (up to 48 months) and rolled up to C1–C4; operations run O1–O10 annually. "
            "Returns use dated cash flows (XIRR).",
            "IFRS: IAS 2 units for sale, IAS 40 let space (cost model), IAS 16 owner-occupied, IAS 23 capitalised borrowing cost, "
            "IFRS 15 sales, IFRS 16 rent, IAS 12 tax, IAS 36 impairment test, IAS 7 indirect cash flow.",
            "Sensitivity tables are Excel Data Tables. If Excel is set to 'Automatic except data tables', press F9.",
            "Disclaimer: a planning tool only – not investment, tax or legal advice. Confirm tax rates (sheet 12) before use.",
            "ማሳሰቢያ፡ ይህ የዕቅድ መሣሪያ ብቻ ነው፤ የኢንቨስትመንት፣ የግብር ወይም የሕግ ምክር አይደለም።",
        ]
        for k, t in enumerate(notes):
            c = ws.cell(r2 + 1 + k, 3, t)
            c.font = F_NOTE if k < 5 else font(9, False, SEC, italic=True, name=AMFONT)
        ws.print_area = "A1:K60"
        ws.page_setup.fitToHeight = 1
    sh.raw(_s)
    sh.r = 70
    return sh


LISTS_DEFAULT = {
    "mode": LISTS["L_MODE"][3], "mode_blank": LISTS["L_MODE"][0],
    "data": LISTS["L_DATASET"][2], "data_blank": LISTS["L_DATASET"][0],
    "scen": LISTS["L_SCEN"][0],
}


# ---------------------------------------------------------------------- HOME ESTIMATE
def build_home(b):
    sh = Sheet(b, "02_Home", "Home Estimate", "የቤት ግንባታ ወጪ ግምት", TAB_REP, kind="free", n=0, width_label=27)
    ws = sh.ws
    cm = b.sheets["23_Construction"]
    ln = b.sheets["27_Loan"]

    def hdr_row(r, cols, labels):
        for c, h in zip(cols, labels):
            x = ws.cell(r, c, h)
            x.font = font(9, True, "FFFFFF")
            x.fill = fill(NAVY)
            x.alignment = Alignment(horizontal="left" if c in (2, 7) else "right")

    def title(r, c, t):
        x = ws.cell(r, c, t)
        x.font = font(11, True, NAVY)

    def _h(_r):
        for col, w in (("A", 2), ("B", 27), ("C", 14), ("D", 10), ("E", 9), ("F", 2), ("G", 9), ("H", 13), ("I", 13),
                       ("J", 13), ("K", 2)):
            ws.column_dimensions[col].width = w
        ws["B5"] = ('=IF(PROJECT_MODE<>0,"Commercial mode is on – this page still shows the cost estimate for the same inputs.  ·  '
                    'ይህ ገጽ ለተመሳሳይ ግብዓቶች የወጪ ግምቱን ያሳያል።","")')
        ws["B5"].font = font(9, True, AMBER)
        info = [(6, "Project · ፕሮጀክት", "=PROJ_NAME"),
                (7, "Building · ሕንፃ", '=BTYPE_CALC&"  ·  "&TEXT(GFA,"#,##0")&" m² gross  ·  "&INDEX(L_FINISH,FIN_IX)&" finish"'),
                (8, "Where & how long · ቦታና ጊዜ", '=INDEX(L_LOC,LOC_TYPE)&"  ·  "&CONSTR_EFF&" months of construction"')]
        for r, lab, f in info:
            ws.cell(r, 2, lab).font = font(9, True, SEC)
            c = ws.cell(r, 3, f)
            c.font = font(10, True, NAVY)
        # tiles
        tiles = [((2, 2), "CONSTRUCTION COST · የግንባታ ወጪ", "=CONSTR_COST", "n0",
                  '="Range "&TEXT(CONSTR_COST*(1+EST_LOW)/1000000,"#,##0.0")&" – "&TEXT(CONSTR_COST*(1+EST_HIGH)/1000000,"#,##0.0")&" m"'),
                 ((3, 5), "COST PER m² · ዋጋ በካሬ ሜትር", "=CC_M2", "n0",
                  '="All-in incl. land & finance: "&TEXT(TDC_M2G,"#,##0")'),
                 ((7, 8), "MONTHLY INSTALMENT · ወርሃዊ ክፍያ", "=LN_INSTAL_M", "n0",
                  '=IF(HH_INCOME>0,TEXT(AFFORD_PCT,"0%")&" of income · "&AFFORD_TXT,"enter income")'),
                 ((9, 10), "FUNDING GAP · የገንዘብ ክፍተት", "=FUND_GAP", "n0",
                  '=IF(FUND_GAP>0.5,"find more funds","covered · በቂ ነው")')]
        accents = [BLUE, TEAL, VIOLET, AMBER]
        for k, ((c1, c2), lab, f, fmt, sub) in enumerate(tiles):
            card(ws, 10, c1, 13, c2, bg=CANVAS)
            for cc in range(c1, c2 + 1):
                ws.cell(10, cc).border = Border(top=Side(style="thick", color=accents[k]))
            if c2 > c1:
                for rr in (10, 11, 13):
                    ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c2)
            ws.cell(10, c1, lab).font = font(8, True, SEC)
            v = ws.cell(11, c1, f)
            v.font = font(16, True, NAVY)
            v.number_format = NF[fmt]
            v.alignment = Alignment(horizontal="left")
            s_ = ws.cell(13, c1, sub)
            s_.font = font(8, False, SEC)
        ws.row_dimensions[11].height = 24
        # cost breakdown
        title(15, 2, "COST BREAKDOWN · የወጪ ክፍፍል")
        title(15, 7, "WHERE THE MONEY GOES")
        hdr_row(16, (2, 3, 4), ("Item", "ETB", "% total"))
        labs_am = ["መሠረትና ምድር ቤት", "ፍሬምና ግድግዳ", "የውስጥና ማጠናቀቂያ", "ኤሌክትሪክ፣ ቧንቧ፣ ሲስተሞች", "ቅድመ-ሥራና የውጭ ሥራ",
                   "የዋጋ ጭማሪና መጠባበቂያ", "ክፍያዎችና ፈቃድ", "የቤት ዕቃዎች", "መሬት", "የብድር ወጪ"]
        for k, cr in enumerate(b.CAT_ROWS):
            r = 17 + k
            ws.cell(r, 2, f"='22_Capex'!B{cr}").font = font(9)
            ws.cell(r, 5, labs_am[k]).font = font(7, False, SEC, name=AMFONT)
            c = ws.cell(r, 3, f"='22_Capex'!D{cr}")
            c.number_format = NF["n0"]
            c.font = font(9)
            c2 = ws.cell(r, 4, f"=IF(HOME_TOTAL>0,C{r}/HOME_TOTAL,0)")
            c2.number_format = NF["pct"]
            c2.font = font(9)
        tr = 17 + len(b.CAT_ROWS)
        ws.cell(tr, 2, "TOTAL PROJECT COST · ጠቅላላ").font = F_BOLD
        c = ws.cell(tr, 3, f"=SUM(C17:C{tr - 1})")
        c.number_format = NF["n0"]
        c.font = F_BOLD
        c.fill = FILL_KEY
        b.name("HOME_TOTAL", f"'02_Home'!$C${tr}")
        for col in (2, 3, 4, 5):
            ws.cell(tr, col).border = Border(top=Side(style="thin", color=INK))
        ch = BarChart()
        ch.type = "bar"
        ch.style = 10
        ch.legend = None
        ch.y_axis.numFmt = '#,##0.0,,"m"'
        ch.y_axis.majorGridlines = None
        ch.add_data(Reference(ws, min_col=3, min_row=17, max_row=tr - 1), titles_from_data=False)
        ch.set_categories(Reference(ws, min_col=2, min_row=17, max_row=tr - 1))
        ch.series[0].graphicalProperties.solidFill = BLUE
        ch.series[0].graphicalProperties.line.solidFill = BLUE
        ch.x_axis.scaling.orientation = "maxMin"
        ch.height, ch.width = 6.2, 9.6
        ch.gapWidth = 40
        axis_font(ch.x_axis, 650)
        axis_font(ch.y_axis, 650)
        ws.add_chart(ch, "G16")
        # payment stages
        r = tr + 2
        title(r, 2, "PAYING THE CONTRACTOR · ለኮንትራክተር ክፍያ")
        title(r, 7, "CASH NEEDED EACH MONTH (S-curve)")
        hdr_row(r + 1, (2, 3, 4, 5), ("Stage", "ETB", "Cum. %", "≈ month"))
        st0 = r + 2
        stage_am = ["ቅድመ-ዝግጅት", "መሠረት", "ፍሬም", "ግድግዳና ጣሪያ", "ኤሌክትሪክ፣ ቧንቧ", "ማጠናቀቂያ"]
        cumrow = f"'23_Construction'!$E${cm.rows['CUM']}:${pc(MON_C)}${cm.rows['CUM']}"
        for k, sr in enumerate(b.STAGE_ROWS):
            rr = st0 + k
            ws.cell(rr, 2, f"='22_Capex'!B{sr}").font = font(9)
            c = ws.cell(rr, 3, f"='22_Capex'!D{sr}")
            c.number_format = NF["n0"]
            c.font = font(9)
            c2 = ws.cell(rr, 4, f"=IF(STAGE_TOT>0,SUM($C${st0}:C{rr})/STAGE_TOT,0)")
            c2.number_format = NF["pct"]
            c2.font = font(9)
            c3 = ws.cell(rr, 5, f'=MIN(CONSTR_EFF,COUNTIF({cumrow},"<"&(D{rr}-0.0001))+1)')
            c3.number_format = NF["int"]
            c3.font = font(9)
        rr = st0 + len(b.STAGE_ROWS)
        ws.cell(rr, 2, "Retention released after completion").font = F_NOTE
        c = ws.cell(rr, 3, "=RET_TOT")
        c.number_format = NF["n0"]
        c.font = font(9)
        c = ws.cell(rr, 5, "=CONSTR_EFF+RET_REL_M")
        c.number_format = NF["int"]
        c.font = font(9)
        ch2 = BarChart()
        ch2.style = 10
        ch2.y_axis.numFmt = '#,##0.0,,"m"'
        ch2.y_axis.majorGridlines = None
        cws = b.wb["23_Construction"]
        ch2.add_data(Reference(cws, min_col=5, max_col=4 + MON_C, min_row=cm.rows["NEED"]), from_rows=True, titles_from_data=False)
        ch2.set_categories(Reference(b.wb["20_Timeline"], min_col=5, max_col=4 + MON_C, min_row=b.TL["M_T"]))
        ch2.series[0].graphicalProperties.solidFill = AMBER
        ch2.series[0].graphicalProperties.line.solidFill = AMBER
        ch2.legend = None
        ch2.height, ch2.width = 4.2, 9.6
        ch2.gapWidth = 30
        axis_font(ch2.x_axis, 600)
        axis_font(ch2.y_axis, 650)
        ws.add_chart(ch2, f"G{st0 - 1}")
        # funding & loan
        r = rr + 2
        title(r, 2, "FUNDING & HOME LOAN · ገንዘብና ብድር")
        title(r, 7, "HOME-LOAN SCHEDULE · የብድር ሰንጠረዥ")
        fl = [("Own savings · ቁጠባ", "=SAVINGS", "n0"), ("Family / diaspora · ቤተሰብ", "=FAMILY", "n0"),
              ("Home loan drawn · የወጣ ብድር", "=DEBT_DRAWN", "n0"), ("Funding gap · ክፍተት", "=FUND_GAP", "n0"),
              ("Interest rate · ወለድ", "=RATE_EFF", "pct"), ("Tenor after completion (yrs)", "=TENOR_Y", "int"),
              ("Monthly instalment · ወርሃዊ ክፍያ", "=LN_INSTAL_M", "n0"), ("Total interest · ጠቅላላ ወለድ", "=LN_INT_TOT", "n0"),
              ("Loan repaid by · የሚያልቅበት", "=LN_FINAL_DATE", "date"),
              ("Instalment ÷ income · ከገቢ ድርሻ", "=IF(HH_INCOME>0,LN_INSTAL_M/HH_INCOME,0)", "pct"),
              ("Affordability · የመክፈል አቅም", '=IF(HH_INCOME<=0,"n/a",IF(AFFORD_PCT<=AFFORD_MAX,"Affordable · ይቻላል",'
                                               'IF(AFFORD_PCT<=AFFORD_MAX*1.25,"Stretch · ከባድ","Not affordable · አይቻልም")))', "gen")]
        for k, (lab, f, fmt) in enumerate(fl):
            rr2 = r + 1 + k
            ws.cell(rr2, 2, lab).font = font(9)
            ws.merge_cells(start_row=rr2, start_column=3, end_row=rr2, end_column=5)
            c = ws.cell(rr2, 3, f)
            c.number_format = NF[fmt]
            c.font = font(9, True, NAVY)
            c.alignment = Alignment(horizontal="right")
        ar_ = r + len(fl)
        b.name("AFFORD_PCT", f"'02_Home'!$C${ar_ - 1}")
        b.name("AFFORD_TXT", f"'02_Home'!$C${ar_}")
        ws.conditional_formatting.add(f"C{ar_}:E{ar_}", FormulaRule(formula=["AND(HH_INCOME>0,AFFORD_PCT<=AFFORD_MAX)"], fill=fill(GOOD_BG), font=Font(color=GOOD_FG, bold=True)))
        ws.conditional_formatting.add(f"C{ar_}:E{ar_}", FormulaRule(formula=["AND(HH_INCOME>0,AFFORD_PCT>AFFORD_MAX)"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add(f"C{r + 4}:E{r + 4}", FormulaRule(formula=["FUND_GAP>0.5"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        hdr_row(r + 1, (7, 8, 9, 10), ("Year", "Interest", "Principal", "Closing"))
        LYr = f"'27_Loan'!$E${ln.rows['LY']}:${pc(MON_L)}${ln.rows['LY']}"
        Kr = f"'27_Loan'!$E${ln.rows['K']}:${pc(MON_L)}${ln.rows['K']}"
        rowr = lambda k: f"'27_Loan'!$E${ln.rows[k]}:${pc(MON_L)}${ln.rows[k]}"
        c = ws.cell(r + 2, 7, "Build")
        c.font = font(8)
        c = ws.cell(r + 2, 8, f"=SUMIFS({rowr('INT')},{LYr},0)")
        c.number_format = NF["n0"]
        c.font = font(8)
        c = ws.cell(r + 2, 10, f"=SUMIFS({rowr('CLOSE')},{Kr},0)")
        c.number_format = NF["n0"]
        c.font = font(8)
        for y in range(1, 21):
            rr2 = r + 2 + y
            c = ws.cell(rr2, 7, y)
            c.number_format = NF["int"]
            c.font = font(8)
            c.alignment = Alignment(horizontal="left")
            for col, k in ((8, "INT"), (9, "PRIN")):
                f = f"=SUMIFS({rowr(k)},{LYr},$G{rr2})" + (f"+SUMIFS({rowr('SWP')},{LYr},$G{rr2})" if k == "PRIN" else "")
                c = ws.cell(rr2, col, f)
                c.number_format = NF["n0"]
                c.font = font(8)
            c = ws.cell(rr2, 10, f"=SUMIFS({rowr('CLOSE')},{Kr},12*$G{rr2})")
            c.number_format = NF["n0"]
            c.font = font(8)
        # rent-out upgrade
        ru = ar_ + 2
        title(ru, 2, "OPTIONAL: RENT OUT PART · ከፊሉን ማከራየት")
        rvr = b.sheets["24_Revenue"].rows
        rows = [("Switched on? (Inputs block 2)", '=IF(RENT_HOME_F=1,"Yes · አዎ","No · አይ")', "gen"),
                ("Rent income, first year · ኪራይ", f"=IF(RENT_HOME_F=1,'24_Revenue'!$I${rvr['RENT']},0)", "n0"),
                ("Gross yield on total cost · ገቢ ከወጪ", f"=IF(AND(RENT_HOME_F=1,HOME_TOTAL>0),'24_Revenue'!$K${rvr['RENT']}/HOME_TOTAL,0)", "pct"),
                ("Simple payback (years) · የመመለሻ ጊዜ", f"=IF(AND(RENT_HOME_F=1,'24_Revenue'!$K${rvr['RENT']}>0),HOME_TOTAL/'24_Revenue'!$K${rvr['RENT']},0)", "n1"),
                ("Rent ÷ monthly instalment · የሚሸፍነው", f"=IF(AND(RENT_HOME_F=1,LN_INSTAL_M>0),'24_Revenue'!$I${rvr['RENT']}/12/LN_INSTAL_M,0)", "pct")]
        for k, (lab, f, fmt) in enumerate(rows):
            rr2 = ru + 1 + k
            ws.cell(rr2, 2, lab).font = font(9)
            ws.merge_cells(start_row=rr2, start_column=3, end_row=rr2, end_column=5)
            c = ws.cell(rr2, 3, f)
            c.number_format = NF[fmt]
            c.font = font(9, True, NAVY)
            c.alignment = Alignment(horizontal="right")
        last = max(ru + 6, r + 23) + 1
        ws.cell(last, 2, "Planning estimate only – get a priced BOQ from a licensed quantity surveyor before signing a contract.").font = F_NOTE
        ws.cell(last + 1, 2, "ይህ ግምት ብቻ ነው፤ ውል ከመፈረምዎ በፊት ፈቃድ ካለው ባለሙያ የተዘጋጀ የሥራ መጠን ዝርዝር (BOQ) ያግኙ።").font = font(9, False, SEC, italic=True, name=AMFONT)
        ws.print_area = f"A1:K{last + 2}"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToHeight = 1
        ws.page_setup.fitToWidth = 1
        ws.freeze_panes = None
    sh.raw(_h)
    sh.r = 95
    return sh


# ---------------------------------------------------------------------- DASHBOARD
def build_dashboard(b):
    sh = Sheet(b, "01_Dashboard", "Dashboard", "ዳሽቦርድ – ፕሮጀክቴ ያተርፋል?", TAB_REP, kind="free", n=0, width_label=3)
    ws = sh.ws

    def _d(_r):
        widths = {"A": 2, "B": 15, "C": 15, "D": 3, "E": 15, "F": 15, "G": 3, "H": 15, "I": 15, "J": 3, "K": 15, "L": 15,
                  "M": 3, "N": 15, "O": 15, "P": 3, "Q": 15, "R": 15, "S": 2}
        for k, v in widths.items():
            ws.column_dimensions[k].width = v
        card(ws, 5, 2, 6, 18, bg=CANVAS)
        for c, t in (("B5", "Scenario"), ("E5", "Checks"), ("H5", "Mode"), ("K5", "Verdict · ውሳኔ"), ("Q5", "Equity IRR hurdle")):
            ws[c] = t
            ws[c].font = font(9, True, SEC)
        for c, f in (("C5", "=SCEN_NAME"), ("F5", '=CHK_PASS&" / "&CHK_TOTAL&" pass"'), ("I5", "=MODE_NAME"),
                     ("L5", '=VERDICT&"  ·  "&VERDICT_AM'), ("R5", "=HURDLE")):
            ws[c] = f
            ws[c].font = font(11, True, NAVY)
        ws["R5"].number_format = NF["pct"]
        ws.merge_cells("L5:O5")
        ws["L5"].alignment = Alignment(horizontal="center")
        for rule_f, bg, fg in (('VERDICT="VIABLE"', GOOD_BG, GOOD_FG), ('VERDICT="MARGINAL"', WARN_BG, WARN_FG),
                               ('VERDICT="NOT VIABLE"', CRIT_BG, CRIT_FG)):
            ws.conditional_formatting.add("L5:O5", FormulaRule(formula=[rule_f], fill=fill(bg), font=Font(color=fg, bold=True)))
        ws.conditional_formatting.add("F5", FormulaRule(formula=["CHK_FAIL>0"], font=Font(color=CRIT_FG, bold=True)))
        ws["B6"] = ('=IF(PROJECT_MODE=0,"Home view – tiles show the cost estimate; open the Home Estimate page for the full picture.",'
                    '"Commercial view – ETB; IRR on dated cash flows (XIRR); verdict = equity IRR vs hurdle, profit on cost and DSCR covenant.")')
        ws["B6"].font = font(8, False, SEC, italic=True)
        tiles = [
            ('=IF(PROJECT_MODE=0,"TOTAL PROJECT COST","TOTAL DEVELOPMENT COST")', "=TDC", "m",
             '="ETB  ·  "&TEXT(TDC_M2G,"#,##0")&" per m² gross"'),
            ('=IF(PROJECT_MODE=0,"COST PER m²","GROSS DEVELOPMENT VALUE")', "=IF(PROJECT_MODE=0,TEXT(CC_M2,\"#,##0\"),GDV)",
             "m", '=IF(PROJECT_MODE=0,"construction cost per m²","GDV ÷ TDC = "&TEXT(GDV_TDC,"0.00")&"x")'),
            ('=IF(PROJECT_MODE=0,"MONTHLY INSTALMENT","PROFIT ON COST")', "=IF(PROJECT_MODE=0,TEXT(LN_INSTAL_M,\"#,##0\"),POC)",
             "pct", '=IF(PROJECT_MODE=0,"home loan","margin on GDV "&TEXT(DEV_MARGIN,"0.0%"))'),
            ('=IF(PROJECT_MODE=0,"FUNDING GAP","EQUITY IRR")', "=IF(PROJECT_MODE=0,TEXT(FUND_GAP,\"#,##0\"),EIRR)", "pct",
             '=IF(PROJECT_MODE=0,IF(FUND_GAP>0.5,"to be found","fully funded"),"hurdle "&TEXT(HURDLE,"0.0%")&IF(EIRR>=HURDLE,"  ✓ above","  ✗ below"))'),
            ('=IF(PROJECT_MODE=0,"LOAN DRAWN","PEAK DEBT")', "=IF(PROJECT_MODE=0,DEBT_DRAWN,PEAK_DEBT)", "m",
             '="incl. capitalised IDC "&TEXT(IDC_TOT/1000000,"#,##0.0")&" m"'),
            ('=IF(PROJECT_MODE=0,"AFFORDABILITY","MIN DSCR")', "=IF(PROJECT_MODE=0,TEXT(AFFORD_PCT,\"0%\"),MIN_DSCR)", "x",
             '=IF(PROJECT_MODE=0,AFFORD_TXT,"covenant "&TEXT(DSCR_COV,"0.00")&"x"&IF(DSCR_OK=1,"  ✓ met","  ✗ breached"))'),
        ]
        cols = [2, 5, 8, 11, 14, 17]
        acc = [BLUE, TEAL, VIOLET, AMBER, ROSE, NAVY]
        for k, t in enumerate(tiles):
            lab, f = t[0], t[1]
            fmt = t[2] if len(t) == 4 else None
            sub = t[-1]
            c0 = cols[k]
            card(ws, 8, c0, 12, c0 + 1, bg=CARD)
            for cc in range(c0, c0 + 2):
                ws.cell(8, cc).border = Border(top=Side(style="thick", color=acc[k]))
            for rr in (8, 10, 12):
                ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=c0 + 1)
            ws.cell(8, c0, lab).font = font(8, True, SEC)
            v = ws.cell(10, c0, f)
            v.font = font(18, True, NAVY)
            v.alignment = Alignment(horizontal="left", vertical="center")
            if fmt:
                v.number_format = NF.get(fmt, fmt)
            s_ = ws.cell(12, c0, sub)
            s_.font = font(8, False, SEC)
        ws.row_dimensions[10].height = 28
        ws.conditional_formatting.add("K10:L10", FormulaRule(formula=["AND(PROJECT_MODE<>0,EIRR<HURDLE)"], font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("Q10:R10", FormulaRule(formula=["AND(PROJECT_MODE<>0,DSCR_OK=0)"], font=Font(color=CRIT_FG, bold=True)))
        # chart data block
        d0 = 62
        ws.cell(d0, 2, "Chart data (linked – do not edit)").font = font(9, True, SEC)
        rt = b.sheets["50_Returns"].rows
        pb = b.sheets["51_Payback"].rows
        rv = b.sheets["24_Revenue"].rows
        lnr = b.sheets["27_Loan"].rows
        ra = b.sheets["43_Ratios"].rows
        series = [("Period", lambda i: f"='20_Timeline'!{pc(i)}{b.TL['T_LABEL']}"),
                  ("Project cash flow", lambda i: f"='50_Returns'!{pc(i)}{rt['PCF']}/1000000"),
                  ("Cumulative", lambda i: f"='51_Payback'!{pc(i)}{pb['PCUM']}/1000000"),
                  ("Loan balance", lambda i: f"='27_Loan'!{pc(i)}{lnr['A_CLOSE']}/1000000"),
                  ("DSCR", lambda i: f"=IF(ISNUMBER('43_Ratios'!{pc(i)}{ra['DSCR']}),MIN(10,'43_Ratios'!{pc(i)}{ra['DSCR']}),0)"),
                  ("Covenant", lambda i: "=DSCR_COV"),
                  ("Sales", lambda i: f"='24_Revenue'!{pc(i)}{rv['SALES_REV']}/1000000"),
                  ("Rent", lambda i: f"='24_Revenue'!{pc(i)}{rv['RENT']}/1000000"),
                  ("Other", lambda i: f"=('24_Revenue'!{pc(i)}{rv['KEYA']}+'24_Revenue'!{pc(i)}{rv['SCR']}+'24_Revenue'!{pc(i)}{rv['OTH']})/1000000")]
        for k, (lab, f) in enumerate(series):
            ws.cell(d0 + 1 + k, 2, lab).font = font(8, False, SEC)
            for i in range(1, ANN_N + 1):
                c = ws.cell(d0 + 1 + k, 2 + i, f(i))
                c.number_format = "0.0" if k else "@"
                c.font = font(8, False, SEC)
        R = lambda k: d0 + 1 + k
        cats = Reference(ws, min_col=3, max_col=2 + ANN_N, min_row=R(0))

        def mk_bar(rows, colors, stacked=False):
            ch = BarChart()
            ch.style = 10
            for k, row in enumerate(rows):
                ch.add_data(Reference(ws, min_col=2, max_col=2 + ANN_N, min_row=row), from_rows=True, titles_from_data=True)
                ch.series[-1].graphicalProperties.solidFill = colors[k]
                ch.series[-1].graphicalProperties.line.solidFill = colors[k]
            ch.set_categories(cats)
            if stacked:
                ch.grouping = "stacked"
                ch.overlap = 100
            ch.y_axis.majorGridlines = None
            ch.y_axis.numFmt = "#,##0"
            ch.x_axis.tickLblPos = "low"
            ch.legend.position = "b"
            ch.height, ch.width = 6.0, 11.6
            ch.gapWidth = 50
            axis_font(ch.x_axis, 650)
            axis_font(ch.y_axis, 650)
            return ch

        def ctitle(cell, t):
            ws[cell] = t
            ws[cell].font = font(10, True, NAVY)
        ctitle("B14", "(a) Project cash flow and cumulative · ETB m")
        ctitle("H14", "(b) Loan balance at period end · ETB m")
        ctitle("N14", "(b) DSCR by year vs covenant · x (capped at 10)")
        ctitle("B28", "(c) Revenue mix · ETB m")
        ctitle("H28", "(d) Development cost breakdown · ETB m")
        ctitle("N28", "Tornado – Equity IRR swing (±20%, delay +6 mo)")
        a = mk_bar([R(1)], [BLUE])
        ln_ = LineChart()
        ln_.add_data(Reference(ws, min_col=2, max_col=2 + ANN_N, min_row=R(2)), from_rows=True, titles_from_data=True)
        ln_.series[0].graphicalProperties.line.solidFill = AMBER
        ln_.series[0].graphicalProperties.line.width = 28000
        a += ln_
        ws.add_chart(a, "B15")
        ws.add_chart(mk_bar([R(3)], [VIOLET]), "H15")
        dch = mk_bar([R(4)], [TEAL])
        cl = LineChart()
        cl.add_data(Reference(ws, min_col=2, max_col=2 + ANN_N, min_row=R(5)), from_rows=True, titles_from_data=True)
        cl.series[0].graphicalProperties.line.solidFill = ROSE
        cl.series[0].graphicalProperties.line.dashStyle = "dash"
        dch.y_axis.numFmt = "0.0"
        dch += cl
        ws.add_chart(dch, "N15")
        ws.add_chart(mk_bar([R(6), R(7), R(8)], [BLUE, TEAL, VIOLET], stacked=True), "B29")
        cx = b.wb["22_Capex"]
        dd = BarChart()
        dd.type = "bar"
        dd.style = 10
        dd.legend = None
        dd.add_data(Reference(cx, min_col=4, min_row=b.CAT_ROWS[0], max_row=b.CAT_ROWS[-1]), titles_from_data=False)
        dd.set_categories(Reference(cx, min_col=2, min_row=b.CAT_ROWS[0], max_row=b.CAT_ROWS[-1]))
        dd.series[0].graphicalProperties.solidFill = AMBER
        dd.series[0].graphicalProperties.line.solidFill = AMBER
        dd.y_axis.numFmt = '#,##0,,'
        dd.y_axis.majorGridlines = None
        dd.x_axis.scaling.orientation = "maxMin"
        dd.height, dd.width = 6.0, 11.6
        dd.gapWidth = 30
        axis_font(dd.x_axis, 650)
        axis_font(dd.y_axis, 650)
        ws.add_chart(dd, "H29")
        se = b.wb["52_Sensitivity"]
        t0 = b.TORNADO["s0"]
        n = b.TORNADO["n"]
        tc = BarChart()
        tc.type = "bar"
        tc.style = 10
        tc.add_data(Reference(se, min_col=5, max_col=6, min_row=t0 + 1, max_row=t0 + n), titles_from_data=False)
        tc.set_categories(Reference(se, min_col=2, min_row=t0 + 1, max_row=t0 + n))
        from openpyxl.chart.series import SeriesLabel
        tc.series[0].tx = SeriesLabel(v="Low case")
        tc.series[1].tx = SeriesLabel(v="High case")
        tc.series[0].graphicalProperties.solidFill = ROSE
        tc.series[0].graphicalProperties.line.solidFill = ROSE
        tc.series[1].graphicalProperties.solidFill = TEAL
        tc.series[1].graphicalProperties.line.solidFill = TEAL
        tc.grouping = "clustered"
        tc.overlap = 100
        tc.gapWidth = 40
        tc.x_axis.scaling.orientation = "maxMin"
        tc.x_axis.tickLblPos = "low"
        tc.y_axis.numFmt = "0%"
        tc.y_axis.majorGridlines = None
        tc.legend.position = "b"
        tc.height, tc.width = 6.0, 11.6
        axis_font(tc.x_axis, 650)
        axis_font(tc.y_axis, 650)
        ws.add_chart(tc, "N29")
        # sources & uses
        s0 = 42
        ws.cell(s0, 2, "SOURCES & USES – construction period (ETB)").font = font(11, True, NAVY)
        cmr = b.sheets["23_Construction"].rows
        uses = [("TDC", "=TDC"),
                ("Capex VAT (refundable)", f"=SUM('23_Construction'!$E${cmr['VATC']}:${pc(MON_C)}${cmr['VATC']})"),
                ("Lease paid in build", f"=SUM('23_Construction'!$E${cmr['LSI']}:${pc(MON_C)}${cmr['LSINT']})")]
        srcs = [("Equity committed", "=EQ_DRAWN"), ("Top-up / gap", "=FUND_GAP"), ("Bank debt drawn", "=DEBT_DRAWN"),
                ("Presale cash used", "=PS_USED"), ("IDC capitalised", "=IDC_TOT"),
                ("Land non-cash", "=LAND_TOT-LANDCASH_TOT"), ("Retention held", "=RET_TOT")]
        for c, t in ((2, "Uses"), (5, "Sources")):
            x = ws.cell(s0 + 1, c, t)
            x.font = font(9, True, "FFFFFF")
            x.fill = fill(NAVY)
        for c in (3, 6):
            ws.cell(s0 + 1, c).fill = fill(NAVY)
        for k, (lab, f) in enumerate(uses):
            ws.cell(s0 + 2 + k, 2, lab).font = font(9)
            c = ws.cell(s0 + 2 + k, 3, f)
            c.number_format = NF["n0"]
            c.font = font(9)
        for k, (lab, f) in enumerate(srcs):
            ws.cell(s0 + 2 + k, 5, lab).font = font(9)
            c = ws.cell(s0 + 2 + k, 6, f)
            c.number_format = NF["n0"]
            c.font = font(9)
        tr = s0 + 2 + max(len(uses), len(srcs))
        for c, t, f in ((2, "Total uses", f"=SUM(C{s0 + 2}:C{tr - 1})"), (5, "Total sources", f"=SUM(F{s0 + 2}:F{tr - 1})")):
            ws.cell(tr, c, t).font = font(9, True)
            x = ws.cell(tr, c + 1, f)
            x.number_format = NF["n0"]
            x.font = font(9, True)
            for cc in (c, c + 1):
                ws.cell(tr, cc).border = Border(top=Side(style="thin", color=INK))
        ws.cell(tr + 1, 5, "Check (= 0)").font = F_NOTE
        x = ws.cell(tr + 1, 6, f"=ROUND(F{tr}-C{tr},0)")
        x.number_format = NF["n0"]
        x.font = F_NOTE
        # alerts
        ws.cell(s0, 11, "ALERTS · ማስጠንቀቂያዎች").font = font(11, True, NAVY)
        alerts = [
            '=IF(CHK_FAIL>0,"✗ "&CHK_FAIL&" model check(s) failing – see sheet 60","✓ All model checks pass")',
            '=IF(PROJECT_MODE=0,"–",IF(N_BREACH>0,"✗ DSCR below covenant in "&N_BREACH&" year(s)","✓ DSCR covenant met"))',
            '=IF(FUND_GAP>0.5,"✗ Funding gap of "&TEXT(FUND_GAP,"#,##0")&" ETB","✓ Funding fully covered")',
            '=IF(PROJECT_MODE=0,"–",IF(N_LOWCASH>0,"! Cash below minimum in "&N_LOWCASH&" year(s)","✓ Cash above minimum"))',
            '=IF(PROJECT_MODE=0,"–",IF(N_IMPAIR>0,"! IAS 36 impairment indicator in "&N_IMPAIR&" year(s)","✓ No impairment indicator"))',
            '=IF(AND(FAR>0,FAR_ACT>FAR),"! FAR "&TEXT(FAR_ACT,"0.00")&" exceeds permitted "&TEXT(FAR,"0.00"),"✓ FAR within limit")',
            '=IF(PROJECT_MODE=0,"–",IF(EIRR<HURDLE,"! Equity IRR below hurdle","✓ Equity IRR above hurdle"))',
            '="! Tax rates must be verified (sheet 12) · የግብር መጣኔዎችን ያረጋግጡ"',
        ]
        for k, f in enumerate(alerts):
            ws.cell(s0 + 1 + k, 11, f).font = font(9)
        rng = f"K{s0 + 1}:K{s0 + len(alerts)}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="✗"'], font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="!"'], font=Font(color=WARN_FG, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="✓"'], font=Font(color=GOOD_FG)))
        ws.print_area = "A1:S52"
        ws.page_setup.fitToHeight = 1
        ws.freeze_panes = None
    sh.raw(_d)
    sh.r = 80
    return sh
