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
from inputs_def import (INPUTS, LISTS, RATE_LIB, EL_KEYS, ELEMENTS, ELEMENTS_AM, SAMPLE_NOTE, confidence, SPACE_TYPES,
                        N_SPACES, SPACES_USER, SPACES_S1, SPACES_S2, T_HOTEL, T_RES, T_OWN, T_FREE)
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
    b.LIST_RANGE = {}
    ws.sheet_state = "hidden"
    for j, (nm, vals) in enumerate(LISTS.items()):
        col = j + 1
        ws.cell(1, col, nm).font = F_BOLD
        for k, v in enumerate(vals):
            ws.cell(2 + k, col, v)
        b.name(nm, f"Lists!${CL(col)}$2:${CL(col)}${1 + len(vals)}")
        b.LIST_RANGE[nm] = f"Lists!${CL(col)}$2:${CL(col)}${1 + len(vals)}"
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
             ("F", "SAMPLE 1\nB+G+8 mixed" if not blank else "SAMPLE 1\n(cleared)"),
             ("G", "SAMPLE 2\nTrade centre" if not blank else "SAMPLE 2\n(cleared)"),
             ("H", "ACTIVE\n(used by model)"), ("I", "Range name"), ("J", "Notes / source"), ("K", "Confidence"),
             ("L", "Source / date")]
    for col, h in heads:
        c = ws[f"{col}{r}"]
        c.value = h
        c.font = font(9, True, "FFFFFF")
        c.fill = fill(NAVY if col not in ("E",) else BLUE)
        c.alignment = Alignment(horizontal="center" if col in "EFGHK" else "left", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 30
    ws["B7"] = ('Blue cells on amber are inputs. Column H picks the data set chosen on the Start sheet '
                '(1 = My project, 2 = Sample 1, 3 = Sample 2). Sample columns are examples – "Sample value — not market advice". '
                'Columns K–L record how sure you are of each number and where it came from.')
    ws["B7"].font = F_NOTE
    r = 9
    for item in INPUTS:
        if item[0] == "sec":
            c = ws.cell(r, 2, item[1])
            c.font = font(11, True, NAVY)
            ws.cell(r, 3, item[2]).font = font(10, False, SEC, name=AMFONT)
            for col in range(2, 13):
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
            if col in (6, 7) and not blank and v not in (None, "", 0) and fmt != "txt":
                c.comment = Comment(SAMPLE_NOTE, "Rapha")
        ch = f"CHOOSE(DATA_SET,E{r},F{r},G{r})"
        if isinstance(dv, str) and dv.startswith("L_"):
            f = f"IFERROR(MATCH({ch},{dv},0),1)"
            hf = "int"
        elif dv == "yn":
            f = f'IF({ch}="Yes",1,0)'
            hf = "int"
        elif dv == "txt":
            f = f'{ch}&""'
            hf = "gen"
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
        cf_ = ws.cell(r, 11, confidence(key))
        cf_.font = font(9, False, INPUT_FG)
        cf_.fill = FILL_INPUT
        cf_.border = BOX
        cf_.alignment = Alignment(horizontal="center")
        sc_ = ws.cell(r, 12, "")
        sc_.font = font(9, False, INPUT_FG)
        sc_.fill = FILL_INPUT
        sc_.border = BOX
        b.name(key, f"'10_Inputs'!$H${r}")
        b.INPUT_ROWS[key] = r
        # validation
        rng = f"E{r}:G{r}"
        if isinstance(dv, str) and dv.startswith("L_"):
            b.DV.append(("10_Inputs", "list", dv, rng))
        elif dv == "yn":
            b.DV.append(("10_Inputs", "list", "L_YN", rng))
        elif isinstance(dv, tuple):
            b.DV.append(("10_Inputs", dv[0], dv, rng))
        r += 1
    b.CONF_RNG = f"'10_Inputs'!$K$9:$K${r - 1}"
    b.DV.append(("10_Inputs", "list", "L_CONF", f"K9:K{r - 1}"))
    b.name("N_LOWCONF", f"'10_Inputs'!$K${r + 1}")
    ws.cell(r + 1, 10, "Inputs marked Low confidence").font = F_BOLD
    c_ = ws.cell(r + 1, 11, f'=COUNTIF(K9:K{r - 1},"Low")')
    c_.font = F_BOLD
    c_.fill = FILL_KEY
    c_.alignment = Alignment(horizontal="center")
    r += 2
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
            b.DV.append(("10_Inputs", "dec", ("dec", 0, 1e7), f"{CL(5 + jj)}{r}"))
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
        b.DV.append(("10_Inputs", "dec", ("dec", 0.1, 5), f"E{r}"))
        r += 1
    for col, w in (("C", 30), ("D", 10), ("E", 15), ("F", 15), ("G", 15), ("H", 15), ("I", 16), ("J", 44), ("K", 12), ("L", 24)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "E9"
    sh.r = r + 1
    return sh


def add_validations(b):
    """Google-Sheets-safe validations: list sources point at sheet ranges (not names)."""
    for sheet, typ, arg, rng in b.DV:
        ws = b.wb[sheet]
        if typ == "list":
            dv = DataValidation(type="list", formula1=f"={b.LIST_RANGE[arg]}", allow_blank=True)
            dv.error = "Pick a value from the list"
            dv.showErrorMessage = True
        else:
            lo, hi = arg[1], arg[2]
            dv = DataValidation(type="decimal" if typ == "dec" else "whole", operator="between",
                                formula1=str(lo), formula2=str(hi), allow_blank=True)
            dv.error = f"Enter a value between {lo:g} and {hi:g}"
            dv.showErrorMessage = True
        ws.add_data_validation(dv)
        dv.add(rng)


# ---------------------------------------------------------------------- SPACES (rent roll)
def build_spaces(b, blank=False):
    sh = Sheet(b, "13_Spaces", "13 · Space schedule & rent roll", "የቦታ ዝርዝርና የኪራይ መዝገብ", TAB_IN, kind="free", n=0,
               width_label=5)
    ws = sh.ws
    cols = ["#", "Space / unit name", "Type", "Net area m²", "Sold % (Mixed)", "Sale price ETB/m²", "Rent ETB/m²/month",
            "Stabilised occupancy"]
    fmts = [None, None, None, "n0", "pct", "n0", "n0", "pct"]
    blocks = [("MY PROJECT – your spaces (type here)", SPACES_USER),
              ("SAMPLE 1 – B+G+8 mixed-use", SPACES_USER if blank else SPACES_S1),
              ("SAMPLE 2 – B+G+10 trade centre & hotel", SPACES_USER if blank else SPACES_S2)]
    b.SP_ROWS = []
    r = 8
    starts = []
    for title, data in blocks:
        starts.append(r + 2)
        r += 2 + N_SPACES + 1
    act0 = r + 2
    tot = act0 + N_SPACES
    sh.r = tot + 12

    def _sp(_r):
        ws["B6"] = ("Every lettable or saleable space, one row each. The ACTIVE block at the bottom uses the data set chosen on Start "
                    "and the project type: Build-to-Sell sells all saleable space, Build-to-Rent keeps it, Mixed uses the Sold % column. "
                    "Hotel, owner-occupied and rent-free space is never sold. Hotel income comes from Inputs block 7.")
        ws["B6"].font = F_NOTE
        for (title, data), r0 in zip(blocks, starts):
            ws.cell(r0 - 2, 2, title).font = font(11, True, NAVY)
            for col in range(2, 10):
                ws.cell(r0 - 2, col).border = BOTTOM_AMBER
            for j, h in enumerate(cols):
                c = ws.cell(r0 - 1, 2 + j, h)
                c.font = font(9, True, "FFFFFF")
                c.fill = fill(BLUE if r0 == starts[0] else NAVY)
                c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
            ws.row_dimensions[r0 - 1].height = 28
            for k in range(N_SPACES):
                rr = r0 + k
                ws.cell(rr, 2, k + 1).font = F_UNIT
                for j, v in enumerate(data[k]):
                    c = ws.cell(rr, 3 + j, v)
                    c.font = F_INPUT
                    c.fill = FILL_INPUT
                    c.border = BOX
                    if fmts[1 + j]:
                        c.number_format = NF[fmts[1 + j]]
                    if r0 != starts[0] and j >= 2 and v not in (0, None, ""):
                        c.comment = Comment(SAMPLE_NOTE, "Rapha")
            rng = lambda col: f"{col}{r0}:{col}{r0 + N_SPACES - 1}"
            b.DV.append(("13_Spaces", "list", "L_SPACE", rng("D")))
            b.DV.append(("13_Spaces", "dec", ("dec", 0, 1e7), rng("E")))
            b.DV.append(("13_Spaces", "dec", ("dec", 0, 1), rng("F")))
            b.DV.append(("13_Spaces", "dec", ("dec", 0, 1e8), rng("G")))
            b.DV.append(("13_Spaces", "dec", ("dec", 0, 1e7), rng("H")))
            b.DV.append(("13_Spaces", "dec", ("dec", 0, 1), rng("I")))
        # active block
        ws.cell(act0 - 2, 2, "ACTIVE – used by the model (do not type here)").font = font(11, True, NAVY)
        ws.cell(act0 - 2, 11, "ያገለገለ – ሞዴሉ የሚጠቀመው").font = font(10, False, SEC, name=AMFONT)
        for col in range(2, 21):
            ws.cell(act0 - 2, col).border = BOTTOM_AMBER
        heads = cols + ["Type code", "Share sold", "Share owner-held", "Share let", "Sold m²", "Let m²", "Owner-held m²",
                        "Sales value (today)", "Potential rent p.a.", "Rent at stabilised occupancy"]
        for j, h in enumerate(heads):
            c = ws.cell(act0 - 1, 2 + j, h)
            c.font = font(9, True, "FFFFFF")
            c.fill = fill(NAVY)
            c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        ws.row_dimensions[act0 - 1].height = 40
        for k in range(N_SPACES):
            rr = act0 + k
            rm, r1, r2 = (st + k for st in starts)
            ws.cell(rr, 2, k + 1).font = F_UNIT
            for col in "CD":
                ws[f"{col}{rr}"] = f"=CHOOSE(DATA_SET,{col}{rm},{col}{r1},{col}{r2})&\"\""
            for col in "EFGHI":
                ws[f"{col}{rr}"] = f"=CHOOSE(DATA_SET,{col}{rm},{col}{r1},{col}{r2})"
            f = {
                "J": f"=IF(E{rr}>0,IFERROR(MATCH(D{rr},L_SPACE,0),11),0)",
                "K": f"=IF(OR(J{rr}=0,J{rr}={T_HOTEL},J{rr}={T_OWN},J{rr}={T_FREE}),0,CHOOSE(PROJECT_MODE,1,0,MIN(1,MAX(0,F{rr}))))",
                "L": f"=IF(OR(J{rr}={T_OWN},J{rr}={T_FREE}),1,0)",
                "M": f"=IF(J{rr}=0,0,1-K{rr}-L{rr})",
                "N": f"=E{rr}*K{rr}", "O": f"=E{rr}*M{rr}", "P": f"=E{rr}*L{rr}",
                "Q": f"=N{rr}*G{rr}",
                "R": f"=IF(J{rr}={T_HOTEL},0,O{rr}*H{rr}*12)",
                "S": f"=R{rr}*MIN(1,MAX(0,I{rr}))",
            }
            for col, v in f.items():
                ws[f"{col}{rr}"] = v
            for j, col in enumerate("CDEFGHIJKLMNOPQRS"):
                c = ws[f"{col}{rr}"]
                c.font = font(9)
                c.number_format = {"E": NF["n0"], "F": NF["pct"], "G": NF["n0"], "H": NF["n0"], "I": NF["pct"],
                                   "J": NF["int"], "K": NF["pct"], "L": NF["pct"], "M": NF["pct"]}.get(col, NF["n0"])
        ws.cell(tot, 3, "TOTAL").font = F_BOLD
        for col in "ENOPQRS":
            c = ws[f"{col}{tot}"]
            c.value = f"=SUM({col}{act0}:{col}{tot - 1})"
            c.font = F_BOLD
            c.fill = FILL_KEY
            c.number_format = NF["n0"]
            c.border = Border(top=Side(style="thin", color=INK))
        nm = {"E": "NET_AREA", "N": "SALE_AREA", "O": "LET_AREA", "P": "OWN_AREA", "Q": "SALEVAL_TODAY", "R": "GPR_TODAY",
              "S": "GPR_STAB"}
        for col, n_ in nm.items():
            b.name(n_, f"'13_Spaces'!${col}${tot}")
        rngJ = f"$J${act0}:$J${tot - 1}"
        extra = [("Hotel space (net m²)", f"=SUMIF({rngJ},{T_HOTEL},$E${act0}:$E${tot - 1})", "HOTEL_AREA"),
                 ("Residential rent at stabilised occupancy", f"=SUMIF({rngJ},{T_RES},$S${act0}:$S${tot - 1})", "RES_RENT_STAB"),
                 ("Spaces entered", f'=COUNTIF({rngJ},">0")', "N_SPACES_USED")]
        for k, (lab, f, n_) in enumerate(extra):
            rr = tot + 2 + k
            ws.cell(rr, 3, lab).font = F_BODY
            c = ws.cell(rr, 5, f)
            c.number_format = NF["n0"]
            c.font = F_BOLD
            b.name(n_, f"'13_Spaces'!$E${rr}")
        ws.cell(tot + 6, 3, "Mix by type (net m²) – chart source").font = font(9, True, NAVY)
        b.SP_MIX = (tot + 7, len(SPACE_TYPES))
        for k, t in enumerate(SPACE_TYPES):
            rr = tot + 7 + k
            ws.cell(rr, 3, t).font = font(9)
            c = ws.cell(rr, 5, f"=SUMIF({rngJ},{k + 1},$E${act0}:$E${tot - 1})")
            c.number_format = NF["n0"]
            c.font = font(9)
        for col, w in (("B", 5), ("C", 30), ("D", 24), ("E", 11), ("F", 10), ("G", 13), ("H", 12), ("I", 12), ("J", 8),
                       ("K", 9), ("L", 9), ("M", 9), ("N", 10), ("O", 10), ("P", 10), ("Q", 14), ("R", 14), ("S", 14)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "C8"
    sh.raw(_sp)
    sh.r = tot + 7 + len(SPACE_TYPES) + 2
    return sh


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
                                      FormulaRule(formula=[f"COLUMN(E{hdr + 1})-4=$D${sh.scal['SCEN']}"], fill=fill("FFE9C7")))
        for col, w in (("C", 20), ("D", 8), ("E", 11), ("F", 11), ("G", 11), ("H", 11), ("I", 12), ("J", 14)):
            ws.column_dimensions[col].width = w
    sh.raw(_t)
    sh.sec("Effective drivers used by the engine (scenario × flex cell)", "ሞዴሉ የሚጠቀማቸው ውጤታማ ነጂዎች")
    eff = [
        ("SCEN", "Active scenario (1–4)", "SCEN_SEL", "int"),
        ("D_PRICE_F", "Sale price factor", "(1+SC_PRICE)*FLX_PRICE", "n2"),
        ("D_RENT_F", "Rent factor", "(1+SC_RENT)*FLX_RENT", "n2"),
        ("D_COST_F", "Construction cost factor", "(1+SC_COST)*FLX_COST", "n2"),
        ("D_OPEX_F", "Opex factor", "(1+SC_OPEX)*FLX_OPEX", "n2"),
        ("D_OCC_ADD", "Occupancy points added", "SC_OCC", "pct"),
        ("D_OCC_F", "Occupancy factor", "FLX_OCC", "n2"),
        ("D_RATE_ADD", "Interest rate points added", "SC_RATE", "pct2"),
        ("D_RATE_F", "Interest rate factor", "FLX_RATE", "n2"),
        ("RATE_EFF", "Effective base interest rate", "MAX(0,(INT_RATE+D_RATE_ADD)*D_RATE_F)", "pct2"),
        ("CONSTR_EFF", "Effective construction months (incl. delay)", "MAX(1,MIN(48,ROUND(CONSTR_M+SC_DELAY+FLX_DELAY,0)))", "int"),
        ("D_PRESALE_F", "Presale factor", "MAX(0,1+SC_PRESALE)", "n2"),
        ("EXIT_CAP_EFF", "Effective exit cap rate", "MAX(0.001,EXIT_CAP+SC_EXITCAP)", "pct2"),
        ("ESC_EFF", "Effective cost escalation", "ESC_PA+SC_ESC", "pct"),
        ("RESC_EFF", "Effective rent escalation", "RENT_ESC+SC_RESC", "pct"),
        ("PESC_EFF", "Effective sale-price escalation", "PRICE_ESC+SC_PESC", "pct"),
        ("DEFAULT_EFF", "Effective default allowance", "MIN(1,MAX(0,DEFAULT_PCT+SC_DEFAULT))", "pct"),
        ("DISC_EFF", "Effective project discount rate", "DISC_RATE+SC_DISC", "pct"),
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
    ("VAT on rent (share taxable – computed from the rent roll)", "RENT_VAT_PCT", "Proc. 1341/2024 – residential rent exempt; commercial rent taxable. Model: 1 − exempt residential rent ÷ total rent"),
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
COM_SHEETS = [("01_Dashboard", "Dashboard", "ዳሽቦርድ"), ("10_Inputs", "Inputs", "ግብዓቶች"), ("13_Spaces", "Space schedule & rent roll", "የቦታ ዝርዝር"),
              ("11_Scenarios", "Scenarios", "ሁኔታዎች"), ("40_PL", "Profit & loss", "ትርፍና ኪሳራ"), ("41_BS", "Balance sheet", "ሀብትና ዕዳ"),
              ("42_CF", "Cash flow", "የገንዘብ ፍሰት"), ("27_Loan", "Loan schedule", "የብድር ሰንጠረዥ"), ("50_Returns", "Returns", "ተመላሽ"),
              ("52_Sensitivity", "Sensitivity & break-even", "የስጋት ትንተና"), ("60_Checks", "Checks", "ማረጋገጫዎች"),
              ("12_TaxLaw", "Tax law register", "የግብር ሕግ")]
PRI_SHEETS = [("P3_Estimate", "My house – estimate (one page)", "የቤት ግምት"), ("P1_Inputs", "My house – inputs", "ግብዓቶች"),
              ("P2_Calc", "Calculations & loan schedule", "ስሌትና የብድር ሰንጠረዥ"), ("P4_Checks", "Checks", "ማረጋገጫዎች"),
              ("61_Glossary", "Glossary", "መዝገበ ቃላት")]
ALL_SHEETS = ["P1_Inputs", "P2_Calc", "P3_Estimate", "P4_Checks", "01_Dashboard", "10_Inputs", "11_Scenarios", "12_TaxLaw",
              "13_Spaces", "20_Timeline", "21_Area", "22_Capex", "23_Construction", "24_Revenue", "25_Opex", "26_Assets", "27_Loan",
              "28_Tax", "29_VAT", "30_WorkingCap", "40_PL", "41_BS", "42_CF", "43_Ratios", "50_Returns", "51_Payback",
              "52_Sensitivity", "60_Checks", "61_Glossary"]


def build_start(b, blank=False):
    from core import link
    sh = Sheet(b, "00_Start", "Rapha Development Profitability Model", "የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል", TAB_SETUP, kind="free",
               n=0, width_label=3)
    ws = sh.ws

    def dd(r, c1, c2, lst, default):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        c = ws.cell(r, c1, default)
        c.font = font(11, True, INPUT_FG)
        for cc in range(c1, c2 + 1):
            ws.cell(r, cc).fill = FILL_INPUT
            ws.cell(r, cc).border = BOX
        dv = DataValidation(type="list", formula1=f"={b.LIST_RANGE[lst]}", allow_blank=False)
        dv.error = "Pick a value from the list"
        dv.showErrorMessage = True
        ws.add_data_validation(dv)
        dv.add(ws.cell(r, c1).coordinate)
        return ws.cell(r, c1).coordinate

    def status(r, c, lab, f, fmt=None, cv=None):
        ws.cell(r, c, lab).font = F_BODY
        x = ws.cell(r, cv or c + 1, f)
        x.font = font(11, True, NAVY)
        x.alignment = Alignment(horizontal="right")
        if fmt:
            x.number_format = NF.get(fmt, fmt)
        return x

    def _s(_r):
        for col, w in (("A", 2), ("B", 3), ("C", 30), ("D", 17), ("E", 17), ("F", 3), ("G", 3), ("H", 30), ("I", 17),
                       ("J", 17), ("K", 3), ("L", 3), ("M", 3), ("N", 3), ("O", 10), ("P", 14)):
            ws.column_dimensions[col].width = w
        ws["C5"] = "Plan it, cost it, test it  ·  ያቅዱ፣ ዋጋውን ይወቁ፣ ይፈትሹ"
        ws["C5"].font = font(13, True, NAVY)
        ws["C6"] = ("Two tools in one workbook: A – a private-house cost estimate with a home-loan plan; "
                    "B – a full profitability model for a commercial building. Pick one in step 1, then use that card only.")
        ws["C6"].font = F_NOTE
        # ---- step 1
        card(ws, 8, 2, 10, 11, bg=CANVAS)
        ws["C8"] = "STEP 1 · What are you planning?  ·  ምን እያቀዱ ነው?"
        ws["C8"].font = font(11, True, NAVY)
        t = dd(9, 4, 10, "L_TRACK", LISTS["L_TRACK"][1])
        ws["C9"] = "Choose ▸"
        ws["C9"].font = F_BOLD
        ws["C10"] = '=IF($P$9=1,"→ Use card A (left). Card B is switched off.","→ Use card B (right). Card A is switched off.")'
        ws["C10"].font = font(10, True, AMBER)
        # ---- helpers (column P) – every conditional format on this sheet reads only these cells
        ws["O8"] = "helpers"
        ws["O8"].font = font(8, False, SEC)
        helpers = [(9, "TRACK", f"IFERROR(MATCH({t},L_TRACK,0),2)"), (10, "P_DATA", "IFERROR(MATCH(D15,L_PDATASET,0),1)"),
                   (11, "PROJECT_MODE", "IFERROR(MATCH(I15,L_MODE,0),3)"), (12, "DATA_SET", "IFERROR(MATCH(I17,L_DATASET,0),1)"),
                   (13, "SCEN_SEL", "IFERROR(MATCH(I19,L_SCEN,0),1)"),
                   (14, "MODE_NAME", 'CHOOSE(PROJECT_MODE,"Sell","Hold","Mixed")'),
                   (15, "SCEN_NAME", 'CHOOSE(SCEN,"Base","Upside","Downside","Lender")'),
                   (16, "UNIT_LBL", '"ETB \'000"'),
                   (17, None, "VERDICT"), (18, None, "CHK_FAIL"), (19, None, "PCHK_FAIL"), (20, None, "P_GAP"),
                   (21, None, "AND(P_INCOME>0,P_DTI>P_AFF_MAX)")]
        for r, nm_, f in helpers:
            c = ws.cell(r, 16, "=" + f)
            c.font = font(8, False, SEC)
            ws.cell(r, 15, nm_ or "cf helper").font = font(7, False, SEC)
            if nm_:
                b.name(nm_, f"'00_Start'!$P${r}")
        # ---- card A (private)
        card(ws, 12, 2, 37, 5, bg=CARD)
        ws["C12"] = "A · PRIVATE RESIDENCE"
        ws["C12"].font = font(12, True, TEAL)
        ws["C13"] = "Cost estimate, cash per month and home-loan plan  ·  የግል ቤት"
        ws["C13"].font = F_NOTE
        ws["C15"] = "Which data?"
        ws["C15"].font = F_BOLD
        ws["C16"] = "የትኛውን መረጃ?"
        ws["C16"].font = font(10, False, SEC, name=AMFONT)
        dd(15, 4, 5, "L_PDATASET", LISTS["L_PDATASET"][0] if blank else LISTS["L_PDATASET"][1])
        ws["C18"] = "STATUS  ·  ሁኔታ"
        ws["C18"].font = font(10, True, NAVY)
        status(19, 3, "Checks passed", '=PCHK_PASS&" / "&PCHK_TOTAL', cv=5)
        status(20, 3, "Model status", "=PCHK_STATUS", cv=5)
        status(21, 3, "Total cost of the house (ETB)", "=P_TOTAL", "n0", cv=5)
        status(22, 3, "Monthly instalment (ETB)", "=P_INSTAL", "n0", cv=5)
        status(23, 3, "Affordability", '=P_AFF_TXT', cv=5)
        status(24, 3, "Funding gap (ETB)", "=P_GAP", "n0", cv=5)
        ws["C26"] = "GO TO  ·  ወደ"
        ws["C26"].font = font(10, True, NAVY)
        for k, (sname, en, am) in enumerate(PRI_SHEETS):
            link(ws, 27 + k, 3, sname, f"▸ {en}", align="left")
            ws.cell(27 + k, 4, am).font = font(9, False, SEC, name=AMFONT)
        # ---- card B (commercial)
        card(ws, 12, 7, 37, 11, bg=CARD)
        ws["H12"] = "B · COMMERCIAL PROJECT"
        ws["H12"].font = font(12, True, BLUE)
        ws["H13"] = "Profitability, IFRS statements, loan and lender metrics  ·  የንግድ ፕሮጀክት"
        ws["H13"].font = F_NOTE
        for r, en, am, lst, dflt in ((15, "Project type", "የፕሮጀክት ዓይነት", "L_MODE", LISTS["L_MODE"][2]),
                                     (17, "Which data?", "የትኛውን መረጃ?", "L_DATASET", LISTS["L_DATASET"][0 if blank else 1]),
                                     (19, "Scenario", "ሁኔታ", "L_SCEN", LISTS["L_SCEN"][0])):
            ws.cell(r, 8, en).font = F_BOLD
            ws.cell(r + 1, 8, am).font = font(10, False, SEC, name=AMFONT)
            dd(r, 9, 10, lst, dflt)
        ws["H21"] = ('=IF(AND(DATA_SET=2,PROJECT_MODE<>3),"Tip: Sample 1 is built for type 3 (Mixed).",'
                     'IF(AND(DATA_SET=3,PROJECT_MODE<>2),"Tip: Sample 2 is built for type 2 (Hold).",""))')
        ws["H21"].font = font(9, True, AMBER)
        ws["H22"] = "STATUS  ·  ሁኔታ"
        ws["H22"].font = font(10, True, NAVY)
        status(23, 8, "Checks passed", '=CHK_PASS&" / "&CHK_TOTAL', cv=10)
        status(24, 8, "Model status", "=CHK_STATUS", cv=10)
        status(25, 8, "Verdict  ·  ውሳኔ", '=VERDICT&"  ·  "&VERDICT_AM', cv=10)
        status(26, 8, "Total development cost (ETB)", "=TDC", "n0", cv=10)
        status(27, 8, "Equity IRR", "=EIRR", "pct", cv=10)
        status(28, 8, "Minimum DSCR", "=MIN_DSCR", "x", cv=10)
        ws["H30"] = "GO TO  ·  ወደ"
        ws["H30"].font = font(10, True, NAVY)
        for k, (sname, en, am) in enumerate(COM_SHEETS):
            r, c = (31 + k, 8) if k < 6 else (31 + k - 6, 9)
            link(ws, r, c, sname, f"▸ {en}", align="left")
        # inactive card greyed out (rule reads the TRACK helper on this sheet)
        grey = dict(fill=fill("ECEEF1"), font=Font(color="9AA3AF"))
        ws.conditional_formatting.add("B12:E37", FormulaRule(formula=["$P$9<>1"], **grey))
        ws.conditional_formatting.add("G12:K37", FormulaRule(formula=["$P$9<>2"], **grey))
        ws.conditional_formatting.add("J24", FormulaRule(formula=["$P$18>0"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("J24", FormulaRule(formula=["$P$18=0"], fill=fill(GOOD_BG), font=Font(color=GOOD_FG, bold=True)))
        ws.conditional_formatting.add("E20", FormulaRule(formula=["$P$19>0"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("E20", FormulaRule(formula=["$P$19=0"], fill=fill(GOOD_BG), font=Font(color=GOOD_FG, bold=True)))
        for rule_f, bg, fg in (('$P$17="VIABLE"', GOOD_BG, GOOD_FG), ('$P$17="MARGINAL"', WARN_BG, WARN_FG),
                               ('$P$17="NOT VIABLE"', CRIT_BG, CRIT_FG)):
            ws.conditional_formatting.add("J25", FormulaRule(formula=[rule_f], fill=fill(bg), font=Font(color=fg, bold=True)))
        ws.conditional_formatting.add("E24", FormulaRule(formula=["$P$20>0.5"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("E23", FormulaRule(formula=["$P$21"], fill=fill(CRIT_BG), font=Font(color=CRIT_FG, bold=True)))
        # ---- all sheets index
        r0 = 39
        ws.cell(r0, 3, "ALL SHEETS – click to jump  ·  ሁሉም ገጾች").font = font(11, True, NAVY)
        per = 10
        for k, sname in enumerate(ALL_SHEETS):
            col = (3, 4, 8)[k // per] if k // per < 3 else 9
            r = r0 + 1 + (k % per)
            link(ws, r, col, sname, sname.replace("_", " "), font(9, False, BLUE, underline="single"), align="left")
        # ---- how to use
        r1 = r0 + per + 2
        ws.cell(r1, 3, "HOW TO USE  ·  አጠቃቀም").font = font(11, True, NAVY)
        steps = [("1", "Step 1: choose Private residence or Commercial project. Only the matching card is active.",
                  "ደረጃ 1፡ የግል መኖሪያ ቤት ወይም የንግድ ፕሮጀክት ይምረጡ።"),
                 ("2", "Private: pick 'My house', type your values in column E of P1 Inputs, then open the Estimate.",
                  "ለግል ቤት፡ 'የእኔ ቤት'ን ይምረጡ፣ በP1 ዓምድ E ላይ ይጻፉ፣ ከዚያ ግምቱን ይክፈቱ።"),
                 ("3", "Commercial: pick the project type and 'My project', fill 10 Inputs (column E) and 13 Spaces (top block).",
                  "ለንግድ ፕሮጀክት፡ ዓይነቱንና 'የእኔ ፕሮጀክት'ን ይምረጡ፣ ገጽ 10 እና 13ን ይሙሉ።"),
                 ("4", "Read the status box (all checks must pass), then the Dashboard. Stress-test with scenarios and the dials on 52.",
                  "የማረጋገጫ ሳጥኑን ይመልከቱ፣ ከዚያ ዳሽቦርዱን፤ በሁኔታዎችና በገጽ 52 ይፈትሹ።")]
        for k, (nmb, en, am) in enumerate(steps):
            r = r1 + 1 + 2 * k
            ws.cell(r, 2, nmb).font = font(13, True, AMBER)
            ws.cell(r, 3, en).font = F_BODY
            ws.cell(r + 1, 3, am).font = font(10, False, SEC, name=AMFONT)
        r2 = r1 + 10
        ws.cell(r2, 3, "COLOUR CODE  ·  የቀለም ኮድ").font = font(11, True, NAVY)
        leg = [("Input – type here", "ግብዓት – እዚህ ይጻፉ", F_INPUT, FILL_INPUT),
               ("Formula – do not overwrite", "ቀመር – አይጻፉበት", font(10, False, INK), fill("FFFFFF")),
               ("Link from another sheet", "ከሌላ ገጽ ማገናኛ", font(10, False, LINK_FG), fill("FFFFFF")),
               ("Key output / check", "ቁልፍ ውጤት / ማረጋገጫ", font(10, True, INK), FILL_KEY)]
        for k, (en, am, f, fl) in enumerate(leg):
            c = ws.cell(r2 + 1 + k, 3, en)
            c.font = f
            c.fill = fl
            c.border = BOX
            ws.cell(r2 + 1 + k, 4, am).font = font(10, False, SEC, name=AMFONT)
        ws.cell(r2, 8, "GOOGLE SHEETS  ·  ጉግል ሺትስ").font = font(11, True, NAVY)
        gs = ["Upload the .xlsx to Google Drive → right-click → Open with → Google Sheets (or File → Save as Google Sheets).",
              "Every link on this page and every '◂ Start' / 'Dashboard ▸' link in the sheet headers jumps to its sheet.",
              "Dropdowns, checks, charts and the sensitivity sheet all recalculate live – no macros, no Data Tables.",
              "File → Settings → Calculation: leave 'Iterative calculation' OFF (the model has no circular references)."]
        for k, t_ in enumerate(gs):
            ws.merge_cells(start_row=r2 + 1 + k, start_column=8, end_row=r2 + 1 + k, end_column=11)
            c = ws.cell(r2 + 1 + k, 8, t_)
            c.font = font(9, False, INK)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r2 + 1 + k].height = 25
        r3 = r2 + 7
        ws.cell(r3, 3, "METHOD NOTES").font = font(11, True, NAVY)
        notes = [
            "No circularity: interest during construction (IDC) is built month by month on the opening balance plus draws; "
            "the arrangement fee is charged on a facility base that excludes IDC; undistributed-profits tax uses a closed form.",
            "Commercial timeline: construction monthly (up to 48 months) rolled up to C1–C4; operations O1–O10 annually; returns use "
            "dated cash flows (XIRR). Private track: 48 construction months plus a 240-month loan schedule.",
            "IFRS: IAS 2 units for sale, IAS 40 let space (cost model), IAS 16 owner-held, IAS 23 capitalised borrowing cost, "
            "IFRS 15 sales, IFRS 16 rent, IAS 12 tax, IAS 36 impairment test, IAS 7 indirect cash flow.",
            "Disclaimer: a planning tool only – not investment, tax or legal advice. Confirm tax rates (sheet 12) before use.",
            "ማሳሰቢያ፡ ይህ የዕቅድ መሣሪያ ብቻ ነው፤ የኢንቨስትመንት፣ የግብር ወይም የሕግ ምክር አይደለም።",
        ]
        for k, t_ in enumerate(notes):
            ws.merge_cells(start_row=r3 + 1 + k, start_column=3, end_row=r3 + 1 + k, end_column=11)
            c = ws.cell(r3 + 1 + k, 3, t_)
            c.font = F_NOTE if k < 4 else font(9, False, SEC, italic=True, name=AMFONT)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r3 + 1 + k].height = 25
        ws.print_area = f"A1:K{r3 + 7}"
        ws.page_setup.fitToHeight = 1
        ws.freeze_panes = None
    sh.raw(_s)
    sh.r = 80
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
        for c, t in (("B5", "Scenario"), ("E5", "Checks"), ("H5", "Project type"), ("K5", "Verdict · ውሳኔ"), ("Q5", "Equity IRR hurdle")):
            ws[c] = t
            ws[c].font = font(9, True, SEC)
        for c, f in (("C5", "=SCEN_NAME"), ("F5", '=CHK_PASS&" / "&CHK_TOTAL&" pass"'), ("I5", "=MODE_NAME"),
                     ("L5", '=VERDICT&"  ·  "&VERDICT_AM'), ("R5", "=HURDLE")):
            ws[c] = f
            ws[c].font = font(11, True, NAVY)
        ws["R5"].number_format = NF["pct"]
        ws.merge_cells("L5:O5")
        ws["L5"].alignment = Alignment(horizontal="center")
        # local helpers (column U) – Google Sheets conditional formats may only read cells on the same sheet
        ws["U4"] = "helpers"
        for r_, f_ in ((5, "=VERDICT"), (6, "=CHK_FAIL"), (7, "=IF(EIRR<HURDLE,1,0)"), (8, "=DSCR_OK"), (9, "=FLEX_ON")):
            ws.cell(r_, 21, f_).font = font(8, False, SEC)
        ws["U4"].font = font(8, False, SEC)
        for rule_f, bg, fg in (('$U$5="VIABLE"', GOOD_BG, GOOD_FG), ('$U$5="MARGINAL"', WARN_BG, WARN_FG),
                               ('$U$5="NOT VIABLE"', CRIT_BG, CRIT_FG)):
            ws.conditional_formatting.add("L5:O5", FormulaRule(formula=[rule_f], fill=fill(bg), font=Font(color=fg, bold=True)))
        ws.conditional_formatting.add("F5", FormulaRule(formula=["$U$6>0"], font=Font(color=CRIT_FG, bold=True)))
        ws["B6"] = ('=IF(FLEX_ON=1,"⚠ What-if dials on sheet 52 are changed – these are NOT base-case results.",'
                    '"ETB; IRR on dated cash flows (XIRR); verdict = equity IRR vs hurdle, profit on cost and DSCR covenant.")')
        ws.conditional_formatting.add("B6:R6", FormulaRule(formula=["$U$9=1"], fill=fill(WARN_BG), font=Font(color=WARN_FG, bold=True)))
        ws["B6"].font = font(8, False, SEC, italic=True)
        tiles = [
            ("TOTAL DEVELOPMENT COST", "=TDC", "m", '="ETB  ·  "&TEXT(TDC_M2G,"#,##0")&" per m² gross"'),
            ("GROSS DEVELOPMENT VALUE", "=GDV", "m", '="GDV ÷ TDC = "&TEXT(GDV_TDC,"0.00")&"x"'),
            ("PROFIT ON COST", "=POC", "pct", '="margin on GDV "&TEXT(DEV_MARGIN,"0.0%")'),
            ("EQUITY IRR", "=EIRR", "pct", '="hurdle "&TEXT(HURDLE,"0.0%")&IF(EIRR>=HURDLE,"  ✓ above","  ✗ below")'),
            ("PEAK DEBT", "=PEAK_DEBT", "m", '="incl. capitalised IDC "&TEXT(IDC_TOT/1000000,"#,##0.0")&" m"'),
            ("MIN DSCR", "=MIN_DSCR", "x", '="covenant "&TEXT(DSCR_COV,"0.00")&"x"&IF(DSCR_OK=1,"  ✓ met","  ✗ breached")'),
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
        ws.conditional_formatting.add("K10:L10", FormulaRule(formula=["$U$7=1"], font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add("Q10:R10", FormulaRule(formula=["$U$8=0"], font=Font(color=CRIT_FG, bold=True)))
        # chart data block
        d0 = 70
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
        ctitle("N28", "Tornado – change in project value, ETB m (±20%; delay +6 mo)")
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
        tc.series[0].tx = SeriesLabel(v="Adverse")
        tc.series[1].tx = SeriesLabel(v="Favourable")
        tc.series[0].graphicalProperties.solidFill = ROSE
        tc.series[0].graphicalProperties.line.solidFill = ROSE
        tc.series[1].graphicalProperties.solidFill = TEAL
        tc.series[1].graphicalProperties.line.solidFill = TEAL
        tc.grouping = "clustered"
        tc.overlap = 100
        tc.gapWidth = 40
        tc.x_axis.scaling.orientation = "maxMin"
        tc.x_axis.tickLblPos = "low"
        tc.y_axis.numFmt = "#,##0"
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
                ("Capex VAT (refund)", f"=SUM('23_Construction'!$E${cmr['VATC']}:${pc(MON_C)}${cmr['VATC']})"),
                ("Lease paid in build", f"=SUM('23_Construction'!$E${cmr['LSI']}:${pc(MON_C)}${cmr['LSINT']})"),
                ("Pre-opening cost", "=PREOP_TOT")]
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
            '=IF(N_BREACH>0,"✗ DSCR below covenant in "&N_BREACH&" year(s)","✓ DSCR covenant met")',
            '=IF(FUND_GAP>0.5,"✗ Funding gap of "&TEXT(FUND_GAP,"#,##0")&" ETB","✓ Funding fully covered")',
            '=IF(N_LOWCASH>0,"! Cash below minimum in "&N_LOWCASH&" year(s)","✓ Cash above minimum")',
            '=IF(N_IMPAIR>0,"! IAS 36 impairment indicator in "&N_IMPAIR&" year(s)","✓ No impairment indicator")',
            '=IF(AND(FAR>0,FAR_ACT>FAR),"! FAR "&TEXT(FAR_ACT,"0.00")&" exceeds permitted "&TEXT(FAR,"0.00"),"✓ FAR within limit")',
            '=IF(EIRR<HURDLE,"! Equity IRR below hurdle","✓ Equity IRR above hurdle")',
            '=IF(N_LOWCONF>0,"! "&N_LOWCONF&" input(s) marked Low confidence – firm them up (10 Inputs, column K)","✓ No Low-confidence inputs")',
            '="! Tax rates must be verified (sheet 12) · የግብር መጣኔዎችን ያረጋግጡ"',
        ]
        for k, f in enumerate(alerts):
            ws.cell(s0 + 1 + k, 11, f).font = font(9)
        rng = f"K{s0 + 1}:K{s0 + len(alerts)}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="✗"'], font=Font(color=CRIT_FG, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="!"'], font=Font(color=WARN_FG, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(K{s0 + 1},1)="✓"'], font=Font(color=GOOD_FG)))
        # key findings – plain-language commentary that rewrites itself
        f0 = 54
        ws.cell(f0, 2, "KEY FINDINGS – written by the model from the live numbers  ·  ዋና ዋና ግኝቶች").font = font(11, True, NAVY)
        for cc in range(2, 19):
            ws.cell(f0, cc).border = BOTTOM_AMBER
        finds = [
            '=PROJ_NAME&" · type: "&MODE_NAME&" · "&SCEN_NAME&" scenario → verdict: "&VERDICT&" ("&VERDICT_AM&")."',
            '="Returns: equity IRR "&TEXT(EIRR,"0.0%")&" against a "&TEXT(HURDLE,"0.0%")&" hurdle; project IRR "&TEXT(PIRR,"0.0%")&'
            '"; profit on cost "&TEXT(POC,"0.0%")&"; equity multiple "&TEXT(EMOIC,"0.00")&"x."',
            '="Cost & funding: TDC ETB "&TEXT(TDC/1000000,"#,##0.0")&" m ("&TEXT(TDC_M2G,"#,##0")&" per m² gross), funded by equity "&'
            'TEXT(EQ_DRAWN/1000000,"#,##0.0")&" m, bank debt "&TEXT(DEBT_DRAWN/1000000,"#,##0.0")&" m and presales "&TEXT(PS_USED/1000000,"#,##0.0")&" m"&'
            'IF(FUND_GAP>0.5,"; a funding gap of "&TEXT(FUND_GAP/1000000,"#,##0.0")&" m is still open.",".")',
            '=IF(N_DS_YEARS=0,"Lender view: no debt service in the horizon.","Lender view: minimum DSCR "&TEXT(MIN_DSCR,"0.00")&"x against a "&'
            'TEXT(DSCR_COV,"0.00")&"x covenant"&IF(N_BREACH>0," – breached in "&N_BREACH&" year(s), when dividends are locked up."," – met in every repayment year.")&'
            '" Peak debt ETB "&TEXT(PEAK_DEBT/1000000,"#,##0.0")&" m, repaid by "&TEXT(LN_FINAL_DATE,"mmm yyyy")&".")',
            '="Income: stabilised NOI ETB "&TEXT(NOI_STAB/1000000,"#,##0.0")&" m a year → yield on cost "&TEXT(YOC,"0.0%")&" vs exit cap rate "&'
            'TEXT(EXIT_CAP_EFF,"0.0%")&IF(YOC>EXIT_CAP_EFF," (positive development spread)."," (no development spread – value relies on sales).")',
            '="Biggest risk on the tornado: "&TOP_RISK&". "&IF(PNPV<0,"Project NPV is negative.",IF(BE_RENT_NPV>0,"Rents can fall "&'
            'TEXT(1-BE_RENT_NPV,"0%")&" before NPV reaches zero","")&IF(BE_COST>0,"; costs can overrun "&TEXT(BE_COST,"0%")&" before equity IRR drops below the hurdle.","."))',
            '=IF(NCI_PCT>0,"Land partner contributes the plot in kind and holds "&TEXT(NCI_PCT,"0%")&" of equity as a non-controlling interest.",'
            '"Land: "&INDEX(L_LAND,LAND_TYPE)&".")&IF(N_LOWCONF>0," "&N_LOWCONF&" input(s) are marked Low confidence – test them on sheet 52.",'
            '" No input is marked Low confidence.")',
        ]
        for k, f in enumerate(finds):
            r_ = f0 + 1 + k
            ws.cell(r_, 2, "•").font = font(10, True, AMBER)
            ws.cell(r_, 2).alignment = Alignment(horizontal="right", vertical="top")
            ws.merge_cells(start_row=r_, start_column=3, end_row=r_, end_column=18)
            c = ws.cell(r_, 3, f)
            c.font = font(10, False, INK)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r_].height = 27
        ws.print_area = "A1:S62"
        ws.page_setup.fitToHeight = 1
        ws.freeze_panes = None
    sh.raw(_d)
    sh.r = 80
    return sh
