"""Input register: every assumption, with My-project (neutral), Sample A and Sample B values.

Row tuple: (key, label_en, label_am, unit, fmt, user, sampleA, sampleB, note, validation)
validation: list-name (text dropdown -> code), 'yn' (Yes/No -> 1/0), 'use' (text kept),
            ('dec', lo, hi) or ('whole', lo, hi) or None.
"""
import datetime as dt

D0 = dt.datetime(2027, 1, 1)

USES = ["Residential", "Retail", "Office", "Parking", "Other"]
USES_AM = ["መኖሪያ", "የንግድ ሱቅ", "ቢሮ", "መኪና ማቆሚያ", "ሌላ"]
FLOORS = ["Basement 2", "Basement 1", "Ground floor"] + [f"Floor {i}" for i in range(1, 14)]
FLOORS_AM = ["ምድር ቤት 2", "ምድር ቤት 1", "ምድር ወለል"] + [f"{i}ኛ ፎቅ" for i in range(1, 14)]
ELEMENTS = ["Preliminaries", "Substructure", "Superstructure (frame)", "Envelope (walls, windows, roof)",
            "Internal works (partitions, doors)", "Finishes", "MEP (electrical, mechanical, plumbing)",
            "External works", "Specialist systems (lifts, fire, CCTV)"]
ELEMENTS_AM = ["ቅድመ-ሥራዎች", "የመሠረት ሥራ", "ከመሬት በላይ መዋቅር (ፍሬም)", "ግድግዳ፣ መስኮትና ጣሪያ",
               "የውስጥ ሥራዎች (ክፍልፋይ፣ በር)", "የማጠናቀቂያ ሥራዎች", "ኤሌክትሪክ፣ ሜካኒካልና ቧንቧ",
               "የውጭ ሥራዎች", "ልዩ ሲስተሞች (ሊፍት፣ እሳት፣ ካሜራ)"]
EL_KEYS = ["PRELIM", "SUBSTR", "SUPER", "ENVEL", "INTERN", "FINISH", "MEP", "EXTW", "SPEC"]
# library: ETB per m2 of gross floor area, Addis Ababa, budget date (Sample values — not market advice)
RATE_LIB = {
    "PRELIM": (1300, 1600, 2000),
    "SUBSTR": (3500, 3800, 4100),
    "SUPER": (7300, 7800, 8300),
    "ENVEL": (3800, 4800, 6600),
    "INTERN": (1800, 2400, 3400),
    "FINISH": (3300, 5600, 9200),
    "MEP": (3000, 4400, 6600),
    "EXTW": (800, 1200, 1800),
    "SPEC": (0, 400, 1200),
}

LISTS = {
    "L_MODE": ["0 – Personal residence (cost estimate only) · የግል መኖሪያ ቤት (የወጪ ግምት ብቻ)",
               "1 – Build-to-Sell · ለሽያጭ መገንባት",
               "2 – Build-to-Rent / Hold · ለኪራይ መገንባት",
               "3 – Mixed (sell part, rent part) · ቅይጥ (ከፊል ሽያጭ፣ ከፊል ኪራይ)"],
    "L_DATASET": ["1 – My project · የእኔ ፕሮጀክት", "2 – SAMPLE A: G+1 family house · ምሳሌ ሀ",
                  "3 – SAMPLE B: B+G+8 mixed-use · ምሳሌ ለ"],
    "L_SCEN": ["1 – Base · መሠረታዊ", "2 – Upside · ብሩህ", "3 – Downside · አስቸጋሪ", "4 – Lender case · የባንክ"],
    "L_UNITS": ["ETB", "ETB '000", "ETB million"],
    "L_FINISH": ["Basic", "Standard", "Premium"],
    "L_YN": ["Yes", "No"],
    "L_USE": USES,
    "L_METHOD": ["Rate x area", "Lump sum (BOQ)"],
    "L_BTYPE": ["G+0 villa", "G+1", "G+2", "G+3", "Other / commercial"],
    "L_LOC": ["Addis Ababa", "Regional town"],
    "L_LAND": ["Lease (cash)", "Already owned (in-kind by owner)", "Partner in-kind (NCI)"],
    "L_SCURVE": ["S-curve (sigmoid)", "Linear", "User % row"],
    "L_DRAW": ["Equity first", "Pro-rata"],
    "L_RATETYPE": ["Fixed", "Variable path by year"],
    "L_GRACE": ["Interest-only (paid)", "Capitalised"],
    "L_REPAY": ["Annuity (PMT)", "Equal principal", "Balloon %"],
    "L_FREQ": ["Monthly", "Quarterly", "Annual"],
    "L_REV": ["Point in time (handover)", "Over time (% complete)"],
    "L_OPEXM": ["% of rent", "ETB per m2 per year"],
    "L_EXIT": ["Exit cap rate", "EBITDA multiple"],
}

SAMPLE_NOTE = "Sample value — not market advice"


def S(t, am=""):
    return ("sec", t, am)


INPUTS = [
    S("1. Project", "1. ፕሮጀክት"),
    ("PROJ_NAME", "Project name", "የፕሮጀክት ስም", "text", "txt", "My project",
     "SAMPLE A — G+1 family house, Addis Ababa", "SAMPLE B — B+G+8 mixed-use building, Addis Ababa",
     "Shown in every sheet header", None),
    ("LOCATION", "Location (sub-city / town)", "አካባቢ (ክ/ከተማ / ከተማ)", "text", "txt", "", "Addis Ababa – Bole sub-city",
     "Addis Ababa – Kirkos sub-city", "Free text", None),
    ("LOC_TYPE", "Location class for cost factor", "የቦታ ዓይነት (ለወጪ ማስተካከያ)", "list", "gen", "Addis Ababa",
     "Addis Ababa", "Addis Ababa", "Regional towns use the regional location factor", "L_LOC"),
    ("PLOT_AREA", "Plot area", "የቦታ ስፋት", "m²", "n0", 0, 300, 1000, "From title deed / lease contract", ("dec", 0, 1e7)),
    ("FAR", "Max floor-area ratio (FAR) permitted", "የተፈቀደ የወለል ስፋት ጥምርታ", "x", "n2", 0, 1.5, 8.0,
     "From the city master plan / building permit; 0 = not checked", ("dec", 0, 50)),
    ("START_DATE", "Construction start date", "የግንባታ መጀመሪያ ቀን", "date", "date", D0, D0, D0,
     "First month of construction (budget date for prices)", None),
    ("CONSTR_M", "Construction period", "የግንባታ ጊዜ", "months", "int", 12, 14, 30,
     "1–48 months; scenario delay months are added", ("whole", 1, 48)),
    ("FX_USD", "Memo: ETB per USD", "ብር በአንድ ዶላር (ማስታወሻ)", "ETB/USD", "n1", 155, 155, 155,
     "Memo only – the model runs in ETB. Verify current NBE rate", ("dec", 0, 10000)),

    S("2. Home estimate card (Mode 0)", "2. የቤት ወጪ ግምት መረጃ (ዓይነት 0)"),
    ("BLDG_TYPE", "Building type", "የሕንፃው ዓይነት", "list", "gen", "G+1", "G+1", "Other / commercial",
     "Checked against the floors entered in block 3", "L_BTYPE"),
    ("BASEMENT_YN", "Basement?", "ምድር ቤት አለው?", "Y/N", "gen", "No", "No", "Yes", "Checked against block 3", "yn"),
    ("FINISH", "Finish level", "የማጠናቀቂያ ደረጃ", "list", "gen", "Standard", "Standard", "Standard",
     "Selects the column of the cost library (block 4)", "L_FINISH"),
    ("HH_INCOME", "Household monthly income (for affordability)", "የቤተሰብ ወርሃዊ ገቢ", "ETB/month", "n0", 0, 300000, 0,
     "Net income of all earners", ("dec", 0, 1e9)),
    ("AFFORD_MAX", "Max instalment as % of income", "ከገቢ ከፍተኛ የክፍያ ድርሻ", "%", "pct", 0.35, 0.35, 0.35,
     "Typical bank rule of thumb 30–40%", ("dec", 0, 1)),
    ("EST_LOW", "Estimate range – low side", "የግምት ዝቅተኛ ወሰን", "%", "pct", -0.10, -0.10, -0.10,
     "Applied to construction cost for the low estimate", ("dec", -0.9, 0)),
    ("EST_HIGH", "Estimate range – high side", "የግምት ከፍተኛ ወሰን", "%", "pct", 0.15, 0.15, 0.15,
     "Applied to construction cost for the high estimate", ("dec", 0, 2)),
    ("SAVINGS", "Own savings available", "የራስ ቁጠባ", "ETB", "n0", 0, 5000000, 0, "Mode 0 funding, drawn first", ("dec", 0, 1e12)),
    ("FAMILY", "Family / diaspora contributions", "የቤተሰብ / ዳያስፖራ መዋጮ", "ETB", "n0", 0, 2500000, 0,
     "Mode 0 funding, drawn after savings", ("dec", 0, 1e12)),
    ("HOME_LOAN", "Home loan amount (maximum)", "የቤት ብድር መጠን", "ETB", "n0", 0, 7000000, 0,
     "Mode 0: bank home-loan limit. Uses the loan terms in block 9", ("dec", 0, 1e12)),
    ("RENT_HOME_F", "I may rent out part of the house", "የቤቱን ከፊል ለማከራየት እችላለሁ", "Y/N", "gen", "No", "No", "No",
     "Upgrade path inside Mode 0: adds rent, payback and yield", "yn"),
    ("RENT_HOME_AREA", "Area rented out", "የሚከራየው ስፋት", "m²", "n0", 0, 60, 0, "e.g. ground-floor unit or service quarter",
     ("dec", 0, 1e6)),
    ("RENT_HOME_MONTHLY", "Expected rent for that part (today's prices)", "የሚጠበቅ ወርሃዊ ኪራይ", "ETB/month", "n0", 0, 25000, 0,
     "Escalated with rent escalation (block 7)", ("dec", 0, 1e9)),

    S("3. Area programme – gross floor area by floor", "3. የወለል ስፋት በፎቅ"),
]
_A = {"Ground floor": 130, "Floor 1": 120}
_B = {"Basement 1": 900, "Ground floor": 650, "Floor 1": 650, **{f"Floor {i}": 600 for i in range(2, 9)}}
for i, fl in enumerate(FLOORS):
    INPUTS.append((f"FL_A{i+1}", f"{fl} – gross area", f"{FLOORS_AM[i]} – ጠቅላላ ስፋት", "m²", "n0", 0, _A.get(fl, 0),
                   _B.get(fl, 0), "Gross floor area incl. walls & common areas", ("dec", 0, 1e6)))
INPUTS.append(S("3b. Use of each floor", "3ለ. የእያንዳንዱ ፎቅ አገልግሎት"))
_BU = {"Basement 1": "Parking", "Ground floor": "Retail", "Floor 1": "Retail", "Basement 2": "Parking"}
for i, fl in enumerate(FLOORS):
    INPUTS.append((f"FL_U{i+1}", f"{fl} – use", f"{FLOORS_AM[i]} – አገልግሎት", "list", "gen", "Residential", "Residential",
                   _BU.get(fl, "Residential"), "Residential / Retail / Office / Parking / Other", "use"))
INPUTS.append(S("3c. Efficiency and disposition by use", "3ሐ. ቅልጥፍናና አጠቃቀም በአገልግሎት"))
_eff = {"Residential": (0.80, 0.85, 0.80), "Retail": (0.85, 0.85, 0.85), "Office": (0.82, 0.82, 0.82),
        "Parking": (0.0, 0.0, 0.0), "Other": (0.0, 0.0, 0.0)}
for u, am in zip(USES, USES_AM):
    e = _eff[u]
    INPUTS.append((f"EFF_{u[:3].upper()}", f"Net-to-gross efficiency – {u}", f"የተጣራ/ጠቅላላ ጥምርታ – {am}", "%", "pct",
                   e[0], e[1], e[2], "Saleable or lettable m² ÷ gross m²", ("dec", 0, 1)))
_sold = {"Residential": (1.0, 0, 1.0), "Retail": (0.0, 0, 0.0), "Office": (0.0, 0, 0.0), "Parking": (0, 0, 0), "Other": (0, 0, 0)}
for u, am in zip(USES, USES_AM):
    s = _sold[u]
    INPUTS.append((f"SOLD_{u[:3].upper()}", f"Mode 3 (Mixed): % of {u} area sold", f"ቅይጥ፡ የሚሸጠው የ{am} ድርሻ", "%", "pct",
                   s[0], s[1], s[2], "Rest is retained and let (IAS 40). Modes 1/2 override this", ("dec", 0, 1)))
INPUTS += [
    ("OWNUSE_PCT", "Commercial modes: % of retained area owner-occupied", "ለባለቤቱ አገልግሎት የሚውል ድርሻ", "%", "pct", 0, 0, 0,
     "Owner-occupied = IAS 16; let = IAS 40; sold = IAS 2", ("dec", 0, 1)),

    S("4. Construction cost", "4. የግንባታ ወጪ"),
]
for k, en, am in zip(EL_KEYS, ELEMENTS, ELEMENTS_AM):
    INPUTS.append((f"METH_{k}", f"{en} – entry method", f"{am} – አገባብ", "list", "gen", "Rate x area", "Rate x area",
                   "Lump sum (BOQ)" if k == "SPEC" else "Rate x area",
                   "Rate x area uses the cost library below; Lump sum = detailed BOQ override", "L_METHOD"))
for k, en, am in zip(EL_KEYS, ELEMENTS, ELEMENTS_AM):
    INPUTS.append((f"LUMP_{k}", f"{en} – lump sum (if BOQ)", f"{am} – ጠቅላላ ዋጋ (BOQ)", "ETB", "n0", 0, 0,
                   45000000 if k == "SPEC" else 0, "Used only when method = Lump sum (BOQ)", ("dec", 0, 1e13)))
INPUTS += [
    ("STOREY_UPLIFT", "Superstructure uplift per storey above G+3", "ከG+3 በላይ ለእያንዳንዱ ፎቅ ጭማሪ", "%", "pct", 0.02, 0.02, 0.02,
     "Taller frames cost more per m²", ("dec", 0, 0.2)),
    ("BASEMENT_RATE", "Basement extra (excavation, retaining, waterproofing)", "የምድር ቤት ተጨማሪ ወጪ", "ETB/m²", "n0", 12000, 12000,
     12000, "Applied to basement gross area", ("dec", 0, 1e6)),
    ("COMPOUND_YN", "Compound wall?", "የአጥር ግንብ?", "Y/N", "gen", "No", "Yes", "Yes", "", "yn"),
    ("COMPOUND_LEN", "Compound wall length", "የአጥር ርዝመት", "m", "n0", 0, 70, 120, "", ("dec", 0, 1e5)),
    ("COMPOUND_RATE", "Compound wall rate", "የአጥር ዋጋ በሜትር", "ETB/m", "n0", 18000, 18000, 20000, "Incl. gate", ("dec", 0, 1e7)),
    ("SEPTIC_YN", "Septic tank?", "የፍሳሽ ጉድጓድ?", "Y/N", "gen", "No", "Yes", "No", "Not needed where sewer exists", "yn"),
    ("SEPTIC_QTY", "Septic tanks – number", "የፍሳሽ ጉድጓድ ብዛት", "no.", "n0", 0, 1, 0, "", ("dec", 0, 100)),
    ("SEPTIC_RATE", "Septic tank – cost each", "የአንዱ ዋጋ", "ETB", "n0", 450000, 450000, 450000, "", ("dec", 0, 1e9)),
    ("WTANK_YN", "Water tanks?", "የውሃ ታንከር?", "Y/N", "gen", "No", "Yes", "Yes", "", "yn"),
    ("WTANK_QTY", "Water tanks – number", "የውሃ ታንከር ብዛት", "no.", "n0", 0, 2, 4, "", ("dec", 0, 100)),
    ("WTANK_RATE", "Water tank – cost each (incl. stand, pump)", "የአንዱ ዋጋ", "ETB", "n0", 60000, 60000, 250000, "",
     ("dec", 0, 1e9)),
    ("EXT_M2", "Paving / landscaping area", "የንጣፍ / የአትክልት ስፋት", "m²", "n0", 0, 80, 250, "", ("dec", 0, 1e6)),
    ("EXT_RATE", "Paving / landscaping rate", "የንጣፍ ዋጋ በካሬ", "ETB/m²", "n0", 3500, 3500, 4500, "", ("dec", 0, 1e6)),
    ("CONT_PCT", "Contingency (% of escalated hard cost)", "ያልተጠበቀ ወጪ መጠባበቂያ", "%", "pct", 0.10, 0.10, 0.08, "",
     ("dec", 0, 1)),
    ("DESIGN_PCT", "Design fee (% of hard cost)", "የዲዛይን ክፍያ", "%", "pct", 0.04, 0.04, 0.05, "", ("dec", 0, 1)),
    ("DESIGN_UPFRONT", "Design fee paid upfront (month 1)", "በመጀመሪያ ወር የሚከፈል ድርሻ", "%", "pct", 0.5, 0.5, 0.5,
     "Rest follows the S-curve", ("dec", 0, 1)),
    ("SUPERV_PCT", "Supervision (% of escalated hard cost)", "የቁጥጥር ክፍያ", "%", "pct", 0.02, 0.02, 0.03, "", ("dec", 0, 1)),
    ("PERMITS", "Permits & municipal fees", "ፈቃድና የማዘጋጃ ቤት ክፍያዎች", "ETB", "n0", 0, 120000, 2500000, "Paid month 1",
     ("dec", 0, 1e12)),
    ("FFE", "Furniture, fittings & equipment (FF&E)", "የቤት ዕቃዎችና መገልገያዎች", "ETB", "n0", 0, 1000000, 6000000,
     "Paid in the completion month", ("dec", 0, 1e12)),
    ("ESC_PA", "Construction cost escalation", "የግንባታ ዋጋ ጭማሪ በዓመት", "% p.a.", "pct", 0.15, 0.15, 0.15,
     "Applied per month from the budget date", ("dec", -0.5, 2)),
    ("INDEX_USE", "Use Rapha Construction Cost Index row instead", "የራፋ የግንባታ ዋጋ ኢንዴክስ ይጠቀሙ", "Y/N", "gen", "No", "No",
     "No", "Enter index values in the row on sheet 23", "yn"),
    ("INDEX_BASE", "Cost index at budget date", "ኢንዴክስ በበጀት ቀን", "index", "n1", 0, 0, 0, "Rapha CCI value at START_DATE",
     ("dec", 0, 1e6)),
    ("SCURVE_TYPE", "Spend profile", "የወጪ ስርጭት", "list", "gen", "S-curve (sigmoid)", "S-curve (sigmoid)", "S-curve (sigmoid)",
     "", "L_SCURVE"),
    ("SCURVE_A", "S-curve steepness", "የS-ኩርባ ቅርጽ", "x", "n1", 2.0, 2.0, 2.0, "1 = linear, 2–3 = typical", ("dec", 0.5, 6)),
    ("RETENTION", "Contractor retention", "ከኮንትራክተር የሚያዝ ገንዘብ", "%", "pct", 0.05, 0.05, 0.05, "Withheld from hard cost",
     ("dec", 0, 0.5)),
    ("RET_REL_M", "Retention released – months after completion", "የሚለቀቅበት ጊዜ", "months", "int", 12, 6, 12, "",
     ("whole", 0, 60)),
    ("VAT_CAPEX_PCT", "Share of capex carrying input VAT", "ተ.እ.ታ ያለበት የወጪ ድርሻ", "%", "pct", 0.7, 0.7, 0.7,
     "Commercial modes only. Mode 0: rates are VAT-inclusive (households cannot reclaim VAT)", ("dec", 0, 1)),

    S("5. Land", "5. መሬት"),
    ("LAND_TYPE", "Land arrangement", "የመሬት ይዞታ ዓይነት", "list", "gen", "Lease (cash)", "Already owned (in-kind by owner)",
     "Lease (cash)", "In-kind partner = non-controlling interest (NCI)", "L_LAND"),
    ("LAND_VALUE", "Land lease price / market value", "የሊዝ ዋጋ / የገበያ ዋጋ", "ETB", "n0", 0, 0, 150000000,
     "Sample A plot already owned – excluded from cost", ("dec", 0, 1e13)),
    ("LEASE_DP", "Lease down payment", "የሊዝ ቅድመ ክፍያ", "%", "pct", 0.10, 0.10, 0.10, "Paid in month 1", ("dec", 0, 1)),
    ("LEASE_N", "Lease instalment years", "የሊዝ ክፍያ ዓመታት", "years", "int", 30, 30, 30, "Annual instalments",
     ("whole", 1, 99)),
    ("LEASE_INT", "Interest on unpaid lease balance", "ባልተከፈለ ሊዝ ላይ ወለድ", "% p.a.", "pct", 0.07, 0.07, 0.07,
     "Per lease contract; verify", ("dec", 0, 1)),

    S("6. Sales", "6. ሽያጭ"),
]
_price = {"Residential": (0, 0, 125000), "Retail": (0, 0, 180000), "Office": (0, 0, 110000), "Parking": (0, 0, 0), "Other": (0, 0, 0)}
for u, am in zip(USES, USES_AM):
    p = _price[u]
    INPUTS.append((f"PRICE_{u[:3].upper()}", f"Sale price – {u} (today's prices)", f"የሽያጭ ዋጋ – {am}", "ETB/m²", "n0",
                   p[0], p[1], p[2], "Per net saleable m²", ("dec", 0, 1e9)))
INPUTS += [
    ("PRICE_ESC", "Sale price escalation", "የሽያጭ ዋጋ ጭማሪ", "% p.a.", "pct", 0.10, 0.10, 0.10, "", ("dec", -0.5, 2)),
    ("PRESALE_1", "Presold in construction year 1 (% of saleable)", "በ1ኛው የግንባታ ዓመት ቅድመ-ሽያጭ", "%", "pct", 0, 0, 0.25, "",
     ("dec", 0, 1)),
    ("PRESALE_2", "Presold in construction year 2", "በ2ኛው ዓመት", "%", "pct", 0, 0, 0.25, "", ("dec", 0, 1)),
    ("PRESALE_3", "Presold in construction year 3", "በ3ኛው ዓመት", "%", "pct", 0, 0, 0.20, "", ("dec", 0, 1)),
    ("PRESALE_4", "Presold in construction year 4", "በ4ኛው ዓመት", "%", "pct", 0, 0, 0, "Total presales capped at 100%",
     ("dec", 0, 1)),
    ("ABS_1", "After completion: % of unsold sold in O1", "ከግንባታ በኋላ በ1ኛው ዓመት የሚሸጥ", "%", "pct", 0.6, 0.6, 0.6, "",
     ("dec", 0, 1)),
    ("ABS_2", "… sold in O2", "በ2ኛው ዓመት", "%", "pct", 0.4, 0.4, 0.4, "", ("dec", 0, 1)),
    ("ABS_3", "… sold in O3", "በ3ኛው ዓመት", "%", "pct", 0, 0, 0, "", ("dec", 0, 1)),
    ("ABS_4", "… sold in O4 (any remainder sells in O5)", "በ4ኛው ዓመት (ቀሪው በ5ኛው)", "%", "pct", 0, 0, 0,
     "Guarantees full sell-out by O5", ("dec", 0, 1)),
    ("PAY_DP", "Payment plan: down payment at contract", "ውል ሲፈረም ቅድመ ክፍያ", "%", "pct", 0.3, 0.3, 0.3, "", ("dec", 0, 1)),
    ("PAY_INST", "Payment plan: instalments tied to construction progress", "በግንባታ ሂደት የሚከፈል", "%", "pct", 0.4, 0.4, 0.4,
     "Balance is paid at handover", ("dec", 0, 1)),
    ("COMM_PCT", "Sales commission", "የሽያጭ ኮሚሽን", "%", "pct", 0.02, 0.02, 0.02, "Expensed when revenue is recognised",
     ("dec", 0, 0.2)),
    ("DEFAULT_PCT", "Default / cancellation allowance", "የውል ማፍረስ ቅናሽ", "%", "pct", 0.02, 0.02, 0.02,
     "Deducted from contract value", ("dec", 0, 0.5)),
    ("MKT_PCT", "Marketing (% of contract value signed)", "ማስታወቂያ", "%", "pct", 0.01, 0.01, 0.01, "", ("dec", 0, 0.2)),
    ("REV_POLICY", "IFRS 15 revenue recognition for sales", "የገቢ ዕውቅና ፖሊሲ", "list", "gen", "Point in time (handover)",
     "Point in time (handover)", "Point in time (handover)", "Contract liability holds cash received before revenue",
     "L_REV"),
    ("VAT_SALES_PCT", "Share of sales value subject to VAT", "ተ.እ.ታ የሚከፈልበት የሽያጭ ድርሻ", "%", "pct", 0, 0, 0,
     "Treatment of building sales must be verified (VAT Proc. 1341/2024)", ("dec", 0, 1)),

    S("7. Rental", "7. ኪራይ"),
]
_rent = {"Residential": (0, 0, 650), "Retail": (0, 0, 2000), "Office": (0, 0, 900), "Parking": (0, 0, 0), "Other": (0, 0, 0)}
for u, am in zip(USES, USES_AM):
    p = _rent[u]
    INPUTS.append((f"RENT_{u[:3].upper()}", f"Rent – {u} (today's prices)", f"ኪራይ – {am}", "ETB/m²/mo", "n0",
                   p[0], p[1], p[2], "Per net lettable m² per month", ("dec", 0, 1e7)))
INPUTS += [
    ("OCC_O1", "Occupancy O1", "የተያዘ መጠን በ1ኛው ዓመት", "%", "pct", 0.6, 0.9, 0.6, "", ("dec", 0, 1)),
    ("OCC_O2", "Occupancy O2", "በ2ኛው ዓመት", "%", "pct", 0.8, 0.95, 0.8, "", ("dec", 0, 1)),
    ("OCC_STAB", "Stabilised occupancy (O3+)", "የተረጋጋ የተያዘ መጠን", "%", "pct", 0.92, 0.95, 0.92, "", ("dec", 0, 1)),
    ("RENT_ESC", "Rent escalation", "የኪራይ ጭማሪ", "% p.a.", "pct", 0.08, 0.08, 0.08, "From today's prices", ("dec", -0.5, 2)),
    ("KEY_MONTHS", "Key money (months of rent, received O1)", "የቁልፍ ገንዘብ (በወራት ኪራይ)", "months", "n1", 0, 0, 6,
     "Deferred and released over the period below (IFRS 16)", ("dec", 0, 60)),
    ("KEY_AMORT_Y", "Key money release period", "የቁልፍ ገንዘብ የሚታወቅበት ጊዜ", "years", "int", 5, 5, 5, "", ("whole", 1, 10)),
    ("SC_RECOV", "Service charge recovery (% of property opex)", "የአገልግሎት ክፍያ ማስመለሻ", "%", "pct", 0.6, 0, 0.6, "",
     ("dec", 0, 1)),
    ("BAD_DEBT", "Bad debt (% of rent)", "የማይሰበሰብ ዕዳ", "%", "pct", 0.02, 0.02, 0.02, "", ("dec", 0, 1)),
    ("OTHER_INC", "Other income (signage, telecom) – year 1", "ሌላ ገቢ", "ETB/yr", "n0", 0, 0, 600000,
     "Escalates with opex escalation", ("dec", 0, 1e12)),
    ("RENT_VAT_PCT", "Share of rent subject to VAT", "ተ.እ.ታ የሚከፈልበት የኪራይ ድርሻ", "%", "pct", 1, 0, 1,
     "Commercial rent is VAT-able; residential rent usually exempt – verify", ("dec", 0, 1)),

    S("8. Operating costs", "8. የሥራ ማስኬጃ ወጪ"),
    ("OPEX_METHOD", "Property opex basis", "የወጪ አሰላል", "list", "gen", "% of rent", "% of rent", "% of rent", "", "L_OPEXM"),
    ("OPEX_PCT", "Property opex (% of rent)", "የንብረት ወጪ (ከኪራይ %)", "%", "pct", 0.15, 0.15, 0.15, "", ("dec", 0, 1)),
    ("OPEX_M2", "Property opex (ETB/m²/yr, today's prices)", "የንብረት ወጪ በካሬ", "ETB/m²/yr", "n0", 1200, 1200, 1200, "",
     ("dec", 0, 1e6)),
    ("MGMT_PCT", "Management fee (% of rent)", "የአስተዳደር ክፍያ", "%", "pct", 0.03, 0, 0.03, "", ("dec", 0, 1)),
    ("INS_PCT", "Insurance (% of building value p.a.)", "መድን", "%", "pct2", 0.0015, 0.0015, 0.0015, "", ("dec", 0, 0.1)),
    ("PROPTAX_PCT", "Property tax (% of land + building value p.a.)", "የንብረት ግብር", "%", "pct2", 0.001, 0.001, 0.001,
     "Verify current city proclamation", ("dec", 0, 0.1)),
    ("MAINT_PCT", "Maintenance reserve (% of building value p.a.)", "የጥገና መጠባበቂያ", "%", "pct2", 0.005, 0.005, 0.005, "",
     ("dec", 0, 0.2)),
    ("OVERHEAD", "Owner overheads (year 1, today's prices)", "አጠቃላይ የአስተዳደር ወጪ", "ETB/yr", "n0", 0, 0, 1200000, "",
     ("dec", 0, 1e12)),
    ("OPEX_ESC", "Opex escalation", "የወጪ ጭማሪ", "% p.a.", "pct", 0.10, 0.10, 0.10, "", ("dec", -0.5, 2)),
    ("OPEX_VAT_PCT", "Share of opex carrying input VAT", "ተ.እ.ታ ያለበት የወጪ ድርሻ", "%", "pct", 0.5, 0.5, 0.5, "",
     ("dec", 0, 1)),
    ("RECV_DAYS", "Rent receivable days", "የኪራይ ተሰብሳቢ ቀናት", "days", "int", 30, 30, 30, "", ("whole", 0, 365)),
    ("PAY_DAYS", "Opex payable days", "የወጪ ተከፋይ ቀናት", "days", "int", 30, 30, 30, "", ("whole", 0, 365)),

    S("9. Financing", "9. ፋይናንስ"),
    ("EQ_PCT", "Commercial: equity share of funding need", "የባለቤት ካፒታል ድርሻ", "%", "pct", 0.4, 0.4, 0.4,
     "Debt = 1 − equity, capped by the loan cap", ("dec", 0, 1)),
    ("DRAW_ORDER", "Draw order", "የአወጣጥ ቅደም ተከተል", "list", "gen", "Equity first", "Equity first", "Equity first", "",
     "L_DRAW"),
    ("PRESALE_USE", "Use presale cash to fund construction?", "የቅድመ-ሽያጭ ገንዘብ ለግንባታ ይውላል?", "Y/N", "gen", "Yes", "Yes",
     "Yes", "", "yn"),
    ("PRESALE_USE_PCT", "Share of presale cash usable (rest in escrow)", "ጥቅም ላይ የሚውለው ድርሻ", "%", "pct", 0.8, 0.8, 0.8, "",
     ("dec", 0, 1)),
    ("LOAN_CAP", "Commercial: loan cap (0 = no cap)", "የብድር ጣሪያ", "ETB", "n0", 0, 0, 0, "", ("dec", 0, 1e13)),
    ("RATE_TYPE", "Interest rate type", "የወለድ ዓይነት", "list", "gen", "Fixed", "Fixed", "Fixed", "", "L_RATETYPE"),
    ("RATE", "Interest rate (fixed, and during construction)", "የወለድ መጠን", "% p.a.", "pct", 0.15, 0.14, 0.15,
     "Verify bank term sheet", ("dec", 0, 1)),
]
for y in range(1, 11):
    INPUTS.append((f"RATE_O{y}", f"Variable path – rate in O{y}", f"ተለዋዋጭ ወለድ O{y}", "% p.a.", "pct", 0.15, 0.14, 0.15,
                   "Used when type = Variable", ("dec", 0, 1)))
INPUTS += [
    ("GRACE_M", "Grace period after completion", "ከግንባታ በኋላ የእፎይታ ጊዜ", "months", "int", 0, 0, 12,
     "Interest during construction is always capitalised (IDC)", ("whole", 0, 120)),
    ("GRACE_TYPE", "Interest during grace", "በእፎይታ ጊዜ ወለድ", "list", "gen", "Interest-only (paid)", "Interest-only (paid)",
     "Interest-only (paid)", "", "L_GRACE"),
    ("TENOR_Y", "Tenor after completion (incl. grace)", "የብድር ጊዜ (እፎይታን ጨምሮ)", "years", "int", 10, 15, 10, "Max 20",
     ("whole", 1, 20)),
    ("REPAY_TYPE", "Repayment type", "የአከፋፈል ዓይነት", "list", "gen", "Annuity (PMT)", "Annuity (PMT)", "Annuity (PMT)", "",
     "L_REPAY"),
    ("BALLOON_PCT", "Balloon (% of balance at repayment start)", "የመጨረሻ ትልቅ ክፍያ", "%", "pct", 0, 0, 0,
     "Only for repayment type = Balloon", ("dec", 0, 1)),
    ("FREQ", "Repayment frequency", "የክፍያ ድግግሞሽ", "list", "gen", "Monthly", "Monthly", "Quarterly", "", "L_FREQ"),
    ("ARR_FEE", "Arrangement fee (% of facility, excl. IDC)", "የብድር ማመቻቻ ክፍያ", "%", "pct", 0.015, 0.01, 0.015,
     "Base excludes IDC → no circularity", ("dec", 0, 0.2)),
    ("COMMIT_FEE", "Commitment fee on undrawn (% p.a.)", "ባልወጣ ብድር ላይ ክፍያ", "%", "pct", 0.005, 0, 0.005, "", ("dec", 0, 0.2)),
    ("DSRA_M", "Debt service reserve (months of next year's debt service)", "የዕዳ ክፍያ መጠባበቂያ", "months", "int", 0, 0, 6, "",
     ("whole", 0, 24)),
    ("SWEEP", "Cash sweep (% of sales proceeds to prepay loan)", "ከሽያጭ ለብድር ቅድመ ክፍያ", "%", "pct", 0, 0, 0.40, "",
     ("dec", 0, 1)),
    ("DSCR_COV", "DSCR covenant", "የዕዳ ክፍያ ሽፋን ግዴታ", "x", "x", 1.3, 1.3, 1.3, "Dividends locked up below covenant",
     ("dec", 0, 10)),
    ("DIST_SURPLUS", "Return surplus cash to shareholders (capital reduction)", "ትርፍ ገንዘብን ለባለአክሲዮኖች መመለስ", "Y/N", "gen",
     "Yes", "No", "Yes", "Subject to the DSCR lock-up test: first returns paid-in capital, then pays a special dividend from retained earnings", "yn"),
    ("MIN_CASH", "Minimum cash balance (operations)", "ዝቅተኛ ጥሬ ገንዘብ", "ETB", "n0", 0, 0, 2000000, "", ("dec", 0, 1e12)),

    S("10. Tax & accounting", "10. ግብርና ሂሳብ አያያዝ"),
    ("CIT", "Corporate income tax", "የንግድ ሥራ ገቢ ግብር", "%", "pct", 0.30, 0.30, 0.30, "See TaxLaw register", ("dec", 0, 1)),
    ("TAXDEP_BLD", "Tax depreciation – buildings (straight line)", "የሕንፃ የግብር እርጅና ቅናሽ", "% p.a.", "pct", 0.05, 0.05, 0.05,
     "", ("dec", 0, 1)),
    ("TAXDEP_PLANT", "Tax depreciation – plant & equipment pool (diminishing)", "የመሣሪያ የግብር እርጅና ቅናሽ", "% p.a.", "pct",
     0.20, 0.20, 0.20, "", ("dec", 0, 1)),
    ("LCF_YEARS", "Tax loss carry-forward period", "ኪሳራ የሚሸጋገርበት ጊዜ", "years", "int", 5, 5, 5, "", ("whole", 0, 20)),
    ("LCF_CAP", "Loss offset cap (% of taxable profit)", "የኪሳራ ማካካሻ ጣሪያ", "%", "pct", 1, 1, 1, "", ("dec", 0, 1)),
    ("VAT_R", "VAT rate", "የተጨማሪ እሴት ታክስ መጣኔ", "%", "pct", 0.15, 0.15, 0.15, "", ("dec", 0, 1)),
    ("VAT_REFUND_F", "Excess input VAT refunded next year?", "ትርፍ የግብዓት ታክስ ይመለሳል?", "Y/N", "gen", "Yes", "Yes", "Yes",
     "No = carried forward against output VAT", "yn"),
    ("UPT_R", "Undistributed profits tax", "ያልተከፋፈለ ትርፍ ግብር", "%", "pct", 0.10, 0.10, 0.10, "Closed-form, no circularity",
     ("dec", 0, 1)),
    ("DIV_PAYOUT", "Dividend payout ratio", "የትርፍ ድርሻ ክፍፍል መጠን", "%", "pct", 0.5, 0, 0.5, "", ("dec", 0, 1)),
    ("DIV_WHT", "Dividend withholding tax (memo)", "የትርፍ ድርሻ ተቀናሽ ግብር", "%", "pct", 0.10, 0.10, 0.10, "", ("dec", 0, 1)),
    ("LR_PCT", "Legal reserve transfer (% of profit)", "ሕጋዊ መጠባበቂያ", "%", "pct", 0.05, 0.05, 0.05, "", ("dec", 0, 1)),
    ("LR_CAP", "Legal reserve cap (% of share capital)", "የሕጋዊ መጠባበቂያ ጣሪያ", "%", "pct", 0.10, 0.10, 0.10, "",
     ("dec", 0, 1)),
    ("LIFE_BLD", "Useful life – building structure", "የሕንፃ አገልግሎት ዘመን", "years", "int", 50, 50, 50, "", ("whole", 1, 100)),
    ("LIFE_SYS", "Useful life – building systems (MEP, lifts)", "የሲስተሞች አገልግሎት ዘመን", "years", "int", 20, 20, 20, "",
     ("whole", 1, 100)),
    ("LIFE_FFE", "Useful life – FF&E", "የዕቃዎች አገልግሎት ዘመን", "years", "int", 7, 7, 7, "", ("whole", 1, 50)),
    ("FFE_LIFE_CAP", "FF&E life cap", "የዕቃዎች ዘመን ጣሪያ", "years", "int", 10, 10, 10, "Fix from reference model",
     ("whole", 1, 50)),

    S("11. Valuation", "11. ግምገማ"),
    ("DISC", "Project discount rate", "የፕሮጀክት የቅናሽ መጠን", "%", "pct", 0.18, 0.18, 0.18, "", ("dec", 0, 1)),
    ("COE", "Cost of equity", "የባለቤት ካፒታል ወጪ", "%", "pct", 0.22, 0.22, 0.22, "", ("dec", 0, 1)),
    ("EXIT_METHOD", "Terminal value method", "የመጨረሻ ዋጋ ዘዴ", "list", "gen", "Exit cap rate", "Exit cap rate", "Exit cap rate",
     "", "L_EXIT"),
    ("EXIT_CAP", "Exit cap rate", "የመውጫ ካፕ ሬት", "%", "pct", 0.09, 0.09, 0.09, "On next-year NOI", ("dec", 0.001, 1)),
    ("EXIT_MULT", "Exit EBITDA multiple", "የመውጫ ብዜት", "x", "x", 10, 10, 10, "", ("dec", 0, 100)),
    ("EXIT_COST", "Exit / disposal costs", "የሽያጭ ወጪ", "%", "pct", 0.02, 0.02, 0.02, "", ("dec", 0, 0.5)),
    ("HURDLE", "Equity IRR hurdle", "ዝቅተኛ ተፈላጊ ተመላሽ", "%", "pct", 0.20, 0.20, 0.20, "", ("dec", 0, 2)),
    ("MARGIN_BAND", "Marginal band below hurdle", "የድንበር ክልል", "%", "pct", 0.03, 0.03, 0.03,
     "Within this band below the hurdle = MARGINAL", ("dec", 0, 1)),
]
