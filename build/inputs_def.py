"""Commercial input register: every assumption, with My-project (neutral), Sample 1 and Sample 2 values.

Row tuple: (key, label_en, label_am, unit, fmt, user, sample1, sample2, note, validation)
validation: list-name (text dropdown -> code), 'yn' (Yes/No -> 1/0), 'txt' (free text),
            ('dec', lo, hi) or ('whole', lo, hi) or None.

Sample 1 = B+G+8 mixed-use building (apartments for sale, shops held).
Sample 2 = B+G+10 trade centre with a hotel operator lease, land contributed in kind by a city partner (NCI).
"""
import datetime as dt

D0 = dt.datetime(2027, 1, 1)

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

# space types for the rent roll (13_Spaces)
SPACE_TYPES = ["Retail", "Office", "Residential", "Showroom / wholesale", "Food & beverage", "Warehouse / storage",
               "Hotel (operator lease)", "Parking", "Owner-occupied", "Rent-free / public service", "Other"]
SPACE_TYPES_AM = ["የንግድ ሱቅ", "ቢሮ", "መኖሪያ", "ማሳያ / ጅምላ ንግድ", "ምግብና መጠጥ", "መጋዘን", "ሆቴል (በኦፕሬተር ኪራይ)",
                  "መኪና ማቆሚያ", "ለባለቤቱ አገልግሎት", "ያለ ኪራይ / የሕዝብ አገልግሎት", "ሌላ"]
T_HOTEL, T_RES, T_OWN, T_FREE = 7, 3, 9, 10      # codes (1-based)

LISTS = {
    "L_TRACK": ["Private residence – cost estimate & home loan · የግል መኖሪያ ቤት",
                "Commercial project – full profitability · የንግድ ፕሮጀክት"],
    "L_MODE": ["1 – Build-to-Sell · ለሽያጭ መገንባት",
               "2 – Build-to-Rent / Hold · ለኪራይ መገንባት",
               "3 – Mixed (sell part, hold part) · ቅይጥ"],
    "L_DATASET": ["1 – My project · የእኔ ፕሮጀክት", "2 – SAMPLE 1: B+G+8 mixed-use · ምሳሌ 1",
                  "3 – SAMPLE 2: B+G+10 trade centre & hotel · ምሳሌ 2"],
    "L_PDATASET": ["1 – My house · የእኔ ቤት", "2 – SAMPLE: G+1 family house · ምሳሌ ቤት"],
    "L_SCEN": ["1 – Base · መሠረታዊ", "2 – Upside · ብሩህ", "3 – Downside · አስቸጋሪ", "4 – Lender case · የባንክ"],
    "L_FINISH": ["Basic", "Standard", "Premium"],
    "L_YN": ["Yes", "No"],
    "L_SPACE": SPACE_TYPES,
    "L_CONF": ["High", "Medium", "Low"],
    "L_METHOD": ["Rate x area", "Lump sum (BOQ)"],
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
    # private track
    "L_BTYPE": ["G+0 villa", "G+1", "G+2", "G+3"],
    "L_IDC": ["Paid monthly during construction", "Capitalised into the loan"],
    "L_PLAND": ["Already owned – no cost", "Buying / lease payment in the build period"],
}

SAMPLE_NOTE = "Sample value — not market advice"


def S(t, am=""):
    return ("sec", t, am)


INPUTS = [
    S("1. Project", "1. ፕሮጀክት"),
    ("PROJ_NAME", "Project name", "የፕሮጀክት ስም", "text", "txt", "My project",
     "SAMPLE 1 — B+G+8 mixed-use building, Addis Ababa", "SAMPLE 2 — B+G+10 trade centre & hotel, Addis Ababa",
     "Shown in every sheet header", "txt"),
    ("LOCATION", "Location (sub-city / town)", "አካባቢ (ክ/ከተማ / ከተማ)", "text", "txt", "", "Addis Ababa – Kirkos sub-city",
     "Addis Ababa – Lideta sub-city", "Free text", "txt"),
    ("LOC_TYPE", "Location class for cost factor", "የቦታ ዓይነት (ለወጪ ማስተካከያ)", "list", "gen", "Addis Ababa",
     "Addis Ababa", "Addis Ababa", "Regional towns use the regional location factor", "L_LOC"),
    ("PLOT_AREA", "Plot area", "የቦታ ስፋት", "m²", "n0", 0, 1000, 2000, "From title deed / lease contract", ("dec", 0, 1e7)),
    ("FAR", "Max floor-area ratio (FAR) permitted", "የተፈቀደ የወለል ስፋት ጥምርታ", "x", "n2", 0, 8.0, 8.0,
     "From the city master plan / building permit; 0 = not checked", ("dec", 0, 50)),
    ("START_DATE", "Construction start date", "የግንባታ መጀመሪያ ቀን", "date", "date", D0, D0, D0,
     "First month of construction (budget date for prices)", None),
    ("CONSTR_M", "Construction period", "የግንባታ ጊዜ", "months", "int", 24, 30, 36,
     "1–48 months; scenario delay months are added", ("whole", 1, 48)),
    ("FX_USD", "Memo: ETB per USD", "ብር በአንድ ዶላር (ማስታወሻ)", "ETB/USD", "n1", 155, 155, 155,
     "Memo only – the model runs in ETB. Verify current NBE rate", ("dec", 0, 10000)),
    ("FINISH", "Finish level", "የማጠናቀቂያ ደረጃ", "list", "gen", "Standard", "Standard", "Standard",
     "Selects the column of the cost library (block 12)", "L_FINISH"),

    S("2. Gross floor area by floor (net lettable / saleable areas go on sheet 13 Spaces)", "2. ጠቅላላ የወለል ስፋት በፎቅ"),
]
_S1 = {"Basement 1": 900, "Ground floor": 650, "Floor 1": 650, **{f"Floor {i}": 600 for i in range(2, 9)}}
_S2 = {"Basement 2": 1800, "Basement 1": 1800, "Ground floor": 1400, **{f"Floor {i}": 1300 for i in range(1, 11)}}
for i, fl in enumerate(FLOORS):
    INPUTS.append((f"FL_A{i+1}", f"{fl} – gross area", f"{FLOORS_AM[i]} – ጠቅላላ ስፋት", "m²", "n0", 0, _S1.get(fl, 0),
                   _S2.get(fl, 0), "Gross floor area incl. walls, cores & common areas", ("dec", 0, 1e6)))
INPUTS += [
    S("3. Construction cost", "3. የግንባታ ወጪ"),
]
for k, en, am in zip(EL_KEYS, ELEMENTS, ELEMENTS_AM):
    INPUTS.append((f"METH_{k}", f"{en} – entry method", f"{am} – አገባብ", "list", "gen", "Rate x area",
                   "Lump sum (BOQ)" if k == "SPEC" else "Rate x area", "Lump sum (BOQ)" if k == "SPEC" else "Rate x area",
                   "Rate x area uses the cost library (block 12); Lump sum = detailed BOQ override", "L_METHOD"))
for k, en, am in zip(EL_KEYS, ELEMENTS, ELEMENTS_AM):
    INPUTS.append((f"LUMP_{k}", f"{en} – lump sum (if BOQ)", f"{am} – ጠቅላላ ዋጋ (BOQ)", "ETB", "n0", 0,
                   45000000 if k == "SPEC" else 0, 95000000 if k == "SPEC" else 0,
                   "Used only when method = Lump sum (BOQ)", ("dec", 0, 1e13)))
INPUTS += [
    ("STOREY_UPLIFT", "Superstructure uplift per storey above G+3", "ከG+3 በላይ ለእያንዳንዱ ፎቅ ጭማሪ", "%", "pct", 0.02, 0.02, 0.02,
     "Taller frames cost more per m²", ("dec", 0, 0.2)),
    ("BASEMENT_RATE", "Basement extra (excavation, retaining, waterproofing)", "የምድር ቤት ተጨማሪ ወጪ", "ETB/m²", "n0", 12000, 12000,
     12000, "Applied to basement gross area", ("dec", 0, 1e6)),
    ("COMPOUND_YN", "Perimeter wall?", "የአጥር ግንብ?", "Y/N", "gen", "No", "Yes", "Yes", "", "yn"),
    ("COMPOUND_LEN", "Perimeter wall length", "የአጥር ርዝመት", "m", "n0", 0, 120, 180, "", ("dec", 0, 1e5)),
    ("COMPOUND_RATE", "Perimeter wall rate", "የአጥር ዋጋ በሜትር", "ETB/m", "n0", 18000, 20000, 20000, "Incl. gates", ("dec", 0, 1e7)),
    ("SEPTIC_YN", "Septic tank / treatment plant?", "የፍሳሽ ማስወገጃ?", "Y/N", "gen", "No", "No", "No", "Not needed where sewer exists", "yn"),
    ("SEPTIC_QTY", "Septic tanks – number", "የፍሳሽ ጉድጓድ ብዛት", "no.", "n0", 0, 0, 0, "", ("dec", 0, 100)),
    ("SEPTIC_RATE", "Septic tank – cost each", "የአንዱ ዋጋ", "ETB", "n0", 450000, 450000, 450000, "", ("dec", 0, 1e9)),
    ("WTANK_YN", "Water tanks?", "የውሃ ታንከር?", "Y/N", "gen", "No", "Yes", "Yes", "", "yn"),
    ("WTANK_QTY", "Water tanks – number", "የውሃ ታንከር ብዛት", "no.", "n0", 0, 4, 6, "", ("dec", 0, 100)),
    ("WTANK_RATE", "Water tank – cost each (incl. stand, pump)", "የአንዱ ዋጋ", "ETB", "n0", 250000, 250000, 250000, "",
     ("dec", 0, 1e9)),
    ("EXT_M2", "Paving / landscaping area", "የንጣፍ / የአትክልት ስፋት", "m²", "n0", 0, 250, 400, "", ("dec", 0, 1e6)),
    ("EXT_RATE", "Paving / landscaping rate", "የንጣፍ ዋጋ በካሬ", "ETB/m²", "n0", 4500, 4500, 4500, "", ("dec", 0, 1e6)),
    ("CONT_PCT", "Contingency (% of escalated hard cost)", "ያልተጠበቀ ወጪ መጠባበቂያ", "%", "pct", 0.10, 0.08, 0.08, "",
     ("dec", 0, 1)),
    ("DESIGN_PCT", "Design fee (% of hard cost)", "የዲዛይን ክፍያ", "%", "pct", 0.05, 0.05, 0.05, "", ("dec", 0, 1)),
    ("DESIGN_UPFRONT", "Design fee paid upfront (month 1)", "በመጀመሪያ ወር የሚከፈል ድርሻ", "%", "pct", 0.5, 0.5, 0.5,
     "Rest follows the S-curve", ("dec", 0, 1)),
    ("SUPERV_PCT", "Supervision (% of escalated hard cost)", "የቁጥጥር ክፍያ", "%", "pct", 0.03, 0.03, 0.03, "", ("dec", 0, 1)),
    ("CONSULT_PCT", "Specialist consultancy (% of escalated hard cost)", "የልዩ ባለሙያ አማካሪ ክፍያ", "%", "pct", 0, 0, 0.015,
     "Hotel, façade, MEP, fire or traffic specialists; capitalised, follows the S-curve", ("dec", 0, 1)),
    ("PERMITS", "Permits & municipal fees", "ፈቃድና የማዘጋጃ ቤት ክፍያዎች", "ETB", "n0", 0, 2500000, 6000000, "Paid month 1",
     ("dec", 0, 1e12)),
    ("FFE", "Furniture, fittings & equipment (FF&E)", "የቤት ዕቃዎችና መገልገያዎች", "ETB", "n0", 0, 6000000, 25000000,
     "Paid in the completion month; depreciated over the FF&E life", ("dec", 0, 1e12)),
    ("PREOPEN", "Pre-opening, leasing launch & marketing (expensed)", "ከመክፈቻ በፊት ወጪና የገበያ ማስተዋወቂያ", "ETB", "n0", 0,
     1500000, 18000000, "Paid in the completion month; expensed in P&L (not capitalised) but funded like construction cost",
     ("dec", 0, 1e12)),
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
    ("RET_REL_M", "Retention released – months after completion", "የሚለቀቅበት ጊዜ", "months", "int", 12, 12, 12, "",
     ("whole", 0, 60)),
    ("VAT_CAPEX_PCT", "Share of capex carrying input VAT", "ተ.እ.ታ ያለበት የወጪ ድርሻ", "%", "pct", 0.7, 0.7, 0.7,
     "Reclaimable against output VAT or refunded", ("dec", 0, 1)),

    S("4. Land", "4. መሬት"),
    ("LAND_TYPE", "Land arrangement", "የመሬት ይዞታ ዓይነት", "list", "gen", "Lease (cash)", "Lease (cash)",
     "Partner in-kind (NCI)", "In-kind partner (e.g. city PPP) = non-controlling interest (NCI)", "L_LAND"),
    ("LAND_VALUE", "Land lease price / market value", "የሊዝ ዋጋ / የገበያ ዋጋ", "ETB", "n0", 0, 150000000, 250000000,
     "Lease price, or agreed value of the land contributed in kind", ("dec", 0, 1e13)),
    ("LEASE_DP", "Lease down payment", "የሊዝ ቅድመ ክፍያ", "%", "pct", 0.10, 0.10, 0.10, "Paid in month 1", ("dec", 0, 1)),
    ("LEASE_N", "Lease instalment years", "የሊዝ ክፍያ ዓመታት", "years", "int", 30, 30, 30, "Annual instalments",
     ("whole", 1, 99)),
    ("LEASE_INT", "Interest on unpaid lease balance", "ባልተከፈለ ሊዝ ላይ ወለድ", "% p.a.", "pct", 0.07, 0.07, 0.07,
     "Per lease contract; verify", ("dec", 0, 1)),

    S("5. Sales (prices per space are on sheet 13 Spaces)", "5. ሽያጭ"),
    ("PRICE_ESC", "Sale price escalation", "የሽያጭ ዋጋ ጭማሪ", "% p.a.", "pct", 0.10, 0.10, 0.10, "", ("dec", -0.5, 2)),
    ("PRESALE_1", "Presold in construction year 1 (% of saleable)", "በ1ኛው የግንባታ ዓመት ቅድመ-ሽያጭ", "%", "pct", 0, 0.25, 0.10, "",
     ("dec", 0, 1)),
    ("PRESALE_2", "Presold in construction year 2", "በ2ኛው ዓመት", "%", "pct", 0, 0.25, 0.20, "", ("dec", 0, 1)),
    ("PRESALE_3", "Presold in construction year 3", "በ3ኛው ዓመት", "%", "pct", 0, 0.20, 0.20, "", ("dec", 0, 1)),
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

    S("6. Leasing (rent and stabilised occupancy per space are on sheet 13 Spaces)", "6. ኪራይ"),
    ("RAMP_O1", "Lease-up: occupancy in O1 as % of stabilised", "በ1ኛው ዓመት የተያዘ (ከተረጋጋው %)", "%", "pct", 0.65, 0.65, 0.70,
     "Applied to every space's stabilised occupancy", ("dec", 0, 1)),
    ("RAMP_O2", "Lease-up: occupancy in O2 as % of stabilised", "በ2ኛው ዓመት", "%", "pct", 0.87, 0.87, 0.90,
     "O3 onwards = 100% of stabilised", ("dec", 0, 1)),
    ("RENT_ESC", "Rent escalation", "የኪራይ ጭማሪ", "% p.a.", "pct", 0.08, 0.08, 0.08, "From today's prices", ("dec", -0.5, 2)),
    ("KEY_MONTHS", "Key money (months of rent, received O1)", "የቁልፍ ገንዘብ (በወራት ኪራይ)", "months", "n1", 0, 6, 6,
     "Deferred and released over the period below (IFRS 16)", ("dec", 0, 60)),
    ("KEY_AMORT_Y", "Key money release period", "የቁልፍ ገንዘብ የሚታወቅበት ጊዜ", "years", "int", 5, 5, 5, "", ("whole", 1, 10)),
    ("SC_RECOV", "Service charge recovery (% of property opex)", "የአገልግሎት ክፍያ ማስመለሻ", "%", "pct", 0.6, 0.6, 0.6, "",
     ("dec", 0, 1)),
    ("BAD_DEBT", "Bad debt (% of rent)", "የማይሰበሰብ ዕዳ", "%", "pct", 0.02, 0.02, 0.02, "", ("dec", 0, 1)),
    ("RES_VAT_EXEMPT", "Residential rent exempt from VAT?", "የመኖሪያ ኪራይ ከተ.እ.ታ ነፃ ነው?", "Y/N", "gen", "Yes", "Yes", "Yes",
     "Share of rent subject to VAT is computed from the rent roll – verify", "yn"),

    S("7. Hotel operator lease (used when sheet 13 has 'Hotel (operator lease)' space)", "7. የሆቴል ኦፕሬተር ኪራይ"),
    ("HOT_KEYS", "Hotel rooms (keys)", "የሆቴል ክፍሎች ብዛት", "no.", "n0", 0, 0, 85, "", ("whole", 0, 2000)),
    ("HOT_ADR", "Average daily rate (today's prices)", "አማካይ የቀን ክፍል ዋጋ", "ETB/night", "n0", 0, 0, 12000,
     "Room revenue ÷ rooms sold", ("dec", 0, 1e7)),
    ("HOT_ADR_ESC", "ADR escalation", "የክፍል ዋጋ ጭማሪ", "% p.a.", "pct", 0.10, 0.10, 0.10, "", ("dec", -0.5, 2)),
    ("HOT_OCC1", "Hotel occupancy O1", "የሆቴል ይዞታ O1", "%", "pct", 0.45, 0.45, 0.45, "", ("dec", 0, 1)),
    ("HOT_OCC2", "Hotel occupancy O2", "የሆቴል ይዞታ O2", "%", "pct", 0.58, 0.58, 0.58, "", ("dec", 0, 1)),
    ("HOT_OCC3", "Hotel occupancy – stabilised (O3+)", "የተረጋጋ የሆቴል ይዞታ", "%", "pct", 0.65, 0.65, 0.65, "", ("dec", 0, 1)),
    ("HOT_FB", "F&B and other hotel revenue (% of rooms revenue)", "የምግብና ሌላ ገቢ (ከክፍል ገቢ %)", "%", "pct", 0.6, 0.6, 0.6,
     "", ("dec", 0, 5)),
    ("HOT_TURN", "Turnover rent (% of total hotel revenue)", "ከጠቅላላ ገቢ የሚከፈል ኪራይ", "%", "pct", 0.18, 0.18, 0.18,
     "Paid by the operator to the owner", ("dec", 0, 1)),
    ("HOT_MIN", "Minimum guaranteed rent p.a. (today's prices)", "ዝቅተኛ ዋስትና ያለው ኪራይ", "ETB/yr", "n0", 0, 0, 20000000,
     "Escalated with rent escalation; owner receives the higher of minimum and turnover rent", ("dec", 0, 1e12)),

    S("8. Other income streams (year-1 amounts, today's prices)", "8. ሌሎች ገቢዎች"),
    ("OINC1_NAME", "Stream 1 – name", "ገቢ 1 – ስም", "text", "txt", "Signage & billboards", "Signage & billboards",
     "Signage & billboards", "", "txt"),
    ("OINC1", "Stream 1 – amount", "ገቢ 1 – መጠን", "ETB/yr", "n0", 0, 400000, 2400000, "", ("dec", 0, 1e12)),
    ("OINC2_NAME", "Stream 2 – name", "ገቢ 2 – ስም", "text", "txt", "Telecom rooftop", "Telecom rooftop", "Telecom rooftop", "",
     "txt"),
    ("OINC2", "Stream 2 – amount", "ገቢ 2 – መጠን", "ETB/yr", "n0", 0, 200000, 600000, "", ("dec", 0, 1e12)),
    ("OINC3_NAME", "Stream 3 – name", "ገቢ 3 – ስም", "text", "txt", "Events & exhibitions", "Events & exhibitions",
     "Events & exhibitions", "", "txt"),
    ("OINC3", "Stream 3 – amount", "ገቢ 3 – መጠን", "ETB/yr", "n0", 0, 0, 3000000, "", ("dec", 0, 1e12)),
    ("OINC4_NAME", "Stream 4 – name", "ገቢ 4 – ስም", "text", "txt", "Other", "Other", "Visitor parking (hourly)", "", "txt"),
    ("OINC4", "Stream 4 – amount", "ገቢ 4 – መጠን", "ETB/yr", "n0", 0, 0, 1800000, "", ("dec", 0, 1e12)),
    ("OI_ESC", "Other income escalation", "የሌሎች ገቢዎች ጭማሪ", "% p.a.", "pct", 0.08, 0.08, 0.08, "", ("dec", -0.5, 2)),

    S("9. Operating costs", "9. የሥራ ማስኬጃ ወጪ"),
    ("OPEX_METHOD", "Property opex basis", "የወጪ አሰላል", "list", "gen", "% of rent", "% of rent", "% of rent", "", "L_OPEXM"),
    ("OPEX_PCT", "Property opex (% of rent)", "የንብረት ወጪ (ከኪራይ %)", "%", "pct", 0.15, 0.15, 0.15, "", ("dec", 0, 1)),
    ("OPEX_M2", "Property opex (ETB/m²/yr, today's prices)", "የንብረት ወጪ በካሬ", "ETB/m²/yr", "n0", 1200, 1200, 1200, "",
     ("dec", 0, 1e6)),
    ("MGMT_PCT", "Management fee (% of rent)", "የአስተዳደር ክፍያ", "%", "pct", 0.03, 0.03, 0.03, "", ("dec", 0, 1)),
    ("INS_PCT", "Insurance (% of building value p.a.)", "መድን", "%", "pct2", 0.0015, 0.0015, 0.0015, "", ("dec", 0, 0.1)),
    ("PROPTAX_PCT", "Property tax (% of land + building value p.a.)", "የንብረት ግብር", "%", "pct2", 0.001, 0.001, 0.001,
     "Verify current city proclamation", ("dec", 0, 0.1)),
    ("MAINT_PCT", "Maintenance reserve (% of building value p.a.)", "የጥገና መጠባበቂያ", "%", "pct2", 0.005, 0.005, 0.005, "",
     ("dec", 0, 0.2)),
    ("OVERHEAD", "Owner overheads (year 1, today's prices)", "አጠቃላይ የአስተዳደር ወጪ", "ETB/yr", "n0", 0, 1200000, 3000000, "",
     ("dec", 0, 1e12)),
    ("OPEX_ESC", "Opex escalation", "የወጪ ጭማሪ", "% p.a.", "pct", 0.10, 0.10, 0.10, "", ("dec", -0.5, 2)),
    ("OPEX_VAT_PCT", "Share of opex carrying input VAT", "ተ.እ.ታ ያለበት የወጪ ድርሻ", "%", "pct", 0.5, 0.5, 0.5, "",
     ("dec", 0, 1)),
    ("RECV_DAYS", "Rent receivable days", "የኪራይ ተሰብሳቢ ቀናት", "days", "int", 30, 30, 30, "", ("whole", 0, 365)),
    ("PAY_DAYS", "Opex payable days", "የወጪ ተከፋይ ቀናት", "days", "int", 30, 30, 30, "", ("whole", 0, 365)),

    S("10. Financing", "10. ፋይናንስ"),
    ("EQ_PCT", "Equity share of funding need", "የባለቤት ካፒታል ድርሻ", "%", "pct", 0.4, 0.4, 0.5,
     "Debt = 1 − equity, capped by the loan cap", ("dec", 0, 1)),
    ("DRAW_ORDER", "Draw order", "የአወጣጥ ቅደም ተከተል", "list", "gen", "Equity first", "Equity first", "Equity first", "",
     "L_DRAW"),
    ("PRESALE_USE", "Use presale cash to fund construction?", "የቅድመ-ሽያጭ ገንዘብ ለግንባታ ይውላል?", "Y/N", "gen", "Yes", "Yes",
     "Yes", "", "yn"),
    ("PRESALE_USE_PCT", "Share of presale cash usable (rest in escrow)", "ጥቅም ላይ የሚውለው ድርሻ", "%", "pct", 0.8, 0.8, 0.8, "",
     ("dec", 0, 1)),
    ("DEBT_HEADROOM", "Facility headroom for bank fees (% of debt)", "ለባንክ ክፍያዎች ተጨማሪ የብድር መጠን", "%", "pct", 0.03, 0.03, 0.03,
     "Lets the loan also fund the arrangement and commitment fees, so no gap is left for them", ("dec", 0, 0.5)),
    ("LOAN_CAP", "Loan cap (0 = no cap)", "የብድር ጣሪያ", "ETB", "n0", 0, 0, 0, "", ("dec", 0, 1e13)),
    ("RATE_TYPE", "Interest rate type", "የወለድ ዓይነት", "list", "gen", "Fixed", "Fixed", "Fixed", "", "L_RATETYPE"),
    ("INT_RATE", "Interest rate (fixed, and during construction)", "የወለድ መጠን", "% p.a.", "pct", 0.15, 0.15, 0.15,
     "Verify bank term sheet", ("dec", 0, 1)),
]
for y in range(1, 11):
    INPUTS.append((f"RATE_O{y}", f"Variable path – rate in O{y}", f"ተለዋዋጭ ወለድ O{y}", "% p.a.", "pct", 0.15, 0.15, 0.15,
                   "Used when type = Variable", ("dec", 0, 1)))
INPUTS += [
    ("GRACE_M", "Grace period after completion", "ከግንባታ በኋላ የእፎይታ ጊዜ", "months", "int", 0, 12, 12,
     "Interest during construction is always capitalised (IDC)", ("whole", 0, 120)),
    ("GRACE_TYPE", "Interest during grace", "በእፎይታ ጊዜ ወለድ", "list", "gen", "Interest-only (paid)", "Interest-only (paid)",
     "Capitalised", "Capitalised suits lease-up: no debt service while occupancy builds", "L_GRACE"),
    ("TENOR_Y", "Tenor after completion (incl. grace)", "የብድር ጊዜ (እፎይታን ጨምሮ)", "years", "int", 10, 10, 15, "Max 20",
     ("whole", 1, 20)),
    ("REPAY_TYPE", "Repayment type", "የአከፋፈል ዓይነት", "list", "gen", "Annuity (PMT)", "Annuity (PMT)", "Annuity (PMT)", "",
     "L_REPAY"),
    ("BALLOON_PCT", "Balloon (% of balance at repayment start)", "የመጨረሻ ትልቅ ክፍያ", "%", "pct", 0, 0, 0,
     "Only for repayment type = Balloon", ("dec", 0, 1)),
    ("FREQ", "Repayment frequency", "የክፍያ ድግግሞሽ", "list", "gen", "Monthly", "Quarterly", "Quarterly", "", "L_FREQ"),
    ("ARR_FEE", "Arrangement fee (% of facility, excl. IDC)", "የብድር ማመቻቻ ክፍያ", "%", "pct", 0.015, 0.015, 0.015,
     "Base excludes IDC → no circularity", ("dec", 0, 0.2)),
    ("COMMIT_FEE", "Commitment fee on undrawn (% p.a.)", "ባልወጣ ብድር ላይ ክፍያ", "%", "pct", 0.005, 0.005, 0.005, "", ("dec", 0, 0.2)),
    ("DSRA_M", "Debt service reserve (months of next year's debt service)", "የዕዳ ክፍያ መጠባበቂያ", "months", "int", 0, 6, 6, "",
     ("whole", 0, 24)),
    ("SWEEP", "Cash sweep (% of sales proceeds to prepay loan)", "ከሽያጭ ለብድር ቅድመ ክፍያ", "%", "pct", 0, 0.40, 0.40, "",
     ("dec", 0, 1)),
    ("DSCR_COV", "DSCR covenant", "የዕዳ ክፍያ ሽፋን ግዴታ", "x", "x", 1.3, 1.3, 1.3, "Dividends locked up below covenant",
     ("dec", 0, 10)),
    ("DIST_SURPLUS", "Return surplus cash to shareholders (capital reduction)", "ትርፍ ገንዘብን ለባለአክሲዮኖች መመለስ", "Y/N", "gen",
     "Yes", "Yes", "Yes", "Subject to the DSCR lock-up test: first returns paid-in capital, then pays a special dividend from retained earnings", "yn"),
    ("MIN_CASH", "Minimum cash balance (operations)", "ዝቅተኛ ጥሬ ገንዘብ", "ETB", "n0", 0, 2000000, 5000000, "", ("dec", 0, 1e12)),

    S("11. Tax, accounting & valuation", "11. ግብር፣ ሂሳብ አያያዝና ግምገማ"),
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
    ("DIV_PAYOUT", "Dividend payout ratio", "የትርፍ ድርሻ ክፍፍል መጠን", "%", "pct", 0.5, 0.5, 0.5, "", ("dec", 0, 1)),
    ("DIV_WHT", "Dividend withholding tax (memo)", "የትርፍ ድርሻ ተቀናሽ ግብር", "%", "pct", 0.10, 0.10, 0.10, "", ("dec", 0, 1)),
    ("LR_PCT", "Legal reserve transfer (% of profit)", "ሕጋዊ መጠባበቂያ", "%", "pct", 0.05, 0.05, 0.05, "", ("dec", 0, 1)),
    ("LR_CAP", "Legal reserve cap (% of share capital)", "የሕጋዊ መጠባበቂያ ጣሪያ", "%", "pct", 0.10, 0.10, 0.10, "",
     ("dec", 0, 1)),
    ("LIFE_BLD", "Useful life – building structure", "የሕንፃ አገልግሎት ዘመን", "years", "int", 50, 50, 50, "", ("whole", 1, 100)),
    ("LIFE_SYS", "Useful life – building systems (MEP, lifts)", "የሲስተሞች አገልግሎት ዘመን", "years", "int", 20, 20, 20, "",
     ("whole", 1, 100)),
    ("LIFE_FFE", "Useful life – FF&E", "የዕቃዎች አገልግሎት ዘመን", "years", "int", 7, 7, 7, "", ("whole", 1, 50)),
    ("FFE_LIFE_CAP", "FF&E life cap", "የዕቃዎች ዘመን ጣሪያ", "years", "int", 10, 10, 10, "", ("whole", 1, 50)),
    ("DISC_RATE", "Project discount rate", "የፕሮጀክት የቅናሽ መጠን", "%", "pct", 0.18, 0.18, 0.18, "", ("dec", 0, 1)),
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

# default confidence per input (column K of 10_Inputs, user-editable)
LOW_CONF = {"HOT_ADR", "HOT_OCC3", "EXIT_CAP", "RAMP_O1", "PRESALE_1", "PRESALE_2", "PRESALE_3"}
HIGH_CONF = {"PROJ_NAME", "LOCATION", "LOC_TYPE", "PLOT_AREA", "START_DATE", "FX_USD", "VAT_R", "CIT", "LR_PCT", "LR_CAP",
             "LAND_TYPE", "LAND_VALUE"}


def confidence(key):
    if key in LOW_CONF:
        return "Low"
    if key in HIGH_CONF or key.startswith("FL_A"):
        return "High"
    return "Medium"


# ---------------------------------------------------------------- space schedule (13_Spaces)
# (name, type, net m², sold % (Mixed mode), sale price ETB/m², rent ETB/m²/month, stabilised occupancy)
N_SPACES = 12
SPACES_USER = [("", "Retail", 0, 0, 0, 0, 0.92)] * N_SPACES
SPACES_S1 = [
    ("Ground-floor shops", "Retail", 552, 0, 180000, 2000, 0.92),
    ("First-floor shops", "Retail", 553, 0, 180000, 2000, 0.92),
    ("Apartments F2–F8", "Residential", 3360, 1.0, 125000, 650, 0.92),
]
SPACES_S2 = [
    ("Ground-floor shops", "Retail", 1000, 0.3, 230000, 2200, 0.95),
    ("Retail galleria F1–F2", "Retail", 1900, 0.3, 190000, 1500, 0.90),
    ("Wholesale showrooms B1", "Showroom / wholesale", 1300, 0, 120000, 900, 0.90),
    ("Food court F3", "Food & beverage", 800, 0, 0, 1400, 0.90),
    ("Offices F4–F6", "Office", 2900, 0, 110000, 800, 0.85),
    ("Hotel F7–F10 (85 keys)", "Hotel (operator lease)", 4100, 0, 0, 0, 1.0),
    ("Owner head office F3", "Owner-occupied", 200, 0, 0, 0, 1.0),
    ("City one-stop service desk", "Rent-free / public service", 150, 0, 0, 0, 1.0),
    ("Tenant parking B2", "Parking", 1500, 0, 0, 120, 0.85),
    ("Storage units B1", "Warehouse / storage", 300, 0, 0, 500, 0.85),
]
for lst in (SPACES_S1, SPACES_S2):
    while len(lst) < N_SPACES:
        lst.append(("", "Retail", 0, 0, 0, 0, 0.92))
