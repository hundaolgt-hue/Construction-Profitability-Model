"""Private residence track: inputs (P1_Inputs). Columns: My house, Sample house.

Row tuple: (key, label_en, label_am, unit, fmt, user, sample, note, validation)
All names carry the P_ prefix so they never clash with the commercial model.
"""
import datetime as dt

D0 = dt.datetime(2027, 1, 1)

P_FLOORS = ["Basement", "Ground floor", "Floor 1", "Floor 2", "Floor 3"]
P_FLOORS_AM = ["ምድር ቤት", "ምድር ወለል", "1ኛ ፎቅ", "2ኛ ፎቅ", "3ኛ ፎቅ"]

# residential cost library, ETB per m² gross, Basic / Standard / Premium (Sample values — not market advice)
P_EL = [("PRELIM", "Preliminaries & site set-up", "ቅድመ-ሥራዎች", (900, 1100, 1400)),
        ("SUBSTR", "Foundation / substructure", "የመሠረት ሥራ", (3200, 3500, 3900)),
        ("SUPER", "Frame (columns, beams, slabs)", "ፍሬም (ኮለን፣ ቢም፣ ስላብ)", (6200, 6800, 7400)),
        ("ENVEL", "Walls, windows & roof", "ግድግዳ፣ መስኮትና ጣሪያ", (3600, 4500, 6200)),
        ("INTERN", "Internal walls & doors", "የውስጥ ግድግዳና በሮች", (1600, 2200, 3200)),
        ("FINISH", "Finishes (floor, plaster, paint, tiles)", "ማጠናቀቂያ (ወለል፣ ልስን፣ ቀለም፣ ሴራሚክ)", (3000, 5200, 8800)),
        ("MEP", "Electrical, plumbing & sanitary", "ኤሌክትሪክ፣ ቧንቧና ንፅህና", (2600, 3800, 5800)),
        ("KITCH", "Kitchen & built-in cabinets", "ማዕድ ቤትና ቁም ሣጥኖች", (600, 1100, 2200))]

P_STAGES = [("1 Mobilisation & site set-up", "ቅድመ-ዝግጅት", ["PRELIM"]),
            ("2 Foundation", "መሠረት", ["SUBSTR", "BASE"]),
            ("3 Frame (columns, beams, slabs)", "ፍሬም", ["SUPER"]),
            ("4 Walls, windows & roof", "ግድግዳ፣ መስኮትና ጣሪያ", ["ENVEL"]),
            ("5 Electrical, plumbing & internal walls", "ኤሌክትሪክ፣ ቧንቧና የውስጥ ግድግዳ", ["MEP", "INTERN"]),
            ("6 Finishes, kitchen & external works", "ማጠናቀቂያና የውጭ ሥራ", ["FINISH", "KITCH", "EXT"])]


def S(t, am=""):
    return ("sec", t, am)


P_INPUTS = [
    S("1. Your house", "1. ቤትዎ"),
    ("P_NAME", "Project name", "የፕሮጀክት ስም", "text", "txt", "My house", "SAMPLE — G+1 family house, Addis Ababa",
     "Shown on the estimate", "txt"),
    ("P_LOCATION", "Location (sub-city / town)", "አካባቢ", "text", "txt", "", "Addis Ababa – Bole sub-city", "", "txt"),
    ("P_LOC_TYPE", "Location class for cost factor", "የቦታ ዓይነት", "list", "gen", "Addis Ababa", "Addis Ababa",
     "Regional towns use a lower location factor", "L_LOC"),
    ("P_PLOT", "Plot area", "የቦታ ስፋት", "m²", "n0", 0, 300, "From your title deed", ("dec", 0, 1e6)),
    ("P_START", "Construction start date", "የግንባታ መጀመሪያ ቀን", "date", "date", D0, D0, "Budget date for prices", None),
    ("P_MONTHS", "Construction period", "የግንባታ ጊዜ", "months", "int", 12, 14, "1–48 months", ("whole", 1, 48)),
    ("P_BTYPE", "Building type", "የሕንፃው ዓይነት", "list", "gen", "G+1", "G+1", "Checked against the floors below", "L_BTYPE"),
    ("P_FINISH", "Finish level", "የማጠናቀቂያ ደረጃ", "list", "gen", "Standard", "Standard", "Selects the cost-library column",
     "L_FINISH"),
    ("P_BASE_YN", "Basement?", "ምድር ቤት አለው?", "Y/N", "gen", "No", "No", "Checked against the floors below", "yn"),
    S("2. Gross floor area (incl. walls and stairs)", "2. ጠቅላላ የወለል ስፋት"),
]
_PA = {"Ground floor": 130, "Floor 1": 120}
for i, fl in enumerate(P_FLOORS):
    P_INPUTS.append((f"P_FL{i+1}", f"{fl} – gross area", f"{P_FLOORS_AM[i]} – ጠቅላላ ስፋት", "m²", "n0", 0, _PA.get(fl, 0), "",
                     ("dec", 0, 1e5)))
P_INPUTS += [
    S("3. Construction extras and fees", "3. ተጨማሪ ሥራዎችና ክፍያዎች"),
    ("P_BASE_RATE", "Basement extra (excavation, retaining walls)", "የምድር ቤት ተጨማሪ", "ETB/m²", "n0", 10000, 10000,
     "Applied to basement area", ("dec", 0, 1e6)),
    ("P_COMP_YN", "Compound wall?", "የአጥር ግንብ?", "Y/N", "gen", "No", "Yes", "", "yn"),
    ("P_COMP_LEN", "Compound wall length", "የአጥር ርዝመት", "m", "n0", 0, 70, "", ("dec", 0, 1e4)),
    ("P_COMP_RATE", "Compound wall rate (incl. gate)", "የአጥር ዋጋ በሜትር", "ETB/m", "n0", 18000, 18000, "", ("dec", 0, 1e7)),
    ("P_SEPT_YN", "Septic tank?", "የፍሳሽ ጉድጓድ?", "Y/N", "gen", "No", "Yes", "Not needed where sewer exists", "yn"),
    ("P_SEPT_RATE", "Septic tank cost", "የፍሳሽ ጉድጓድ ዋጋ", "ETB", "n0", 450000, 450000, "", ("dec", 0, 1e9)),
    ("P_WTANK_QTY", "Water tanks – number", "የውሃ ታንከር ብዛት", "no.", "n0", 0, 2, "", ("dec", 0, 50)),
    ("P_WTANK_RATE", "Water tank cost each (stand, pump)", "የአንዱ ዋጋ", "ETB", "n0", 60000, 60000, "", ("dec", 0, 1e8)),
    ("P_PAVE_M2", "Paving / garden area", "የንጣፍ / የአትክልት ስፋት", "m²", "n0", 0, 80, "", ("dec", 0, 1e5)),
    ("P_PAVE_RATE", "Paving / garden rate", "የንጣፍ ዋጋ", "ETB/m²", "n0", 3500, 3500, "", ("dec", 0, 1e6)),
    ("P_CONT", "Contingency (% of escalated building cost)", "ያልተጠበቀ ወጪ መጠባበቂያ", "%", "pct", 0.10, 0.10, "", ("dec", 0, 1)),
    ("P_DESIGN", "Design & permit drawings (% of building cost)", "የዲዛይን ክፍያ", "%", "pct", 0.04, 0.04, "Paid in month 1",
     ("dec", 0, 1)),
    ("P_SUPERV", "Site supervision (% of building cost)", "የቁጥጥር ክፍያ", "%", "pct", 0.02, 0.02, "", ("dec", 0, 1)),
    ("P_PERMITS", "Building permit & utility connections", "ፈቃድና የአገልግሎት ቅጥያ", "ETB", "n0", 0, 120000, "Paid in month 1",
     ("dec", 0, 1e10)),
    ("P_FURN", "Furniture & appliances", "የቤት ዕቃዎች", "ETB", "n0", 0, 1000000, "Paid in the last month", ("dec", 0, 1e10)),
    ("P_ESC", "Price escalation", "የዋጋ ጭማሪ", "% p.a.", "pct", 0.15, 0.15, "Building prices rise during construction",
     ("dec", -0.5, 2)),
    ("P_SCURVE", "Spend profile", "የወጪ ስርጭት", "list", "gen", "S-curve (sigmoid)", "S-curve (sigmoid)", "", "L_SCURVE"),
    ("P_RETENTION", "Retention held from contractor", "ከኮንትራክተር የሚያዝ", "%", "pct", 0.05, 0.05,
     "Paid after the defects period", ("dec", 0, 0.5)),
    ("P_RET_M", "Retention released – months after completion", "የሚለቀቅበት ጊዜ", "months", "int", 6, 6, "", ("whole", 0, 36)),
    ("P_EST_LOW", "Estimate range – low side", "የግምት ዝቅተኛ ወሰን", "%", "pct", -0.10, -0.10, "", ("dec", -0.9, 0)),
    ("P_EST_HIGH", "Estimate range – high side", "የግምት ከፍተኛ ወሰን", "%", "pct", 0.15, 0.15, "", ("dec", 0, 2)),
    S("4. Land", "4. መሬት"),
    ("P_LAND_TYPE", "Land", "መሬት", "list", "gen", "Already owned – no cost", "Already owned – no cost", "", "L_PLAND"),
    ("P_LAND_COST", "Land / lease payment during construction", "የመሬት / የሊዝ ክፍያ", "ETB", "n0", 0, 0,
     "Paid in month 1 when land type = Buying", ("dec", 0, 1e12)),
    S("5. Funding", "5. የገንዘብ ምንጭ"),
    ("P_SAVINGS", "Own savings available", "የራስ ቁጠባ", "ETB", "n0", 0, 5000000, "Used first", ("dec", 0, 1e12)),
    ("P_FAMILY", "Family / diaspora contributions", "የቤተሰብ / ዳያስፖራ መዋጮ", "ETB", "n0", 0, 2500000, "Used after savings",
     ("dec", 0, 1e12)),
    ("P_LOAN_MAX", "Home loan – maximum the bank will lend", "ከፍተኛ የቤት ብድር", "ETB", "n0", 0, 7000000, "Used last",
     ("dec", 0, 1e12)),
    ("P_RATE", "Home loan interest rate", "የብድር ወለድ", "% p.a.", "pct", 0.14, 0.14, "Verify the bank offer", ("dec", 0, 1)),
    ("P_TENOR", "Repayment period after completion", "የመክፈያ ጊዜ", "years", "int", 15, 15, "Max 20", ("whole", 1, 20)),
    ("P_GRACE", "Grace period after completion (interest only)", "የእፎይታ ጊዜ", "months", "int", 0, 0, "", ("whole", 0, 60)),
    ("P_IDC", "Interest while building", "በግንባታ ወቅት ወለድ", "list", "gen", "Capitalised into the loan",
     "Capitalised into the loan", "Paid monthly = you pay it from income; Capitalised = added to the loan", "L_IDC"),
    ("P_ARR", "Loan arrangement fee (% of loan limit)", "የብድር ማመቻቻ ክፍያ", "%", "pct", 0.01, 0.01, "Paid in month 1",
     ("dec", 0, 0.2)),
    S("6. Affordability", "6. የመክፈል አቅም"),
    ("P_INCOME", "Household monthly net income", "የቤተሰብ ወርሃዊ የተጣራ ገቢ", "ETB/month", "n0", 0, 300000,
     "All earners, after tax", ("dec", 0, 1e9)),
    ("P_OTHER_DEBT", "Other monthly loan payments", "ሌሎች ወርሃዊ የብድር ክፍያዎች", "ETB/month", "n0", 0, 0, "", ("dec", 0, 1e9)),
    ("P_AFF_MAX", "Max total debt payments as % of income", "ከገቢ ከፍተኛ የክፍያ ድርሻ", "%", "pct", 0.35, 0.35,
     "Typical bank rule 30–40%", ("dec", 0, 1)),
    S("7. Optional: rent out part of the house", "7. አማራጭ፡ ከፊሉን ማከራየት"),
    ("P_RENT_YN", "Rent out part of the house?", "ከፊሉን ያከራያሉ?", "Y/N", "gen", "No", "Yes", "", "yn"),
    ("P_RENT_AREA", "Area rented out", "የሚከራየው ስፋት", "m²", "n0", 0, 60, "e.g. ground-floor unit", ("dec", 0, 1e5)),
    ("P_RENT_M", "Monthly rent (today's prices)", "ወርሃዊ ኪራይ", "ETB/month", "n0", 0, 25000, "", ("dec", 0, 1e9)),
    ("P_RENT_ESC", "Rent increase per year", "የኪራይ ጭማሪ", "% p.a.", "pct", 0.08, 0.08, "", ("dec", -0.5, 2)),
    ("P_VACANCY", "Months empty per year (vacancy)", "ባዶ የሚሆንባቸው ወራት", "%", "pct", 0.08, 0.08, "≈ 1 month a year",
     ("dec", 0, 1)),
    ("P_RENT_COST", "Running costs & rental income tax (% of rent)", "ወጪና የኪራይ ገቢ ግብር", "%", "pct", 0.20, 0.20,
     "Repairs, cleaning, schedule-B rental income tax – verify", ("dec", 0, 1)),
]
