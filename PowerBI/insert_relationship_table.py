"""Insert Power BI relationship schema table into the guide docx XML."""
import xml.sax.saxutils as sax

DOC_PATH = r"D:\Work\Driver Survey\PowerBI\unpacked_guide3\word\document.xml"

def esc(text):
    """Escape XML special characters in text content."""
    return sax.escape(str(text))

# ── Table data ─────────────────────────────────────────────────────────────────
HEADER = ["View / Table", "SQL Sources From", "Time Columns", "Power BI Relationship", "Role / Notes"]

C_HEAD   = ("2E75B6", "FFFFFF")
C_BASE   = ("D9D9D9", "000000")
C_DAX    = ("FCE4D6", "000000")
C_RA     = ("E2EFDA", "000000")
C_WEEKLY = ("DEEAF1", "000000")
C_CITY   = ("FFFFFF", "000000")
C_SEC    = ("F2F2F2", "595959")

ROWS = [
    (C_SEC,  ["BASE SQL TABLES", "", "", "", ""]),
    (C_BASE, ["DriverSurvey_ShortMain",   "-- (base table)",  "yearweek INT, weeknumber INT", "--", "Primary fact table for all Short-survey views"]),
    (C_BASE, ["DriverSurvey_ShortRare",   "-- (base table)",  "--",                           "--", "Rare-question supplement; joined to ShortMain on recordID"]),
    (C_BASE, ["DriverSurvey_WideMain",    "-- (base table)",  "--",                           "--", "One-hot columns for incentive-type questions"]),
    (C_BASE, ["DriverSurvey_LongMain",    "-- (base table)",  "--",                           "--", "Long-format answers for open-ended questions"]),
    (C_BASE, ["DriverSurvey_LongRare",    "-- (base table)",  "--",                           "--", "Long-format rare answers"]),

    (C_SEC,  ["DAX CALCULATED TABLE", "", "", "", ""]),
    (C_DAX,  ["WeekList", "DAX: DISTINCT SELECTCOLUMNS from vw_RA_SatReview", "yearweek (text), yearweek_sort (int)", "One side of all yearweek relationships", "Slicer source. Set Sort by Column: yearweek_sort on yearweek."]),

    (C_SEC,  ["ROW-LEVEL BASE VIEW", "", "", "", ""]),
    (C_WEEKLY, ["vw_ShortBase", "ShortMain LEFT JOIN ShortRare", "yearweek (text), yearweek_sort (int), weeknumber (int)", "--", "Row-level; source for all non-RA aggregate views"]),

    (C_SEC,  ["WEEKLY TIME-SERIES VIEWS  (connect to WeekList)", "", "", "", ""]),
    (C_WEEKLY, ["vw_WeeklySatisfaction",     "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Satisfaction and ride trends per week"]),
    (C_WEEKLY, ["vw_WeeklyNPS",              "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Snapp and Tapsi NPS scores per week"]),
    (C_WEEKLY, ["vw_IncentiveByWeek",        "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Incentive amounts and participation per week"]),
    (C_WEEKLY, ["vw_NavigationByWeek",       "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Nav app usage share per week"]),
    (C_WEEKLY, ["vw_SatisfactionByCityWeek", "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "City x week satisfaction heatmap"]),
    (C_WEEKLY, ["vw_RideShareByCityWeek",    "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "City x week ride-share breakdown"]),

    (C_SEC,  ["ROUTINE ANALYSIS (RA) VIEWS  (connect to WeekList)", "", "", "", ""]),
    (C_RA, ["vw_RA_SatReview",         "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Satisfaction and participation by week/city/coop_type"]),
    (C_RA, ["vw_RA_CitiesOverview",    "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "LOC, joint %, got-message by week/city"]),
    (C_RA, ["vw_RA_RideShare",         "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Ride volumes and share by week/city"]),
    (C_RA, ["vw_RA_PersonaPartTime",   "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Part-time % and rides per boarded by week/city"]),
    (C_RA, ["vw_RA_IncentiveAmounts",  "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Incentive range distribution by week/city"]),
    (C_RA, ["vw_RA_IncentiveDuration", "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Active-duration bucket distribution by week/city"]),
    (C_RA, ["vw_RA_Persona",           "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "All demographic dimensions in long format"]),
    (C_RA, ["vw_RA_CommFree",          "ShortMain",              "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Commission-free incentive metrics by week/city"]),
    (C_RA, ["vw_RA_CSRare",            "ShortMain + ShortRare",  "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Customer support satisfaction (rare questions)"]),
    (C_RA, ["vw_RA_NavReco",           "ShortMain + ShortRare",  "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Navigation and NPS recommendation scores"]),

    (C_SEC,  ["CITY / AGGREGATE VIEWS  (no Power BI relationship needed)", "", "", "", ""]),
    (C_CITY, ["vw_SatisfactionByCity",       "vw_ShortBase",        "--", "--", "Overall satisfaction averages by city"]),
    (C_CITY, ["vw_IncentiveByCity",          "vw_ShortBase",        "--", "--", "Incentive amounts and rates by city"]),
    (C_CITY, ["vw_IncentiveAmountByCity",    "vw_ShortBase",        "--", "--", "Incentive range distribution by city"]),
    (C_CITY, ["vw_SatisfactionByDemographics","vw_ShortBase",       "--", "--", "Satisfaction by age/gender/coop_type/driver_type"]),
    (C_CITY, ["vw_PersonaByCity",            "vw_ShortBase",        "--", "--", "Demographic breakdown by city"]),
    (C_CITY, ["vw_NavigationUsage",          "vw_ShortBase",        "--", "--", "Overall nav app usage (all-time aggregate)"]),
    (C_CITY, ["vw_Demographics",             "vw_ShortBase",        "--", "--", "Overall demographic distribution"]),
    (C_CITY, ["vw_HoneymoonEffect",          "vw_ShortBase",        "--", "--", "Satisfaction by snapp_age tenure bucket"]),
    (C_CITY, ["vw_KPISummary",               "vw_ShortBase",        "--", "--", "Single-row overall KPI card values"]),
    (C_CITY, ["vw_WideIncentiveTypes",       "WideMain",            "--", "--", "Binary incentive-type counts"]),
    (C_CITY, ["vw_WideUnsatisfactionReasons","LongMain / LongRare", "--", "--", "Open-ended dissatisfaction reason counts"]),
    (C_CITY, ["vw_LongSurveyAnswers",        "LongMain",            "--", "--", "Long-format question/answer distribution"]),
    (C_CITY, ["vw_LongRareSurveyAnswers",    "LongRare",            "--", "--", "Long-format rare question distribution"]),
    (C_CITY, ["vw_LongSurveyByCity",         "LongMain",            "--", "--", "Long-format answers broken down by city"]),
]

# ── XML helpers ────────────────────────────────────────────────────────────────
WIDTHS = [1800, 2000, 1600, 1960, 2000]  # total = 9360 DXA

def make_cell(text, width, fill, txt_color="000000", bold=False, sz=16, span=1, italic=False):
    b  = "<w:b/>"  if bold   else ""
    it = "<w:i/>"  if italic else ""
    clr = f'<w:color w:val="{txt_color}"/>' if txt_color != "000000" else ""
    gs  = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
    return (
        f'<w:tc>'
        f'<w:tcPr>{gs}'
        f'<w:tcW w:w="{width}" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        f'</w:tcBorders>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
        f'<w:tcMar>'
        f'<w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
        f'</w:tcMar>'
        f'</w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
        f'{b}{it}<w:sz w:val="{sz}"/>{clr}'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t>'
        f'</w:r></w:p>'
        f'</w:tc>'
    )

def make_row(colour_pair, cells, is_header=False, is_section=False):
    fill, txt = colour_pair
    if is_section:
        # Merged single cell across all 5 columns
        cell_xml = make_cell(cells[0], sum(WIDTHS), fill, txt,
                             bold=True, sz=16, span=5, italic=True)
    else:
        cell_xml = ""
        for i, text in enumerate(cells):
            cell_xml += make_cell(text, WIDTHS[i], fill, txt,
                                  bold=is_header, sz=17 if is_header else 16)
    return f'<w:tr><w:trPr><w:trHeight w:val="260" w:hRule="atLeast"/></w:trPr>{cell_xml}</w:tr>'

# ── Assemble table ─────────────────────────────────────────────────────────────
col_grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in WIDTHS)
header_row = make_row(C_HEAD, HEADER, is_header=True)
data_rows  = "".join(
    make_row(r[0], r[1], is_section=(r[0] == C_SEC))
    for r in ROWS
)

table_xml = (
    f'<w:tbl>'
    f'<w:tblPr>'
    f'<w:tblW w:w="{sum(WIDTHS)}" w:type="dxa"/>'
    f'<w:tblBorders>'
    f'<w:top    w:val="single" w:sz="8" w:space="0" w:color="2E75B6"/>'
    f'<w:left   w:val="single" w:sz="8" w:space="0" w:color="2E75B6"/>'
    f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="2E75B6"/>'
    f'<w:right  w:val="single" w:sz="8" w:space="0" w:color="2E75B6"/>'
    f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
    f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
    f'</w:tblBorders>'
    f'</w:tblPr>'
    f'<w:tblGrid>{col_grid}</w:tblGrid>'
    f'{header_row}{data_rows}'
    f'</w:tbl>'
)

heading_xml = (
    '<w:p>'
    '<w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="480" w:after="120"/></w:pPr>'
    '<w:r>'
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/>'
    '<w:sz w:val="32"/><w:color w:val="2E75B6"/></w:rPr>'
    '<w:t>Appendix &#x2014; Power BI Data Model: View and Relationship Schema</w:t>'
    '</w:r>'
    '</w:p>'
)

intro_xml = (
    '<w:p><w:pPr><w:spacing w:before="0" w:after="160"/></w:pPr>'
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'
    '<w:t xml:space="preserve">'
    'The table below lists every SQL view and base table used in this dashboard, '
    'what it sources from at the SQL level, whether it carries a yearweek time column, '
    'and which Power BI relationship connects it to the WeekList slicer table. '
    'All yearweek relationships are one-to-many (1:*) from WeekList[yearweek] to the '
    'view&#x2019;s [yearweek] column, with single cross-filter direction '
    '(WeekList filters the views, not the reverse). '
    'Views in the bottom section have no time dimension and are used as '
    'standalone aggregates in card or table visuals with no relationship.'
    '</w:t>'
    '</w:r></w:p>'
)

page_break_xml = (
    '<w:p>'
    '<w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0"/></w:pPr>'
    '</w:p>'
)

insert_block = page_break_xml + heading_xml + intro_xml + table_xml

import xml.etree.ElementTree as ET

# ── Inject before </w:body> ────────────────────────────────────────────────────
with open(DOC_PATH, "r", encoding="utf-8") as f:
    doc = f.read()

assert "</w:body>" in doc
doc = doc.replace("</w:body>", insert_block + "\n</w:body>")

with open(DOC_PATH, "w", encoding="utf-8") as f:
    f.write(doc)

# Final validation (full document has namespace declarations)
try:
    ET.parse(DOC_PATH)
    print("Full document.xml is well-formed. Done.")
except ET.ParseError as e:
    print(f"ERROR in final document: {e}")
    # Show context
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        lines = f.readlines()
    lineno = e.position[0]
    for i, l in enumerate(lines[max(0,lineno-3):lineno+2], start=max(0,lineno-3)+1):
        print(f"  {i}: {l.rstrip()[:120]}")
