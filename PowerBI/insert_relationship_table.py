"""Insert a Power BI relationship schema table into the guide docx XML."""

DOC_PATH = r"D:\Work\Driver Survey\PowerBI\unpacked_guide3\word\document.xml"

# ── Table data ─────────────────────────────────────────────────────────────────
# Columns: View/Table | Sources From | yearweek cols | PBI connects to | Notes

HEADER = ["View / Table", "SQL Sources From", "Time Columns", "Power BI Relationship", "Role / Notes"]

# Category colours (RRGGBB fill, text colour)
C_HEAD   = ("2E75B6", "FFFFFF")  # dark blue header
C_BASE   = ("D9D9D9", "000000")  # grey  - base SQL tables
C_DAX    = ("FCE4D6", "000000")  # orange - DAX calculated table
C_RA     = ("E2EFDA", "000000")  # green - RA views
C_WEEKLY = ("DEEAF1", "000000")  # blue  - weekly time views (non-RA)
C_CITY   = ("FFFFFF", "000000")  # white - city/aggregate, no time
C_SEC    = ("F2F2F2", "595959")  # light grey section divider row

ROWS = [
    # (fill, text, [col1, col2, col3, col4, col5])
    (C_SEC,  ["BASE SQL TABLES", "", "", "", ""]),
    (C_BASE, ["DriverSurvey_ShortMain", "—  (base table)", "yearweek INT, weeknumber INT", "—", "Primary fact table for all Short-survey views"]),
    (C_BASE, ["DriverSurvey_ShortRare",  "—  (base table)", "—", "—", "Rare-question supplement; joined to ShortMain on recordID"]),
    (C_BASE, ["DriverSurvey_WideMain",   "—  (base table)", "—", "—", "One-hot columns for incentive-type questions"]),
    (C_BASE, ["DriverSurvey_LongMain",   "—  (base table)", "—", "—", "Long-format answers for open-ended questions"]),
    (C_BASE, ["DriverSurvey_LongRare",   "—  (base table)", "—", "—", "Long-format rare answers"]),

    (C_SEC,  ["DAX CALCULATED TABLE", "", "", "", ""]),
    (C_DAX,  ["WeekList", "DAX: DISTINCT SELECTCOLUMNS from vw_RA_SatReview", "yearweek (text), yearweek_sort (int)", "One side of all yearweek relationships", "Slicer source. Set Sort by Column: yearweek_sort on yearweek."]),

    (C_SEC,  ["ROW-LEVEL BASE VIEW", "", "", "", ""]),
    (C_WEEKLY, ["vw_ShortBase", "ShortMain LEFT JOIN ShortRare", "yearweek (text), yearweek_sort (int), weeknumber (int)", "—", "Row-level; source for all non-RA aggregate views"]),

    (C_SEC,  ["WEEKLY TIME-SERIES VIEWS  (connect to WeekList)", "", "", "", ""]),
    (C_WEEKLY, ["vw_WeeklySatisfaction",    "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Satisfaction & ride trends per week"]),
    (C_WEEKLY, ["vw_WeeklyNPS",             "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Snapp & Tapsi NPS scores per week"]),
    (C_WEEKLY, ["vw_IncentiveByWeek",       "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Incentive amounts & participation per week"]),
    (C_WEEKLY, ["vw_NavigationByWeek",      "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Nav app usage share per week"]),
    (C_WEEKLY, ["vw_SatisfactionByCityWeek","vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "City x week satisfaction heatmap"]),
    (C_WEEKLY, ["vw_RideShareByCityWeek",   "vw_ShortBase", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "City x week ride-share breakdown"]),

    (C_SEC,  ["ROUTINE ANALYSIS (RA) VIEWS  (connect to WeekList)", "", "", "", ""]),
    (C_RA, ["vw_RA_SatReview",        "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Sat & participation by week/city/coop_type"]),
    (C_RA, ["vw_RA_CitiesOverview",   "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "LOC, joint %, got-message by week/city"]),
    (C_RA, ["vw_RA_RideShare",        "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Ride volumes & share by week/city"]),
    (C_RA, ["vw_RA_PersonaPartTime",  "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Part-time % & rides per boarded by week/city"]),
    (C_RA, ["vw_RA_IncentiveAmounts", "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Incentive range distribution by week/city"]),
    (C_RA, ["vw_RA_IncentiveDuration","ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Active-duration bucket distribution by week/city"]),
    (C_RA, ["vw_RA_Persona",          "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "All demographic dimensions long-format"]),
    (C_RA, ["vw_RA_CommFree",         "ShortMain", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Commission-free incentive metrics by week/city"]),
    (C_RA, ["vw_RA_CSRare",           "ShortMain + ShortRare", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Customer support satisfaction (rare questions)"]),
    (C_RA, ["vw_RA_NavReco",          "ShortMain + ShortRare", "yearweek, yearweek_sort", "WeekList[yearweek]  1:*", "Navigation & NPS recommendation scores"]),

    (C_SEC,  ["CITY / AGGREGATE VIEWS  (no Power BI relationship needed)", "", "", "", ""]),
    (C_CITY, ["vw_SatisfactionByCity",      "vw_ShortBase", "—", "—", "Overall satisfaction averages by city"]),
    (C_CITY, ["vw_IncentiveByCity",         "vw_ShortBase", "—", "—", "Incentive amounts & rates by city"]),
    (C_CITY, ["vw_IncentiveAmountByCity",   "vw_ShortBase", "—", "—", "Incentive range distribution by city"]),
    (C_CITY, ["vw_SatisfactionByDemographics","vw_ShortBase","—","—","Satisfaction by age/gender/coop_type/driver_type"]),
    (C_CITY, ["vw_PersonaByCity",           "vw_ShortBase", "—", "—", "Demographic breakdown by city"]),
    (C_CITY, ["vw_NavigationUsage",         "vw_ShortBase", "—", "—", "Overall nav app usage (all-time)"]),
    (C_CITY, ["vw_Demographics",            "vw_ShortBase", "—", "—", "Overall demographic distribution"]),
    (C_CITY, ["vw_HoneymoonEffect",         "vw_ShortBase", "—", "—", "Satisfaction by snapp_age tenure bucket"]),
    (C_CITY, ["vw_KPISummary",              "vw_ShortBase", "—", "—", "Single-row overall KPI card"]),
    (C_CITY, ["vw_WideIncentiveTypes",      "WideMain",     "—", "—", "Binary incentive-type counts"]),
    (C_CITY, ["vw_WideUnsatisfactionReasons","LongMain/LongRare","—","—","Open-ended dissatisfaction reasons"]),
    (C_CITY, ["vw_LongSurveyAnswers",       "LongMain",     "—", "—", "Long-format question/answer distribution"]),
    (C_CITY, ["vw_LongRareSurveyAnswers",   "LongRare",     "—", "—", "Long-format rare question distribution"]),
    (C_CITY, ["vw_LongSurveyByCity",        "LongMain",     "—", "—", "Long-format answers broken down by city"]),
]

# ── XML builder helpers ────────────────────────────────────────────────────────
def shd(fill, color="auto"):
    return f'<w:shd w:val="clear" w:color="{color}" w:fill="{fill}"/>'

def border(style="single", sz=4, color="BFBFBF"):
    return f'<w:top w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/><w:left w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/><w:bottom w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/><w:right w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/>'

def cell(text, width, fill, txt_color="000000", bold=False, sz=16, span_all=False, italic=False):
    b_open  = "<w:b/>"  if bold   else ""
    i_open  = "<w:i/>"  if italic else ""
    color_el = f'<w:color w:val="{txt_color}"/>' if txt_color != "000000" else ""
    grid_span = f'<w:gridSpan w:val="5"/>' if span_all else ""
    xml = f"""<w:tc>
      <w:tcPr>
        {grid_span}
        <w:tcW w:w="{width}" w:type="dxa"/>
        <w:tcBorders>{border()}</w:tcBorders>
        {shd(fill)}
        <w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="left"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{b_open}{i_open}<w:sz w:val="{sz}"/>{color_el}</w:rPr>
          <w:t xml:space="preserve">{text}</w:t>
        </w:r>
      </w:p>
    </w:tc>"""
    return xml

# Column widths (total = 9360 = 6.5" with 1" margins on letter)
WIDTHS = [1800, 2000, 1600, 1960, 2000]

def make_row(fill_txt, cells_data, is_header=False, is_section=False):
    fill, txt_color = fill_txt
    rows_xml = ""
    if is_section:
        # Single merged cell spanning all columns
        rows_xml = cell(cells_data[0], sum(WIDTHS), fill, txt_color,
                        bold=True, sz=16, span_all=True, italic=True)
    else:
        for i, text in enumerate(cells_data):
            rows_xml += cell(text, WIDTHS[i], fill, txt_color,
                             bold=is_header, sz=16 if not is_header else 17)
    return f"<w:tr><w:trPr><w:trHeight w:val='260'/></w:trPr>{rows_xml}</w:tr>"

# ── Build table XML ────────────────────────────────────────────────────────────
col_grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in WIDTHS)

tbl_borders = f"""<w:tblBorders>
    <w:top    w:val="single" w:sz="6"  w:space="0" w:color="2E75B6"/>
    <w:left   w:val="single" w:sz="6"  w:space="0" w:color="2E75B6"/>
    <w:bottom w:val="single" w:sz="6"  w:space="0" w:color="2E75B6"/>
    <w:right  w:val="single" w:sz="6"  w:space="0" w:color="2E75B6"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
</w:tblBorders>"""

header_row = make_row(C_HEAD, HEADER, is_header=True)

data_rows = ""
for entry in ROWS:
    colours = entry[0]
    cols     = entry[1]
    is_sec   = (colours == C_SEC)
    data_rows += make_row(colours, cols, is_section=is_sec)

table_xml = f"""<w:tbl>
  <w:tblPr>
    <w:tblW w:w="{sum(WIDTHS)}" w:type="dxa"/>
    {tbl_borders}
    <w:tblLook w:val="04A0"/>
  </w:tblPr>
  <w:tblGrid>{col_grid}</w:tblGrid>
  {header_row}
  {data_rows}
</w:tbl>"""

# ── Section heading paragraph ──────────────────────────────────────────────────
heading_xml = """<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:spacing w:before="480" w:after="120"/>
  </w:pPr>
  <w:r>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="32"/><w:color w:val="2E75B6"/></w:rPr>
    <w:t>Appendix &#x2014; Power BI Data Model: View &amp; Relationship Schema</w:t>
  </w:r>
</w:p>"""

intro_xml = """<w:p>
  <w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr>
  <w:r>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>
    <w:t xml:space="preserve">The table below lists every SQL view and base table, what it sources from, whether it carries a yearweek column, and which Power BI relationship connects it to the WeekList slicer table. All yearweek relationships are one-to-many (1:*) from WeekList[yearweek] to the view&#x2019;s [yearweek] column, with single cross-filter direction (WeekList filters the views, not the reverse). Views with no Power BI relationship are standalone aggregates used directly in card or table visuals.</w:t>
  </w:r>
</w:p>"""

insert_block = heading_xml + intro_xml + table_xml

# ── Inject before </w:body> ────────────────────────────────────────────────────
with open(DOC_PATH, "r", encoding="utf-8") as f:
    doc = f.read()

assert "</w:body>" in doc, "</w:body> not found"

# Add a page break before the appendix
page_break = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:lastRenderedPageBreak/><w:br w:type="page"/></w:r></w:p>'
doc = doc.replace("</w:body>", page_break + insert_block + "\n</w:body>")

with open(DOC_PATH, "w", encoding="utf-8") as f:
    f.write(doc)

print("Done — relationship schema table inserted before </w:body>")
