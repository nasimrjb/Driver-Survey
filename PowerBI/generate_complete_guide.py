import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT

doc = Document()

# ── rendering helpers ────────────────────────────────────────────────────────
def set_font(run, size=10, bold=False, color=None, name='Arial'):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading(text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.name = 'Arial'
    return p

def body(text='', bold=False, color=None, size=10):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, color=color)
    return p

def code_line(text, is_name_line=False):
    """Render one line of DAX as a code block paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.font.name = 'Courier New'
    run.font.size = Pt(8.5)
    if is_name_line:
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 70, 127)
    else:
        run.font.color.rgb = RGBColor(30, 30, 120)
    return p

def dax_measure(dax_str):
    """Render a complete DAX measure string. First line (MeasureName =) is bolded."""
    # Strip RA<n> prefix from the measure name line
    lines = dax_str.split('\n')
    lines[0] = re.sub(r'^RA\d+\s+', '', lines[0])
    dax_str = '\n'.join(lines)
    doc.add_paragraph()  # spacer before
    lines = dax_str.split('\n')
    for i, line in enumerate(lines):
        if line.strip():
            code_line(line, is_name_line=(i == 0))
    doc.add_paragraph()  # spacer after

def bullet(text, indent=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.3 + indent * 0.2)
    run = p.add_run(text)
    set_font(run, size=9.5)
    return p

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    run = p.add_run('ℹ  ' + text)
    set_font(run, size=9, color=(100, 100, 100))
    return p

def col_table(headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, h in enumerate(headers):
        cell = t.cell(0, i)
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.name = 'Arial'
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r + 1, c)
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(8.5)
            cell.paragraphs[0].runs[0].font.name = 'Courier New'
    return t

# ── DAX string builders ──────────────────────────────────────────────────────
def clean_name(name):
    """Strip 'RA<n> ' prefix so measure names in the doc match Power BI (e.g. 'Snapp NPS')."""
    return re.sub(r'^RA\d+\s+', '', name)

def yw_var(view):
    return (
        f'VAR SelYearWeek = IF(HASONEVALUE({view}[yearweek]),\n'
        f'    SELECTEDVALUE({view}[yearweek]),\n'
        f'    CALCULATE(MAX({view}[yearweek]), ALL({view})))'
    )

def metric_dax(name, view, metric_expr, n_col='n', extra_filters=None):
    """Standard metric with N-gate."""
    name = clean_name(name)
    filters = f', {view}[yearweek] = SelYearWeek'
    if extra_filters:
        filters += f',\n        {extra_filters}'
    n_filters = f'{view}[yearweek] = SelYearWeek'
    if extra_filters:
        n_filters += f',\n        {extra_filters}'
    return (
        f'{name} =\n'
        f'{yw_var(view)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({view}[{n_col}]), {n_filters}) >= MinN,\n'
        f'    CALCULATE({metric_expr}{filters}),\n'
        f'    BLANK())'
    )

def count_dax(name, view, count_col='n', extra_filters=None):
    """Simple count — no N-gate."""
    name = clean_name(name)
    n_filter = f'{view}[yearweek] = SelYearWeek'
    if extra_filters:
        n_filter += f', {extra_filters}'
    return (
        f'{name} =\n'
        f'{yw_var(view)}\n'
        f'RETURN CALCULATE(SUM({view}[{count_col}]), {n_filter})'
    )

def wow_dax(name, view, metric_expr, n_col='n', extra_filters=None):
    """WoW delta measure."""
    name = clean_name(name)
    base_filter = f'{view}[yearweek] = SelYearWeek'
    prev_filter = f'{view}[yearweek] = PrevYearWeek'
    if extra_filters:
        base_filter += f',\n        {extra_filters}'
        prev_filter += f',\n        {extra_filters}'
    n_filter = f'{view}[yearweek] = SelYearWeek'
    if extra_filters:
        n_filter += f', {extra_filters}'
    return (
        f'{name} =\n'
        f'{yw_var(view)}\n'
        f'VAR PrevYearWeek = CALCULATE(MAX({view}[yearweek]),\n'
        f'    ALL({view}), {view}[yearweek] < SelYearWeek)\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'VAR CurrVal = CALCULATE({metric_expr}, {base_filter})\n'
        f'VAR PrevVal = CALCULATE({metric_expr}, {prev_filter})\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({view}[{n_col}]), {n_filter}) >= MinN,\n'
        f'    CurrVal - PrevVal,\n'
        f'    BLANK())'
    )

# ════════════════════════════════════════════════════════════════════════════
# TITLE
# ════════════════════════════════════════════════════════════════════════════
heading('Driver Survey – Power BI Routine Analysis Guide v5', level=1)
heading('Complete DAX Reference for All 17 RA Views', level=2)
body('Database: Cab_Studies  •  Server: 192.168.18.37  •  Schema: [Cab]', color=(80, 80, 80))
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 1. SETUP
# ════════════════════════════════════════════════════════════════════════════
heading('1. Power BI Setup', level=1)
body('Connection', bold=True)
bullet('Home → Get Data → SQL Server')
bullet('Server: 192.168.18.37  |  Database: Cab_Studies')
bullet('Import mode (not DirectQuery)')
body('Import all 17 RA views', bold=True)
for v in [
    'vw_RA_SatReview', 'vw_RA_CitiesOverview', 'vw_RA_RideShare', 'vw_RA_PersonaPartTime',
    'vw_RA_IncentiveAmounts', 'vw_RA_IncentiveDuration', 'vw_RA_Persona',
    'vw_RA_CommFree', 'vw_RA_CSRare', 'vw_RA_NavReco',
    'vw_RA_IncentiveTypes', 'vw_RA_IncentiveUnsatCity', 'vw_RA_IncentiveUnsatNational',
    'vw_RA_Navigation', 'vw_RA_Referral', 'vw_RA_TapsiInactivity', 'vw_RA_LuckyWheel',
]:
    bullet(v, indent=1)
body('Do NOT create relationships between views – each is self-contained.', color=(160, 0, 0))
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 2. GLOBAL HELPERS
# ════════════════════════════════════════════════════════════════════════════
heading('2. Global Helper Measures & Parameter Table', level=1)

body('Min N Cutoff parameter table', bold=True)
note('Create a blank table via Enter Data named "Min N Cutoff" with one column "Min N Cutoff" and one row: 30. '
     'Users can edit the value in the table to change the global threshold.')

body('Helper measures (create in a "Measures" blank table):', bold=True)
dax_measure(
    'Min N Cutoff Value =\n'
    'SELECTEDVALUE(\'Min N Cutoff\'[Min N Cutoff], 0)'
)
note('All metric measures reference [Min N Cutoff Value]. Changing the table value updates all pages at once.')

body('Year-Week sorting', bold=True)
note('In the Model view, select yearweek (TEXT column) → Sort by Column → yearweek_sort (INT). '
     'Apply this to every imported view.')
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# RA-1  vw_RA_SatReview
# ════════════════════════════════════════════════════════════════════════════
heading('3. RA-1 – Satisfaction & Participation Review', level=1)
body('View: vw_RA_SatReview  |  Excel Page: #3', bold=True)
body('Incentive participation rates and satisfaction scores (1–5) for Snapp & Tapsi. '
     'Filter cooperation_type: All Drivers / Part-Time / Full-Time.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek', 'TEXT', '"26-01" style'),
        ('yearweek_sort', 'INT', 'for axis ordering'),
        ('weeknumber', 'INT', ''),
        ('city', 'TEXT', ''),
        ('cooperation_type', 'TEXT', 'All Drivers / Part-Time / Full-Time'),
        ('n', 'INT', 'all respondents in city'),
        ('n_joint', 'INT', 'joint drivers'),
        ('Part_pct_Snapp', 'FLOAT', '% participated in Snapp incentive'),
        ('Part_pct_Jnt_Snapp', 'FLOAT', '% joint participated in Snapp'),
        ('Part_pct_Jnt_Tapsi', 'FLOAT', '% joint participated in Tapsi'),
        ('Part_GotMsg_pct_Snapp', 'FLOAT', '% participated among who got Snapp msg'),
        ('Part_GotMsg_pct_Jnt_Snapp', 'FLOAT', ''),
        ('Part_GotMsg_pct_Jnt_Tapsi', 'FLOAT', ''),
        ('Incentive_Sat_Snapp', 'FLOAT', 'avg 1-5'),
        ('Incentive_Sat_Jnt_Snapp', 'FLOAT', ''),
        ('Incentive_Sat_Jnt_Tapsi', 'FLOAT', ''),
        ('Fare_Sat_Snapp / Jnt_Snapp / Jnt_Tapsi', 'FLOAT', 'avg 1-5'),
        ('Request_Sat_Snapp / Jnt_Snapp / Jnt_Tapsi', 'FLOAT', 'avg 1-5'),
        ('Income_Sat_Snapp / Jnt_Snapp / Jnt_Tapsi', 'FLOAT', 'avg 1-5'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Use a page-level filter on cooperation_type rather than a slicer to avoid cross-filtering other visuals.')

V = 'vw_RA_SatReview'
dax_measure(count_dax('RA1 N All', V))
dax_measure(count_dax('RA1 N Joint', V, count_col='n_joint'))
dax_measure(metric_dax('RA1 Part% Snapp', V, f'AVERAGE({V}[Part_pct_Snapp])'))
dax_measure(metric_dax('RA1 Part% Jnt Snapp', V, f'AVERAGE({V}[Part_pct_Jnt_Snapp])', n_col='n_joint'))
dax_measure(metric_dax('RA1 Part% Jnt Tapsi', V, f'AVERAGE({V}[Part_pct_Jnt_Tapsi])', n_col='n_joint'))
dax_measure(metric_dax('RA1 GotMsg Part% Snapp', V, f'AVERAGE({V}[Part_GotMsg_pct_Snapp])'))
dax_measure(metric_dax('RA1 Incentive Sat Snapp', V, f'AVERAGE({V}[Incentive_Sat_Snapp])'))
dax_measure(metric_dax('RA1 Incentive Sat Jnt Snapp', V, f'AVERAGE({V}[Incentive_Sat_Jnt_Snapp])', n_col='n_joint'))
dax_measure(metric_dax('RA1 Incentive Sat Jnt Tapsi', V, f'AVERAGE({V}[Incentive_Sat_Jnt_Tapsi])', n_col='n_joint'))
dax_measure(metric_dax('RA1 Fare Sat Snapp', V, f'AVERAGE({V}[Fare_Sat_Snapp])'))
dax_measure(metric_dax('RA1 Request Sat Snapp', V, f'AVERAGE({V}[Request_Sat_Snapp])'))
dax_measure(metric_dax('RA1 Income Sat Snapp', V, f'AVERAGE({V}[Income_Sat_Snapp])'))
dax_measure(wow_dax('RA1 WoW Incentive Sat Snapp', V, f'AVERAGE({V}[Incentive_Sat_Snapp])'))
note('Duplicate Incentive/Fare/Request/Income Sat measures for Jnt_Snapp and Jnt_Tapsi columns (use n_joint as MinN gate).')
note('Visual: Matrix – city × satisfaction column, week slicer. Cards for national averages. Line for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# RA-2  vw_RA_CitiesOverview
# ════════════════════════════════════════════════════════════════════════════
heading('4. RA-2 – Cities Overview', level=1)
body('View: vw_RA_CitiesOverview  |  Excel Page: #12', bold=True)
body('Three independent groups: E (all), F (joint), G (tapsi_LOC > 0). No cooperation_type split.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('E_n', 'INT', 'all respondents'),
        ('F_n', 'INT', 'joint drivers'),
        ('G_n', 'INT', 'drivers with tapsi LOC > 0'),
        ('pct_Joint', 'FLOAT', '% who are joint (base E_n)'),
        ('pct_Dual_SU', 'FLOAT', '% with tapsi LOC > 0 (base E_n)'),
        ('AvgLOC_All_Snapp', 'FLOAT', 'avg Snapp LOC (base E_n)'),
        ('GotMsg_All_Snapp', 'FLOAT', '% got Snapp msg (base E_n)'),
        ('AvgLOC_Joint_Snapp', 'FLOAT', 'avg Snapp LOC (base F_n)'),
        ('GotMsg_Joint_Snapp', 'FLOAT', '% joint got Snapp msg (base F_n)'),
        ('GotMsg_Joint_Cmpt', 'FLOAT', '% joint got Tapsi msg (base F_n)'),
        ('AvgLOC_Joint_Cmpt', 'FLOAT', 'avg Tapsi LOC (base F_n)'),
        ('AvgLOC_Joint_Cmpt_SU', 'FLOAT', 'avg Tapsi LOC (base G_n)'),
        ('GotMsg_Joint_Cmpt_SU', 'FLOAT', '% G-group got Tapsi msg (base G_n)'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_CitiesOverview'
dax_measure(count_dax('RA2 E_n', V, count_col='E_n'))
dax_measure(count_dax('RA2 F_n', V, count_col='F_n'))
dax_measure(count_dax('RA2 G_n', V, count_col='G_n'))
dax_measure(metric_dax('RA2 pct Joint', V, f'AVERAGE({V}[pct_Joint])'))
dax_measure(metric_dax('RA2 pct Dual SU', V, f'AVERAGE({V}[pct_Dual_SU])'))
dax_measure(metric_dax('RA2 AvgLOC All Snapp', V, f'AVERAGE({V}[AvgLOC_All_Snapp])'))
dax_measure(metric_dax('RA2 GotMsg All Snapp', V, f'AVERAGE({V}[GotMsg_All_Snapp])'))
dax_measure(metric_dax('RA2 GotMsg Joint Snapp', V, f'AVERAGE({V}[GotMsg_Joint_Snapp])', n_col='F_n'))
dax_measure(metric_dax('RA2 GotMsg Joint Cmpt', V, f'AVERAGE({V}[GotMsg_Joint_Cmpt])', n_col='F_n'))
dax_measure(metric_dax('RA2 GotMsg Joint Cmpt SU', V, f'AVERAGE({V}[GotMsg_Joint_Cmpt_SU])', n_col='G_n'))
dax_measure(wow_dax('RA2 WoW pct Joint', V, f'AVERAGE({V}[pct_Joint])'))
note('Use E_n / F_n / G_n as the MinN gate for corresponding metric groups.')

# ════════════════════════════════════════════════════════════════════════════
# RA-3  vw_RA_RideShare
# ════════════════════════════════════════════════════════════════════════════
heading('5. RA-3 – Ride Share', level=1)
body('View: vw_RA_RideShare  |  Excel Page: #13', bold=True)
body('Total ride counts and Snapp/Tapsi share split by driver segment.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('total_Res', 'INT', 'all respondents'),
        ('Joint_Res', 'INT', 'joint drivers'),
        ('Ex_drivers', 'INT', 'exclusive Snapp drivers'),
        ('Total_Ride', 'FLOAT', 'Snapp + Tapsi combined'),
        ('Total_Ride_Snapp', 'FLOAT', 'all Snapp rides'),
        ('Ex_Ride_Snapp', 'FLOAT', 'exclusive driver Snapp rides'),
        ('Jnt_Snapp_Ride', 'FLOAT', 'joint driver Snapp rides'),
        ('Jnt_Tapsi_Ride', 'FLOAT', 'joint driver Tapsi rides'),
        ('All_Snapp_pct', 'FLOAT', 'Snapp share of total rides'),
        ('Ex_Drivers_Snapp_pct', 'FLOAT', 'exclusive driver share'),
        ('Jnt_at_Snapp_pct', 'FLOAT', 'joint Snapp share'),
        ('Jnt_at_Tapsi_pct', 'FLOAT', 'joint Tapsi share'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_RideShare'
dax_measure(count_dax('RA3 Total Respondents', V, count_col='total_Res'))
dax_measure(count_dax('RA3 Joint Respondents', V, count_col='Joint_Res'))
dax_measure(metric_dax('RA3 All Snapp pct', V, f'AVERAGE({V}[All_Snapp_pct])', n_col='total_Res'))
dax_measure(metric_dax('RA3 Jnt at Snapp pct', V, f'AVERAGE({V}[Jnt_at_Snapp_pct])', n_col='Joint_Res'))
dax_measure(metric_dax('RA3 Jnt at Tapsi pct', V, f'AVERAGE({V}[Jnt_at_Tapsi_pct])', n_col='Joint_Res'))
dax_measure(wow_dax('RA3 WoW All Snapp pct', V, f'AVERAGE({V}[All_Snapp_pct])', n_col='total_Res'))
note('Visual: Clustered bar for pct columns by city. Line chart for WoW trend.')

# ════════════════════════════════════════════════════════════════════════════
# RA-4  vw_RA_PersonaPartTime
# ════════════════════════════════════════════════════════════════════════════
heading('6. RA-4 – Persona Part-Time', level=1)
body('View: vw_RA_PersonaPartTime  |  Excel Page: #15 (Part-Time sub-table)', bold=True)
body('Part-Time rate and average rides per boarded driver, by city.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('total_Res', 'INT', ''),
        ('Joint_Res', 'INT', ''),
        ('Ex_drivers', 'INT', ''),
        ('PT_pct_Joint', 'FLOAT', '% Part-Time among joint'),
        ('PT_pct_Exclusive', 'FLOAT', '% Part-Time among exclusive'),
        ('RidePerBoarded_Snapp', 'FLOAT', 'avg Snapp rides per joint driver'),
        ('RidePerBoarded_Tapsi', 'FLOAT', 'avg Tapsi rides per joint driver'),
        ('AvgAllRides', 'FLOAT', 'avg Snapp rides across all respondents'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_PersonaPartTime'
dax_measure(metric_dax('RA4 PT pct Joint', V, f'AVERAGE({V}[PT_pct_Joint])', n_col='Joint_Res'))
dax_measure(metric_dax('RA4 PT pct Exclusive', V, f'AVERAGE({V}[PT_pct_Exclusive])', n_col='Ex_drivers'))
dax_measure(metric_dax('RA4 RidePerBoarded Snapp', V, f'AVERAGE({V}[RidePerBoarded_Snapp])', n_col='Joint_Res'))
dax_measure(metric_dax('RA4 RidePerBoarded Tapsi', V, f'AVERAGE({V}[RidePerBoarded_Tapsi])', n_col='Joint_Res'))
dax_measure(metric_dax('RA4 AvgAllRides', V, f'AVERAGE({V}[AvgAllRides])', n_col='total_Res'))

# ════════════════════════════════════════════════════════════════════════════
# RA-5  vw_RA_IncentiveAmounts
# ════════════════════════════════════════════════════════════════════════════
heading('7. RA-5 – Incentive Amounts (Long Format)', level=1)
body('View: vw_RA_IncentiveAmounts  |  Excel Pages: #1 (Snapp), #2 (Tapsi)', bold=True)
body('Long format: one row per city × incentive_range × week. '
     'Tapsi rows include joint drivers only. Add a page-level filter on platform.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('platform', 'TEXT', 'Snapp or Tapsi'),
        ('incentive_range', 'TEXT', 'bucket label e.g. "<20k"'),
        ('incentive_range_sort', 'INT', 'for column ordering'),
        ('n_range', 'INT', 'count in this bucket'),
        ('n_total', 'INT', 'total respondents for city × week'),
        ('pct', 'FLOAT', 'n_range / n_total × 100'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Sort incentive_range by incentive_range_sort in the Model view.')

V = 'vw_RA_IncentiveAmounts'

dax_measure(
    f'RA5 n Total (Snapp) =\n'
    f'{yw_var(V)}\n'
    f'RETURN CALCULATE(MAX({V}[n_total]),\n'
    f'    {V}[yearweek] = SelYearWeek,\n'
    f'    {V}[platform] = "Snapp")'
)
dax_measure(
    f'RA5 pct Range (Snapp) =\n'
    f'// Place on a Matrix: rows = city, columns = incentive_range, values = this measure\n'
    f'// Page filter: platform = "Snapp"\n'
    f'{yw_var(V)}\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Snapp") >= MinN,\n'
    f'    CALCULATE(SUM({V}[pct]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Snapp"),\n'
    f'    BLANK())'
)
dax_measure(
    f'RA5 pct Range (Tapsi) =\n'
    f'// Page filter: platform = "Tapsi"\n'
    f'{yw_var(V)}\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Tapsi") >= MinN,\n'
    f'    CALCULATE(SUM({V}[pct]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Tapsi"),\n'
    f'    BLANK())'
)
note('Visual: Matrix – city on rows, incentive_range on columns. Separate pages for Snapp and Tapsi.')

# ════════════════════════════════════════════════════════════════════════════
# RA-6  vw_RA_IncentiveDuration
# ════════════════════════════════════════════════════════════════════════════
heading('8. RA-6 – Incentive Duration (Long Format)', level=1)
body('View: vw_RA_IncentiveDuration  |  Excel Page: #4', bold=True)
body('How long drivers have had an active incentive. Snapp (all) and Tapsi (joint only) rows.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('platform', 'TEXT', 'Snapp or Tapsi'),
        ('duration_bucket', 'TEXT', '"Few Hours","1 Day","1_6 Days","7 Days",">7 Days"'),
        ('duration_bucket_sort', 'INT', '1–5 + 99'),
        ('n_range', 'INT', ''),
        ('n_total', 'INT', ''),
        ('pct', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Sort duration_bucket by duration_bucket_sort in the Model view.')

V = 'vw_RA_IncentiveDuration'
dax_measure(
    f'RA6 n Total =\n'
    f'{yw_var(V)}\n'
    f'RETURN CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek)'
)
dax_measure(
    f'RA6 pct Bucket =\n'
    f'// Matrix: rows = city, columns = duration_bucket; add platform slicer or page filter\n'
    f'{yw_var(V)}\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek) >= MinN,\n'
    f'    CALCULATE(SUM({V}[pct]), {V}[yearweek] = SelYearWeek),\n'
    f'    BLANK())'
)

# ════════════════════════════════════════════════════════════════════════════
# RA-7  vw_RA_Persona
# ════════════════════════════════════════════════════════════════════════════
heading('9. RA-7 – Persona (Long Format)', level=1)
body('View: vw_RA_Persona  |  Excel Page: #15 (all demographic sub-tables)', bold=True)
body('Dimensions: Activity Type, Age Group, Education, Marital Status, Gender, Cooperation Type. '
     'Use a slicer on dimension to switch between them.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('dimension', 'TEXT', '"Activity Type", "Age Group", etc.'),
        ('category', 'TEXT', 'bucket value within the dimension'),
        ('category_sort', 'INT', 'for axis ordering'),
        ('n', 'INT', 'count in this category'),
        ('n_total', 'INT', 'total respondents for city × week × dimension'),
        ('pct', 'FLOAT', 'n / n_total × 100'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Sort category by category_sort in the Model view.')

V = 'vw_RA_Persona'
dax_measure(
    f'RA7 n Total =\n'
    f'{yw_var(V)}\n'
    f'RETURN CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek)'
)
dax_measure(
    f'RA7 pct Category =\n'
    f'// Matrix: rows = city, columns = category; dimension slicer selects which breakdown\n'
    f'{yw_var(V)}\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek) >= MinN,\n'
    f'    CALCULATE(SUM({V}[pct]), {V}[yearweek] = SelYearWeek),\n'
    f'    BLANK())'
)

# ════════════════════════════════════════════════════════════════════════════
# RA-8  vw_RA_CommFree
# ════════════════════════════════════════════════════════════════════════════
heading('10. RA-8 – Commission-Free Incentive', level=1)
body('View: vw_RA_CommFree  |  Excel Page: #18 (Snapp + Tapsi)', bold=True)
body('UNION ALL of Snapp (all drivers) and Tapsi (joint only). '
     'Incentive-type binary flags joined from WideMain. '
     'Hardcode platform filter in each measure rather than using a slicer.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city / platform', '', 'Snapp or Tapsi'),
        ('n', 'INT', 'base respondents'),
        ('Who_Got_Message', 'INT', 'received incentive message'),
        ('GotMsg_Money', 'INT', 'msg received, category = Money'),
        ('GotMsg_FreeComm', 'INT', 'msg received, category = Free-Commission'),
        ('GotMsg_Money_FreeComm', 'INT', 'msg received, category = Money & Free-commission'),
        ('GotMsg_PayRide', 'INT', 'received msg + Pay After Ride type'),
        ('GotMsg_EarnCF', 'INT', 'received msg + Earning-Based CF type'),
        ('GotMsg_RideCF', 'INT', 'received msg + Ride-Based CF type'),
        ('GotMsg_IncGuar', 'INT', 'received msg + Income Guarantee type'),
        ('GotMsg_PayInc', 'INT', 'received msg + Pay After Income type'),
        ('GotMsg_CFSome', 'INT', 'received msg + CF Some Trips type'),
        ('Free_Comm_Drivers', 'INT', 'drivers with CF rides > 0'),
        ('Participated', 'INT', 'participated in incentive'),
        ('pct_Got_Message', 'FLOAT', '% of n'),
        ('pct_Free_Comm_Ride', 'FLOAT', '% of n'),
        ('pct_Participated', 'FLOAT', '% of Who_Got_Message'),
        ('Avg_CF_Rides', 'FLOAT', 'avg CF rides among CF drivers only'),
        ('Avg_Total_Rides', 'FLOAT', 'avg total rides'),
        ('Avg_pct_CF_RideShare', 'FLOAT', 'avg CF% of total rides'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_CommFree'
for plat in ('Snapp', 'Tapsi'):
    dax_measure(
        f'RA8 n ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'RETURN CALCULATE(SUM({V}[n]),\n'
        f'    {V}[yearweek] = SelYearWeek,\n'
        f'    {V}[platform] = "{plat}")'
    )
    dax_measure(
        f'RA8 pct Got Message ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({V}[n]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}") >= MinN,\n'
        f'    CALCULATE(AVERAGE({V}[pct_Got_Message]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}"),\n'
        f'    BLANK())'
    )
    dax_measure(
        f'RA8 pct Free Comm Ride ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({V}[n]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}") >= MinN,\n'
        f'    CALCULATE(AVERAGE({V}[pct_Free_Comm_Ride]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}"),\n'
        f'    BLANK())'
    )
    dax_measure(
        f'RA8 pct Participated ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({V}[n]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}") >= MinN,\n'
        f'    CALCULATE(AVERAGE({V}[pct_Participated]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}"),\n'
        f'    BLANK())'
    )

for plat, type_cols in [
    ('Snapp', ['GotMsg_PayRide', 'GotMsg_EarnCF', 'GotMsg_RideCF',
               'GotMsg_IncGuar', 'GotMsg_PayInc', 'GotMsg_CFSome']),
    ('Tapsi', ['GotMsg_PayRide', 'GotMsg_EarnCF', 'GotMsg_RideCF',
               'GotMsg_IncGuar', 'GotMsg_PayInc', 'GotMsg_CFSome']),
]:
    dax_measure(
        f'RA8 GotMsg by Type% ({plat}) =\n'
        f'// Divide each GotMsg_* count by Who_Got_Message for the type breakdown\n'
        f'{yw_var(V)}\n'
        f'VAR WhoGot = CALCULATE(SUM({V}[Who_Got_Message]),\n'
        f'    {V}[yearweek] = SelYearWeek, {V}[platform] = "{plat}")\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'// Replace GotMsg_PayRide with whichever type column you need:\n'
        f'VAR NType  = CALCULATE(SUM({V}[GotMsg_PayRide]),\n'
        f'    {V}[yearweek] = SelYearWeek, {V}[platform] = "{plat}")\n'
        f'RETURN IF(WhoGot >= MinN, DIVIDE(NType, WhoGot) * 100, BLANK())'
    )
    dax_measure(
        f'RA8 Avg CF Rides ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(SUM({V}[n]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}") >= MinN,\n'
        f'    CALCULATE(AVERAGE({V}[Avg_CF_Rides]),\n'
        f'        {V}[yearweek] = SelYearWeek,\n'
        f'        {V}[platform] = "{plat}"),\n'
        f'    BLANK())'
    )

dax_measure(
    f'RA8 WoW pct Got Message (Snapp) =\n'
    f'{yw_var(V)}\n'
    f'VAR PrevYearWeek = CALCULATE(MAX({V}[yearweek]),\n'
    f'    ALL({V}), {V}[yearweek] < SelYearWeek)\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'VAR CurrVal = CALCULATE(AVERAGE({V}[pct_Got_Message]),\n'
    f'    {V}[yearweek] = SelYearWeek, {V}[platform] = "Snapp")\n'
    f'VAR PrevVal = CALCULATE(AVERAGE({V}[pct_Got_Message]),\n'
    f'    {V}[yearweek] = PrevYearWeek, {V}[platform] = "Snapp")\n'
    f'RETURN IF(\n'
    f'    CALCULATE(SUM({V}[n]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Snapp") >= MinN,\n'
    f'    CurrVal - PrevVal,\n'
    f'    BLANK())'
)
note('Duplicate WoW pattern for pct_Free_Comm_Ride, pct_Participated, Avg_CF_Rides. Apply same pattern to Tapsi.')
note('Visual: Matrix with city rows. Cards for national pct_Got_Message / pct_Free_Comm_Ride.')

# ════════════════════════════════════════════════════════════════════════════
# RA-9  vw_RA_CSRare
# ════════════════════════════════════════════════════════════════════════════
heading('11. RA-9 – Customer Support Satisfaction', level=1)
body('View: vw_RA_CSRare  |  Excel Pages: CS_Sat_Snapp / CS_Sat_Tapsi', bold=True)
body('ShortRare joined to ShortMain. All scores are 1–5 averages.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n', 'INT', 'respondents with ShortRare record'),
        ('Snapp_CS_Overall', 'FLOAT', 'avg overall CS satisfaction 1-5'),
        ('Snapp_CS_WaitTime', 'FLOAT', ''),
        ('Snapp_CS_Solution', 'FLOAT', ''),
        ('Snapp_CS_Behaviour', 'FLOAT', ''),
        ('Snapp_CS_Relevance', 'FLOAT', ''),
        ('Snapp_CS_Solved_pct', 'FLOAT', '% whose issue was solved'),
        ('Tapsi_CS_Overall', 'FLOAT', ''),
        ('Tapsi_CS_WaitTime', 'FLOAT', ''),
        ('Tapsi_CS_Solution', 'FLOAT', ''),
        ('Tapsi_CS_Behaviour', 'FLOAT', ''),
        ('Tapsi_CS_Relevance', 'FLOAT', ''),
        ('Tapsi_CS_Solved_pct', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_CSRare'
dax_measure(count_dax('RA9 n', V))
for col in ['Snapp_CS_Overall', 'Snapp_CS_WaitTime', 'Snapp_CS_Solution',
            'Snapp_CS_Behaviour', 'Snapp_CS_Relevance']:
    short = col.replace('Snapp_CS_', '').replace('_', ' ')
    dax_measure(metric_dax(f'RA9 Snapp CS {short}', V, f'AVERAGE({V}[{col}])'))
dax_measure(metric_dax('RA9 Snapp CS Solved pct', V, f'AVERAGE({V}[Snapp_CS_Solved_pct])'))
dax_measure(metric_dax('RA9 Tapsi CS Overall', V, f'AVERAGE({V}[Tapsi_CS_Overall])'))
dax_measure(wow_dax('RA9 WoW Snapp CS Overall', V, f'AVERAGE({V}[Snapp_CS_Overall])'))
note('Duplicate all five Snapp score measures + Solved_pct for Tapsi columns.')
note('Visual: Matrix city × score column. Line chart for WoW trend.')

# ════════════════════════════════════════════════════════════════════════════
# RA-10  vw_RA_NavReco
# ════════════════════════════════════════════════════════════════════════════
heading('12. RA-10 – Navigation & NPS Recommendation Scores', level=1)
body('View: vw_RA_NavReco  |  Excel Pages: NavReco_Scores / Reco_NPS', bold=True)
body('From ShortRare. NPS scores and navigation-app recommendation scores (0–10).')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n', 'INT', ''),
        ('Snapp_NPS', 'FLOAT', 'avg recommend-Snapp score'),
        ('Tapsi_NPS_SnapDriver', 'FLOAT', 'avg recommend-Tapsi (Snapp driver respondent)'),
        ('Tapsi_NPS_TapsiDriver', 'FLOAT', 'avg recommend-Tapsi (Tapsi driver respondent)'),
        ('Reco_GoogleMap', 'FLOAT', 'avg recommendation score'),
        ('Reco_Waze', 'FLOAT', ''),
        ('Reco_Neshan', 'FLOAT', ''),
        ('Reco_Balad', 'FLOAT', ''),
        ('Snapp_Nav_Sat', 'FLOAT', 'avg Snapp nav app satisfaction'),
        ('Tapsi_Nav_Sat', 'FLOAT', 'avg Tapsi in-app nav satisfaction'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_NavReco'
dax_measure(count_dax('RA10 n', V))
for col, label in [
    ('Snapp_NPS', 'Snapp NPS'),
    ('Tapsi_NPS_SnapDriver', 'Tapsi NPS (Snapp Driver)'),
    ('Tapsi_NPS_TapsiDriver', 'Tapsi NPS (Tapsi Driver)'),
    ('Reco_Neshan', 'Reco Neshan'),
    ('Reco_Balad', 'Reco Balad'),
    ('Reco_GoogleMap', 'Reco GoogleMap'),
    ('Reco_Waze', 'Reco Waze'),
    ('Snapp_Nav_Sat', 'Snapp Nav Sat'),
    ('Tapsi_Nav_Sat', 'Tapsi Nav Sat'),
]:
    dax_measure(metric_dax(f'RA10 {label}', V, f'AVERAGE({V}[{col}])'))
dax_measure(wow_dax('RA10 WoW Snapp NPS', V, f'AVERAGE({V}[Snapp_NPS])'))

# ════════════════════════════════════════════════════════════════════════════
# RA-11  vw_RA_IncentiveTypes
# ════════════════════════════════════════════════════════════════════════════
heading('13. RA-11 – Incentive Type Distribution', level=1)
body('View: vw_RA_IncentiveTypes  |  Excel Pages: #5 (Snapp Excl), #6 (Joint)', bold=True)
body('Multi-select incentive type flags (6 types × 3 segments). Base: n_excl for Excl metrics, n_joint for Joint.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n / n_joint / n_excl', 'INT', ''),
        ('pct_GotMsg_Excl_Snapp', 'FLOAT', '% excl drivers who got Snapp msg'),
        ('pct_GotMsg_Jnt_Snapp', 'FLOAT', '% joint who got Snapp msg'),
        ('pct_GotMsg_Jnt_Tapsi', 'FLOAT', '% joint who got Tapsi msg'),
        ('pct_GotMsg_Both', 'FLOAT', '% joint who got both msgs'),
        ('pct_GotMsg_Diff', 'FLOAT', '% joint who got only one msg'),
        ('pct_PayRide_Excl / JntSn / JntTp', 'FLOAT', 'Pay After Ride type %'),
        ('pct_EarnCF_Excl / JntSn / JntTp', 'FLOAT', 'Earning-Based CF type %'),
        ('pct_RideCF_Excl / JntSn / JntTp', 'FLOAT', 'Ride-Based CF type %'),
        ('pct_IncGuar_Excl / JntSn / JntTp', 'FLOAT', 'Income Guarantee type %'),
        ('pct_PayInc_Excl / JntSn / JntTp', 'FLOAT', 'Pay After Income type %'),
        ('pct_CFSome_Excl / JntSn / JntTp', 'FLOAT', 'CF on Some Trips type %'),
        ('Avg_CF_Rides_Snapp', 'FLOAT', 'avg CF rides (CF drivers only)'),
        ('Avg_CF_Rides_Tapsi', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_IncentiveTypes'
dax_measure(count_dax('RA11 n Excl', V, count_col='n_excl'))
dax_measure(count_dax('RA11 n Joint', V, count_col='n_joint'))
dax_measure(metric_dax('RA11 pct GotMsg Excl Snapp', V, f'AVERAGE({V}[pct_GotMsg_Excl_Snapp])', n_col='n_excl'))
dax_measure(metric_dax('RA11 pct GotMsg Jnt Snapp', V, f'AVERAGE({V}[pct_GotMsg_Jnt_Snapp])', n_col='n_joint'))
dax_measure(metric_dax('RA11 pct GotMsg Jnt Tapsi', V, f'AVERAGE({V}[pct_GotMsg_Jnt_Tapsi])', n_col='n_joint'))
dax_measure(metric_dax('RA11 pct GotMsg Both', V, f'AVERAGE({V}[pct_GotMsg_Both])', n_col='n_joint'))
dax_measure(metric_dax('RA11 pct GotMsg Diff', V, f'AVERAGE({V}[pct_GotMsg_Diff])', n_col='n_joint'))

for seg, base in [('Excl', 'n_excl'), ('JntSn', 'n_joint'), ('JntTp', 'n_joint')]:
    for t in ['PayRide', 'EarnCF', 'RideCF', 'IncGuar', 'PayInc', 'CFSome']:
        col = f'pct_{t}_{seg}'
        dax_measure(metric_dax(f'RA11 {col}', V, f'AVERAGE({V}[{col}])', n_col=base))

dax_measure(metric_dax('RA11 Avg CF Rides Snapp', V, f'AVERAGE({V}[Avg_CF_Rides_Snapp])'))
dax_measure(metric_dax('RA11 Avg CF Rides Tapsi', V, f'AVERAGE({V}[Avg_CF_Rides_Tapsi])', n_col='n_joint'))

# ════════════════════════════════════════════════════════════════════════════
# RA-12  vw_RA_IncentiveUnsatCity
# ════════════════════════════════════════════════════════════════════════════
heading('14. RA-12 – Incentive Dissatisfaction by City', level=1)
body('View: vw_RA_IncentiveUnsatCity  |  Excel Page: #8', bold=True)
body('"Low sat" = driver cited ANY unsatisfaction reason (multi-select from WideMain). '
     'Snapp base = n_sn_low_sat; Tapsi base = n_tp_low_sat (joint only).')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n_all', 'INT', 'all respondents'),
        ('n_joint', 'INT', 'joint drivers'),
        ('n_sn_low_sat', 'INT', 'Snapp dissatisfied (cited any reason)'),
        ('n_tp_low_sat', 'INT', 'Tapsi dissatisfied joint drivers'),
        ('pct_Sn_NoTime', 'FLOAT', '% of n_sn_low_sat citing "Not Available"'),
        ('pct_Sn_ImpAmt', 'FLOAT', '% citing "Improper Amount"'),
        ('pct_Sn_LowTime', 'FLOAT', '% citing "No Time todo"'),
        ('pct_Sn_HardToDo', 'FLOAT', '% citing "difficult"'),
        ('pct_Sn_NonPay', 'FLOAT', '% citing "Non Payment"'),
        ('pct_Tp_NoTime', 'FLOAT', '% of n_tp_low_sat citing each Tapsi reason'),
        ('pct_Tp_ImpAmt / pct_Tp_LowTime / pct_Tp_HardToDo / pct_Tp_NonPay', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_IncentiveUnsatCity'
dax_measure(count_dax('RA12 n All', V, count_col='n_all'))
dax_measure(count_dax('RA12 n Sn Low Sat', V, count_col='n_sn_low_sat'))
dax_measure(count_dax('RA12 n Tp Low Sat', V, count_col='n_tp_low_sat'))
for col, label in [
    ('pct_Sn_NoTime', 'pct Sn NoTime'),
    ('pct_Sn_ImpAmt', 'pct Sn ImpAmt'),
    ('pct_Sn_LowTime', 'pct Sn LowTime'),
    ('pct_Sn_HardToDo', 'pct Sn HardToDo'),
    ('pct_Sn_NonPay', 'pct Sn NonPay'),
]:
    dax_measure(metric_dax(f'RA12 {label}', V, f'AVERAGE({V}[{col}])', n_col='n_sn_low_sat'))
for col, label in [
    ('pct_Tp_NoTime', 'pct Tp NoTime'),
    ('pct_Tp_ImpAmt', 'pct Tp ImpAmt'),
    ('pct_Tp_LowTime', 'pct Tp LowTime'),
    ('pct_Tp_HardToDo', 'pct Tp HardToDo'),
    ('pct_Tp_NonPay', 'pct Tp NonPay'),
]:
    dax_measure(metric_dax(f'RA12 {label}', V, f'AVERAGE({V}[{col}])', n_col='n_tp_low_sat'))

# ════════════════════════════════════════════════════════════════════════════
# RA-13  vw_RA_IncentiveUnsatNational
# ════════════════════════════════════════════════════════════════════════════
heading('15. RA-13 – Incentive Dissatisfaction (National)', level=1)
body('View: vw_RA_IncentiveUnsatNational  |  Excel Page: #9', bold=True)
body('Long format – no city column. Segments: All Snapp / Joint Snapp / Joint Tapsi.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber', '', ''),
        ('segment', 'TEXT', '"All Snapp", "Joint Snapp", "Joint Tapsi"'),
        ('segment_sort', 'INT', '1, 2, 3'),
        ('n', 'INT', 'total respondents in segment'),
        ('n_low_sat', 'INT', 'dissatisfied drivers in segment'),
        ('pct_NoTime', 'FLOAT', '% of n_low_sat citing "Not Available"'),
        ('pct_ImpAmt', 'FLOAT', '% citing "Improper Amount"'),
        ('pct_LowTime', 'FLOAT', '% citing "No Time todo"'),
        ('pct_HardToDo', 'FLOAT', '% citing "difficult"'),
        ('pct_NonPay', 'FLOAT', '% citing "Non Payment"'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Create all five pct measures × three segments = 15 measures. Template below; replace segment name.')

V = 'vw_RA_IncentiveUnsatNational'
for seg in ['All Snapp', 'Joint Snapp', 'Joint Tapsi']:
    seg_var = seg.replace(' ', '')
    for col, label in [
        ('pct_NoTime', 'NoTime'),
        ('pct_ImpAmt', 'ImpAmt'),
        ('pct_LowTime', 'LowTime'),
        ('pct_HardToDo', 'HardToDo'),
        ('pct_NonPay', 'NonPay'),
    ]:
        dax_measure(
            f'RA13 {label} ({seg}) =\n'
            f'{yw_var(V)}\n'
            f'VAR MinN = [Min N Cutoff Value]\n'
            f'RETURN IF(\n'
            f'    CALCULATE(SUM({V}[n_low_sat]),\n'
            f'        {V}[yearweek] = SelYearWeek,\n'
            f'        {V}[segment] = "{seg}") >= MinN,\n'
            f'    CALCULATE(AVERAGE({V}[{col}]),\n'
            f'        {V}[yearweek] = SelYearWeek,\n'
            f'        {V}[segment] = "{seg}"),\n'
            f'    BLANK())'
        )

dax_measure(
    f'RA13 WoW pct NoTime (All Snapp) =\n'
    f'{yw_var(V)}\n'
    f'VAR PrevYearWeek = CALCULATE(MAX({V}[yearweek]),\n'
    f'    ALL({V}), {V}[yearweek] < SelYearWeek)\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'VAR CurrVal = CALCULATE(AVERAGE({V}[pct_NoTime]),\n'
    f'    {V}[yearweek] = SelYearWeek, {V}[segment] = "All Snapp")\n'
    f'VAR PrevVal = CALCULATE(AVERAGE({V}[pct_NoTime]),\n'
    f'    {V}[yearweek] = PrevYearWeek, {V}[segment] = "All Snapp")\n'
    f'RETURN IF(\n'
    f'    CALCULATE(SUM({V}[n_low_sat]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[segment] = "All Snapp") >= MinN,\n'
    f'    CurrVal - PrevVal,\n'
    f'    BLANK())'
)
note('Visual: Clustered bar – reason on x-axis, segment as legend. Line for WoW trend per reason.')

# ════════════════════════════════════════════════════════════════════════════
# RA-14  vw_RA_Navigation
# ════════════════════════════════════════════════════════════════════════════
heading('16. RA-14 – Navigation App Usage by City', level=1)
body('View: vw_RA_Navigation  |  Excel Page: #14', bold=True)
body('UNION ALL Snapp (all) + Tapsi (joint only). '
     'Snapp has GoogleMap & Waze; Tapsi has InAppNav (NULL for the other platform).')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city / platform', '', 'Snapp or Tapsi'),
        ('n', 'INT', 'respondents with non-null navigation answer'),
        ('pct_Neshan', 'FLOAT', ''),
        ('pct_Balad', 'FLOAT', ''),
        ('pct_None', 'FLOAT', '% No Navigation App'),
        ('pct_GoogleMap', 'FLOAT', 'Snapp only (NULL for Tapsi)'),
        ('pct_Waze', 'FLOAT', 'Snapp only (NULL for Tapsi)'),
        ('pct_InAppNav', 'FLOAT', 'Tapsi only (NULL for Snapp)'),
        ('pct_Other', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_Navigation'
for plat, cols in [
    ('Snapp', ['pct_Neshan', 'pct_Balad', 'pct_None', 'pct_GoogleMap', 'pct_Waze', 'pct_Other']),
    ('Tapsi', ['pct_Neshan', 'pct_Balad', 'pct_None', 'pct_InAppNav', 'pct_Other']),
]:
    dax_measure(
        f'RA14 n ({plat}) =\n'
        f'{yw_var(V)}\n'
        f'RETURN CALCULATE(SUM({V}[n]),\n'
        f'    {V}[yearweek] = SelYearWeek, {V}[platform] = "{plat}")'
    )
    for col in cols:
        label = col.replace('pct_', '')
        dax_measure(
            f'RA14 {label} ({plat}) =\n'
            f'{yw_var(V)}\n'
            f'VAR MinN = [Min N Cutoff Value]\n'
            f'RETURN IF(\n'
            f'    CALCULATE(SUM({V}[n]),\n'
            f'        {V}[yearweek] = SelYearWeek,\n'
            f'        {V}[platform] = "{plat}") >= MinN,\n'
            f'    CALCULATE(AVERAGE({V}[{col}]),\n'
            f'        {V}[yearweek] = SelYearWeek,\n'
            f'        {V}[platform] = "{plat}"),\n'
            f'    BLANK())'
        )

dax_measure(
    f'RA14 WoW pct Neshan (Snapp) =\n'
    f'{yw_var(V)}\n'
    f'VAR PrevYearWeek = CALCULATE(MAX({V}[yearweek]),\n'
    f'    ALL({V}), {V}[yearweek] < SelYearWeek)\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'VAR CurrVal = CALCULATE(AVERAGE({V}[pct_Neshan]),\n'
    f'    {V}[yearweek] = SelYearWeek, {V}[platform] = "Snapp")\n'
    f'VAR PrevVal = CALCULATE(AVERAGE({V}[pct_Neshan]),\n'
    f'    {V}[yearweek] = PrevYearWeek, {V}[platform] = "Snapp")\n'
    f'RETURN IF(\n'
    f'    CALCULATE(SUM({V}[n]),\n'
    f'        {V}[yearweek] = SelYearWeek,\n'
    f'        {V}[platform] = "Snapp") >= MinN,\n'
    f'    CurrVal - PrevVal,\n'
    f'    BLANK())'
)
note('Visual: Stacked bar – navigation app as legend, city on axis. Separate Snapp / Tapsi pages.')

# ════════════════════════════════════════════════════════════════════════════
# RA-15  vw_RA_Referral
# ════════════════════════════════════════════════════════════════════════════
heading('17. RA-15 – Referral / Joining Bonus', level=1)
body('View: vw_RA_Referral  |  Excel Page: #16', bold=True)
body('Snapp joining bonus: all drivers. Tapsi joining bonus: joint drivers only.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n_Snapp', 'INT', 'respondents with non-null Snapp joining_bonus'),
        ('joining_Snapp', 'INT', 'count who got Snapp bonus'),
        ('pct_Joining_Snapp', 'FLOAT', ''),
        ('n_Tapsi', 'INT', 'joint respondents with non-null Tapsi joining_bonus'),
        ('joining_Tapsi', 'INT', 'count who got Tapsi bonus'),
        ('pct_Joining_Tapsi', 'FLOAT', ''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_Referral'
dax_measure(count_dax('RA15 n Snapp', V, count_col='n_Snapp'))
dax_measure(count_dax('RA15 n Tapsi', V, count_col='n_Tapsi'))
dax_measure(metric_dax('RA15 pct Joining Snapp', V, f'AVERAGE({V}[pct_Joining_Snapp])', n_col='n_Snapp'))
dax_measure(metric_dax('RA15 pct Joining Tapsi', V, f'AVERAGE({V}[pct_Joining_Tapsi])', n_col='n_Tapsi'))
dax_measure(wow_dax('RA15 WoW pct Joining Snapp', V, f'AVERAGE({V}[pct_Joining_Snapp])', n_col='n_Snapp'))
dax_measure(wow_dax('RA15 WoW pct Joining Tapsi', V, f'AVERAGE({V}[pct_Joining_Tapsi])', n_col='n_Tapsi'))
note('Visual: Matrix with city rows. Cards for national joining pct. Line chart for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# RA-16  vw_RA_TapsiInactivity
# ════════════════════════════════════════════════════════════════════════════
heading('18. RA-16 – Tapsi Inactivity Before Incentive', level=1)
body('View: vw_RA_TapsiInactivity  |  Excel Page: #17', bold=True)
body('Joint drivers only. Long format with inactivity time bucket before receiving the incentive.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('inactivity_bucket', 'TEXT',
         '"Same Day", "1_3 Day Before", "3_7 Days Before", '
         '"8_14 Days Before", "15_30 Days_Before", "1_2 Month Before", '
         '"2_3 Month Before", "3_6Month Before", ">6 Month Before"'),
        ('bucket_sort', 'INT', '1–9 + 99'),
        ('n', 'INT', 'count in this bucket'),
        ('n_total', 'INT', 'total joint drivers in city × week'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
note('Sort inactivity_bucket by bucket_sort in the Model view.')

V = 'vw_RA_TapsiInactivity'
dax_measure(
    f'RA16 n Total =\n'
    f'{yw_var(V)}\n'
    f'RETURN CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek)'
)
dax_measure(
    f'RA16 pct Bucket =\n'
    f'// Matrix: rows = city, columns = inactivity_bucket (sorted by bucket_sort)\n'
    f'{yw_var(V)}\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek) >= MinN,\n'
    f'    DIVIDE(\n'
    f'        CALCULATE(SUM({V}[n]), {V}[yearweek] = SelYearWeek),\n'
    f'        CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek)) * 100,\n'
    f'    BLANK())'
)

for bucket in ['Same Day', '1_3 Day Before', '>6 Month Before']:
    safe = bucket.replace(' ', '_').replace('>', 'gt')
    dax_measure(
        f'RA16 pct {bucket} =\n'
        f'{yw_var(V)}\n'
        f'VAR MinN = [Min N Cutoff Value]\n'
        f'RETURN IF(\n'
        f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek) >= MinN,\n'
        f'    DIVIDE(\n'
        f'        CALCULATE(SUM({V}[n]),\n'
        f'            {V}[yearweek] = SelYearWeek,\n'
        f'            {V}[inactivity_bucket] = "{bucket}"),\n'
        f'        CALCULATE(MAX({V}[n_total]),\n'
        f'            {V}[yearweek] = SelYearWeek)) * 100,\n'
        f'    BLANK())'
    )

dax_measure(
    f'RA16 WoW pct Same Day =\n'
    f'{yw_var(V)}\n'
    f'VAR PrevYearWeek = CALCULATE(MAX({V}[yearweek]),\n'
    f'    ALL({V}), {V}[yearweek] < SelYearWeek)\n'
    f'VAR MinN = [Min N Cutoff Value]\n'
    f'VAR CurrVal = DIVIDE(\n'
    f'    CALCULATE(SUM({V}[n]),\n'
    f'        {V}[yearweek] = SelYearWeek, {V}[inactivity_bucket] = "Same Day"),\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek)) * 100\n'
    f'VAR PrevVal = DIVIDE(\n'
    f'    CALCULATE(SUM({V}[n]),\n'
    f'        {V}[yearweek] = PrevYearWeek, {V}[inactivity_bucket] = "Same Day"),\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = PrevYearWeek)) * 100\n'
    f'RETURN IF(\n'
    f'    CALCULATE(MAX({V}[n_total]), {V}[yearweek] = SelYearWeek) >= MinN,\n'
    f'    CurrVal - PrevVal,\n'
    f'    BLANK())'
)
note('Visual: Clustered/Stacked bar – inactivity_bucket on x-axis, city as legend or matrix rows.')

# ════════════════════════════════════════════════════════════════════════════
# RA-17  vw_RA_LuckyWheel
# ════════════════════════════════════════════════════════════════════════════
heading('19. RA-17 – Lucky Wheel Usage', level=1)
body('View: vw_RA_LuckyWheel  |  Excel Page: #19', bold=True)
body('wheel column = Rial amount won; 0 = did not use.')
body('Columns:', bold=True)
col_table(
    ['Column', 'Type', 'Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city', '', ''),
        ('n', 'INT', 'all respondents'),
        ('n_users', 'INT', 'drivers who used Lucky Wheel (wheel > 0)'),
        ('pct_usage', 'FLOAT', '% who used the wheel'),
        ('avg_wheel_amount', 'FLOAT', 'avg Rial amount among users only'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)

V = 'vw_RA_LuckyWheel'
dax_measure(count_dax('RA17 n', V))
dax_measure(count_dax('RA17 n Users', V, count_col='n_users'))
dax_measure(metric_dax('RA17 pct Usage', V, f'AVERAGE({V}[pct_usage])'))
dax_measure(metric_dax('RA17 Avg Wheel Amount', V, f'AVERAGE({V}[avg_wheel_amount])', n_col='n_users'))
dax_measure(wow_dax('RA17 WoW pct Usage', V, f'AVERAGE({V}[pct_usage])'))
dax_measure(wow_dax('RA17 WoW Avg Wheel Amount', V, f'AVERAGE({V}[avg_wheel_amount])', n_col='n_users'))
note('Visual: Matrix with city rows. Cards for national pct_usage and avg_wheel_amount. Line for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# APPENDIX
# ════════════════════════════════════════════════════════════════════════════
heading('20. Appendix – General DAX Patterns', level=1)

body('Standard Metric Pattern', bold=True)
dax_measure(
    'Metric Name =\n'
    'VAR SelYearWeek = IF(HASONEVALUE(ViewName[yearweek]),\n'
    '    SELECTEDVALUE(ViewName[yearweek]),\n'
    '    CALCULATE(MAX(ViewName[yearweek]), ALL(ViewName)))\n'
    'VAR MinN = [Min N Cutoff Value]\n'
    'RETURN IF(\n'
    '    CALCULATE(SUM(ViewName[n]), ViewName[yearweek] = SelYearWeek) >= MinN,\n'
    '    CALCULATE(AVERAGE(ViewName[metric_col]), ViewName[yearweek] = SelYearWeek),\n'
    '    BLANK())'
)

body('WoW Delta Pattern', bold=True)
dax_measure(
    'WoW Metric Name =\n'
    'VAR SelYearWeek = IF(HASONEVALUE(ViewName[yearweek]),\n'
    '    SELECTEDVALUE(ViewName[yearweek]),\n'
    '    CALCULATE(MAX(ViewName[yearweek]), ALL(ViewName)))\n'
    'VAR PrevYearWeek = CALCULATE(MAX(ViewName[yearweek]),\n'
    '    ALL(ViewName), ViewName[yearweek] < SelYearWeek)\n'
    'VAR MinN    = [Min N Cutoff Value]\n'
    'VAR CurrVal = CALCULATE(AVERAGE(ViewName[metric_col]), ViewName[yearweek] = SelYearWeek)\n'
    'VAR PrevVal = CALCULATE(AVERAGE(ViewName[metric_col]), ViewName[yearweek] = PrevYearWeek)\n'
    'RETURN IF(\n'
    '    CALCULATE(SUM(ViewName[n]), ViewName[yearweek] = SelYearWeek) >= MinN,\n'
    '    CurrVal - PrevVal,\n'
    '    BLANK())'
)

body('Min N Cutoff Value (helper measure)', bold=True)
dax_measure(
    'Min N Cutoff Value =\n'
    "SELECTEDVALUE('Min N Cutoff'[Min N Cutoff], 0)"
)

body('Platform-filtered UNION ALL views (CommFree, Navigation)', bold=True)
note('Hardcode platform = "Snapp" or "Tapsi" inside each measure\'s CALCULATE filter. '
     'Do not rely on a slicer for platform — it bypasses the N-gate logic.')

body('Long-format Matrix views (IncentiveAmounts, Duration, Persona, TapsiInactivity)', bold=True)
note('Place the bucket/category column on the matrix column axis. '
     'Sort it by the _sort column via Sort by Column in the Model view. '
     'Put the pct measure as Values. city on Rows. yearweek slicer filters the whole matrix.')

body('Segment-filtered views (IncentiveUnsatNational)', bold=True)
note('Hardcode segment = "All Snapp" / "Joint Snapp" / "Joint Tapsi" in each measure. '
     'Do not use a slicer for segment filtering.')

body('Sorting yearweek on x-axis / slicer', bold=True)
note('In the Model view, select yearweek (TEXT) → Sort by Column → yearweek_sort (INT). '
     'Apply to every imported view before building any visual.')

doc.add_paragraph()
body('End of Document – Driver Survey Power BI Routine Analysis Guide v5 (Complete)', bold=True, color=(80, 80, 80))

out = r'D:\Work\Driver Survey\PowerBI\PowerBI_Routine_Analysis_Guide_v5_Complete.docx'
doc.save(out)
print(f'Saved: {out}')
