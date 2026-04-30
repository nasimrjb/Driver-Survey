from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── styles ──────────────────────────────────────────────────────────────────
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

def code_block(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = 'Courier New'
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(0, 0, 139)
    return p

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
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
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
            cell = t.cell(r+1, c)
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(8.5)
            cell.paragraphs[0].runs[0].font.name = 'Courier New'
    return t

def dax_section(title, measures):
    """measures = list of (name, dax_string)"""
    p = doc.add_paragraph()
    run = p.add_run(title)
    set_font(run, size=10, bold=True, color=(0, 70, 127))
    for name, dax in measures:
        p2 = doc.add_paragraph()
        r = p2.add_run(f'// {name}')
        set_font(r, size=9, bold=True, color=(0, 100, 0))
        for line in dax.strip().split('\n'):
            code_block(line)
        doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# TITLE
# ════════════════════════════════════════════════════════════════════════════
heading('Driver Survey – Power BI Routine Analysis Guide v5', level=1)
heading('Complete DAX Reference for All 17 RA Views', level=2)
body('Database: Cab_Studies  •  Server: 192.168.18.37  •  Schema: [Cab]', color=(80,80,80))
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
for name in [
    'vw_RA_SatReview','vw_RA_CitiesOverview','vw_RA_RideShare','vw_RA_PersonaPartTime',
    'vw_RA_IncentiveAmounts','vw_RA_IncentiveDuration','vw_RA_Persona',
    'vw_RA_CommFree','vw_RA_CSRare','vw_RA_NavReco',
    'vw_RA_IncentiveTypes','vw_RA_IncentiveUnsatCity','vw_RA_IncentiveUnsatNational',
    'vw_RA_Navigation','vw_RA_Referral','vw_RA_TapsiInactivity','vw_RA_LuckyWheel',
]:
    bullet(name, indent=1)

body('Do NOT create relationships between views – each is self-contained.', bold=False, color=(160,0,0))
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 2. GLOBAL HELPER MEASURES (in a dedicated "Measures" table)
# ════════════════════════════════════════════════════════════════════════════
heading('2. Global Helper Measures', level=1)
note('Create a blank table named "Measures" (Enter Data) and add all measures below to it.')

dax_section('Threshold & Week helpers', [
    ('Min N', '''[Min N] = 30'''),
    ('Selected YearWeek', '''[Selected YearWeek] =
IF(
    HASONEVALUE(vw_RA_SatReview[yearweek]),
    VALUES(vw_RA_SatReview[yearweek]),
    MAX(vw_RA_SatReview[yearweek])
)'''),
])

note('Each view has its own yearweek column – use the view-specific version in each page slicer.')

# ════════════════════════════════════════════════════════════════════════════
# DAX PATTERN HELPERS (reusable strings)
# ════════════════════════════════════════════════════════════════════════════

def yw_var(view):
    return f'''VAR yw = IF(HASONEVALUE({view}[yearweek]), VALUES({view}[yearweek]), MAX({view}[yearweek]))'''

def prev_yw(view):
    return f'''VAR prev_yw = CALCULATE(MAX({view}[yearweek]), {view}[yearweek] < yw)'''

def n_gate(view, n_col='n'):
    return f'''VAR n_val = CALCULATE(SUM({view}[{n_col}]), {view}[yearweek] = yw)
IF(n_val < [Min N], BLANK(),'''

# ════════════════════════════════════════════════════════════════════════════
# RA-1  vw_RA_SatReview
# ════════════════════════════════════════════════════════════════════════════
heading('3. RA-1 – Satisfaction & Participation Review', level=1)
body('View: vw_RA_SatReview  |  Excel Page: #3', bold=True)
body('''Tracks incentive participation rates and satisfaction scores (1-5) for Snapp & Tapsi.
cooperation_type slicer selects: All Drivers / Part-Time / Full-Time.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek','TEXT','e.g. "26-01"'),
        ('yearweek_sort','INT','for axis ordering'),
        ('weeknumber','INT',''),
        ('city','TEXT',''),
        ('cooperation_type','TEXT','All Drivers / Part-Time / Full-Time'),
        ('n','INT','all respondents in city'),
        ('n_joint','INT','joint drivers'),
        ('Part_pct_Snapp','FLOAT','% participated in Snapp incentive'),
        ('Part_pct_Jnt_Snapp','FLOAT','% joint participated in Snapp'),
        ('Part_pct_Jnt_Tapsi','FLOAT','% joint participated in Tapsi'),
        ('Part_GotMsg_pct_Snapp','FLOAT','% participated among who got Snapp msg'),
        ('Part_GotMsg_pct_Jnt_Snapp','FLOAT',''),
        ('Part_GotMsg_pct_Jnt_Tapsi','FLOAT',''),
        ('Incentive_Sat_Snapp','FLOAT','avg 1-5'),
        ('Incentive_Sat_Jnt_Snapp','FLOAT',''),
        ('Incentive_Sat_Jnt_Tapsi','FLOAT',''),
        ('Fare_Sat_Snapp','FLOAT','avg 1-5'),
        ('Fare_Sat_Jnt_Snapp','FLOAT',''),
        ('Fare_Sat_Jnt_Tapsi','FLOAT',''),
        ('Request_Sat_Snapp','FLOAT','avg 1-5'),
        ('Request_Sat_Jnt_Snapp','FLOAT',''),
        ('Request_Sat_Jnt_Tapsi','FLOAT',''),
        ('Income_Sat_Snapp','FLOAT','avg 1-5'),
        ('Income_Sat_Jnt_Snapp','FLOAT',''),
        ('Income_Sat_Jnt_Tapsi','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA1 N (All)', f'''{yw_var("vw_RA_SatReview")}
VAR result = CALCULATE(SUM(vw_RA_SatReview[n]),
    vw_RA_SatReview[yearweek] = yw)
RETURN result'''),
    ('RA1 Part% Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Part_pct_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Part% Jnt Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n_joint]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Part_pct_Jnt_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Part% Jnt Tapsi', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n_joint]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Part_pct_Jnt_Tapsi]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 GotMsg Part% Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Part_GotMsg_pct_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Incentive Sat Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Incentive_Sat_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Incentive Sat Jnt Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n_joint]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Incentive_Sat_Jnt_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Incentive Sat Jnt Tapsi', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n_joint]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Incentive_Sat_Jnt_Tapsi]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Fare Sat Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Fare_Sat_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Request Sat Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Request_Sat_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 Income Sat Snapp', f'''{yw_var("vw_RA_SatReview")}
VAR n_val = CALCULATE(SUM(vw_RA_SatReview[n]), vw_RA_SatReview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_SatReview[Income_Sat_Snapp]), vw_RA_SatReview[yearweek] = yw))'''),
    ('RA1 WoW Incentive Sat Snapp', f'''{yw_var("vw_RA_SatReview")}
{prev_yw("vw_RA_SatReview")}
VAR curr = CALCULATE(AVERAGE(vw_RA_SatReview[Incentive_Sat_Snapp]), vw_RA_SatReview[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_SatReview[Incentive_Sat_Snapp]), vw_RA_SatReview[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Apply filter vw_RA_SatReview[cooperation_type] = "All Drivers" (or the relevant segment) before building visuals. Duplicate measures with Jnt_Snapp / Jnt_Tapsi suffix as needed.')
note('Visual: Matrix with city on rows, week slicer. Cards for national averages. Line chart for WoW trend.')

# ════════════════════════════════════════════════════════════════════════════
# RA-2  vw_RA_CitiesOverview
# ════════════════════════════════════════════════════════════════════════════
heading('4. RA-2 – Cities Overview', level=1)
body('View: vw_RA_CitiesOverview  |  Excel Page: #12', bold=True)
body('Three independent respondent groups: E (all), F (joint), G (tapsi_LOC > 0). No cooperation_type split.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber','INT',''),
        ('city','TEXT',''),
        ('E_n','INT','all respondents'),
        ('F_n','INT','joint drivers'),
        ('G_n','INT','drivers with tapsi LOC > 0'),
        ('pct_Joint','FLOAT','% who are joint'),
        ('pct_Dual_SU','FLOAT','% with tapsi LOC > 0'),
        ('AvgLOC_All_Snapp','FLOAT','avg Snapp LOC'),
        ('GotMsg_All_Snapp','FLOAT','% got Snapp incentive msg'),
        ('AvgLOC_Joint_Snapp','FLOAT','avg Snapp LOC among joint'),
        ('GotMsg_Joint_Snapp','FLOAT','% joint got Snapp msg'),
        ('GotMsg_Joint_Cmpt','FLOAT','% joint got Tapsi msg'),
        ('AvgLOC_Joint_Cmpt','FLOAT','avg Tapsi LOC among joint'),
        ('AvgLOC_Joint_Cmpt_SU','FLOAT','avg Tapsi LOC among G-group'),
        ('GotMsg_Joint_Cmpt_SU','FLOAT','% G-group got Tapsi msg'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA2 E_n', f'''{yw_var("vw_RA_CitiesOverview")}
CALCULATE(SUM(vw_RA_CitiesOverview[E_n]), vw_RA_CitiesOverview[yearweek] = yw)'''),
    ('RA2 pct Joint', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[E_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[pct_Joint]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 pct Dual SU', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[E_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[pct_Dual_SU]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 AvgLOC All Snapp', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[E_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[AvgLOC_All_Snapp]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 GotMsg All Snapp', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[E_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[GotMsg_All_Snapp]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 GotMsg Joint Snapp', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[F_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[GotMsg_Joint_Snapp]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 GotMsg Joint Cmpt', f'''{yw_var("vw_RA_CitiesOverview")}
VAR n_val = CALCULATE(SUM(vw_RA_CitiesOverview[F_n]), vw_RA_CitiesOverview[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CitiesOverview[GotMsg_Joint_Cmpt]), vw_RA_CitiesOverview[yearweek] = yw))'''),
    ('RA2 WoW pct Joint', f'''{yw_var("vw_RA_CitiesOverview")}
{prev_yw("vw_RA_CitiesOverview")}
VAR curr = CALCULATE(AVERAGE(vw_RA_CitiesOverview[pct_Joint]), vw_RA_CitiesOverview[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_CitiesOverview[pct_Joint]), vw_RA_CitiesOverview[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Visual: Matrix with city rows. Use E_n / F_n / G_n as gate for corresponding metric groups.')

# ════════════════════════════════════════════════════════════════════════════
# RA-3  vw_RA_RideShare
# ════════════════════════════════════════════════════════════════════════════
heading('5. RA-3 – Ride Share', level=1)
body('View: vw_RA_RideShare  |  Excel Page: #13', bold=True)
body('Total ride counts and Snapp/Tapsi share by driver segment.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber / city','',''),
        ('total_Res','INT','all respondents'),
        ('Joint_Res','INT','joint drivers'),
        ('Ex_drivers','INT','exclusive Snapp drivers'),
        ('Total_Ride','FLOAT','Snapp + Tapsi rides combined'),
        ('Total_Ride_Snapp','FLOAT','all Snapp rides'),
        ('Ex_Ride_Snapp','FLOAT','exclusive driver Snapp rides'),
        ('Jnt_Snapp_Ride','FLOAT','joint driver Snapp rides'),
        ('Jnt_Tapsi_Ride','FLOAT','joint driver Tapsi rides'),
        ('All_Snapp_pct','FLOAT','Snapp share of total rides'),
        ('Ex_Drivers_Snapp_pct','FLOAT','exclusive driver share'),
        ('Jnt_at_Snapp_pct','FLOAT','joint driver Snapp share'),
        ('Jnt_at_Tapsi_pct','FLOAT','joint driver Tapsi share'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA3 Total Respondents', f'''{yw_var("vw_RA_RideShare")}
CALCULATE(SUM(vw_RA_RideShare[total_Res]), vw_RA_RideShare[yearweek] = yw)'''),
    ('RA3 All Snapp pct', f'''{yw_var("vw_RA_RideShare")}
VAR n_val = CALCULATE(SUM(vw_RA_RideShare[total_Res]), vw_RA_RideShare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_RideShare[All_Snapp_pct]), vw_RA_RideShare[yearweek] = yw))'''),
    ('RA3 Jnt at Snapp pct', f'''{yw_var("vw_RA_RideShare")}
VAR n_val = CALCULATE(SUM(vw_RA_RideShare[Joint_Res]), vw_RA_RideShare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_RideShare[Jnt_at_Snapp_pct]), vw_RA_RideShare[yearweek] = yw))'''),
    ('RA3 Jnt at Tapsi pct', f'''{yw_var("vw_RA_RideShare")}
VAR n_val = CALCULATE(SUM(vw_RA_RideShare[Joint_Res]), vw_RA_RideShare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_RideShare[Jnt_at_Tapsi_pct]), vw_RA_RideShare[yearweek] = yw))'''),
    ('RA3 WoW All Snapp pct', f'''{yw_var("vw_RA_RideShare")}
{prev_yw("vw_RA_RideShare")}
VAR curr = CALCULATE(AVERAGE(vw_RA_RideShare[All_Snapp_pct]), vw_RA_RideShare[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_RideShare[All_Snapp_pct]), vw_RA_RideShare[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Visual: Clustered bar for pct columns. City slicer or matrix rows.')

# ════════════════════════════════════════════════════════════════════════════
# RA-4  vw_RA_PersonaPartTime
# ════════════════════════════════════════════════════════════════════════════
heading('6. RA-4 – Persona Part-Time', level=1)
body('View: vw_RA_PersonaPartTime  |  Excel Page: #15 (Part-Time sub-table)', bold=True)
body('Part-Time rate and average rides per boarded driver, by city.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber / city','',''),
        ('total_Res','INT',''),
        ('Joint_Res','INT',''),
        ('Ex_drivers','INT',''),
        ('PT_pct_Joint','FLOAT','% Part-Time among joint'),
        ('PT_pct_Exclusive','FLOAT','% Part-Time among exclusive'),
        ('RidePerBoarded_Snapp','FLOAT','avg Snapp rides per joint driver'),
        ('RidePerBoarded_Tapsi','FLOAT','avg Tapsi rides per joint driver'),
        ('AvgAllRides','FLOAT','avg Snapp rides across all respondents'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA4 PT pct Joint', f'''{yw_var("vw_RA_PersonaPartTime")}
VAR n_val = CALCULATE(SUM(vw_RA_PersonaPartTime[Joint_Res]), vw_RA_PersonaPartTime[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_PersonaPartTime[PT_pct_Joint]), vw_RA_PersonaPartTime[yearweek] = yw))'''),
    ('RA4 PT pct Exclusive', f'''{yw_var("vw_RA_PersonaPartTime")}
VAR n_val = CALCULATE(SUM(vw_RA_PersonaPartTime[Ex_drivers]), vw_RA_PersonaPartTime[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_PersonaPartTime[PT_pct_Exclusive]), vw_RA_PersonaPartTime[yearweek] = yw))'''),
    ('RA4 RidePerBoarded Snapp', f'''{yw_var("vw_RA_PersonaPartTime")}
VAR n_val = CALCULATE(SUM(vw_RA_PersonaPartTime[Joint_Res]), vw_RA_PersonaPartTime[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_PersonaPartTime[RidePerBoarded_Snapp]), vw_RA_PersonaPartTime[yearweek] = yw))'''),
    ('RA4 RidePerBoarded Tapsi', f'''{yw_var("vw_RA_PersonaPartTime")}
VAR n_val = CALCULATE(SUM(vw_RA_PersonaPartTime[Joint_Res]), vw_RA_PersonaPartTime[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_PersonaPartTime[RidePerBoarded_Tapsi]), vw_RA_PersonaPartTime[yearweek] = yw))'''),
    ('RA4 AvgAllRides', f'''{yw_var("vw_RA_PersonaPartTime")}
VAR n_val = CALCULATE(SUM(vw_RA_PersonaPartTime[total_Res]), vw_RA_PersonaPartTime[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_PersonaPartTime[AvgAllRides]), vw_RA_PersonaPartTime[yearweek] = yw))'''),
])

# ════════════════════════════════════════════════════════════════════════════
# RA-5  vw_RA_IncentiveAmounts
# ════════════════════════════════════════════════════════════════════════════
heading('7. RA-5 – Incentive Amounts (Long Format)', level=1)
body('View: vw_RA_IncentiveAmounts  |  Excel Pages: #1 (Snapp) and #2 (Tapsi)', bold=True)
body('''Long format: one row per city × incentive_range bucket × week.
Filter platform = "Snapp" for page #1, "Tapsi" for page #2.
Tapsi rows use only joint drivers.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber / city / platform','','Snapp or Tapsi'),
        ('incentive_range','TEXT','bucket label e.g. "<20k"'),
        ('incentive_range_sort','INT','for x-axis ordering'),
        ('n_range','INT','count in this bucket'),
        ('n_total','INT','total respondents for city × week'),
        ('pct','FLOAT','n_range / n_total × 100'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA5 n Total (Snapp)', f'''{yw_var("vw_RA_IncentiveAmounts")}
CALCULATE(
    SUM(vw_RA_IncentiveAmounts[n_total]),
    vw_RA_IncentiveAmounts[yearweek] = yw,
    vw_RA_IncentiveAmounts[platform] = "Snapp")'''),
    ('RA5 pct Range (Snapp)', f'''// Place on a matrix: rows = city, columns = incentive_range (sorted by incentive_range_sort)
{yw_var("vw_RA_IncentiveAmounts")}
VAR n_val = CALCULATE(
    MAX(vw_RA_IncentiveAmounts[n_total]),
    vw_RA_IncentiveAmounts[yearweek] = yw,
    vw_RA_IncentiveAmounts[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(
        SUM(vw_RA_IncentiveAmounts[pct]),
        vw_RA_IncentiveAmounts[yearweek] = yw,
        vw_RA_IncentiveAmounts[platform] = "Snapp"))'''),
    ('RA5 pct Range (Tapsi)', f'''{yw_var("vw_RA_IncentiveAmounts")}
VAR n_val = CALCULATE(
    MAX(vw_RA_IncentiveAmounts[n_total]),
    vw_RA_IncentiveAmounts[yearweek] = yw,
    vw_RA_IncentiveAmounts[platform] = "Tapsi")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(
        SUM(vw_RA_IncentiveAmounts[pct]),
        vw_RA_IncentiveAmounts[yearweek] = yw,
        vw_RA_IncentiveAmounts[platform] = "Tapsi"))'''),
])
note('Visual: Matrix (city × incentive_range). Sort incentive_range by incentive_range_sort column. Page-level filter on platform.')

# ════════════════════════════════════════════════════════════════════════════
# RA-6  vw_RA_IncentiveDuration
# ════════════════════════════════════════════════════════════════════════════
heading('8. RA-6 – Incentive Duration (Long Format)', level=1)
body('View: vw_RA_IncentiveDuration  |  Excel Page: #4', bold=True)
body('Long format: how long drivers have had an active incentive. Snapp (all) and Tapsi (joint only) rows.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber / city / platform','','Snapp or Tapsi'),
        ('duration_bucket','TEXT','"Few Hours","1 Day","1_6 Days","7 Days",">7 Days"'),
        ('duration_bucket_sort','INT','1–5 + 99'),
        ('n_range','INT',''),
        ('n_total','INT',''),
        ('pct','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA6 n Total', f'''{yw_var("vw_RA_IncentiveDuration")}
CALCULATE(
    MAX(vw_RA_IncentiveDuration[n_total]),
    vw_RA_IncentiveDuration[yearweek] = yw)'''),
    ('RA6 pct Bucket', f'''// Matrix: rows = city, columns = duration_bucket (sorted by duration_bucket_sort)
// Add platform slicer or page-level filter
{yw_var("vw_RA_IncentiveDuration")}
VAR n_val = CALCULATE(
    MAX(vw_RA_IncentiveDuration[n_total]),
    vw_RA_IncentiveDuration[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(
        SUM(vw_RA_IncentiveDuration[pct]),
        vw_RA_IncentiveDuration[yearweek] = yw))'''),
])
note('Visual: Same matrix pattern as RA-5. Sort duration_bucket by duration_bucket_sort.')

# ════════════════════════════════════════════════════════════════════════════
# RA-7  vw_RA_Persona
# ════════════════════════════════════════════════════════════════════════════
heading('9. RA-7 – Persona (Long Format)', level=1)
body('View: vw_RA_Persona  |  Excel Page: #15 (all demographic sub-tables)', bold=True)
body('''Long format with dimension slicer. Dimensions: Activity Type, Age Group, Education,
Marital Status, Gender, Cooperation Type. One category column holds bucket values.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort','TEXT/INT',''),
        ('weeknumber / city','',''),
        ('dimension','TEXT','e.g. "Activity Type"'),
        ('category','TEXT','bucket value'),
        ('category_sort','INT','for axis ordering'),
        ('n','INT','count in this category'),
        ('n_total','INT','total for city × week × dimension'),
        ('pct','FLOAT','n / n_total × 100'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA7 n Total', f'''{yw_var("vw_RA_Persona")}
CALCULATE(
    MAX(vw_RA_Persona[n_total]),
    vw_RA_Persona[yearweek] = yw)'''),
    ('RA7 pct Category', f'''// Matrix: rows = city, columns = category (sorted by category_sort)
// Dimension slicer controls which demographic is shown
{yw_var("vw_RA_Persona")}
VAR n_val = CALCULATE(
    MAX(vw_RA_Persona[n_total]),
    vw_RA_Persona[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(
        SUM(vw_RA_Persona[pct]),
        vw_RA_Persona[yearweek] = yw))'''),
])
note('Add a slicer on vw_RA_Persona[dimension] so users can switch between demographic breakdowns.')
note('Sort each category column by category_sort using "Sort by Column" in the Model view.')

# ════════════════════════════════════════════════════════════════════════════
# RA-8  vw_RA_CommFree
# ════════════════════════════════════════════════════════════════════════════
heading('10. RA-8 – Commission-Free Incentive', level=1)
body('View: vw_RA_CommFree  |  Excel Page: #18 (Snapp + Tapsi)', bold=True)
body('''UNION ALL of Snapp (all drivers) and Tapsi (joint only) rows. Platform column distinguishes them.
Incentive type binary flags (multi-select) come from WideMain via LEFT JOIN.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city / platform','','Snapp or Tapsi'),
        ('n','INT','base respondents'),
        ('Who_Got_Message','INT','count who received incentive message'),
        ('GotMsg_Money','INT','msg received, category = Money'),
        ('GotMsg_FreeComm','INT','msg received, category = Free-Commission'),
        ('GotMsg_Money_FreeComm','INT','msg received, category = Money & Free-commission'),
        ('GotMsg_PayRide','INT','received msg + PayAfterRide type'),
        ('GotMsg_EarnCF','INT','received msg + EarningBasedCF type'),
        ('GotMsg_RideCF','INT','received msg + RideBasedCF type'),
        ('GotMsg_IncGuar','INT','received msg + IncomeGuarantee type'),
        ('GotMsg_PayInc','INT','received msg + PayAfterIncome type'),
        ('GotMsg_CFSome','INT','received msg + CFSomeTrips type'),
        ('Free_Comm_Drivers','INT','drivers with CF rides > 0'),
        ('Participated','INT','participated in incentive'),
        ('pct_Got_Message','FLOAT','%'),
        ('pct_Free_Comm_Ride','FLOAT','%'),
        ('pct_Participated','FLOAT','% of msg-recipients who participated'),
        ('Avg_CF_Rides','FLOAT','avg CF rides (among CF drivers)'),
        ('Avg_Total_Rides','FLOAT','avg total rides'),
        ('Avg_pct_CF_RideShare','FLOAT','avg CF% of total rides'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA8 n (Snapp)', f'''{yw_var("vw_RA_CommFree")}
CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw,
    vw_RA_CommFree[platform] = "Snapp")'''),
    ('RA8 pct Got Message (Snapp)', f'''{yw_var("vw_RA_CommFree")}
VAR n_val = CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CommFree[pct_Got_Message]),
        vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp"))'''),
    ('RA8 pct Free Comm Ride (Snapp)', f'''{yw_var("vw_RA_CommFree")}
VAR n_val = CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CommFree[pct_Free_Comm_Ride]),
        vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp"))'''),
    ('RA8 pct Participated (Snapp)', f'''{yw_var("vw_RA_CommFree")}
VAR n_val = CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CommFree[pct_Participated]),
        vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp"))'''),
    ('RA8 GotMsg by Type % (Snapp, PayRide)', f'''{yw_var("vw_RA_CommFree")}
VAR who_got = CALCULATE(SUM(vw_RA_CommFree[Who_Got_Message]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
VAR n_type = CALCULATE(SUM(vw_RA_CommFree[GotMsg_PayRide]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(who_got < [Min N], BLANK(), DIVIDE(n_type, who_got) * 100)'''),
    ('RA8 Avg CF Rides (Snapp)', f'''{yw_var("vw_RA_CommFree")}
VAR n_val = CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CommFree[Avg_CF_Rides]),
        vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp"))'''),
    ('RA8 Avg pct CF RideShare (Snapp)', f'''{yw_var("vw_RA_CommFree")}
VAR n_val = CALCULATE(SUM(vw_RA_CommFree[n]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CommFree[Avg_pct_CF_RideShare]),
        vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp"))'''),
    ('RA8 WoW pct Got Message (Snapp)', f'''{yw_var("vw_RA_CommFree")}
{prev_yw("vw_RA_CommFree")}
VAR curr = CALCULATE(AVERAGE(vw_RA_CommFree[pct_Got_Message]),
    vw_RA_CommFree[yearweek] = yw, vw_RA_CommFree[platform] = "Snapp")
VAR prev = CALCULATE(AVERAGE(vw_RA_CommFree[pct_Got_Message]),
    vw_RA_CommFree[yearweek] = prev_yw, vw_RA_CommFree[platform] = "Snapp")
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Duplicate all Snapp measures with platform = "Tapsi" for the Tapsi section of page #18.')
note('Visual: Matrix with city rows. Cards for national pct_Got_Message / pct_Free_Comm_Ride.')

# ════════════════════════════════════════════════════════════════════════════
# RA-9  vw_RA_CSRare
# ════════════════════════════════════════════════════════════════════════════
heading('11. RA-9 – Customer Support Satisfaction', level=1)
body('View: vw_RA_CSRare  |  Excel Pages: CS_Sat_Snapp / CS_Sat_Tapsi', bold=True)
body('Data from ShortRare joined to ShortMain. All scores are 1-5 averages.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n','INT','respondents with ShortRare record'),
        ('Snapp_CS_Overall','FLOAT','avg overall CS satisfaction 1-5'),
        ('Snapp_CS_WaitTime','FLOAT',''),
        ('Snapp_CS_Solution','FLOAT',''),
        ('Snapp_CS_Behaviour','FLOAT',''),
        ('Snapp_CS_Relevance','FLOAT',''),
        ('Snapp_CS_Solved_pct','FLOAT','% whose issue was solved'),
        ('Tapsi_CS_Overall','FLOAT',''),
        ('Tapsi_CS_WaitTime','FLOAT',''),
        ('Tapsi_CS_Solution','FLOAT',''),
        ('Tapsi_CS_Behaviour','FLOAT',''),
        ('Tapsi_CS_Relevance','FLOAT',''),
        ('Tapsi_CS_Solved_pct','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA9 n', f'''{yw_var("vw_RA_CSRare")}
CALCULATE(SUM(vw_RA_CSRare[n]), vw_RA_CSRare[yearweek] = yw)'''),
    ('RA9 Snapp CS Overall', f'''{yw_var("vw_RA_CSRare")}
VAR n_val = CALCULATE(SUM(vw_RA_CSRare[n]), vw_RA_CSRare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CSRare[Snapp_CS_Overall]), vw_RA_CSRare[yearweek] = yw))'''),
    ('RA9 Snapp CS Solved pct', f'''{yw_var("vw_RA_CSRare")}
VAR n_val = CALCULATE(SUM(vw_RA_CSRare[n]), vw_RA_CSRare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CSRare[Snapp_CS_Solved_pct]), vw_RA_CSRare[yearweek] = yw))'''),
    ('RA9 Tapsi CS Overall', f'''{yw_var("vw_RA_CSRare")}
VAR n_val = CALCULATE(SUM(vw_RA_CSRare[n]), vw_RA_CSRare[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_CSRare[Tapsi_CS_Overall]), vw_RA_CSRare[yearweek] = yw))'''),
    ('RA9 WoW Snapp CS Overall', f'''{yw_var("vw_RA_CSRare")}
{prev_yw("vw_RA_CSRare")}
VAR curr = CALCULATE(AVERAGE(vw_RA_CSRare[Snapp_CS_Overall]), vw_RA_CSRare[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_CSRare[Snapp_CS_Overall]), vw_RA_CSRare[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Replicate all five Snapp score measures for Tapsi. Visual: Matrix city × score, Line chart for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# RA-10  vw_RA_NavReco
# ════════════════════════════════════════════════════════════════════════════
heading('12. RA-10 – Navigation & NPS Recommendation Scores', level=1)
body('View: vw_RA_NavReco  |  Excel Pages: NavReco_Scores / Reco_NPS', bold=True)
body('From ShortRare. NPS scores and navigation app recommendation scores (0-10 or 1-10).')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n','INT',''),
        ('Snapp_NPS','FLOAT','avg recommend Snapp score'),
        ('Tapsi_NPS_SnapDriver','FLOAT','avg recommend Tapsi score (Snapp driver respondent)'),
        ('Tapsi_NPS_TapsiDriver','FLOAT','avg recommend Tapsi score (Tapsi driver respondent)'),
        ('Reco_GoogleMap','FLOAT','avg recommendation score'),
        ('Reco_Waze','FLOAT',''),
        ('Reco_Neshan','FLOAT',''),
        ('Reco_Balad','FLOAT',''),
        ('Snapp_Nav_Sat','FLOAT','avg Snapp nav app satisfaction'),
        ('Tapsi_Nav_Sat','FLOAT','avg Tapsi in-app nav satisfaction'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA10 Snapp NPS', f'''{yw_var("vw_RA_NavReco")}
VAR n_val = CALCULATE(SUM(vw_RA_NavReco[n]), vw_RA_NavReco[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_NavReco[Snapp_NPS]), vw_RA_NavReco[yearweek] = yw))'''),
    ('RA10 Tapsi NPS (Snapp Driver)', f'''{yw_var("vw_RA_NavReco")}
VAR n_val = CALCULATE(SUM(vw_RA_NavReco[n]), vw_RA_NavReco[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_NavReco[Tapsi_NPS_SnapDriver]), vw_RA_NavReco[yearweek] = yw))'''),
    ('RA10 Reco Neshan', f'''{yw_var("vw_RA_NavReco")}
VAR n_val = CALCULATE(SUM(vw_RA_NavReco[n]), vw_RA_NavReco[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_NavReco[Reco_Neshan]), vw_RA_NavReco[yearweek] = yw))'''),
    ('RA10 Reco Balad', f'''{yw_var("vw_RA_NavReco")}
VAR n_val = CALCULATE(SUM(vw_RA_NavReco[n]), vw_RA_NavReco[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_NavReco[Reco_Balad]), vw_RA_NavReco[yearweek] = yw))'''),
    ('RA10 Snapp Nav Sat', f'''{yw_var("vw_RA_NavReco")}
VAR n_val = CALCULATE(SUM(vw_RA_NavReco[n]), vw_RA_NavReco[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_NavReco[Snapp_Nav_Sat]), vw_RA_NavReco[yearweek] = yw))'''),
    ('RA10 WoW Snapp NPS', f'''{yw_var("vw_RA_NavReco")}
{prev_yw("vw_RA_NavReco")}
VAR curr = CALCULATE(AVERAGE(vw_RA_NavReco[Snapp_NPS]), vw_RA_NavReco[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_NavReco[Snapp_NPS]), vw_RA_NavReco[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])

# ════════════════════════════════════════════════════════════════════════════
# RA-11  vw_RA_IncentiveTypes
# ════════════════════════════════════════════════════════════════════════════
heading('13. RA-11 – Incentive Type Distribution', level=1)
body('View: vw_RA_IncentiveTypes  |  Excel Pages: #5 (Snapp Excl) and #6 (Joint)', bold=True)
body('''Multi-select incentive types joined from WideMain. Six type flags each for Snapp (Excl/Joint) and Tapsi (Joint).
Base for Excl metrics = n_excl; base for Joint metrics = n_joint.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n / n_joint / n_excl','INT',''),
        ('pct_GotMsg_Excl_Snapp','FLOAT','% excl drivers who got Snapp msg'),
        ('pct_GotMsg_Jnt_Snapp','FLOAT','% joint drivers who got Snapp msg'),
        ('pct_GotMsg_Jnt_Tapsi','FLOAT','% joint drivers who got Tapsi msg'),
        ('pct_GotMsg_Both','FLOAT','% joint who got both msgs'),
        ('pct_GotMsg_Diff','FLOAT','% joint who got only one msg'),
        ('pct_PayRide_Excl','FLOAT','Pay-After-Ride type % (excl Snapp)'),
        ('pct_EarnCF_Excl','FLOAT','Earning-Based CF type % (excl Snapp)'),
        ('pct_RideCF_Excl','FLOAT','Ride-Based CF type % (excl Snapp)'),
        ('pct_IncGuar_Excl','FLOAT','Income Guarantee type % (excl Snapp)'),
        ('pct_PayInc_Excl','FLOAT','Pay-After-Income type % (excl Snapp)'),
        ('pct_CFSome_Excl','FLOAT','CF on Some Trips type % (excl Snapp)'),
        ('pct_PayRide_JntSn','FLOAT','... type % (joint Snapp)'),
        ('pct_EarnCF_JntSn / pct_RideCF_JntSn / pct_IncGuar_JntSn / pct_PayInc_JntSn / pct_CFSome_JntSn','FLOAT',''),
        ('pct_PayRide_JntTp','FLOAT','... type % (joint Tapsi)'),
        ('pct_EarnCF_JntTp / pct_RideCF_JntTp / pct_IncGuar_JntTp / pct_PayInc_JntTp / pct_CFSome_JntTp','FLOAT',''),
        ('Avg_CF_Rides_Snapp','FLOAT','avg CF rides (CF drivers only)'),
        ('Avg_CF_Rides_Tapsi','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA11 n Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)'''),
    ('RA11 n Joint', f'''{yw_var("vw_RA_IncentiveTypes")}
CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)'''),
    ('RA11 pct GotMsg Excl Snapp', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_GotMsg_Excl_Snapp]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct GotMsg Jnt Snapp', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_GotMsg_Jnt_Snapp]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct GotMsg Jnt Tapsi', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_GotMsg_Jnt_Tapsi]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct GotMsg Both', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_GotMsg_Both]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct PayRide Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_PayRide_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct EarnCF Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_EarnCF_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct RideCF Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_RideCF_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct IncGuar Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_IncGuar_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct PayInc Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_PayInc_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct CFSome Excl', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_excl]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_CFSome_Excl]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct PayRide JntSn', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_PayRide_JntSn]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 pct PayRide JntTp', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[pct_PayRide_JntTp]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 Avg CF Rides Snapp', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[Avg_CF_Rides_Snapp]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
    ('RA11 Avg CF Rides Tapsi', f'''{yw_var("vw_RA_IncentiveTypes")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveTypes[n_joint]), vw_RA_IncentiveTypes[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveTypes[Avg_CF_Rides_Tapsi]), vw_RA_IncentiveTypes[yearweek] = yw))'''),
])
note('Duplicate EarnCF / RideCF / IncGuar / PayInc / CFSome measures for JntSn and JntTp segments as needed.')

# ════════════════════════════════════════════════════════════════════════════
# RA-12  vw_RA_IncentiveUnsatCity
# ════════════════════════════════════════════════════════════════════════════
heading('14. RA-12 – Incentive Dissatisfaction by City', level=1)
body('View: vw_RA_IncentiveUnsatCity  |  Excel Page: #8', bold=True)
body('''Multi-select unsatisfaction reasons from WideMain. "Low sat" = driver cited ANY reason.
Snapp base = n_sn_low_sat; Tapsi base = n_tp_low_sat (joint only).''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n_all','INT','all respondents'),
        ('n_joint','INT','joint drivers'),
        ('n_sn_low_sat','INT','Snapp dissatisfied (any reason)'),
        ('n_tp_low_sat','INT','Tapsi dissatisfied joint drivers'),
        ('pct_Sn_NoTime','FLOAT','% of n_sn_low_sat citing "Not Available"'),
        ('pct_Sn_ImpAmt','FLOAT','% citing "Improper Amount"'),
        ('pct_Sn_LowTime','FLOAT','% citing "No Time todo"'),
        ('pct_Sn_HardToDo','FLOAT','% citing "difficult"'),
        ('pct_Sn_NonPay','FLOAT','% citing "Non Payment"'),
        ('pct_Tp_NoTime','FLOAT','% of n_tp_low_sat citing "Not Available"'),
        ('pct_Tp_ImpAmt','FLOAT',''),
        ('pct_Tp_LowTime','FLOAT',''),
        ('pct_Tp_HardToDo','FLOAT',''),
        ('pct_Tp_NonPay','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA12 n Sn Low Sat', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)'''),
    ('RA12 n Tp Low Sat', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_tp_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)'''),
    ('RA12 pct Sn NoTime', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Sn_NoTime]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Sn ImpAmt', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Sn_ImpAmt]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Sn LowTime', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Sn_LowTime]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Sn HardToDo', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Sn_HardToDo]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Sn NonPay', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_sn_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Sn_NonPay]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Tp NoTime', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_tp_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Tp_NoTime]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
    ('RA12 pct Tp NonPay', f'''{yw_var("vw_RA_IncentiveUnsatCity")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatCity[n_tp_low_sat]), vw_RA_IncentiveUnsatCity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatCity[pct_Tp_NonPay]), vw_RA_IncentiveUnsatCity[yearweek] = yw))'''),
])

# ════════════════════════════════════════════════════════════════════════════
# RA-13  vw_RA_IncentiveUnsatNational
# ════════════════════════════════════════════════════════════════════════════
heading('15. RA-13 – Incentive Dissatisfaction (National)', level=1)
body('View: vw_RA_IncentiveUnsatNational  |  Excel Page: #9', bold=True)
body('''Long format – no city column. Segment = All Snapp / Joint Snapp / Joint Tapsi.
Slicer on segment or use segment as axis/legend.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber','',''),
        ('segment','TEXT','"All Snapp","Joint Snapp","Joint Tapsi"'),
        ('segment_sort','INT','1, 2, 3'),
        ('n','INT','total respondents in segment'),
        ('n_low_sat','INT','dissatisfied drivers in segment'),
        ('pct_NoTime','FLOAT','% of n_low_sat citing each reason'),
        ('pct_ImpAmt / pct_LowTime / pct_HardToDo / pct_NonPay','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA13 n Low Sat (All Snapp)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
CALCULATE(SUM(vw_RA_IncentiveUnsatNational[n_low_sat]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")'''),
    ('RA13 pct NoTime (All Snapp)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatNational[n_low_sat]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_NoTime]),
        vw_RA_IncentiveUnsatNational[yearweek] = yw,
        vw_RA_IncentiveUnsatNational[segment] = "All Snapp"))'''),
    ('RA13 pct ImpAmt (All Snapp)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatNational[n_low_sat]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_ImpAmt]),
        vw_RA_IncentiveUnsatNational[yearweek] = yw,
        vw_RA_IncentiveUnsatNational[segment] = "All Snapp"))'''),
    ('RA13 pct LowTime (All Snapp)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatNational[n_low_sat]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_LowTime]),
        vw_RA_IncentiveUnsatNational[yearweek] = yw,
        vw_RA_IncentiveUnsatNational[segment] = "All Snapp"))'''),
    ('RA13 pct HardToDo (Joint Tapsi)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
VAR n_val = CALCULATE(SUM(vw_RA_IncentiveUnsatNational[n_low_sat]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "Joint Tapsi")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_HardToDo]),
        vw_RA_IncentiveUnsatNational[yearweek] = yw,
        vw_RA_IncentiveUnsatNational[segment] = "Joint Tapsi"))'''),
    ('RA13 WoW pct NoTime (All Snapp)', f'''{yw_var("vw_RA_IncentiveUnsatNational")}
{prev_yw("vw_RA_IncentiveUnsatNational")}
VAR curr = CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_NoTime]),
    vw_RA_IncentiveUnsatNational[yearweek] = yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")
VAR prev = CALCULATE(AVERAGE(vw_RA_IncentiveUnsatNational[pct_NoTime]),
    vw_RA_IncentiveUnsatNational[yearweek] = prev_yw,
    vw_RA_IncentiveUnsatNational[segment] = "All Snapp")
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Create all five pct measures for each of the three segments (15 measures total).')
note('Visual: Clustered bar per reason, segment as legend. Line for WoW trend per reason.')

# ════════════════════════════════════════════════════════════════════════════
# RA-14  vw_RA_Navigation
# ════════════════════════════════════════════════════════════════════════════
heading('16. RA-14 – Navigation App Usage by City', level=1)
body('View: vw_RA_Navigation  |  Excel Page: #14', bold=True)
body('''UNION ALL Snapp (all) + Tapsi (joint only). Platform slicer or page filter.
Snapp has GoogleMap & Waze; Tapsi has InAppNav. NULL columns for the opposite platform.''')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('platform','TEXT','Snapp or Tapsi'),
        ('n','INT','respondents with non-null navigation response'),
        ('pct_Neshan','FLOAT',''),
        ('pct_Balad','FLOAT',''),
        ('pct_None','FLOAT','% No Navigation App'),
        ('pct_GoogleMap','FLOAT','Snapp only (NULL for Tapsi)'),
        ('pct_Waze','FLOAT','Snapp only (NULL for Tapsi)'),
        ('pct_InAppNav','FLOAT','Tapsi only (NULL for Snapp)'),
        ('pct_Other','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA14 n (Snapp)', f'''{yw_var("vw_RA_Navigation")}
CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw,
    vw_RA_Navigation[platform] = "Snapp")'''),
    ('RA14 pct Neshan (Snapp)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_Neshan]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp"))'''),
    ('RA14 pct Balad (Snapp)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_Balad]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp"))'''),
    ('RA14 pct GoogleMap (Snapp)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_GoogleMap]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp"))'''),
    ('RA14 pct Waze (Snapp)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_Waze]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp"))'''),
    ('RA14 pct None (Snapp)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_None]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp"))'''),
    ('RA14 pct InAppNav (Tapsi)', f'''{yw_var("vw_RA_Navigation")}
VAR n_val = CALCULATE(SUM(vw_RA_Navigation[n]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Tapsi")
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Navigation[pct_InAppNav]),
        vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Tapsi"))'''),
    ('RA14 WoW pct Neshan (Snapp)', f'''{yw_var("vw_RA_Navigation")}
{prev_yw("vw_RA_Navigation")}
VAR curr = CALCULATE(AVERAGE(vw_RA_Navigation[pct_Neshan]),
    vw_RA_Navigation[yearweek] = yw, vw_RA_Navigation[platform] = "Snapp")
VAR prev = CALCULATE(AVERAGE(vw_RA_Navigation[pct_Neshan]),
    vw_RA_Navigation[yearweek] = prev_yw, vw_RA_Navigation[platform] = "Snapp")
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Duplicate Neshan/Balad/None/Other for Tapsi. Tapsi-only: pct_InAppNav. Snapp-only: pct_GoogleMap, pct_Waze.')
note('Visual: Stacked bar with navigation app as legend, city on axis.')

# ════════════════════════════════════════════════════════════════════════════
# RA-15  vw_RA_Referral
# ════════════════════════════════════════════════════════════════════════════
heading('17. RA-15 – Referral / Joining Bonus', level=1)
body('View: vw_RA_Referral  |  Excel Page: #16', bold=True)
body('Snapp joining bonus: all drivers. Tapsi joining bonus: joint drivers only.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n_Snapp','INT','respondents with non-null Snapp joining_bonus'),
        ('joining_Snapp','INT','count who got Snapp joining bonus'),
        ('pct_Joining_Snapp','FLOAT',''),
        ('n_Tapsi','INT','joint respondents with non-null Tapsi joining_bonus'),
        ('joining_Tapsi','INT','count who got Tapsi joining bonus'),
        ('pct_Joining_Tapsi','FLOAT',''),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA15 n Snapp', f'''{yw_var("vw_RA_Referral")}
CALCULATE(SUM(vw_RA_Referral[n_Snapp]), vw_RA_Referral[yearweek] = yw)'''),
    ('RA15 pct Joining Snapp', f'''{yw_var("vw_RA_Referral")}
VAR n_val = CALCULATE(SUM(vw_RA_Referral[n_Snapp]), vw_RA_Referral[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Referral[pct_Joining_Snapp]), vw_RA_Referral[yearweek] = yw))'''),
    ('RA15 n Tapsi', f'''{yw_var("vw_RA_Referral")}
CALCULATE(SUM(vw_RA_Referral[n_Tapsi]), vw_RA_Referral[yearweek] = yw)'''),
    ('RA15 pct Joining Tapsi', f'''{yw_var("vw_RA_Referral")}
VAR n_val = CALCULATE(SUM(vw_RA_Referral[n_Tapsi]), vw_RA_Referral[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_Referral[pct_Joining_Tapsi]), vw_RA_Referral[yearweek] = yw))'''),
    ('RA15 WoW pct Joining Snapp', f'''{yw_var("vw_RA_Referral")}
{prev_yw("vw_RA_Referral")}
VAR curr = CALCULATE(AVERAGE(vw_RA_Referral[pct_Joining_Snapp]), vw_RA_Referral[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_Referral[pct_Joining_Snapp]), vw_RA_Referral[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Visual: Matrix with city rows. Cards for national joining pct. Line chart for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# RA-16  vw_RA_TapsiInactivity
# ════════════════════════════════════════════════════════════════════════════
heading('18. RA-16 – Tapsi Inactivity Before Incentive', level=1)
body('View: vw_RA_TapsiInactivity  |  Excel Page: #17', bold=True)
body('Joint drivers only. Long format with inactivity time bucket before receiving the incentive.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('inactivity_bucket','TEXT','"Same Day","1_3 Day Before","3_7 Days Before","8_14 Days Before","15_30 Days_Before","1_2 Month Before","2_3 Month Before","3_6Month Before",">6 Month Before"'),
        ('bucket_sort','INT','1–9 + 99'),
        ('n','INT','count in this bucket'),
        ('n_total','INT','total joint drivers in city × week'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA16 n Total', f'''{yw_var("vw_RA_TapsiInactivity")}
CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]), vw_RA_TapsiInactivity[yearweek] = yw)'''),
    ('RA16 pct Bucket', f'''// Matrix: rows = city, columns = inactivity_bucket (sorted by bucket_sort)
{yw_var("vw_RA_TapsiInactivity")}
VAR n_val = CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]), vw_RA_TapsiInactivity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    DIVIDE(
        CALCULATE(SUM(vw_RA_TapsiInactivity[n]), vw_RA_TapsiInactivity[yearweek] = yw),
        CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]), vw_RA_TapsiInactivity[yearweek] = yw)
    ) * 100)'''),
    ('RA16 pct Same Day', f'''{yw_var("vw_RA_TapsiInactivity")}
VAR n_val = CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]), vw_RA_TapsiInactivity[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    DIVIDE(
        CALCULATE(SUM(vw_RA_TapsiInactivity[n]),
            vw_RA_TapsiInactivity[yearweek] = yw,
            vw_RA_TapsiInactivity[inactivity_bucket] = "Same Day"),
        CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]),
            vw_RA_TapsiInactivity[yearweek] = yw)
    ) * 100)'''),
    ('RA16 WoW pct Same Day', f'''{yw_var("vw_RA_TapsiInactivity")}
{prev_yw("vw_RA_TapsiInactivity")}
VAR curr = DIVIDE(
    CALCULATE(SUM(vw_RA_TapsiInactivity[n]),
        vw_RA_TapsiInactivity[yearweek] = yw,
        vw_RA_TapsiInactivity[inactivity_bucket] = "Same Day"),
    CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]),
        vw_RA_TapsiInactivity[yearweek] = yw)) * 100
VAR prev = DIVIDE(
    CALCULATE(SUM(vw_RA_TapsiInactivity[n]),
        vw_RA_TapsiInactivity[yearweek] = prev_yw,
        vw_RA_TapsiInactivity[inactivity_bucket] = "Same Day"),
    CALCULATE(MAX(vw_RA_TapsiInactivity[n_total]),
        vw_RA_TapsiInactivity[yearweek] = prev_yw)) * 100
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Sort inactivity_bucket column by bucket_sort in Model view.')
note('Visual: Clustered/Stacked bar – inactivity_bucket on x-axis, city as legend or matrix rows.')

# ════════════════════════════════════════════════════════════════════════════
# RA-17  vw_RA_LuckyWheel
# ════════════════════════════════════════════════════════════════════════════
heading('19. RA-17 – Lucky Wheel Usage', level=1)
body('View: vw_RA_LuckyWheel  |  Excel Page: #19', bold=True)
body('wheel column = Rial amount won from Lucky Wheel; 0 = did not use.')
body('Columns:', bold=True)
col_table(
    ['Column','Type','Notes'],
    [
        ('yearweek / yearweek_sort / weeknumber / city','',''),
        ('n','INT','all respondents'),
        ('n_users','INT','drivers who used Lucky Wheel (wheel > 0)'),
        ('pct_usage','FLOAT','% who used the wheel'),
        ('avg_wheel_amount','FLOAT','avg Rial amount among users only'),
    ]
)
doc.add_paragraph()
body('DAX Measures:', bold=True)
dax_section('', [
    ('RA17 n', f'''{yw_var("vw_RA_LuckyWheel")}
CALCULATE(SUM(vw_RA_LuckyWheel[n]), vw_RA_LuckyWheel[yearweek] = yw)'''),
    ('RA17 n Users', f'''{yw_var("vw_RA_LuckyWheel")}
CALCULATE(SUM(vw_RA_LuckyWheel[n_users]), vw_RA_LuckyWheel[yearweek] = yw)'''),
    ('RA17 pct Usage', f'''{yw_var("vw_RA_LuckyWheel")}
VAR n_val = CALCULATE(SUM(vw_RA_LuckyWheel[n]), vw_RA_LuckyWheel[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_LuckyWheel[pct_usage]), vw_RA_LuckyWheel[yearweek] = yw))'''),
    ('RA17 Avg Wheel Amount', f'''{yw_var("vw_RA_LuckyWheel")}
VAR n_val = CALCULATE(SUM(vw_RA_LuckyWheel[n_users]), vw_RA_LuckyWheel[yearweek] = yw)
RETURN IF(n_val < [Min N], BLANK(),
    CALCULATE(AVERAGE(vw_RA_LuckyWheel[avg_wheel_amount]), vw_RA_LuckyWheel[yearweek] = yw))'''),
    ('RA17 WoW pct Usage', f'''{yw_var("vw_RA_LuckyWheel")}
{prev_yw("vw_RA_LuckyWheel")}
VAR curr = CALCULATE(AVERAGE(vw_RA_LuckyWheel[pct_usage]), vw_RA_LuckyWheel[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_LuckyWheel[pct_usage]), vw_RA_LuckyWheel[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
    ('RA17 WoW Avg Wheel Amount', f'''{yw_var("vw_RA_LuckyWheel")}
{prev_yw("vw_RA_LuckyWheel")}
VAR curr = CALCULATE(AVERAGE(vw_RA_LuckyWheel[avg_wheel_amount]), vw_RA_LuckyWheel[yearweek] = yw)
VAR prev = CALCULATE(AVERAGE(vw_RA_LuckyWheel[avg_wheel_amount]), vw_RA_LuckyWheel[yearweek] = prev_yw)
RETURN IF(ISBLANK(prev), BLANK(), curr - prev)'''),
])
note('Visual: Matrix with city rows. Cards for national pct_usage and avg_wheel_amount. Line for WoW.')

# ════════════════════════════════════════════════════════════════════════════
# APPENDIX: General DAX Patterns
# ════════════════════════════════════════════════════════════════════════════
heading('20. Appendix – General DAX Patterns', level=1)
body('Week Slicer (per-view)', bold=True)
note('Each view has its own yearweek column. Add a slicer using the view-specific yearweek column on each report page.')

body('Sorting yearweek on x-axis / slicer', bold=True)
note('In the Model view, select yearweek (TEXT) → Sort by Column → yearweek_sort (INT). Apply to every view.')

body('WoW Pattern', bold=True)
code_block('VAR yw      = IF(HASONEVALUE(View[yearweek]), VALUES(View[yearweek]), MAX(View[yearweek]))')
code_block('VAR prev_yw = CALCULATE(MAX(View[yearweek]), View[yearweek] < yw)')
code_block('VAR curr    = CALCULATE(metric, View[yearweek] = yw)')
code_block('VAR prev    = CALCULATE(metric, View[yearweek] = prev_yw)')
code_block('RETURN IF(ISBLANK(prev), BLANK(), curr - prev)')

body('N-Gate Pattern', bold=True)
code_block('VAR n_val = CALCULATE(SUM(View[n]), View[yearweek] = yw, <optional segment filters>)')
code_block('RETURN IF(n_val < [Min N], BLANK(), <metric formula>)')

body('Long-Format Matrix Pattern (IncentiveAmounts, Duration, Persona, TapsiInactivity)', bold=True)
note('Place the bucket/category column on the matrix column axis. Sort it by the _sort column. Put the pct measure as Values. city on Rows. yearweek slicer filters the whole matrix.')

body('Platform Filtering', bold=True)
note('For UNION ALL views (CommFree, Navigation): hardcode platform = "Snapp" or "Tapsi" in the CALCULATE filter. Use separate measures for each platform rather than relying on a slicer, so N-gates are applied independently.')

body('Cooperation Type Filtering (vw_RA_SatReview)', bold=True)
note('Use a page-level filter on cooperation_type instead of a slicer to avoid accidentally cross-filtering other visuals.')

doc.add_paragraph()
body('End of Document – Driver Survey Power BI Routine Analysis Guide v5 (Complete)', bold=True, color=(80,80,80))

out = r"D:\Work\Driver Survey\PowerBI\PowerBI_Routine_Analysis_Guide_v5_Complete.docx"
doc.save(out)
print(f"Saved: {out}")
