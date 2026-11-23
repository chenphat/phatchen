"""Part 5: E11 supply-chain ESG register and E12 ESG theme map (coverage, not a score)."""
import datetime as dt
from wp_lib import C, A, ref, rng

D = dt.date
E10, E11, E12 = 'E10 ESG KPIs', 'E11 SupplyChain', 'E12 ESG Themes'

THEMES = [  # theme, pillar, ISSB position
    ('Climate change', 'Environmental', 'IFRS S2 in full: governance, strategy, risk management, GHG and cross-industry metrics, targets'),
    ('Biodiversity', 'Environmental', 'Screened - no relevant SRO at own sites (SR-05). ISSB nature-related guidance in development'),
    ('Pollution & resources', 'Environmental', 'Relevant via wash-plant chemicals and wastewater (SR-03); SASB CG-AA chemicals topic'),
    ('Water security', 'Environmental', 'Relevant via groundwater stress at wash plant (SR-03); CDSB water guidance considered'),
    ('Labour standards', 'Social', 'Relevant - order suspension and trade-preference risk (SR-01, SR-04); SASB CG-AA labour topic'),
    ('Human rights & community', 'Social', 'Relevant through labour rights and supplier due diligence (SR-01, SR-07)'),
    ('Health & safety', 'Social', 'Relevant - fire safety, fainting and heat stress (SR-02, CR-02)'),
    ('Customer responsibility', 'Social', 'Relevant only for restricted-substances compliance in products (SR-03); quality metrics outside ISSB'),
    ('Corporate governance', 'Governance', 'Required governance disclosures (S1.26-27, S2.5-7) and remuneration metric (S2.29(g))'),
    ('Risk management', 'Governance', 'Required risk-management disclosures (S1.43-44, S2.24-26)'),
    ('Tax transparency', 'Governance', 'Not an ISSB topic unless linked to a material SRO; GRI 207 is the usual framework'),
    ('Anti-corruption', 'Governance', 'No relevant SRO identified; reported under other frameworks'),
    ('Supply chain - environmental', 'Supply chain', 'Relevant - Scope 3, raw-material sourcing and supplier wastewater (CR-03, CR-07, SR-07); S2.B47-B56 primary-data disclosure'),
    ('Supply chain - social', 'Supply chain', 'Relevant - supplier labour conditions and buyer due-diligence rules (SR-07); SASB CG-AA labour in supply chain'),
]


def build_e11(X):
    ws, r = X.sheet(E11, 'E11', 'Supply-chain ESG register - supplier environmental and social assessment',
                    'Assess tier-1 suppliers for environmental and social risk, show where supply-chain risks are concentrated, and measure coverage of supplier assessments. Supports value-chain disclosures and the extent of primary data used for Scope 3.',
                    'IFRS S1.32(b), S1.B2-B5 (value chain), S1.30, 46; IFRS S2.13(b), S2.B40, B47-B56 (Scope 3 data); Industry-based Guidance / SASB CG-AA-430a and 430b',
                    'HB 2.1.40, 3.3.20, 4.1.140; MG Step 1 (value chain); DC 2.2.2, 2.4.1, 3.2.2',
                    [7, 32, 11, 11, 11, 8, 8, 9, 9, 9, 12, 8, 9, 11, 14, 8, 9, 9, 8, 8, 10, 12, 12], prep='Chan M.', rev='Keo V. (CFO)', tab=C['green'], status='In progress', obj_h=40)
    X.sec(ws, r, '1. Supplier register (top tier-1 suppliers individually assessed; others shown in aggregate)', 23)
    H = ['ID', 'Supplier (fictional)', 'Category', 'Country', 'Spend FY2025 (USD\'000)', '% of spend', 'Wet processing?', 'Code of conduct signed?', 'Country env. risk (H/M/L)', 'Country social risk (H/M/L)',
         'Environmental assessment', 'Higg FEM score', 'Primary GHG data?', 'Wastewater meets ZDHC?', 'Social audit', 'Audit grade (A-E)', 'Open critical findings', 'Corrective action plan agreed?',
         'Env. risk score', 'Social risk score', 'Overall risk', 'Last assessed', 'Linked SRO']
    S_ = [
        ('Mill CN-01 - woven cotton (Zhejiang)', 'Fabric', 'China', 12400, 'Yes', 'Yes', 'M', 'H', 'Verified', 72, 'Yes', 'Yes', 'SLCP verified', 'B', 0, 'N/A', D(2025, 6, 18), 'CR-07, SR-07'),
        ('Mill CN-02 - polyester and blends (Jiangsu)', 'Fabric', 'China', 8900, 'Yes', 'Yes', 'M', 'H', 'Verified', 64, 'Yes', 'Yes', 'amfori BSCI', 'C', 1, 'Yes', D(2025, 4, 9), 'CR-07, SR-07'),
        ('Mill VN-01 - cotton shirting (Dong Nai)', 'Fabric', 'Vietnam', 7600, 'Yes', 'Yes', 'M', 'M', 'Self-assessed', 51, 'No', 'Yes', 'SLCP verified', 'B', 0, 'N/A', D(2025, 8, 27), 'CR-07'),
        ('Mill TH-01 - cotton and recycled polyester (Samut Prakan)', 'Fabric', 'Thailand', 5200, 'Yes', 'Yes', 'L', 'M', 'Verified', 78, 'Yes', 'Yes', 'amfori BSCI', 'A', 0, 'N/A', D(2025, 9, 15), 'CO-02'),
        ('Mill CN-03 - denim (Guangdong)', 'Fabric', 'China', 4300, 'Yes', 'No', 'H', 'H', 'None', None, 'No', 'No', 'None', 'n/a', 0, 'N/A', D(2024, 11, 2), 'SR-07'),
        ('Mill BD-01 - knit fabric (Dhaka)', 'Fabric', 'Bangladesh', 1900, 'Yes', 'Yes', 'H', 'H', 'Self-assessed', 44, 'No', 'No', 'amfori BSCI', 'D', 2, 'No', D(2025, 3, 21), 'SR-07'),
        ('Trims CN-04 - zips and buttons', 'Trims', 'China', 2600, 'No', 'Yes', 'M', 'H', 'Self-assessed', 58, 'No', 'N/A', 'amfori BSCI', 'B', 0, 'N/A', D(2025, 5, 30), 'SR-07'),
        ('Trims KH-01 - labels and tags (Phnom Penh)', 'Trims', 'Cambodia', 1450, 'No', 'Yes', 'M', 'M', 'None', None, 'No', 'N/A', 'BFC', 'B', 0, 'N/A', D(2025, 10, 8), 'SR-07'),
        ('Packaging KH-02 - cartons and polybags', 'Packaging', 'Cambodia', 1100, 'No', 'Yes', 'M', 'M', 'None', None, 'No', 'N/A', 'None', 'n/a', 0, 'N/A', D(2024, 12, 12), 'n/a'),
        ('Chemicals TH-02 - dyes and auxiliaries (JV)', 'Chemicals', 'Thailand', 1250, 'No', 'Yes', 'L', 'L', 'Verified', 81, 'Yes', 'N/A', 'SLCP verified', 'A', 0, 'N/A', D(2025, 7, 4), 'SR-03'),
        ('Logistics KH-03 - road haulage', 'Logistics', 'Cambodia', 1800, 'No', 'Yes', 'M', 'M', 'None', None, 'No', 'N/A', 'None', 'n/a', 0, 'N/A', D(2025, 2, 14), 'n/a'),
        ('Forwarder SG-01 - sea freight', 'Logistics', 'Singapore', 1350, 'No', 'Yes', 'L', 'L', 'Self-assessed', None, 'Yes', 'N/A', 'None', 'n/a', 0, 'N/A', D(2025, 1, 20), 'CR-03'),
        ('Fuelwood KH-04 - boiler wood (Kampong Speu)', 'Energy', 'Cambodia', 95, 'No', 'No', 'H', 'H', 'None', None, 'No', 'N/A', 'None', 'n/a', 0, 'N/A', D(2025, 11, 5), 'CR-04'),
        ('Briquettes KH-05 - rice-husk briquettes', 'Energy', 'Cambodia', 63, 'No', 'Yes', 'M', 'M', 'Self-assessed', None, 'No', 'N/A', 'None', 'n/a', 0, 'N/A', D(2025, 11, 5), 'CR-04'),
        ('Services KH-06 - canteen and security contractor', 'Services', 'Cambodia', 780, 'No', 'Yes', 'M', 'H', 'None', None, 'No', 'N/A', 'None', 'n/a', 1, 'Yes', D(2025, 9, 30), 'SR-07'),
        ('Other tier-1 suppliers (31) - not individually assessed', 'Various', 'Various', 4900, 'No', 'No', 'M', 'M', 'None', None, 'No', 'N/A', 'None', 'n/a', 0, 'N/A', None, 'n/a'),
    ]
    n = len(S_)
    a, b = r + 2, r + 1 + n
    E = f'{A(a, 4)}:{A(b, 4)}'
    rows = []
    for i, s in enumerate(S_):
        rows.append([f'S-{i + 1:02d}', s[0], s[1], s[2], s[3], lambda R: f'=E{R}/SUM({E})', s[4], s[5], s[6], s[7], s[8],
                     s[9] if s[9] is not None else 'n/a', s[10], s[11], s[12], s[13], s[14], s[15],
                     lambda R: f'=IF(I{R}="H",3,IF(I{R}="M",2,1))+IF(K{R}="Verified",0,IF(K{R}="Self-assessed",1,2))+IF(N{R}="No",2,0)',
                     lambda R: f'=IF(J{R}="H",3,IF(J{R}="M",2,1))+IF(O{R}="None",2,IF(OR(P{R}="D",P{R}="E"),1,0))+IF(AND(Q{R}>0,R{R}<>"Yes"),2,0)',
                     lambda R: f'=IF(MAX(S{R},T{R})>=6,"High",IF(MAX(S{R},T{R})>=4,"Medium","Low"))',
                     s[16] if s[16] else 'n/a', s[17]])
    a2, b2 = X.table(ws, r + 1, H, rows, nf={5: '0.0%', 18: '0', 19: '0'}, inp=(1, 2, 3, 6, 7, 8, 9, 10, 12, 13, 14, 15, 17, 21, 22))
    assert (a2, b2) == (a, b)
    ws.set_row(r + 1, 44)
    for rr in range(a, b + 1):
        ws.set_row(rr, 26)
    for j, opts in ((6, ['Yes', 'No']), (7, ['Yes', 'No']), (8, ['H', 'M', 'L']), (9, ['H', 'M', 'L']), (10, ['Verified', 'Self-assessed', 'None']),
                    (12, ['Yes', 'No']), (13, ['Yes', 'No', 'N/A']), (14, ['SLCP verified', 'amfori BSCI', 'WRAP', 'BFC', 'None']), (17, ['Yes', 'No', 'N/A'])):
        X.dv(ws, a, b, j, opts)
    X.cf_status(ws, a, b, 20); X.cf_status(ws, a, b, 13); X.cf_status(ws, a, b, 17)
    X.put(ws, b + 1, 1, 'TOTAL tier-1 spend', bold=True)
    X.put(ws, b + 1, 4, f'=SUM({E})', bold=True, bg_color=C['grey'])
    X.put(ws, b + 1, 5, f'=SUM({A(a, 5)}:{A(b, 5)})', num_format='0%', bold=True)
    X.name('SC_Spend', E11, b + 1, 4)
    X.put(ws, b + 2, 1, 'Check: suppliers in register vs tier-1 count in E1', bold=True)
    X.put(ws, b + 2, 4, f'=IF(COUNTA({A(a, 0)}:{A(b - 1, 0)})+31={X.ad["e1_tier1"]},"Agrees","DIFFERENCE")', bold=True)
    X.note(ws, b + 3, 'Scoring (entity-developed, S1.50): country risk H=3 / M=2 / L=1, plus 2 if not assessed (1 if self-assessed), plus 2 if wastewater fails ZDHC (environmental); plus 2 if no social audit (1 for grade D/E), '
                      'plus 2 if critical findings lack a corrective action plan (social). Overall = the higher of the two: 6+ High, 4-5 Medium, below 4 Low. Country ratings are illustrative - use a documented index.', 23, h=30)
    r = b + 5
    col = lambda j: f'{A(a, j)}:{A(b, j)}'
    X.sec(ws, r, '2. Supply-chain metrics (feed E10 KPIs and the Dashboard)', 8)
    M = [('Tier-1 spend with a verified environmental assessment (Higg FEM verified)', f'=SUMIF({col(10)},"Verified",{E})/SC_Spend', 'SC_Env', '0%', 'IBG / SASB CG-AA-430a.2'),
         ('Fabric spend from mills providing primary GHG data', f'=SUMIFS({E},{col(12)},"Yes",{col(2)},"Fabric")/SUMIF({col(2)},"Fabric",{E})', 'SC_GHG', '0%', 'S2.B47-B49, B56(a)'),
         ('Wet-processing suppliers meeting ZDHC wastewater limits', f'=COUNTIFS({col(6)},"Yes",{col(13)},"Yes")/COUNTIF({col(6)},"Yes")', 'SC_WW', '0%', 'SASB CG-AA-430a.1'),
         ('Tier-1 spend covered by a social audit', f'=SUMIF({col(14)},"<>None",{E})/SC_Spend', 'SC_Soc', '0%', 'SASB CG-AA-430b.1'),
         ('Tier-1 spend with supplier code of conduct signed', f'=SUMIF({col(7)},"Yes",{E})/SC_Spend', 'SC_COC', '0%', 'S1.46(b)'),
         ('Suppliers with critical findings that have an agreed corrective action plan', f'=COUNTIFS({col(16)},">0",{col(17)},"Yes")/COUNTIF({col(16)},">0")', 'SC_CAP', '0%', 'SASB CG-AA-430b.2'),
         ('High-risk suppliers (number)', f'=COUNTIF({col(20)},"High")', 'SC_High', '0', 'S1.32(b)'),
         ('Tier-1 spend with high-risk suppliers', f'=SUMIF({col(20)},"High",{E})/SC_Spend', 'SC_HighSpend', '0%', 'S1.32(b); S2.13(b)')]
    X.hdr(ws, r + 1, ['', 'Metric', '', '', '', 'Value', 'Source / ISSB link', ''])
    for i, (l, f, nm, nf, src) in enumerate(M):
        R = r + 2 + i
        X.put(ws, R, 0, f'SC-{i + 1}')
        ws.merge_range(R, 1, R, 4, l, X.f(border=1, border_color=C['mid'], text_wrap=True))
        X.put(ws, R, 5, f, num_format=nf, bold=True)
        ws.merge_range(R, 6, R, 7, src, X.f(border=1, border_color=C['mid']))
        X.name(nm, E11, R, 5)
    r = r + 2 + len(M) + 1
    X.sec(ws, r, '3. Where supply-chain risk is concentrated (S1.32(b), S2.13(b))', 8)
    X.hdr(ws, r + 1, ['', 'Overall risk', 'Suppliers', 'Spend (USD\'000)', '% of spend', '', '', ''])
    for i, lv in enumerate(['High', 'Medium', 'Low']):
        R = r + 2 + i
        X.put(ws, R, 0, ''); X.put(ws, R, 1, lv, bold=True)
        X.put(ws, R, 2, f'=COUNTIF({col(20)},{A(R, 1)})')
        X.put(ws, R, 3, f'=SUMIF({col(20)},{A(R, 1)},{E})')
        X.put(ws, R, 4, f'={A(R, 3)}/SC_Spend', num_format='0%')
    X.cf_status(ws, r + 2, r + 4, 1)
    ch = X.wb.add_chart({'type': 'column'})
    ch.add_series({'name': 'Spend by supplier risk', 'categories': [E11, r + 2, 1, r + 4, 1], 'values': [E11, r + 2, 3, r + 4, 3], 'gap': 70,
                   'points': [{'fill': {'color': C['red']}}, {'fill': {'color': C['amber']}}, {'fill': {'color': C['green']}}],
                   'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'name': 'Arial', 'size': 8}}})
    ch.set_title({'name': 'Tier-1 spend by supplier risk (USD\'000)', 'name_font': {'name': 'Arial', 'size': 10, 'color': C['navy']}, 'overlay': False})
    ch.set_legend({'none': True}); ch.set_size({'width': 420, 'height': 230})
    ch.set_x_axis({'num_font': {'name': 'Arial', 'size': 8}}); ch.set_y_axis({'num_font': {'name': 'Arial', 'size': 8}, 'num_format': '#,##0', 'major_gridlines': {'visible': True, 'line': {'color': '#E6EAF0'}}})
    ws.insert_chart(r + 1, 9, ch)
    X.note(ws, r + 6, 'Concentration: China supplies over half of tier-1 spend and is rated high social risk; buyer due-diligence rules can block shipments containing affected inputs. '
                      'Three mills can supply primary GHG data that is not yet used in E5 (F5 action). Suppliers beyond tier 1 are mapped only partially (E1: 61).', 8, h=44)
    ws.freeze_panes(10, 2)


def build_e12(X):
    ws, r = X.sheet(E12, 'E12', 'ESG theme map - coverage of the ESG themes and their place in ISSB reporting',
                    'Map the KPI library (E10) to the fourteen themes commonly used in ESG ratings, including supply-chain environmental and social, to show coverage, evidence and which themes belong in the ISSB disclosures. This is a coverage map, not an ESG score: ISSB requires decision-useful information on material risks and opportunities, not a rating.',
                    'IFRS S1.17-18, B13-B30 (materiality, not obscuring information), S1.55-58 (other frameworks); IFRS S2 (climate theme)',
                    'MG Steps 2-5; HB 2.2.30, 2.2.130-2.2.150, 6.1.60 (other frameworks). Theme structure adapted from a common "ESG scores" wheel',
                    [26, 13, 10, 11, 11, 11, 11, 56], prep='Chan M.', tab=C['green'], status='In progress', obj_h=44)
    ka, kb = X.ad['kpi']
    th, inc, ev = rng(E10, 15, ka, kb), rng(E10, 10, ka, kb), rng(E10, 13, ka, kb)
    H = ['ESG theme', 'Pillar', 'KPIs in library', 'In ISSB disclosures', 'Monitored only', 'Evidence on file %', 'Theme in ISSB report?', 'Position under ISSB']
    rows = []
    for t, p, pos in THEMES:
        rows.append([t, p, lambda R: f'=COUNTIF({th},A{R})', lambda R: f'=COUNTIFS({th},A{R},{inc},"Yes*")', lambda R: f'=C{R}-D{R}',
                     lambda R: f'=IF(C{R}=0,0,COUNTIFS({th},A{R},{ev},"Yes")/C{R})', lambda R: f'=IF(D{R}>0,"Yes","No")', pos])
    a, b = X.table(ws, r, H, rows, nf={5: '0%'}, inp=(7,))
    for rr in range(a, b + 1):
        ws.set_row(rr, 30)
    X.cf_status(ws, a, b, 6)
    ws.conditional_format(a, 5, b, 5, {'type': 'data_bar', 'bar_color': '#9BC2E6', 'bar_solid': True, 'min_type': 'num', 'min_value': 0, 'max_type': 'num', 'max_value': 1})
    X.put(ws, b + 1, 0, 'TOTAL', bold=True); X.put(ws, b + 1, 1, '')
    for j in (2, 3, 4):
        X.put(ws, b + 1, j, f'=SUM({A(a, j)}:{A(b, j)})', bold=True)
    X.put(ws, b + 1, 5, f'=COUNTIF({ev},"Yes")/C{b + 2}', num_format='0%', bold=True)
    X.put(ws, b + 1, 6, f'=COUNTIF({A(a, 6)}:{A(b, 6)},"Yes")&" of {len(THEMES)}"', bold=True)
    X.put(ws, b + 1, 7, f'=IF(C{b + 2}=KPI_Total,"All E10 KPIs mapped to a theme","CHECK - some KPIs have no theme")', bold=True)
    X.name('Theme_ISSB', E12, b + 1, 6)
    ch = X.wb.add_chart({'type': 'bar', 'subtype': 'stacked'})
    ch.add_series({'name': 'In ISSB disclosures', 'categories': [E12, a, 0, b, 0], 'values': [E12, a, 3, b, 3], 'fill': {'color': C['teal']}, 'gap': 50})
    ch.add_series({'name': 'Monitored only', 'categories': [E12, a, 0, b, 0], 'values': [E12, a, 4, b, 4], 'fill': {'color': '#C9D1DC'}})
    ch.set_title({'name': 'KPIs by ESG theme', 'name_font': {'name': 'Arial', 'size': 10, 'color': C['navy']}, 'overlay': False})
    ch.set_legend({'position': 'bottom', 'font': {'name': 'Arial', 'size': 8}}); ch.set_size({'width': 520, 'height': 380})
    ch.set_y_axis({'reverse': True, 'num_font': {'name': 'Arial', 'size': 8}}); ch.set_x_axis({'num_font': {'name': 'Arial', 'size': 8}, 'major_gridlines': {'visible': True, 'line': {'color': '#E6EAF0'}}})
    ws.insert_chart(b + 3, 0, ch)
    rd = X.wb.add_chart({'type': 'radar', 'subtype': 'filled'})
    rd.add_series({'name': 'Evidence on file %', 'categories': [E12, a, 0, b, 0], 'values': [E12, a, 5, b, 5], 'fill': {'color': C['blue'], 'transparency': 55}, 'line': {'color': C['blue']}})
    rd.set_title({'name': 'Data readiness by ESG theme (evidence on file)', 'name_font': {'name': 'Arial', 'size': 10, 'color': C['navy']}, 'overlay': False})
    rd.set_legend({'none': True}); rd.set_size({'width': 520, 'height': 380})
    rd.set_y_axis({'min': 0, 'max': 1, 'num_format': '0%', 'num_font': {'name': 'Arial', 'size': 7}, 'major_gridlines': {'visible': True, 'line': {'color': '#D9DEE6'}}})
    rd.set_x_axis({'num_font': {'name': 'Arial', 'size': 7}})
    ws.insert_chart(b + 3, 5, rd)
    X.note(ws, b + 24, 'Why there is no single ESG score here: ESG ratings combine these themes with each provider\'s own weights, so the same company can score differently across providers. IFRS S1 instead asks for material information about the risks and opportunities that could affect the entity\'s prospects (S1.17). '
                       'A theme with "No" in column G is still managed and can be reported under other frameworks (e.g. GRI); adding it to the ISSB report could obscure material information (S1.B25, B30).', 8, h=56)
    ws.freeze_panes(9, 1)
