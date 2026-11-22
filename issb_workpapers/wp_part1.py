"""Part 1: A-series (foundation), B (governance), C (strategy), D (risk management)."""
import datetime as dt
from wp_lib import C, A, ref, rng

D = dt.date
STATUS = ['Complete', 'In progress', 'Not started', 'N/A']


def build_a1(X):
    S = 'A1 Basis'
    ws, r = X.sheet(S, 'A1', 'Entity profile, basis of preparation and key parameters',
                    'Document the reporting entity, the related financial statements, reporting period, presentation currency, sources of guidance, transition reliefs and the statement of compliance. Holds the global parameters used by all other working papers.',
                    'IFRS S1.20-24, S1.54-59, S1.60-73, S1.B38, S1.E1-E6; IFRS S2.C1-C5',
                    'KPMG IFRS Sustainability handbook (HB) 2.1.10, 2.3.50, 6.1.10-6.1.30, 6.2.50; Disclosure checklist (DC) 1.3-1.9, 1.13, App. 3',
                    [38, 26, 30, 44, 40], tab=C['blue'])
    X.sec(ws, r, '1. Key parameters (single source of truth - referenced by other working papers)', 5)
    P = [
        ('Reporting entity', 'Tonle Threads Manufacturing Plc. and its subsidiary', 'Same reporting entity as the consolidated financial statements', 'IFRS S1.20, B38', 'Entity'),
        ('Related financial statements', 'Consolidated FS for the year ended 31-Dec-2025 (CIFRS)', 'Cambodian International Financial Reporting Standards (= IFRS Accounting Standards)', 'IFRS S1.22', None),
        ('Reporting period - start', D(2025, 1, 1), 'Same period as the financial statements', 'IFRS S1.64', None),
        ('Reporting period - end', D(2025, 12, 31), '12-month period', 'IFRS S1.64-66', 'PeriodEnd'),
        ('Status date of these working papers', D(2026, 2, 20), 'KEY INPUT: drives ageing / overdue analysis. Replace with =TODAY() for live tracking', 'Engagement admin', 'StatusDate'),
        ('Planned date of authorisation for issue', D(2026, 3, 27), 'Same date as the financial statements', 'IFRS S1.64, S1.67-68', None),
        ('Presentation currency', 'USD', 'FS presented in USD with KHR translation (Law on Accounting and Auditing)', 'IFRS S1.24', None),
        ('KHR per USD (closing rate, for translation only)', 4025, 'ASSUMPTION - replace with official NBC closing rate', 'IFRS S1.24', 'FX_KHR'),
        ('Revenue FY2025 (USD\'000)', 85400, 'Agrees to consolidated statement of profit or loss - see F3', 'IFRS S1.21(b)(ii), B42', 'Rev25'),
        ('Revenue FY2024 (USD\'000)', 78900, 'Prior-year FS', 'IFRS S1.70', 'Rev24'),
        ('Profit before tax FY2025 (USD\'000)', 5120, 'Agrees to FS - used as a quantitative materiality indicator in A4', 'IFRS S1.B19-B22', 'PBT25'),
        ('Total assets at 31-Dec-2025 (USD\'000)', 61300, 'Agrees to statement of financial position', 'IFRS S1.B42', 'Assets25'),
        ('Property, plant and equipment (USD\'000)', 24800, 'Carrying amount, FS note 9', 'IFRS S2.29(b)-(c)', 'PPE25'),
        ('Inventories (USD\'000)', 11900, 'Carrying amount, FS note 12', 'IFRS S2.29(c)', 'Inv25'),
        ('Capital expenditure FY2025 (USD\'000)', 4800, 'Additions to PPE, FS note 9', 'IFRS S2.29(e)', 'Capex25'),
        ('Employee benefits expense FY2025 (USD\'000)', 27600, 'FS note 6', 'IFRS S2.22', 'StaffCost25'),
        ('Electricity expense FY2025 (USD\'000) - consolidated group', 1165, 'General ledger a/c 6210; used for reasonableness test in F3', 'IFRS S1.B42(c)', 'ElecExp25'),
        ('Cash and undrawn committed facilities (USD\'000)', 9400, 'Cash 5,900 + undrawn 3,500 (FS note 20)', 'IFRS S2.22(a)(iii)(1)', 'Liquidity25'),
        ('Key management personnel remuneration (USD\'000)', 1850, 'FS note 27 (related parties)', 'IFRS S2.29(g)', 'ExecRem25'),
        ('First period of applying IFRS S1 / S2', 'FY2024 (voluntary application)', 'FY2025 is therefore the SECOND reporting period', 'IFRS S1.E1-E2, S2.C1-C2', None),
        ('Industry (SASB / IFRS S2 industry-based guidance)', 'Apparel, Accessories & Footwear (CG-AA)', 'Contract garment manufacturer (cut-make-trim, FOB) for global brands', 'IFRS S1.55(a), 58(a), 59; S2.12, 32', None),
    ]
    X.hdr(ws, r + 1, ['Parameter', 'Value', 'Notes / source', 'ISSB reference', 'Defined name (used in formulas)'])
    rr = r + 2
    for lab, v, note, isb, nm in P:
        X.put(ws, rr, 0, lab, bold=True)
        kw = {'bg_color': C['yellow']} if nm in ('StatusDate', 'FX_KHR') else {}
        if isinstance(v, str):
            kw['font_color'] = C['inp']
        X.put(ws, rr, 1, v, **kw)
        X.put(ws, rr, 2, note)
        X.put(ws, rr, 3, isb)
        X.put(ws, rr, 4, nm or '')
        if nm:
            X.name(nm, S, rr, 1)
        rr += 1
    r = rr + 1
    X.sec(ws, r, '2. Reporting entity and organisational boundary (IFRS S1.20, B38; IFRS S2.29(a)(iv), B27)', 5)
    rows = [
        ('Tonle Threads Manufacturing Plc. - Factory A (Phnom Penh SEZ)', 'Parent', 'Consolidated', 'Included - consolidated accounting group', 'Cutting, sewing, finishing; head office'),
        ('TTM Kampong Speu Co., Ltd. - Factory B (Kampong Speu)', '100% subsidiary', 'Consolidated (IFRS 10)', 'Included - consolidated accounting group', 'Sewing and finishing; located in Prek Thnot river flood plain'),
        ('Speu Wash & Dye JV Co., Ltd. (wash plant)', '50% joint venture', 'Equity method (IFRS 11 / IAS 28)', 'Included at 100% - TTM has operational control under the JV agreement; reported as "other investees"', 'Garment washing / dyeing; highest water and thermal-energy use'),
    ]
    a, b = X.table(ws, r + 1, ['Entity / facility', 'Relationship', 'Treatment in financial statements', 'Treatment for GHG emissions (operational control approach)', 'Activities'], rows, inp=(0, 1, 2, 3, 4))
    r = b + 2
    X.sec(ws, r, '3. Sources of guidance applied (IFRS S1.54-59)', 5)
    rows = [
        ('IFRS S2 Climate-related Disclosures (June 2023)', 'Climate-related risks and opportunities', 'Applied in full', 'IFRS S1.54, 56', 'Dec-2025 amendments to S2 (GHG) effective 1-Jan-2027: NOT early-applied (S2.C1B)'),
        ('Industry-based Guidance on Implementing IFRS S2 - Apparel, Accessories & Footwear', 'Industry-based climate metrics', 'Referred to and applicability considered', 'IFRS S2.12, 32', 'See E8'),
        ('SASB Standard - Apparel, Accessories & Footwear (CG-AA)', 'Non-climate SROs and metrics: labour conditions, chemicals, water / wastewater', 'Referred to and applicability considered ("shall refer to")', 'IFRS S1.55(a), 58(a)', 'See A3, E8'),
        ('GHG Protocol Corporate Standard (2004) and Corporate Value Chain (Scope 3) Standard (2011)', 'GHG measurement', 'Applied', 'IFRS S2.29(a)(ii), B23-B25, B32', 'See E3-E6'),
        ('CDSB Framework application guidance - water', 'Water-related disclosures (wash plant)', 'May refer to - considered', 'IFRS S1.55(b)(i), 58(b)', 'See E8'),
        ('GRI Standards / ESRS', 'Considered only to the extent they meet investor information needs', 'May refer to - considered, no conflict with ISSB', 'IFRS S1.58(c), App. C', 'GRI 403 definitions used for LTIFR'),
        ('Better Factories Cambodia (ILO/IFC) assessment protocol; ZDHC wastewater guidelines; Higg FEM', 'Industry practice for labour and environmental metrics', 'Industry practice considered', 'IFRS S1.55(b)(iii), 57', 'Entity-developed metrics disclosed per S1.50'),
    ]
    a, b = X.table(ws, r + 1, ['Source', 'Used for', 'How used', 'ISSB reference', 'Comment / WP link'], rows, inp=(0, 1, 2, 4))
    r = b + 2
    X.sec(ws, r, '4. Transition reliefs register (IFRS S1.E3-E6; IFRS S2.C3-C5; DC Appendix 3)', 5)
    rows = [
        ('Comparative information not required in first year', 'IFRS S1.E3; S2.C3', 'Used in FY2024', 'FY2025: comparatives now required for climate metrics (Scope 1 and 2) - see E6', 'Complete'),
        ('Reporting after the financial statements in first year', 'IFRS S1.E4', 'Not used', 'Sustainability disclosures issued with the FS in both years', 'N/A'),
        ('"Climate-first" - only climate-related SROs in first year', 'IFRS S1.E5', 'Used in FY2024', 'FY2025 is the first year covering non-climate SROs (labour, OHS, water). Fact of using relief disclosed in FY2024', 'Complete'),
        ('No comparatives for non-climate SROs in second year', 'IFRS S1.E6', 'USED in FY2025', 'No FY2024 comparatives for non-climate metrics in E8', 'Complete'),
        ('Continue existing GHG method (not GHG Protocol) in first year', 'IFRS S2.C4(a)', 'Not used', 'GHG Protocol applied from FY2024', 'N/A'),
        ('Scope 3 GHG emissions not required in first year', 'IFRS S2.C4(b)', 'Used in FY2024', 'Scope 3 measured for the first time in FY2025 (E5)', 'Complete'),
        ('No Scope 3 comparatives in second year', 'IFRS S2.C5', 'USED in FY2025', 'Scope 3 FY2024 comparative shown as "n/a" in E6', 'Complete'),
    ]
    a, b = X.table(ws, r + 1, ['Relief', 'ISSB reference', 'Election', 'Effect on FY2025 disclosures', 'WP status'], rows, inp=(2, 3, 4))
    X.dv(ws, a, b, 4, STATUS)
    X.cf_status(ws, a, b, 4)
    r = b + 2
    X.sec(ws, r, '5. Statement of compliance, location and timing (IFRS S1.60-73)', 5)
    rows = [
        ('Statement of compliance', 'IFRS S1.72-73', 'An explicit and unreserved statement may be made ONLY if all requirements are met. Link: F1 checklist must show 100% complete for applicable items.', "=IF(Chk_Pct=1,\"Unreserved statement supportable\",\"NOT yet supportable - open checklist items remain\")", 'In progress'),
        ('Location of disclosures', 'IFRS S1.60-63, B45-B47', 'Separate "Sustainability-related financial disclosures" section of the annual report, cross-referenced to FS notes (same terms, same time).', 'Annual report pp. 48-79 (draft)', 'Complete'),
        ('Timing of reporting', 'IFRS S1.64-69', 'Issued at the same time as FS; events after the reporting period assessed to authorisation date (S1.67-68).', 'Subsequent-events review scheduled in F5 (A-018)', 'In progress'),
        ('Local law / regulation', 'IFRS S1.B31-B33; S1.73', 'Confirm the current status of ISSB adoption with the Accounting and Auditing Regulator (ACAR) and any SERC / CSX listing requirements.', 'ASSUMPTION: voluntary application - verify before finalisation', 'In progress'),
        ('Commercially sensitive opportunity information', 'IFRS S1.B34-B37', 'Exemption considered for CO-02 (pricing of preferred-supplier programme): NOT used - information disclosed at aggregated level.', 'See C1', 'Complete'),
    ]
    a, b = X.table(ws, r + 1, ['Matter', 'ISSB reference', 'Consideration', 'Conclusion / evidence', 'WP status'], rows, inp=(2, 3, 4))
    X.dv(ws, a, b, 4, STATUS)
    X.cf_status(ws, a, b, 4)
    ws.freeze_panes(8, 1)


def build_a2(X):
    S = 'A2 ValueChain'
    ws, r = X.sheet(S, 'A2', 'Value chain map - impacts and dependencies on resources and relationships',
                    'Step 1 of the materiality process: understand the population of impacts and dependencies across the value chain, as the starting point for identifying sustainability-related risks and opportunities (SROs).',
                    'IFRS S1.2, S1.B1-B5, S1.32, S1.B6; IFRS S2.13',
                    'Materiality How-to guide (MG) Step 1 and illustrative example (TerraBrew) Step 1; HB 2.1.40, 2.2.20',
                    [6, 14, 34, 22, 44, 44, 22, 28], prep='Chan M.', tab=C['blue'])
    H = ['#', 'Value-chain stage', 'Activity', 'Geography', 'IMPACTS on resources / relationships', 'DEPENDENCIES on resources / relationships', 'Capital / resource type', 'Information source']
    rows = [
        (1, 'Upstream', 'Fabric and yarn sourcing (woven cotton, polyester, blends) - nominated mills', 'China, Vietnam, Thailand', 'Natural: GHG from spinning / weaving / dyeing, water use, chemical discharge. Social: labour conditions at mills', 'Natural: cotton yields, water availability, climate stability. Financial: stable input prices. Relationship: buyer-nominated mills', 'Natural; social & relationship; financial', 'Purchase ledger; mill Higg FEM data; buyer nomination lists'),
        (2, 'Upstream', 'Trims, accessories and packaging', 'China, Cambodia', 'Natural: plastics and paper use, packaging waste', 'Manufactured: reliable supply; Financial: price stability', 'Natural; manufactured', 'Purchase ledger'),
        (3, 'Upstream', 'Inbound logistics - sea freight to Sihanoukville, road haulage to factories', 'Regional sea lanes; NR4 / NR3', 'Natural: GHG emissions, air pollution. Social: driver safety', 'Manufactured: port and road infrastructure (flood-exposed), customs clearance', 'Manufactured; natural', 'Forwarder invoices; bills of lading'),
        (4, 'Upstream', 'Energy supply - EDC grid electricity, wood biomass for boilers, diesel', 'Cambodia', 'Natural: grid emissions (coal / hydro mix), forest degradation risk from fuelwood, air emissions', 'Natural: hydrology (hydropower), forests. Manufactured: grid reliability. Financial: tariff stability', 'Natural; manufactured; financial', 'EDC bills; wood delivery notes; fuel cards'),
        (5, 'Own operations', 'Cutting, sewing, finishing - Factory A (Phnom Penh SEZ)', 'Phnom Penh', 'Natural: energy use, fabric off-cut waste. Social: employment of ~3,700 workers (84% women), wages, working hours, OHS', 'Human: skilled sewing operators, supervisors. Social: labour relations, union dialogue. Manufactured: machinery', 'Human; social & relationship; manufactured', 'HR system; production system; BFC reports'),
        (6, 'Own operations', 'Sewing and finishing - Factory B (Kampong Speu)', 'Kampong Speu (Prek Thnot flood plain)', 'As above; plus community employment in rural district', 'Natural: flood-free site, tolerable working temperatures. Human: local labour pool', 'Human; natural; manufactured', 'Flood-risk study (2025); insurance survey'),
        (7, 'Own operations (JV)', 'Washing and dyeing - Speu Wash & Dye JV (operational control)', 'Kampong Speu', 'Natural: high water abstraction, wastewater and chemical discharge, boiler emissions', 'Natural: groundwater availability and quality. Regulatory: discharge permits (MoE). Relationship: ZDHC conformance demanded by buyers', 'Natural; social & relationship', 'Water meters; wastewater test reports; MoE permit'),
        (8, 'Downstream', 'Outbound logistics - FOB delivery to port; buyer-controlled sea / air freight', 'Sihanoukville to EU, US, Japan', 'Natural: GHG emissions (air freight when orders run late)', 'Manufactured: port operations. Financial: freight rates; trade-preference access', 'Manufactured; financial', 'Shipping documents; buyer logistics data'),
        (9, 'Downstream', 'Customers - 11 global apparel brands (top 3 = 71% of revenue)', 'EU 46%, US 31%, Japan 14%, other 9%', 'Social: brand reputation exposure to supplier conduct', 'Relationship: preferred-supplier status; brand science-based targets cascade to suppliers; social-compliance audit results', 'Social & relationship; financial', 'Customer contracts; supplier codes of conduct; scorecards'),
        (10, 'Downstream', 'Product use and end-of-life', 'Consumer markets', 'Natural: use-phase energy / water (laundering), textile waste', 'Regulatory: EU textile EPR, eco-design and due-diligence rules influence buyer specifications', 'Natural; social & relationship', 'Buyer product specifications; regulatory horizon scan'),
        (11, 'Cross-cutting', 'Finance providers and investors', 'Cambodia / regional banks; CSX investors', 'Financial: capital allocation to transition', 'Financial: access to and cost of capital; sustainability-linked loan covenants', 'Financial', 'Loan agreements; investor Q&A log'),
    ]
    a, b = X.table(ws, r, H, rows, inp=tuple(range(1, 8)))
    X.note(ws, b + 2, 'Reassessment trigger (IFRS S1.B11): the value-chain map is reassessed on a significant event or change in circumstances (e.g. new facility, major customer change, the October 2025 flood). Last reassessed: November 2025.', 8, h=28)
    ws.freeze_panes(9, 3)


SRO = [
    # id, title, description, topic, std, R/O, type, value chain, horizon, L, M, relevant, source, metrics WP
    ('CR-01', 'River and pluvial flooding - Factory B', 'Prek Thnot river overflow and intense rainfall interrupt production, damage inventory and PPE (Oct-2025 event: 9 production days lost)', 'Climate', 'IFRS S2', 'Risk', 'Physical - acute', 'Own operations', 'Short', 4, 4, 'Yes', 'IFRS S2.10(b); flood-risk study', 'C2, C3, E7'),
    ('CR-02', 'Heat stress on workforce productivity and health', 'Rising wet-bulb temperatures reduce sewing-line efficiency, increase fainting incidents and cooling costs', 'Climate', 'IFRS S2', 'Risk', 'Physical - chronic', 'Own operations', 'Medium', 4, 3, 'Yes', 'IFRS S2.10(b); HB 4.1.30', 'C3, E8'),
    ('CR-03', 'Buyer decarbonisation requirements', 'Brands with science-based targets require suppliers to cut Scope 1 and 2 emissions and phase out coal / unsustainable biomass; non-conforming suppliers lose order allocation', 'Climate', 'IFRS S2', 'Risk', 'Transition - market', 'Downstream', 'Short', 4, 5, 'Yes', 'Customer supplier codes; IFRS S2.10', 'E6, E7, E9'),
    ('CR-04', 'Biomass (fuelwood) boiler fuel - sourcing restrictions', 'Forest-wood use for boilers faces tightening enforcement and buyer deforestation-free policies; switching cost to electric / agri-residue boilers', 'Climate', 'IFRS S2', 'Risk', 'Transition - policy & legal', 'Own operations / upstream', 'Medium', 4, 3, 'Yes', 'IFRS S2 IBG CG-AA Raw Materials Sourcing', 'C1, E9'),
    ('CR-05', 'Grid electricity cost, reliability and carbon intensity', 'Tariff increases and outages (diesel back-up); rising coal share in grid mix increases Scope 2', 'Climate', 'IFRS S2', 'Risk', 'Transition - market', 'Upstream / own operations', 'Short', 3, 3, 'Yes', 'EDC tariff schedule; IFRS S2.10', 'E4, C3'),
    ('CR-06', 'Carbon pricing and product regulation in export markets', 'EU due-diligence, eco-design and textile EPR rules and possible carbon pricing pass-through alter buyer specifications and costs', 'Climate', 'IFRS S2', 'Risk', 'Transition - policy & legal', 'Downstream', 'Long', 3, 3, 'Yes', 'Regulatory horizon scan', 'C3, E7'),
    ('CR-07', 'Climate-driven raw-material supply disruption', 'Cotton yield volatility and mill disruptions (floods, drought, heat) raise fabric cost and lead times', 'Climate', 'IFRS S2', 'Risk', 'Physical - chronic (value chain)', 'Upstream', 'Medium', 3, 3, 'Yes', 'IFRS S2 IBG CG-AA-440a.3', 'E8'),
    ('CR-08', 'Sea-level rise affecting port access (Sihanoukville)', 'Long-term disruption of port logistics', 'Climate', 'IFRS S2', 'Risk', 'Physical - chronic', 'Upstream / downstream', 'Long', 2, 2, 'No', 'Scenario screening (C3)', 'n/a'),
    ('CO-01', 'Rooftop solar and energy efficiency', 'Self-generation and efficient motors / LED / heat recovery lower energy cost and emissions', 'Climate', 'IFRS S2', 'Opportunity', 'Resource efficiency / energy source', 'Own operations', 'Short', 5, 3, 'Yes', 'IFRS S2.10(a); energy audit 2024', 'E7, E9'),
    ('CO-02', 'Preferred low-carbon supplier status', 'Higher order allocation and longer contracts from brands consolidating to verified low-carbon suppliers; recycled-content product lines', 'Climate', 'IFRS S2', 'Opportunity', 'Products & services / markets', 'Downstream', 'Medium', 3, 4, 'Yes', 'Buyer scorecards', 'E7'),
    ('SR-01', 'Labour conditions and worker rights', 'Wages, overtime, freedom of association; adverse audit findings lead to order suspension; trade-preference conditionality', 'Human capital / labour', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations', 'Short', 3, 5, 'Yes', 'SASB CG-AA Labour Conditions in the Supply Chain', 'E8'),
    ('SR-02', 'Occupational health and safety', 'Fire and building safety, mass-fainting events, commuting accidents on collective transport', 'Health & safety', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations', 'Short', 3, 4, 'Yes', 'S1.55(b)(iii) industry practice (BFC); GRI 403', 'E8'),
    ('SR-03', 'Water use, wastewater and chemicals - wash plant', 'Groundwater stress and discharge non-compliance (ZDHC / MoE limits) risk permit suspension and buyer delisting', 'Water & chemicals', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations (JV)', 'Medium', 3, 4, 'Yes', 'SASB CG-AA Environmental Impacts in the Supply Chain; CDSB water guidance', 'E8'),
    ('SR-04', 'Workforce retention and skills', 'High turnover of sewing operators raises training cost and lowers line efficiency', 'Human capital / labour', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations', 'Medium', 4, 2, 'Yes', 'Management KPI', 'E8'),
    ('SO-01', 'Verified social-compliance leadership', 'Strong BFC / buyer audit results support premium customers and access to sustainability-linked finance', 'Human capital / labour', 'IFRS S1', 'Opportunity', 'n/a (non-climate)', 'Downstream / finance', 'Medium', 3, 3, 'Yes', 'Loan term sheet; buyer scorecards', 'E8'),
    ('SR-05', 'Biodiversity and land use at factory sites', 'Sites are in established industrial zones; no significant dependency or impact identified', 'Nature', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations', 'Long', 1, 2, 'No', 'Screening only', 'n/a'),
    ('SR-06', 'Data privacy and cyber security', 'Limited personal data (employees only); managed within IT general controls', 'Governance', 'IFRS S1', 'Risk', 'n/a (non-climate)', 'Own operations', 'Short', 2, 2, 'No', 'Screening only', 'n/a'),
]


def build_a3(X):
    S = 'A3 SRO Register'
    ws, r = X.sheet(S, 'A3', 'Register of sustainability- and climate-related risks and opportunities (SROs)',
                    'Steps 2-3 of the materiality process: identify potentially relevant SROs from impacts and dependencies (A2), assess them using the criteria in D1 and determine which could reasonably be expected to affect the entity\'s prospects (cash flows, access to finance, cost of capital).',
                    'IFRS S1.3, S1.30-31, S1.54-55, S1.B1-B12; IFRS S2.10-12; IFRS S1.44(a), S2.25(a)',
                    'MG Steps 2-3; HB 2.2.10-2.2.60, 3.3.10, 4.1.30; DC 2.2.1, 3.2.1',
                    [8, 26, 46, 16, 9, 11, 22, 18, 10, 9, 9, 8, 10, 11, 30, 12, 34], prep='Chan M.', tab=C['blue'], obj_h=40)
    H = ['SRO ID', 'Risk / opportunity', 'Description', 'Topic', 'Standard', 'Risk or opportunity', 'Type (S2.10(b)) ', 'Where concentrated (S1.32(b) / S2.13(b))',
         'Time horizon (S1.30(b) / S2.10(c))', 'Likelihood (1-5)', 'Magnitude (1-5)', 'Score', 'Rating', 'Relevant SRO? (judgement)', 'Source of guidance / evidence', 'Linked WPs', 'Rationale for conclusion']
    rows = []
    for s in SRO:
        rat = ('Could reasonably be expected to affect prospects - carried to A4 and core-content WPs' if s[11] == 'Yes'
               else 'Not reasonably expected to affect prospects over any horizon - documented, monitored annually (S1.B11)')
        rows.append(list(s[:9]) + [s[9], s[10], lambda R: f'=J{R}*K{R}',
                                   lambda R: f'=IF(L{R}>=15,"Critical",IF(L{R}>=10,"High",IF(L{R}>=5,"Medium","Low")))',
                                   s[11], s[12], s[13], rat])
    a, b = X.table(ws, r, H, rows, inp=tuple(range(1, 9)) + (13, 14, 15, 16))
    X.ad['sro'] = (S, a, b)
    X.dv(ws, a, b, 8, ['Short', 'Medium', 'Long'])
    X.dv(ws, a, b, 13, ['Yes', 'No'])
    X.dv(ws, a, b, 5, ['Risk', 'Opportunity'])
    X.cf_status(ws, a, b, 12)
    X.cf_status(ws, a, b, 13)
    r = b + 2
    X.sec(ws, r, 'Summary (feeds Dashboard)', 6)
    lab = ['Total SROs identified', 'Relevant SROs (reasonably expected to affect prospects)', '  - of which climate-related (IFRS S2)', '  - of which other sustainability topics (IFRS S1)', '  - risks', '  - opportunities']
    fx = [f'=COUNTA({A(a, 0)}:{A(b, 0)})', f'=COUNTIF({A(a, 13)}:{A(b, 13)},"Yes")',
          f'=COUNTIFS({A(a, 13)}:{A(b, 13)},"Yes",{A(a, 4)}:{A(b, 4)},"IFRS S2")',
          f'=COUNTIFS({A(a, 13)}:{A(b, 13)},"Yes",{A(a, 4)}:{A(b, 4)},"IFRS S1")',
          f'=COUNTIFS({A(a, 13)}:{A(b, 13)},"Yes",{A(a, 5)}:{A(b, 5)},"Risk")',
          f'=COUNTIFS({A(a, 13)}:{A(b, 13)},"Yes",{A(a, 5)}:{A(b, 5)},"Opportunity")']
    for i, (l, f) in enumerate(zip(lab, fx)):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 2, l, X.f(border=1, border_color=C['mid'], bold=(i < 2)))
        X.put(ws, r + 1 + i, 3, f, bold=True)
    X.name('SRO_Total', S, r + 1, 3)
    X.name('SRO_Relevant', S, r + 2, 3)
    r += 8
    X.sec(ws, r, 'Time-horizon definitions (IFRS S1.30(c), S2.10(d)) - linked to planning horizons', 6)
    rows = [('Short', '0-2 years', 'Annual budget and order book cycle; customer contracts typically 1-2 seasons'),
            ('Medium', '2-5 years', 'Five-year business plan; payback period of energy capex; loan tenor'),
            ('Long', '5-15 years', 'Useful life of buildings and boilers; land-lease term; buyers\' 2030-2040 climate targets')]
    X.table(ws, r + 1, ['Horizon', 'Period', 'Link to planning horizon used for strategic decision-making'], [(x[0], x[1], x[2]) for x in rows], inp=(1, 2))
    ws.merge_range(r + 1, 2, r + 1, 5, 'Link to planning horizon used for strategic decision-making', X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], border=1, text_wrap=True, valign='vcenter'))
    for i, x in enumerate(rows):
        ws.merge_range(r + 2 + i, 2, r + 2 + i, 5, x[2], X.f(border=1, border_color=C['mid'], font_color=C['inp'], text_wrap=True))
    ws.freeze_panes(9, 2)


def build_a4(X):
    S = 'A4 Materiality'
    ws, r = X.sheet(S, 'A4', 'Materiality assessment - identifying and assessing material information',
                    'Steps 4-5 of the materiality process: for each relevant SRO, identify potentially material information and judge whether omitting, misstating or obscuring it could reasonably be expected to influence decisions of primary users. Quantitative indicators are a screen only - qualitative factors and entity-specific context drive the conclusion.',
                    'IFRS S1.17-19, S1.B13-B28 (materiality), S1.B29-B30 (aggregation), S1.B19-B24 (judgement); IFRS Practice Statement 2 by analogy',
                    'MG Steps 4-6 and illustrative example; HB 2.2.70-2.2.150; DC "Key concepts - Materiality"',
                    [8, 9, 40, 20, 13, 12, 13, 40, 11, 11, 40, 20], prep='Chan M.', tab=C['blue'], obj_h=44)
    X.sec(ws, r, 'Quantitative indicators (screening only - not a bright line; IFRS S1.B19-B22)', 6)
    X.put(ws, r + 1, 0, 'Indicator 1'); ws.merge_range(r + 1, 1, r + 1, 3, '5% of profit before tax (USD\'000)', X.f(border=1, border_color=C['mid']))
    X.put(ws, r + 1, 4, '=5%*PBT25')
    X.put(ws, r + 2, 0, 'Indicator 2'); ws.merge_range(r + 2, 1, r + 2, 3, '0.5% of revenue (USD\'000)', X.f(border=1, border_color=C['mid']))
    X.put(ws, r + 2, 4, '=0.5%*Rev25')
    X.put(ws, r + 3, 0, 'Screen used'); ws.merge_range(r + 3, 1, r + 3, 3, 'Lower of the two indicators (USD\'000)', X.f(border=1, border_color=C['mid'], bold=True))
    X.put(ws, r + 3, 4, f'=MIN({A(r + 1, 4)},{A(r + 2, 4)})', bold=True, bg_color=C['yellow'])
    thr = A(r + 3, 4)
    r += 5
    H = ['Info ID', 'Linked SRO', 'Potentially material information item (Step 4)', 'ISSB requirement', 'Est. annual financial effect (USD\'000)', '% of PBT', 'Exceeds quantitative screen?',
         'Qualitative factors considered (Step 5)', 'Investor interest (H/M/L)', 'MATERIAL? (judgement)', 'Rationale for conclusion', 'Report section (Step 6)']
    I = [
        ('M-01', 'CR-01', 'Flood exposure of Factory B: current-period loss, share of assets and revenue exposed, adaptation capex', 'IFRS S2.13, 16, 29(c), 29(e)', 835, 'Entity-specific; recurrent hazard; affects going-concern-neutral but covenant headroom; trend worsening under scenarios', 'H', 'Yes', 'Quantitatively and qualitatively material; event occurred in period', 'Strategy; Metrics'),
        ('M-02', 'CR-02', 'Heat-stress effect on productivity and planned cooling investment', 'IFRS S2.13, 14(a)(ii), 16(d)', 310, 'Affects 84%-female workforce; links to OHS (SR-02); trend certain under all scenarios', 'M', 'Yes', 'Above screen and connected to workforce disclosures', 'Strategy'),
        ('M-03', 'CR-03', 'Share of revenue subject to buyer decarbonisation requirements and progress against them', 'IFRS S2.13, 29(b), 33-36', 9520, 'Customer concentration (top 3 = 71%); could alter business model; highly decision-useful', 'H', 'Yes', 'Largest potential effect on prospects', 'Strategy; Metrics; Targets'),
        ('M-04', 'CR-04', 'Biomass boiler transition: fuel mix, certified share, replacement capex', 'IFRS S2.14(a)(i)-(iv), 29(e), 32', 620, 'Regulatory and reputational sensitivity; buyers ask directly', 'H', 'Yes', 'Core element of the transition plan', 'Strategy; Industry metrics'),
        ('M-05', 'CR-05', 'Electricity cost / reliability exposure and Scope 2 intensity', 'IFRS S2.16, 29(a)', 260, 'Combined with CO-01 for disclosure (S1.B29 aggregation of similar items)', 'M', 'Yes', 'Material in aggregate with energy transition information', 'Strategy; Metrics'),
        ('M-06', 'CR-06', 'Anticipated effects of export-market regulation and carbon pricing', 'IFRS S2.15(b), 16(c)-(d), 19-21', 0, 'High measurement uncertainty; qualitative disclosure plus scenario range (C3)', 'M', 'Yes', 'Qualitatively material; quantitative single estimate not useful (S2.19(b))', 'Strategy; Resilience'),
        ('M-07', 'CR-07', 'Priority raw materials, climate threats to sourcing and mitigation', 'IFRS S2.32; IBG CG-AA-440a.3/4', 180, 'Fabric = 62% of cost of sales; buyer-nominated mills limit control', 'M', 'Yes', 'Industry-based metric considered applicable', 'Industry metrics'),
        ('M-08', 'All climate', 'Scope 1, 2 and 3 GHG emissions, methods and emission factors', 'IFRS S2.29(a), B19-B57', 0, 'Required cross-industry metric; material by nature for an entity with climate-related risks', 'H', 'Yes', 'Qualitative factor: information by its nature is expected to be material', 'Metrics'),
        ('M-09', 'CO-01', 'Solar and efficiency programme: capex deployed and savings', 'IFRS S2.14, 29(d)-(e)', 148, 'Shows execution of transition plan', 'M', 'Yes', 'Disclosed with M-05 (aggregated)', 'Strategy; Metrics'),
        ('M-10', 'CO-02', 'Revenue aligned with low-carbon / recycled programmes', 'IFRS S2.29(d); S1.B34-B37', 1998, 'Commercial sensitivity assessed - disclosed at aggregated level', 'M', 'Yes', 'Decision-useful growth indicator', 'Metrics'),
        ('M-11', 'SR-01', 'Labour-conditions audit results, non-conformance rates, wage and overtime indicators', 'IFRS S1.30, 46, 55(a)', 4270, 'Order suspension risk; investor and lender covenants; trade-preference linkage', 'H', 'Yes', 'Could change order allocation - material', 'Other topics'),
        ('M-12', 'SR-02', 'Injury rates, fainting incidents, fire-safety remediation', 'IFRS S1.30, 46, 50', 210, 'Human-life severity; reputational', 'H', 'Yes', 'Qualitatively material irrespective of amount (MG Example 19)', 'Other topics'),
        ('M-13', 'SR-03', 'Water withdrawal, wastewater conformance, chemical management at wash plant', 'IFRS S1.30, 46, 55(a)-(b)', 340, 'Permit suspension would halt JV; buyer delisting', 'M', 'Yes', 'Material - licence to operate', 'Other topics'),
        ('M-14', 'SR-04', 'Workforce turnover rate and training hours', 'IFRS S1.46(b)', 190, 'Managed operationally; no evidence it would influence investor decisions separately', 'L', 'No', 'Not material on its own - turnover rate disclosed within SR-01 narrative only', 'Not separately disclosed'),
        ('M-15', 'SO-01', 'Sustainability-linked loan KPIs and margin ratchet', 'IFRS S1.33, 35', 45, 'Small amount; but connects sustainability performance to cost of capital', 'M', 'Yes', 'Qualitative: direct link to financing cost', 'Other topics; FS note 20'),
    ]
    rows = []
    for x in I:
        rows.append([x[0], x[1], x[2], x[3], x[4], lambda R: f'=IF(PBT25=0,0,E{R}/PBT25)',
                     lambda R, t=thr: f'=IF(E{R}>={t},"Yes","No")', x[5], x[6], x[7], x[8], x[9]])
    a, b = X.table(ws, r, H, rows, nf={5: '0.0%'}, inp=(2, 3, 7, 8, 9, 10, 11))
    X.dv(ws, a, b, 9, ['Yes', 'No'])
    X.dv(ws, a, b, 8, ['H', 'M', 'L'])
    X.cf_status(ws, a, b, 9)
    X.cf_status(ws, a, b, 6)
    X.put(ws, b + 1, 2, 'Information items assessed as material', bold=True)
    X.put(ws, b + 1, 9, f'=COUNTIF({A(a, 9)}:{A(b, 9)},"Yes")', bold=True)
    X.name('Mat_Items', S, b + 1, 9)
    X.note(ws, b + 3, 'Note: an item below the quantitative screen can still be material because of qualitative factors (e.g. M-12 safety, M-08 GHG). Conversely, immaterial information need not be disclosed even if a standard specifies it (IFRS S1.B25). Materiality judgements are reassessed at each reporting date (IFRS S1.B28).', 12, h=30)
    ws.freeze_panes(a, 3)


def req_table(X, ws, r, rows, title, ncols=7):
    """Core-content requirement table: ref | requirement | client information | evidence | owner | status | checklist ID."""
    X.sec(ws, r, title, ncols)
    a, b = X.table(ws, r + 1, ['ISSB reference', 'Disclosure requirement (summarised)', 'Client information gathered / draft disclosure content', 'Evidence obtained', 'Information owner', 'Status', 'Comment / follow-up'],
                   rows, inp=(2, 3, 4, 5, 6))
    X.dv(ws, a, b, 5, STATUS)
    X.cf_status(ws, a, b, 5)
    return b + 2


def build_b1(X):
    S = 'B1 Governance'
    ws, r = X.sheet(S, 'B1', 'Governance - oversight bodies and management\'s role',
                    'Gather and evidence the information needed to disclose the governance processes, controls and procedures used to monitor, manage and oversee sustainability- and climate-related risks and opportunities. Governance is disclosed on an integrated basis to avoid duplication (IFRS S2.7).',
                    'IFRS S1.26-27; IFRS S2.5-7; IFRS S1.B42(b)',
                    'HB 3.2.10-3.2.20; DC 2.1 and 3.1',
                    [16, 44, 56, 34, 20, 13, 30], prep='Ly S.', tab=C['teal'])
    rows = [
        ('S2.6(a) / S1.27(a)', 'Identify the governance body(s) or individual(s) responsible for oversight', 'Board of Directors (7 members; 3 independent). Delegated to the Audit, Risk & Sustainability Committee (ARSC) - 3 members, chaired by an independent director.', 'Board charter; ARSC terms of reference (rev. Mar-2025)', 'Company Secretary', 'Complete', ''),
        ('S2.6(a)(i) / S1.27(a)(i)', 'How responsibilities are reflected in terms of reference, mandates, role descriptions and policies', 'ARSC ToR s.4.3 covers climate and sustainability risk, the transition plan and ISSB reporting. Sustainability Policy approved Mar-2025.', 'ARSC ToR; Sustainability Policy', 'Company Secretary', 'Complete', ''),
        ('S2.6(a)(ii) / S1.27(a)(ii)', 'How the body determines whether appropriate skills and competencies are available or will be developed', 'Annual board skills matrix: 2 of 7 directors have sustainability / climate competence; ISSB training held Jun-2025 (4 hrs). External adviser retained for scenario analysis.', 'Skills matrix 2025; training attendance record', 'Company Secretary', 'Complete', ''),
        ('S2.6(a)(iii) / S1.27(a)(iii)', 'How and how often the body is informed', 'ARSC receives a quarterly sustainability dashboard (GHG, energy, safety, audit results); Board receives ARSC report semi-annually. See meeting log below.', 'ARSC packs Q1-Q4 2025', 'Head of Sustainability', 'Complete', ''),
        ('S2.6(a)(iv) / S1.27(a)(iv)', 'How SROs are taken into account in overseeing strategy, major transactions and risk management, including trade-offs', 'Board approved solar and flood-defence capex (May-2025) using an internal carbon price; trade-off considered: electric boilers raise Scope 2 and tariff exposure but remove fuelwood risk.', 'Board minutes 22-May-2025; capex paper', 'CFO', 'Complete', 'Trade-off narrative to be cross-referred to C1'),
        ('S2.6(a)(v) / S1.27(a)(v)', 'How the body oversees target-setting and monitors progress, including links to remuneration', 'Board approved 2030 targets (E9) in Mar-2024; progress reviewed semi-annually. 10% of the executive bonus pool is linked to Scope 1+2 reduction and LTIFR.', 'Remuneration policy 2025; board minutes', 'HR Director', 'In progress', 'Remuneration committee minutes for FY2025 outcome outstanding (A-007)'),
        ('S2.6(b)(i) / S1.27(b)(i)', 'Whether management\'s role is delegated to a specific position or committee and how oversight is exercised', 'Sustainability Steering Committee (CEO chair; CFO, COO, HR, Head of Sustainability) meets monthly; reports to ARSC.', 'SSC charter; minutes (12 meetings)', 'Head of Sustainability', 'Complete', ''),
        ('S2.6(b)(ii) / S1.27(b)(ii)', 'Whether management uses controls and procedures, and how they are integrated with other internal functions', 'Sustainability data controls manual (v1.1): monthly energy-data reconciliation to invoices, reviewer sign-off, internal-audit walkthrough of GHG process in Q4-2025.', 'Controls manual; internal audit report IA-2025-09', 'Financial Controller', 'In progress', 'Two IA findings open - see F4 / F5'),
        ('S2.7', 'Avoid unnecessary duplication between S1 and S2 governance disclosures', 'Single integrated governance section drafted covering climate and other SROs.', 'Draft annual report s.3.1', 'Head of Sustainability', 'Complete', ''),
    ]
    r = req_table(X, ws, r, rows, '1. Governance disclosure requirements')
    X.sec(ws, r, '2. Supporting data - governance meeting log FY2025 (evidence for S2.6(a)(iii))', 7)
    M = [(D(2025, 2, 18), 'ARSC', 'FY2024 ISSB disclosures (first-time application); transition reliefs', 'Approved'),
         (D(2025, 3, 20), 'Board', 'Sustainability Policy; ARSC ToR revision', 'Approved'),
         (D(2025, 5, 22), 'Board', 'Capex: rooftop solar 1.2 MWp and Factory B flood wall; internal carbon price', 'Approved'),
         (D(2025, 6, 12), 'Board / ARSC', 'ISSB and climate-scenario training (external facilitator)', 'Training'),
         (D(2025, 8, 14), 'ARSC', 'H1 sustainability dashboard; Scope 3 screening results', 'Noted'),
         (D(2025, 10, 30), 'ARSC (special)', 'October flood at Factory B - loss assessment, insurance, adaptation plan', 'Actions set'),
         (D(2025, 11, 20), 'ARSC', 'Q3 dashboard; materiality assessment (A4); internal audit of GHG data', 'Approved'),
         (D(2025, 12, 11), 'Board', 'FY2026 budget incl. boiler-conversion pilot; target progress review', 'Approved')]
    rows = [(m[0], m[1], m[2], m[3], 'Company Secretary', 'Complete', 'Minutes on file') for m in M]
    a, b = X.table(ws, r + 1, ['Date', 'Body', 'Sustainability / climate agenda item', 'Outcome', 'Owner', 'Minutes obtained', 'Comment'], rows, inp=(1, 2, 3, 4, 5, 6))
    X.cf_status(ws, a, b, 5)
    X.put(ws, b + 1, 1, 'Meetings with sustainability on agenda', bold=True)
    X.put(ws, b + 1, 2, f'=COUNTA({A(a, 0)}:{A(b, 0)})', bold=True, align='left')
    ws.freeze_panes(9, 2)


def build_c1(X):
    S = 'C1 Strategy'
    ws, r = X.sheet(S, 'C1', 'Strategy - effects on business model, value chain, strategy and decision-making; transition plan',
                    'For each relevant SRO document the current and anticipated effects on the business model and value chain, where they are concentrated, the entity\'s response (mitigation, adaptation), resourcing, progress against previously disclosed plans and trade-offs. Section 2 documents the climate-related transition plan.',
                    'IFRS S1.28-33; IFRS S2.8-14; IFRS S1.33(c) (trade-offs)',
                    'HB 3.3.10-3.3.30, 4.1.60-4.1.80, 5.1 (transition plans); DC 2.2.1-2.2.3, 3.2.1-3.2.3',
                    [8, 26, 40, 30, 44, 30, 34, 13], prep='Ly S.', tab=C['teal'], status='In review')
    X.sec(ws, r, '1. Effects and responses by relevant SRO (S1.32-33; S2.13-14)', 8)
    H = ['SRO ID', 'Risk / opportunity', 'Current and anticipated effects on business model and value chain (S2.13(a))', 'Where concentrated (S2.13(b))',
         'Response - direct / indirect mitigation and adaptation (S2.14(a)(i)-(iii))', 'Resourcing (S2.14(b))', 'Progress vs prior-period plans (S2.14(c)) and trade-offs (S1.33(c))', 'Status']
    R = [
        ('CR-01', 'Flooding - Factory B', 'Current: 9 production days lost Oct-2025; orders re-routed to Factory A with overtime. Anticipated: more frequent events; possible higher insurance deductible', 'Factory B (41% of sewing capacity; 39% of PPE)', 'Direct adaptation: 1.8 m perimeter flood wall and pump station (phase 1 done), raise electrical rooms and fabric store, business-continuity plan with Factory A swing capacity', 'Capex USD 0.54m FY2025; USD 0.65m FY2026 from operating cash flow', 'Prior plan (FY2024): feasibility study - DONE; wall phase 1 completed Sep-2025 (60% of critical assets protected). Trade-off: capex diverted from line automation', 'Complete'),
        ('CR-02', 'Heat stress', 'Current: efficiency -1.1% in Apr-May; 14 heat-related clinic visits. Anticipated: more hot days above 35C', 'Sewing floors, both factories', 'Evaporative cooling and roof insulation; shift start moved to 06:30 in hot season; hydration and rest protocol', 'Opex USD 95k; capex in FY2026 budget USD 0.3m', 'New in FY2025 (no prior plan)', 'Complete'),
        ('CR-03', 'Buyer decarbonisation requirements', 'Current: 62% of revenue from brands with supplier climate requirements; two brands require -30% Scope 1+2 by 2030 and coal / forest-wood phase-out. Anticipated: order allocation tied to verified performance', 'Top-3 customers (EU / US)', 'Transition plan (section 2); Higg FEM verification; annual buyer climate disclosure; engagement with nominated mills (indirect mitigation)', 'Sustainability team 4 FTE; verification USD 60k p.a.', 'On plan: Scope 1+2 down vs base year (see E9). Trade-off: electrification raises grid exposure', 'Complete'),
        ('CR-04', 'Biomass boiler fuel', 'Current: 7,200 t fuelwood p.a.; 35% from certified plantation / agri-residue. Anticipated: restrictions and buyer bans on forest wood by 2028', 'Boiler houses at A, B and JV', 'Switch to rice-husk briquettes (2026), electric heat-pump boilers pilot at Factory A (2026), full phase-out of uncertified wood by 2028', 'Capex USD 0.62m FY2025 pilot; USD 2.1m 2026-2028; green loan under negotiation', 'Pilot ordered Dec-2025 (3 months behind FY2024 plan). Trade-off: briquette supply seasonality', 'In progress'),
        ('CR-05', 'Electricity cost and reliability', 'Current: 46 outage hours; diesel back-up cost USD 0.26m. Anticipated: tariff +3% to +12% by 2030 (C3)', 'All sites', 'Rooftop solar; power-factor correction; energy-management system (ISO 50001 alignment)', 'Included in CO-01', 'Outage hours down 18% vs FY2024', 'Complete'),
        ('CR-06', 'Export-market regulation / carbon pricing', 'Current: none. Anticipated: buyers pass through EPR / due-diligence data requests; possible carbon cost in supply contracts', 'EU-destined production (46% of revenue)', 'Regulatory horizon scanning with GMAC / TAFTAC; product traceability system; internal carbon price in capex appraisal', 'Within existing teams', 'New in FY2025', 'In progress'),
        ('CR-07', 'Raw-material supply', 'Current: none significant. Anticipated: fabric price and lead-time volatility', 'Upstream mills (China, Vietnam)', 'Dual-sourcing with buyers\' consent; increase recycled polyester share; 6-week buffer for core fabrics', 'Working capital +USD 0.4m', 'New in FY2025', 'In progress'),
        ('CO-01', 'Solar and energy efficiency', 'Current: 1.2 MWp commissioned Apr-2025; 1,090 MWh generated. Anticipated: phase 2 at Factory B (0.9 MWp)', 'Factory A (phase 1); Factory B (phase 2)', 'Phase 2 solar; high-efficiency servo motors; compressed-air leak programme', 'Capex USD 1.45m FY2025; USD 1.0m FY2026', 'FY2024 plan: commission by Mar-2025 - achieved Apr-2025', 'Complete'),
        ('CO-02', 'Preferred low-carbon supplier status', 'Current: 13% of revenue under recycled-content / low-carbon programmes. Anticipated: long-term agreements with 2 brands', 'Top-3 customers', 'Third-party verified GHG data; product carbon-footprint capability', 'Verification + PCF tool USD 80k', 'New disclosure (S1.B34 exemption not used)', 'Complete'),
        ('SR-01', 'Labour conditions', 'Current: BFC assessment - 3 non-compliance points (overtime records). Anticipated: stricter due-diligence audits', 'Both factories', 'Digital time-recording; grievance mechanism; supervisor training', 'USD 120k', 'First year of disclosure (S1.E5 relief used in FY2024)', 'Complete'),
        ('SR-02', 'Occupational health and safety', 'Current: 31 lost-time injuries; no fatalities. Anticipated: stable', 'Both factories; commuting', 'Fire-safety remediation; safe-transport programme with truck operators', 'Capex USD 0.21m', 'First year of disclosure', 'Complete'),
        ('SR-03', 'Water and wastewater - wash plant', 'Current: 1 minor discharge exceedance (COD) in Jun-2025; no penalty. Anticipated: tighter limits; groundwater stress', 'JV wash plant', 'ETP upgrade; 30% water recycling by 2027; ZDHC MRSL conformance', 'JV capex USD 0.48m (TTM share 50%)', 'First year of disclosure', 'In progress'),
    ]
    a, b = X.table(ws, r + 1, H, R, inp=tuple(range(1, 8)))
    X.dv(ws, a, b, 7, STATUS)
    X.cf_status(ws, a, b, 7)
    r = b + 2
    X.sec(ws, r, '2. Climate-related transition plan (S2.14(a)(iv)-(v)); key assumptions and dependencies', 8)
    H = ['Lever', 'Action', 'Timeline', 'Capex (USD\'000)', 'Key assumptions', 'Dependencies', 'Expected abatement (tCO2e p.a.) by 2030', 'Status']
    T = [
        ('L1', 'Rooftop solar 2.1 MWp (A phase 1, B phase 2)', '2025-2026', 2450, 'Yield 1,400 kWh/kWp; self-consumption 95%', 'EDC rooftop-solar regulations; roof structural capacity', 1180, 'In progress'),
        ('L2', 'Energy efficiency (servo motors, LED, compressed air, heat recovery)', '2025-2027', 760, '8% electricity saving vs FY2023 baseline', 'Production volume stable', 560, 'In progress'),
        ('L3', 'Boiler conversion - electric heat pumps / briquettes; forest-wood phase-out', '2026-2028', 2720, 'Heat-pump COP 3.0; briquette supply 6,000 t p.a.', 'Grid capacity upgrade; briquette supply chain; green-loan approval', 190, 'In progress'),
        ('L4', 'Replace diesel back-up with battery storage (Factory A)', '2027-2029', 900, 'Outage hours < 40 p.a.', 'Battery cost decline; grid reliability', 310, 'Not started'),
        ('L5', 'Fleet: electric shuttle vehicles and forklifts', '2026-2030', 380, '60% of vehicle fuel displaced', 'Charging infrastructure', 185, 'Not started'),
        ('L6', 'Renewable electricity procurement (PPA / I-REC) for residual grid use', '2028-2030', 0, 'Instruments available and credible in Cambodia', 'Regulatory framework for corporate PPAs', 1400, 'Not started'),
        ('L7', 'Supplier engagement - fabric mills (Scope 3 cat. 1)', '2026-2030', 0, 'Buyers co-fund mill programmes', 'Buyer nomination of lower-carbon mills', 0, 'Not started'),
    ]
    a, b = X.table(ws, r + 1, H, T, inp=(1, 2, 4, 5, 7))
    X.dv(ws, a, b, 7, STATUS)
    X.cf_status(ws, a, b, 7)
    X.put(ws, b + 1, 2, 'Total', bold=True)
    X.put(ws, b + 1, 3, f'=SUM({A(a, 3)}:{A(b, 3)})', bold=True)
    X.put(ws, b + 1, 6, f'=SUM({A(a, 6)}:{A(b, 6)})', bold=True)
    X.name('TP_Capex', S, b + 1, 3)
    X.name('TP_Abate', S, b + 1, 6)
    X.note(ws, b + 3, 'Abatement of L7 is not quantified (Scope 3). Planned abatement is compared with the reduction required by target T1 in E9. The plan does not rely on carbon credits (S2.36(e)).', 8)
    ws.freeze_panes(10, 2)


def build_c2(X):
    S = 'C2 FinEffects'
    ws, r = X.sheet(S, 'C2', 'Current and anticipated financial effects',
                    'Quantify (single amount or range) how SROs affected financial position, performance and cash flows in the period and how they are expected to do so over the short, medium and long term; document where quantitative information is not provided and why.',
                    'IFRS S1.34-40; IFRS S2.15-21; IFRS S1.21(b)(ii), B40-B44 (connectivity)',
                    'HB 3.3.40-3.3.90; DC 2.2.4, 3.2.4',
                    [8, 40, 22, 26, 13, 13, 40, 16, 24], prep='Sok D.', tab=C['teal'], status='In review')
    X.sec(ws, r, '1. CURRENT financial effects - FY2025 (S2.16(a)-(b); S1.35(a)-(b)).  Sign: positive = cost / outflow / asset addition; negative = income / saving', 9)
    H = ['SRO ID', 'Effect', 'Financial statement', 'Line item affected', 'Amount (USD\'000)', 'Recognised in FS?', 'Basis of quantification', 'Significant risk of material adjustment next period? (S2.16(b))', 'Evidence / FS tie (see F3)']
    R = [
        ('CR-01', 'Flood clean-up and repairs', 'Profit or loss', 'Other operating expenses', 420, 'Yes', 'Invoices and internal labour charge', 'No', 'GL 7450; FS note 7'),
        ('CR-01', 'Inventory written off (water-damaged fabric and WIP)', 'Profit or loss', 'Cost of sales', 185, 'Yes', 'Stock count 14-Oct-2025', 'No', 'FS note 12'),
        ('CR-01', 'Contribution lost during 9-day stoppage (estimate)', 'Profit or loss', 'Revenue / gross profit (not separately recognised)', 610, 'No - estimate', 'Lost pieces x average contribution per piece', 'No', 'Production report; estimate memo'),
        ('CR-01', 'Insurance recovery recognised', 'Profit or loss / SOFP', 'Other income / other receivables', -380, 'Yes', 'Insurer settlement letter 12-Dec-2025', 'Yes', 'FS note 14 - receivable subject to final adjustment'),
        ('CR-01', 'Flood wall and pump station', 'SOFP / cash flows', 'PPE additions (investing)', 540, 'Yes', 'Capex register', 'No', 'FS note 9'),
        ('CR-02', 'Cooling running costs and heat-related absence', 'Profit or loss', 'Cost of sales', 95, 'Yes', 'Utilities and payroll analysis', 'No', 'GL analysis'),
        ('CR-03', 'Higg FEM verification and GHG data assurance readiness', 'Profit or loss', 'Administrative expenses', 60, 'Yes', 'Invoices', 'No', 'GL 6890'),
        ('CR-04', 'Electric heat-pump boiler pilot (deposit and works)', 'SOFP / cash flows', 'PPE - assets under construction', 620, 'Yes', 'Capex register', 'No', 'FS note 9'),
        ('CR-05', 'Diesel for back-up generation during outages', 'Profit or loss', 'Cost of sales - utilities', 260, 'Yes', 'Fuel purchase ledger', 'No', 'GL 6215'),
        ('CO-01', 'Rooftop solar phase 1', 'SOFP / cash flows', 'PPE additions (investing)', 1450, 'Yes', 'Capex register', 'No', 'FS note 9'),
        ('CO-01', 'Electricity cost avoided by solar generation', 'Profit or loss', 'Cost of sales - utilities', "=-ROUND('E1 ActivityData'!$E$22*ElecExp25/('E1 ActivityData'!$B$22+'E1 ActivityData'!$C$22),0)", 'Yes (lower expense)', 'Solar MWh x average group tariff (derived: electricity expense / MWh purchased)', 'No', 'E1; inverter logs'),
        ('CO-01', 'LED and efficient-motor programme', 'SOFP / cash flows', 'PPE additions (investing)', 310, 'Yes', 'Capex register', 'No', 'FS note 9'),
        ('SR-02', 'Fire-safety remediation', 'SOFP / cash flows', 'PPE additions (investing)', 210, 'Yes', 'Capex register', 'No', 'FS note 9'),
        ('SR-01', 'Digital time-recording and grievance system', 'Profit or loss', 'Administrative expenses', 120, 'Yes', 'Invoices', 'No', 'GL 6710'),
        ('SO-01', 'Interest saving - sustainability-linked loan margin ratchet', 'Profit or loss', 'Finance costs', -45, 'Yes', '15 bp on USD 30m average drawn', 'No', 'FS note 20'),
    ]
    a, b = X.table(ws, r + 1, H, R, inp=(1, 2, 3, 5, 6, 7, 8))
    X.ad['c2cur'] = (S, a, b)
    X.dv(ws, a, b, 7, ['Yes', 'No'])
    lab = [('Net effect on profit before tax - recognised items (cost +)', f'=SUMIFS({A(a, 4)}:{A(b, 4)},{A(a, 2)}:{A(b, 2)},"Profit or loss*",{A(a, 5)}:{A(b, 5)},"Yes*")', 'PL_Effect'),
           ('Capital expenditure recognised on SROs (ties to E7, S2.29(e))', f'=SUMIFS({A(a, 4)}:{A(b, 4)},{A(a, 3)}:{A(b, 3)},"PPE*")', 'Capex_SRO'),
           ('Net PBT effect as % of profit before tax', f'={A(b + 1, 4)}/PBT25', None)]
    for i, (l, f, nm) in enumerate(lab):
        ws.merge_range(b + 1 + i, 0, b + 1 + i, 3, l, X.f(bold=True, border=1, border_color=C['mid']))
        X.put(ws, b + 1 + i, 4, f, bold=True, num_format='0.0%' if i == 2 else '#,##0')
        if nm:
            X.name(nm, S, b + 1 + i, 4)
    r = b + 5
    X.sec(ws, r, '2. ANTICIPATED financial effects (S2.16(c)-(d), 17-21; S1.35(c)-(d), 37-40) - annual effect on profit before tax (negative = reduces profit); ranges from scenario analysis in C3', 9)
    H = ['SRO ID', 'Anticipated effect', 'Time horizon', 'FS line items likely affected (S2.21(b))', 'Range - lower effect (USD\'000 p.a.)', 'Range - higher effect (USD\'000 p.a.)', 'Basis / link', 'Quantitative info provided?', 'If not: reason (S2.19-21)']
    c3 = "'C3 Resilience'!"
    R = [
        ('CR-01', 'Expected annual flood loss (after phase-1 defences)', 'Short-Medium', 'Other operating expenses; cost of sales; PPE (impairment)', f"=-{c3}$C$35", f"=-{c3}$E$35", 'C3 scenario table - 2030, orderly vs hot-house', 'Yes', ''),
        ('CR-02', 'Heat-related productivity loss', 'Medium', 'Cost of sales (direct labour)', f"=-{c3}$C$36", f"=-{c3}$E$36", 'C3 scenario table - 2030', 'Yes', ''),
        ('CR-03', 'Contribution at risk if buyer climate requirements are not met', 'Short-Medium', 'Revenue; gross profit', f"=-{c3}$E$38", f"=-{c3}$C$38", 'C3 scenario table - 2030', 'Yes', ''),
        ('CR-05', 'Electricity tariff increase', 'Medium', 'Cost of sales - utilities', f"=-{c3}$E$37", f"=-{c3}$C$37", 'C3 scenario table - 2030', 'Yes', ''),
        ('CR-06', 'Carbon-cost pass-through from buyers / regulation', 'Long', 'Revenue (price pressure); cost of sales', f"=-{c3}$E$34", f"=-{c3}$C$34", 'C3 scenario table - 2030', 'Yes', 'Range only - high measurement uncertainty'),
        ('CR-04', 'Boiler conversion: depreciation and fuel-cost change', 'Medium', 'PPE; depreciation; cost of sales', None, None, 'Engineering study due Q2-2026', 'No', 'S2.19(b): uncertainty too high pending pilot results - qualitative information given; line items identified'),
        ('CR-07', 'Fabric price volatility', 'Medium', 'Cost of sales; inventories', None, None, 'Not separately identifiable from general commodity price movements', 'No', 'S2.19(a): effects not separately identifiable'),
        ('CO-01', 'Energy cost savings (solar phases 1-2 and efficiency)', 'Short-Medium', 'Cost of sales - utilities', f"=-{c3}$E$39", f"=-{c3}$C$39", 'C3 scenario table - 2030 (saving = positive effect on profit)', 'Yes', ''),
        ('SR-01', 'Order suspension following adverse labour audit', 'Short', 'Revenue', None, None, 'Low-probability / high-impact; no reliable estimate', 'No', 'S1.38(b): measurement uncertainty; qualitative disclosure with customer-concentration data'),
    ]
    a, b = X.table(ws, r + 1, H, R, inp=(1, 2, 3, 6, 7, 8))
    X.dv(ws, a, b, 7, ['Yes', 'No'])
    X.cf_status(ws, a, b, 7)
    r = b + 2
    X.sec(ws, r, '3. Investment plans and funding (S2.16(c)(i)-(ii)) - USD\'000', 9)
    H = ['Ref', 'Planned investment', 'Committed?', 'Funding source', 'FY2026', 'FY2027-2030', 'Link', '', '']
    R = [('I1', 'Solar phase 2 - Factory B', 'Yes', 'Operating cash flow', 1000, 0, 'C1 L1', '', ''),
         ('I2', 'Boiler conversion', 'Partly', 'Green loan (under negotiation) 70% / cash 30%', 700, 1400, 'C1 L3', '', ''),
         ('I3', 'Flood defences phase 2 and asset raising', 'Yes', 'Operating cash flow', 650, 0, 'CR-01', '', ''),
         ('I4', 'Efficiency, storage and fleet', 'No', 'Operating cash flow / leasing', 450, 1280, 'C1 L2, L4, L5', '', ''),
         ('I5', 'Cooling and OHS', 'Yes', 'Operating cash flow', 300, 400, 'CR-02, SR-02', '', '')]
    a, b = X.table(ws, r + 1, H, R, inp=(1, 2, 3, 6))
    X.put(ws, b + 1, 3, 'Total planned', bold=True)
    X.put(ws, b + 1, 4, f'=SUM({A(a, 4)}:{A(b, 4)})', bold=True)
    X.put(ws, b + 1, 5, f'=SUM({A(a, 5)}:{A(b, 5)})', bold=True)
    X.put(ws, b + 2, 3, 'Available liquidity (A1)', bold=True)
    X.put(ws, b + 2, 4, '=Liquidity25', bold=True)
    X.put(ws, b + 3, 3, 'FY2026 plan covered by liquidity?', bold=True)
    X.put(ws, b + 3, 4, f'=IF({A(b + 2, 4)}>={A(b + 1, 4)},"Yes","No")', bold=True)
    ws.freeze_panes(10, 2)


def build_c3(X):
    S = 'C3 Resilience'
    ws, r = X.sheet(S, 'C3', 'Climate resilience - scenario analysis',
                    'Document the climate-related scenario analysis (inputs, assumptions, results) and the assessment of the resilience of strategy and business model. Approach is commensurate with circumstances: quantitative for the main drivers, qualitative narrative for the rest.',
                    'IFRS S2.22-23, B1-B18; IFRS S1.41-42',
                    'HB 3.3.100, 4.1.90-4.1.110, 5.2.10-5.2.70; DC 2.2.5, 3.2.5',
                    [44, 16, 14, 14, 14, 14, 46], prep='Sok D.', rev='Keo V. (CFO) / external adviser', tab=C['teal'], status='In review')
    X.sec(ws, r, '1. Scenario inputs (S2.22(b)(i))', 7)
    rows = [
        ('Scenarios used and sources', 'A) "Orderly transition" ~1.5C: NGFS Net Zero 2050 (transition) with IPCC SSP1-2.6 (physical).   B) "Hot-house world" ~3C+: NGFS Current Policies (transition) with IPCC SSP3-7.0 (physical)'),
        ('Diverse range? Paris-aligned scenario included?', 'Yes - two contrasting scenarios; Scenario A is aligned with the latest international agreement on climate change (S2.22(b)(i)(4))'),
        ('Associated with transition or physical risks', 'A: mainly transition risks (CR-03, 04, 05, 06).  B: mainly physical risks (CR-01, 02, 07)'),
        ('Why relevant', 'Customers\' targets follow 1.5C pathways (A); Cambodia is highly exposed to flood and heat under high-warming pathways (B)'),
        ('Time horizons', '2030 (medium term) and 2040 (long term)'),
        ('Scope of operations', 'Factories A and B and the JV wash plant (100% of own operations); upstream fabric supply assessed qualitatively'),
        ('Reporting period in which analysis was carried out (S2.22(b)(iii), B18)', 'FY2025 (Q3-Q4 2025). To be updated at least in line with the strategic-planning cycle (every 3 years) - resilience assessment itself is updated annually'),
        ('Approach selected (S2.B1-B15)', 'Exposure: high; skills and resources: moderate -> quantitative modelling of key drivers with external support; reasonable and supportable information available without undue cost or effort'),
    ]
    for i, (l, v) in enumerate(rows):
        X.put(ws, r + 1 + i, 0, l, bold=True)
        ws.merge_range(r + 1 + i, 1, r + 1 + i, 6, v, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.set_row(r + 1 + i, 28)
    r += len(rows) + 2            # r = 18
    X.sec(ws, r, '2. Key assumptions (S2.22(b)(ii)) - yellow cells are scenario levers', 7)
    H = ['Assumption', 'Unit', 'A: Orderly 2030', 'A: Orderly 2040', 'B: Hot-house 2030', 'B: Hot-house 2040', 'Source / basis (illustrative - replace with entity-specific study)']
    Y = {'bg_color': C['yellow']}
    asm = [
        ('Carbon cost passed through by buyers / regulation (policy)', 'USD / tCO2e', 40, 95, 5, 10, 'NGFS carbon-price trajectories, scaled for pass-through to a Cambodian supplier'),
        ('Scope 1+2 emissions as % of FY2025 (after planned actions)', '%', 0.58, 0.30, 0.90, 0.85, 'Transition plan C1 (A) vs partial delivery (B)'),
        ('Annual probability of a damaging flood at Factory B (physical)', '%', 0.06, 0.07, 0.08, 0.12, 'Flood-risk study 2025; return periods shift under SSP3-7.0'),
        ('Loss per flood event after defences (damage + interruption)', 'USD\'000', 1400, 1400, 2400, 2800, 'A assumes phase-2 defences completed; B assumes phase 1 only'),
        ('Heat: loss of direct-labour productivity', '% of labour cost', 0.010, 0.015, 0.018, 0.035, 'ILO heat-stress studies for SE Asia garment sector (illustrative)'),
        ('Electricity tariff change vs FY2025 (energy mix)', '%', 0.12, 0.18, 0.03, 0.05, 'Power Development Plan; coal / LNG / hydro mix'),
        ('Share of climate-conditional revenue lost if requirements not fully met (macro / market)', '%', 0.10, 0.15, 0.02, 0.03, 'Management estimate from buyer scorecard rules'),
        ('Annual energy-cost savings from solar and efficiency (technology)', 'USD\'000', 420, 640, 310, 400, 'Engineering estimates C1 L1-L2'),
        ('Contribution margin on revenue', '%', 0.18, 0.18, 0.18, 0.18, 'FY2025 management accounts'),
        ('Share of employee cost that is direct labour', '%', 0.80, 0.80, 0.80, 0.80, 'Payroll analysis'),
    ]
    a, b = X.table(ws, r + 1, H, asm, nf={}, inp=(6,), extra={2: Y, 3: Y, 4: Y, 5: Y})
    for i, x in enumerate(asm):      # percent formats
        if x[1].startswith('%'):
            for j in range(2, 6):
                X.put(ws, a + i, j, x[j], num_format='0.0%', bg_color=C['yellow'])
    q = a  # first assumption row (0-based)
    r = b + 2
    X.sec(ws, r, '3. Results - modelled annual effect on profit before tax, USD\'000 (cost shown positive; saving negative)', 7)
    H = ['Driver (linked SRO)', 'Type', 'A: Orderly 2030', 'A: Orderly 2040', 'B: Hot-house 2030', 'B: Hot-house 2040', 'Formula logic']
    s12 = 'GHG_S12'
    trev = 'Rev_Transition'

    def cols(fn):
        return [fn(chr(ord('C') + k)) for k in range(4)]
    R = [
        ['Carbon cost on Scope 1+2 (CR-06)', 'Transition'] + cols(lambda c: f"={c}{q + 1}*{c}{q + 2}*{s12}/1000") + ['Carbon price x residual Scope 1+2 emissions (E6)'],
        ['Expected annual flood loss (CR-01)', 'Physical'] + cols(lambda c: f"={c}{q + 3}*{c}{q + 4}") + ['Probability x loss per event'],
        ['Heat-related productivity loss (CR-02)', 'Physical'] + cols(lambda c: f"={c}{q + 5}*{c}{q + 10}*StaffCost25") + ['% x direct labour cost'],
        ['Electricity tariff increase (CR-05)', 'Transition'] + cols(lambda c: f"={c}{q + 6}*ElecExp25") + ['% x FY2025 electricity expense (A1)'],
        ['Contribution at risk - buyer requirements (CR-03)', 'Transition'] + cols(lambda c: f"={c}{q + 7}*{c}{q + 9}*{trev}") + ['% lost x margin x climate-conditional revenue (E7)'],
        ['Energy savings from solar / efficiency (CO-01)', 'Opportunity'] + cols(lambda c: f"=-{c}{q + 8}") + ['Savings (negative cost)'],
    ]
    a2, b2 = X.table(ws, r + 1, H, R)
    assert a2 == 33, a2   # C2 links rely on rows 34-39 (Excel numbering)
    X.put(ws, b2 + 1, 0, 'Net modelled effect on profit before tax', bold=True)
    X.put(ws, b2 + 1, 1, '')
    for k in range(4):
        X.put(ws, b2 + 1, 2 + k, f'=SUM({A(a2, 2 + k)}:{A(b2, 2 + k)})', bold=True)
        X.put(ws, b2 + 2, 2 + k, f'={A(b2 + 1, 2 + k)}/PBT25', num_format='0.0%')
        X.put(ws, b2 + 3, 2 + k, f'=IF({A(b2 + 1, 2 + k)}<=Liquidity25*0.25,"Resilient","Stress")')
    X.put(ws, b2 + 2, 0, 'As % of FY2025 profit before tax'); X.put(ws, b2 + 2, 1, '')
    X.put(ws, b2 + 3, 0, 'Resilience test: annual effect <= 25% of available liquidity'); X.put(ws, b2 + 3, 1, '')
    X.put(ws, b2 + 1, 6, 'Disclosed as a range (S2.22: single amount or range)')
    X.put(ws, b2 + 2, 6, ''); X.put(ws, b2 + 3, 6, 'Threshold is a management judgement (F2 J-05)')
    X.name('Scen_A30', S, b2 + 1, 2)
    X.name('Scen_B30', S, b2 + 1, 4)
    r = b2 + 5
    X.sec(ws, r, '4. Resilience assessment at the reporting date (S2.22(a))', 7)
    rows = [
        ('Implications for strategy and business model (S2.22(a)(i))', 'Strategy remains viable under both scenarios. A: the transition plan must be delivered on time to protect order allocation - boiler conversion is the critical path. B: Factory B needs phase-2 flood defences and a decision by 2028 on relocating the fabric store.'),
        ('Significant areas of uncertainty (S2.22(a)(ii))', 'Pace of buyer requirements; Cambodian grid mix and tariff; flood frequency at basin level; availability of renewable-energy instruments; productivity response to heat.'),
        ('Capacity to adjust - financial resources (S2.22(a)(iii)(1))', '=\"Cash and undrawn committed facilities of USD \"&TEXT(Liquidity25,\"#,##0\")&\"k versus FY2026 climate capex plan in C2; sustainability-linked facility renewable 2027.\"'),
        ('Capacity to adjust - redeploy, repurpose, upgrade or decommission assets (S2.22(a)(iii)(2))', 'Sewing lines are mobile (swing capacity between A and B within 3 weeks); boilers can be replaced modularly; land lease at Factory B runs to 2041.'),
        ('Effect of current and planned investments (S2.22(a)(iii)(3))', 'Planned mitigation and adaptation capex (C1 transition plan; C2 investment plan) reduces modelled 2030 net effect - compare columns A vs B.'),
    ]
    for i, (l, v) in enumerate(rows):
        X.put(ws, r + 1 + i, 0, l, bold=True)
        if v.startswith('='):
            ws.merge_range(r + 1 + i, 1, r + 1 + i, 6, '', X.f(text_wrap=True, border=1, border_color=C['mid']))
            ws.write_formula(r + 1 + i, 1, v, X.f(text_wrap=True, border=1, border_color=C['mid']))
        else:
            ws.merge_range(r + 1 + i, 1, r + 1 + i, 6, v, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.set_row(r + 1 + i, 42)
    ws.freeze_panes(8, 1)


def build_d1(X):
    S = 'D1 RiskMgmt'
    ws, r = X.sheet(S, 'D1', 'Risk management - processes to identify, assess, prioritise and monitor SROs',
                    'Document the processes and policies used to identify, assess, prioritise and monitor sustainability- and climate-related risks and opportunities and how they are integrated into overall (enterprise) risk management. Section 2 holds the scoring criteria applied in A3.',
                    'IFRS S1.43-44; IFRS S2.24-26',
                    'HB 3.4; DC 2.3 and 3.3',
                    [16, 44, 56, 34, 20, 13, 30], prep='Ly S.', tab=C['teal'])
    rows = [
        ('S2.25(a)(i) / S1.44(a)(i)', 'Inputs and parameters used (data sources, scope of operations)', 'Inputs: value-chain map (A2), flood-risk study, buyer codes, BFC reports, energy audit, regulatory scan. Scope: all three facilities plus tier-1 fabric mills.', 'ERM procedure v3; A2', 'Risk Manager', 'Complete', ''),
        ('S2.25(a)(ii) / S1.44(a)(ii)', 'Whether and how scenario analysis informs risk identification', 'Climate: two scenarios (C3) used to stress likelihood ratings of CR-01, 02, 03, 06. Non-climate: not used.', 'C3', 'Risk Manager', 'Complete', ''),
        ('S2.25(a)(iii) / S1.44(a)(iii)', 'How nature, likelihood and magnitude of effects are assessed', '5x5 matrix below; magnitude calibrated to profit before tax; qualitative overlay for safety and licence-to-operate.', 'Risk criteria (section 2)', 'Risk Manager', 'Complete', ''),
        ('S2.25(a)(iv) / S1.44(a)(iv)', 'Whether and how SROs are prioritised relative to other risk types', 'Same matrix as for financial and operational risks; climate risks CR-01 and CR-03 rank in the group top-10 register.', 'Group risk register Q4-2025', 'Risk Manager', 'Complete', ''),
        ('S2.25(a)(v) / S1.44(a)(v)', 'How risks are monitored', 'Key risk indicators reported monthly to SSC (energy, outage hours, river level alerts, audit scores); quarterly to ARSC.', 'KRI dashboard', 'Head of Sustainability', 'Complete', ''),
        ('S2.25(a)(vi) / S1.44(a)(vi)', 'Changes in processes compared with the previous period', 'FY2025 changes: non-climate SROs added to register (end of climate-first relief); Scope 3 screening; flood KRI added after October event.', 'ERM change log', 'Risk Manager', 'In progress', 'Change log to be approved by ARSC (A-009)'),
        ('S2.25(b) / S1.44(b)', 'Processes to identify, assess, prioritise and monitor opportunities (incl. use of scenario analysis)', 'Opportunities captured in annual strategy workshop and capex appraisal; scenario A used to size CO-01 and CO-02.', 'Strategy workshop pack Sep-2025', 'CFO', 'Complete', ''),
        ('S2.25(c) / S1.44(c)', 'Extent and manner of integration into the overall risk management process', 'Single ERM framework (ISO 31000-aligned); sustainability risks owned by line management; second-line review by Risk Manager; internal audit provides assurance.', 'ERM policy', 'Risk Manager', 'Complete', ''),
        ('S2.26', 'Avoid duplication between S1 and S2 risk-management disclosures', 'Integrated risk-management section drafted.', 'Draft annual report s.3.3', 'Head of Sustainability', 'Complete', ''),
    ]
    r = req_table(X, ws, r, rows, '1. Risk-management disclosure requirements')
    X.sec(ws, r, '2. Assessment criteria applied in A3 (S2.25(a)(iii))', 7)
    rows = [(5, 'Almost certain (> 80% in horizon)', '=\"> \"&TEXT(0.2*PBT25,\"#,##0\")', 'Fatality / permanent loss of key customer or licence', 'Critical if score >= 15'),
            (4, 'Likely (50-80%)', '=TEXT(0.1*PBT25,\"#,##0\")&\" - \"&TEXT(0.2*PBT25,\"#,##0\")', 'Order suspension by a top-3 customer; regulatory sanction', 'High if score 10-14'),
            (3, 'Possible (20-50%)', '=TEXT(0.05*PBT25,\"#,##0\")&\" - \"&TEXT(0.1*PBT25,\"#,##0\")', 'Adverse audit finding; local media attention', 'Medium if score 5-9'),
            (2, 'Unlikely (5-20%)', '=TEXT(0.01*PBT25,\"#,##0\")&\" - \"&TEXT(0.05*PBT25,\"#,##0\")', 'Minor non-conformance', 'Low if score < 5'),
            (1, 'Rare (< 5%)', '=\"< \"&TEXT(0.01*PBT25,\"#,##0\")', 'Negligible', '')]
    X.table(ws, r + 1, ['Score', 'Likelihood', 'Magnitude - annual effect on profit before tax (USD\'000)', 'Magnitude - qualitative', 'Rating bands (score = L x M)'], rows, inp=(1, 3, 4))
    ws.freeze_panes(9, 2)


def build_part1(X):
    build_a1(X); build_a2(X); build_a3(X); build_a4(X)
    build_b1(X); build_c1(X); build_c2(X); build_c3(X); build_d1(X)
