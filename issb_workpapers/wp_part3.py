"""Part 3: F-series (completion), Index, Calc (dashboard data) and Dashboard."""
import datetime as dt
from wp_lib import C, A, ref, rng, CLIENT, PERIOD
from wp_part2 import E1, E5, E9, MONTHS
from wp_part4 import build_g1

D = dt.date
F1, F2, F3, F4, F5 = 'F1 Checklist', 'F2 Judgements', 'F3 Connectivity', 'F4 Evidence', 'F5 Actions'
PILLARS = ['General requirements', 'Governance', 'Strategy', 'Risk management', 'Metrics and targets']
G, GV, ST, RM, MT = PILLARS
OWN = {G: 'Financial Controller', GV: 'Company Secretary', ST: 'CFO', RM: 'Risk Manager', MT: 'Head of Sustainability'}
DUE = {G: D(2026, 3, 13), GV: D(2026, 2, 27), ST: D(2026, 3, 6), RM: D(2026, 2, 27), MT: D(2026, 3, 6)}
Cc, Ip, Ns = 'Complete', 'In progress', 'Not started'

CHK = [
    # (DC section, pillar, std, ref, requirement, applicable, status, WP, handbook)
    ('1.1', G, 'S1', 'S1.B32-B33', 'Material information disclosed even if local law permits non-disclosure; where law prohibits disclosure, identify type of information omitted and source of restriction', 'Yes', Cc, 'A1', '6.1.50'),
    ('1.2', G, 'S1', 'S1.B34-B36', 'Commercially sensitive opportunity information: exemption criteria assessed; if used, disclose the fact for each item omitted and reassess each period', 'Yes', Cc, 'A1', '2.3.190'),
    ('1.3', G, 'S1', 'S1.20, B38', 'Disclosures are for the same reporting entity as the related financial statements', 'Yes', Cc, 'A1', '2.1.10'),
    ('1.4', G, 'S1', 'S1.22', 'Identify the financial statements to which the disclosures relate', 'Yes', Cc, 'A1', '6.1.80'),
    ('1.4', G, 'S1', 'S1.24', 'Use the presentation currency of the related financial statements where currency is the unit of measure', 'Yes', Cc, 'A1', '6.1.80'),
    ('1.4', G, 'S1', 'S1.21, B39-B41', 'Provide information enabling users to understand connections between SROs, between core-content disclosures and with the financial statements', 'Yes', Ip, 'F3', '6.1.90-6.1.110'),
    ('1.4', G, 'S1', 'S1.23, B42(c)', 'Data and assumptions consistent with the financial statements to the extent possible; significant differences disclosed', 'Yes', Ip, 'F3', '6.1.120'),
    ('-', G, 'S1', 'S1.17-18, B13-B28', 'Material information about SROs identified and disclosed; materiality judgements documented and reassessed', 'Yes', Cc, 'A4', '2.2.70-2.2.120'),
    ('-', G, 'S1', 'S1.B29-B30', 'Information aggregated / disaggregated so that material information is not obscured', 'Yes', Ip, 'A4', '2.2.140-2.2.150'),
    ('1.5', G, 'S1', 'S1.59(a)', 'Identify the specific standards, pronouncements, industry practice and other sources of guidance applied, incl. SASB disclosure topics', 'Yes', Cc, 'A1', '2.2.30'),
    ('1.5', G, 'S1', 'S1.59(b)', 'Identify the industry(s) specified in ISSB Standards, SASB Standards or other sources applied, incl. for metrics', 'Yes', Cc, 'A1', '2.2.30'),
    ('1.6', G, 'S1', 'S1.60', 'Disclosures provided as part of the general purpose financial reports', 'Yes', Cc, 'A1', '6.1.30'),
    ('1.6', G, 'S1', 'S1.B45-B47', 'Information included by cross-reference: report identified, access explained, precise cross-reference, same terms and time', 'Yes', Cc, 'A1', '6.1.130'),
    ('1.7', G, 'S1', 'S1.64', 'Reported at the same time as, and for the same period as, the related financial statements', 'Yes', Ip, 'A1', '2.3.50'),
    ('1.7', G, 'S1', 'S1.66', 'Change of reporting date: period covered, reason and non-comparability disclosed', 'No', 'N/A', 'A1', '2.3.50'),
    ('1.7', G, 'S1', 'S1.67', 'Disclosures updated for information received after the reporting date about conditions existing at that date', 'Yes', Ip, 'F5', '2.3.70'),
    ('1.7', G, 'S1', 'S1.68', 'Material transactions, events and conditions after the reporting date and before authorisation disclosed', 'Yes', Ns, 'F5', '2.3.70'),
    ('1.8', G, 'S1', 'S1.70, B49', 'Comparative information for all amounts disclosed; narrative comparatives where useful', 'Yes', Cc, 'E6', '2.3.110'),
    ('1.8', G, 'S1', 'S1.B50-B51', 'Revised comparative for prior-period estimates when new information arises; difference and reasons disclosed', 'Yes', Ip, 'F2', '2.3.130'),
    ('1.8', G, 'S1', 'S1.B52-B53', 'Redefined, replaced or new metrics: revised comparative, explanation and reasons', 'Yes', Cc, 'E6', '2.3.140'),
    ('1.8', G, 'S1', 'S1.B54', 'Fact disclosed if it is impracticable to revise a comparative amount', 'No', 'N/A', 'F2', '2.3.120'),
    ('1.9', G, 'S1', 'S1.72', 'Explicit and unreserved statement of compliance only if all requirements are met', 'Yes', Ip, 'A1', '6.1.20'),
    ('1.10', G, 'S1', 'S1.74-75', 'Judgements (other than estimates) with the most significant effect on the information disclosed', 'Yes', Ip, 'F2', '2.3.80'),
    ('1.11', G, 'S1', 'S1.77', 'Most significant uncertainties affecting amounts reported', 'Yes', Ip, 'F2', '2.3.100'),
    ('1.11', G, 'S1', 'S1.78', 'Amounts subject to high measurement uncertainty: sources of uncertainty; assumptions, approximations and judgements', 'Yes', Ip, 'F2', '2.3.100'),
    ('1.12', G, 'S1', 'S1.83, B55-B59', 'Material prior-period errors corrected by restating comparatives; nature and correction disclosed', 'Yes', Cc, 'F2', '2.3.150'),
    ('1.13', G, 'S1', 'S1.E5-E6; S2.C4-C5', 'Use of transition reliefs disclosed (climate-first; Scope 3; comparatives)', 'Yes', Cc, 'A1', '6.2.50'),
    # Governance
    ('2.1', GV, 'S2', 'S2.6(a)', 'Identify governance body(s) or individual(s) responsible for oversight of climate-related risks and opportunities', 'Yes', Cc, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(a)(i)', 'How responsibilities are reflected in terms of reference, mandates, role descriptions and policies', 'Yes', Cc, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(a)(ii)', 'How the body determines that appropriate skills and competencies are available or will be developed', 'Yes', Cc, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(a)(iii)', 'How and how often the body is informed', 'Yes', Cc, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(a)(iv)', 'How the body takes SROs into account in overseeing strategy, major transactions, risk management; trade-offs', 'Yes', Cc, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(a)(v)', 'How the body oversees target-setting and monitors progress, incl. links to remuneration policies', 'Yes', Ip, 'B1', '3.2.10'),
    ('2.1', GV, 'S2', 'S2.6(b)(i)', 'Management\'s role: delegation to a position or committee and how oversight is exercised', 'Yes', Cc, 'B1', '3.2.20'),
    ('2.1', GV, 'S2', 'S2.6(b)(ii)', 'Management\'s controls and procedures and their integration with other internal functions', 'Yes', Ip, 'B1', '3.2.20'),
    ('3.1', GV, 'S1', 'S1.27(a)', 'Governance body oversight of other (non-climate) SROs - items (i) to (v), integrated with climate disclosure', 'Yes', Cc, 'B1', '3.2.10'),
    ('3.1', GV, 'S1', 'S1.27(b)', 'Management\'s role for other (non-climate) SROs - items (i) to (ii)', 'Yes', Cc, 'B1', '3.2.20'),
    # Strategy - climate
    ('2.2.1', ST, 'S2', 'S2.10(a)', 'Describe climate-related risks and opportunities that could reasonably be expected to affect prospects', 'Yes', Cc, 'A3', '4.1.30'),
    ('2.2.1', ST, 'S2', 'S2.10(b)', 'Classify each climate-related risk as physical or transition', 'Yes', Cc, 'A3', '4.1.30'),
    ('2.2.1', ST, 'S2', 'S2.10(c)', 'Time horizons over which effects of each risk and opportunity could reasonably be expected to occur', 'Yes', Cc, 'A3', '3.3.10'),
    ('2.2.1', ST, 'S2', 'S2.10(d)', 'Definitions of short, medium and long term and link to planning horizons', 'Yes', Cc, 'A3', '3.3.10'),
    ('2.2.1', ST, 'S2', 'S2.11-12', 'Industry-based disclosure topics (IFRS S2 IBG) referred to and applicability considered', 'Yes', Cc, 'A3 / E8', '4.1.30'),
    ('2.2.2', ST, 'S2', 'S2.13(a)', 'Current and anticipated effects on business model and value chain', 'Yes', Cc, 'C1', '3.3.20'),
    ('2.2.2', ST, 'S2', 'S2.13(b)', 'Where in business model and value chain risks and opportunities are concentrated', 'Yes', Cc, 'C1', '3.3.20'),
    ('2.2.3', ST, 'S2', 'S2.14(a)(i)', 'Current and anticipated changes to business model, incl. resource allocation', 'Yes', Cc, 'C1', '4.1.60-4.1.80'),
    ('2.2.3', ST, 'S2', 'S2.14(a)(ii)', 'Current and anticipated direct mitigation and adaptation efforts', 'Yes', Cc, 'C1', '4.1.80'),
    ('2.2.3', ST, 'S2', 'S2.14(a)(iii)', 'Current and anticipated indirect mitigation and adaptation efforts', 'Yes', Ip, 'C1', '4.1.80'),
    ('2.2.3', ST, 'S2', 'S2.14(a)(iv)', 'Climate-related transition plan, key assumptions and dependencies', 'Yes', Ip, 'C1', '5.1.10-5.1.40'),
    ('2.2.3', ST, 'S2', 'S2.14(a)(v)', 'How the entity plans to achieve climate-related targets incl. GHG targets', 'Yes', Ip, 'C1 / E9', '4.1.80'),
    ('2.2.3', ST, 'S1', 'S1.33(c)', 'Trade-offs between SROs considered', 'Yes', Cc, 'C1', '3.3.30'),
    ('2.2.3', ST, 'S2', 'S2.14(b)', 'How activities are resourced and planned to be resourced', 'Yes', Cc, 'C1', '4.1.80'),
    ('2.2.3', ST, 'S2', 'S2.14(c)', 'Quantitative and qualitative progress of plans disclosed in previous periods', 'Yes', Cc, 'C1', '4.1.80'),
    ('2.2.4', ST, 'S2', 'S2.15(a), 16(a)', 'Current financial effects on financial position, performance and cash flows', 'Yes', Cc, 'C2', '3.3.50'),
    ('2.2.4', ST, 'S2', 'S2.16(b)', 'SROs with a significant risk of material adjustment to carrying amounts within the next annual period', 'Yes', Ip, 'C2', '3.3.50'),
    ('2.2.4', ST, 'S2', 'S2.16(c)', 'Expected changes in financial position: investment and disposal plans; planned sources of funding', 'Yes', Cc, 'C2', '3.3.60'),
    ('2.2.4', ST, 'S2', 'S2.16(d)', 'Expected changes in financial performance and cash flows over short, medium and long term', 'Yes', Ip, 'C2', '3.3.60'),
    ('2.2.4', ST, 'S2', 'S2.17', 'Quantitative information given as a single amount or a range', 'Yes', Cc, 'C2', '3.3.60'),
    ('2.2.4', ST, 'S2', 'S2.19-21', 'Where quantitative information not provided: explain why, give qualitative information incl. FS line items, and combined effects if useful', 'Yes', Ip, 'C2', '3.3.70-3.3.90'),
    ('2.2.5', ST, 'S2', 'S2.22(a)(i)', 'Resilience assessment: implications for strategy and business model', 'Yes', Ip, 'C3', '4.1.90-4.1.110'),
    ('2.2.5', ST, 'S2', 'S2.22(a)(ii)', 'Significant areas of uncertainty considered in the resilience assessment', 'Yes', Cc, 'C3', '4.1.110'),
    ('2.2.5', ST, 'S2', 'S2.22(a)(iii)', 'Capacity to adjust or adapt: financial resources, assets, effect of investments', 'Yes', Cc, 'C3', '4.1.110'),
    ('2.2.5', ST, 'S2', 'S2.22(b)(i)', 'Scenario-analysis inputs: scenarios and sources, diversity, risk type, Paris-aligned scenario, relevance, horizons, scope', 'Yes', Cc, 'C3', '5.2.30-5.2.70'),
    ('2.2.5', ST, 'S2', 'S2.22(b)(ii)', 'Key assumptions: policies, macroeconomic trends, national / regional variables, energy usage and mix, technology', 'Yes', Ip, 'C3', '5.2.60'),
    ('2.2.5', ST, 'S2', 'S2.22(b)(iii)', 'Reporting period in which the scenario analysis was carried out', 'Yes', Cc, 'C3', '5.2.30'),
    # Strategy - other topics
    ('3.2.1', ST, 'S1', 'S1.30(a)-(c)', 'Other SROs: description, time horizons and definitions', 'Yes', Cc, 'A3', '3.3.10'),
    ('3.2.2', ST, 'S1', 'S1.32(a)-(b)', 'Other SROs: effects on business model and value chain and where concentrated', 'Yes', Cc, 'C1', '3.3.20'),
    ('3.2.3', ST, 'S1', 'S1.33(a)-(b)', 'Other SROs: responses in strategy and decision-making; progress against prior plans', 'Yes', Cc, 'C1', '3.3.30'),
    ('3.2.4', ST, 'S1', 'S1.34-35', 'Other SROs: current and anticipated financial effects', 'Yes', Ip, 'C2', '3.3.40-3.3.60'),
    ('3.2.4', ST, 'S1', 'S1.38-40', 'Other SROs: explanation where quantitative information is not provided', 'Yes', Ip, 'C2', '3.3.70-3.3.90'),
    ('3.2.5', ST, 'S1', 'S1.41', 'Other SROs: qualitative (and if applicable quantitative) assessment of resilience', 'Yes', Ns, 'C3', '3.3.100'),
    # Risk management
    ('2.3', RM, 'S2', 'S2.25(a)(i)', 'Inputs and parameters used to identify, assess, prioritise and monitor climate-related risks', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(a)(ii)', 'Whether and how scenario analysis informs identification of climate-related risks', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(a)(iii)', 'How nature, likelihood and magnitude of effects are assessed', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(a)(iv)', 'Whether and how climate-related risks are prioritised relative to other risks', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(a)(v)', 'How climate-related risks are monitored', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(a)(vi)', 'Changes in processes compared with the previous reporting period', 'Yes', Ip, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(b)', 'Processes for climate-related opportunities, incl. use of scenario analysis', 'Yes', Cc, 'D1', '3.4'),
    ('2.3', RM, 'S2', 'S2.25(c)', 'Integration into the overall risk-management process', 'Yes', Cc, 'D1', '3.4'),
    ('3.3', RM, 'S1', 'S1.44(a)', 'Other SROs: risk processes and related policies - items (i) to (vi)', 'Yes', Cc, 'D1', '3.4'),
    ('3.3', RM, 'S1', 'S1.44(b)-(c)', 'Other SROs: opportunity processes and integration into overall risk management', 'Yes', Cc, 'D1', '3.4'),
    # Metrics and targets
    ('2.4.1', MT, 'S2', 'S2.29(a)(i)(1)', 'Absolute gross Scope 1 GHG emissions (tCO2e)', 'Yes', Cc, 'E3 / E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(i)(2)', 'Absolute gross Scope 2 GHG emissions (tCO2e)', 'Yes', Cc, 'E4 / E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(i)(3)', 'Absolute gross Scope 3 GHG emissions (tCO2e)', 'Yes', Ip, 'E5 / E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.B19', 'Value-chain data for different reporting periods: conditions met and effects of significant events disclosed', 'No', 'N/A', 'E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(ii), B23-B25', 'Measured in accordance with the GHG Protocol Corporate Standard (2004) unless another method is required by a jurisdiction / exchange', 'Yes', Cc, 'E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(iii), B26-B27', 'Measurement approach, inputs and assumptions; reasons; changes during the period', 'Yes', Cc, 'E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.B29', 'Emission factors used - information enabling users to understand them', 'Yes', Ip, 'E2', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(iv)', 'Scope 1 and 2 disaggregated between consolidated accounting group and other investees', 'Yes', Cc, 'E3 / E4 / E6', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(v), B30-B31', 'Location-based Scope 2 and information about contractual instruments', 'Yes', Cc, 'E4', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.29(a)(vi)(1), B32', 'Scope 3 categories included, per the GHG Protocol Scope 3 Standard (2011)', 'Yes', Cc, 'E5', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.B40, B55-B56', 'Scope 3 measurement approach: characteristics of data; extent of primary and verified inputs', 'Yes', Ip, 'E5', '4.1.140'),
    ('2.4.1', MT, 'S2', 'S2.B57', 'If impracticable to estimate Scope 3: how Scope 3 emissions are managed', 'No', 'N/A', 'E5', '4.1.140'),
    ('2.4.4', MT, 'S2', 'S2.29(a)(vi)(2), B58-B63', 'Financed emissions (asset management, commercial banking, insurance)', 'No', 'N/A', 'E5', '4.1.200'),
    ('2.4.2', MT, 'S2', 'S2.29(b)', 'Amount and percentage of assets or business activities vulnerable to transition risks', 'Yes', Cc, 'E7', '4.1.150'),
    ('2.4.2', MT, 'S2', 'S2.29(c)', 'Amount and percentage of assets or business activities vulnerable to physical risks', 'Yes', Cc, 'E7', '4.1.150'),
    ('2.4.2', MT, 'S2', 'S2.29(d)', 'Amount and percentage of assets or business activities aligned with climate-related opportunities', 'Yes', Ip, 'E7', '4.1.150'),
    ('2.4.2', MT, 'S2', 'S2.29(e)', 'Capital expenditure, financing or investment deployed towards climate-related risks and opportunities', 'Yes', Cc, 'E7', '4.1.160'),
    ('2.4.2', MT, 'S2', 'S2.29(f)', 'Internal carbon price: whether and how applied; price per tonne', 'Yes', Cc, 'E7', '4.1.170'),
    ('2.4.2', MT, 'S2', 'S2.29(g)', 'Remuneration: whether and how climate considerations are factored in; percentage linked', 'Yes', Ip, 'E7', '4.1.180'),
    ('2.4.3', MT, 'S2', 'S2.32', 'Industry-based metrics (IFRS S2 IBG) referred to and applicability considered', 'Yes', Ip, 'E8', '4.1.190'),
    ('2.4.5', MT, 'S1', 'S1.46(b)', 'Other metrics used to measure and monitor each risk / opportunity and performance', 'Yes', Cc, 'E8', '3.5.10'),
    ('2.4.5', MT, 'S1', 'S1.49', 'Metrics taken from sources other than ISSB Standards: source and metric identified', 'Yes', Cc, 'E8', '3.5.10'),
    ('2.4.5', MT, 'S1', 'S1.50; S2.B67', 'Entity-developed metrics: definition, type, third-party validation, method and inputs', 'Yes', Ip, 'E8', '3.5.10'),
    ('2.4.5', MT, 'S1', 'S1.53', 'Metrics and targets labelled and defined with meaningful, clear and precise names', 'Yes', Cc, 'E8 / E9', '3.5.10'),
    ('2.4.6', MT, 'S2', 'S2.33(a)-(d)', 'Each climate target: metric, objective, part of entity, period', 'Yes', Cc, 'E9', '4.1.210-4.1.230'),
    ('2.4.6', MT, 'S2', 'S2.33(e)-(h)', 'Each climate target: base period, milestones, absolute / intensity, how informed by latest international agreement', 'Yes', Cc, 'E9', '4.1.230'),
    ('2.4.6', MT, 'S2', 'S2.34', 'Approach to setting and reviewing targets: validation, review process, monitoring metrics, revisions', 'Yes', Ip, 'E9', '4.1.230'),
    ('2.4.6', MT, 'S2', 'S2.35', 'Performance against each target and analysis of trends or changes', 'Yes', Cc, 'E9', '4.1.230'),
    ('2.4.6', MT, 'S2', 'S2.36(a)-(d)', 'GHG targets: gases, scopes, gross or net, sectoral decarbonisation approach', 'Yes', Cc, 'E9', '4.1.230'),
    ('2.4.6', MT, 'S2', 'S2.36(e), B70-B71', 'Planned use of carbon credits to achieve net GHG targets', 'No', 'N/A', 'E9', '4.1.240'),
    ('3.4', MT, 'S1', 'S1.46, 48', 'Other SROs: metrics required or used, incl. industry-specific metrics (SASB)', 'Yes', Ip, 'E8', '3.5.10'),
    ('3.4', MT, 'S1', 'S1.51', 'Other SROs: targets - metric, quantitative / qualitative target, period, base period, milestones, performance, revisions', 'Yes', Ip, 'E9', '3.5.20'),
]


def build_f1(X):
    ws, r = X.sheet(F1, 'F1', 'ISSB disclosure checklist and completion tracker',
                    'Track compliance with every specific disclosure requirement of IFRS S1 and IFRS S2 that applies to the entity (second year of application: general, climate and other topics). Requirements are summarised - always read the Standard itself. Status feeds the Dashboard and the statement-of-compliance assessment (A1).',
                    'IFRS S1 (all disclosure paragraphs and Appendices B, E); IFRS S2 (all disclosure paragraphs and Appendices B, C)',
                    'KPMG Guide to sustainability reporting - Disclosure checklist (Oct 2026) chapters 1-3 and App. 2-3; HB cross-references per row',
                    [7, 8, 20, 9, 18, 62, 11, 13, 12, 24, 20, 13, 14, 30], prep='Sok D.', tab=C['purple'], status='In progress', obj_h=44)
    n = len(CHK)
    t0 = r + 9            # header row of the detailed table (0-based)
    a, b = t0 + 1, t0 + n
    pil, app, sta = f'{A(a, 2)}:{A(b, 2)}', f'{A(a, 6)}:{A(b, 6)}', f'{A(a, 7)}:{A(b, 7)}'
    X.sec(ws, r, 'Completion summary by pillar (applicable requirements only)', 14)
    X.hdr(ws, r + 1, ['', '', 'Pillar', '', 'Applicable', 'Complete / In progress / Not started', 'Complete', 'In progress', 'Not started', 'N/A (not applicable)', '% complete'])
    for i, p in enumerate(PILLARS + ['TOTAL']):
        R = r + 2 + i
        tot = p == 'TOTAL'
        X.put(ws, R, 2, p, bold=True); X.put(ws, R, 0, ''); X.put(ws, R, 1, ''); X.put(ws, R, 3, '')
        if tot:
            for j in (4, 6, 7, 8, 9):
                X.put(ws, R, j, f'=SUM({A(r + 2, j)}:{A(R - 1, j)})', bold=True)
        else:
            X.put(ws, R, 4, f'=COUNTIFS({pil},{A(R, 2)},{app},"Yes")')
            X.put(ws, R, 6, f'=COUNTIFS({pil},{A(R, 2)},{app},"Yes",{sta},"Complete")')
            X.put(ws, R, 7, f'=COUNTIFS({pil},{A(R, 2)},{app},"Yes",{sta},"In progress")')
            X.put(ws, R, 8, f'=COUNTIFS({pil},{A(R, 2)},{app},"Yes",{sta},"Not started")')
            X.put(ws, R, 9, f'=COUNTIFS({pil},{A(R, 2)},{app},"No")')
        X.put(ws, R, 5, f'={A(R, 6)}&" / "&{A(R, 7)}&" / "&{A(R, 8)}', align='center')
        X.put(ws, R, 10, f'=IF({A(R, 4)}=0,0,{A(R, 6)}/{A(R, 4)})', num_format='0%', bold=True)
    X.ad['f1sum'] = (r + 2, r + 6)       # pillar rows
    tr = r + 7
    X.name('Chk_Total', F1, tr, 4); X.name('Chk_Done', F1, tr, 6); X.name('Chk_IP', F1, tr, 7); X.name('Chk_NS', F1, tr, 8)
    X.name('Chk_Pct', F1, tr, 10)
    H = ['ID', 'DC section', 'Pillar', 'Standard', 'Paragraph', 'Disclosure requirement (summarised)', 'Applicable?', 'Status', 'Working paper', 'Location in draft report', 'Owner', 'Due date', 'Handbook ref', 'Comment']
    loc = {G: 'Basis of preparation', GV: 's.3.1 Governance', ST: 's.3.2 Strategy', RM: 's.3.3 Risk management', MT: 's.3.4 Metrics and targets'}
    rows = []
    for i, c in enumerate(CHK):
        cm = '' if c[6] in (Cc, 'N/A') else ('See F5 action log' if c[6] == Ip else 'Not yet drafted - see F5')
        if c[6] == 'N/A':
            cm = 'Not applicable to the entity - rationale in WP'
        rows.append([f'C-{i + 1:03d}', c[0], c[1], 'IFRS ' + c[2], c[3], c[4], c[5], c[6], c[7], loc[c[1]] if c[5] == 'Yes' else 'n/a', OWN[c[1]], DUE[c[1]], 'HB ' + c[8], cm])
    a2, b2 = X.table(ws, t0, H, rows, inp=(6, 7, 9, 10, 13))
    assert (a2, b2) == (a, b)
    X.dv(ws, a, b, 6, ['Yes', 'No']); X.dv(ws, a, b, 7, ['Complete', 'In progress', 'Not started', 'N/A'])
    X.cf_status(ws, a, b, 7)
    ws.autofilter(t0, 0, b, 13)
    ws.freeze_panes(t0 + 1, 6)


def build_f2(X):
    ws, r = X.sheet(F2, 'F2', 'Judgements, measurement uncertainty, revised comparatives and errors',
                    'Identify the judgements with the most significant effect on the disclosures, the amounts subject to high measurement uncertainty (with sources, assumptions and sensitivities), revisions of comparatives and any prior-period errors.',
                    'IFRS S1.74-82 (judgements and uncertainty), S1.83-86 and B55-B59 (errors), S1.B49-B54 (comparatives)',
                    'HB 2.3.80-2.3.150; DC 1.8, 1.10-1.12',
                    [8, 34, 60, 18, 46, 16, 13], prep='Sok D.', tab=C['purple'], status='In progress')
    X.sec(ws, r, '1. Significant judgements (S1.74-76)', 7)
    J = [('J-01', 'Organisational boundary - operational control over the JV', 'JV agreement gives TTM the right to set operating and EHS policies; 100% of JV emissions included although equity-accounted in the FS', 'E3, E4, E6', 'Alternative (equity share) would include 50%: Scope 1+2 lower by c. 0.7 kt', 'S1.74-75; S2.B27', 'Complete'),
         ('J-02', 'Identification of relevant SROs and material information', 'Which SROs could reasonably be expected to affect prospects; SR-04 judged not separately material', 'A3, A4', 'Qualitative factors overrode quantitative screen for M-08 and M-12', 'S1.75(a), (c)', 'Complete'),
         ('J-03', 'Definition of assets and activities "vulnerable" to climate risks', 'Transition: revenue from brands with binding supplier requirements. Physical: assets inside the 1-in-25-year flood extent', 'E7', 'Using 1-in-100 extent would add part of Factory A warehouse (+USD 2.1m)', 'S1.75; S2.29(b)-(c)', 'In progress'),
         ('J-04', 'Sources of guidance for non-climate topics', 'SASB CG-AA metrics adapted from a brand perspective to a manufacturer\'s own facilities', 'A1, E8', 'Differences from SASB definitions disclosed (S1.50(a))', 'S1.75(b), 57-58', 'Complete'),
         ('J-05', 'Resilience conclusion threshold', 'Strategy judged resilient if modelled annual effect is below 25% of available liquidity', 'C3', 'Threshold is entity-specific', 'S1.75; S2.22', 'In progress'),
         ('J-06', 'Scope 3 category 9 boundary under FOB terms', 'Buyer-controlled freight included in category 9; use-phase (category 11 indirect) excluded', 'E5', 'Category 9 sea freight = c. 1.5 kt', 'S1.75; S2.B32-B33', 'Complete')]
    a, b = X.table(ws, r + 1, ['ID', 'Judgement', 'Description', 'WP', 'Effect / alternative considered', 'ISSB reference', 'Status'], J, inp=(1, 2, 3, 4, 6))
    X.cf_status(ws, a, b, 6)
    r = b + 2
    X.sec(ws, r, '2. Amounts subject to high measurement uncertainty (S1.77-82) with sensitivity', 7)
    H = ['ID', 'Amount / metric', 'Sources of uncertainty and assumptions (S1.78(a)-(b))', 'Reported amount', 'Sensitivity', 'Sensitivity result', 'Status']
    U = [('U-01', 'Scope 3 category 1 - fabric (tCO2e)', 'Secondary industry-average factor for a cotton / polyester mix; no mill-specific data; fibre mix approximated', '=' + ref(E5, X.ad['s3cat'][0], 3), 'Emission factor +/- 30%', lambda R: f'=D{R}*0.3', 'In progress'),
         ('U-02', 'Scope 2 location-based (tCO2e)', 'National grid factor vintage; grid mix changes year to year', '=GHG_S2', 'Grid factor +/- 10%', lambda R: f'=D{R}*0.1', 'Complete'),
         ('U-03', 'Scope 3 category 7 - commuting (tCO2e)', 'Survey sample (19% of workforce); distance and mode split extrapolated', '=' + ref(E5, X.ad['s3cat'][0] + 6, 3), 'Distance +/- 25%', lambda R: f'=D{R}*0.25', 'Complete'),
         ('U-04', 'Fugitive HFC emissions (tCO2e)', 'Top-up quantity used as proxy for leakage; blend composition', '=S1_Fug', 'Top-up records +/- 20%', lambda R: f'=D{R}*0.2', 'Complete'),
         ('U-05', 'Lost contribution from flood stoppage (USD\'000)', 'Estimate of pieces not produced and recovery through overtime', 610, 'Recovery rate +/- 15 pts', lambda R: f'=D{R}*0.15', 'In progress'),
         ('U-06', 'Anticipated financial effects - scenario results 2030 (USD\'000)', 'Forward-looking; depends on carbon price, flood probability and buyer behaviour', '=Scen_B30', 'Difference between scenarios A and B', '=ABS(Scen_A30-Scen_B30)', 'In progress'),
         ('U-07', 'Assets vulnerable to physical risk (USD\'000)', 'Flood-extent mapping resolution; allocation of inventory by site at year end', '=Assets_Physical', 'Alternative flood extent (J-03)', 2100, 'In progress')]
    a, b = X.table(ws, r + 1, H, U, nf={3: '#,##0', 5: '#,##0'}, inp=(1, 2, 4, 6))
    X.cf_status(ws, a, b, 6)
    r = b + 2
    X.sec(ws, r, '3. Revised comparatives and prior-period errors (S1.B50-B59, S1.83-86)', 7)
    H = ['ID', 'Item', 'Nature', 'FY2024 as previously reported', 'FY2024 revised', 'Difference', 'Status']
    Rw = [('R-01', 'Scope 1 - consolidated group (tCO2e)', 'ERROR (S1.83): LPG for canteens omitted in FY2024 inventory - identified by internal audit; corrected by restating the comparative', 1104, 1214, lambda R: f'=E{R}-D{R}', 'Complete'),
          ('R-02', 'Scope 2 - other investees (tCO2e)', 'REVISED ESTIMATE (S1.B50): December 2024 JV consumption was estimated; actual invoice received', 1052, 1070, lambda R: f'=E{R}-D{R}', 'Complete'),
          ('R-03', 'Scope 2 - consolidated group (tCO2e)', 'No revision', 4030, 4030, lambda R: f'=E{R}-D{R}', 'Complete')]
    a, b = X.table(ws, r + 1, H, Rw, nf={5: '#,##0;-#,##0;-'}, inp=(1, 2, 6))
    X.cf_status(ws, a, b, 6)
    X.note(ws, b + 2, 'Revised FY2024 amounts agree to the comparative column in E6. Disclose: the revised amount, the difference and the reasons (S1.B50); for the error: its nature and the correction for each prior period (S1.B58).', 7)
    ws.freeze_panes(9, 2)


def build_f3(X):
    ws, r = X.sheet(F3, 'F3', 'Connectivity - reconciliation of sustainability data to the financial statements',
                    'Demonstrate that financial data and assumptions used in the sustainability-related financial disclosures are consistent with the related financial statements, explain significant differences, and perform analytical reasonableness tests of activity data against the general ledger.',
                    'IFRS S1.21-24, B39-B44; IFRS S2.15-21 (financial effects)',
                    'HB 6.1.80-6.1.130; DC 1.4',
                    [8, 40, 20, 16, 16, 14, 18, 44], prep='Sok D.', tab=C['purple'], status='In progress')
    X.sec(ws, r, '1. Tie-out of amounts used in sustainability disclosures to the financial statements (USD\'000)', 8)
    H = ['Ref', 'Data point used in sustainability disclosures', 'Used in WP', 'Amount per sustainability WPs', 'Amount per financial statements / GL', 'Difference', 'Result', 'FS reference / explanation of difference (S1.B42(c))']
    c2s, c2a, c2b = X.ad['c2cur']
    capex_all = f"SUMIFS({rng(c2s, 4, c2a, c2b)},{rng(c2s, 3, c2a, c2b)},\"PPE*\")"
    T = [('K-01', 'Revenue FY2025', 'E6, E7', '=Rev25', 85400, 'Statement of profit or loss'),
         ('K-02', 'Profit before tax', 'A4, D1', '=PBT25', 5120, 'Statement of profit or loss'),
         ('K-03', 'Property, plant and equipment', 'E7', '=PPE25', 24800, 'Note 9'),
         ('K-04', 'Inventories', 'E7', '=Inv25', 11900, 'Note 12'),
         ('K-05', 'Total additions to PPE', 'E5 cat 2, E7', '=Capex25', 4800, 'Note 9'),
         ('K-06', 'SRO-related capital expenditure (climate and safety)', 'C2, E7', '=' + capex_all, 3130, 'Capex register by project code - subset of K-05'),
         ('K-07', 'Flood - repairs and inventory write-off', 'C2', 605, 605, 'Notes 7 and 12 (420 + 185)'),
         ('K-08', 'Insurance recovery receivable', 'C2', 380, 380, 'Note 14'),
         ('K-09', 'Employee benefits expense', 'C3', '=StaffCost25', 27600, 'Note 6'),
         ('K-10', 'Key management remuneration', 'E7', '=ExecRem25', 1850, 'Note 27'),
         ('K-11', 'Cash and undrawn facilities', 'C2, C3', '=Liquidity25', 9400, 'Notes 16 and 20'),
         ('K-12', 'Carrying amount of assets at Factory B', 'E7', '=Assets_Physical', 12650, 'Fixed-asset register 9,700 + inventory by site 2,950. Difference: goods in transit allocated to B in sustainability data - TO BE RESOLVED (A-012)')]
    rows = [[k, d, w, s_, f, lambda R: f'=D{R}-E{R}', lambda R: f'=IF(ABS(F{R})<1,"Agrees","DIFFERENCE")', e] for k, d, w, s_, f, e in T]
    a, b = X.table(ws, r + 1, H, rows, nf={5: '#,##0;-#,##0;-'}, inp=(1, 2, 7))
    ws.conditional_format(a, 6, b, 6, {'type': 'cell', 'criteria': '==', 'value': '"Agrees"', 'format': X.wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})})
    ws.conditional_format(a, 6, b, 6, {'type': 'cell', 'criteria': '==', 'value': '"DIFFERENCE"', 'format': X.wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
    X.put(ws, b + 1, 1, 'Items agreeing to the financial statements', bold=True)
    X.put(ws, b + 1, 6, f'=COUNTIF({A(a, 6)}:{A(b, 6)},"Agrees")&" of "&COUNTA({A(a, 6)}:{A(b, 6)})', bold=True)
    X.name('Tie_OK', F3, b + 2, 6); X.put(ws, b + 2, 1, '% agreeing', bold=True)
    X.put(ws, b + 2, 6, f'=COUNTIF({A(a, 6)}:{A(b, 6)},"Agrees")/COUNTA({A(a, 6)}:{A(b, 6)})', num_format='0%', bold=True)
    r = b + 5
    X.sec(ws, r, '2. Analytical reasonableness tests - activity data vs general ledger', 8)
    H = ['Ref', 'Test', 'Activity data (WP)', 'Quantity', 'GL expense (USD\'000)', 'Implied unit cost', 'Expected range', 'Conclusion']
    t = X.ad['e1t']
    e = X.ad
    T = [('AR-01', 'Grid electricity - consolidated group: implied tariff (USD/kWh)', 'E1 Factory A + B MWh', f"={ref(E1, t, 1)}+{ref(E1, t, 2)}", '=ElecExp25', lambda R: f'=E{R}/D{R}', '0.12 - 0.16', lambda R: f'=IF(AND(F{R}>=0.12,F{R}<=0.16),"Reasonable","Investigate")'),
         ('AR-02', 'Diesel - generators and vehicles: implied price (USD/litre)', 'E1 group litres', f"={ref(E1, t, 5)}+{ref(E1, t, 6)}+{e['e1_veh_dsl']}", 222, lambda R: f'=E{R}*1000/D{R}', '0.85 - 1.15', lambda R: f'=IF(AND(F{R}>=0.85,F{R}<=1.15),"Reasonable","Investigate")'),
         ('AR-03', 'Wood biomass - group: implied price (USD/tonne)', 'E1 Factory A + B tonnes', f"={ref(E1, t, 8)}+{ref(E1, t, 9)}", 158, lambda R: f'=E{R}*1000/D{R}', '30 - 50', lambda R: f'=IF(AND(F{R}>=30,F{R}<=50),"Reasonable","Investigate")'),
         ('AR-04', 'Fabric: implied price (USD/kg)', 'E1 tonnes purchased', '=' + e['e1_fabric'], 41900, lambda R: f'=E{R}/D{R}', '3.5 - 6.0', lambda R: f'=IF(AND(F{R}>=3.5,F{R}<=6),"Reasonable","Investigate")'),
         ('AR-05', 'Payroll: employee cost per head (USD)', 'E1 average headcount', '=' + e['e1_hc'], '=StaffCost25', lambda R: f'=E{R}*1000/D{R}', '3,500 - 5,000', lambda R: f'=IF(AND(F{R}>=3500,F{R}<=5000),"Reasonable","Investigate")')]
    a, b = X.table(ws, r + 1, H, T, nf={5: '#,##0.000'}, inp=(1, 2, 6))
    for i in (2, 4):
        X.put(ws, a + i, 5, T[i][5](a + i + 1), num_format='#,##0')
    ws.conditional_format(a, 7, b, 7, {'type': 'cell', 'criteria': '==', 'value': '"Reasonable"', 'format': X.wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})})
    ws.conditional_format(a, 7, b, 7, {'type': 'cell', 'criteria': '==', 'value': '"Investigate"', 'format': X.wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
    r = b + 2
    X.sec(ws, r, '3. Connected information between disclosures (S1.21(b), B40-B44) - consistency checks', 8)
    T = [('X-01', 'Flood (CR-01): strategy narrative (C1) <-> current financial effect (C2) <-> physical-risk metric (E7) <-> adaptation target T5 (E9) <-> FS notes 7, 9, 12, 14', 'Consistent'),
         ('X-02', 'Buyer requirements (CR-03): transition plan (C1) <-> GHG emissions (E6) <-> target T1 (E9) <-> revenue concentration disclosure in FS note 4 (segment / major customers)', 'Consistent'),
         ('X-03', 'Useful lives of wood-fired boilers (FS note 9: 15 years) vs planned replacement by 2028 in transition plan', 'DIFFERENCE - useful lives to be reassessed under IAS 16.51 (A-011)'),
         ('X-04', 'Impairment testing (IAS 36) cash-flow assumptions for Factory B vs flood probabilities in C3', 'Under review with financial reporting team (A-013)'),
         ('X-05', 'Internal carbon price (E7) vs carbon-price assumptions in C3 scenario A', 'Explained - shadow price is a management tool; scenario prices are external pathways')]
    X.hdr(ws, r + 1, ['Ref', 'Connection tested', '', '', '', '', 'Result', ''])
    for i, (k, d, res) in enumerate(T):
        X.put(ws, r + 2 + i, 0, k)
        ws.merge_range(r + 2 + i, 1, r + 2 + i, 5, d, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.merge_range(r + 2 + i, 6, r + 2 + i, 7, res, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.set_row(r + 2 + i, 30)
    ws.freeze_panes(9, 2)


def build_f4(X):
    ws, r = X.sheet(F4, 'F4', 'Evidence tracker, data controls and assurance readiness',
                    'Record, for each data stream, the source, the control performed, evidence held, testing performed and a data-quality score, to support "reasonable and supportable information" and readiness for external assurance.',
                    'IFRS S1.B8-B10 (reasonable and supportable information), S1.D21-D24 (verifiability), S1.27(b)(ii), S2.6(b)(ii), S2.B40(d) (verified inputs)',
                    'HB 2.3.170, 3.2.20. Assurance is outside ISSB Standards - ISSA 5000 (IAASB) is the reference if assurance is engaged',
                    [7, 30, 9, 20, 28, 18, 34, 11, 11, 9, 10, 10, 12, 11, 30], prep='Ly S.', rev='Internal Audit', tab=C['purple'], status='In progress')
    H = ['ID', 'Data stream', 'WP', 'Pillar', 'Source document / system', 'Data owner', 'Control performed', 'Frequency', 'Evidence on file?', 'Sample tested (n)', 'Exceptions', 'Data-quality score (1-5)', 'Primary or estimated', 'Assurance-ready?', 'Comment']
    Dd = [('Grid electricity', 'E1', MT, 'EDC invoices', 'Finance', 'Invoice-to-meter reconciliation; reviewer sign-off', 'Monthly', 'Yes', 36, 0, 5, 'Primary', ''),
          ('Solar generation', 'E1', MT, 'Inverter portal export', 'Maintenance', 'Monthly export reviewed against bi-directional meter', 'Monthly', 'Yes', 9, 0, 4, 'Primary', ''),
          ('Diesel - generators', 'E1', MT, 'Fuel issue ledger; purchase invoices', 'Maintenance', 'Stock reconciliation (opening + purchases - closing)', 'Monthly', 'Yes', 24, 1, 4, 'Primary', 'One month JV ledger unsigned (A-004)'),
          ('Wood biomass', 'E1', MT, 'Weighbridge tickets / delivery notes', 'Procurement', 'Weighbridge ticket matched to invoice', 'Monthly', 'Partial', 20, 3, 3, 'Primary', 'Factory B deliveries recorded by volume, converted with assumed density (A-005)'),
          ('Vehicle fuels and LPG', 'E1', MT, 'Fuel cards; supplier invoices', 'Logistics', 'Card statement review', 'Monthly', 'Yes', 12, 0, 4, 'Primary', ''),
          ('Refrigerants', 'E1', MT, 'HVAC service records', 'Maintenance', 'Annual top-up log reviewed', 'Annual', 'Partial', 6, 1, 3, 'Primary', 'JV log incomplete'),
          ('Emission factors and GWP', 'E2', MT, 'Published sources', 'Sustainability', 'Annual factor review and source file', 'Annual', 'Partial', 25, 0, 3, 'Estimated', 'Scope 3 factors to be replaced by documented sources (A-002)'),
          ('Purchased goods (fabric, trims)', 'E1 / E5', MT, 'Purchase ledger; packing lists', 'Procurement', 'ERP quantity report reconciled to GL', 'Quarterly', 'Yes', 30, 0, 4, 'Primary', 'Activity primary; factor secondary'),
          ('Logistics tonne-km', 'E5', MT, 'Forwarder reports', 'Logistics', 'Analytical review', 'Annual', 'Partial', 10, 0, 3, 'Estimated', 'Average distances'),
          ('Waste', 'E1', MT, 'Contractor manifests', 'EHS', 'Manifest sign-off', 'Monthly', 'Yes', 24, 0, 4, 'Primary', ''),
          ('Commuting survey', 'E5', MT, 'Survey Aug-2025', 'HR', 'Survey design reviewed', 'Annual', 'Yes', 1, 0, 3, 'Estimated', 'Extrapolated'),
          ('Production volume', 'E1', MT, 'Production system', 'Operations', 'Reconciled to shipments', 'Monthly', 'Yes', 24, 0, 5, 'Primary', ''),
          ('Water withdrawal and wastewater tests', 'E1 / E8', MT, 'Meter logs; lab reports', 'JV EHS', 'Meter calibration; accredited lab', 'Monthly', 'Yes', 24, 0, 4, 'Primary', ''),
          ('OHS incidents and hours worked', 'E1 / E8', MT, 'Incident register; payroll', 'EHS / HR', 'Incident classification review by EHS manager', 'Monthly', 'Yes', 28, 0, 4, 'Primary', ''),
          ('Workforce data', 'E1 / E8', MT, 'HR system', 'HR', 'System report', 'Monthly', 'Yes', 12, 0, 5, 'Primary', ''),
          ('Financial data for metrics', 'E7 / F3', MT, 'GL; fixed-asset register', 'Finance', 'Tie-out to FS (F3)', 'Annual', 'Partial', 12, 1, 4, 'Primary', 'K-12 difference open'),
          ('Governance records', 'B1', GV, 'Minutes; ToR; skills matrix', 'Company Secretary', 'Minutes approved by chair', 'Per meeting', 'Yes', 8, 0, 5, 'Primary', ''),
          ('Remuneration linkage', 'B1 / E7', GV, 'Remuneration committee papers', 'HR', 'Committee approval', 'Annual', 'No', 0, 0, 2, 'Primary', 'FY2025 outcome paper outstanding (A-007)'),
          ('SRO register and materiality', 'A3 / A4', ST, 'Workshop records; ARSC approval', 'Risk Manager', 'ARSC approval of assessment', 'Annual', 'Yes', 1, 0, 4, 'Primary', ''),
          ('Scenario analysis model', 'C3', ST, 'Adviser model; assumptions book', 'CFO', 'Model review by CFO; assumptions approved by SSC', 'Annual', 'Partial', 1, 0, 3, 'Estimated', 'Assumption sources to be documented (A-010)'),
          ('Transition plan and capex', 'C1 / C2', ST, 'Board papers; capex register', 'CFO', 'Board approval', 'Annual', 'Yes', 4, 0, 4, 'Primary', ''),
          ('Risk-management process', 'D1', RM, 'ERM policy; risk register', 'Risk Manager', 'Quarterly register review', 'Quarterly', 'Yes', 4, 0, 4, 'Primary', ''),
          ('Targets and base-year data', 'E9', MT, 'Board minutes; base-year inventory', 'Sustainability', 'Annual recalculation check', 'Annual', 'Partial', 7, 0, 3, 'Primary', 'Base-year (2023) inventory not yet re-performed after R-01 error (A-006)')]
    rows = [[f'D-{i + 1:02d}'] + list(d[:12]) + [lambda R: f'=IF(AND(I{R}="Yes",K{R}=0,L{R}>=4),"Yes","No")', d[12]] for i, d in enumerate(Dd)]
    a, b = X.table(ws, r, H, rows, inp=tuple(range(1, 9)) + (12, 14))
    X.dv(ws, a, b, 8, ['Yes', 'Partial', 'No'])
    X.cf_status(ws, a, b, 8); X.cf_status(ws, a, b, 13)
    ws.conditional_format(a, 11, b, 11, {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})
    r = b + 2
    Sm = [('Data streams', f'=COUNTA({A(a, 0)}:{A(b, 0)})', 'Ev_Total', '#,##0'),
          ('Evidence fully on file (%)', f'=COUNTIF({A(a, 8)}:{A(b, 8)},"Yes")/COUNTA({A(a, 0)}:{A(b, 0)})', 'Ev_Pct', '0%'),
          ('Assurance-ready data streams (%)', f'=COUNTIF({A(a, 13)}:{A(b, 13)},"Yes")/COUNTA({A(a, 0)}:{A(b, 0)})', 'Ev_Ready', '0%'),
          ('Average data-quality score (1-5)', f'=AVERAGE({A(a, 11)}:{A(b, 11)})', 'Ev_DQ', '0.0'),
          ('Exceptions noted in testing', f'=SUM({A(a, 10)}:{A(b, 10)})', 'Ev_Exc', '#,##0'),
          ('Primary-data streams (%)', f'=COUNTIF({A(a, 12)}:{A(b, 12)},"Primary")/COUNTA({A(a, 0)}:{A(b, 0)})', 'Ev_Primary', '0%')]
    X.sec(ws, r, 'Summary (feeds Dashboard)', 6)
    for i, (l, f, nm, nf) in enumerate(Sm):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 3, l, X.f(bold=True, border=1, border_color=C['mid']))
        X.put(ws, r + 1 + i, 4, f, num_format=nf, bold=True, align='left')
        X.name(nm, F4, r + 1 + i, 4)
    X.note(ws, r + 8, 'Data-quality score: 5 = metered / system-generated with reconciliation; 4 = primary records with review; 3 = primary with gaps or documented estimates; 2 = largely estimated; 1 = no support.', 15)
    ws.freeze_panes(9, 2)


def build_f5(X):
    ws, r = X.sheet(F5, 'F5', 'Action log - review points, open items and remediation',
                    'Log every open point arising from preparation and review of the working papers, with owner, due date and status. Ageing and overdue analysis are computed against the status date in A1 and feed the Dashboard.',
                    'Engagement administration; supports IFRS S1.72 (statement of compliance) and S1.67-68 (events after the reporting period)',
                    'DC (all sections) - one action per open checklist theme',
                    [5, 8, 10, 20, 44, 44, 20, 9, 12, 12, 12, 9, 10, 10, 7, 16], prep='Sok D.', tab=C['purple'], status='In progress')
    H = ['No.', 'Action ID', 'WP ref', 'Pillar', 'Issue / review point', 'Action required', 'Owner', 'Priority', 'Raised', 'Due', 'Status', 'Days open', 'Ageing bucket', 'Overdue?', 'Rank (open, by due date)', 'ISSB reference']
    Ac = [('A1', G, 'Confirm jurisdictional status of ISSB adoption and any local modifications', 'Obtain written confirmation from regulator / legal counsel; update A1 s.5', 'Financial Controller', 'Medium', D(2025, 11, 12), D(2026, 1, 30), 'In progress', 'S1.B31-B33'),
          ('E2', MT, 'Scope 3 emission factors are illustrative placeholders', 'Replace with documented, current factors; record source, vintage, geography', 'Head of Sustainability', 'High', D(2025, 11, 18), D(2026, 2, 13), 'In progress', 'S2.B29, B40'),
          ('E5', MT, 'Scope 3 cat. 1 uses only secondary factors although three mills can provide primary GHG data (E11)', 'Obtain and use mill-specific data in E5; request data from remaining mills; update S2.B56 disclosure', 'Procurement', 'High', D(2025, 11, 18), D(2026, 3, 6), 'Open', 'S2.B47-B49'),
          ('F4', MT, 'JV diesel ledger for September unsigned', 'Obtain JV manager sign-off and stock reconciliation', 'JV Finance', 'Low', D(2025, 12, 3), D(2026, 1, 16), 'Closed', 'S2.6(b)(ii)'),
          ('E1', MT, 'Factory B wood deliveries recorded by volume (stere), not weight', 'Install weighbridge procedure; document density assumption and sensitivity', 'Procurement', 'Medium', D(2025, 12, 3), D(2026, 2, 6), 'In progress', 'S1.78'),
          ('E9', MT, 'Base-year (2023) inventory not recalculated after LPG omission error', 'Recalculate base year; assess effect on target T1; disclose per S1.B58', 'Head of Sustainability', 'High', D(2025, 12, 10), D(2026, 2, 20), 'In progress', 'S1.83, B58; S2.34(d)'),
          ('B1', GV, 'FY2025 bonus outcome paper not yet available', 'Obtain remuneration-committee minutes; update E7 metric (g)', 'HR Director', 'Medium', D(2025, 12, 15), D(2026, 2, 27), 'Open', 'S2.6(a)(v), 29(g)'),
          ('B1', GV, 'Two internal-audit findings on GHG data controls open', 'Close findings IA-2025-09 #2 and #4; evidence to F4', 'Financial Controller', 'High', D(2025, 12, 15), D(2026, 2, 13), 'In progress', 'S2.6(b)(ii)'),
          ('D1', RM, 'ERM change log not yet approved by ARSC', 'Table at ARSC meeting 26-Feb-2026', 'Risk Manager', 'Low', D(2026, 1, 8), D(2026, 2, 27), 'Open', 'S2.25(a)(vi)'),
          ('C3', ST, 'Scenario assumptions lack documented external sources', 'Compile assumptions book with sources; CFO sign-off', 'CFO', 'High', D(2026, 1, 8), D(2026, 3, 6), 'In progress', 'S2.22(b)(ii), B11-B15'),
          ('F3', ST, 'Useful lives of wood-fired boilers inconsistent with transition plan', 'Financial-reporting team to reassess useful lives / residual values (IAS 16) before FS sign-off', 'Financial Controller', 'High', D(2026, 1, 14), D(2026, 2, 27), 'Open', 'S1.21(b)(ii), B42'),
          ('F3', G, 'K-12: carrying amount of Factory B assets differs from fixed-asset register by USD 150k', 'Reconcile goods-in-transit allocation; correct E7 if needed', 'Financial Controller', 'Medium', D(2026, 1, 14), D(2026, 2, 13), 'Open', 'S1.23, B42(c)'),
          ('F3', ST, 'Impairment-test assumptions for Factory B to reflect flood scenarios', 'Align IAS 36 cash-flow forecasts with C3; document in FS working papers', 'CFO', 'High', D(2026, 1, 14), D(2026, 3, 6), 'In progress', 'S1.23; S2.16(b)'),
          ('C2', ST, 'Anticipated financial effect of boiler conversion not quantified', 'Complete engineering study; if still not quantifiable, finalise S2.19-21 explanation', 'CFO', 'Medium', D(2026, 1, 20), D(2026, 3, 13), 'Open', 'S2.19-21'),
          ('C3', ST, 'Resilience assessment for non-climate SROs not drafted', 'Draft qualitative resilience assessment for SR-01 to SR-03', 'Head of Sustainability', 'Medium', D(2026, 1, 20), D(2026, 3, 6), 'Open', 'S1.41'),
          ('E7', MT, 'Definition of revenue "aligned with opportunities" needs approval', 'SSC to approve definition; apply consistently to comparatives next year', 'CFO', 'Medium', D(2026, 1, 27), D(2026, 2, 27), 'In progress', 'S2.29(d); S1.74'),
          ('E8', MT, 'SASB metric codes to be verified against current version; adapted metrics need S1.50 disclosure', 'Verify codes; draft "entity-developed metric" definitions', 'Head of Sustainability', 'Low', D(2026, 1, 27), D(2026, 3, 6), 'Open', 'S1.49-50, 58(a)'),
          ('A1', G, 'Events after the reporting period review not yet performed', 'Perform review to authorisation date (flood season claims, buyer notices, regulation)', 'Financial Controller', 'High', D(2026, 2, 3), D(2026, 3, 20), 'Open', 'S1.67-68'),
          ('F2', G, 'Judgements and uncertainty note in draft report incomplete', 'Draft note from F2; include sensitivities for U-01 and U-06', 'Sok D.', 'Medium', D(2026, 2, 3), D(2026, 3, 6), 'In progress', 'S1.74-78'),
          ('A4', G, 'Aggregation review - energy-related disclosures risk obscuring flood information', 'Review draft for obscured or duplicated information', 'Head of Sustainability', 'Low', D(2026, 2, 10), D(2026, 3, 13), 'Open', 'S1.B29-B30'),
          ('E6', MT, 'Comparative restatement (R-01) disclosure drafted', 'None - reviewed and cleared', 'Sok D.', 'Medium', D(2025, 12, 10), D(2026, 1, 23), 'Closed', 'S1.B58'),
          ('A3', ST, 'Time-horizon definitions to be linked to planning horizons', 'None - updated and cleared', 'Chan M.', 'Low', D(2025, 11, 12), D(2025, 12, 19), 'Closed', 'S2.10(d)'),
          ('E4', MT, 'Contractual-instrument disclosure wording', 'None - cleared', 'Chan M.', 'Low', D(2025, 11, 25), D(2025, 12, 19), 'Closed', 'S2.29(a)(v)'),
          ('E11', ST, 'New supply-chain SRO (SR-07) not yet in the materiality assessment (A4) or strategy response (C1)', 'Assess materiality; document response, resourcing and metrics; disclose where risks are concentrated', 'Head of Sustainability', 'High', D(2026, 2, 12), D(2026, 3, 13), 'Open', 'S1.30, 32-33, B2-B5'),
          ('E11', ST, 'Two high-risk suppliers (CN-03, BD-01) without an agreed corrective action plan or assessment', 'Agree corrective action plans or exit plans with Procurement; re-assess by Q2-2026', 'Procurement', 'Medium', D(2026, 2, 12), D(2026, 4, 30), 'Open', 'S1.33(a); SASB CG-AA-430b.2'),
          ('F1', G, 'Statement of compliance cannot be finalised until all checklist items are complete', 'Final checklist review before authorisation; CFO sign-off', 'Keo V. (CFO)', 'High', D(2026, 2, 10), D(2026, 3, 25), 'Open', 'S1.72')]
    n = len(Ac)
    a, b = r + 1, r + n
    st, due, no = f'{A(a, 10)}:{A(b, 10)}', f'{A(a, 9)}:{A(b, 9)}', f'{A(a, 0)}:{A(b, 0)}'
    rows = []
    for i, x in enumerate(Ac):
        rows.append([i + 1, f'A-{i + 1:03d}', x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], x[8],
                     lambda R: f'=IF(K{R}="Closed","",StatusDate-I{R})',
                     lambda R: f'=IF(K{R}="Closed","",IF(L{R}<=30,"0-30 days",IF(L{R}<=60,"31-60 days",IF(L{R}<=90,"61-90 days","Over 90 days"))))',
                     lambda R: f'=IF(AND(K{R}<>"Closed",J{R}<StatusDate),"Overdue","")',
                     lambda R: f'=IF(K{R}="Closed","",COUNTIFS({st},"<>Closed",{due},"<"&J{R})+COUNTIFS({st},"<>Closed",{due},J{R},{no},"<"&A{R})+1)',
                     x[9]])
    a2, b2 = X.table(ws, r, H, rows, nf={11: '0', 14: '0'}, inp=(2, 3, 4, 5, 6, 7, 10))
    assert (a2, b2) == (a, b)
    X.dv(ws, a, b, 10, ['Open', 'In progress', 'Closed']); X.dv(ws, a, b, 7, ['High', 'Medium', 'Low'])
    X.cf_status(ws, a, b, 10); X.cf_status(ws, a, b, 7); X.cf_status(ws, a, b, 13)
    ws.autofilter(r, 0, b, 15)
    X.ad['act'] = (a, b)
    r = b + 2
    Sm = [('Total actions', f'=COUNTA({A(a, 1)}:{A(b, 1)})', 'Act_Total'), ('Open or in progress', f'=COUNTIF({st},"<>Closed")', 'Act_Open'),
          ('Closed', f'=COUNTIF({st},"Closed")', 'Act_Closed'), ('Overdue at status date', f'=COUNTIF({A(a, 13)}:{A(b, 13)},"Overdue")', 'Act_Overdue'),
          ('High-priority open', f'=COUNTIFS({st},"<>Closed",{A(a, 7)}:{A(b, 7)},"High")', 'Act_High')]
    X.sec(ws, r, 'Summary (feeds Dashboard)', 6)
    for i, (l, f, nm) in enumerate(Sm):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 3, l, X.f(bold=True, border=1, border_color=C['mid']))
        X.put(ws, r + 1 + i, 4, f, bold=True, align='left')
        X.name(nm, F5, r + 1 + i, 4)
    ws.freeze_panes(9, 2)


def build_index(X):
    ws = X.idx
    ws.hide_gridlines(2)
    for j, w in enumerate([8, 18, 58, 50, 14, 22, 13, 46]):
        ws.set_column(j, j, w)
    ws.merge_range(0, 0, 0, 7, CLIENT, X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], font_size=14, valign='vcenter'))
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, 7, 'WORKING PAPER INDEX - Sustainability-related financial disclosures under IFRS S1 and IFRS S2  |  ' + PERIOD, X.f(bold=True, font_size=12, font_color=C['navy']))
    ws.merge_range(2, 0, 2, 7, 'IMPORTANT: the client, all data and all amounts are fictional and for teaching / illustration only. Emission factors marked "illustrative" must be replaced before real use. Requirements are summarised - refer to the Standards. '
                   'Colour code: blue font = input, black = formula, green = cross-working-paper link, yellow fill = key assumption. Guidance abbreviations: HB = KPMG IFRS Sustainability handbook (Sep 2026); MG = Materiality How-to guide; DC = Disclosure checklist (Oct 2026). ISSA 5000 (assurance) is referenced by topic in G1 - it was not among the source documents.',
                   X.f(italic=True, text_wrap=True, font_color='#9C0006', bg_color='#FFF2CC', border=1, border_color=C['mid']))
    ws.set_row(2, 46)
    ws.merge_range(3, 0, 3, 1, '', X.f(font_color=C['blue'], underline=1, bold=True))
    ws.write_formula(3, 0, '=HYPERLINK("#\'Dashboard\'!A1","<< Dashboard")', X.f(font_color=C['blue'], underline=1, bold=True))
    ws.set_landscape(); ws.fit_to_pages(1, 0)
    H = ['WP ref', 'Sheet (click to open)', 'Title', 'ISSB reference', 'Prepared by', 'Reviewed by', 'Status', 'Guidance used']
    X.hdr(ws, 4, H)
    secs = {'A': 'A. FOUNDATION - reporting entity, value chain, SROs and materiality', 'B': 'B. GOVERNANCE', 'C': 'C. STRATEGY', 'D': 'D. RISK MANAGEMENT', 'E': 'E. METRICS AND TARGETS', 'F': 'F. COMPLETION, CONNECTIVITY AND QUALITY', 'G': 'G. ASSURANCE READINESS (ISSA 5000)'}
    r = 5
    last = ''
    first = None
    for w in X.wps:
        k = w['wp'][0]
        if k != last:
            ws.merge_range(r, 0, r, 7, secs[k], X.f(bold=True, bg_color=C['sec'], font_color=C['navy'], border=1, border_color=C['mid']))
            last = k
            r += 1
        if first is None:
            first = r
        X.put(ws, r, 0, w['wp'], bold=True)
        ws.write_url(r, 1, f"internal:'{w['sheet']}'!A1", X.f(font_color=C['blue'], underline=1, border=1, border_color=C['mid']), w['sheet'])
        X.put(ws, r, 2, w['title']); X.put(ws, r, 3, w['issb']); X.put(ws, r, 4, w['prep']); X.put(ws, r, 5, w['rev'])
        X.put(ws, r, 6, w['status'], font_color=C['inp']); X.put(ws, r, 7, w['guide'])
        r += 1
    X.dv(ws, first, r - 1, 6, ['Reviewed', 'In review', 'In progress', 'Not started'])
    X.cf_status(ws, first, r - 1, 6)
    st = f'{A(first, 6)}:{A(r - 1, 6)}'
    r += 1
    for i, (l, f, nm) in enumerate([('Working papers', f'=COUNTIF({st},"*")', 'WP_Total'), ('Reviewed', f'=COUNTIF({st},"Reviewed")', 'WP_Rev'),
                                    ('In review', f'=COUNTIF({st},"In review")', 'WP_InRev'), ('In progress', f'=COUNTIF({st},"In progress")', 'WP_IP'),
                                    ('Not started', f'=COUNTIF({st},"Not started")', 'WP_NS')]):
        X.put(ws, r + i, 2, l, bold=True); X.put(ws, r + i, 3, f, bold=True, align='left')
        X.name(nm, 'Index', r + i, 3)
    ws.freeze_panes(5, 2)


def build_calc(X):
    """Dashboard data sheet: every chart on the Dashboard reads from here."""
    S = 'Calc'
    ws = X.wb.add_worksheet(S)
    ws.hide_gridlines(2)
    ws.set_tab_color('#7F7F7F')
    ws.set_column(0, 0, 34); ws.set_column(1, 8, 15)
    ws.merge_range(0, 0, 0, 8, 'DASHBOARD DATA - all cells are formulas linked to the working papers (do not overwrite)', X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], font_size=12))
    cz = {}
    # A. checklist by pillar
    r = 2
    X.hdr(ws, r, ['A. Disclosure status by pillar', 'Complete', 'In progress', 'Not started', '% complete'])
    p0, p1 = X.ad['f1sum']
    for i in range(5):
        X.put(ws, r + 1 + i, 0, '=' + ref(F1, p0 + i, 2))
        for j, c in enumerate((6, 7, 8, 10)):
            X.put(ws, r + 1 + i, 1 + j, '=' + ref(F1, p0 + i, c), num_format='0%' if c == 10 else '0')
    cz['pil'] = (r + 1, r + 5)
    # B. GHG by scope
    r = 10
    X.hdr(ws, r, ['B. GHG emissions by scope (tCO2e)', 'FY2025', 'FY2024', 'Share of total', 'Chart label'])
    for i, (l, f, py) in enumerate([('Scope 1', '=GHG_S1', '=GHG_S1_PY'), ('Scope 2 (location-based)', '=GHG_S2', '=GHG_S2_PY'), ('Scope 3', '=GHG_S3', '')]):
        X.put(ws, r + 1 + i, 0, l); X.put(ws, r + 1 + i, 1, f); X.put(ws, r + 1 + i, 2, py if py else 'n/a')
        X.put(ws, r + 1 + i, 3, f'={A(r + 1 + i, 1)}/GHG_Total', num_format='0.0%')
        X.put(ws, r + 1 + i, 4, f'={A(r + 1 + i, 0)}&": "&TEXT({A(r + 1 + i, 1)},"#,##0")&" ("&TEXT({A(r + 1 + i, 3)},"0.0%")&")"')
    cz['scope'] = (r + 1, r + 3)
    # C. Scope 1+2 by source
    r = 16
    X.hdr(ws, r, ['C. Scope 1 + 2 by source (tCO2e)', 'FY2025'])
    for i, (l, f) in enumerate([('Grid electricity (Scope 2)', '=GHG_S2'), ('Diesel generators', '=S1_Gen'), ('Refrigerants (HFCs)', '=S1_Fug'), ('Vehicles', '=S1_Mob'),
                                ('Biomass boilers (CH4, N2O)', '=S1_Bio'), ('LPG', '=S1_LPG')]):
        X.put(ws, r + 1 + i, 0, l); X.put(ws, r + 1 + i, 1, f)
    cz['src'] = (r + 1, r + 6)
    # D. SRO ratings
    r = 25
    s, a, b = X.ad['sro']
    X.hdr(ws, r, ['D. Relevant SROs by rating', 'Count'])
    for i, l in enumerate(['Critical', 'High', 'Medium', 'Low']):
        X.put(ws, r + 1 + i, 0, l)
        X.put(ws, r + 1 + i, 1, f'=COUNTIFS({rng(s, 12, a, b)},"{l}",{rng(s, 13, a, b)},"Yes")')
    cz['sro'] = (r + 1, r + 4)
    # E. action ageing
    r = 32
    a, b = X.ad['act']
    X.hdr(ws, r, ['E. Open actions by age', 'Count'])
    for i, l in enumerate(['0-30 days', '31-60 days', '61-90 days', 'Over 90 days']):
        X.put(ws, r + 1 + i, 0, l); X.put(ws, r + 1 + i, 1, f'=COUNTIF({rng(F5, 12, a, b)},"{l}")')
    cz['age'] = (r + 1, r + 4)
    # F. monthly Scope 1+2
    r = 39
    X.hdr(ws, r, ['F. Monthly emissions FY2025', 'Scope 1 - generators and boilers (tCO2e)', 'Scope 2 - grid electricity (tCO2e)', 'Production (\'000 pcs)', 'Intensity (kg CO2e per garment)'])
    m0, m1 = X.ad['e1m']
    for i in range(12):
        R = m0 + i
        X.put(ws, r + 1 + i, 0, '=' + ref(E1, R, 0))
        X.put(ws, r + 1 + i, 1, f"=(SUM({ref(E1, R, 5)}:{ref(E1, R, 7).split('!')[1]})*EF_Diesel_Stat+SUM({ref(E1, R, 8)}:{ref(E1, R, 10).split('!')[1]})*EF_Wood)/1000", num_format='#,##0.0')
        X.put(ws, r + 1 + i, 2, f"=SUM({ref(E1, R, 1)}:{ref(E1, R, 3).split('!')[1]})*EF_Grid/1000", num_format='#,##0.0')
        X.put(ws, r + 1 + i, 3, f"={ref(E1, R, 11)}+{ref(E1, R, 12)}")
        X.put(ws, r + 1 + i, 4, f'=({A(r + 1 + i, 1)}+{A(r + 1 + i, 2)})/{A(r + 1 + i, 3)}', num_format='0.000')
    cz['mon'] = (r + 1, r + 12)
    X.note(ws, r + 13, 'Monthly series covers sources metered monthly (generators, boilers, grid electricity). Vehicles, LPG and refrigerants are recorded annually (E3).', 5)
    # G. targets
    r = 55
    a, b = X.ad['tgt']
    X.hdr(ws, r, ['G. Targets', 'Progress to target', 'Expected (linear path)', 'Status'])
    for i in range(b - a + 1):
        X.put(ws, r + 1 + i, 0, f'={ref(E9, a + i, 0)}&" "&{ref(E9, a + i, 24)}')
        X.put(ws, r + 1 + i, 1, '=' + ref(E9, a + i, 13), num_format='0%')
        X.put(ws, r + 1 + i, 2, '=' + ref(E9, a + i, 14), num_format='0%')
        X.put(ws, r + 1 + i, 3, '=' + ref(E9, a + i, 15))
    cz['tgt'] = (r + 1, r + 1 + b - a)
    # H. readiness indicators
    r = 65
    X.hdr(ws, r, ['H. Reporting-readiness indicators', '%'])
    for i, (l, f) in enumerate([('Disclosure checklist complete', '=Chk_Pct'), ('Working papers reviewed', '=WP_Rev/WP_Total'), ('Evidence fully on file', '=Ev_Pct'),
                                ('Assurance-ready data streams', '=Ev_Ready'), ('FS tie-out items agreeing', '=Tie_OK'), ('Scope 3 from primary activity data', '=S3_Primary'),
                                ('Targets on track / achieved', '=Tgt_OnTrack/(Tgt_OnTrack+Tgt_Behind)'),
                                ('ESG KPIs with evidence on file', '=KPI_Ev/KPI_Total'), ('ISSA 5000 readiness items complete', '=ISSA_Ready'), ('Tier-1 spend social-audited (E11)', '=SC_Soc')]):
        X.put(ws, r + 1 + i, 0, l); X.put(ws, r + 1 + i, 1, f, num_format='0%')
    cz['rdy'] = (r + 1, r + 10)
    # I. top scope 3 categories
    r = 77
    a, b = X.ad['s3cat']
    X.hdr(ws, r, ['I. Largest Scope 3 categories', 'tCO2e', '% of Scope 3'])
    v = rng(E5, 3, a, b)
    for i in range(5):
        X.put(ws, r + 1 + i, 1, f'=LARGE({v},{i + 1})')
        X.put(ws, r + 1 + i, 0, f'="Cat. "&INDEX({rng(E5, 0, a, b)},MATCH({A(r + 1 + i, 1)},{v},0))&" "&INDEX({rng(E5, 1, a, b)},MATCH({A(r + 1 + i, 1)},{v},0))')
        X.put(ws, r + 1 + i, 2, f'={A(r + 1 + i, 1)}/GHG_S3', num_format='0.0%')
    cz['s3'] = (r + 1, r + 5)
    # J. priority actions
    r = 85
    a, b = X.ad['act']
    X.hdr(ws, r, ['J. Priority open actions (by due date)', 'Action ID', 'WP', 'Owner', 'Due', 'Status', 'Overdue?', 'Rank'])
    rk = rng(F5, 14, a, b)
    for i in range(7):
        X.put(ws, r + 1 + i, 7, i + 1, font_color='#000000')
        for j, c in ((0, 5), (1, 1), (2, 2), (3, 6), (4, 9), (5, 10), (6, 13)):
            X.put(ws, r + 1 + i, j, f'=IFERROR(INDEX({rng(F5, c, a, b)},MATCH({A(r + 1 + i, 7)},{rk},0)),"")', num_format='dd-mmm-yy' if c == 9 else 'General')
    cz['pri'] = (r + 1, r + 7)
    X.cz = cz
    return ws


def build_dashboard(X):
    ws = X.dash
    wb = X.wb
    S = 'Calc'
    cz = X.cz
    PAGE = '#EEF2F7'
    ws.hide_gridlines(2)
    ws.set_zoom(85)
    navy = X.f(bg_color=C['navy'], font_color='#FFFFFF')
    page = X.f(bg_color=PAGE)
    ws.set_column(0, 0, 25, navy)
    ws.set_column(1, 1, 1.4, page)
    ws.set_column(2, 25, 8.43, page)
    ws.set_column(26, 26, 1.4, page)
    for rr in range(0, 60):
        ws.set_row(rr, 15)
    # ---------- sidebar
    ws.merge_range(0, 0, 1, 0, 'ISSB REPORTING', X.f(bg_color=C['navy'], font_color='#FFFFFF', bold=True, font_size=14, align='center', valign='vcenter'))
    ws.write(2, 0, 'WORKING PAPERS  FY2025', X.f(bg_color=C['navy'], font_color='#9FB8E0', bold=True, font_size=8, align='center'))
    nav = [('Dashboard', 'Dashboard', True), ('WP Index', 'Index', False), (None, 'FOUNDATION', None)] + \
          [(w['wp'] + '  ' + w['sheet'][len(w['wp']) + 1:], w['sheet'], False) for w in X.wps]
    grp = {'B': 'CORE CONTENT', 'E': 'METRICS AND TARGETS', 'F': 'COMPLETION', 'G': 'ASSURANCE'}
    r = 4
    seen = set()
    lf = X.f(bg_color=C['navy'], font_color='#FFFFFF', font_size=9, indent=1, valign='vcenter')
    hf = X.f(bg_color=C['navy'], font_color='#7FA3D9', font_size=8, bold=True, indent=1, valign='bottom')
    for lab, sh, act in nav:
        while r in (8, 22, 36):      # thin spacer rows - keep sidebar labels off them
            r += 1
        if lab is None:
            ws.write(r, 0, sh, hf); r += 1
            continue
        k = lab[0]
        if not act and sh not in ('Index',) and k in grp and k not in seen:
            seen.add(k)
            ws.write(r, 0, grp[k], hf); r += 1
            while r in (8, 22, 36):
                r += 1
        f = X.f(bg_color=C['blue'], font_color='#FFFFFF', bold=True, font_size=10, indent=1, valign='vcenter') if act else lf
        ws.write_formula(r, 0, f'=HYPERLINK("#\'{sh}\'!A1","{lab}")', f)
        r += 1
    ws.write_formula(r + 1, 0, '=HYPERLINK("#\'Calc\'!A1","Dashboard data (Calc)")', lf)
    ws.merge_range(r + 3, 0, r + 6, 0, 'Sample client and data are fictional - for teaching and illustration only.', X.f(bg_color=C['navy'], font_color='#9FB8E0', italic=True, font_size=8, text_wrap=True, indent=1))
    # ---------- top band
    tb = X.f(bg_color=C['navy'], font_color='#FFFFFF', bold=True, font_size=18, valign='vcenter', indent=1)
    ws.merge_range(0, 2, 1, 11, 'SUSTAINABILITY REPORTING DASHBOARD', tb)
    ws.merge_range(2, 2, 2, 11, 'IFRS S1 General Requirements  |  IFRS S2 Climate-related Disclosures  |  100% formula-driven from the working papers',
                   X.f(bg_color=C['navy'], font_color='#9FB8E0', font_size=9, indent=1, valign='vcenter'))
    boxes = [(12, 16, 'Entity', '="Tonle Threads Manufacturing Plc."', None), (17, 19, 'Reporting period end', '=PeriodEnd', 'dd-mmm-yyyy'),
             (20, 22, 'Status date (A1)', '=StatusDate', 'dd-mmm-yyyy'), (23, 25, 'Year of application', '="2nd year (S1 + S2)"', None)]
    for c0, c1, lab, f, nf in boxes:
        ws.merge_range(0, c0, 0, c1, lab, X.f(bg_color=C['navy'], font_color='#9FB8E0', font_size=8, valign='bottom', indent=1))
        k = dict(bg_color='#FFFFFF', font_color=C['navy'], bold=True, font_size=10, valign='vcenter', indent=1, border=5, border_color=C['navy'])
        if nf:
            k['num_format'] = nf
            k['align'] = 'left'
        ws.merge_range(1, c0, 2, c1, '', X.f(**k))
        ws.write_formula(1, c0, f, X.f(**k))
    ws.set_row(3, 8)
    # ---------- KPI tiles
    tiles = [('DISCLOSURE REQUIREMENTS IN SCOPE', '=Chk_Total', '#,##0', '=Chk_Done&" complete  |  "&Chk_IP&" in progress  |  "&Chk_NS&" not started"', C['blue']),
             ('DISCLOSURE READINESS', '=Chk_Pct', '0%', '="Statement of compliance: "&IF(Chk_Pct=1,"supportable","not yet supportable")', C['green']),
             ('TOTAL GHG EMISSIONS (tCO2e)', '=GHG_Total', '#,##0', '="Scope 1: "&TEXT(GHG_S1,"#,##0")&"  |  Scope 2: "&TEXT(GHG_S2,"#,##0")&"  |  Scope 3: "&TEXT(GHG_S3,"#,##0")', C['teal']),
             ('SCOPE 1+2 vs FY2024', '=GHG_S12/GHG_S12_PY-1', '+0.0%;-0.0%', '=TEXT(GHG_S12,"#,##0")&" tCO2e vs "&TEXT(GHG_S12_PY,"#,##0")&"  |  "&TEXT(GHG_Int_Rev,"0.0")&" t per USD m revenue"', C['purple']),
             ('RELEVANT RISKS AND OPPORTUNITIES', '=SRO_Relevant', '0', '="of "&SRO_Total&" identified  |  "&Mat_Items&" material information items"', C['orange']),
             ('OVERDUE ACTIONS', '=Act_Overdue', '0', '=Act_Open&" open actions  |  "&Act_High&" high priority"', C['red'])]
    for i, (lab, f, nf, sub, col) in enumerate(tiles):
        c0, c1 = 2 + i * 4, 5 + i * 4
        bd = dict(left=5, right=5, left_color=PAGE, right_color=PAGE)
        ws.merge_range(4, c0, 4, c1, lab, X.f(bg_color='#FFFFFF', font_color='#5B6B7F', bold=True, font_size=8, indent=1, valign='vcenter', top=2, top_color=col, **bd))
        ws.merge_range(5, c0, 6, c1, '', X.f(bg_color='#FFFFFF', **bd))
        ws.write_formula(5, c0, f, X.f(bg_color='#FFFFFF', font_color=col, bold=True, font_size=22, num_format=nf, align='left', indent=1, valign='vcenter', **bd))
        ws.merge_range(7, c0, 7, c1, '', X.f(bg_color='#FFFFFF', **bd))
        ws.write_formula(7, c0, sub, X.f(bg_color='#FFFFFF', font_color='#5B6B7F', font_size=7.5, indent=1, valign='vcenter', **bd))
    ws.set_row(4, 16); ws.set_row(5, 18); ws.set_row(6, 18); ws.set_row(7, 15); ws.set_row(8, 8)

    # ---------- charts
    def base(ch, title, legend=True):
        ch.set_title({'name': title, 'name_font': {'name': 'Arial', 'size': 10, 'bold': True, 'color': C['navy']}, 'overlay': False, 'layout': {'x': 0.03, 'y': 0.03}})
        ch.set_chartarea({'border': {'color': C['mid']}, 'fill': {'color': '#FFFFFF'}})
        ch.set_plotarea({'fill': {'none': True}})
        if legend:
            ch.set_legend({'position': 'bottom', 'font': {'name': 'Arial', 'size': 8}})
        else:
            ch.set_legend({'none': True})
        ch.set_size({'width': 380, 'height': 258})
        return ch

    def cat(a, b, c=0):
        return [S, a, c, b, c]
    axf = {'num_font': {'name': 'Arial', 'size': 8}, 'line': {'color': C['mid']}, 'major_gridlines': {'visible': False}}
    axg = {'num_font': {'name': 'Arial', 'size': 8}, 'line': {'none': True}, 'major_gridlines': {'visible': True, 'line': {'color': '#E6EAF0', 'width': 0.5}}}
    dl = {'value': True, 'font': {'name': 'Arial', 'size': 8}}
    R1, R2 = 9, 23
    # 1 disclosure status by pillar
    a, b = cz['pil']
    ch = base(wb.add_chart({'type': 'bar', 'subtype': 'stacked'}), '1. Disclosure checklist status by pillar')
    for j, (nm, col) in enumerate((('Complete', C['green']), ('In progress', C['amber']), ('Not started', C['red']))):
        ch.add_series({'name': nm, 'categories': cat(a, b), 'values': [S, a, 1 + j, b, 1 + j], 'fill': {'color': col}, 'gap': 60,
                       'data_labels': {'value': True, 'num_format': '0;;', 'font': {'name': 'Arial', 'size': 8, 'color': '#FFFFFF', 'bold': True}}})
    ch.set_x_axis({'visible': False, 'major_gridlines': {'visible': False}}); ch.set_y_axis(dict(axf, reverse=True))
    ws.insert_chart(R1, 2, ch, {'x_offset': 2, 'y_offset': 2})
    # 2 GHG by scope (doughnut)
    a, b = cz['scope']
    ch = base(wb.add_chart({'type': 'doughnut'}), '2. GHG emissions by scope (tCO2e)')
    ch.add_series({'name': 'FY2025', 'categories': cat(a, b, 4), 'values': [S, a, 1, b, 1],
                   'points': [{'fill': {'color': C['orange']}}, {'fill': {'color': C['blue']}}, {'fill': {'color': C['teal']}}]})
    ch.set_hole_size(60)
    ch.set_legend({'position': 'right', 'font': {'name': 'Arial', 'size': 8}})
    ws.insert_chart(R1, 8, ch, {'x_offset': 2, 'y_offset': 2})
    # 3 Scope 1+2 by source
    a, b = cz['src']
    ch = base(wb.add_chart({'type': 'bar'}), '3. Scope 1 + 2 emissions by source (tCO2e)', legend=False)
    ch.add_series({'name': 'FY2025', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'fill': {'color': C['blue']}, 'gap': 60,
                   'data_labels': dict(dl, num_format='#,##0')})
    ch.set_x_axis({'visible': False, 'major_gridlines': {'visible': False}}); ch.set_y_axis(dict(axf, reverse=True))
    ws.insert_chart(R1, 14, ch, {'x_offset': 2, 'y_offset': 2})
    # 4 SRO ratings
    a, b = cz['sro']
    ch = base(wb.add_chart({'type': 'column'}), '4. Relevant risks and opportunities by rating', legend=False)
    ch.add_series({'name': 'SROs', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'gap': 80, 'data_labels': dl,
                   'points': [{'fill': {'color': C['red']}}, {'fill': {'color': C['orange']}}, {'fill': {'color': C['amber']}}, {'fill': {'color': C['blue']}}]})
    ch.set_y_axis(dict(axg, num_format='0')); ch.set_x_axis(dict(axf))
    ws.insert_chart(R1, 20, ch, {'x_offset': 2, 'y_offset': 2})
    # 5 action ageing
    a, b = cz['age']
    ch = base(wb.add_chart({'type': 'column'}), '5. Open actions - ageing at status date', legend=False)
    ch.add_series({'name': 'Open actions', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'gap': 80, 'data_labels': dl,
                   'points': [{'fill': {'color': C['green']}}, {'fill': {'color': C['amber']}}, {'fill': {'color': C['orange']}}, {'fill': {'color': C['red']}}]})
    ch.set_y_axis(dict(axg, num_format='0')); ch.set_x_axis(dict(axf))
    ws.insert_chart(R2, 2, ch, {'x_offset': 2, 'y_offset': 2})
    # 6 monthly emissions + intensity
    a, b = cz['mon']
    ch = base(wb.add_chart({'type': 'column', 'subtype': 'stacked'}), '6. Monthly Scope 1 + 2 emissions and intensity')
    ch.add_series({'name': 'Scope 2 - grid', 'categories': cat(a, b), 'values': [S, a, 2, b, 2], 'fill': {'color': C['blue']}, 'gap': 50})
    ch.add_series({'name': 'Scope 1 - fuels', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'fill': {'color': C['orange']}})
    ln = wb.add_chart({'type': 'line'})
    ln.add_series({'name': 'kg CO2e per garment', 'categories': cat(a, b), 'values': [S, a, 4, b, 4], 'y2_axis': True,
                   'line': {'color': C['navy'], 'width': 1.75}, 'marker': {'type': 'circle', 'size': 4, 'fill': {'color': C['navy']}, 'border': {'color': C['navy']}}})
    ln.set_y2_axis({'num_font': {'name': 'Arial', 'size': 8}, 'num_format': '0.00', 'line': {'none': True}, 'min': 0})
    ch.combine(ln)
    ch.set_y_axis(dict(axg, num_format='#,##0')); ch.set_x_axis(dict(axf))
    ws.insert_chart(R2, 8, ch, {'x_offset': 2, 'y_offset': 2})
    # 7 targets
    a, b = cz['tgt']
    ch = base(wb.add_chart({'type': 'bar'}), '7. Progress against targets (S2.33-36, S1.51)')
    ch.add_series({'name': 'Progress to target', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'fill': {'color': C['teal']}, 'gap': 40, 'overlap': -10,
                   'data_labels': dict(dl, num_format='0%')})
    ch.add_series({'name': 'Expected (linear path)', 'categories': cat(a, b), 'values': [S, a, 2, b, 2], 'fill': {'color': '#B8C4D6'}})
    ch.set_x_axis(dict(axg, num_format='0%', min=0, max=1)); ch.set_y_axis(dict(axf, reverse=True, crossing='max', num_font={'name': 'Arial', 'size': 8}))
    ws.insert_chart(R2, 14, ch, {'x_offset': 2, 'y_offset': 2})
    # 8 readiness indicators
    a, b = cz['rdy']
    ch = base(wb.add_chart({'type': 'bar'}), '8. Reporting-readiness indicators', legend=False)
    ch.add_series({'name': 'Readiness', 'categories': cat(a, b), 'values': [S, a, 1, b, 1], 'fill': {'color': C['purple']}, 'gap': 50, 'data_labels': dict(dl, num_format='0%')})
    ch.set_x_axis({'visible': False, 'major_gridlines': {'visible': False}, 'min': 0, 'max': 1}); ch.set_y_axis(dict(axf, reverse=True, num_font={'name': 'Arial', 'size': 8}))
    ws.insert_chart(R2, 20, ch, {'x_offset': 2, 'y_offset': 2})
    ws.set_row(22, 8); ws.set_row(36, 8)

    # ---------- bottom tables
    T = 37
    th = X.f(bg_color=C['navy'], font_color='#FFFFFF', bold=True, font_size=9, valign='vcenter', indent=1)
    hd = X.f(bg_color='#DCE5F2', font_color=C['navy'], bold=True, font_size=8, valign='vcenter', border=1, border_color='#FFFFFF')

    def cell(**k):
        d = dict(bg_color='#FFFFFF', font_size=8.5, valign='vcenter', bottom=1, bottom_color='#E6EAF0')
        d.update(k)
        return X.f(**d)
    # priority actions
    ws.merge_range(T, 2, T, 13, 'PRIORITY OPEN ACTIONS (earliest due date first - F5)', th)
    spans = [(2, 2, 'ID', 1), (3, 3, 'WP', 2), (4, 9, 'Issue / review point', 0), (10, 11, 'Owner', 3), (12, 12, 'Due', 4), (13, 13, 'Status', 5)]
    for c0, c1, lab, _ in spans:
        if c0 == c1:
            ws.write(T + 1, c0, lab, hd)
        else:
            ws.merge_range(T + 1, c0, T + 1, c1, lab, hd)
    a, b = cz['pri']
    for i in range(7):
        rr = T + 2 + i
        ws.set_row(rr, 24)
        for c0, c1, lab, src in spans:
            f = '=' + ref(S, a + i, src) + '&""' if src != 4 else '=' + ref(S, a + i, src)
            fm = cell(text_wrap=(src == 0), num_format='dd-mmm-yy' if src == 4 else 'General', align='left')
            if c0 != c1:
                ws.merge_range(rr, c0, rr, c1, '', fm)
            ws.write_formula(rr, c0, f, fm)
    X.cf_status(ws, T + 2, T + 8, 13)
    ws.conditional_format(T + 2, 12, T + 8, 12, {'type': 'formula', 'criteria': f'=AND(ISNUMBER($M{T + 3}),$M{T + 3}<StatusDate)', 'format': wb.add_format({'font_color': C['red'], 'bold': True})})
    # GHG snapshot
    ws.merge_range(T, 14, T, 19, 'GHG SNAPSHOT (tCO2e) - E6', th)
    for c0, c1, lab in ((14, 16, 'Metric'), (17, 17, 'FY2025'), (18, 18, 'FY2024'), (19, 19, 'Change')):
        if c0 == c1:
            ws.write(T + 1, c0, lab, hd)
        else:
            ws.merge_range(T + 1, c0, T + 1, c1, lab, hd)
    G_ = [('Scope 1', '=GHG_S1', '=GHG_S1_PY'), ('Scope 2 (location-based)', '=GHG_S2', '=GHG_S2_PY'), ('Scope 1 + 2', '=GHG_S12', '=GHG_S12_PY'),
          ('Scope 3 (first measured FY2025)', '=GHG_S3', None), ('Biogenic CO2 (memo)', '=GHG_Biogenic', None),
          ('Renewable electricity share', '=RE_Share', None), ('Climate capex (USD\'000)', '=Capex_Climate', None)]
    for i, (l, f, py) in enumerate(G_):
        rr = T + 2 + i
        nf = '0.0%' if 'share' in l else '#,##0'
        ws.merge_range(rr, 14, rr, 16, l, cell(indent=1, bold=(i == 2)))
        ws.write_formula(rr, 17, f, cell(num_format=nf, bold=(i == 2)))
        if py:
            ws.write_formula(rr, 18, py, cell(num_format='#,##0'))
            ws.write_formula(rr, 19, f'={A(rr, 17, False)}/{A(rr, 18, False)}-1', cell(num_format='+0.0%;-0.0%', font_color=C['green'], bold=True))
        else:
            ws.write(rr, 18, 'n/a', cell(align='right', font_color='#8A96A8')); ws.write(rr, 19, '', cell())
    # WP status + scope 3 + cross-industry
    ws.merge_range(T, 20, T, 25, 'CROSS-INDUSTRY METRICS AND STATUS', th)
    for c0, c1, lab in ((20, 23, 'Indicator'), (24, 25, 'Value')):
        ws.merge_range(T + 1, c0, T + 1, c1, lab, hd)
    K_ = [('Revenue exposed to transition risk - S2.29(b)', '=Rev_Transition_Pct', '0%'), ('Assets exposed to physical risk - S2.29(c)', '=Assets_Physical_Pct', '0%'),
          ('Revenue aligned with opportunities - S2.29(d)', '=Rev_Opp_Pct', '0%'), ('Climate capex / total capex - S2.29(e)', '=Capex_Climate_Pct', '0%'),
          ('Internal carbon price USD/t - S2.29(f)', '=ICP', '#,##0'), ('Exec. remuneration climate-linked - S2.29(g)', '=Rem_Pct', '0.0%'),
          ('ESG KPIs in ISSB disclosures (E10)', '=KPI_ISSB&" of "&KPI_Total', 'General')]
    for i, (l, f, nf) in enumerate(K_):
        rr = T + 2 + i
        ws.merge_range(rr, 20, rr, 23, l, cell(indent=1))
        ws.merge_range(rr, 24, rr, 25, '', cell(num_format=nf, bold=True, align='right', font_color=C['navy']))
        ws.write_formula(rr, 24, f, cell(num_format=nf, bold=True, align='right', font_color=C['navy']))
    ws.merge_range(T + 10, 2, T + 10, 25, 'Sources: F1 checklist (KPIs 1-2, chart 1) | E3-E6 GHG inventory (KPIs 3-4, charts 2, 3, 6) | A3-A4 (KPI 5, chart 4) | F5 action log (KPI 6, chart 5) | E9 targets (chart 7) | F3-F4, E10-E12, G1, Index (chart 8). '
                   'Change inputs (blue cells) in the working papers and the dashboard recalculates.', X.f(bg_color=PAGE, font_color='#5B6B7F', italic=True, font_size=8, text_wrap=True, valign='top'))
    ws.set_row(T + 10, 26)
    ws.set_first_sheet(); ws.activate()
    ws.set_landscape(); ws.fit_to_pages(1, 1); ws.set_paper(8)


def build_part3(X):
    build_f1(X); build_f2(X); build_f3(X); build_f4(X); build_f5(X)
    build_g1(X)
    build_index(X)
    build_calc(X)
    build_dashboard(X)
