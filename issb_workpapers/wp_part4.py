"""Part 4: E10 ESG KPI library (mapped to ISSB) and G1 ISSA 5000 assurance readiness."""
from wp_lib import C, A, ref, rng
from wp_part2 import E1, E9

E10, G1 = 'E10 ESG KPIs', 'G1 ISSA 5000'
REQ = 'Yes - required'
MAT = 'Yes - material SRO'
NO = 'No - monitored only'


def build_e10(X):
    ws, r = X.sheet(E10, 'E10', 'ESG key-metric library - Environmental, Social, Governance and supply-chain KPIs mapped to ISSB',
                    'Extend the metric set to a full Environmental / Social / Governance KPI library and decide, KPI by KPI, whether it belongs in the ISSB disclosures. ISSB does not prescribe a fixed ESG KPI list: a metric is disclosed when an ISSB Standard requires it (IFRS S2.29, 32) or when the entity uses it to measure and monitor a risk or opportunity that could reasonably be expected to affect its prospects (IFRS S1.46). Other KPIs are kept for internal monitoring or other frameworks.',
                    'IFRS S1.45-53, S1.55-58 (sources of guidance), S1.B25, B29-B30 (immaterial information; not obscuring); IFRS S2.29, 32, B64-B67',
                    'HB 3.5.10, 4.1.120-4.1.190; DC 2.4.2-2.4.5, 3.4; MG Steps 4-5. KPI structure adapted from "Key metrics for sustainability" (GTS / A. Vizcaya Abdo)',
                    [8, 13, 22, 28, 40, 13, 16, 13, 11, 28, 19, 26, 18, 11, 12, 22], prep='Chan M.', tab=C['green'], status='In review', obj_h=58)
    e = X.ad
    t = e['e1t']
    prod = f"({ref(E1, t, 11)}+{ref(E1, t, 12)})"
    waste = f"({e['e1_w_rec']}+{e['e1_w_land']}+{e['e1_w_haz']})"
    base = ref(E9, e['tgt'][0], 6)
    P, N0, N1, N2 = '0.0%', '#,##0', '#,##0.0', '#,##0.00'
    NA = 'n/a (S1.E6)'
    Env, Soc, Gov = 'Environmental', 'Social', 'Governance'
    K = [
        # pillar, category, KPI, metric, value, unit, fmt, FY2024, SRO, ISSB basis, in ISSB?, metric source, WP, evidence, assurance
        (Env, 'Carbon emissions and climate', 'Total GHG emissions - Scope 1', 'Absolute gross Scope 1, tCO2e', '=GHG_S1', 'tCO2e', N0, '=GHG_S1_PY', 'All climate', 'S2.29(a)(i)(1)', REQ, 'IFRS S2 / GHG Protocol', 'E3, E6', 'Yes', 'Yes'),
        (Env, 'Carbon emissions and climate', 'Total GHG emissions - Scope 2', 'Location-based Scope 2, tCO2e', '=GHG_S2', 'tCO2e', N0, '=GHG_S2_PY', 'All climate', 'S2.29(a)(i)(2), (v)', REQ, 'IFRS S2 / GHG Protocol', 'E4, E6', 'Yes', 'Yes'),
        (Env, 'Carbon emissions and climate', 'Total GHG emissions - Scope 3', 'Absolute gross Scope 3 by category, tCO2e', '=GHG_S3', 'tCO2e', N0, 'n/a (S2.C5)', 'All climate', 'S2.29(a)(i)(3), (vi)', REQ, 'IFRS S2 / GHG Protocol', 'E5, E6', 'Partial', 'No'),
        (Env, 'Carbon emissions and climate', 'Emission reduction progress', 'Change in Scope 1+2 versus 2023 base year', f'=GHG_S12/{base}-1', '%', P, f'=GHG_S12_PY/{base}-1', 'CR-03', 'S2.33-35', REQ, 'IFRS S2 (target T1)', 'E9', 'Partial', 'Yes'),
        (Env, 'Carbon emissions and climate', 'Carbon intensity - revenue', 'Scope 1+2 tCO2e per USD million revenue', '=GHG_Int_Rev', 't / USD m', N1, '=GHG_S12_PY/(Rev24/1000)', 'CR-03', 'S1.46(b), 50; S2.B67', MAT, 'Entity-developed', 'E6', 'Yes', 'Yes'),
        (Env, 'Carbon emissions and climate', 'Carbon intensity - production', 'Scope 1+2 kg CO2e per garment', '=GHG_Int_Pc', 'kg / garment', '0.000', 0.2745, 'CR-03', 'S1.46(b), 50', MAT, 'Entity-developed', 'E6', 'Yes', 'Yes'),
        (Env, 'Energy use and efficiency', 'Total energy consumption', 'Electricity and fuels, all sites', '=Energy_GJ', 'GJ', N0, 171900, 'CR-05, CO-01', 'S1.46(b), 50', MAT, 'Entity-developed (GRI 302-1 definition)', 'E8', 'Yes', 'Yes'),
        (Env, 'Energy use and efficiency', 'Renewable energy share', 'Renewable electricity / total electricity consumed', '=RE_Share', '%', P, 0.0, 'CO-01', 'S2.33-35; S1.46(b)', MAT, 'Entity-developed (target T2)', 'E4, E9', 'Yes', 'Yes'),
        (Env, 'Energy use and efficiency', 'Energy efficiency improvement', 'Change in energy intensity (MJ per garment) year on year', f"=Energy_GJ*1000/({prod}*1000)/7.21-1", '%', P, 'n/a', 'CO-01', 'S1.46(b), 50', MAT, 'Entity-developed', 'E8', 'Yes', 'No'),
        (Env, 'Water management', 'Total water usage', 'Water withdrawn, all sites', '=Water_m3', 'm3', N0, NA, 'SR-03', 'S1.46, 55(a)-(b)', MAT, 'CDSB water guidance / GRI 303-3', 'E8', 'Yes', 'Yes'),
        (Env, 'Water management', 'Water efficiency', 'Water withdrawn per thousand garments', f'=Water_m3/{prod}', 'm3 / \'000 pcs', N1, NA, 'SR-03', 'S1.46(b), 50', MAT, 'Entity-developed', 'E8', 'Yes', 'No'),
        (Env, 'Water management', 'Water recycling rate', 'Process water recycled - wash plant', '=' + e['e1_water_rec'], '%', P, NA, 'SR-03', 'S1.51 (target T7)', MAT, 'Entity-developed', 'E9', 'Yes', 'No'),
        (Env, 'Water management', 'Wastewater conformance', 'Samples within permit / ZDHC limits', '=' + e['e1_ww_conf'], '%', P, NA, 'SR-03', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-430a.1 (adapted)', 'E8', 'Yes', 'No'),
        (Env, 'Waste reduction and recycling', 'Total waste generated', 'All waste streams', '=' + waste, 'tonnes', N0, NA, 'SR-03', 'S1.46(b)', NO, 'GRI 306-3', 'E1', 'Yes', 'No'),
        (Env, 'Waste reduction and recycling', 'Recycling rate', 'Waste diverted from landfill / incineration', f"={e['e1_w_rec']}/{waste}", '%', P, NA, 'SR-03', 'S1.46(b)', NO, 'GRI 306-4', 'E8', 'Yes', 'No'),
        (Env, 'Waste reduction and recycling', 'Hazardous waste management', 'Hazardous waste treated by licensed contractor', '=' + e['e1_w_haz'], 'tonnes', N0, NA, 'SR-03', 'S1.46(b)', MAT, 'GRI 306-3', 'E1', 'Yes', 'No'),
        (Env, 'Pollution and resources', 'Air-emissions compliance', 'Boiler stack tests within permit limits', f"={e['e1_stack_pass']}/{e['e1_stack_tests']}", '%', P, NA, 'CR-04', 'S1.B25 - not disclosed', NO, 'Entity-developed (MoE permit)', 'E1', 'Yes', 'No'),
        (Env, 'Pollution and resources', 'Chemical management', 'Chemical inventory conforming to ZDHC MRSL', '=' + e['e1_mrsl'], '%', P, NA, 'SR-03', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-250a (adapted)', 'E1', 'Yes', 'No'),
        (Env, 'Pollution and resources', 'Plastic packaging', 'Polybags used', '=' + e['e1_plastic'], 'tonnes', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Env, 'Land use and biodiversity', 'Land area preserved', 'Hectares conserved or restored', '=' + e['e1_land_ha'], 'hectares', N0, NA, 'SR-05 (not relevant)', 'S1.B25 - not disclosed', NO, 'GRI 304-3', 'E1', 'Yes', 'No'),
        (Env, 'Land use and biodiversity', 'Operations in sensitive areas', 'Sites in or adjacent to protected areas', '=' + e['e1_sites_sens'], 'sites', N0, NA, 'SR-05 (not relevant)', 'S1.B25 - not disclosed', NO, 'GRI 304-1', 'E1', 'Yes', 'No'),
        (Env, 'Land use and biodiversity', 'Certified sourcing of biomass', 'Boiler biomass from certified plantation or agri-residue', '=' + e['e1_wood_cert'], '%', P, 0.27, 'CR-04', 'S2.33-35 (target T3)', MAT, 'Entity-developed', 'E9', 'Partial', 'No'),
        (Env, 'Supply chain - environmental', 'Supplier environmental assessment', 'Tier-1 spend with verified Higg FEM assessment', '=SC_Env', '%', P, NA, 'CR-03, SR-07', 'S1.46(b); S2.32', MAT, 'IBG / SASB CG-AA-430a.2', 'E11', 'Yes', 'No'),
        (Env, 'Supply chain - environmental', 'Supplier primary GHG data', 'Fabric spend from mills providing primary GHG data', '=SC_GHG', '%', P, NA, 'CR-03', 'S2.29(a)(vi), B56(a)', REQ, 'IFRS S2', 'E11, E5', 'Yes', 'No'),
        (Env, 'Supply chain - environmental', 'Supplier wastewater', 'Wet-processing suppliers meeting ZDHC limits', '=SC_WW', '%', P, NA, 'SR-07', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-430a.1', 'E11', 'Partial', 'No'),
        (Soc, 'Employee satisfaction and retention', 'Employee turnover rate', 'Leavers / average headcount', '=' + e['e1_turnover'], '%', P, NA, 'SR-04', 'S1.46(b)', NO, 'GRI 401-1', 'E8', 'Yes', 'No'),
        (Soc, 'Employee satisfaction and retention', 'Employee satisfaction', 'Engagement-survey score', '=' + e['e1_eng'], 'score / 100', N0, NA, 'SR-04', 'S1.46(b)', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Soc, 'Employee satisfaction and retention', 'Training and development', 'Average training hours per employee', f"={e['e1_train_h']}/{e['e1_hc']}", 'hours', N1, NA, 'SR-04', 'S1.46(b)', NO, 'GRI 404-1', 'E1', 'Yes', 'No'),
        (Soc, 'Diversity, equity and inclusion', 'Gender representation', 'Women in management (supervisor and above)', '=' + e['e1_women_mgmt'], '%', P, NA, 'SR-01', 'S1.46(b)', NO, 'GRI 405-1', 'E1', 'Yes', 'No'),
        (Soc, 'Diversity, equity and inclusion', 'Pay equity', 'Female-to-male basic pay ratio, same grade', '=' + e['e1_pay_eq'], 'ratio', N2, NA, 'SR-01', 'S1.46(b)', MAT, 'GRI 405-2', 'E1', 'Partial', 'No'),
        (Soc, 'Diversity, equity and inclusion', 'Workforce diversity', 'Female share of workforce', '=' + e['e1_female'], '%', P, NA, 'SR-01', 'S1.46(b)', MAT, 'GRI 405-1', 'E8', 'Yes', 'No'),
        (Soc, 'Community engagement', 'Community investment', 'Financial contributions to local communities', '=' + e['e1_comm'], 'USD\'000', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 201-1', 'E1', 'Yes', 'No'),
        (Soc, 'Community engagement', 'Local hiring', 'New hires from host provinces', '=' + e['e1_local'], '%', P, NA, 'SR-04', 'S1.B25 - not disclosed', NO, 'GRI 202-2 (adapted)', 'E1', 'Yes', 'No'),
        (Soc, 'Community engagement', 'Stakeholder engagement activities', 'Community consultations held', '=' + e['e1_consult'], 'number', N0, NA, 'CR-01, SR-03', 'S1.B25 - not disclosed', NO, 'GRI 413-1', 'E1', 'Yes', 'No'),
        (Soc, 'Human rights and community', 'Human-rights due diligence', 'Own operations covered by HRDD assessment', '=' + e['e1_hrdd'], '%', P, NA, 'SR-01', 'S1.46(b)', MAT, 'UNGP / OECD due-diligence guidance', 'E1', 'Yes', 'No'),
        (Soc, 'Human rights and community', 'Child and forced labour', 'Cases identified in own operations', '=' + e['e1_cl_fl'], 'number', N0, NA, 'SR-01', 'S1.46(b)', MAT, 'GRI 408 / 409', 'E1', 'Yes', 'No'),
        (Soc, 'Human rights and community', 'Community grievances', 'Grievances received from communities', '=' + e['e1_comm_griev'], 'number', N0, NA, 'CR-01, SR-03', 'S1.B25 - not disclosed', NO, 'GRI 413', 'E1', 'Yes', 'No'),
        (Soc, 'Health and safety', 'Injury rate', 'Recordable injuries per 200,000 hours worked', f"={e['e1_rec_inj']}*200000/{e['e1_hours']}", 'per 200k hours', N2, NA, 'SR-02', 'S1.46(b), 49', MAT, 'OSHA / GRI 403-9', 'E1', 'Yes', 'Yes'),
        (Soc, 'Health and safety', 'Lost-time injury frequency rate', 'Lost-time injuries per million hours worked', '=LTIFR', 'per million hours', N2, NA, 'SR-02', 'S1.46(b), 51 (target T6)', MAT, 'GRI 403-9', 'E8', 'Yes', 'Yes'),
        (Soc, 'Health and safety', 'Safety training completion', 'Employees completing required safety training', '=' + e['e1_safe_tr'], '%', P, NA, 'SR-02', 'S1.46(b)', MAT, 'GRI 403-5', 'E1', 'Yes', 'No'),
        (Soc, 'Health and safety', 'Occupational health monitoring', 'Occupational ill-health cases reported', '=' + e['e1_occ'], 'number', N0, NA, 'SR-02, CR-02', 'S1.46(b)', MAT, 'GRI 403-10', 'E1', 'Yes', 'No'),
        (Soc, 'Health and safety', 'Heat-related incidents', 'Persons treated for fainting on site', '=' + e['e1_faint'], 'number', N0, NA, 'CR-02, SR-02', 'S1.46(b); S2.B65', MAT, 'Entity-developed', 'E8', 'Yes', 'No'),
        (Soc, 'Employee well-being', 'Participation in wellness programmes', 'Employees using wellness initiatives', '=' + e['e1_well'], '%', P, NA, 'SR-04', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Partial', 'No'),
        (Soc, 'Employee well-being', 'Health-support access', 'Employees with access to on-site clinic', '=' + e['e1_clinic'], '%', P, NA, 'SR-02', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Soc, 'Labour conditions', 'Social-audit performance', 'Buyer audits passed without critical finding', '=' + e['e1_audit_pass'], '%', P, NA, 'SR-01', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-430b.2 (adapted)', 'E8', 'Yes', 'No'),
        (Soc, 'Labour conditions', 'BFC non-compliance points', 'Non-compliant critical assessment items', '=' + e['e1_bfc_nc'], 'number', N0, NA, 'SR-01', 'S1.46(b), 50(c)', MAT, 'Better Factories Cambodia', 'E8', 'Yes', 'No'),
        (Soc, 'Supply chain - social', 'Supplier social audits', 'Tier-1 spend covered by a social audit', '=SC_Soc', '%', P, NA, 'SR-07', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-430b.1', 'E11', 'Yes', 'No'),
        (Soc, 'Supply chain - social', 'Supplier corrective actions', 'Suppliers with critical findings that have an agreed plan', '=SC_CAP', '%', P, NA, 'SR-07', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-430b.2', 'E11', 'Yes', 'No'),
        (Soc, 'Supply chain - social', 'Supplier code of conduct', 'Tier-1 spend with code of conduct signed', '=SC_COC', '%', P, NA, 'SR-07', 'S1.46(b)', NO, 'Entity-developed', 'E11', 'Yes', 'No'),
        (Soc, 'Supply chain - social', 'High-risk supplier concentration', 'Tier-1 spend with high-risk suppliers', '=SC_HighSpend', '%', P, NA, 'SR-07', 'S1.32(b)', MAT, 'Entity-developed', 'E11', 'Partial', 'No'),
        (Soc, 'Customer responsibility', 'Restricted substances in products', 'Product RSL tests passed', f"={e['e1_rsl_pass']}/{e['e1_rsl_tests']}", '%', P, NA, 'SR-03', 'S1.55(a), 58(a)', MAT, 'SASB CG-AA-250a.1', 'E1', 'Yes', 'No'),
        (Soc, 'Customer responsibility', 'Product recalls', 'Recalls of products supplied', '=' + e['e1_recalls'], 'number', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Soc, 'Customer responsibility', 'Customer quality claims', 'Claims as % of revenue', '=' + e['e1_claims'], '%', P, NA, 'None', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Soc, 'Customer responsibility', 'Data privacy', 'Substantiated privacy complaints', '=' + e['e1_privacy'], 'number', N0, NA, 'SR-06 (not relevant)', 'S1.B25 - not disclosed', NO, 'GRI 418', 'E1', 'Yes', 'No'),
        (Gov, 'Board oversight and structure', 'Board independence', 'Independent directors / board members', f"={e['e1_bd_ind']}/{e['e1_bd_n']}", '%', P, NA, 'Governance', 'S1.27(a) narrative only', NO, 'GRI 2-9', 'B1', 'Yes', 'No'),
        (Gov, 'Board oversight and structure', 'Sustainability oversight', 'Board-level committee with sustainability mandate (1 = yes)', '=' + e['e1_bd_cttee'], 'yes = 1', N0, NA, 'Governance', 'S1.27(a); S2.6(a)', REQ, 'IFRS S1 / S2', 'B1', 'Yes', 'No'),
        (Gov, 'Board oversight and structure', 'Sustainability on board agenda', 'Board / committee meetings with sustainability items', '=' + e['e1_bd_meet'], 'meetings', N0, NA, 'Governance', 'S1.27(a)(iii); S2.6(a)(iii)', REQ, 'IFRS S1 / S2', 'B1', 'Yes', 'No'),
        (Gov, 'Board oversight and structure', 'Board diversity', 'Women on the board', f"={e['e1_bd_w']}/{e['e1_bd_n']}", '%', P, NA, 'Governance', 'S1.B25 - not disclosed', NO, 'GRI 405-1', 'B1', 'Yes', 'No'),
        (Gov, 'Executive accountability', 'Climate-linked compensation', 'Executive remuneration linked to climate considerations', '=Rem_Pct', '%', P, 0.0, 'CR-03', 'S2.29(g); S2.6(a)(v)', REQ, 'IFRS S2', 'E7', 'No', 'No'),
        (Gov, 'Executive accountability', 'CEO pay ratio', 'CEO remuneration as multiple of median employee', '=' + e['e1_ceo'], 'multiple', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 2-21', 'E1', 'Yes', 'No'),
        (Gov, 'Executive accountability', 'Performance-based incentives', 'Bonus pool linked to strategic goals', '=' + e['e1_bonus'], '%', P, NA, 'None', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Partial', 'No'),
        (Gov, 'Ethics and compliance', 'Code-of-conduct training', 'Employees completing annual ethics training', '=' + e['e1_coc'], '%', P, NA, 'SR-01', 'S1.46(b)', NO, 'GRI 205-2', 'E1', 'Yes', 'No'),
        (Gov, 'Ethics and compliance', 'Anti-corruption cases', 'Confirmed corruption incidents', '=' + e['e1_corr'], 'number', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 205-3', 'E1', 'Yes', 'No'),
        (Gov, 'Ethics and compliance', 'Whistleblower system use', 'Reports resolved / reports received', f"={e['e1_wb_res']}/{e['e1_wb_rec']}", '%', P, NA, 'SR-01', 'S1.46(b)', MAT, 'GRI 2-26', 'E1', 'Yes', 'No'),
        (Gov, 'Tax transparency', 'Effective tax rate', 'Income tax expense / profit before tax', f"={e['e1_tax_exp']}/PBT25", '%', P, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 207 / FS note 8', 'E1', 'Yes', 'No'),
        (Gov, 'Tax transparency', 'Total tax contribution', 'Income tax, payroll taxes, customs and withholding paid', '=' + e['e1_tax_total'], 'USD\'000', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 207-4', 'E1', 'Yes', 'No'),
        (Gov, 'Tax transparency', 'Tax strategy published', 'Board-approved tax strategy published (1 = yes)', '=' + e['e1_tax_policy'], 'yes = 1', N0, NA, 'None', 'S1.B25 - not disclosed', NO, 'GRI 207-1', 'E1', 'Yes', 'No'),
        (Gov, 'Risk management', 'Business-continuity plans', 'Critical sites with a plan updated in the year', f"={e['e1_bcp']}/3", '%', P, 0.33, 'CR-01', 'S2.25(a); S2.14(a)(ii)', MAT, 'Entity-developed', 'D1', 'Yes', 'No'),
        (Gov, 'Risk management', 'Risk-mitigation coverage', 'Relevant SROs with a documented response (C1)', f"={e['e1_sro_resp']}/SRO_Relevant", '%', P, 0.78, 'All', 'S1.44; S2.25', MAT, 'Entity-developed', 'A3, C1', 'Yes', 'No'),
        (Gov, 'Risk management', 'Cyber-security audits', 'Audits conducted in the year', '=' + e['e1_cyber'], 'number', N0, NA, 'SR-06 (not relevant)', 'S1.B25 - not disclosed', NO, 'Entity-developed', 'E1', 'Yes', 'No'),
        (Gov, 'Transparency and reporting', 'Disclosure frequency', 'Sustainability-related financial disclosures issued with the annual report', '=' + e['e1_reports'], 'reports', N0, 1, 'Governance', 'S1.64, 72', REQ, 'IFRS S1', 'A1', 'Yes', 'No'),
        (Gov, 'Transparency and reporting', 'Disclosure readiness', 'ISSB checklist items complete (F1)', '=Chk_Pct', '%', P, 'n/a', 'Governance', 'S1.72 (internal KPI)', NO, 'Entity-developed', 'F1', 'Yes', 'No'),
        (Gov, 'Transparency and reporting', 'ESG rating performance', 'Third-party rating score', '=' + e['e1_rating'], 'score / 100', N0, NA, 'CO-02', 'S1.B25 - not disclosed', NO, 'Third-party rating', 'E1', 'Yes', 'No'),
        (Gov, 'Transparency and reporting', 'Grievance resolution', 'Average days to resolve worker grievances', '=' + e['e1_griev'], 'days', N0, NA, 'SR-01', 'S1.46(b)', MAT, 'Entity-developed', 'E1', 'Yes', 'No'),
    ]
    H = ['KPI ID', 'Pillar', 'Category', 'KPI', 'Metric definition', 'FY2025', 'Unit', 'FY2024', 'Linked SRO (A3)', 'ISSB basis', 'In ISSB disclosures?',
         'Metric source (S1.49-50)', 'Working paper', 'Evidence on file? (F4)', 'In assurance scope? (G1)', 'ESG theme (E12)']
    THEME = {'Carbon emissions and climate': 'Climate change', 'Energy use and efficiency': 'Climate change', 'Water management': 'Water security',
             'Waste reduction and recycling': 'Pollution & resources', 'Pollution and resources': 'Pollution & resources', 'Land use and biodiversity': 'Biodiversity',
             'Supply chain - environmental': 'Supply chain - environmental', 'Supply chain - social': 'Supply chain - social',
             'Employee satisfaction and retention': 'Labour standards', 'Diversity, equity and inclusion': 'Labour standards', 'Labour conditions': 'Labour standards',
             'Human rights and community': 'Human rights & community', 'Community engagement': 'Human rights & community',
             'Health and safety': 'Health & safety', 'Employee well-being': 'Health & safety', 'Customer responsibility': 'Customer responsibility',
             'Board oversight and structure': 'Corporate governance', 'Executive accountability': 'Corporate governance', 'Transparency and reporting': 'Corporate governance',
             'Ethics and compliance': 'Anti-corruption', 'Risk management': 'Risk management', 'Tax transparency': 'Tax transparency',
             'Certified sourcing of biomass': 'Supply chain - environmental', 'Code-of-conduct training': 'Anti-corruption'}
    rows = []
    for i, k in enumerate(K):
        rows.append([f'{k[0][0]}-{i + 1:02d}', k[0], k[1], k[2], k[3], k[4], k[5], k[7], k[8], k[9], k[10], k[11], k[12], k[13], k[14], THEME.get(k[2], THEME[k[1]])])
    a, b = X.table(ws, r, H, rows, inp=(4, 6, 8, 9, 10, 11, 12, 13, 14, 15))
    for i, k in enumerate(K):
        X.put(ws, a + i, 5, k[4], num_format=k[6])
        v = k[7]
        if isinstance(v, str) and not v.startswith('='):
            X.put(ws, a + i, 7, v, font_color=C['inp'], align='right')
        else:
            X.put(ws, a + i, 7, v, num_format=k[6])
    X.dv(ws, a, b, 10, [REQ, MAT, NO]); X.dv(ws, a, b, 13, ['Yes', 'Partial', 'No']); X.dv(ws, a, b, 14, ['Yes', 'No'])
    X.cf_status(ws, a, b, 13); X.cf_status(ws, a, b, 14)
    for word, bg, fg in ((REQ, '#C6EFCE', '#006100'), (MAT, '#DDEBF7', '#1F4E78'), (NO, '#EDEDED', '#595959')):
        ws.conditional_format(a, 10, b, 10, {'type': 'cell', 'criteria': '==', 'value': f'"{word}"', 'format': X.wb.add_format({'bg_color': bg, 'font_color': fg})})
    ws.autofilter(r, 0, b, 15)
    X.ad['kpi'] = (a, b)
    r = b + 2
    X.sec(ws, r, 'Summary (feeds Dashboard)', 9)
    X.hdr(ws, r + 1, ['', 'Pillar', 'KPIs in library', 'Required by ISSB', 'Material SRO metric', 'Monitored only', 'In ISSB disclosures', 'Evidence on file', 'In assurance scope'])
    pil, inc, ev, asr = (f'{A(a, c)}:{A(b, c)}' for c in (1, 10, 13, 14))
    for i, p in enumerate([Env, Soc, Gov, 'TOTAL']):
        R = r + 2 + i
        X.put(ws, R, 0, ''); X.put(ws, R, 1, p, bold=True)
        if p == 'TOTAL':
            for j in range(2, 9):
                X.put(ws, R, j, f'=SUM({A(r + 2, j)}:{A(R - 1, j)})', bold=True)
        else:
            X.put(ws, R, 2, f'=COUNTIF({pil},"{p}")')
            X.put(ws, R, 3, f'=COUNTIFS({pil},"{p}",{inc},"{REQ}")')
            X.put(ws, R, 4, f'=COUNTIFS({pil},"{p}",{inc},"{MAT}")')
            X.put(ws, R, 5, f'=COUNTIFS({pil},"{p}",{inc},"{NO}")')
            X.put(ws, R, 6, f'={A(R, 3)}+{A(R, 4)}')
            X.put(ws, R, 7, f'=COUNTIFS({pil},"{p}",{ev},"Yes")')
            X.put(ws, R, 8, f'=COUNTIFS({pil},"{p}",{asr},"Yes")')
    X.ad['kpisum'] = (r + 2, r + 4)
    tr = r + 5
    for nm, c in (('KPI_Total', 2), ('KPI_ISSB', 6), ('KPI_Ev', 7), ('KPI_Assured', 8)):
        X.name(nm, E10, tr, c)
    X.note(ws, tr + 2, 'How to read "In ISSB disclosures?": "Yes - required" = a specific IFRS S1 / S2 requirement. "Yes - material SRO" = the entity uses the KPI to measure and monitor a relevant risk or opportunity and the information is material (S1.46(b)); '
                       'the source is identified (S1.49) and entity-developed metrics are defined (S1.50). "No - monitored only" = not linked to a relevant SRO or not material: ISSB does not require it and including it could obscure material information (S1.B25, B30); '
                       'it can still be reported under other frameworks (e.g. GRI). No FY2024 comparative is given for non-climate KPIs in the first year they are reported (S1.E6).', 16, h=54)
    ws.freeze_panes(9, 4)


def build_g1(X):
    ws, r = X.sheet(G1, 'G1', 'ISSA 5000 assurance readiness - preconditions, assertions, materiality and differences',
                    'Assess whether the sustainability information is ready for an external assurance engagement under ISSA 5000. ISSA 5000 is applied by the assurance practitioner, not by the entity: the entity\'s task is to be "assurance-ready" - suitable criteria, a documented reporting process, internal control, and sufficient appropriate evidence for each disclosure in scope.',
                    'IFRS S1.72 (compliance), S1.B8-B10, S1.D21-D24 (verifiability), S1.27(b)(ii), S2.6(b)(ii), S2.B40(d). Assurance itself is outside IFRS S1 / S2',
                    'IAASB ISSA 5000 General Requirements for Sustainability Assurance Engagements (referenced by topic - ADD PARAGRAPH REFERENCES FROM THE STANDARD; it was not among the source documents). IESBA sustainability ethics standards (IESSA)',
                    [8, 26, 44, 44, 34, 20, 13, 30], prep='Ly S.', rev='Internal Audit / ARSC', tab='#7F3F98', status='In progress', obj_h=44)
    X.sec(ws, r, '1. Proposed assurance engagement parameters (to be agreed with the practitioner)', 8)
    P = [('Assurance standard', 'ISSA 5000 (IAASB). Check the effective date and local adoption with the regulator and the practitioner before engagement', 'Engagement basis'),
         ('Level of assurance', 'Limited assurance in FY2025 on the scope below; move to reasonable assurance on Scope 1 and 2 when controls mature', 'Limited / reasonable assurance'),
         ('Sustainability information in scope', 'Scope 1 and Scope 2 GHG emissions, energy, water, LTIFR and injury rate, target T1 progress (see E10 column "In assurance scope")', 'Scope and boundary of the engagement'),
         ('Reporting boundary', 'Same reporting entity as the financial statements plus the operationally controlled JV for GHG (A1 s.2)', 'Reporting boundary'),
         ('Applicable criteria', 'IFRS S1 and IFRS S2; GHG Protocol Corporate Standard (2004); entity-developed metric definitions in E8 / E10 (made available to users in the report)', 'Suitable and available criteria'),
         ('Not in scope this year', 'Scope 3 emissions (secondary factors), scenario analysis and anticipated financial effects, narrative governance / strategy disclosures', 'Scope limitation to be disclosed clearly'),
         ('Practitioner independence and ethics', 'Practitioner to confirm independence; non-assurance services (GHG advisory) to be pre-approved by ARSC', 'Relevant ethical requirements')]
    for i, (l, v, tpc) in enumerate(P):
        X.put(ws, r + 1 + i, 0, f'P-{i + 1}')
        X.put(ws, r + 1 + i, 1, l, bold=True)
        ws.merge_range(r + 1 + i, 2, r + 1 + i, 5, v, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.merge_range(r + 1 + i, 6, r + 1 + i, 7, 'ISSA 5000 topic: ' + tpc, X.f(text_wrap=True, border=1, border_color=C['mid']))
        ws.set_row(r + 1 + i, 30)
    r += len(P) + 2
    X.sec(ws, r, '2. Readiness checklist', 8)
    H = ['ID', 'ISSA 5000 topic', 'What the practitioner will need', 'Entity position', 'Evidence / working paper', 'Owner', 'Status', 'Gap / action']
    Cc, Ip, Ns = 'Complete', 'In progress', 'Not started'
    R = [('Preconditions', 'Roles and responsibilities of management and those charged with governance are acknowledged', 'Board statement of responsibility drafted; ARSC oversees reporting', 'B1; draft annual report', 'Company Secretary', Cc, ''),
         ('Preconditions', 'Criteria are suitable (relevant, complete, reliable, neutral, understandable) and available to intended users', 'ISSB Standards are public; entity-developed metric definitions documented', 'A1 s.3; E8; E10', 'Head of Sustainability', Ip, 'Adapted SASB metrics need published definitions (F5 A-017)'),
         ('Preconditions', 'The sustainability matters and reporting boundary are appropriate and clearly identified', 'Boundary documented incl. JV under operational control', 'A1 s.2; E6', 'Financial Controller', Cc, ''),
         ('Preconditions', 'The practitioner expects to be able to obtain the evidence needed; engagement has a rational purpose', 'Evidence tracker in place; scope limited to assurance-ready streams', 'F4', 'Internal Audit', Ip, 'See F4 exceptions'),
         ('Entity process', 'Process to identify the sustainability information to be reported (materiality process) is documented', 'Six-step materiality process approved by ARSC', 'A2, A3, A4', 'Risk Manager', Cc, ''),
         ('Entity process', 'Understanding of the entity, its value chain and its legal / regulatory environment', 'Value-chain map and regulatory scan', 'A2; A1 s.5', 'Head of Sustainability', Ip, 'Jurisdictional status to confirm (F5 A-001)'),
         ('Internal control', 'Control environment and oversight of sustainability reporting', 'ARSC mandate; Sustainability Steering Committee', 'B1', 'Company Secretary', Cc, ''),
         ('Internal control', 'Risk-assessment process covers risks of misstatement in sustainability information, incl. fraud / greenwashing', 'ERM covers SROs; reporting-misstatement risk not yet assessed', 'D1', 'Risk Manager', Ns, 'Add reporting-risk assessment incl. incentives linked to targets (bonus KPI)'),
         ('Internal control', 'Information system: data flows from source to report are documented for each metric in scope', 'Controls manual v1.1; data-flow maps for energy only', 'F4; controls manual', 'Financial Controller', Ip, 'Extend data-flow maps to water, OHS, fuels'),
         ('Internal control', 'Control activities: preparer / reviewer segregation, reconciliations, IT access to spreadsheets', 'Monthly reconciliations; spreadsheet access not restricted', 'F4', 'Financial Controller', Ip, 'IA findings #2 and #4 open (F5 A-008)'),
         ('Internal control', 'Monitoring: internal audit coverage of sustainability data', 'Internal audit walkthrough of GHG process Q4-2025', 'IA-2025-09', 'Internal Audit', Cc, ''),
         ('Evidence', 'Audit trail from each reported amount to source documents; calculations re-performable', 'Linked workbook (E1-E6) with source references', 'E1-E6', 'Head of Sustainability', Cc, ''),
         ('Evidence', 'Emission factors, GWPs and conversion factors supported by cited sources', 'Fuel factors documented; grid and Scope 3 factors are placeholders', 'E2', 'Head of Sustainability', Ip, 'Replace placeholders (F5 A-002)'),
         ('Evidence', 'Estimates and forward-looking information: methods, assumptions and data documented', 'Scenario model exists; assumption sources incomplete', 'C3; F2', 'CFO', Ip, 'Assumptions book (F5 A-010)'),
         ('Evidence', 'Information from the value chain and from experts (adviser, laboratory, engineer) - competence and source documented', 'Lab accreditation and engineer certificate held; mill data not obtained', 'F4', 'Procurement', Ip, 'Mill data request (F5 A-003)'),
         ('Evidence', 'Comparatives, restatements and errors are documented', 'Restatement R-01 documented and disclosed', 'F2 s.3; E6', 'Financial Controller', Cc, ''),
         ('Reporting', 'Consistency of sustainability information with the financial statements and other information in the annual report', 'Tie-out performed; one difference and two connectivity points open', 'F3', 'Financial Controller', Ip, 'F5 A-011 to A-013'),
         ('Reporting', 'Disclosures are complete against the applicable criteria', 'Checklist tracked', 'F1', 'Financial Controller', Ip, '=TEXT(Chk_Pct,"0%")&" of applicable checklist items complete"'),
         ('Reporting', 'Events after the reporting period considered up to the date of the assurance report', 'Review planned', 'A1', 'Financial Controller', Ns, 'F5 A-018'),
         ('Engagement', 'Written representations from management and those charged with governance', 'Template to be agreed with practitioner', '-', 'CFO', Ns, 'Draft representation letter'),
         ('Engagement', 'Practitioner appointed, terms of engagement agreed, timetable aligned with financial-statement audit', 'Proposals requested from two firms', 'ARSC paper Feb-2026', 'CFO', Ip, 'Appointment at ARSC 26-Feb-2026'),
         ('Engagement', 'Access to people, sites and records, incl. the JV and outsourced data providers', 'JV agreement grants access', 'JV agreement cl. 14', 'CFO', Cc, '')]
    rows = [[f'R-{i + 1:02d}'] + list(x) for i, x in enumerate(R)]
    a, b = X.table(ws, r + 1, H, rows, inp=(1, 3, 4, 5, 6, 7))
    X.dv(ws, a, b, 6, ['Complete', 'In progress', 'Not started', 'N/A'])
    X.cf_status(ws, a, b, 6)
    st = f'{A(a, 6)}:{A(b, 6)}'
    r = b + 1
    Sm = [('Readiness items', f'=COUNTA({A(a, 0)}:{A(b, 0)})', None, '0'), ('Complete', f'=COUNTIF({st},"Complete")', None, '0'),
          ('In progress', f'=COUNTIF({st},"In progress")', None, '0'), ('Not started', f'=COUNTIF({st},"Not started")', None, '0'),
          ('ISSA 5000 readiness (%)', f'={A(r + 1, 3)}/{A(r, 3)}', 'ISSA_Ready', '0%')]
    for i, (l, f, nm, nf) in enumerate(Sm):
        ws.merge_range(r + i, 0, r + i, 2, l, X.f(bold=True, border=1, border_color=C['mid']))
        X.put(ws, r + i, 3, f, bold=True, num_format=nf, align='left')
        if nm:
            X.name(nm, G1, r + i, 3)
    r += len(Sm) + 1
    X.sec(ws, r, '3. Disclosures in scope - assertions, risk of material misstatement and planned evidence', 8)
    H = ['ID', 'Disclosure', 'Relevant assertions', 'What could go wrong (risk)', 'Key control', 'Inherent risk', 'Control in place?', 'Evidence the practitioner is likely to request']
    Rr = [('Scope 1 GHG emissions', 'Completeness; accuracy; cut-off; classification', 'Sources omitted (as for LPG in FY2024); wrong factor or unit; fuel purchased vs consumed', 'Source inventory reviewed annually; fuel stock reconciliation', 'High', 'Yes', 'Source register, fuel ledgers, invoices, factor sources, recalculation'),
          ('Scope 2 GHG emissions', 'Completeness; accuracy; cut-off', 'Missing meter or month; outdated grid factor', 'Invoice-to-meter reconciliation', 'Medium', 'Yes', 'EDC invoices (12 x 3 sites), grid-factor source'),
          ('Organisational boundary (JV)', 'Completeness; presentation', 'JV wrongly excluded or shown in the wrong column', 'Boundary memo approved by CFO', 'Medium', 'Yes', 'JV agreement, boundary memo'),
          ('Energy consumption and renewable share', 'Accuracy; classification', 'Solar generation over-stated; conversion factors', 'Inverter data reviewed against meter', 'Medium', 'Yes', 'Inverter logs, bi-directional meter data'),
          ('Water withdrawal', 'Completeness; accuracy', 'Unmetered boreholes; meter drift', 'Meter calibration; monthly log sign-off', 'Medium', 'Partial', 'Meter logs, calibration certificates'),
          ('LTIFR and injury rate', 'Completeness; occurrence; classification', 'Under-reporting of injuries; hours worked mis-stated', 'Incident classification review; payroll hours', 'High', 'Yes', 'Incident register, clinic log, payroll hours, interviews'),
          ('Target T1 progress and base year', 'Accuracy; consistency; presentation', 'Base year not recalculated after error; selective presentation', 'Annual base-year recalculation check', 'High', 'Partial', 'Base-year inventory, recalculation policy'),
          ('Climate-related capital expenditure', 'Occurrence; classification; consistency with FS', 'Non-climate capex classified as climate', 'Project-code approval by CFO', 'Medium', 'Yes', 'Capex register, board papers, FS note 9'),
          ('Assets exposed to physical risk', 'Accuracy; classification', 'Flood-extent definition applied inconsistently', 'Definition approved; asset register by location', 'Medium', 'Partial', 'Flood study, asset register, inventory by site')]
    rows = [[f'D-{i + 1}'] + list(x) for i, x in enumerate(Rr)]
    a, b = X.table(ws, r + 1, H, rows, inp=tuple(range(1, 8)))
    X.dv(ws, a, b, 5, ['High', 'Medium', 'Low']); X.dv(ws, a, b, 6, ['Yes', 'Partial', 'No'])
    X.cf_status(ws, a, b, 5); X.cf_status(ws, a, b, 6)
    r = b + 2
    X.sec(ws, r, '4. Indicative quantitative materiality and schedule of identified differences (management\'s own pre-assurance review)', 8)
    H = ['ID', 'Metric', 'Reported amount', 'Indicative materiality %', 'Materiality (amount)', 'Performance materiality (75%)', 'Identified differences (absolute)', 'Conclusion / nature of difference']
    Mm = [('Scope 1 GHG (tCO2e)', '=GHG_S1', 0.05, 12, 'Wood recorded by volume at Factory B - density assumption (F5 A-005)'),
          ('Scope 2 GHG (tCO2e)', '=GHG_S2', 0.05, 0, 'None identified'),
          ('Total energy (GJ)', '=Energy_GJ', 0.05, 1400, 'Same wood-density point'),
          ('Water withdrawn (m3)', '=Water_m3', 0.05, 0, 'None identified'),
          ('Climate capex (USD\'000)', '=Capex_Climate', 0.05, 0, 'None identified'),
          ('Assets exposed to physical risk (USD\'000)', '=Assets_Physical', 0.05, 150, 'F3 K-12 goods-in-transit allocation')]
    rows = [[f'M-{i + 1}', m[0], m[1], m[2], lambda R: f'=C{R}*D{R}', lambda R: f'=E{R}*0.75', m[3],
             lambda R, t=m[4]: f'=IF(G{R}>E{R},"EXCEEDS materiality - ","Below materiality - ")&"{t}"'] for i, m in enumerate(Mm)]
    a, b = X.table(ws, r + 1, H, rows, nf={3: '0.0%'}, inp=(1,), extra={3: {'bg_color': C['yellow']}})
    X.note(ws, b + 2, 'Materiality percentages are management\'s illustrative assumptions for self-review only. Under ISSA 5000 the practitioner sets materiality for quantitative disclosures and considers it for qualitative disclosures, '
                      'from the perspective of intended users; qualitative factors (e.g. safety, target achievement, remuneration triggers) can make a small difference material.', 8, h=40)
    ws.freeze_panes(9, 2)
