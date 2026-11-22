"""Part 2: E-series - activity data, emission factors, GHG inventory, metrics and targets."""
from wp_lib import C, A, ref, rng

E1, E2, E3, E4, E5, E6, E7, E8, E9 = ('E1 ActivityData', 'E2 EF-GWP', 'E3 Scope1', 'E4 Scope2', 'E5 Scope3',
                                      'E6 GHG Summary', 'E7 CrossIndustry', 'E8 Industry', 'E9 Targets')
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def build_e1(X):
    ws, r = X.sheet(E1, 'E1', 'Activity data - energy, production, water and other source data',
                    'Central register of primary activity data extracted from source documents (utility bills, fuel ledgers, production and HR systems). All GHG and metric calculations link to this sheet - enter data here only.',
                    'IFRS S2.29(a), B19 (reporting period), B26-B29 (inputs); IFRS S1.46, S1.B42(c) (consistent data)',
                    'HB 3.5.10, 4.1.140; DC 2.4.1; KPMG GHG emissions reporting handbook (referred to in HB)',
                    [12] + [11] * 15, prep='Chan M.', tab=C['green'])
    assert r == 8
    s = [0.96, 0.86, 1.00, 1.04, 1.08, 1.07, 1.05, 1.03, 0.99, 0.93, 1.00, 0.99]
    o = [1.0, 1.0, 1.2, 1.6, 1.7, 1.3, 1.0, 0.9, 0.9, 0.9, 0.8, 0.8]
    sol = [0, 0, 0, 118, 132, 128, 122, 118, 112, 108, 124, 128]
    H = ['Month FY2025', 'Grid electricity - Factory A (MWh)', 'Grid electricity - Factory B (MWh)', 'Grid electricity - JV wash plant (MWh)', 'Solar self-generated - Factory A (MWh)',
         'Diesel generators - A (litres)', 'Diesel generators - B (litres)', 'Diesel generators - JV (litres)', 'Wood biomass - A (tonnes)', 'Wood biomass - B (tonnes)', 'Wood biomass - JV (tonnes)',
         'Production - A (\'000 pcs)', 'Production - B (\'000 pcs)', 'Water withdrawn - A (m3)', 'Water withdrawn - B (m3)', 'Water withdrawn - JV (m3)']
    rows = []
    for i, m in enumerate(MONTHS):
        fb = 0.70 if m == 'Oct' else 1.0      # October flood at Factory B
        rows.append([m, round(480 * s[i] - 0.95 * sol[i]), round(325 * s[i] * fb), round(217 * s[i]), sol[i],
                     int(round(5200 * o[i], -1)), int(round(4600 * o[i] * (2.2 if m == 'Oct' else 1), -1)), int(round(3900 * o[i], -1)),
                     round(190 * s[i]), round(160 * s[i] * fb), round(250 * s[i]),
                     round(1210 * s[i]), round(850 * s[i] * fb), int(round(9500 * s[i], -1)), int(round(7200 * s[i] * fb, -1)), int(round(21000 * s[i], -1))])
    a, b = X.table(ws, r, H, rows)
    ws.set_row(r, 54)
    X.put(ws, b + 1, 0, 'TOTAL FY2025', bold=True)
    for j in range(1, 16):
        X.put(ws, b + 1, j, f'=SUM({A(a, j, False)}:{A(b, j, False)})', bold=True, bg_color=C['grey'])
    X.ad['e1m'] = (a, b)             # month rows (0-based)
    X.ad['e1t'] = b + 1              # total row
    assert b + 1 == 21
    X.put(ws, b + 2, 0, 'Source', italic=True)
    src = ['EDC invoices'] * 3 + ['Inverter log'] + ['Fuel issue ledger'] * 3 + ['Weighbridge / delivery notes'] * 3 + ['Production system'] * 2 + ['Water meter log'] * 3
    for j, v in enumerate(src):
        X.put(ws, b + 2, j + 1, v, italic=True)
    X.note(ws, b + 3, 'October: flood at Factory B (9 production days lost; additional generator running). April: rooftop solar commissioned at Factory A. Sample data for illustration - replace with client records.', 16)
    r = b + 5
    X.sec(ws, r, '2. Annual and other activity data', 7)
    H = ['Ref', 'Data item', 'FY2025 value', 'Unit', 'Source document', 'Owner', 'Used in']
    D_ = [
        ('veh_dsl', 'Diesel - company vehicles and trucks', 95000, 'litres', 'Fuel-card statements', 'Logistics', 'E3'),
        ('veh_gas', 'Gasoline - company vehicles', 22000, 'litres', 'Fuel-card statements', 'Logistics', 'E3'),
        ('lpg', 'LPG - canteens (A and B)', 38000, 'kg', 'Supplier invoices', 'Admin', 'E3'),
        ('r410a', 'Refrigerant top-up R-410A - consolidated group', 120, 'kg', 'HVAC service records', 'Maintenance', 'E3'),
        ('r134a', 'Refrigerant top-up R-134a - JV chillers', 20, 'kg', 'HVAC service records', 'JV Maintenance', 'E3'),
        ('fabric', 'Fabric purchased', 9200, 'tonnes', 'Purchase ledger / packing lists', 'Procurement', 'E5 cat 1, E8'),
        ('trims', 'Trims, accessories and packaging purchased', 1150, 'tonnes', 'Purchase ledger', 'Procurement', 'E5 cat 1'),
        ('ogs', 'Other purchased goods and services (spend)', 6500, 'USD\'000', 'General ledger', 'Finance', 'E5 cat 1'),
        ('km_sea_in', 'Average inbound sea distance (mills to Sihanoukville)', 3500, 'km', 'Forwarder data', 'Logistics', 'E5 cat 4'),
        ('km_road', 'Average road distance port to factories', 230, 'km', 'Route data', 'Logistics', 'E5 cat 4, 9'),
        ('fg', 'Finished goods shipped', 8300, 'tonnes', 'Shipping documents', 'Logistics', 'E5 cat 9, 12'),
        ('km_sea_out', 'Average outbound sea distance to destination ports', 15500, 'km', 'Buyer logistics data', 'Logistics', 'E5 cat 9'),
        ('w_rec', 'Waste - fabric off-cuts sent for recycling', 780, 'tonnes', 'Waste-contractor manifests', 'EHS', 'E5 cat 5, E8'),
        ('w_land', 'Waste - general waste to landfill', 410, 'tonnes', 'Waste-contractor manifests', 'EHS', 'E5 cat 5, E8'),
        ('w_haz', 'Waste - hazardous (sludge, chemical containers) incinerated', 36, 'tonnes', 'Licensed contractor manifests', 'EHS', 'E5 cat 5, E8'),
        ('air_pkm', 'Business air travel', 420000, 'passenger-km', 'Travel-agent report', 'Admin', 'E5 cat 6'),
        ('hc', 'Average number of employees', 6480, 'headcount', 'HR system', 'HR', 'E5 cat 7, E8'),
        ('days', 'Working days', 290, 'days', 'HR calendar', 'HR', 'E5 cat 7'),
        ('km_commute', 'Average one-way commuting distance', 9, 'km', 'Commuting survey Aug-2025 (n=1,240)', 'HR', 'E5 cat 7'),
        ('sh_truck', 'Share of workers commuting by collective truck / bus', 0.65, '%', 'Commuting survey', 'HR', 'E5 cat 7'),
        ('hours', 'Hours worked', 16300000, 'hours', 'Payroll / time records', 'HR', 'E8 LTIFR'),
        ('lti', 'Lost-time injuries', 28, 'number', 'OHS incident register', 'EHS', 'E8'),
        ('fatal', 'Work-related fatalities', 0, 'number', 'OHS incident register', 'EHS', 'E8'),
        ('faint', 'Fainting incidents (persons)', 37, 'number', 'Clinic log', 'EHS', 'E8'),
        ('female', 'Female share of workforce', 0.84, '%', 'HR system', 'HR', 'E8'),
        ('turnover', 'Annual employee turnover', 0.28, '%', 'HR system', 'HR', 'E8'),
        ('wood_cert', 'Share of boiler biomass from certified plantation or agri-residue', 0.35, '%', 'Supplier declarations / chain-of-custody', 'Procurement', 'E8, E9'),
        ('flood_prot', 'Share of critical assets at Factory B protected to 1-in-100 flood level', 0.60, '%', 'Engineer\'s certificate Sep-2025', 'Engineering', 'E9'),
        ('water_rec', 'Share of wash-plant process water recycled', 0.12, '%', 'ETP flow meters', 'JV EHS', 'E8, E9'),
        ('ww_conf', 'Wastewater samples conforming to permit / ZDHC limits', 0.958, '%', 'Lab test reports (24 samples)', 'JV EHS', 'E8'),
        ('rpet', 'Recycled polyester share of polyester fabric purchased', 0.17, '%', 'GRS transaction certificates', 'Procurement', 'E8'),
        ('cotton_cert', 'Share of cotton fabric certified (BCI / organic)', 0.41, '%', 'Supplier certificates', 'Procurement', 'E8'),
        ('bfc_nc', 'BFC assessment - non-compliance points (of 52 critical items)', 3, 'number', 'BFC report Nov-2025', 'Compliance', 'E8'),
        ('audit_pass', 'Buyer social audits passed without critical finding', 0.92, '%', '12 buyer audits', 'Compliance', 'E8'),
        ('tier1', 'Tier-1 suppliers (fabric mills, trims)', 46, 'number', 'Approved supplier list', 'Procurement', 'E8'),
        ('tier2', 'Suppliers beyond tier 1 (mapped)', 61, 'number', 'Traceability system', 'Procurement', 'E8'),
        ('land_ha', 'Land conserved or restored', 0, 'hectares', 'EHS records', 'EHS', 'E10'),
        ('sites_sens', 'Sites in or adjacent to protected areas', 0, 'sites', 'Site screening 2025', 'EHS', 'E10'),
        ('eng', 'Employee engagement-survey score', 71, 'score / 100', 'Survey Sep-2025', 'HR', 'E10'),
        ('train_h', 'Training hours delivered', 142600, 'hours', 'Training register', 'HR', 'E10'),
        ('women_mgmt', 'Women in management (supervisor and above)', 0.46, '%', 'HR system', 'HR', 'E10'),
        ('pay_eq', 'Female-to-male basic pay ratio (same grade)', 0.97, 'ratio', 'Payroll analysis', 'HR', 'E10'),
        ('comm', 'Community investment', 185, 'USD\'000', 'General ledger', 'Finance', 'E10'),
        ('local', 'New hires from host provinces', 0.88, '%', 'HR system', 'HR', 'E10'),
        ('consult', 'Community consultations held', 14, 'number', 'Stakeholder log', 'Compliance', 'E10'),
        ('rec_inj', 'Recordable injuries (incl. lost-time)', 64, 'number', 'OHS incident register', 'EHS', 'E10'),
        ('safe_tr', 'Employees completing required safety training', 0.96, '%', 'Training register', 'EHS', 'E10'),
        ('occ', 'Occupational ill-health cases', 6, 'number', 'Clinic log', 'EHS', 'E10'),
        ('well', 'Employees using wellness initiatives', 0.41, '%', 'HR records', 'HR', 'E10'),
        ('clinic', 'Employees with access to on-site clinic', 1.0, '%', 'HR records', 'HR', 'E10'),
        ('bd_n', 'Board members', 7, 'number', 'Company register', 'Company Secretary', 'E10'),
        ('bd_ind', 'Independent directors', 3, 'number', 'Company register', 'Company Secretary', 'E10'),
        ('bd_w', 'Women on the board', 2, 'number', 'Company register', 'Company Secretary', 'E10'),
        ('bd_cttee', 'Board committee with sustainability mandate (1 = yes)', 1, 'yes = 1', 'ARSC terms of reference', 'Company Secretary', 'E10'),
        ('bd_meet', 'Board / committee meetings with sustainability items', 8, 'meetings', 'Minutes (B1 log)', 'Company Secretary', 'E10'),
        ('ceo', 'CEO remuneration as multiple of median employee', 86, 'multiple', 'Payroll analysis', 'HR', 'E10'),
        ('bonus', 'Bonus pool linked to strategic goals', 0.60, '%', 'Remuneration policy', 'HR', 'E10'),
        ('coc', 'Employees completing annual ethics training', 0.93, '%', 'Training register', 'Compliance', 'E10'),
        ('corr', 'Confirmed corruption incidents', 0, 'number', 'Compliance register', 'Compliance', 'E10'),
        ('wb_rec', 'Whistleblower / grievance reports received', 23, 'number', 'Grievance system', 'Compliance', 'E10'),
        ('wb_res', 'Whistleblower / grievance reports resolved', 21, 'number', 'Grievance system', 'Compliance', 'E10'),
        ('bcp', 'Critical sites with business-continuity plan updated in year', 2, 'sites (of 3)', 'BCP register', 'Risk Manager', 'E10'),
        ('sro_resp', 'Relevant SROs with a documented response in C1', 12, 'number', 'C1 count', 'Risk Manager', 'E10'),
        ('cyber', 'Cyber-security audits conducted', 2, 'number', 'IT audit reports', 'IT', 'E10'),
        ('reports', 'Sustainability reports issued with annual report', 1, 'reports', 'Annual report', 'Finance', 'E10'),
        ('rating', 'Third-party ESG rating score', 58, 'score / 100', 'Rating scorecard', 'Sustainability', 'E10'),
        ('griev', 'Average days to resolve worker grievances', 9, 'days', 'Grievance system', 'Compliance', 'E10'),
    ]
    rows = [(k, d, v, u, s_, o_, w) for k, d, v, u, s_, o_, w in D_]
    a, b = X.table(ws, r + 1, H, rows, inp=(1, 3, 4, 5))
    for i, x in enumerate(D_):
        if x[3] == '%':
            X.put(ws, a + i, 2, x[2], num_format='0.0%')
        X.ad['e1_' + x[0]] = ref(E1, a + i, 2)
    ws.set_column(1, 1, 30)
    ws.freeze_panes(9, 1)


def build_e2(X):
    ws, r = X.sheet(E2, 'E2', 'Emission factors and global warming potentials',
                    'Document every emission factor and GWP value used, with source, so that users can understand the measurement approach, inputs and assumptions. Factors are prioritised: (1) supplier/activity-specific, (2) national, (3) IPCC / international defaults.',
                    'IFRS S2.29(a)(i)-(iii), B19-B22 (CO2e and GWP), B26-B29 (emission factors), B38-B54 (Scope 3 measurement framework)',
                    'HB 4.1.140; DC 2.4.1; GHG Protocol Corporate Standard (2004)',
                    [38, 14, 14, 13, 13, 13, 13, 13, 13, 11, 14, 48], prep='Chan M.', tab=C['green'], status='In review')
    X.sec(ws, r, '1. Global warming potentials - 100-year, latest IPCC assessment at the reporting date (S2.B21-B22): AR6', 6)
    G = [('CO2', 1, 'co2'), ('CH4 - fossil origin', 29.8, 'ch4f'), ('CH4 - non-fossil (biomass)', 27.0, 'ch4b'), ('N2O', 273, 'n2o'),
         ('HFC blend R-410A (50% R-32 / 50% R-125)', 2256, 'r410a'), ('HFC-134a', 1530, 'r134a')]
    a, b = X.table(ws, r + 1, ['Gas', 'GWP-100', 'Source'], [(g[0], g[1], 'IPCC Sixth Assessment Report (AR6) WG1, Table 7.SM.7') for g in G], nf={1: '#,##0.0'})
    ws.merge_range(r + 1, 2, r + 1, 5, 'Source', X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], border=1))
    for i, g in enumerate(G):
        ws.merge_range(a + i, 2, a + i, 5, 'IPCC Sixth Assessment Report (AR6) WG1, Table 7.SM.7 - verify blend values', X.f(border=1, border_color=C['mid']))
        X.ad['gwp_' + g[2]] = ref(E2, a + i, 1)
    r = b + 2
    X.sec(ws, r, '2. Fuel combustion factors (Scope 1) - derived from IPCC 2006 Guidelines default values', 12)
    H = ['Fuel / use', 'Activity unit', 'Net calorific value (MJ per unit)', 'CO2 (kg/TJ)', 'CH4 (kg/TJ)', 'N2O (kg/TJ)', 'kg CO2 per unit', 'kg CH4 per unit', 'kg N2O per unit', 'Biogenic CO2?', 'kg CO2e per unit (excl. biogenic CO2)', 'Source / notes']
    Fz = [('dsl_stat', 'Diesel - stationary (generators)', 'litre', 36.12, 74100, 3.0, 0.6, 'No', 'IPCC 2006 Vol.2 Ch.2 Tables 2.2-2.3; NCV 43.0 MJ/kg x density 0.84 kg/l'),
          ('dsl_mob', 'Diesel - mobile (road vehicles)', 'litre', 36.12, 74100, 3.9, 3.9, 'No', 'IPCC 2006 Vol.2 Ch.3 Tables 3.2.1-3.2.2'),
          ('gas_mob', 'Gasoline - mobile (road vehicles)', 'litre', 32.78, 69300, 25.0, 8.0, 'No', 'IPCC 2006 Vol.2 Ch.3; NCV 44.3 MJ/kg x density 0.74 kg/l'),
          ('lpg', 'LPG - stationary (canteens)', 'kg', 47.3, 63100, 5.0, 0.1, 'No', 'IPCC 2006 Vol.2 Ch.2 Table 2.4 (commercial)'),
          ('wood', 'Wood biomass - boilers', 'tonne', 15600.0, 112000, 30.0, 4.0, 'Yes', 'IPCC 2006 Vol.2 Ch.2 Table 2.3; NCV 15.6 MJ/kg. Biogenic CO2 reported separately from the scopes (GHG Protocol)')]
    rows = []
    for k, n, u, ncv, c, m, n2, bio, src in Fz:
        rows.append([n, u, ncv, c, m, n2, lambda R: f'=C{R}*D{R}/1000000', lambda R: f'=C{R}*E{R}/1000000', lambda R: f'=C{R}*F{R}/1000000', bio,
                     (lambda R: f"=H{R}*{X.ad['gwp_ch4b'].split('!')[1]}+I{R}*{X.ad['gwp_n2o'].split('!')[1]}") if bio == 'Yes'
                     else (lambda R: f"=G{R}+H{R}*{X.ad['gwp_ch4f'].split('!')[1]}+I{R}*{X.ad['gwp_n2o'].split('!')[1]}"), src])
    a, b = X.table(ws, r + 1, H, rows, nf={2: '#,##0.00', 3: '#,##0', 4: '#,##0.0', 5: '#,##0.0', 6: '#,##0.0000', 7: '#,##0.000000', 8: '#,##0.000000', 10: '#,##0.0000'}, inp=(0, 1, 9, 11))
    ws.set_row(r + 1, 42)
    for i, f in enumerate(Fz):
        X.ad['ef_' + f[0]] = dict(co2=ref(E2, a + i, 6), ch4=ref(E2, a + i, 7), n2o=ref(E2, a + i, 8), co2e=ref(E2, a + i, 10), ncv=ref(E2, a + i, 2))
    X.name('EF_Diesel_Stat', E2, a + 0, 10)
    X.name('EF_Wood', E2, a + 4, 10)
    r = b + 2
    X.sec(ws, r, '3. Electricity and Scope 3 factors (kg CO2e per activity unit) - yellow = key assumptions to be replaced with current official / supplier-specific values', 12)
    H = ['Factor', 'Value', 'Unit', 'Data type (S2.B40)', 'Source / notes']
    Oz = [('grid', 'Cambodia national grid - location-based', 420.0, 'kg CO2e / MWh', 'Secondary - national average', 'ILLUSTRATIVE. Replace with the latest grid emission factor published by the Ministry of Environment / IGES list of grid emission factors'),
          ('elec_up', 'Electricity - upstream fuel (well-to-tank) and T&D losses', 105.0, 'kg CO2e / MWh', 'Secondary', 'Illustrative: 60 upstream + 45 T&D losses'),
          ('fabric', 'Woven fabric (cotton / polyester mix), cradle-to-gate', 11000.0, 'kg CO2e / tonne', 'Secondary - industry average', 'Illustrative LCA average; replace with mill-specific Higg FEM / product-footprint data (S2.B47-B49 prioritise primary data)'),
          ('trims', 'Trims, accessories and packaging', 2500.0, 'kg CO2e / tonne', 'Secondary - industry average', 'Illustrative'),
          ('spend_ogs', 'Other goods and services - spend-based', 300.0, 'kg CO2e / USD\'000', 'Secondary - EEIO', 'Illustrative environmentally-extended input-output factor'),
          ('spend_cap', 'Capital goods - spend-based', 450.0, 'kg CO2e / USD\'000', 'Secondary - EEIO', 'Illustrative'),
          ('wtt_dsl', 'Diesel - well-to-tank', 0.62, 'kg CO2e / litre', 'Secondary', 'Illustrative (order of magnitude of published WTT factors)'),
          ('wtt_gas', 'Gasoline - well-to-tank', 0.60, 'kg CO2e / litre', 'Secondary', 'Illustrative'),
          ('wtt_lpg', 'LPG - well-to-tank', 0.35, 'kg CO2e / kg', 'Secondary', 'Illustrative'),
          ('wtt_wood', 'Wood biomass - harvesting and transport', 40.0, 'kg CO2e / tonne', 'Secondary', 'Illustrative'),
          ('sea', 'Sea freight - container vessel', 0.012, 'kg CO2e / tonne-km', 'Secondary', 'Illustrative (GLEC-type default)'),
          ('road', 'Road freight - heavy truck', 0.11, 'kg CO2e / tonne-km', 'Secondary', 'Illustrative'),
          ('w_rec', 'Waste - recycling', 21.0, 'kg CO2e / tonne', 'Secondary', 'Illustrative'),
          ('w_land', 'Waste - landfill (mixed)', 520.0, 'kg CO2e / tonne', 'Secondary', 'Illustrative'),
          ('w_haz', 'Waste - hazardous incineration', 900.0, 'kg CO2e / tonne', 'Secondary', 'Illustrative'),
          ('air', 'Air travel - average passenger', 0.15, 'kg CO2e / passenger-km', 'Secondary', 'Illustrative; excludes radiative forcing'),
          ('truck_p', 'Commuting - collective truck / bus', 0.03, 'kg CO2e / passenger-km', 'Secondary', 'Illustrative'),
          ('moto', 'Commuting - motorcycle', 0.07, 'kg CO2e / passenger-km', 'Secondary', 'Illustrative'),
          ('eol', 'End-of-life treatment of sold textiles', 450.0, 'kg CO2e / tonne', 'Secondary', 'Illustrative blend of landfill / incineration / reuse')]
    rows = [(n, v, u, t, s_) for k, n, v, u, t, s_ in Oz]
    a, b = X.table(ws, r + 1, H, rows, nf={1: '#,##0.000'}, inp=(0, 2, 3), extra={1: {'bg_color': C['yellow']}})
    ws.merge_range(r + 1, 4, r + 1, 11, 'Source / notes', X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], border=1, valign='vcenter'))
    for i, o_ in enumerate(Oz):
        ws.merge_range(a + i, 4, a + i, 11, o_[5], X.f(border=1, border_color=C['mid'], text_wrap=True))
        X.ad['f_' + o_[0]] = ref(E2, a + i, 1)
    X.name('EF_Grid', E2, a, 1)
    X.note(ws, b + 2, 'IMPORTANT: Section 3 factors are placeholders of realistic magnitude for teaching purposes. Before use on a real engagement, replace each with a documented, current factor and record the source, vintage and geography (IFRS S2.B29). Changes in factors between periods are disclosed under S2.29(a)(iii)(3).', 12, h=30)
    ws.freeze_panes(8, 1)


def build_e3(X):
    ws, r = X.sheet(E3, 'E3', 'Scope 1 GHG emissions - calculation',
                    'Calculate absolute gross Scope 1 emissions (tCO2e) by source, by gas and disaggregated between the consolidated accounting group and other investees. Activity data from E1; factors from E2.',
                    'IFRS S2.29(a)(i)(1), 29(a)(ii)-(iv), B19-B29',
                    'HB 4.1.140; DC 2.4.1; GHG Protocol Corporate Standard (2004) ch. 3-6',
                    [5, 30, 16, 20, 13, 13, 9, 11, 12, 12, 11, 11, 11, 11, 12, 13, 26], prep='Chan M.', tab=C['green'])
    t = X.ad['e1t']
    L = [('Diesel - back-up generators', 'Factory A', 'Consolidated group', 'Stationary combustion', ref(E1, t, 5), 'litres', 'dsl_stat'),
         ('Diesel - back-up generators', 'Factory B', 'Consolidated group', 'Stationary combustion', ref(E1, t, 6), 'litres', 'dsl_stat'),
         ('Diesel - back-up generators', 'JV wash plant', 'Other investees', 'Stationary combustion', ref(E1, t, 7), 'litres', 'dsl_stat'),
         ('Wood biomass - steam boilers', 'Factory A', 'Consolidated group', 'Stationary combustion', ref(E1, t, 8), 'tonnes', 'wood'),
         ('Wood biomass - steam boilers', 'Factory B', 'Consolidated group', 'Stationary combustion', ref(E1, t, 9), 'tonnes', 'wood'),
         ('Wood biomass - steam boilers', 'JV wash plant', 'Other investees', 'Stationary combustion', ref(E1, t, 10), 'tonnes', 'wood'),
         ('LPG - canteens', 'Factories A and B', 'Consolidated group', 'Stationary combustion', X.ad['e1_lpg'], 'kg', 'lpg'),
         ('Diesel - vehicles', 'Group fleet', 'Consolidated group', 'Mobile combustion', X.ad['e1_veh_dsl'], 'litres', 'dsl_mob'),
         ('Gasoline - vehicles', 'Group fleet', 'Consolidated group', 'Mobile combustion', X.ad['e1_veh_gas'], 'litres', 'gas_mob'),
         ('Refrigerant leakage R-410A', 'Factories A and B', 'Consolidated group', 'Fugitive emissions', X.ad['e1_r410a'], 'kg', 'R:r410a'),
         ('Refrigerant leakage R-134a', 'JV wash plant', 'Other investees', 'Fugitive emissions', X.ad['e1_r134a'], 'kg', 'R:r134a')]
    H = ['#', 'Emission source', 'Facility', 'Boundary (S2.29(a)(iv))', 'Category', 'Activity data', 'Unit', 'kg CO2 / unit', 'kg CH4 / unit', 'kg N2O / unit',
         'CO2 (t)', 'CH4 (tCO2e)', 'N2O (tCO2e)', 'HFC (tCO2e)', 'TOTAL Scope 1 (tCO2e)', 'Memo: biogenic CO2 (t) - outside scopes', 'Factor source']
    rows = []
    for i, (s_, fac, bnd, cat, act, u, k) in enumerate(L):
        if k.startswith('R:'):
            g = X.ad['gwp_' + k[2:]]
            rows.append([i + 1, s_, fac, bnd, cat, '=' + act, u, '', '', '', 0, 0, 0, lambda R, g=g: f'=F{R}*{g}/1000',
                         lambda R: f'=SUM(K{R}:N{R})', 0, 'E2 s.1 (GWP); mass-balance of top-ups'])
        else:
            e = X.ad['ef_' + k]
            bio = k == 'wood'
            ch4 = X.ad['gwp_ch4b'] if bio else X.ad['gwp_ch4f']
            rows.append([i + 1, s_, fac, bnd, cat, '=' + act, u, '=' + e['co2'], '=' + e['ch4'], '=' + e['n2o'],
                         0 if bio else (lambda R: f'=F{R}*H{R}/1000'),
                         lambda R, c=ch4: f'=F{R}*I{R}*{c}/1000', lambda R: f"=F{R}*J{R}*{X.ad['gwp_n2o']}/1000", 0,
                         lambda R: f'=SUM(K{R}:N{R})', (lambda R: f'=F{R}*H{R}/1000') if bio else 0, 'E2 s.2 (IPCC 2006)'])
    a, b = X.table(ws, r, H, rows, nf={7: '#,##0.0000', 8: '#,##0.000000', 9: '#,##0.000000', 10: '#,##0.0', 11: '#,##0.0', 12: '#,##0.0', 13: '#,##0.0', 14: '#,##0.0', 15: '#,##0.0'},
                   inp=(1, 2, 3, 4, 6, 16))
    ws.set_row(r, 44)
    X.put(ws, b + 1, 1, 'TOTAL', bold=True)
    for j in range(10, 16):
        X.put(ws, b + 1, j, f'=SUM({A(a, j)}:{A(b, j)})', bold=True, num_format='#,##0.0', bg_color=C['grey'])
    X.ad['e3'] = (a, b)
    tot = b + 1
    r = b + 3
    X.sec(ws, r, 'Summary', 6)
    Sm = [('Scope 1 - consolidated accounting group (S2.29(a)(iv)(1))', f'=SUMIF({A(a, 3)}:{A(b, 3)},"Consolidated group",{A(a, 14)}:{A(b, 14)})', 'GHG_S1_G'),
          ('Scope 1 - other investees: JV under operational control (S2.29(a)(iv)(2))', f'=SUMIF({A(a, 3)}:{A(b, 3)},"Other investees",{A(a, 14)}:{A(b, 14)})', 'GHG_S1_O'),
          ('TOTAL Scope 1 (tCO2e)', f'={A(tot, 14)}', 'GHG_S1'),
          ('  Stationary combustion', f'=SUMIF({A(a, 4)}:{A(b, 4)},"Stationary combustion",{A(a, 14)}:{A(b, 14)})', 'S1_Stat'),
          ('  Mobile combustion', f'=SUMIF({A(a, 4)}:{A(b, 4)},"Mobile combustion",{A(a, 14)}:{A(b, 14)})', 'S1_Mob'),
          ('  Fugitive emissions', f'=SUMIF({A(a, 4)}:{A(b, 4)},"Fugitive emissions",{A(a, 14)}:{A(b, 14)})', 'S1_Fug'),
          ('  of which diesel generators', f'=SUMIF({A(a, 1)}:{A(b, 1)},"Diesel - back-up generators",{A(a, 14)}:{A(b, 14)})', 'S1_Gen'),
          ('  of which biomass boilers (CH4 and N2O only)', f'=SUMIF({A(a, 1)}:{A(b, 1)},"Wood biomass*",{A(a, 14)}:{A(b, 14)})', 'S1_Bio'),
          ('  of which LPG', f'=SUMIF({A(a, 1)}:{A(b, 1)},"LPG*",{A(a, 14)}:{A(b, 14)})', 'S1_LPG'),
          ('By gas: CO2 (t)', f'={A(tot, 10)}', 'S1_CO2'), ('By gas: CH4 (tCO2e)', f'={A(tot, 11)}', 'S1_CH4'),
          ('By gas: N2O (tCO2e)', f'={A(tot, 12)}', 'S1_N2O'), ('By gas: HFCs (tCO2e)', f'={A(tot, 13)}', 'S1_HFC'),
          ('Memo: biogenic CO2 from biomass combustion (t) - reported separately', f'={A(tot, 15)}', 'GHG_Biogenic')]
    for i, (l, f, nm) in enumerate(Sm):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 4, l, X.f(border=1, border_color=C['mid'], bold=not l.startswith('  ')))
        X.put(ws, r + 1 + i, 5, f, num_format='#,##0.0', bold=not l.startswith('  '))
        X.name(nm, E3, r + 1 + i, 5)
    X.note(ws, r + len(Sm) + 2, 'Method: activity data x emission factor x GWP (GHG Protocol calculation approach). Organisational boundary: operational control (S2.B27). CO2 from biomass is biogenic and shown as a memo item; CH4 and N2O from biomass are in Scope 1. No direct measurement (CEMS) is used.', 17, h=30)
    ws.freeze_panes(9, 3)


def build_e4(X):
    ws, r = X.sheet(E4, 'E4', 'Scope 2 GHG emissions - location-based calculation and contractual instruments',
                    'Calculate location-based Scope 2 emissions from purchased electricity and document information about contractual instruments needed to inform users\' understanding.',
                    'IFRS S2.29(a)(i)(2), 29(a)(iv), 29(a)(v), B30-B31',
                    'HB 4.1.140; DC 2.4.1; GHG Protocol Scope 2 Guidance (for definitions)',
                    [28, 22, 16, 16, 16, 44], prep='Chan M.', tab=C['green'])
    t = X.ad['e1t']
    rows = [('Factory A', 'Consolidated group', '=' + ref(E1, t, 1), '=EF_Grid', lambda R: f'=C{R}*D{R}/1000', 'EDC medium-voltage supply; 12 monthly invoices agreed'),
            ('Factory B', 'Consolidated group', '=' + ref(E1, t, 2), '=EF_Grid', lambda R: f'=C{R}*D{R}/1000', 'EDC medium-voltage supply; 12 monthly invoices agreed'),
            ('JV wash plant', 'Other investees', '=' + ref(E1, t, 3), '=EF_Grid', lambda R: f'=C{R}*D{R}/1000', 'EDC supply via SEZ operator; sub-meter readings')]
    a, b = X.table(ws, r, ['Facility', 'Boundary (S2.29(a)(iv))', 'Purchased electricity (MWh)', 'Grid factor (kg CO2e/MWh)', 'Scope 2 location-based (tCO2e)', 'Evidence'], rows,
                   nf={4: '#,##0.0', 3: '#,##0.0'}, inp=(0, 1, 5))
    X.put(ws, b + 1, 0, 'TOTAL', bold=True); X.put(ws, b + 1, 1, '')
    X.put(ws, b + 1, 2, f'=SUM({A(a, 2)}:{A(b, 2)})', bold=True)
    X.put(ws, b + 1, 3, '')
    X.put(ws, b + 1, 4, f'=SUM({A(a, 4)}:{A(b, 4)})', bold=True, num_format='#,##0.0')
    X.name('Elec_MWh', E4, b + 1, 2)
    r = b + 3
    Sm = [('Scope 2 (location-based) - consolidated accounting group', f'=SUMIF({A(a, 1)}:{A(b, 1)},"Consolidated group",{A(a, 4)}:{A(b, 4)})', 'GHG_S2_G'),
          ('Scope 2 (location-based) - other investees', f'=SUMIF({A(a, 1)}:{A(b, 1)},"Other investees",{A(a, 4)}:{A(b, 4)})', 'GHG_S2_O'),
          ('TOTAL Scope 2 location-based (tCO2e)', f'={A(b + 1, 4)}', 'GHG_S2'),
          ('Memo: solar electricity self-generated and consumed (MWh) - zero Scope 2', '=' + ref(E1, t, 4), 'Solar_MWh'),
          ('Memo: renewable share of total electricity consumed', f'={A(r + 4, 4)}/({A(r + 4, 4)}+{A(b + 1, 2)})', 'RE_Share')]
    for i, (l, f, nm) in enumerate(Sm):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 3, l, X.f(border=1, border_color=C['mid'], bold=True))
        X.put(ws, r + 1 + i, 4, f, num_format='0.0%' if nm == 'RE_Share' else '#,##0.0', bold=True)
        X.name(nm, E4, r + 1 + i, 4)
    r += 8
    X.sec(ws, r, 'Contractual instruments (S2.29(a)(v), B30-B31)', 6)
    rows = [('Energy attribute certificates (I-REC etc.)', 'None purchased in FY2025', 'n/a', '', '', 'Considered for 2028-2030 (C1 L6)'),
            ('Power purchase agreements', 'None', 'n/a', '', '', ''),
            ('Rooftop solar', 'Owned asset, behind-the-meter; no certificates sold', 'On-site generation - not a contractual instrument', '', '', 'Reduces purchased MWh'),
            ('Market-based Scope 2', 'Not disclosed - no contractual instruments exist, so market-based would equal location-based using a residual mix (not available for Cambodia)', 'n/a', '', '', 'Location-based figure is the required disclosure')]
    X.table(ws, r + 1, ['Instrument', 'Position FY2025', 'Treatment', '', '', 'Comment'], rows, inp=(1, 2, 5))
    ws.freeze_panes(9, 1)


def build_e5(X):
    ws, r = X.sheet(E5, 'E5', 'Scope 3 GHG emissions - calculation by category',
                    'Measure Scope 3 emissions for all 15 GHG Protocol categories considered, using the IFRS S2 measurement framework (prioritising direct measurement, primary data, timely and verified data where available without undue cost or effort), and document inclusion / exclusion of each category.',
                    'IFRS S2.29(a)(i)(3), 29(a)(vi), B32-B57; transition relief S2.C4(b), C5',
                    'HB 4.1.140; DC 2.4.1; GHG Protocol Corporate Value Chain (Scope 3) Standard (2011)',
                    [6, 30, 34, 15, 13, 13, 18, 14, 22, 14, 10, 30], prep='Chan M.', tab=C['green'], status='In review')
    e = X.ad
    t = e['e1t']
    dsl_all = f"({ref(E1, t, 5)}+{ref(E1, t, 6)}+{ref(E1, t, 7)}+{e['e1_veh_dsl']})"
    wood_all = f"({ref(E1, t, 8)}+{ref(E1, t, 9)}+{ref(E1, t, 10)})"
    pkm = f"{e['e1_hc']}*{e['e1_days']}*2*{e['e1_km_commute']}"
    L = [(1, 'Purchased goods and services', 'Fabric', '=' + e['e1_fabric'], 'tonnes', 'fabric', 'Average-data (mass)', 'Secondary', 'No'),
         (1, 'Purchased goods and services', 'Trims, accessories and packaging', '=' + e['e1_trims'], 'tonnes', 'trims', 'Average-data (mass)', 'Secondary', 'No'),
         (1, 'Purchased goods and services', 'Other goods and services', '=' + e['e1_ogs'], 'USD\'000', 'spend_ogs', 'Spend-based', 'Secondary', 'No'),
         (2, 'Capital goods', 'PPE additions FY2025', '=Capex25', 'USD\'000', 'spend_cap', 'Spend-based', 'Secondary', 'No'),
         (3, 'Fuel- and energy-related activities', 'Diesel - well-to-tank', '=' + dsl_all, 'litres', 'wtt_dsl', 'Average-data', 'Primary activity / secondary EF', 'No'),
         (3, 'Fuel- and energy-related activities', 'Gasoline - well-to-tank', '=' + e['e1_veh_gas'], 'litres', 'wtt_gas', 'Average-data', 'Primary activity / secondary EF', 'No'),
         (3, 'Fuel- and energy-related activities', 'LPG - well-to-tank', '=' + e['e1_lpg'], 'kg', 'wtt_lpg', 'Average-data', 'Primary activity / secondary EF', 'No'),
         (3, 'Fuel- and energy-related activities', 'Wood biomass - harvesting and transport', '=' + wood_all, 'tonnes', 'wtt_wood', 'Average-data', 'Primary activity / secondary EF', 'No'),
         (3, 'Fuel- and energy-related activities', 'Electricity - upstream and T&D losses', '=Elec_MWh', 'MWh', 'elec_up', 'Average-data', 'Primary activity / secondary EF', 'No'),
         (4, 'Upstream transportation and distribution', 'Inbound sea freight', f"=({e['e1_fabric']}+{e['e1_trims']})*{e['e1_km_sea_in']}", 'tonne-km', 'sea', 'Distance-based', 'Primary activity / secondary EF', 'No'),
         (4, 'Upstream transportation and distribution', 'Inbound road haulage', f"=({e['e1_fabric']}+{e['e1_trims']})*{e['e1_km_road']}", 'tonne-km', 'road', 'Distance-based', 'Primary activity / secondary EF', 'No'),
         (5, 'Waste generated in operations', 'Fabric off-cuts recycled', '=' + e['e1_w_rec'], 'tonnes', 'w_rec', 'Waste-type-specific', 'Primary activity / secondary EF', 'No'),
         (5, 'Waste generated in operations', 'General waste to landfill', '=' + e['e1_w_land'], 'tonnes', 'w_land', 'Waste-type-specific', 'Primary activity / secondary EF', 'No'),
         (5, 'Waste generated in operations', 'Hazardous waste incinerated', '=' + e['e1_w_haz'], 'tonnes', 'w_haz', 'Waste-type-specific', 'Primary activity / secondary EF', 'No'),
         (6, 'Business travel', 'Air travel', '=' + e['e1_air_pkm'], 'passenger-km', 'air', 'Distance-based', 'Primary activity / secondary EF', 'No'),
         (7, 'Employee commuting', 'Collective truck / bus', f"={pkm}*{e['e1_sh_truck']}", 'passenger-km', 'truck_p', 'Distance-based (survey)', 'Secondary (survey estimate)', 'No'),
         (7, 'Employee commuting', 'Motorcycle', f"={pkm}*(1-{e['e1_sh_truck']})", 'passenger-km', 'moto', 'Distance-based (survey)', 'Secondary (survey estimate)', 'No'),
         (9, 'Downstream transportation and distribution', 'Road haulage to port (FOB - paid by TTM; reported here by policy)', f"={e['e1_fg']}*{e['e1_km_road']}", 'tonne-km', 'road', 'Distance-based', 'Primary activity / secondary EF', 'No'),
         (9, 'Downstream transportation and distribution', 'Sea freight to destination (paid by buyers)', f"={e['e1_fg']}*{e['e1_km_sea_out']}", 'tonne-km', 'sea', 'Distance-based', 'Secondary', 'No'),
         (12, 'End-of-life treatment of sold products', 'Garments sold', '=' + e['e1_fg'], 'tonnes', 'eol', 'Average-data', 'Secondary', 'No')]
    H = ['Cat.', 'Scope 3 category', 'Line item', 'Activity data', 'Unit', 'Emission factor (kg CO2e/unit)', 'Calculation method (GHG Protocol)', 'tCO2e',
         'Input type (S2.B40(b), B47-B49)', 'Based mainly on primary activity data?', 'Verified? (B40(d))', 'Evidence / comment']
    rows = []
    for c_, n, li, act, u, k, meth, typ, ver in L:
        rows.append([c_, n, li, act, u, '=' + e['f_' + k], meth, lambda R: f'=D{R}*F{R}/1000', typ,
                     'Yes' if typ.startswith('Primary') else 'No', ver, 'E1 / E2'])
    a, b = X.table(ws, r, H, rows, nf={5: '#,##0.000', 7: '#,##0.0'}, inp=(1, 2, 4, 6, 8, 9, 10, 11))
    ws.set_row(r, 44)
    X.put(ws, b + 1, 2, 'TOTAL Scope 3 (tCO2e)', bold=True)
    X.put(ws, b + 1, 7, f'=SUM({A(a, 7)}:{A(b, 7)})', bold=True, num_format='#,##0.0', bg_color=C['grey'])
    X.name('GHG_S3', E5, b + 1, 7)
    X.put(ws, b + 2, 2, 'Share of Scope 3 measured using primary activity data (S2.B56(a))', bold=True)
    X.put(ws, b + 2, 7, f'=SUMIF({A(a, 9)}:{A(b, 9)},"Yes",{A(a, 7)}:{A(b, 7)})/{A(b + 1, 7)}', bold=True, num_format='0.0%')
    X.name('S3_Primary', E5, b + 2, 7)
    X.put(ws, b + 3, 2, 'Share of Scope 3 measured using verified inputs (S2.B56(b))', bold=True)
    X.put(ws, b + 3, 7, f'=SUMIF({A(a, 10)}:{A(b, 10)},"Yes",{A(a, 7)}:{A(b, 7)})/{A(b + 1, 7)}', bold=True, num_format='0.0%')
    da, db = a, b
    r = b + 6
    X.sec(ws, r, 'Category summary and inclusion assessment (S2.29(a)(vi)(1), B32-B37)', 12)
    CAT = [(1, 'Purchased goods and services', 'Yes', 'Largest category - fabric is 62% of cost of sales'),
           (2, 'Capital goods', 'Yes', 'Significant capex in FY2025'),
           (3, 'Fuel- and energy-related activities (not in Scope 1 or 2)', 'Yes', 'Upstream of all fuels and electricity incl. JV'),
           (4, 'Upstream transportation and distribution', 'Yes', 'Inbound logistics paid by TTM'),
           (5, 'Waste generated in operations', 'Yes', 'Third-party treatment'),
           (6, 'Business travel', 'Yes', 'Small but measured'),
           (7, 'Employee commuting', 'Yes', '6,480 workers; survey-based'),
           (8, 'Upstream leased assets', 'No', 'No leased assets outside Scope 1 and 2 boundary'),
           (9, 'Downstream transportation and distribution', 'Yes', 'FOB terms - buyer-controlled freight'),
           (10, 'Processing of sold products', 'No', 'Finished garments - no further processing'),
           (11, 'Use of sold products', 'No', 'Indirect use-phase emissions (laundering) are optional under the GHG Protocol; excluded and disclosed as such'),
           (12, 'End-of-life treatment of sold products', 'Yes', 'Estimated from mass shipped'),
           (13, 'Downstream leased assets', 'No', 'None'),
           (14, 'Franchises', 'No', 'None'),
           (15, 'Investments', 'No', 'JV already included at 100% in Scope 1 and 2 under operational control; no other investments. Entity has no asset-management, banking or insurance activities (S2.29(a)(vi)(2), B58-B63 not applicable)')]
    rows = [[c_, n, inc, lambda R, da=da, db=db: f'=SUMIF({A(da, 0)}:{A(db, 0)},A{R},{A(da, 7)}:{A(db, 7)})',
             lambda R, da=da, db=db: f'=IF({A(db + 1, 7)}=0,0,D{R}/{A(db + 1, 7)})', why] for c_, n, inc, why in CAT]
    a, b = X.table(ws, r + 1, ['Cat.', 'Scope 3 category', 'Included?', 'tCO2e FY2025', '% of Scope 3', 'Rationale / comment'], rows, nf={3: '#,##0.0', 4: '0.0%'}, inp=(2, 5))
    ws.merge_range(r + 1, 5, r + 1, 11, 'Rationale / comment', X.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], border=1, valign='vcenter'))
    for i, c_ in enumerate(CAT):
        ws.merge_range(a + i, 5, a + i, 11, c_[3], X.f(border=1, border_color=C['mid'], text_wrap=True, font_color=C['inp']))
    X.cf_status(ws, a, b, 2)
    X.ad['s3cat'] = (a, b)
    X.put(ws, b + 1, 1, 'Check: summary = detail', bold=True)
    X.put(ws, b + 1, 3, f'=SUM({A(a, 3)}:{A(b, 3)})-{A(db + 1, 7)}', num_format='#,##0.0;-#,##0.0;"OK"')
    ws.freeze_panes(9, 3)


def build_e6(X):
    ws, r = X.sheet(E6, 'E6', 'GHG emissions summary - disclosure table',
                    'Summarise absolute gross GHG emissions by scope for disclosure, with comparatives, disaggregation, intensity ratios and the description of the measurement approach, inputs and assumptions.',
                    'IFRS S2.29(a)(i)-(vi), B19-B57; IFRS S1.70, B49-B59 (comparatives); S2.C3-C5',
                    'HB 4.1.140, 2.3.110; DC 2.4.1, 1.8',
                    [52, 15, 15, 14, 12, 22, 44], prep='Sok D.', tab=C['green'], status='In review')
    H = ['tCO2e', 'FY2025', 'FY2024 (comparative)', 'Change', 'Change %', 'ISSB reference', 'Notes']
    X.hdr(ws, r, H)
    s3a, s3b = X.ad['s3cat']
    L = [('Scope 1 - consolidated accounting group', '=GHG_S1_G', 1214.0, 'S2.29(a)(i)(1), (iv)(1)', 'E3'),
         ('Scope 1 - other investees (JV, operational control)', '=GHG_S1_O', 262.0, 'S2.29(a)(iv)(2)', 'E3'),
         ('TOTAL SCOPE 1', None, None, 'S2.29(a)(i)(1)', ''),
         ('Scope 2 (location-based) - consolidated accounting group', '=GHG_S2_G', 4030.0, 'S2.29(a)(i)(2), (v)', 'E4. FY2024 before rooftop solar'),
         ('Scope 2 (location-based) - other investees', '=GHG_S2_O', 1070.0, 'S2.29(a)(iv)(2)', 'E4'),
         ('TOTAL SCOPE 2 (location-based)', None, None, 'S2.29(a)(i)(2)', ''),
         ('TOTAL SCOPE 1 + 2', None, None, '', 'Basis of target T1 (E9)')]
    rr = r + 1
    row_of = {}
    for i, (l, f, py, isb, nt) in enumerate(L):
        R = rr + i + 1  # excel row
        tot = l.startswith('TOTAL')
        X.put(ws, rr + i, 0, l, bold=tot)
        if l == 'TOTAL SCOPE 1':
            X.put(ws, rr + i, 1, f'=B{R - 2}+B{R - 1}', bold=True, num_format='#,##0'); X.put(ws, rr + i, 2, f'=C{R - 2}+C{R - 1}', bold=True)
        elif l.startswith('TOTAL SCOPE 2'):
            X.put(ws, rr + i, 1, f'=B{R - 2}+B{R - 1}', bold=True); X.put(ws, rr + i, 2, f'=C{R - 2}+C{R - 1}', bold=True)
        elif l == 'TOTAL SCOPE 1 + 2':
            X.put(ws, rr + i, 1, f'=B{R - 4}+B{R - 1}', bold=True, bg_color=C['grey']); X.put(ws, rr + i, 2, f'=C{R - 4}+C{R - 1}', bold=True, bg_color=C['grey'])
        else:
            X.put(ws, rr + i, 1, f); X.put(ws, rr + i, 2, py, num_format='#,##0')
        X.put(ws, rr + i, 3, f'=B{R}-C{R}', bold=tot, num_format='#,##0;-#,##0')
        X.put(ws, rr + i, 4, f'=IF(C{R}=0,0,D{R}/C{R})', num_format='0.0%', bold=tot)
        X.put(ws, rr + i, 5, isb); X.put(ws, rr + i, 6, nt)
        row_of[l] = rr + i
    X.name('GHG_S12', E6, row_of['TOTAL SCOPE 1 + 2'], 1)
    X.name('GHG_S12_PY', E6, row_of['TOTAL SCOPE 1 + 2'], 2)
    X.name('GHG_S1_PY', E6, row_of['TOTAL SCOPE 1'], 2)
    X.name('GHG_S2_PY', E6, row_of['TOTAL SCOPE 2 (location-based)'], 2)
    rr += len(L)
    X.put(ws, rr, 0, 'Scope 3 by category:', bold=True, italic=True)
    for j in range(1, 7):
        X.put(ws, rr, j, '')
    rr += 1
    c0 = rr
    cats = ['1 Purchased goods and services', '2 Capital goods', '3 Fuel- and energy-related activities', '4 Upstream transportation and distribution', '5 Waste generated in operations',
            '6 Business travel', '7 Employee commuting', '8 Upstream leased assets', '9 Downstream transportation and distribution', '10 Processing of sold products',
            '11 Use of sold products', '12 End-of-life treatment of sold products', '13 Downstream leased assets', '14 Franchises', '15 Investments']
    for i, cname in enumerate(cats):
        inc = ref(E5, s3a + i, 2)
        X.put(ws, rr, 0, '   Category ' + cname)
        X.put(ws, rr, 1, f'=IF({inc}="Yes",{ref(E5, s3a + i, 3)},"excluded")', num_format='#,##0')
        X.put(ws, rr, 2, 'n/a', font_color=C['inp'], align='right')
        X.put(ws, rr, 3, 'n/a', align='right'); X.put(ws, rr, 4, 'n/a', align='right')
        X.put(ws, rr, 5, 'S2.29(a)(vi)(1)'); X.put(ws, rr, 6, 'E5' if i else 'E5. No comparative: relief S2.C4(b), C5')
        rr += 1
    X.ad['e6cat'] = (c0, rr - 1)
    X.put(ws, rr, 0, 'TOTAL SCOPE 3', bold=True); X.put(ws, rr, 1, '=GHG_S3', bold=True)
    X.put(ws, rr, 2, 'n/a', align='right'); X.put(ws, rr, 3, 'n/a', align='right'); X.put(ws, rr, 4, 'n/a', align='right')
    X.put(ws, rr, 5, 'S2.29(a)(i)(3)'); X.put(ws, rr, 6, 'First measured FY2025')
    rr += 1
    X.put(ws, rr, 0, 'TOTAL GROSS GHG EMISSIONS (Scope 1 + 2 + 3)', bold=True, bg_color=C['grey'])
    X.put(ws, rr, 1, '=GHG_S12+GHG_S3', bold=True, bg_color=C['grey'])
    X.name('GHG_Total', E6, rr, 1)
    for j, v in ((2, 'n/a'), (3, 'n/a'), (4, 'n/a'), (5, 'S2.29(a)(i)'), (6, 'Gross - no offsets or removals deducted')):
        X.put(ws, rr, j, v, align='right' if j < 5 else 'left')
    rr += 1
    X.put(ws, rr, 0, 'Memo: biogenic CO2 from biomass combustion (outside scopes)'); X.put(ws, rr, 1, '=GHG_Biogenic')
    X.put(ws, rr, 2, 12950.0, num_format='#,##0'); X.put(ws, rr, 3, f'=B{rr + 1}-C{rr + 1}', num_format='#,##0;-#,##0'); X.put(ws, rr, 4, f'=D{rr + 1}/C{rr + 1}', num_format='0.0%')
    X.put(ws, rr, 5, 'GHG Protocol; S1.B29'); X.put(ws, rr, 6, 'Disclosed separately for transparency (linked to CR-04)')
    rr += 2
    X.sec(ws, rr, 'Intensity ratios (entity-developed metrics - S1.46(b), S1.50; S2.B67)', 7)
    prod = f"({ref(E1, X.ad['e1t'], 11)}+{ref(E1, X.ad['e1t'], 12)})"
    I = [('Scope 1+2 intensity - tCO2e per USD million revenue', '=GHG_S12/(Rev25/1000)', '=GHG_S12_PY/(Rev24/1000)', 'GHG_Int_Rev'),
         ('Scope 1+2 intensity - kg CO2e per garment produced', f'=GHG_S12*1000/({prod}*1000)', 0.2745, 'GHG_Int_Pc'),
         ('Total (Scope 1+2+3) intensity - tCO2e per USD million revenue', '=GHG_Total/(Rev25/1000)', 'n/a', None)]
    for l, f, py, nm in I:
        X.put(ws, rr + 1, 0, l)
        X.put(ws, rr + 1, 1, f, num_format='#,##0.000' if 'garment' in l else '#,##0.0')
        if isinstance(py, str) and py.startswith('='):
            X.put(ws, rr + 1, 2, py, num_format='#,##0.0')
        elif py == 'n/a':
            X.put(ws, rr + 1, 2, py, align='right')
        else:
            X.put(ws, rr + 1, 2, py, num_format='#,##0.000')
        X.put(ws, rr + 1, 3, f'=IF(ISNUMBER(C{rr + 2}),B{rr + 2}-C{rr + 2},"n/a")', num_format='#,##0.000;-#,##0.000', align='right')
        X.put(ws, rr + 1, 4, f'=IF(ISNUMBER(C{rr + 2}),D{rr + 2}/C{rr + 2},"n/a")', num_format='0.0%', align='right')
        X.put(ws, rr + 1, 5, 'S1.50; S2.B67'); X.put(ws, rr + 1, 6, 'Defined, labelled and method disclosed (S1.50, 53)')
        if nm:
            X.name(nm, E6, rr + 1, 1)
        rr += 1
    rr += 2
    X.sec(ws, rr, 'Measurement approach, inputs and assumptions (S2.29(a)(ii)-(iii), B23-B29)', 7)
    T = [('Standard applied', 'GHG Protocol Corporate Accounting and Reporting Standard (2004); Scope 3 categories per the Corporate Value Chain (Scope 3) Standard (2011)', 'S2.29(a)(ii), B23-B25'),
         ('Organisational boundary', 'Operational control. The 50% JV wash plant is equity-accounted in the FS but operationally controlled; 100% of its Scope 1 and 2 emissions are included and shown as "other investees"', 'S2.29(a)(iii)(1), (iv), B27'),
         ('Reason for the approach', 'Operational control best reflects emissions that management can directly reduce and matches the scope of buyers\' supplier requirements', 'S2.29(a)(iii)(2), B27(b)'),
         ('Gases and GWP', 'CO2, CH4, N2O and HFCs; no PFCs, SF6 or NF3 sources. GWP-100 from IPCC AR6', 'S2.B20-B22'),
         ('Emission factors', 'IPCC 2006 default factors for fuels; national grid factor for electricity; secondary industry-average factors for Scope 3 (see E2)', 'S2.B29'),
         ('Changes in approach during the period', 'Scope 3 measured for the first time. No change to Scope 1 and 2 methods or factors versus FY2024', 'S2.29(a)(iii)(3)'),
         ('Reporting periods of value-chain data', 'All data relate to calendar 2025; no value-chain data with a different reporting period used', 'S2.B19'),
         ('December 2025 amendments to IFRS S2 (GHG)', 'Effective 1-Jan-2027; not early-applied. No expected effect (no financed emissions; AR6 GWPs already used)', 'S2.C1B (as amended)')]
    for l, v, isb in T:
        X.put(ws, rr + 1, 0, l, bold=True)
        ws.merge_range(rr + 1, 1, rr + 1, 4, v, X.f(text_wrap=True, border=1, border_color=C['mid'], font_color=C['inp']))
        ws.merge_range(rr + 1, 5, rr + 1, 6, isb, X.f(text_wrap=True, border=1, border_color=C['mid']))
        ws.set_row(rr + 1, 30)
        rr += 1
    ws.freeze_panes(9, 1)


def build_e7(X):
    ws, r = X.sheet(E7, 'E7', 'Cross-industry climate-related metrics (other than GHG)',
                    'Compute the cross-industry metric categories: assets / business activities vulnerable to transition and physical risks, aligned with opportunities, capital deployment, internal carbon price and remuneration.',
                    'IFRS S2.29(b)-(g), 30-31, B64-B65',
                    'HB 4.1.150-4.1.180; DC 2.4.2',
                    [46, 44, 15, 15, 11, 16, 40], prep='Sok D.', tab=C['green'], status='In review')
    H = ['Metric', 'Definition / basis used', 'Amount (USD\'000)', 'Denominator (USD\'000)', '%', 'ISSB reference', 'Source / connectivity to FS (F3)']
    X.hdr(ws, r, H)
    rows = [
        ('(b) Transition risk - revenue from customers with contractual supplier-decarbonisation requirements', 'Revenue from brands whose supplier codes set binding GHG / fuel-switch requirements (CR-03)', 52900, '=Rev25', 'S2.29(b)', 'Sales ledger by customer; contracts reviewed', 'Rev_Transition'),
        ('(c) Physical risk - carrying amount of PPE and inventories at Factory B (flood plain)', 'Assets located inside the 1-in-25-year flood extent (CR-01)', 12800, '=PPE25+Inv25', 'S2.29(c)', 'Fixed-asset register by location; inventory by site', 'Assets_Physical'),
        ('(c) Physical risk - revenue from production at Factory B', 'Business activity exposed to acute flood risk', 35200, '=Rev25', 'S2.29(c)', 'Production-based revenue allocation', None),
        ('(d) Opportunities - revenue from recycled-content / verified low-carbon programmes', 'Orders placed under buyers\' preferred low-carbon supplier programmes (CO-02)', 11100, '=Rev25', 'S2.29(d)', 'Sales ledger by programme code', 'Rev_Opp'),
    ]
    rr = r + 1
    for l, d, amt, den, isb, src, nm in rows:
        R = rr + 1
        X.put(ws, rr, 0, l, bold=True); X.put(ws, rr, 1, d, font_color=C['inp']); X.put(ws, rr, 2, amt); X.put(ws, rr, 3, den)
        X.put(ws, rr, 4, f'=C{R}/D{R}', num_format='0.0%', bold=True); X.put(ws, rr, 5, isb); X.put(ws, rr, 6, src)
        if nm:
            X.name(nm, E7, rr, 2)
            X.name(nm + '_Pct', E7, rr, 4)
        rr += 1
    rr += 1
    X.sec(ws, rr, '(e) Capital deployment towards climate-related risks and opportunities - S2.29(e)', 7)
    cap = [('Rooftop solar phase 1 (CO-01)', 1450), ('Electric heat-pump boiler pilot (CR-04)', 620), ('Flood wall and pump station - Factory B (CR-01)', 540), ('LED and efficient motors (CO-01)', 310)]
    a = rr + 1
    for l, v in cap:
        X.put(ws, rr + 1, 0, l); X.put(ws, rr + 1, 1, 'PPE additions - capex register', font_color=C['inp']); X.put(ws, rr + 1, 2, v)
        X.put(ws, rr + 1, 3, ''); X.put(ws, rr + 1, 4, ''); X.put(ws, rr + 1, 5, 'S2.29(e)'); X.put(ws, rr + 1, 6, 'C2 section 1')
        rr += 1
    X.put(ws, rr + 1, 0, 'Total climate-related capital expenditure', bold=True); X.put(ws, rr + 1, 1, 'As % of total capex (A1)')
    X.put(ws, rr + 1, 2, f'=SUM({A(a, 2)}:{A(rr, 2)})', bold=True); X.put(ws, rr + 1, 3, '=Capex25')
    X.put(ws, rr + 1, 4, f'=C{rr + 2}/D{rr + 2}', num_format='0.0%', bold=True); X.put(ws, rr + 1, 5, 'S2.29(e)')
    X.put(ws, rr + 1, 6, f"=IF(ABS(C{rr + 2}+{210}-Capex_SRO)<1,\"Agrees to C2 (C2 also includes SR-02 fire-safety capex of 210)\",\"DIFFERENCE vs C2 - investigate\")")
    X.name('Capex_Climate', E7, rr + 1, 2)
    X.name('Capex_Climate_Pct', E7, rr + 1, 4)
    rr += 3
    X.sec(ws, rr, '(f) Internal carbon price - S2.29(f)', 7)
    X.put(ws, rr + 1, 0, 'Internal (shadow) carbon price', bold=True); X.put(ws, rr + 1, 1, 'Applied in appraisal of all capex above USD 100k and in scenario analysis (C3); not used for transfer pricing', font_color=C['inp'])
    X.put(ws, rr + 1, 2, 20, bg_color=C['yellow']); X.put(ws, rr + 1, 3, 'USD per tCO2e'); X.put(ws, rr + 1, 4, ''); X.put(ws, rr + 1, 5, 'S2.29(f)(i)-(ii)'); X.put(ws, rr + 1, 6, 'Board paper 22-May-2025')
    X.name('ICP', E7, rr + 1, 2)
    X.put(ws, rr + 2, 0, 'Implied annual cost of Scope 1+2 at internal price (USD\'000)'); X.put(ws, rr + 2, 1, 'Internal price x Scope 1+2 (E6) - memo only')
    X.put(ws, rr + 2, 2, '=ICP*GHG_S12/1000'); X.put(ws, rr + 2, 3, '=PBT25'); X.put(ws, rr + 2, 4, f'=C{rr + 3}/D{rr + 3}', num_format='0.0%'); X.put(ws, rr + 2, 5, 'S2.29(f)'); X.put(ws, rr + 2, 6, '')
    rr += 4
    X.sec(ws, rr, '(g) Remuneration - S2.29(g); see also S2.6(a)(v)', 7)
    X.put(ws, rr + 1, 0, 'Executive remuneration linked to climate-related considerations', bold=True)
    X.put(ws, rr + 1, 1, '10% of executive bonus pool tied to Scope 1+2 reduction (60%) and LTIFR (40%); climate-linked portion recognised in FY2025', font_color=C['inp'])
    X.put(ws, rr + 1, 2, 111); X.put(ws, rr + 1, 3, '=ExecRem25'); X.put(ws, rr + 1, 4, f'=C{rr + 2}/D{rr + 2}', num_format='0.0%', bold=True)
    X.put(ws, rr + 1, 5, 'S2.29(g)(i)-(ii)'); X.put(ws, rr + 1, 6, 'Remuneration committee paper; FS note 27')
    X.name('Rem_Pct', E7, rr + 1, 4)
    X.note(ws, rr + 3, 'S2.30: in preparing (b)-(d) the entity uses all reasonable and supportable information available at the reporting date without undue cost or effort. Amounts are single-counted within each metric but an asset can be exposed to more than one risk. Definitions are entity-specific judgements - see F2 J-03.', 7, h=30)
    ws.freeze_panes(9, 1)


def build_e8(X):
    ws, r = X.sheet(E8, 'E8', 'Industry-based and entity-specific metrics',
                    'Document industry-based metrics considered (IFRS S2 industry-based guidance and SASB Standard for Apparel, Accessories & Footwear), their applicability, and other metrics the entity uses to measure and monitor each relevant SRO.',
                    'IFRS S2.32, B65(d); IFRS S1.45-50, 53, 55(a), 58(a); Industry-based Guidance on Implementing IFRS S2 (Vol. Apparel, Accessories & Footwear)',
                    'HB 3.5.10, 4.1.190; DC 2.4.3, 2.4.5, 3.4',
                    [15, 26, 46, 14, 13, 13, 11, 16, 34], prep='Chan M.', tab=C['green'], status='In review')
    e = X.ad
    t = e['e1t']
    X.sec(ws, r, '1. Industry-based metrics considered (verify codes against the current SASB / IBG version before publication)', 9)
    H = ['Code', 'Disclosure topic', 'Metric', 'FY2025', 'Unit', 'Applicable?', 'Linked SRO', 'ISSB reference', 'Source / comment']
    R = [('CG-AA-440a.3', 'Raw materials sourcing (IFRS S2 IBG)', 'Priority raw materials; environmental / social factors threatening sourcing; business risks; management strategy', 'Narrative', 'n/a', 'Yes', 'CR-07', 'S2.32; IBG CG-AA', 'Cotton and polyester fabric; narrative in C1'),
         ('CG-AA-440a.4', 'Raw materials sourcing (IFRS S2 IBG)', '(1) Amount of priority raw materials purchased - fabric', '=' + e['e1_fabric'], 'tonnes', 'Yes', 'CR-07', 'S2.32; IBG CG-AA', 'E1'),
         ('CG-AA-440a.4', 'Raw materials sourcing (IFRS S2 IBG)', '(2) Share certified to a third-party standard - cotton (BCI / organic)', '=' + e['e1_cotton_cert'], '%', 'Yes', 'CR-07', 'S2.32; IBG CG-AA', 'Supplier certificates'),
         ('CG-AA-440a.4', 'Raw materials sourcing (IFRS S2 IBG)', '(2) Share certified - recycled polyester (GRS)', '=' + e['e1_rpet'], '%', 'Yes', 'CO-02', 'S2.32; IBG CG-AA', 'Transaction certificates'),
         ('CG-AA-000.A', 'Activity metric', 'Number of (1) tier-1 suppliers', '=' + e['e1_tier1'], 'number', 'Yes', 'All', 'S2.32; S1.48', 'Approved supplier list'),
         ('CG-AA-000.A', 'Activity metric', 'Number of (2) suppliers beyond tier 1', '=' + e['e1_tier2'], 'number', 'Yes', 'All', 'S2.32; S1.48', 'Traceability system (partial mapping)'),
         ('CG-AA-430a.1', 'Environmental impacts in the supply chain (SASB)', 'Wastewater discharge conformance - own wash plant (adapted: metric is designed for brands\' supplier facilities)', '=' + e['e1_ww_conf'], '%', 'Yes', 'SR-03', 'S1.55(a), 58(a), 50', 'Adapted metric - differences disclosed per S1.50(a)'),
         ('CG-AA-430a.2', 'Environmental impacts in the supply chain (SASB)', 'Facilities that completed the Higg FEM (or equivalent) assessment', 1.0, '%', 'Yes', 'CR-03', 'S1.55(a), 58(a)', '3 of 3 facilities verified'),
         ('CG-AA-430b.1', 'Labour conditions in the supply chain (SASB)', 'Facilities audited to a labour code of conduct (own facilities)', 1.0, '%', 'Yes', 'SR-01', 'S1.55(a), 58(a)', 'BFC + buyer audits'),
         ('CG-AA-430b.2', 'Labour conditions in the supply chain (SASB)', 'Buyer social audits passed without critical finding', '=' + e['e1_audit_pass'], '%', 'Yes', 'SR-01', 'S1.55(a), 58(a), 50', 'Adapted metric'),
         ('CG-AA-430b.3', 'Labour conditions in the supply chain (SASB)', 'Greatest labour and environmental, health and safety risks - description', 'Narrative', 'n/a', 'Yes', 'SR-01, SR-02', 'S1.55(a)', 'Draft narrative s.5'),
         ('CG-AA-250a.1 / .2', 'Management of chemicals in products (SASB)', 'Processes to maintain compliance with restricted-substances regulations; chemical risk assessment', 'Narrative', 'n/a', 'Yes', 'SR-03', 'S1.55(a)', 'ZDHC MRSL conformance at wash plant'),
         ('FB / other', 'Other IBG volumes', 'No other industry activities', 'n/a', 'n/a', 'No', 'n/a', 'S2.32; HB Ex. 2.2.100', 'Single-industry entity')]
    a, b = X.table(ws, r + 1, H, R, inp=(1, 2, 4, 5, 6, 8))
    for i, x in enumerate(R):
        if x[4] == '%':
            X.put(ws, a + i, 3, x[3], num_format='0.0%')
    X.cf_status(ws, a, b, 5)
    r = b + 2
    X.sec(ws, r, '2. Entity-specific metrics used to measure and monitor SROs (S1.46(b), S1.50, S2.B67) - no FY2024 comparatives for non-climate metrics (relief S1.E6)', 9)
    prod = f"({ref(E1, t, 11)}+{ref(E1, t, 12)})"
    dsl = f"({ref(E1, t, 5)}+{ref(E1, t, 6)}+{ref(E1, t, 7)}+{e['e1_veh_dsl']})"
    wood = f"({ref(E1, t, 8)}+{ref(E1, t, 9)}+{ref(E1, t, 10)})"
    water = f"({ref(E1, t, 13)}+{ref(E1, t, 14)}+{ref(E1, t, 15)})"
    H = ['Metric ID', 'Metric', 'Definition and method (S1.50(a), (d))', 'FY2025', 'Unit', 'Type (S1.50(b))', 'Linked SRO', 'Third-party validated? (S1.50(c))', 'Source']
    en = (f"=Elec_MWh*3.6+Solar_MWh*3.6+{dsl}*{e['ef_dsl_stat']['ncv']}/1000+{e['e1_veh_gas']}*{e['ef_gas_mob']['ncv']}/1000"
          f"+{e['e1_lpg']}*{e['ef_lpg']['ncv']}/1000+{wood}*{e['ef_wood']['ncv']}/1000")
    R = [('EM-01', 'Total energy consumed', 'Electricity (grid + solar) x 3.6 GJ/MWh plus fuels x net calorific value (E2)', en, 'GJ', 'Absolute', 'CR-05, CO-01', 'No', 'E1, E2'),
         ('EM-02', 'Renewable electricity share', 'Solar self-generation / total electricity consumed', '=RE_Share', '%', 'Ratio', 'CO-01', 'No', 'E4'),
         ('EM-03', 'Biomass share of total energy', 'Wood energy / total energy', lambda R: f"={wood}*{e['ef_wood']['ncv']}/1000/D{R - 2}", '%', 'Ratio', 'CR-04', 'No', 'E1'),
         ('EM-04', 'Energy intensity', 'Total energy / garments produced', lambda R: f"=D{R - 3}*1000/({prod}*1000)", 'MJ per garment', 'Ratio', 'CO-01', 'No', 'E1'),
         ('EM-05', 'Certified / agri-residue share of boiler biomass', 'Mass with chain-of-custody evidence / total biomass', '=' + e['e1_wood_cert'], '%', 'Ratio', 'CR-04', 'No', 'E1'),
         ('EM-06', 'Production days lost to climate-related events', 'Factory-days of stoppage due to flood or extreme heat', 9, 'days', 'Absolute', 'CR-01, CR-02', 'No', 'Production reports'),
         ('EM-07', 'Total water withdrawn', 'Metered groundwater and municipal supply, all sites', '=' + water, 'm3', 'Absolute', 'SR-03', 'No', 'E1'),
         ('EM-08', 'Wash-plant share of water withdrawn', 'JV water / total water', lambda R: f"={ref(E1, t, 15)}/D{R - 1}", '%', 'Ratio', 'SR-03', 'No', 'E1'),
         ('EM-09', 'Process water recycled - wash plant', 'Recycled volume / process water demand', '=' + e['e1_water_rec'], '%', 'Ratio', 'SR-03', 'No', 'ETP flow meters'),
         ('EM-10', 'Waste diverted from disposal', 'Recycled / (recycled + landfill + incinerated)', f"={e['e1_w_rec']}/({e['e1_w_rec']}+{e['e1_w_land']}+{e['e1_w_haz']})", '%', 'Ratio', 'SR-03', 'No', 'E1'),
         ('EM-11', 'Lost-time injury frequency rate (LTIFR)', 'Lost-time injuries x 1,000,000 / hours worked (GRI 403 definition)', f"={e['e1_lti']}*1000000/{e['e1_hours']}", 'per million hours', 'Ratio', 'SR-02', 'No', 'OHS register; payroll'),
         ('EM-12', 'Work-related fatalities', 'Employees and contractors', '=' + e['e1_fatal'], 'number', 'Absolute', 'SR-02', 'No', 'OHS register'),
         ('EM-13', 'Fainting incidents', 'Persons treated for fainting on site', '=' + e['e1_faint'], 'number', 'Absolute', 'SR-02, CR-02', 'No', 'Clinic log'),
         ('EM-14', 'BFC non-compliance points', 'Count of non-compliant items among critical assessment items', '=' + e['e1_bfc_nc'], 'number', 'Absolute', 'SR-01', 'Yes - ILO/IFC Better Factories Cambodia', 'BFC report'),
         ('EM-15', 'Female share of workforce', 'Female headcount / total headcount at year end', '=' + e['e1_female'], '%', 'Ratio', 'SR-01', 'No', 'HR system'),
         ('EM-16', 'Employee turnover', 'Leavers / average headcount', '=' + e['e1_turnover'], '%', 'Ratio', 'SR-04', 'No', 'HR system')]
    a, b = X.table(ws, r + 1, H, R, inp=(1, 2, 4, 5, 6, 7, 8))
    for i, x in enumerate(R):
        if x[4] == '%':
            X.put(ws, a + i, 3, x[3](a + i + 1) if callable(x[3]) else x[3], num_format='0.0%')
        elif x[0] in ('EM-11', 'EM-04'):
            X.put(ws, a + i, 3, x[3](a + i + 1) if callable(x[3]) else x[3], num_format='#,##0.00')
    X.name('LTIFR', E8, a + 10, 3)
    X.name('Energy_GJ', E8, a, 3)
    X.name('Water_m3', E8, a + 6, 3)
    ws.freeze_panes(9, 2)


def build_e9(X):
    ws, r = X.sheet(E9, 'E9', 'Targets - definition, approach and performance',
                    'Document each quantitative and qualitative target, how it was set and is reviewed, and performance against it, including the additional information required for GHG emissions targets.',
                    'IFRS S2.33-37, B66-B71; IFRS S1.51-53',
                    'HB 3.5.20, 4.1.210-4.1.240; DC 2.4.6',
                    [6, 34, 26, 16, 18, 8, 12, 8, 12, 24, 11, 14, 12, 10, 10, 11, 16, 14, 10, 10, 20, 16, 28, 28, 22], prep='Sok D.', tab=C['green'], status='In review')
    e = X.ad
    H = ['ID', 'Target', 'Metric used (S2.33(a))', 'Objective (S2.33(b))', 'Part of entity (S2.33(c))', 'Base year (S2.33(e))', 'Base value', 'Target year (S2.33(d))', 'Target value',
         'Milestones / interim targets (S2.33(f))', 'Absolute or intensity (S2.33(g))', 'Unit', 'FY2025 actual', 'Progress to target %', 'Expected % (linear path)', 'Status (S2.35)',
         'GHG covered (S2.36(a))', 'Scopes (S2.36(b))', 'Gross / net (S2.36(c))', 'Sectoral decarbonisation approach? (S2.36(d))', 'Carbon credits (S2.36(e))',
         'Third-party validation (S2.34(a))', 'How informed by latest international agreement (S2.33(h))', 'Review process and revisions (S2.34(b), (d))', 'Short name (dashboard label)']
    cat1 = ref(E5, X.ad['s3cat'][0], 3)
    T = [('T1', 'Reduce absolute Scope 1+2 GHG emissions by 42% by 2030', 'Scope 1 + Scope 2 (location-based) tCO2e, incl. JV', 'Mitigation; aligned with 1.5C cross-sector pathway', 'Entire entity incl. JV', 2023, 6900, 2030, lambda R: f'=G{R}*(1-0.42)',
          '-20% by 2027', 'Absolute', 'tCO2e', '=GHG_S12', 'CO2, CH4, N2O, HFCs', 'Scope 1 and 2', 'Gross', 'No - absolute contraction', 'None planned', 'Method reviewed by external adviser; SBTi submission planned 2026', 'Consistent with 1.5C goal of the Paris Agreement and Cambodia\'s updated NDC / LTS4CN', 'Annual review by ARSC; no revision in FY2025'),
         ('T2', 'Reach 30% renewable electricity by 2030', 'Renewable electricity / total electricity consumed', 'Mitigation', 'Entire entity incl. JV', 2023, 0.0, 2030, 0.30,
          '15% by 2027', 'Intensity (ratio)', '%', '=RE_Share', 'n/a', 'Supports Scope 2', 'n/a', 'n/a', 'n/a', 'No', 'Supports NDC energy targets', 'Annual review; no revision'),
         ('T3', 'Phase out uncertified forest wood in boilers by 2028', 'Certified / agri-residue share of boiler biomass', 'Mitigation; conformance with buyer deforestation-free policies', 'All boiler houses', 2023, 0.20, 2028, 1.0,
          '60% by 2026', 'Intensity (ratio)', '%', '=' + e['e1_wood_cert'], 'n/a', 'n/a', 'n/a', 'n/a', 'n/a', 'No', 'Forest-sector commitments in NDC', 'Timeline revised in FY2025 from 2027 to 2028 - pilot delay (disclose per S2.34(d))'),
         ('T4', 'Reduce Scope 3 category 1 intensity by 25% by 2030', 'tCO2e per tonne of fabric purchased', 'Mitigation (value chain)', 'Purchased fabric and trims', 2025, lambda R: f"={cat1}/{e['e1_fabric']}", 2030, lambda R: f'=G{R}*(1-0.25)',
          'Mill-specific data for 50% of volume by 2027', 'Intensity', 'tCO2e / t', lambda R: f"={cat1}/{e['e1_fabric']}", 'CO2e', 'Scope 3 (cat. 1)', 'Gross', 'No', 'None planned', 'No', 'Aligned with buyers\' science-based Scope 3 targets', 'Set in FY2025 (new target)'),
         ('T5', 'Protect 100% of critical assets at Factory B to 1-in-100 flood level by 2027', 'Share of critical assets protected', 'Adaptation', 'Factory B', 2024, 0.0, 2027, 1.0,
          '60% by 2025', 'Intensity (ratio)', '%', '=' + e['e1_flood_prot'], 'n/a', 'n/a', 'n/a', 'n/a', 'n/a', 'Engineer-certified', 'National Adaptation Plan priorities', 'Reviewed after Oct-2025 flood; no revision'),
         ('T6', 'Reduce LTIFR to 1.0 by 2027', 'Lost-time injury frequency rate', 'Worker safety (non-climate, S1.51)', 'Factories A and B', 2023, 2.6, 2027, 1.0,
          '1.8 by 2025', 'Intensity', 'per million hours', '=LTIFR', 'n/a', 'n/a', 'n/a', 'n/a', 'n/a', 'No', 'n/a (non-climate target)', 'Annual review by ARSC'),
         ('T7', 'Recycle 30% of wash-plant process water by 2027', 'Recycled share of process water', 'Water stewardship (non-climate, S1.51)', 'JV wash plant', 2024, 0.0, 2027, 0.30,
          '10% by 2025', 'Intensity (ratio)', '%', '=' + e['e1_water_rec'], 'n/a', 'n/a', 'n/a', 'n/a', 'n/a', 'No', 'n/a (non-climate target)', 'Annual review by JV board')]
    SHORT = {'T1': 'Scope 1+2 -42%', 'T2': 'Renewable power 30%', 'T3': 'Forest-wood phase-out', 'T4': 'Scope 3 cat.1 intensity', 'T5': 'Flood protection', 'T6': 'LTIFR 1.0', 'T7': 'Water recycling 30%'}
    rows = []
    for x in T:
        rows.append(list(x[:12]) + [x[12],
                                     lambda R: f'=IF(I{R}=G{R},0,MAX(0,MIN(1,(M{R}-G{R})/(I{R}-G{R}))))',
                                     lambda R: f'=IF(H{R}=F{R},0,MIN(1,(YEAR(PeriodEnd)-F{R})/(H{R}-F{R})))',
                                     lambda R: f'=IF(N{R}>=1,"Achieved",IF(N{R}>=O{R},"On track","Behind"))'] + list(x[13:]) + [SHORT[x[0]]])
    a, b = X.table(ws, r, H, rows, nf={5: '0', 7: '0', 13: '0.0%', 14: '0.0%'}, inp=tuple(range(1, 5)) + (9, 10, 11) + tuple(range(16, 25)))
    ws.set_row(r, 56)
    for i, x in enumerate(T):
        pct = x[11] == '%'
        nf = '0.0%' if pct else ('#,##0.00' if x[0] in ('T4', 'T6') else '#,##0')
        for j in (6, 8, 12):
            v = x[j]
            X.put(ws, a + i, j, v(a + i + 1) if callable(v) else v, num_format=nf)
    X.cf_status(ws, a, b, 15)
    X.ad['tgt'] = (a, b)
    X.put(ws, b + 2, 1, 'Targets on track or achieved', bold=True)
    X.put(ws, b + 2, 2, f'=COUNTIF({A(a, 15)}:{A(b, 15)},"On track")+COUNTIF({A(a, 15)}:{A(b, 15)},"Achieved")', bold=True, align='left')
    X.put(ws, b + 3, 1, 'Targets behind', bold=True)
    X.put(ws, b + 3, 2, f'=COUNTIF({A(a, 15)}:{A(b, 15)},"Behind")', bold=True, align='left')
    X.put(ws, b + 4, 1, 'T1 check: abatement planned in transition plan (C1) vs reduction still required (tCO2e)', bold=True)
    X.put(ws, b + 4, 2, f'=IF(TP_Abate>=GHG_S12-{A(a, 8)},"Plan covers the gap","Plan does NOT cover the gap")', bold=True, align='left')
    X.name('Tgt_OnTrack', E9, b + 2, 2)
    X.name('Tgt_Behind', E9, b + 3, 2)
    X.note(ws, b + 6, 'Progress % = (actual - base) / (target - base). Expected % assumes a straight-line path from base year to target year - a simplification used for monitoring only. Net-zero or net targets: none; no reliance on carbon credits (S2.36(c), (e), B68-B71).', 12, h=30)
    ws.freeze_panes(9, 2)


def build_part2(X):
    build_e1(X); build_e2(X); build_e3(X); build_e4(X); build_e5(X); build_e6(X); build_e7(X); build_e8(X); build_e9(X)
