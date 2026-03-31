from __future__ import annotations

import hashlib
import json
import math
import random
from collections import OrderedDict
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import xlsxwriter

OUTPUT_ROOT = Path("output")
INPUT_CSV = OUTPUT_ROOT / "matrix_500.csv"
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
OUT_JSON = OUTPUT_ROOT / 'charts.json'
OUT_XLSX = OUTPUT_ROOT / 'charts.xlsx'
OUT_MANIFEST = OUTPUT_ROOT / 'manifest.json'

STYLE_REF = {
    'id': 'office-custom-theme-pendent-v1',
    'href': './styles/office-custom-theme-pendent-v1.json',
    'description': 'Reference pending the final corporate style definition.'
}

MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4']
YEARS = [str(y) for y in range(1995, 2026)]
REGIONS_CAT = ['High Pyrenees', 'Western Plains', 'Tarragona Camp', 'Metropolitan Barcelona', 'Girona', 'Ebro Lands']

# Accessible visual constants prepared for future pattern-based rendering.
ACCESSIBLE_COLOR_PALETTE: List[str] = [
    '#4472C4',
    '#C0504D',
    '#9BBB59',
    '#8064A2',
    '#4BACC6',
    '#F79646',
    '#7F7F7F',
    '#1F4E78',
]

ACCESSIBLE_PATTERN_PALETTE: List[Dict[str, str]] = [
    {'id': 'P1', 'name': 'diagonal-right', 'token': 'slash', 'description': 'Right-leaning diagonal lines'},
    {'id': 'P2', 'name': 'diagonal-left', 'token': 'backslash', 'description': 'Left-leaning diagonal lines'},
    {'id': 'P3', 'name': 'vertical', 'token': 'vertical', 'description': 'Vertical lines'},
    {'id': 'P4', 'name': 'horizontal', 'token': 'horizontal', 'description': 'Horizontal lines'},
    {'id': 'P5', 'name': 'dots', 'token': 'dots', 'description': 'Separated dots'},
    {'id': 'P6', 'name': 'cross', 'token': 'cross', 'description': 'Simple cross'},
    {'id': 'P7', 'name': 'diag-cross', 'token': 'diag-cross', 'description': 'Diagonal cross'},
    {'id': 'P8', 'name': 'grid', 'token': 'grid', 'description': 'Grid pattern'},
]

ACCESSIBLE_SERIES_STYLE_TABLE: List[Dict[str, str]] = [
    {'series_slot': '1', 'color': '#4472C4', 'pattern_id': 'P1'},
    {'series_slot': '2', 'color': '#C0504D', 'pattern_id': 'P2'},
    {'series_slot': '3', 'color': '#9BBB59', 'pattern_id': 'P3'},
    {'series_slot': '4', 'color': '#8064A2', 'pattern_id': 'P4'},
    {'series_slot': '5', 'color': '#4BACC6', 'pattern_id': 'P5'},
    {'series_slot': '6', 'color': '#F79646', 'pattern_id': 'P6'},
    {'series_slot': '7', 'color': '#7F7F7F', 'pattern_id': 'P7'},
    {'series_slot': '8', 'color': '#1F4E78', 'pattern_id': 'P8'},
]

ACCESSIBLE_LINE_MARKER_TABLE: List[Dict[str, str]] = [
    {'series_slot': '1', 'dash': 'solid', 'marker': 'circle'},
    {'series_slot': '2', 'dash': 'dash', 'marker': 'square'},
    {'series_slot': '3', 'dash': 'dot', 'marker': 'diamond'},
    {'series_slot': '4', 'dash': 'dash_dot', 'marker': 'triangle'},
    {'series_slot': '5', 'dash': 'short_dash', 'marker': 'x'},
    {'series_slot': '6', 'dash': 'long_dash', 'marker': 'circle'},
]

ACCESSIBLE_STYLE_POLICY: Dict[str, Dict[str, str]] = {
    'column': {'mode': 'pattern-per-series', 'legend': 'pattern+color'},
    'bar': {'mode': 'pattern-per-series', 'legend': 'pattern+color'},
    'area': {'mode': 'pattern-per-series-light', 'legend': 'pattern+color'},
    'pie': {'mode': 'pattern-per-point', 'legend': 'pattern+color'},
    'doughnut': {'mode': 'pattern-per-point', 'legend': 'pattern+color'},
    'radar': {'mode': 'pattern-per-series-light', 'legend': 'pattern+color'},
    'combo-column': {'mode': 'pattern-per-series', 'legend': 'pattern+color'},
    'combo-line': {'mode': 'line-marker', 'legend': 'line+marker'},
    'line': {'mode': 'line-marker', 'legend': 'line+marker'},
    'scatter': {'mode': 'marker-only', 'legend': 'marker+color'},
    'stock': {'mode': 'line-marker', 'legend': 'line+marker'},
    'histogram': {'mode': 'single-pattern', 'legend': 'pattern+color'},
    'funnel': {'mode': 'pattern-per-point', 'legend': 'pattern+color'},
    'surface': {'mode': 'contour-scale', 'legend': 'scale'},
}

ACCESSIBLE_PATTERN_TOKEN_TO_XLSXWRITER: Dict[str, str] = {
    'slash': 'light_upward_diagonal',
    'backslash': 'light_downward_diagonal',
    'vertical': 'light_vertical',
    'horizontal': 'light_horizontal',
    'dots': 'dotted_grid',
    'cross': 'small_grid',
    'diag-cross': 'trellis',
    'grid': 'large_grid',
}

ENABLE_CHART_PATTERNS = False

DOMAIN_CFG: Dict[str, Dict[str, Any]] = {
    'climate': {
        'unit_abs': 'ktCO₂e',
        'unit_pct': '% contribution',
        'simple_categories': ['Transport', 'Buildings', 'Industry', 'Waste', 'Agriculture', 'Energy', 'Aviation', 'Livestock', 'Land use', 'Forestation', 'Industrial processes', 'Fisheries'],
        'long_categories': [
            'Urban and interurban passenger transport',
            'Aging residential building stock',
            'High-consumption industrial estates',
            'Mixed municipal waste management',
            'Local agri-food activity',
            'Heavy freight vehicle fleet',
            'Heating and cooling in tertiary and retail spaces',
            'Fossil-fuel-based electricity generation',
            'Extensive livestock farming and slurry production',
            'Water extraction and industrial water use'
        ],
        'geo_categories': ['Pyrenees', 'Lleida Plain', 'Tarragona Camp', 'Barcelona Metropolitan Area', 'Girona Counties', 'Ebro Lands', 'Central Catalonia', 'Alt Emporda'],
        'series': ['CO₂', 'Methane', 'Nitrous oxide'],
        'radar_axes': ['Mitigation', 'Adaptation', 'Exposure', 'Resilience', 'Efficiency', 'Monitoring', 'Governance', 'Social impact'],
        'scatter_x': ('Average temperature (°C)', 10, 28),
        'scatter_y': ('Per-capita emissions (tCO₂e)', 2, 14),
        'map_metric': 'Per-capita emissions (tCO₂e)',
        'combo_primary': 'Monthly emissions (ktCO₂e)',
        'combo_secondary': 'Temperature anomaly (°C)',
        'stock_label': 'Electricity futures price (€/MWh)'
    },
    'demographics': {
        'unit_abs': 'thousands of people',
        'unit_pct': '% of total',
        'simple_categories': ['Children', 'Youth', 'Adults', 'Seniors', 'Dependency', 'Migration', 'Housing', 'Independent living', 'Diversity', 'New citizens', 'Families', 'Loneliness'],
        'long_categories': [
            'Population aged 0 to 14',
            'Population aged 15 to 29',
            'Population aged 30 to 44',
            'Population aged 45 to 64',
            'Population aged 65 or older',
            'People born outside the European Union',
            'Single-person households in consolidated urban areas',
            'Young adults in delayed emancipation',
            'People with recognized disabilities',
            'Children in single-parent households'
        ],
        'geo_categories': ['Aran Valley', 'Western Plains', 'Central Counties', 'Barcelona', 'Tarragona Camp', 'Girona', 'Ebro Lands', 'Pyrenees'],
        'series': ['Men', 'Women', 'Unspecified'],
        'radar_axes': ['Growth', 'Aging', 'Density', 'Mobility', 'Birth rate', 'Attractiveness', 'Territorial balance', 'Generational renewal'],
        'scatter_x': ('Density (inhabitants/km²)', 30, 600),
        'scatter_y': ('Annual growth (%)', -1.0, 3.0),
        'map_metric': 'Population density (inhabitants/km²)',
        'combo_primary': 'Monthly births',
        'combo_secondary': 'Migration balance (%)',
        'stock_label': 'Housing price index'
    },
    'education': {
        'unit_abs': 'thousands of students',
        'unit_pct': '% of students',
        'simple_categories': ['Primary', 'Lower secondary', 'Upper secondary', 'Vocational training', 'University', 'Early childhood', 'Higher vocational training', 'Doctorate', 'Adults', 'Special education', 'Languages', 'Basic computing'],
        'long_categories': [
            'Primary students in public schools',
            'Lower-secondary students in vulnerable settings',
            'Science and technology upper-secondary students',
            'Students in dual vocational training',
            'First-year university students',
            'Children enrolled in 0-3 early-childhood centers',
            'Students in higher vocational cycles',
            'Doctoral students in Catalan public universities',
            'Adults in reskilling programs',
            'Students with specific educational needs'
        ],
        'geo_categories': ['Girona', 'Barcelona', 'Lleida', 'Tarragona', 'Ebro Lands', 'Pyrenees', 'High Pyrenees', 'Central Counties'],
        'series': ['Public', 'State-funded private', 'Private'],
        'radar_axes': ['Mathematics', 'Reading', 'Science', 'Digital skills', 'Languages', 'Participation', 'Equity', 'Academic continuity'],
        'scatter_x': ('Student-to-teacher ratio', 8, 32),
        'scatter_y': ('Average score', 50, 95),
        'map_metric': 'Early school leaving (%)',
        'combo_primary': 'Monthly enrollments',
        'combo_secondary': 'Coverage ratio (%)',
        'stock_label': 'Education cost index'
    },
    'energy': {
        'unit_abs': 'GWh',
        'unit_pct': '% of generation',
        'simple_categories': ['Nuclear', 'Hydropower', 'Wind', 'Solar', 'Gas', 'Biomass', 'Marine', 'Geothermal', 'Cogeneration', 'Waste', 'Hydrogen', 'Smart grid'],
        'long_categories': [
            'Asco nuclear power station',
            'Annual regulation reservoirs',
            'Terra Alta wind farms',
            'Segria photovoltaic facilities',
            'Support combined-cycle plants',
            'Agro-industrial biogas plants',
            'Experimental marine energy platforms',
            'Deep geothermal heat pumps',
            'Efficient industrial cogeneration units',
            'Energy recovery from municipal solid waste'
        ],
        'geo_categories': ['Pyrenees', 'Western districts', 'Tarragona Camp', 'Metropolitan Area', 'Central Catalonia', 'Ebro', 'High Pyrenees', 'Girona Counties'],
        'series': ['Baseload', 'Renewable', 'Support'],
        'radar_axes': ['Availability', 'Cost', 'Emissions', 'Flexibility', 'Risk', 'Coverage', 'Autonomy', 'Grid stability'],
        'scatter_x': ('Installed capacity (MW)', 20, 1200),
        'scatter_y': ('Annual production (GWh)', 10, 6500),
        'map_metric': 'Renewable generation per resident (MWh)',
        'combo_primary': 'Monthly generation (GWh)',
        'combo_secondary': 'Pool price (€/MWh)',
        'stock_label': 'Gas futures price (€/MWh)'
    },
    'finance': {
        'unit_abs': 'M€',
        'unit_pct': '% of total',
        'simple_categories': ['Revenue', 'Costs', 'Margin', 'Capex', 'Debt', 'Dividends', 'Taxes', 'Reserves', 'Depreciation', 'Working capital', 'Cash', 'Equity'],
        'long_categories': [
            'Recurring subscription revenue',
            'Commercial and acquisition cost',
            'Operating margin before depreciation',
            'Investment in technology platform',
            'Adjusted net financial debt',
            'Dividends distributed to shareholders',
            'Effective tax burden on earnings',
            'Reserves for future contingencies',
            'Depreciation of key intangible assets',
            'Working capital and cash position'
        ],
        'geo_categories': ['Northeast', 'Central', 'East', 'South', 'Portugal', 'International', 'Asia-Pacific', 'Latin America'],
        'series': ['Actual', 'Budget', 'Target'],
        'radar_axes': ['Liquidity', 'Profitability', 'Growth', 'Solvency', 'Efficiency', 'Conversion', 'Leverage', 'Forward visibility'],
        'scatter_x': ('Risk (volatility %)', 5, 40),
        'scatter_y': ('Annual return (%)', -10, 35),
        'map_metric': 'Revenue by province (M€)',
        'combo_primary': 'Monthly revenue (M€)',
        'combo_secondary': 'EBITDA margin (%)',
        'stock_label': 'Share price (€)'
    },
    'geography': {
        'unit_abs': 'territorial index',
        'unit_pct': '% of area',
        'simple_categories': ['Coast', 'Plain', 'Pre-coastal', 'Mountain', 'Urban', 'Islands', 'River', 'Dryland', 'Wetlands', 'Forest', 'Cropland', 'Delta'],
        'long_categories': [
            'High-density consolidated urban land',
            'Intensively irrigated agricultural land',
            'Mid-mountain forest areas',
            'Tourism-oriented coastal corridors',
            'Highly fragile protected natural areas',
            'Wetlands of special ecological value',
            'Expanding peri-urban industrial estates',
            'Protected priority agricultural areas',
            'Integrated river-management areas',
            'Rural landscapes with recognized heritage value'
        ],
        'geo_categories': ['Aran Valley', 'Pallars', 'Segria', 'Barcelones', 'Emporda', 'Ebro Delta', 'Barbera Basin', 'Lake Plain'],
        'series': ['Land cover', 'Urban use', 'Protected'],
        'radar_axes': ['Accessibility', 'Slopes', 'Density', 'Facilities', 'Risk', 'Attractiveness', 'Connectivity', 'Urban pressure'],
        'scatter_x': ('Average altitude (m)', 0, 2200),
        'scatter_y': ('Population density (inhabitants/km²)', 5, 1500),
        'map_metric': 'Population density (inhabitants/km²)',
        'combo_primary': 'Monthly tourist flow (thousands)',
        'combo_secondary': 'Occupancy (%)',
        'stock_label': 'Territorial index'
    },
    'manufacturing': {
        'unit_abs': 'thousands of units',
        'unit_pct': '% of production',
        'simple_categories': ['Line A', 'Line B', 'Line C', 'Line D', 'Rework', 'Quality', 'Packaging', 'Painting', 'Testing', 'Assembly', 'Accessories', 'Maintenance'],
        'long_categories': [
            'Fine-electronics assembly line',
            'Critical-parts machining line',
            'Robotic surface-finishing cell',
            'Validation and dimensional-control area',
            'Rework and non-conformity circuit',
            'Packaging and shipment-preparation section',
            'Electrostatic powder-coating booth',
            'Testing bench and functional validation',
            'Final-equipment assembly area',
            'Preventive and predictive maintenance workshop'
        ],
        'geo_categories': ['North Plant', 'Central Plant', 'South Plant', 'Technical warehouse', 'Laboratory', 'Subcontractor', 'East Plant', 'Logistics platform'],
        'series': ['Accepted', 'Reworked', 'Rejected'],
        'radar_axes': ['Quality', 'Productivity', 'Safety', 'Maintenance', 'Cost', 'Flexibility', 'Traceability', 'Robustness'],
        'scatter_x': ('Cycle time (s)', 20, 220),
        'scatter_y': ('Defects per million', 30, 1800),
        'map_metric': 'Production by plant (thousands of units)',
        'combo_primary': 'Monthly production (thousands)',
        'combo_secondary': 'OEE (%)',
        'stock_label': 'Industrial cost index'
    },
    'operations': {
        'unit_abs': 'thousands of orders',
        'unit_pct': '% of activity',
        'simple_categories': ['Receiving', 'Picking', 'Dispatch', 'Quality', 'Returns', 'Supply', 'Customer care', 'Billing', 'Planning', 'Transport', 'Warehouse', 'Reverse logistics'],
        'long_categories': [
            'Receiving supplier merchandise',
            'E-commerce order picking',
            'Dispatch to physical stores',
            'Outbound quality control',
            'Returns and incident handling',
            'Procurement of strategic materials',
            'Multichannel post-sales customer care',
            'Order billing and reconciliation',
            'Monthly demand planning',
            'Long-haul transport management'
        ],
        'geo_categories': ['Barcelona hub', 'Girona hub', 'Lleida hub', 'Tarragona hub', 'Cross-dock', 'Last mile', 'Distribution center', 'Logistics platform'],
        'series': ['Planned', 'Executed', 'Incidents'],
        'radar_axes': ['Service', 'Time', 'Cost', 'Capacity', 'Quality', 'Flexibility', 'On-time performance', 'Scalability'],
        'scatter_x': ('Preparation time (min)', 5, 120),
        'scatter_y': ('Orders/hour', 20, 400),
        'map_metric': 'On-time deliveries (%)',
        'combo_primary': 'Monthly orders (thousands)',
        'combo_secondary': 'On-time performance (%)',
        'stock_label': 'Logistics cost index'
    },
    'health': {
        'unit_abs': 'thousands of patients',
        'unit_pct': '% of cases',
        'simple_categories': ['Primary care', 'Emergency', 'Hospitalization', 'Diagnosis', 'Rehabilitation', 'Mental health', 'Pharmacy', 'Oncology', 'Pediatrics', 'Geriatrics', 'Surgery', 'Cardiology'],
        'long_categories': [
            'Patients under post-operative remote monitoring',
            'Cases handled through the rapid-diagnosis pathway',
            'High-complexity internal medicine admissions',
            'Community cardiovascular prevention program',
            'Intensive functional rehabilitation unit',
            'Primary-care mental-health consultations',
            'Pharmaceutical dispensing of chronic medication',
            'Patients in active oncology treatment',
            'Pediatric visits in primary-care centers',
            'Geriatric care in residences and at home'
        ],
        'geo_categories': ['Barcelona', 'Girona', 'Lleida', 'Tarragona', 'Central Catalonia', 'Ebro Lands', 'High Pyrenees', 'Western Valles'],
        'series': ['Resolved', 'Under follow-up', 'Referred'],
        'radar_axes': ['Accessibility', 'Quality', 'Waiting time', 'Resolution', 'Satisfaction', 'Prevention', 'Continuity of care', 'Coverage'],
        'scatter_x': ('Waiting days', 2, 180),
        'scatter_y': ('Resolved cases (%)', 40, 98),
        'map_metric': 'Average waiting list (days)',
        'combo_primary': 'Monthly visits (thousands)',
        'combo_secondary': 'Average waiting time (days)',
        'stock_label': 'Healthcare cost index'
    },
    'sales': {
        'unit_abs': 'M€',
        'unit_pct': '% of sales',
        'simple_categories': ['Physical store', 'E-commerce', 'Distributor', 'Marketplace', 'Phone channel', 'Direct-to-consumer', 'Franchise', 'Subscription', 'Export', 'Tender', 'Wholesaler', 'Flash sale'],
        'long_categories': [
            'Premium home appliances',
            'Connected accessories for household use',
            'Efficient small kitchen appliances',
            'Maintenance and extended-warranty service',
            'Professional installer channel',
            'Direct online sales to end customers',
            'Franchise operators in the regional network',
            'Recurring-service subscription model',
            'Exports to emerging European markets',
            'Public tenders for large contracts'
        ],
        'geo_categories': ['North', 'Central', 'East', 'South', 'Balearic Islands', 'Portugal', 'Canary Islands', 'International'],
        'series': ['Nord', 'Centre', 'Sud'],
        'series': ['North', 'Central', 'South'],
        'radar_axes': ['Volume', 'Margin', 'Turnover', 'Loyalty', 'Average ticket', 'Conversion', 'Penetration', 'Recurrence'],
        'scatter_x': ('Average discount (%)', 0, 35),
        'scatter_y': ('Sales per outlet (k€)', 20, 500),
        'map_metric': 'Sales per resident (€)',
        'combo_primary': 'Monthly revenue (M€)',
        'combo_secondary': 'Gross margin (%)',
        'stock_label': 'Retail price index'
    },
    'web analytics': {
        'unit_abs': 'thousands of sessions',
        'unit_pct': '% of sessions',
        'simple_categories': ['Direct', 'Organic', 'Social', 'Email', 'Referral', 'Paid Search', 'Affiliate', 'Display', 'Push', 'App', 'QR', 'Video'],
        'long_categories': [
            'B2B lead-generation search campaign',
            'Loyalty program through segmented email',
            'Seasonal promotion on social networks',
            'Branded traffic from external comparison sites',
            'Recurring visits from the mobile app',
            'Display campaign across content networks',
            'Paid-advertising landing pages',
            'Traffic from affiliate blog articles',
            'Push notifications for returning users',
            'Traffic from QR codes in physical advertising'
        ],
        'geo_categories': ['SEO', 'SEM', 'CRM', 'Social', 'Affiliate', 'App', 'Display', 'Referral'],
        'series': ['Low bounce', 'Average browsing', 'Conversion'],
        'radar_axes': ['Acquisition', 'Conversion', 'Retention', 'Depth', 'Speed', 'ROI', 'Engagement', 'Traffic quality'],
        'scatter_x': ('Load time (s)', 0.5, 5.0),
        'scatter_y': ('Conversion (%)', 0.2, 8.0),
        'map_metric': 'Sessions by territory (thousands)',
        'combo_primary': 'Monthly sessions (thousands)',
        'combo_secondary': 'Conversion (%)',
        'stock_label': 'CPC index (€)'
    },
}

DOMAIN_VALUE_RANGE_CFG: Dict[str, Dict[str, Dict[str, List[Tuple[float, float]]]]] = {
    'climate': {
        'category_value_ranges': {
            'low': [(110.0, 240.0), (140.0, 280.0)],
            'medium': [(70.0, 180.0), (90.0, 210.0)],
            'alta': [(40.0, 120.0), (55.0, 145.0)],
            'very high': [(22.0, 80.0), (30.0, 95.0)],
        },
        'map_value_ranges': {
            'low': [(6.5, 12.0), (5.5, 10.5)],
            'medium': [(4.5, 9.0), (3.8, 7.8)],
            'alta': [(3.0, 6.5), (2.6, 5.5)],
            'very high': [(2.0, 4.8), (1.8, 4.0)],
        },
    },
    'demographics': {
        'category_value_ranges': {
            'low': [(220.0, 950.0), (180.0, 820.0)],
            'medium': [(140.0, 700.0), (110.0, 580.0)],
            'alta': [(90.0, 460.0), (70.0, 360.0)],
            'very high': [(45.0, 240.0), (35.0, 190.0)],
        },
        'map_value_ranges': {
            'low': [(220.0, 620.0), (180.0, 520.0)],
            'medium': [(140.0, 420.0), (110.0, 340.0)],
            'alta': [(80.0, 260.0), (60.0, 220.0)],
            'very high': [(35.0, 160.0), (25.0, 120.0)],
        },
    },
    'education': {
        'category_value_ranges': {
            'low': [(80.0, 360.0), (60.0, 280.0)],
            'medium': [(50.0, 240.0), (35.0, 190.0)],
            'alta': [(25.0, 130.0), (18.0, 100.0)],
            'very high': [(10.0, 70.0), (8.0, 55.0)],
        },
        'map_value_ranges': {
            'low': [(14.0, 28.0), (11.0, 24.0)],
            'medium': [(9.0, 21.0), (7.0, 17.0)],
            'alta': [(6.0, 15.0), (4.5, 12.0)],
            'very high': [(3.0, 9.0), (2.5, 7.5)],
        },
    },
    'energy': {
        'category_value_ranges': {
            'low': [(600.0, 4200.0), (450.0, 3200.0)],
            'medium': [(300.0, 2500.0), (220.0, 1800.0)],
            'alta': [(140.0, 1200.0), (100.0, 900.0)],
            'very high': [(60.0, 520.0), (40.0, 380.0)],
        },
        'map_value_ranges': {
            'low': [(2.0, 8.5), (1.5, 6.5)],
            'medium': [(1.2, 5.8), (0.9, 4.5)],
            'alta': [(0.7, 3.8), (0.5, 2.8)],
            'very high': [(0.3, 1.8), (0.2, 1.3)],
        },
    },
    'finance': {
        'category_value_ranges': {
            'low': [(40.0, 240.0), (30.0, 180.0)],
            'medium': [(22.0, 150.0), (16.0, 110.0)],
            'alta': [(10.0, 85.0), (8.0, 65.0)],
            'very high': [(4.0, 38.0), (3.0, 28.0)],
        },
        'map_value_ranges': {
            'low': [(60.0, 380.0), (45.0, 300.0)],
            'medium': [(35.0, 220.0), (25.0, 170.0)],
            'alta': [(18.0, 120.0), (12.0, 90.0)],
            'very high': [(7.0, 55.0), (5.0, 40.0)],
        },
    },
    'geography': {
        'category_value_ranges': {
            'low': [(45.0, 95.0), (35.0, 85.0)],
            'medium': [(28.0, 78.0), (22.0, 68.0)],
            'alta': [(16.0, 58.0), (12.0, 48.0)],
            'very high': [(8.0, 34.0), (6.0, 26.0)],
        },
        'map_value_ranges': {
            'low': [(250.0, 1500.0), (180.0, 1200.0)],
            'medium': [(120.0, 900.0), (90.0, 700.0)],
            'alta': [(55.0, 420.0), (40.0, 320.0)],
            'very high': [(12.0, 180.0), (8.0, 130.0)],
        },
    },
    'manufacturing': {
        'category_value_ranges': {
            'low': [(40.0, 240.0), (30.0, 180.0)],
            'medium': [(24.0, 155.0), (18.0, 120.0)],
            'alta': [(12.0, 82.0), (9.0, 64.0)],
            'very high': [(5.0, 34.0), (4.0, 26.0)],
        },
        'map_value_ranges': {
            'low': [(50.0, 260.0), (40.0, 210.0)],
            'medium': [(28.0, 170.0), (22.0, 130.0)],
            'alta': [(14.0, 90.0), (10.0, 70.0)],
            'very high': [(6.0, 40.0), (4.0, 28.0)],
        },
    },
    'operations': {
        'category_value_ranges': {
            'low': [(24.0, 180.0), (18.0, 140.0)],
            'medium': [(14.0, 110.0), (10.0, 82.0)],
            'alta': [(7.0, 56.0), (5.0, 42.0)],
            'very high': [(3.0, 24.0), (2.0, 18.0)],
        },
        'map_value_ranges': {
            'low': [(90.0, 99.0), (86.0, 97.0)],
            'medium': [(82.0, 96.0), (78.0, 94.0)],
            'alta': [(74.0, 92.0), (70.0, 88.0)],
            'very high': [(66.0, 84.0), (62.0, 80.0)],
        },
    },
    'health': {
        'category_value_ranges': {
            'low': [(60.0, 380.0), (45.0, 290.0)],
            'medium': [(35.0, 220.0), (25.0, 170.0)],
            'alta': [(18.0, 120.0), (12.0, 90.0)],
            'very high': [(7.0, 55.0), (5.0, 42.0)],
        },
        'map_value_ranges': {
            'low': [(35.0, 140.0), (28.0, 120.0)],
            'medium': [(22.0, 100.0), (16.0, 80.0)],
            'alta': [(12.0, 60.0), (9.0, 46.0)],
            'very high': [(5.0, 28.0), (4.0, 22.0)],
        },
    },
    'sales': {
        'category_value_ranges': {
            'low': [(25.0, 180.0), (18.0, 140.0)],
            'medium': [(14.0, 105.0), (10.0, 82.0)],
            'alta': [(7.0, 56.0), (5.0, 44.0)],
            'very high': [(3.0, 26.0), (2.0, 18.0)],
        },
        'map_value_ranges': {
            'low': [(650.0, 1600.0), (500.0, 1350.0)],
            'medium': [(380.0, 1200.0), (300.0, 920.0)],
            'alta': [(180.0, 720.0), (140.0, 560.0)],
            'very high': [(80.0, 340.0), (60.0, 250.0)],
        },
    },
    'web analytics': {
        'category_value_ranges': {
            'low': [(140.0, 900.0), (100.0, 700.0)],
            'medium': [(80.0, 560.0), (60.0, 420.0)],
            'alta': [(35.0, 260.0), (25.0, 190.0)],
            'very high': [(12.0, 110.0), (8.0, 80.0)],
        },
        'map_value_ranges': {
            'low': [(90.0, 450.0), (70.0, 360.0)],
            'medium': [(55.0, 260.0), (40.0, 210.0)],
            'alta': [(24.0, 130.0), (18.0, 95.0)],
            'very high': [(8.0, 55.0), (6.0, 40.0)],
        },
    },
}

for domain_key, range_cfg in DOMAIN_VALUE_RANGE_CFG.items():
    DOMAIN_CFG[domain_key].update(range_cfg)


def stable_seed(text: str) -> int:
    return int(hashlib.md5(text.encode('utf-8')).hexdigest()[:8], 16)


def rng_for(case_id: str) -> random.Random:
    return random.Random(stable_seed(case_id))


def normalize_difficulty(difficulty: str) -> str:
    return str(difficulty).strip().lower()


def difficulty_count(difficulty: str, kind: str, rnd: random.Random | None = None) -> int:
    if kind in {'scatter', 'surface'}:
        return {'low': 16, 'medium': 24, 'high': 32, 'very high': 40}.get(normalize_difficulty(difficulty), 24)
    intervals: Dict[str, Tuple[int, int]] = {
        'low': (2, 6),
        'medium': (7, 12),
        'high': (13, 18),
        'very high': (19, 30),
    }
    lo, hi = intervals.get(normalize_difficulty(difficulty), (7, 12))
    return rnd.randint(lo, hi) if rnd is not None else random.randint(lo, hi)


def radar_axis_count(difficulty: str, max_axes: int) -> int:
    target = {
        'low': 4,
        'medium': 5,
        'high': 6,
        'very high': 8,
    }.get(normalize_difficulty(difficulty), 5)
    return max(3, min(target, max_axes))


def choose_value_range(cfg: Dict[str, Any], difficulty: str, range_key: str, rnd: random.Random) -> Tuple[float, float]:
    difficulty_key = normalize_difficulty(difficulty)
    options = cfg.get(range_key, {}).get(difficulty_key)
    if not options:
        return (60.0, 130.0)
    lo, hi = rnd.choice(options)
    return float(lo), float(hi)


def base_spread_for_value_range(
    value_range: Tuple[float, float],
    rnd: random.Random,
    series_index: int = 0,
    series_count: int = 1,
) -> Tuple[float, float, float, float]:
    lo, hi = value_range
    width = max(1.0, hi - lo)
    center = (lo + hi) / 2
    if series_count > 1:
        center += (series_index - (series_count - 1) / 2) * width * 0.12
    spread = width * rnd.uniform(0.55, 0.85)
    base = center - spread * 0.45
    return lo, hi, base, spread


def clamp_series_values(values: List[float], lo: float, hi: float, allow_negative: bool = False) -> List[float]:
    if allow_negative:
        lower_bound = min(lo, -0.35 * hi)
        upper_bound = hi * 1.05
    else:
        lower_bound = max(0.0, lo * 0.9)
        upper_bound = hi * 1.05
    return [round(min(upper_bound, max(lower_bound, v)), 1) for v in values]


def clean_row(row: Dict[str, Any]) -> Dict[str, Any]:
    out = {}
    for k, v in row.items():
        if pd.isna(v):
            out[k] = None
        elif hasattr(v, 'item'):
            out[k] = v.item()
        else:
            out[k] = v

    return out


SOURCE_CASE_KEY_MAP = {
    'Excel_family': 'excel_family',
    'subtype_excel': 'excel_subtype',
    'style_varian': 'style_variant',
    'quote_subtype_5000': 'quota_subtype_5000',
    'quote_subtype_500': 'quota_subtype_500',
    'codi_subtipus': 'code_subtype',
    'familia_excel': 'excel_family',
    'subtipus_excel': 'excel_subtype',
    'estructura_dades': 'data_structure',
    'patro_estadistic': 'statistic_pattern',
    'domini_semantic': 'semantic_domain',
    'dificultat': 'difficulty',
}

DIFFICULTY_OUTPUT_MAP = {
    'baixa': 'low',
    'mitjana': 'medium',
    'alta': 'high',
    'molt alta': 'very high',
}

DOMAIN_OUTPUT_MAP = {
    'vendes': 'sales',
    'salut': 'health',
    'educació': 'education',
    'demografia': 'demographics',
    'operacions': 'operations',
    'energia': 'energy',
    'finances': 'finance',
    'manufactura': 'manufacturing',
    'clima': 'climate',
    'geografia': 'geography',
    'web analytics': 'web analytics',
}

STRUCTURE_OUTPUT_MAP = {
    'categòric simple': 'simple categorical',
    'categòric multiserie': 'multi-series categorical',
    'categòric amb etiquetes llargues': 'categorical with long labels',
    'categòric ordenat': 'ordered categorical',
    'categòric ordenat + acumulat': 'ordered categorical + cumulative',
    'sèrie temporal': 'time series',
    'sèrie temporal discreta': 'discrete time series',
    'multiserie temporal': 'multi-series temporal',
    'multiserie temporal acumulada': 'cumulative multi-series temporal',
    'parts del tot simple': 'simple part-of-whole',
    'parts del tot amb cua llarga': 'part-of-whole with long tail',
    'parts del tot multianell': 'multi-ring part-of-whole',
    'bivariant numèrica': 'numeric bivariate',
    'OHLC temporal': 'temporal OHLC',
    'graella numèrica 2D': '2D numeric grid',
    'multivariable sobre eixos comuns': 'multivariable on common axes',
    'distribució univariant': 'univariate distribution',
    'procés per etapes': 'stage-by-stage process',
    'multiserie mixta': 'mixed multi-series',
    'temporal amb eix secundari': 'temporal with secondary axis',
    'geogràfica': 'geographic',
}

PATTERN_OUTPUT_MAP = {
    'diferències clares': 'clear differences',
    'valors molt propers': 'very close values',
    'cua llarga': 'long tail',
    'pic local': 'local peak',
    'valors negatius': 'negative values',
    'creixement': 'growth',
    'decreixement': 'decline',
    'estacionalitat': 'seasonality',
    'canvi de règim': 'regime change',
    'soroll alt': 'high noise',
    'pic sobtat': 'sudden spike',
    'vall sobtada': 'sudden valley',
    'parts equilibrades': 'balanced parts',
    'part dominant': 'dominant part',
    'segments molt propers': 'very close segments',
    'acumulació creixent': 'growing accumulation',
    'correlació positiva': 'positive correlation',
    'correlació negativa': 'negative correlation',
    'correlació nul·la': 'no correlation',
    'outlier clar': 'clear outlier',
    'clústers': 'clusters',
    'relació corba': 'curved relationship',
    'alta volatilitat': 'high volatility',
    'baixa volatilitat': 'low volatility',
    'tendència alcista': 'upward trend',
    'tendència baixista': 'downward trend',
    'gap puntual': 'single gap',
    'superfície suau': 'smooth surface',
    'superfície rugosa': 'rough surface',
    'pic central': 'central peak',
    'doble pic': 'double peak',
    'gradient diagonal': 'diagonal gradient',
    'perfil equilibrat': 'balanced profile',
    'perfil espigat': 'spiky profile',
    'dues sèries contrastades': 'two contrasting series',
    'una dimensió dominant': 'one dominant dimension',
    'simètrica': 'symmetric',
    'asimètrica': 'asymmetric',
    'bimodal': 'bimodal',
    '80/20 marcat': 'clear 80/20',
    'concentració moderada': 'moderate concentration',
    'concentració forta': 'strong concentration',
    'caiguda uniforme': 'uniform drop',
    'bottleneck clar': 'clear bottleneck',
    'caiguda tardana': 'late drop',
    'caiguda inicial forta': 'strong initial drop',
    'creixement amb taxa': 'growth with rate',
    'volum i percentatge': 'volume and percentage',
    'dues escales diferents': 'two different scales',
    'sèrie principal + objectiu': 'main series + target',
}

RECORD_FIELD_MAP = {
    'categoria': 'category',
    'serie': 'series',
    'valor': 'value',
    'regio': 'region',
    'grup': 'group',
    'data': 'date',
    'obertura': 'open',
    'maxim': 'high',
    'minim': 'low',
    'tancament': 'close',
    'volum': 'volume',
    'eix': 'axis',
}


def translate_case_value(key: str, value: Any) -> Any:
    if value is None:
        return None
    text = str(value).strip().lower()
    if key == 'difficulty':
        return DIFFICULTY_OUTPUT_MAP.get(text, value)
    if key == 'semantic_domain':
        return DOMAIN_OUTPUT_MAP.get(text, value)
    if key == 'data_structure':
        return STRUCTURE_OUTPUT_MAP.get(text, value)
    if key == 'statistic_pattern':
        return PATTERN_OUTPUT_MAP.get(text, value)
    if key == 'question_template':
        return str(value).replace('{element_objetiu}', '{target_element}')
    return value


def translate_source_case(row: Dict[str, Any]) -> Dict[str, Any]:
    translated: Dict[str, Any] = {}
    for key, value in row.items():
        english_key = SOURCE_CASE_KEY_MAP.get(key, key)
        translated[english_key] = translate_case_value(english_key, value)
    return translated


def translate_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    translated_records: List[Dict[str, Any]] = []
    for record in records:
        translated_records.append({RECORD_FIELD_MAP.get(key, key): value for key, value in record.items()})
    return translated_records


def translate_encoding(encoding: Dict[str, Any]) -> Dict[str, Any]:
    translated: Dict[str, Any] = {}
    for channel, spec in encoding.items():
        translated[channel] = {
            **spec,
            'field': RECORD_FIELD_MAP.get(spec.get('field'), spec.get('field')),
        }
    return translated


def translate_axes(axes: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    translated: Dict[str, Dict[str, Any]] = {}
    for axis_name, axis in axes.items():
        translated[axis_name] = {
            **axis,
            'field': RECORD_FIELD_MAP.get(axis.get('field'), axis.get('field')),
        }
    return translated


def export_instance(instance: Dict[str, Any]) -> Dict[str, Any]:
    data = instance.get('data', {})
    source = data.get('source', {})
    exported = dict(instance)
    exported['data'] = {
        **data,
        'source': {
            **source,
            'values': translate_records(source.get('values', [])),
        },
    }
    exported['encoding'] = translate_encoding(instance.get('encoding', {}))
    exported['axes'] = translate_axes(instance.get('axes', {}))
    return exported


def choose_categories(cfg: Dict[str, Any], structure: str, count: int) -> List[str]:
    if structure == 'categorical with long labels':
        return cfg['long_categories'][:count]
    if structure == 'ordered categorical':
        return ['Very low', 'Low', 'Medium', 'High', 'Very high', 'Very high +'][0:count]
    if structure == 'geographic':
        return cfg['geo_categories'][:count]
    if 'temporal' in structure:
        if count <= 4:
            return QUARTERS[:count]
        if count <= 6:
            return MONTHS[:count]
        return YEARS[-count:]
    return cfg['simple_categories'][:count]


def short_label(label: str, max_words: int = 3, max_chars: int = 22) -> str:
    text = ' '.join(str(label).split())
    if len(text) <= max_chars:
        return text
    words = text.split()
    shortened = ' '.join(words[:max_words])
    if len(shortened) <= max_chars:
        return shortened
    return shortened[:max_chars - 1].rstrip() + '.'


def unique_preserve_order(values: List[str]) -> List[str]:
    return list(OrderedDict((value, None) for value in values).keys())


def difficulty_segment_count(difficulty: str, max_segments: int) -> int:
    target = {
        'low': 3,
        'medium': 5,
        'high': 8,
        'very high': 12,
    }.get(normalize_difficulty(difficulty), 5)
    return max(3, min(target, max_segments))


def difficulty_stacked_series_count(difficulty: str, max_series: int) -> int:
    target = {
        'low': 2,
        'medium': 3,
        'high': 4,
        'very high': 5,
    }.get(normalize_difficulty(difficulty), 3)
    return max(2, min(target, max_series))


def choose_parts_categories(cfg: Dict[str, Any], difficulty: str) -> List[str]:
    category_pool = unique_preserve_order(
        list(cfg['simple_categories'])
        + [short_label(label) for label in cfg.get('long_categories', [])]
        + list(cfg.get('geo_categories', []))
    )
    count = difficulty_segment_count(difficulty, len(category_pool))
    return category_pool[:count]


def choose_stacked_area_series(cfg: Dict[str, Any], difficulty: str) -> List[str]:
    base_series = list(cfg['series'])
    extras = [
        'Others',
        'Residual',
        'Complementary',
    ]
    series_pool = unique_preserve_order(base_series + extras)
    count = difficulty_stacked_series_count(difficulty, len(series_pool))
    return series_pool[:count]


def choose_series(cfg: Dict[str, Any], row: Dict[str, Any], structure: str) -> List[str]:
    subtype = str(row['subtype_excel'])
    family = str(row['Excel_family'])
    if structure in {'time series', 'discrete time series', 'simple categorical', 'ordered categorical', 'geographic'}:
        if family in {'Pie', 'Doughnut'}:
            return ['Share']
        if '100% Stacked' in subtype and family in {'Bar', 'Column', 'Area', 'Line'}:
            if structure == 'simple categorical':
                return ['Relative share']
        return ['Value']
    if structure == 'multi-ring part-of-whole':
        return ['Actual', 'Target']
    if family == 'Area' and 'stacked' in subtype.lower():
        return choose_stacked_area_series(cfg, str(row['difficulty']))
    if 'multi-series' in structure or structure in {'multivariable on common axes', 'temporal with secondary axis'}:
        return cfg['series']
    return cfg['series']


def line_pattern(values_n: int, pattern: str, base: float, spread: float, rnd: random.Random) -> List[float]:
    xs = list(range(values_n))
    vals: List[float] = []
    for i in xs:
        if pattern in {'growth', 'growth with rate'}:
            v = base + spread * (i / max(1, values_n - 1)) * (1.1 if pattern == 'growth with rate' else 1.0)
            v += rnd.uniform(-0.06, 0.06) * spread
        elif pattern == 'decline':
            v = base + spread * (1 - i / max(1, values_n - 1))
            v += rnd.uniform(-0.06, 0.06) * spread
        elif pattern == 'seasonality':
            v = base + spread * (0.5 + 0.45 * math.sin(2 * math.pi * i / max(3, values_n)))
            v += rnd.uniform(-0.04, 0.04) * spread
        elif pattern == 'local peak':
            peak = values_n // 2
            v = base + spread * (0.25 + 0.75 * math.exp(-((i - peak) ** 2) / max(1, values_n / 3)))
        elif pattern == 'sudden spike':
            peak = rnd.randint(max(1, values_n // 3), max(1, values_n - 2))
            v = base + spread * (0.25 + (1.2 if i == peak else 0.15))
        elif pattern == 'sudden valley':
            dip = rnd.randint(max(1, values_n // 3), max(1, values_n - 2))
            v = base + spread * (0.8 if i != dip else 0.15)
        elif pattern == 'regime change':
            cut = max(1, values_n // 2)
            v = base + (0.35 if i < cut else 0.85) * spread + rnd.uniform(-0.05, 0.05) * spread
        elif pattern == 'low volatility':
            v = base + 0.1 * spread + rnd.uniform(-0.03, 0.03) * spread
        elif pattern == 'high noise':
            v = base + 0.5 * spread + rnd.uniform(-0.4, 0.4) * spread
        elif pattern == 'negative values':
            v = base - spread / 2 + spread * (i / max(1, values_n - 1)) + rnd.uniform(-0.12, 0.12) * spread
        elif pattern == 'very close values':
            center = base + 0.5 * spread
            v = center + rnd.uniform(-0.05, 0.05) * center
        elif pattern == 'clear differences':
            v = base + (0.15 + 0.8 * i / max(1, values_n - 1)) * spread
        else:
            v = base + spread * (0.45 + 0.12 * math.sin(i))
        vals.append(round(v, 1))
    return vals


def proportions(n: int, pattern: str, rnd: random.Random) -> List[float]:
    if pattern in {'balanced parts', 'balanced profile', 'very close values', 'very close segments', 'uniform distribution'}:
        raw = [1 + rnd.uniform(-0.06, 0.06) for _ in range(n)]
    elif pattern in {'dominant part', 'long tail'}:
        raw = [n * 2.5] + [max(0.2, 1 / (i + 1) + rnd.uniform(0, 0.2)) for i in range(1, n)]
    elif pattern == 'two dominant regions':
        raw = [n * 1.8, n * 1.6] + [max(0.3, 0.8 + rnd.uniform(-0.1, 0.15)) for _ in range(max(0, n - 2))]
    elif pattern in {'spiky profile'}:
        spike = rnd.randint(0, n - 1)
        raw = [0.7 + rnd.uniform(-0.1, 0.1) for _ in range(n)]
        raw[spike] = n * 2.2
    else:
        raw = [1 + rnd.uniform(-0.25, 0.25) for _ in range(n)]
    s = sum(raw)
    vals = [round(100 * r / s, 1) for r in raw]
    vals[-1] = round(vals[-1] + (100 - sum(vals)), 1)
    return vals


def scale_to_percent(values: List[float]) -> List[float]:
    min_v = min(values)
    max_v = max(values)
    if math.isclose(max_v, min_v):
        return [50.0 for _ in values]
    scaled = [round((v - min_v) / (max_v - min_v) * 100, 1) for v in values]
    return scaled


def estimate_numeric_interval(values: List[float]) -> float | None:
    unique_sorted = sorted(set(round(float(v), 6) for v in values))
    if len(unique_sorted) < 2:
        return None
    diffs = [unique_sorted[i] - unique_sorted[i - 1] for i in range(1, len(unique_sorted))]
    positive_diffs = [d for d in diffs if d > 0]
    if not positive_diffs:
        return None
    return round(sum(positive_diffs) / len(positive_diffs), 3)


def infer_unit_from_title(title: str | None) -> str | None:
    if not title:
        return None
    t = str(title)
    start = t.rfind('(')
    end = t.rfind(')')
    if 0 <= start < end:
        return t[start + 1:end].strip()
    return None


def build_axes_metadata(inst: Dict[str, Any], records: List[Dict[str, Any]], meta: Dict[str, Any], encoding: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    axes: Dict[str, Dict[str, Any]] = {}
    axis_keys = ['x', 'y', 'secondary_y', 'z']
    for axis_key in axis_keys:
        spec = encoding.get(axis_key)
        if not spec:
            continue
        field = spec.get('field')
        axis_type = spec.get('type')
        title = spec.get('title')
        values = [r.get(field) for r in records if r.get(field) is not None]

        axis_unit = infer_unit_from_title(title)
        if axis_key == 'y':
            axis_unit = meta.get('unit') or axis_unit
        elif axis_key == 'secondary_y':
            axis_unit = spec.get('title') or axis_unit
        elif axis_type == 'temporal':
            axis_unit = axis_unit or 'date'
        elif axis_type == 'nominal':
            axis_unit = axis_unit or 'category'

        min_value: Any = None
        max_value: Any = None
        interval: Any = None

        if axis_key == 'y' and inst['chart'].get('stack_mode') == 'percent':
            min_value = 0.0
            max_value = 100.0
            interval = 10.0
            axis_unit = axis_unit or '%'
        elif axis_type == 'quantitative':
            numeric_values = []
            for v in values:
                try:
                    numeric_values.append(float(v))
                except (TypeError, ValueError):
                    continue
            if numeric_values:
                min_value = round(min(numeric_values), 3)
                max_value = round(max(numeric_values), 3)
                interval = estimate_numeric_interval(numeric_values)
        elif axis_type == 'temporal':
            dt = pd.to_datetime(values, errors='coerce')
            valid = [x for x in dt if pd.notna(x)]
            if valid:
                valid_sorted = sorted(set(valid))
                min_value = valid_sorted[0].strftime('%Y-%m-%d')
                max_value = valid_sorted[-1].strftime('%Y-%m-%d')
                if len(valid_sorted) > 1:
                    deltas = [
                        (valid_sorted[i] - valid_sorted[i - 1]).total_seconds() / 86400
                        for i in range(1, len(valid_sorted))
                    ]
                    interval = round(sum(deltas) / len(deltas), 2)
        else:
            if values:
                min_value = values[0]
                max_value = values[-1]
                interval = 1

        axes[axis_key] = {
            'field': field,
            'title': title,
            'type': axis_type,
            'min': min_value,
            'max': max_value,
            'interval': interval,
            'unit': axis_unit,
        }
    return axes


def build_excel_metadata(title: str, axes: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    axis_labels = {'x': 'X', 'y': 'Y', 'secondary_y': 'Y2', 'z': 'Z'}
    axes_payload: Dict[str, Dict[str, Any]] = {}
    for axis_key, axis_label in axis_labels.items():
        axis = axes.get(axis_key)
        if not axis:
            continue
        axes_payload[axis_key] = {
            'label': axis_label,
            'title': axis.get('title'),
            'start': axis.get('min'),
            'end': axis.get('max'),
            'unit': axis.get('unit'),
            'interval': axis.get('interval'),
        }
    return {
        'chart_title': title,
        'axes': axes_payload,
    }


def count_series_in_records(records: List[Dict[str, Any]]) -> int:
    if not records:
        return 0
    if 'serie' in records[0]:
        return len(OrderedDict((r['serie'], None) for r in records))
    if 'series' in records[0]:
        return len(OrderedDict((r['series'], None) for r in records))
    if 'grup' in records[0]:
        return len(OrderedDict((r['grup'], None) for r in records))
    if 'group' in records[0]:
        return len(OrderedDict((r['group'], None) for r in records))
    return 1


def generate_time_long(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    structure = row['data_structure']
    family = row['Excel_family']
    pattern = row['statistic_pattern']
    count = difficulty_count(row['difficulty'], 'time', rnd)
    categories = choose_categories(cfg, structure, count)
    series = choose_series(cfg, row, structure)
    share_like = ('100% Stacked' in row['subtype_excel']) or family in {'Pie', 'Doughnut'} or 'part-of-whole' in structure

    records: List[Dict[str, Any]] = []
    if share_like and len(series) > 1:
        for cat in categories:
            vals = proportions(len(series), pattern, rnd)
            for ser, v in zip(series, vals):
                records.append({'categoria': cat, 'serie': ser, 'valor': v})
        unit = cfg['unit_pct']
        y_title = cfg['unit_pct']
    else:
        unit = cfg['unit_abs']
        y_title = cfg['unit_abs']
        value_range = choose_value_range(cfg, row['difficulty'], 'category_value_ranges', rnd)
        for s_idx, ser in enumerate(series):
            lo, hi, base, spread = base_spread_for_value_range(value_range, rnd, s_idx, len(series))
            vals = line_pattern(len(categories), pattern, base, spread, rnd)
            if '100% Stacked' in row['subtype_excel'] and len(series) == 1:
                vals = scale_to_percent(vals)
                unit = cfg['unit_pct']
                y_title = cfg['unit_pct']
            else:
                vals = clamp_series_values(vals, lo, hi, allow_negative=pattern == 'negative values')
            for cat, v in zip(categories, vals):
                records.append({'categoria': cat, 'serie': ser, 'valor': float(v)})
    meta = {
        'x_title': 'Period' if 'temporal' in structure else 'Category',
        'y_title': y_title,
        'unit': unit,
        'series': series,
        'categories': categories,
    }
    return records, meta


def generate_parts(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    if row['data_structure'] == 'multi-ring part-of-whole':
        categories = choose_parts_categories(cfg, str(row['difficulty']))
        series = ['Actual', 'Target']
        records = []
        for ser in series:
            vals = proportions(len(categories), row['statistic_pattern'], rnd)
            for cat, v in zip(categories, vals):
                records.append({'categoria': cat, 'serie': ser, 'valor': v})
        return records, {
            'x_title': 'Category',
            'y_title': cfg['unit_pct'],
            'unit': cfg['unit_pct'],
            'series': series,
            'categories': categories,
        }
    else:
        categories = choose_parts_categories(cfg, str(row['difficulty']))
        vals = proportions(len(categories), row['statistic_pattern'], rnd)
        records = [{'categoria': cat, 'serie': 'Share', 'valor': v} for cat, v in zip(categories, vals)]
        return records, {
            'x_title': 'Category',
            'y_title': cfg['unit_pct'],
            'unit': cfg['unit_pct'],
            'series': ['Share'],
            'categories': categories,
        }


def generate_geo(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    categories = cfg['geo_categories'][:max(5, difficulty_count(row['difficulty'], 'geo', rnd))]
    if row['Excel_family'] == 'Map' or row['data_structure'] == 'geographic':
        pattern = row['statistic_pattern']
        lo, hi = choose_value_range(cfg, row['difficulty'], 'map_value_ranges', rnd)
        if pattern == 'north-south gradient':
            vals = list(reversed(clamp_series_values(line_pattern(len(categories), 'growth', lo, hi - lo, rnd), lo, hi)))
        elif pattern == 'regional hotspot':
            vals = [round(rnd.uniform(lo, lo + 0.55 * (hi - lo)), 1) for _ in categories]
            vals[rnd.randrange(len(vals))] = round(rnd.uniform(lo + 0.78 * (hi - lo), hi), 1)
        else:
            vals = [round(rnd.uniform(lo, hi), 1) for _ in categories]
        records = [{'regio': cat, 'valor': v} for cat, v in zip(categories, vals)]
        return records, {
            'x_title': 'Region',
            'y_title': cfg['map_metric'],
            'unit': cfg['map_metric'],
            'series': ['Value'],
            'categories': categories,
        }
    return generate_time_long(row, cfg, rnd)


def generate_scatter(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = difficulty_count(row['difficulty'], 'scatter', rnd)
    x_name, x_min, x_max = cfg['scatter_x']
    y_name, y_min, y_max = cfg['scatter_y']
    pattern = row['statistic_pattern']
    groups = ['Group 1', 'Group 2'] if pattern == 'clusters' else ['Sample']
    records = []
    for i in range(n):
        if pattern == 'positive correlation':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_min + ratio * (y_max - y_min) + rnd.uniform(-0.08, 0.08) * (y_max - y_min)
        elif pattern == 'negative correlation':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_max - ratio * (y_max - y_min) + rnd.uniform(-0.08, 0.08) * (y_max - y_min)
        elif pattern == 'no correlation':
            x = rnd.uniform(x_min, x_max)
            y = rnd.uniform(y_min, y_max)
        elif pattern == 'curved relationship':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_min + (ratio ** 2) * (y_max - y_min) + rnd.uniform(-0.04, 0.04) * (y_max - y_min)
        elif pattern == 'clusters':
            grp = groups[i % 2]
            if grp == 'Group 1':
                x = rnd.uniform(x_min, x_min + 0.35 * (x_max - x_min))
                y = rnd.uniform(y_min, y_min + 0.35 * (y_max - y_min))
            else:
                x = rnd.uniform(x_min + 0.55 * (x_max - x_min), x_max)
                y = rnd.uniform(y_min + 0.55 * (y_max - y_min), y_max)
            records.append({'x': round(x, 2), 'y': round(y, 2), 'grup': grp})
            continue
        elif pattern == 'clear outlier':
            if i == 0:
                x = x_max * 0.98
                y = y_max * 0.98
            else:
                x = rnd.uniform(x_min, x_min + 0.6 * (x_max - x_min))
                y = rnd.uniform(y_min, y_min + 0.6 * (y_max - y_min))
        else:
            x = rnd.uniform(x_min, x_max)
            y = rnd.uniform(y_min, y_max)
        records.append({'x': round(x, 2), 'y': round(y, 2), 'grup': groups[0]})
    return records, {
        'x_title': x_name,
        'y_title': y_name,
        'unit': y_name,
        'series': groups,
        'categories': None,
    }


def radar_series_profile(cats: List[str], pattern: str, series_index: int, series_count: int, rnd: random.Random) -> List[float]:
    n = len(cats)
    if n == 0:
        return []

    dominant_idx = (series_index * max(1, n // max(1, series_count)) + series_index) % n
    opposite_idx = (dominant_idx + max(1, n // 2)) % n
    secondary_idx = (dominant_idx + 1) % n

    values: List[float] = []
    for idx in range(n):
        circular_distance = min((idx - dominant_idx) % n, (dominant_idx - idx) % n)
        proximity = 1.0 - (circular_distance / max(1, n // 2))

        if pattern == 'perfil equilibrat':
            value = 62 + 6 * math.sin((idx + series_index) * 0.9) + rnd.uniform(-4, 4)
        elif pattern == 'perfil espigat':
            value = 34 + 52 * max(0.0, proximity) + rnd.uniform(-5, 5)
        elif pattern == 'una dimensió dominant':
            value = 42 + 38 * max(0.0, proximity) + rnd.uniform(-4, 4)
        elif pattern == 'dues sèries contrastades':
            if idx == dominant_idx:
                value = 86 + rnd.uniform(-3, 3)
            elif idx == opposite_idx:
                value = 28 + rnd.uniform(-4, 4)
            else:
                value = 48 + 14 * math.sin((idx + 1) * (series_index + 1)) + rnd.uniform(-5, 5)
        else:
            value = 50 + 24 * max(0.0, proximity) + 10 * math.sin((idx + series_index) * 0.8) + rnd.uniform(-6, 6)

        if idx == secondary_idx and pattern in {'perfil espigat', 'una dimensió dominant'}:
            value += 8
        values.append(round(min(100, max(0, value)), 1))
    return values


def generate_radar(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    cats = cfg['radar_axes'][:radar_axis_count(row['difficulty'], len(cfg['radar_axes']))]
    series = cfg['series'][:3]
    records = []
    for s_idx, ser in enumerate(series):
        vals = radar_series_profile(cats, row['statistic_pattern'], s_idx, len(series), rnd)
        for cat, v in zip(cats, vals):
            records.append({'categoria': cat, 'serie': ser, 'valor': round(v, 1)})
    return records, {
        'x_title': 'Indicator',
        'y_title': 'Score (0-100)',
        'unit': 'Score (0-100)',
        'series': series,
        'categories': cats,
    }


def generate_combo(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = difficulty_count(row['difficulty'], 'combo', rnd)
    cats = MONTHS[:n] if n <= 6 else YEARS[-n:]
    pattern = row['statistic_pattern']
    lo, hi, base, spread = base_spread_for_value_range(choose_value_range(cfg, row['difficulty'], 'category_value_ranges', rnd), rnd)
    primary = line_pattern(
        len(cats),
        'growth' if pattern in {'two different scales', 'main series + target', 'growth with rate'} else pattern,
        base,
        spread,
        rnd,
    )
    primary = clamp_series_values(primary, lo, hi, allow_negative=pattern == 'negative values')
    if pattern == 'two different scales':
        secondary = [round(10 + 8 * math.sin(i), 1) for i in range(len(cats))]
    elif pattern == 'main series + target':
        secondary = [round(sum(primary) / len(primary) * 0.95, 1) for _ in cats]
    else:
        secondary = line_pattern(len(cats), 'low volatility', 25, 12, rnd)
    records = []
    for cat, p, s in zip(cats, primary, secondary):
        records.append({'categoria': cat, 'serie': 'Primary', 'valor': float(p), 'eix': 'primary'})
        records.append({'categoria': cat, 'serie': 'Secondary', 'valor': float(s), 'eix': 'secondary'})
    return records, {
        'x_title': 'Period',
        'y_title': cfg['combo_primary'],
        'unit': cfg['combo_primary'],
        'secondary_title': cfg['combo_secondary'],
        'series': ['Primary', 'Secondary'],
        'categories': cats,
    }


def generate_stock(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = max(10, min(20, difficulty_count(row['difficulty'], 'stock', rnd) + 8))
    dates = [f'2024-{m:02d}-{d:02d}' for m, d in zip(([1] * n), range(1, n + 1))]
    pattern = row['statistic_pattern']
    if pattern in {'growth', 'low volatility'}:
        closes = line_pattern(n, pattern, 70, 18, rnd)
    else:
        closes = line_pattern(n, 'low volatility', 70, 18, rnd)
    records = []
    prev_close = closes[0]
    for i, (date, close) in enumerate(zip(dates, closes)):
        open_ = round(prev_close + rnd.uniform(-1.5, 1.5), 2)
        high = round(max(open_, close) + rnd.uniform(0.8, 4.0), 2)
        low = round(min(open_, close) - rnd.uniform(0.8, 4.0), 2)
        volume = int(rnd.uniform(50000, 250000))
        records.append({'data': date, 'obertura': round(open_, 2), 'maxim': high, 'minim': low, 'tancament': round(close, 2), 'volum': volume})
        prev_close = close
    return records, {
        'x_title': 'Date',
        'y_title': cfg['stock_label'],
        'unit': cfg['stock_label'],
        'series': None,
        'categories': dates,
    }


def generate_surface(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    xs = [1, 2, 3, 4, 5, 6]
    ys = [1, 2, 3, 4, 5, 6]
    records = []
    for x in xs:
        for y in ys:
            z = 30 + 8 * math.sin(x) + 10 * math.cos(y / 2)
            if row['statistic_pattern'] == 'rough surface':
                z += rnd.uniform(-7, 7)
            records.append({'x': x, 'y': y, 'z': round(z, 1)})
    return records, {
        'x_title': 'X Axis',
        'y_title': 'Y Axis',
        'unit': 'Intensity',
        'series': None,
        'categories': None,
    }


def title_for(row: Dict[str, Any], cfg: Dict[str, Any], meta: Dict[str, Any]) -> str:
    domain = str(row['semantic_domain'])
    structure = row['data_structure']
    family = row['Excel_family']
    case_label = str(row['case_id']).removeprefix('CASE_')
    if family == 'Scatter':
        return f"Relationship between {domain} indicators ({case_label})"
    if family == 'Stock':
        return f"Daily evolution of {cfg['stock_label'].lower()} ({case_label})"
    if family == 'Map':
        return f"Territorial distribution of {domain} ({case_label})"
    if family == 'Surface':
        return f"Energy response surface ({case_label})"
    if structure == 'multi-ring part-of-whole':
        return f"Compared category composition in {domain} ({case_label})"
    if 'part-of-whole' in structure:
        return f"Relative share by category in {domain} ({case_label})"
    if 'temporal' in structure:
        return f"Evolution of {domain} by period ({case_label})"
    if 'multivariable' in structure:
        return f"Compared profile of {domain} indicators ({case_label})"
    if structure == 'geographic':
        return f"Territorial indicator for {domain} ({case_label})"
    return f"{domain} indicators by category ({case_label})"


def make_canonical_instance(row: Dict[str, Any]) -> Dict[str, Any]:
    cfg = DOMAIN_CFG[str(row['semantic_domain']).strip().lower()]
    rnd = rng_for(row['case_id'])
    structure = row['data_structure']
    family = row['Excel_family']

    if family == 'Scatter' or structure == 'numeric bivariate':
        records, meta = generate_scatter(row, cfg, rnd)
        data_format = 'scatter'
        encoding = {
            'x': {'field': 'x', 'type': 'quantitative', 'title': meta['x_title']},
            'y': {'field': 'y', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'grup', 'type': 'nominal', 'title': 'Group'}
        }
    elif family == 'Stock' or structure == 'OHLC + volume':
        records, meta = generate_stock(row, cfg, rnd)
        data_format = 'stock'
        encoding = {
            'x': {'field': 'data', 'type': 'temporal', 'title': meta['x_title']},
            'y': {'field': 'tancament', 'type': 'quantitative', 'title': meta['y_title']},
        }
    elif family == 'Surface' or structure == '2D numeric grid':
        records, meta = generate_surface(row, cfg, rnd)
        data_format = 'surface'
        encoding = {
            'x': {'field': 'x', 'type': 'quantitative', 'title': meta['x_title']},
            'y': {'field': 'y', 'type': 'quantitative', 'title': meta['y_title']},
            'z': {'field': 'z', 'type': 'quantitative', 'title': meta['unit']}
        }
    elif family == 'Radar' or structure == 'multivariable on common axes':
        records, meta = generate_radar(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Series'}
        }
    elif family == 'Combo' or structure == 'temporal with secondary axis':
        records, meta = generate_combo(row, cfg, rnd)
        data_format = 'combo'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Series'},
            'secondary_y': {'field': 'valor', 'type': 'quantitative', 'title': meta['secondary_title']}
        }
    elif family == 'Map' or structure == 'geographic':
        records, meta = generate_geo(row, cfg, rnd)
        data_format = 'map'
        encoding = {
            'x': {'field': 'regio', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
        }
    elif 'part-of-whole' in structure or family in {'Pie', 'Doughnut'}:
        records, meta = generate_parts(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Series'}
        }
    else:
        records, meta = generate_time_long(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Series'}
        }

    title = title_for(row, cfg, meta)
    source_case = translate_source_case(row)

    exact_excel_supported = family not in {'Map', 'Surface'} and row['subtype_excel'] not in {'Pie of Pie', 'Bar of Pie', 'Filled Map', 'Contour'}
    render_mode = 'exact' if exact_excel_supported else 'approximate_or_metadata'

    provisional_inst = {
        'chart': {
            'stack_mode': 'percent' if '100% Stacked' in row['subtype_excel'] else ('normal' if 'Stacked' in row['subtype_excel'] else None),
        }
    }
    axes = build_axes_metadata(provisional_inst, records, meta, encoding)
    excel_metadata = build_excel_metadata(title, axes)
    series_count = count_series_in_records(records)

    return {
        '$schema': 'https://example.org/chart-canonical.schema.json',
        'version': '1.0-exploratory',
        'id': row['case_id'],
        'title': title,
        'description': f"Synthetic canonical instance for the {row['subtipus_excel']} chart in the {source_case['semantic_domain']} domain.",
        'style_ref': STYLE_REF,
        'source_case': source_case,
        'data': {
            'source': {'kind': 'inline', 'values': records},
            'structure': source_case['data_structure'],
            'statistical_pattern': source_case['statistic_pattern'],
            'semantic_domain': source_case['semantic_domain'],
            'unit': meta['unit'],
            'format': data_format,
        },
        'chart': {
            'family': row['familia_excel'].lower(),
            'excel_family': row['Excel_family'],
            'excel_subtype': row['subtype_excel'],
            'excel_subtype_code': row['code_subtype'],
            'stack_mode': 'percent' if '100% Stacked' in row['subtype_excel'] else ('normal' if 'Stacked' in row['subtype_excel'] else None),
            'orientation': 'horizontal' if row['Excel_family'] == 'Bar' else 'vertical',
            'render_mode': render_mode,
            'exact_excel_supported': exact_excel_supported,
        },
        'excel_metadata': excel_metadata,
        'axes': axes,
        'encoding': encoding,
        'labels': {
            'title_visible': True,
            'legend_visible': series_count > 1,
            'data_labels': 'data labels' in str(row.get('variant_estil', '')).lower() or 'etiquetes de dades' in str(row.get('variant_estil', '')).lower(),
        },
        'xlsxwriter': {
            'style': 10,
            'legend_position': 'bottom',
            'size': {'width': 640, 'height': 360},
            'render_mode': render_mode,
        },
    }


def safe_sheet_name(case_id: str, used: set[str]) -> str:
    name = case_id[:31]
    base = name
    i = 2
    while name in used:
        suffix = f'_{i}'
        name = (base[:31-len(suffix)] + suffix)
        i += 1
    used.add(name)
    return name


def case_sort_key(inst: Dict[str, Any]) -> Tuple[int, str]:
    case_id = str(inst.get('id', ''))
    digits = ''.join(ch for ch in case_id if ch.isdigit())
    return (int(digits) if digits else 10**9, case_id)


def add_metadata(ws, formats, inst, sheet_name):
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 50)
    ws.write('A1', inst['title'], formats['title'])
    meta_rows = [
        ('Case ID', inst['id']),
        ('Excel family', inst['chart']['excel_family']),
        ('Excel subtype', inst['chart']['excel_subtype']),
        ('Data structure', inst['source_case']['data_structure']),
        ('Pattern', inst['source_case']['statistic_pattern']),
        ('Difficulty', inst['source_case']['difficulty']),
        ('Domain', inst['source_case']['semantic_domain']),
        ('Style ref', inst['style_ref']['href']),
        ('Excel render mode', inst['chart']['render_mode']),
    ]

    axes = inst.get('axes', {})
    axis_labels = [('x', 'X'), ('y', 'Y'), ('secondary_y', 'Y2'), ('z', 'Z')]
    for axis_key, axis_label in axis_labels:
        axis = axes.get(axis_key)
        if not axis:
            continue
        meta_rows.extend([
            (f'Axis {axis_label} - Minimum', axis.get('min')),
            (f'Axis {axis_label} - Maximum', axis.get('max')),
            (f'Axis {axis_label} - Interval', axis.get('interval')),
            (f'Axis {axis_label} - Unit', axis.get('unit')),
        ])

    r = 2
    for k, v in meta_rows:
        ws.write(r, 0, k, formats['label'])
        ws.write(r, 1, '' if v is None else v, formats['text'])
        r += 1
    return r


def add_summary_sheet(workbook, instances, sheet_map, render_stats):
    ws = workbook.add_worksheet('Summary')
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9E2F3', 'border': 1})
    fmt_link = workbook.add_format({'font_color': 'blue', 'underline': 1})
    headers = ['case_id', 'title', 'excel_family', 'excel_subtype', 'data_structure', 'statistic_pattern', 'difficulty', 'semantic_domain', 'excel_render_mode']
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt_header)
    for r, inst in enumerate(instances, start=1):
        src = inst['source_case']
        ws.write_url(r, 0, f"internal:'{sheet_map[inst['id']]}'!A1", fmt_link, string=inst['id'])
        ws.write(r, 1, inst['title'])
        ws.write(r, 2, src['excel_family'])
        ws.write(r, 3, src['excel_subtype'])
        ws.write(r, 4, src['data_structure'])
        ws.write(r, 5, src['statistic_pattern'])
        ws.write(r, 6, src['difficulty'])
        ws.write(r, 7, src['semantic_domain'])
        ws.write(r, 8, inst['chart']['render_mode'])
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(instances), len(headers) - 1)
    ws.set_column('A:A', 14)
    ws.set_column('B:B', 55)
    ws.set_column('C:I', 22)
    start = len(instances) + 3
    ws.write(start, 0, 'Render statistics', fmt_header)
    for i, (k, v) in enumerate(render_stats.items(), start=start + 1):
        ws.write(i, 0, k)
        ws.write(i, 1, v)


def build_pivot_long(records: List[Dict[str, Any]], category_field='categoria', series_field='serie', value_field='valor'):
    categories = list(OrderedDict((r[category_field], None) for r in records).keys())
    series = list(OrderedDict((r[series_field], None) for r in records).keys())
    matrix = {s: {c: None for c in categories} for s in series}
    for r in records:
        matrix[r[series_field]][r[category_field]] = r[value_field]
    return categories, series, matrix


def excel_chart_args(inst):
    fam = inst['chart']['excel_family']
    subtype = inst['chart']['excel_subtype']
    if fam == 'Area':
        args: Dict[str, Any] = {'type': 'area'}
    elif fam == 'Bar':
        args = {'type': 'bar'}
    elif fam == 'Column':
        args = {'type': 'column'}
    elif fam == 'Line':
        args = {'type': 'line'}
    elif fam == 'Pie':
        args = {'type': 'pie'}
    elif fam == 'Doughnut':
        args = {'type': 'doughnut'}
    elif fam == 'Scatter':
        args = {'type': 'scatter'}
    elif fam == 'Radar':
        args = {'type': 'radar'}
    else:
        return None
    args['has_markers'] = False
    lower = subtype.lower()
    if '100% stacked' in lower:
        args['subtype'] = 'percent_stacked'
    elif 'stacked' in lower:
        args['subtype'] = 'stacked'
    elif subtype == 'Line with Markers':
        args['has_markers'] = True
    elif subtype == 'Stacked Line with Markers':
        args['subtype'] = 'stacked'
        args['has_markers'] = True
    elif subtype == '100% Stacked Line with Markers':
        args['subtype'] = 'percent_stacked'
        args['has_markers'] = True
    elif subtype == 'Scatter with Only Markers':
        args['subtype'] = 'marker_only'
    elif subtype == 'Scatter with Straight Lines':
        args['subtype'] = 'straight'
    elif subtype == 'Scatter with Straight Lines and Markers':
        args['subtype'] = 'straight_with_markers'
    elif subtype == 'Scatter with Smooth Lines':
        args['subtype'] = 'smooth'
    elif subtype == 'Scatter with Smooth Lines and Markers':
        args['subtype'] = 'smooth_with_markers'
    elif subtype == 'Radar with Markers':
        args['subtype'] = 'with_markers'
    elif subtype == 'Filled Radar':
        args['subtype'] = 'filled'
    return args


def series_style(
    chart_type: str,
    series_index: int,
    chart_subtype: str | None = None,
    has_markers: bool = False,
) -> Dict[str, Any]:
    style_entry = ACCESSIBLE_SERIES_STYLE_TABLE[series_index % len(ACCESSIBLE_SERIES_STYLE_TABLE)]
    color = style_entry['color']
    pattern_id = style_entry['pattern_id']
    pattern_entry = next(
        item for item in ACCESSIBLE_PATTERN_PALETTE if item['id'] == pattern_id
    )
    pattern = {
        'pattern': ACCESSIBLE_PATTERN_TOKEN_TO_XLSXWRITER[pattern_entry['token']],
        'fg_color': color,
        'bg_color': '#FFFFFF',
    }
    if chart_type in {'area', 'bar', 'column', 'pie', 'doughnut'}:
        style = {
            'fill': {'color': color, 'transparency': 50},
            'line': {'color': color, 'transparency': 50},
        }
        if ENABLE_CHART_PATTERNS:
            style['pattern'] = pattern
        return style
    if chart_type in {'line', 'radar'}:
        if chart_type == 'radar' and chart_subtype == 'filled':
            style = {
                'fill': {'color': color, 'transparency': 60},
                'line': {'color': color, 'transparency': 35},
            }
            if ENABLE_CHART_PATTERNS:
                style['pattern'] = pattern
            return style
        style = {'line': {'color': color, 'transparency': 50}}
        if chart_type == 'line' and has_markers:
            marker_cfg = ACCESSIBLE_LINE_MARKER_TABLE[series_index % len(ACCESSIBLE_LINE_MARKER_TABLE)]
            style['marker'] = {
                'type': marker_cfg['marker'],
                'size': 6,
                'fill': {'color': color},
                'border': {'color': color},
            }
        return style
    if chart_type == 'scatter':
        return {
            'marker': {
                'type': 'automatic',
                'fill': {'color': color, 'transparency': 50},
                'border': {'color': color, 'transparency': 50},
            },
        }
    return {'line': {'color': color, 'transparency': 50}}


def apply_legend_for_series_count(chart, legend_position: str, series_count: int) -> None:
    if series_count > 1:
        chart.set_legend({'position': legend_position})
    else:
        chart.set_legend({'none': True})


def excel_axis_options(name: str, extra: Dict[str, Any] | None = None) -> Dict[str, Any]:
    axis = {
        'name': name,
        'label_position': 'next_to',
        'num_font': {'rotation': 0},
        'name_font': {'rotation': 0},
    }
    if extra:
        axis.update(extra)
    return axis


def point_styles_for_categories(count: int) -> List[Dict[str, Any]]:
    styles: List[Dict[str, Any]] = []
    for idx in range(count):
        style_entry = ACCESSIBLE_SERIES_STYLE_TABLE[idx % len(ACCESSIBLE_SERIES_STYLE_TABLE)]
        color = style_entry['color']
        pattern_id = style_entry['pattern_id']
        pattern_entry = next(
            item for item in ACCESSIBLE_PATTERN_PALETTE if item['id'] == pattern_id
        )
        style = {
            'fill': {'color': color, 'transparency': 20},
            'line': {'color': color, 'transparency': 10},
        }
        if ENABLE_CHART_PATTERNS:
            style['pattern'] = {
                'pattern': ACCESSIBLE_PATTERN_TOKEN_TO_XLSXWRITER[pattern_entry['token']],
                'fg_color': color,
                'bg_color': '#FFFFFF',
            }
        styles.append(style)
    return styles


def data_label_options(chart_type: str) -> Dict[str, Any]:
    if chart_type == 'bar':
        return {'value': True, 'position': 'inside_end'}
    return {'value': True}


def insert_doughnut_ring_labels(ws, anchor_cell: str, series: List[str]) -> None:
    if len(series) < 2:
        return
    if len(series) == 2:
        y_offsets = [112, 148]
    else:
        start_y = 92
        step = 28
        y_offsets = [start_y + idx * step for idx in range(len(series))]

    for idx, ser in enumerate(series):
        ws.insert_textbox(
            anchor_cell,
            ser,
            {
                'x_offset': 235,
                'y_offset': y_offsets[idx],
                'width': 118,
                'height': 22,
                'font': {'bold': True, 'color': '#404040', 'size': 10},
                'align': {'vertical': 'middle', 'horizontal': 'center'},
                'fill': {'none': True},
                'line': {'none': True},
            },
        )


def render_long_chart(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    categories, series, matrix = build_pivot_long(records)
    start_col = 0
    ws.write(start_row, start_col, inst['encoding']['x']['title'], formats['header'])
    for idx, s in enumerate(series, start=1):
        ws.write(start_row, start_col + idx, s, formats['header'])
    for r_idx, cat in enumerate(categories, start=1):
        ws.write(start_row + r_idx, start_col, cat, formats['text'])
        for c_idx, s in enumerate(series, start=1):
            val = matrix[s][cat]
            ws.write_number(start_row + r_idx, start_col + c_idx, float(val), formats['num'])

    chart_args = excel_chart_args(inst)
    if not chart_args:
        return False, 'no exact support'
    exact_subtypes_unsupported = {'Pie of Pie', 'Bar of Pie', 'Filled Map', 'Contour'}
    if inst['chart']['excel_subtype'] in exact_subtypes_unsupported:
        if inst['chart']['excel_family'] == 'Pie':
            chart_args = {'type': 'pie'}
        else:
            return False, 'subtype not exactly supported'
    chart = workbook.add_chart(chart_args)
    for c_idx, s in enumerate(series, start=1):
        series_kwargs = {
            'name': [sheet_name, start_row, start_col + c_idx],
            'categories': [sheet_name, start_row + 1, start_col, start_row + len(categories), start_col],
            'values': [sheet_name, start_row + 1, start_col + c_idx, start_row + len(categories), start_col + c_idx],
        }
        series_kwargs.update(
            series_style(
                chart_args['type'],
                c_idx - 1,
                chart_args.get('subtype'),
                bool(chart_args.get('has_markers')),
            )
        )
        if chart_args['type'] in {'pie', 'doughnut'}:
            series_kwargs['points'] = point_styles_for_categories(len(categories))
        if inst['labels']['data_labels']:
            if chart_args['type'] in {'pie', 'doughnut'}:
                series_kwargs['data_labels'] = {'category': True, 'value': True}
            else:
                series_kwargs['data_labels'] = data_label_options(chart_args['type'])
        chart.add_series(series_kwargs)
    chart.set_title({'name': inst['title']})
    if inst['chart']['excel_family'] in {'Pie', 'Doughnut'}:
        chart.set_legend({'none': True})
    else:
        apply_legend_for_series_count(chart, inst['xlsxwriter']['legend_position'], len(series))
        y_axis_extra: Dict[str, Any] = {}
        if inst['chart']['stack_mode'] == 'percent':
            y_axis_extra['min'] = 0
            y_axis_extra['max'] = 1
            y_axis_extra['num_format'] = '0%'
        chart.set_x_axis(excel_axis_options(inst['encoding']['x']['title']))
        chart.set_y_axis(excel_axis_options(inst['encoding']['y']['title'], y_axis_extra))
    chart.set_style(inst['xlsxwriter']['style'])
    chart.set_size(inst['xlsxwriter']['size'])
    chart_anchor = f'G{chart_row + 1}'
    ws.insert_chart(chart_anchor, chart)
    if inst['chart']['excel_family'] == 'Doughnut' and len(series) > 1:
        insert_doughnut_ring_labels(ws, chart_anchor, series)
    return True, 'exact or near-exact render'


def render_scatter(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    headers = ['x', 'y', 'group']
    for c, h in enumerate(headers):
        ws.write(start_row, c, h, formats['header'])
    for r, rec in enumerate(records, start=1):
        ws.write_number(start_row + r, 0, float(rec['x']), formats['num'])
        ws.write_number(start_row + r, 1, float(rec['y']), formats['num'])
        ws.write(start_row + r, 2, rec['grup'], formats['text'])

    args = {'type': 'scatter', 'subtype': 'marker_only'}
    chart = workbook.add_chart(args)
    groups = list(OrderedDict((r['grup'], None) for r in records).keys())
    for grp in groups:
        grp_rows = [i for i, r in enumerate(records, start=1) if r['grup'] == grp]
        if not grp_rows:
            continue
        # contiguous assumption not guaranteed; write per group helper table
        helper_col = 5 + groups.index(grp) * 3
        ws.write(start_row, helper_col, f'x_{grp}', formats['header'])
        ws.write(start_row, helper_col + 1, f'y_{grp}', formats['header'])
        idx = 1
        for rec in records:
            if rec['grup'] == grp:
                ws.write_number(start_row + idx, helper_col, float(rec['x']), formats['num'])
                ws.write_number(start_row + idx, helper_col + 1, float(rec['y']), formats['num'])
                idx += 1
        chart.add_series({
            'name': grp,
            'categories': [sheet_name, start_row + 1, helper_col, start_row + idx - 1, helper_col],
            'values': [sheet_name, start_row + 1, helper_col + 1, start_row + idx - 1, helper_col + 1],
            **series_style('scatter', groups.index(grp)),
        })
    chart.set_title({'name': inst['title']})
    chart.set_x_axis(excel_axis_options(inst['encoding']['x']['title']))
    chart.set_y_axis(excel_axis_options(inst['encoding']['y']['title']))
    apply_legend_for_series_count(chart, inst['xlsxwriter']['legend_position'], len(groups))
    chart.set_style(inst['xlsxwriter']['style'])
    chart.set_size(inst['xlsxwriter']['size'])
    ws.insert_chart(f'G{chart_row + 1}', chart)
    return True, 'exact render'


def render_combo(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    categories = list(OrderedDict((r['categoria'], None) for r in records).keys())
    ws.write_row(start_row, 0, ['Period', 'Primary', 'Secondary'], formats['header'])
    prim = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Primària'}
    sec = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Secundària'}
    prim = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Primary'} or prim
    sec = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Secondary'} or sec
    for i, cat in enumerate(categories, start=1):
        ws.write(start_row + i, 0, cat, formats['text'])
        ws.write_number(start_row + i, 1, float(prim[cat]), formats['num'])
        ws.write_number(start_row + i, 2, float(sec[cat]), formats['num'])

    col_chart = workbook.add_chart({'type': 'column'})
    col_chart.add_series({
        'name': [sheet_name, start_row, 1],
        'categories': [sheet_name, start_row + 1, 0, start_row + len(categories), 0],
        'values': [sheet_name, start_row + 1, 1, start_row + len(categories), 1],
        **series_style('column', 0),
    })
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'name': [sheet_name, start_row, 2],
        'categories': [sheet_name, start_row + 1, 0, start_row + len(categories), 0],
        'values': [sheet_name, start_row + 1, 2, start_row + len(categories), 2],
        'y2_axis': True,
        **series_style('line', 1),
    })
    col_chart.combine(line_chart)
    col_chart.set_title({'name': inst['title']})
    col_chart.set_x_axis(excel_axis_options(inst['encoding']['x']['title']))
    col_chart.set_y_axis(excel_axis_options(inst['encoding']['y']['title']))
    col_chart.set_y2_axis(excel_axis_options(inst['encoding']['secondary_y']['title']))
    apply_legend_for_series_count(col_chart, inst['xlsxwriter']['legend_position'], 2)
    col_chart.set_style(inst['xlsxwriter']['style'])
    col_chart.set_size(inst['xlsxwriter']['size'])
    ws.insert_chart(f'G{chart_row + 1}', col_chart)
    return True, 'exact render (combo)'


def render_stock(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    headers = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    ws.write_row(start_row, 0, headers, formats['header'])
    for i, rec in enumerate(records, start=1):
        ws.write(start_row + i, 0, rec['data'], formats['text'])
        ws.write_number(start_row + i, 1, rec['obertura'], formats['num'])
        ws.write_number(start_row + i, 2, rec['maxim'], formats['num'])
        ws.write_number(start_row + i, 3, rec['minim'], formats['num'])
        ws.write_number(start_row + i, 4, rec['tancament'], formats['num'])
        ws.write_number(start_row + i, 5, rec['volum'], formats['num'])
    # XlsxWriter exact stock variants are limited; use a close line+column preview.
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'name': 'Close',
        'categories': [sheet_name, start_row + 1, 0, start_row + len(records), 0],
        'values': [sheet_name, start_row + 1, 4, start_row + len(records), 4],
        **series_style('line', 0),
    })
    vol_chart = workbook.add_chart({'type': 'column'})
    vol_chart.add_series({
        'name': 'Volume',
        'categories': [sheet_name, start_row + 1, 0, start_row + len(records), 0],
        'values': [sheet_name, start_row + 1, 5, start_row + len(records), 5],
        'y2_axis': True,
        **series_style('column', 1),
    })
    line_chart.combine(vol_chart)
    line_chart.set_title({'name': inst['title']})
    line_chart.set_x_axis(excel_axis_options('Date'))
    line_chart.set_y_axis(excel_axis_options(inst['encoding']['y']['title']))
    line_chart.set_y2_axis(excel_axis_options('Volume'))
    apply_legend_for_series_count(line_chart, inst['xlsxwriter']['legend_position'], 2)
    line_chart.set_style(inst['xlsxwriter']['style'])
    line_chart.set_size(inst['xlsxwriter']['size'])
    ws.insert_chart(f'H{chart_row + 1}', line_chart)
    return True, 'approximate render (line + volume)'


def render_map(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    ws.write_row(start_row, 0, ['Region', 'Value'], formats['header'])
    for i, rec in enumerate(records, start=1):
        ws.write(start_row + i, 0, rec['regio'], formats['text'])
        ws.write_number(start_row + i, 1, rec['valor'], formats['num'])
    chart = workbook.add_chart({'type': 'bar'})
    chart.add_series({
        'name': inst['encoding']['y']['title'],
        'categories': [sheet_name, start_row + 1, 0, start_row + len(records), 0],
        'values': [sheet_name, start_row + 1, 1, start_row + len(records), 1],
        **series_style('bar', 0),
    })
    chart.set_title({'name': inst['title'] + ' - bar preview'})
    chart.set_x_axis(excel_axis_options(inst['encoding']['y']['title']))
    chart.set_y_axis(excel_axis_options('Region'))
    chart.set_style(inst['xlsxwriter']['style'])
    chart.set_size(inst['xlsxwriter']['size'])
    ws.write(start_row - 1, 0, 'Note: Excel/XlsxWriter does not support Filled Map; a bar-chart preview is shown instead.', formats['note'])
    ws.insert_chart(f'G{chart_row + 1}', chart)
    return True, 'approximate render (bars)'


def render_surface(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    xs = sorted(OrderedDict((r['x'], None) for r in records).keys())
    ys = sorted(OrderedDict((r['y'], None) for r in records).keys())
    ws.write(start_row, 0, 'Y\\X', formats['header'])
    for j, x in enumerate(xs, start=1):
        ws.write_number(start_row, j, x, formats['header'])
    for i, y in enumerate(ys, start=1):
        ws.write_number(start_row + i, 0, y, formats['header'])
        for j, x in enumerate(xs, start=1):
            z = next(r['z'] for r in records if r['x'] == x and r['y'] == y)
            ws.write_number(start_row + i, j, z, formats['num'])
    ws.conditional_format(start_row + 1, 1, start_row + len(ys), len(xs), {
        'type': '3_color_scale',
        'min_color': '#F8696B',
        'mid_color': '#FFEB84',
        'max_color': '#63BE7B',
    })
    ws.write(start_row - 1, 0, 'Note: Excel/XlsxWriter does not support Surface/Contour charts; the grid is shown with a color scale.', formats['note'])
    return True, 'approximate render (cell heatmap)'


def render_instance_sheet(workbook, sheet_name, inst):
    ws = workbook.add_worksheet(sheet_name)
    formats = {
        'title': workbook.add_format({'bold': True, 'font_size': 14}),
        'label': workbook.add_format({'bold': True, 'bg_color': '#D9E2F3'}),
        'text': workbook.add_format({'text_wrap': True}),
        'header': workbook.add_format({'bold': True, 'bg_color': '#D9E2F3', 'border': 1}),
        'num': workbook.add_format({'num_format': '0.0'}),
        'note': workbook.add_format({'italic': True, 'font_color': '#666666'}),
    }
    metadata_end_row = add_metadata(ws, formats, inst, sheet_name)
    data_start_row = metadata_end_row + 2
    chart_start_row = metadata_end_row + 1
    ws.set_zoom(90)
    family = inst['chart']['excel_family']
    fmt = inst['data']['format']
    if fmt == 'scatter':
        return render_scatter(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)
    if fmt == 'combo':
        return render_combo(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)
    if fmt == 'stock':
        return render_stock(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)
    if fmt == 'map':
        return render_map(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)
    if fmt == 'surface':
        return render_surface(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)
    return render_long_chart(workbook, ws, formats, inst, sheet_name, data_start_row, chart_start_row)


def main():
    df = pd.read_csv(INPUT_CSV)
    rows = [clean_row(r) for r in df.to_dict(orient='records')]
    instances = [make_canonical_instance(r) for r in rows]
    sorted_instances = sorted(instances, key=case_sort_key)
    export_instances = [export_instance(instance) for instance in instances]

    with OUT_JSON.open('w', encoding='utf-8') as f:
        json.dump(export_instances, f, ensure_ascii=False, indent=2)

    workbook = xlsxwriter.Workbook(str(OUT_XLSX))
    used = set()
    sheet_map = {}
    render_stats = {'exact': 0, 'approximate_or_metadata': 0}
    detail_counts = {}

    for inst in sorted_instances:
        sheet_name = safe_sheet_name(inst['id'], used)
        sheet_map[inst['id']] = sheet_name
        ok, detail = render_instance_sheet(workbook, sheet_name, inst)
        detail_counts[detail] = detail_counts.get(detail, 0) + 1
        render_stats[inst['chart']['render_mode']] += 1

    add_summary_sheet(workbook, sorted_instances, sheet_map, {**render_stats, **detail_counts})
    workbook.close()

    manifest = {
        'input_rows': len(rows),
        'json': str(OUT_JSON),
        'xlsx': str(OUT_XLSX),
        'style_ref': STYLE_REF,
        'render_stats': render_stats,
        'detail_counts': detail_counts,
        'domains': sorted({str(translate_case_value('semantic_domain', row['domini_semantic'])) for row in rows if row.get('domini_semantic') is not None}),
        'families': dict(OrderedDict((family, sum(1 for row in rows if row.get('familia_excel') == family)) for family in OrderedDict((row.get('familia_excel'), None) for row in rows if row.get('familia_excel') is not None).keys())),
    }
    with OUT_MANIFEST.open('w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print('Generated:')
    print(OUT_JSON)
    print(OUT_XLSX)
    print(OUT_MANIFEST)
    print(render_stats)


if __name__ == '__main__':
    main()
