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

INPUT_CSV = Path('./matriu_500.csv')
OUT_DIR = Path('./sortida_instancies_completa')
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_JSON = OUT_DIR / 'instancies_canoniques.json'
OUT_XLSX = OUT_DIR / 'instancies_canoniques.xlsx'
OUT_MANIFEST = OUT_DIR / 'manifest.json'

STYLE_REF = {
    'id': 'office-custom-theme-pendent-v1',
    'href': './styles/office-custom-theme-pendent-v1.json',
    'description': 'Referència pendent a la definició d’estil corporatiu.'
}

MONTHS = ['Gen', 'Feb', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Des']
QUARTERS = ['T1', 'T2', 'T3', 'T4']
YEARS = [str(y) for y in range(1995, 2026)]
REGIONS_CAT = ['Alt Pirineu', 'Ponent', 'Camp de Tarragona', 'Barcelona Metropolitana', 'Girona', 'Terres de l\'Ebre']

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
    {'id': 'P1', 'name': 'diagonal-right', 'token': 'slash', 'description': 'Ratlles diagonals cap a la dreta'},
    {'id': 'P2', 'name': 'diagonal-left', 'token': 'backslash', 'description': 'Ratlles diagonals cap a l\'esquerra'},
    {'id': 'P3', 'name': 'vertical', 'token': 'vertical', 'description': 'Línies verticals'},
    {'id': 'P4', 'name': 'horizontal', 'token': 'horizontal', 'description': 'Línies horitzontals'},
    {'id': 'P5', 'name': 'dots', 'token': 'dots', 'description': 'Punts separats'},
    {'id': 'P6', 'name': 'cross', 'token': 'cross', 'description': 'Creu simple'},
    {'id': 'P7', 'name': 'diag-cross', 'token': 'diag-cross', 'description': 'Creu diagonal'},
    {'id': 'P8', 'name': 'grid', 'token': 'grid', 'description': 'Quadriculat'},
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
    'clima': {
        'unit_abs': 'ktCO₂e',
        'unit_pct': '% de contribució',
        'simple_categories': ['Transport', 'Edificis', 'Indústria', 'Residus', 'Agricultura', 'Energia', 'Aviació', 'Ramaderia', 'Usos del sòl', 'Forestació', 'Processos industrials', 'Pesca'],
        'long_categories': [
            'Transport urbà i interurbà de passatgers',
            'Parc d\'edificis residencials antics',
            'Polígons industrials de consum intensiu',
            'Gestió de residus municipals barrejats',
            'Activitat agroalimentària de proximitat',
            'Flota de vehicles pesants de mercaderies',
            'Climatització d\'espais terciaris i comercials',
            'Generació elèctrica amb combustibles fòssils',
            'Ramaderia extensiva i producció de purins',
            'Extracció i usos industrials de l\'aigua'
        ],
        'geo_categories': ['Pirineu', 'Pla de Lleida', 'Camp de Tarragona', 'AMB', 'Comarques Gironines', 'Terres de l\'Ebre', 'Catalunya Central', 'Alt Empordà'],
        'series': ['CO₂', 'Metà', 'Òxid nitrós'],
        'radar_axes': ['Mitigació', 'Adaptació', 'Exposició', 'Resiliència', 'Eficiència', 'Seguiment', 'Governança', 'Impacte social'],
        'scatter_x': ('Temperatura mitjana (°C)', 10, 28),
        'scatter_y': ('Emissions per càpita (tCO₂e)', 2, 14),
        'map_metric': 'Emissions per càpita (tCO₂e)',
        'combo_primary': 'Emissions mensuals (ktCO₂e)',
        'combo_secondary': 'Anomalia de temperatura (°C)',
        'stock_label': 'Preu del futur d\'electricitat (€/MWh)'
    },
    'demografia': {
        'unit_abs': 'milers de persones',
        'unit_pct': '% del total',
        'simple_categories': ['Infància', 'Joventut', 'Adults', 'Sèniors', 'Dependència', 'Migració', 'Habitatge', 'Emancipació', 'Diversitat', 'Nova ciutadania', 'Famílies', 'Soledat'],
        'long_categories': [
            'Població de 0 a 14 anys',
            'Població de 15 a 29 anys',
            'Població de 30 a 44 anys',
            'Població de 45 a 64 anys',
            'Població de 65 anys o més',
            'Persones nascudes fora de la Unió Europea',
            'Llars unipersonals en entorn urbà consolidat',
            'Joves en procés d\'emancipació tardana',
            'Persones amb diversitat funcional reconeguda',
            'Fills en unitats familiars monoparentals'
        ],
        'geo_categories': ['Vall d\'Aran', 'Ponent', 'Comarques Centrals', 'Barcelona', 'Camp de Tarragona', 'Girona', 'Terres de l\'Ebre', 'Pirineu'],
        'series': ['Homes', 'Dones', 'No especificat'],
        'radar_axes': ['Creixement', 'Envelliment', 'Densitat', 'Mobilitat', 'Natalitat', 'Atracció', 'Equilibri territorial', 'Relleu generacional'],
        'scatter_x': ('Densitat (hab/km²)', 30, 600),
        'scatter_y': ('Creixement anual (%)', -1.0, 3.0),
        'map_metric': 'Densitat de població (hab/km²)',
        'combo_primary': 'Naixements mensuals',
        'combo_secondary': 'Saldo migratori (%)',
        'stock_label': 'Índex de preu de l’habitatge'
    },
    'educació': {
        'unit_abs': 'milers d\'alumnes',
        'unit_pct': '% d\'alumnat',
        'simple_categories': ['Primària', 'ESO', 'Batxillerat', 'FP', 'Universitat', 'Infantil', 'FP grau superior', 'Doctorat', 'Adults', 'Educació especial', 'Idiomes', 'Informàtica bàsica'],
        'long_categories': [
            'Alumnat de primària en centres públics',
            'Alumnat d\'ESO en entorns vulnerables',
            'Alumnat de batxillerat científic-tecnològic',
            'Alumnat de formació professional dual',
            'Alumnat universitari de primer curs',
            'Menors matriculats a escoletes de 0-3 anys',
            'Alumnat de cicles formatius de grau superior',
            'Doctorands en universitats públiques catalanes',
            'Persones adultes en programes de requalificació',
            'Alumnat amb necessitats educatives específiques'
        ],
        'geo_categories': ['Girona', 'Barcelona', 'Lleida', 'Tarragona', 'Terres de l\'Ebre', 'Pirineu', 'Alt Pirineu', 'Comarques Centrals'],
        'series': ['Públic', 'Concertat', 'Privat'],
        'radar_axes': ['Matemàtiques', 'Lectura', 'Ciències', 'Digital', 'Idiomes', 'Participació', 'Equitat', 'Continuïtat acadèmica'],
        'scatter_x': ('Ràtio alumnes/docent', 8, 32),
        'scatter_y': ('Puntuació mitjana', 50, 95),
        'map_metric': 'Abandonament escolar (%)',
        'combo_primary': 'Matrícules mensuals',
        'combo_secondary': 'Ràtio de cobertura (%)',
        'stock_label': 'Índex de cost educatiu'
    },
    'energia': {
        'unit_abs': 'GWh',
        'unit_pct': '% de generació',
        'simple_categories': ['Nuclear', 'Hidràulica', 'Eòlica', 'Solar', 'Gas', 'Biomassa', 'Marina', 'Geotèrmica', 'Cogeneració', 'Residus', 'Hidrogen', 'Xarxa intel·ligent'],
        'long_categories': [
            'Central nuclear d\'Ascó',
            'Embassaments de regulació anual',
            'Parcs eòlics de la Terra Alta',
            'Instal·lacions fotovoltaiques del Segrià',
            'Cicles combinats de suport',
            'Plantes de biogàs agroindustrial',
            'Plataformes d\'energia marina experimental',
            'Bombes de calor geotèrmiques profundes',
            'Unitats de cogeneració industrial eficient',
            'Valorització energètica de residus sòlids urbans'
        ],
        'geo_categories': ['Pirineu', 'Ponents', 'Camp de Tarragona', 'Àrea Metropolitana', 'Catalunya Central', 'Ebre', 'Alt Pirineu', 'Comarques Gironines'],
        'series': ['Base', 'Renovable', 'Suport'],
        'radar_axes': ['Disponibilitat', 'Cost', 'Emissions', 'Flexibilitat', 'Risc', 'Cobertura', 'Autonomia', 'Estabilitat de xarxa'],
        'scatter_x': ('Potència instal·lada (MW)', 20, 1200),
        'scatter_y': ('Producció anual (GWh)', 10, 6500),
        'map_metric': 'Generació renovable per habitant (MWh)',
        'combo_primary': 'Generació mensual (GWh)',
        'combo_secondary': 'Preu pool (€/MWh)',
        'stock_label': 'Preu del futur del gas (€/MWh)'
    },
    'finances': {
        'unit_abs': 'M€',
        'unit_pct': '% del total',
        'simple_categories': ['Ingressos', 'Costos', 'Marge', 'Capex', 'Deute', 'Dividends', 'Taxes', 'Reserves', 'Amortitzacions', 'Circulant', 'Tresoreria', 'Patrimoni'],
        'long_categories': [
            'Ingressos recurrents de subscripció',
            'Cost comercial i captació',
            'Marge operatiu abans d\'amortitzacions',
            'Inversió en plataforma tecnològica',
            'Deute financer net ajustat',
            'Dividends distribuïts als accionistes',
            'Càrrega fiscal efectiva sobre el resultat',
            'Reserves per a contingències futures',
            'Amortització d\'actius intangibles clau',
            'Fons de maniobra i posició de tresoreria'
        ],
        'geo_categories': ['Nord-est', 'Centre', 'Llevant', 'Sud', 'Portugal', 'Internacional', 'Àsia-Pacífic', 'Amèrica Llatina'],
        'series': ['Real', 'Pressupost', 'Objectiu'],
        'radar_axes': ['Liquiditat', 'Rendibilitat', 'Creixement', 'Solvència', 'Eficiència', 'Conversió', 'Endeutament', 'Visibilitat futura'],
        'scatter_x': ('Risc (volatilitat %)', 5, 40),
        'scatter_y': ('Rendiment anual (%)', -10, 35),
        'map_metric': 'Ingressos per província (M€)',
        'combo_primary': 'Facturació mensual (M€)',
        'combo_secondary': 'Marge EBITDA (%)',
        'stock_label': 'Preu de l\'acció (€)'
    },
    'geografia': {
        'unit_abs': 'índex territorial',
        'unit_pct': '% de superfície',
        'simple_categories': ['Costa', 'Plana', 'Prelitoral', 'Muntanya', 'Urbà', 'Illes', 'Riu', 'Secà', 'Aiguamolls', 'Bosc', 'Conreus', 'Delta'],
        'long_categories': [
            'Sòl urbà consolidat de densitat alta',
            'Espais agraris de regadiu intensiu',
            'Zones forestals de mitja muntanya',
            'Corredors litorals d\'ús turístic',
            'Àrees naturals protegides d\'alta fragilitat',
            'Zones humides de valor ecològic especial',
            'Polígons industrials periurbans d\'expansió',
            'Àrees d\'interès agrícola preferent protegides',
            'Espais fluvials de gestió integrada',
            'Paisatges rurals de valor patrimonial reconegut'
        ],
        'geo_categories': ['Val d\'Aran', 'Pallars', 'Segrià', 'Barcelonès', 'Empordà', 'Delta de l\'Ebre', 'Conca de Barberà', 'Pla de l\'Estany'],
        'series': ['Coberta', 'Ús urbà', 'Protegit'],
        'radar_axes': ['Accessibilitat', 'Pendents', 'Densitat', 'Equipaments', 'Risc', 'Atracció', 'Connectivitat', 'Pressió urbana'],
        'scatter_x': ('Altitud mitjana (m)', 0, 2200),
        'scatter_y': ('Densitat de població (hab/km²)', 5, 1500),
        'map_metric': 'Densitat de població (hab/km²)',
        'combo_primary': 'Flux turístic mensual (milers)',
        'combo_secondary': 'Ocupació (%)',
        'stock_label': 'Índex territorial'
    },
    'manufactura': {
        'unit_abs': 'milers d\'unitats',
        'unit_pct': '% de producció',
        'simple_categories': ['Línia A', 'Línia B', 'Línia C', 'Línia D', 'Retraball', 'Qualitat', 'Embalatge', 'Pintura', 'Proves', 'Ensamblatge', 'Accessoris', 'Manteniment'],
        'long_categories': [
            'Línia d\'assemblatge d\'electrònica fina',
            'Línia de mecanitzat de peces crítiques',
            'Cel·la robotitzada d\'acabats superficials',
            'Àrea de validació i control dimensional',
            'Circuit de retraball i no conformitats',
            'Secció d\'embalatge i preparació d\'expedició',
            'Cabina de pintura en pols electrostàtica',
            'Banc de proves i validació funcional',
            'Zona d\'ensamblatge d\'equips finals',
            'Taller de manteniment preventiu i predictiu'
        ],
        'geo_categories': ['Planta Nord', 'Planta Centre', 'Planta Sud', 'Magatzem tècnic', 'Laboratori', 'Subcontracte', 'Planta Est', 'Plataforma logística'],
        'series': ['Bona', 'Reprocessada', 'Descartada'],
        'radar_axes': ['Qualitat', 'Productivitat', 'Seguretat', 'Manteniment', 'Cost', 'Flexibilitat', 'Traçabilitat', 'Robustesa'],
        'scatter_x': ('Temps de cicle (s)', 20, 220),
        'scatter_y': ('Defectes per milió', 30, 1800),
        'map_metric': 'Producció per planta (milers d\'unitats)',
        'combo_primary': 'Producció mensual (milers)',
        'combo_secondary': 'OEE (%)',
        'stock_label': 'Índex de cost industrial'
    },
    'operacions': {
        'unit_abs': 'milers de comandes',
        'unit_pct': '% d\'activitat',
        'simple_categories': ['Recepció', 'Preparació', 'Expedició', 'Qualitat', 'Devolucions', 'Aprovisionament', 'Atenció al client', 'Facturació', 'Planificació', 'Transport', 'Magatzem', 'Retorn'],
        'long_categories': [
            'Recepció de mercaderia de proveïdors',
            'Preparació de comandes e-commerce',
            'Expedició cap a botigues físiques',
            'Control de qualitat de sortida',
            'Gestió de devolucions i incidències',
            'Aprovisionament de materials estratègics',
            'Atenció al client postvenda multicanal',
            'Facturació i conciliació de comandes',
            'Planificació de la demanda mensual',
            'Gestió del transport de llarg recorregut'
        ],
        'geo_categories': ['Hub Barcelona', 'Hub Girona', 'Hub Lleida', 'Hub Tarragona', 'Cross-dock', 'Última milla', 'Centre de distribució', 'Plataforma logística'],
        'series': ['Planificat', 'Executat', 'Incidències'],
        'radar_axes': ['Servei', 'Temps', 'Cost', 'Capacitat', 'Qualitat', 'Flexibilitat', 'Puntualitat', 'Escalabilitat'],
        'scatter_x': ('Temps de preparació (min)', 5, 120),
        'scatter_y': ('Comandes/hora', 20, 400),
        'map_metric': 'Lliuraments a temps (%)',
        'combo_primary': 'Comandes mensuals (milers)',
        'combo_secondary': 'Puntualitat (%)',
        'stock_label': 'Índex de cost logístic'
    },
    'salut': {
        'unit_abs': 'milers de pacients',
        'unit_pct': '% de casos',
        'simple_categories': ['Atenció primària', 'Urgències', 'Hospitalització', 'Diagnòstic', 'Rehabilitació', 'Salut mental', 'Farmàcia', 'Oncologia', 'Pediatria', 'Geriatria', 'Cirurgia', 'Cardiologia'],
        'long_categories': [
            'Pacients amb seguiment telemàtic postoperatori',
            'Casos atesos al circuit ràpid de diagnòstic',
            'Ingressos de medicina interna d\'alta complexitat',
            'Programa de prevenció cardiovascular comunitària',
            'Unitat de rehabilitació funcional intensiva',
            'Consultes de salut mental en atenció primària',
            'Dispensació farmacèutica de medicació crònica',
            'Pacients en tractament oncològic actiu',
            'Visites pediàtriques en centres d\'atenció primària',
            'Atenció geriàtrica en residències i domicili'
        ],
        'geo_categories': ['Barcelona', 'Girona', 'Lleida', 'Tarragona', 'Catalunya Central', 'Terres de l\'Ebre', 'Alt Pirineu', 'Vallès Occidental'],
        'series': ['Resolts', 'Amb seguiment', 'Derivats'],
        'radar_axes': ['Accessibilitat', 'Qualitat', 'Temps d\'espera', 'Resolució', 'Satisfacció', 'Prevenció', 'Continuïtat assistencial', 'Cobertura'],
        'scatter_x': ('Dies d\'espera', 2, 180),
        'scatter_y': ('Casos resolts (%)', 40, 98),
        'map_metric': 'Llista d\'espera mitjana (dies)',
        'combo_primary': 'Visites mensuals (milers)',
        'combo_secondary': 'Temps mitjà d\'espera (dies)',
        'stock_label': 'Índex de cost sanitari'
    },
    'vendes': {
        'unit_abs': 'M€',
        'unit_pct': '% de vendes',
        'simple_categories': ['Botiga física', 'E-commerce', 'Distribuïdor', 'Marketplace', 'Canal telefònic', 'Direct-to-consumer', 'Franquícia', 'Subscripció', 'Export', 'Licitació', 'Majorista', 'Flash sale'],
        'long_categories': [
            'Electrodomèstics premium per a la llar',
            'Accessoris connectats per consum domèstic',
            'Petits aparells de cuina eficients',
            'Servei de manteniment i ampliació de garantia',
            'Canal professional d\'instal·ladors',
            'Venda directa al consumidor final en línia',
            'Operadors franquiciats de la xarxa regional',
            'Model de subscripció de servei recurrent',
            'Exportació a mercats europeus emergents',
            'Licitació pública per a grans contractes'
        ],
        'geo_categories': ['Nord', 'Centre', 'Llevant', 'Sud', 'Balears', 'Portugal', 'Canàries', 'Internacional'],
        'series': ['Nord', 'Centre', 'Sud'],
        'radar_axes': ['Volum', 'Marge', 'Rotació', 'Fidelització', 'Ticket mitjà', 'Conversió', 'Penetració', 'Recurrència'],
        'scatter_x': ('Descompte mitjà (%)', 0, 35),
        'scatter_y': ('Vendes per punt (k€)', 20, 500),
        'map_metric': 'Vendes per habitant (€)',
        'combo_primary': 'Facturació mensual (M€)',
        'combo_secondary': 'Marge brut (%)',
        'stock_label': 'Índex de preu al detall'
    },
    'web analytics': {
        'unit_abs': 'milers de sessions',
        'unit_pct': '% de sessions',
        'simple_categories': ['Directe', 'Orgànic', 'Social', 'Email', 'Referència', 'Paid Search', 'Afiliació', 'Display', 'Push', 'App', 'QR', 'Vídeo'],
        'long_categories': [
            'Campanya de captació de leads B2B per cercadors',
            'Programa de fidelització mitjançant email segmentat',
            'Promoció de temporada en xarxes socials',
            'Trànsit de marca des de comparadors externs',
            'Visites recurrents des de l\'aplicació mòbil',
            'Campanya de display en xarxes de contingut',
            'Pàgines d\'aterratge de publicitat de pagament',
            'Trànsit provinent de lectures en blogs afiliats',
            'Notificacions push per a usuaris recurrents',
            'Trànsit des de codis QR en publicitat física'
        ],
        'geo_categories': ['SEO', 'SEM', 'CRM', 'Social', 'Afiliació', 'App', 'Display', 'Referral'],
        'series': ['Rebot baix', 'Navegació mitjana', 'Conversió'],
        'radar_axes': ['Captació', 'Conversió', 'Retenció', 'Profunditat', 'Velocitat', 'ROI', 'Engagement', 'Qualitat del trànsit'],
        'scatter_x': ('Temps de càrrega (s)', 0.5, 5.0),
        'scatter_y': ('Conversió (%)', 0.2, 8.0),
        'map_metric': 'Sessions per territori (milers)',
        'combo_primary': 'Sessions mensuals (milers)',
        'combo_secondary': 'Conversió (%)',
        'stock_label': 'Índex de CPC (€)'
    },
}

DOMAIN_VALUE_RANGE_CFG: Dict[str, Dict[str, Dict[str, List[Tuple[float, float]]]]] = {
    'clima': {
        'category_value_ranges': {
            'baixa': [(110.0, 240.0), (140.0, 280.0)],
            'mitjana': [(70.0, 180.0), (90.0, 210.0)],
            'alta': [(40.0, 120.0), (55.0, 145.0)],
            'molt alta': [(22.0, 80.0), (30.0, 95.0)],
        },
        'map_value_ranges': {
            'baixa': [(6.5, 12.0), (5.5, 10.5)],
            'mitjana': [(4.5, 9.0), (3.8, 7.8)],
            'alta': [(3.0, 6.5), (2.6, 5.5)],
            'molt alta': [(2.0, 4.8), (1.8, 4.0)],
        },
    },
    'demografia': {
        'category_value_ranges': {
            'baixa': [(220.0, 950.0), (180.0, 820.0)],
            'mitjana': [(140.0, 700.0), (110.0, 580.0)],
            'alta': [(90.0, 460.0), (70.0, 360.0)],
            'molt alta': [(45.0, 240.0), (35.0, 190.0)],
        },
        'map_value_ranges': {
            'baixa': [(220.0, 620.0), (180.0, 520.0)],
            'mitjana': [(140.0, 420.0), (110.0, 340.0)],
            'alta': [(80.0, 260.0), (60.0, 220.0)],
            'molt alta': [(35.0, 160.0), (25.0, 120.0)],
        },
    },
    'educació': {
        'category_value_ranges': {
            'baixa': [(80.0, 360.0), (60.0, 280.0)],
            'mitjana': [(50.0, 240.0), (35.0, 190.0)],
            'alta': [(25.0, 130.0), (18.0, 100.0)],
            'molt alta': [(10.0, 70.0), (8.0, 55.0)],
        },
        'map_value_ranges': {
            'baixa': [(14.0, 28.0), (11.0, 24.0)],
            'mitjana': [(9.0, 21.0), (7.0, 17.0)],
            'alta': [(6.0, 15.0), (4.5, 12.0)],
            'molt alta': [(3.0, 9.0), (2.5, 7.5)],
        },
    },
    'energia': {
        'category_value_ranges': {
            'baixa': [(600.0, 4200.0), (450.0, 3200.0)],
            'mitjana': [(300.0, 2500.0), (220.0, 1800.0)],
            'alta': [(140.0, 1200.0), (100.0, 900.0)],
            'molt alta': [(60.0, 520.0), (40.0, 380.0)],
        },
        'map_value_ranges': {
            'baixa': [(2.0, 8.5), (1.5, 6.5)],
            'mitjana': [(1.2, 5.8), (0.9, 4.5)],
            'alta': [(0.7, 3.8), (0.5, 2.8)],
            'molt alta': [(0.3, 1.8), (0.2, 1.3)],
        },
    },
    'finances': {
        'category_value_ranges': {
            'baixa': [(40.0, 240.0), (30.0, 180.0)],
            'mitjana': [(22.0, 150.0), (16.0, 110.0)],
            'alta': [(10.0, 85.0), (8.0, 65.0)],
            'molt alta': [(4.0, 38.0), (3.0, 28.0)],
        },
        'map_value_ranges': {
            'baixa': [(60.0, 380.0), (45.0, 300.0)],
            'mitjana': [(35.0, 220.0), (25.0, 170.0)],
            'alta': [(18.0, 120.0), (12.0, 90.0)],
            'molt alta': [(7.0, 55.0), (5.0, 40.0)],
        },
    },
    'geografia': {
        'category_value_ranges': {
            'baixa': [(45.0, 95.0), (35.0, 85.0)],
            'mitjana': [(28.0, 78.0), (22.0, 68.0)],
            'alta': [(16.0, 58.0), (12.0, 48.0)],
            'molt alta': [(8.0, 34.0), (6.0, 26.0)],
        },
        'map_value_ranges': {
            'baixa': [(250.0, 1500.0), (180.0, 1200.0)],
            'mitjana': [(120.0, 900.0), (90.0, 700.0)],
            'alta': [(55.0, 420.0), (40.0, 320.0)],
            'molt alta': [(12.0, 180.0), (8.0, 130.0)],
        },
    },
    'manufactura': {
        'category_value_ranges': {
            'baixa': [(40.0, 240.0), (30.0, 180.0)],
            'mitjana': [(24.0, 155.0), (18.0, 120.0)],
            'alta': [(12.0, 82.0), (9.0, 64.0)],
            'molt alta': [(5.0, 34.0), (4.0, 26.0)],
        },
        'map_value_ranges': {
            'baixa': [(50.0, 260.0), (40.0, 210.0)],
            'mitjana': [(28.0, 170.0), (22.0, 130.0)],
            'alta': [(14.0, 90.0), (10.0, 70.0)],
            'molt alta': [(6.0, 40.0), (4.0, 28.0)],
        },
    },
    'operacions': {
        'category_value_ranges': {
            'baixa': [(24.0, 180.0), (18.0, 140.0)],
            'mitjana': [(14.0, 110.0), (10.0, 82.0)],
            'alta': [(7.0, 56.0), (5.0, 42.0)],
            'molt alta': [(3.0, 24.0), (2.0, 18.0)],
        },
        'map_value_ranges': {
            'baixa': [(90.0, 99.0), (86.0, 97.0)],
            'mitjana': [(82.0, 96.0), (78.0, 94.0)],
            'alta': [(74.0, 92.0), (70.0, 88.0)],
            'molt alta': [(66.0, 84.0), (62.0, 80.0)],
        },
    },
    'salut': {
        'category_value_ranges': {
            'baixa': [(60.0, 380.0), (45.0, 290.0)],
            'mitjana': [(35.0, 220.0), (25.0, 170.0)],
            'alta': [(18.0, 120.0), (12.0, 90.0)],
            'molt alta': [(7.0, 55.0), (5.0, 42.0)],
        },
        'map_value_ranges': {
            'baixa': [(35.0, 140.0), (28.0, 120.0)],
            'mitjana': [(22.0, 100.0), (16.0, 80.0)],
            'alta': [(12.0, 60.0), (9.0, 46.0)],
            'molt alta': [(5.0, 28.0), (4.0, 22.0)],
        },
    },
    'vendes': {
        'category_value_ranges': {
            'baixa': [(25.0, 180.0), (18.0, 140.0)],
            'mitjana': [(14.0, 105.0), (10.0, 82.0)],
            'alta': [(7.0, 56.0), (5.0, 44.0)],
            'molt alta': [(3.0, 26.0), (2.0, 18.0)],
        },
        'map_value_ranges': {
            'baixa': [(650.0, 1600.0), (500.0, 1350.0)],
            'mitjana': [(380.0, 1200.0), (300.0, 920.0)],
            'alta': [(180.0, 720.0), (140.0, 560.0)],
            'molt alta': [(80.0, 340.0), (60.0, 250.0)],
        },
    },
    'web analytics': {
        'category_value_ranges': {
            'baixa': [(140.0, 900.0), (100.0, 700.0)],
            'mitjana': [(80.0, 560.0), (60.0, 420.0)],
            'alta': [(35.0, 260.0), (25.0, 190.0)],
            'molt alta': [(12.0, 110.0), (8.0, 80.0)],
        },
        'map_value_ranges': {
            'baixa': [(90.0, 450.0), (70.0, 360.0)],
            'mitjana': [(55.0, 260.0), (40.0, 210.0)],
            'alta': [(24.0, 130.0), (18.0, 95.0)],
            'molt alta': [(8.0, 55.0), (6.0, 40.0)],
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
        return {'baixa': 16, 'mitjana': 24, 'alta': 32, 'molt alta': 40}.get(normalize_difficulty(difficulty), 24)
    intervals: Dict[str, Tuple[int, int]] = {
        'baixa':     (2,  6),
        'mitjana':   (7,  12),
        'alta':      (13, 18),
        'molt alta': (19, 30),
    }
    lo, hi = intervals.get(normalize_difficulty(difficulty), (7, 12))
    return rnd.randint(lo, hi) if rnd is not None else random.randint(lo, hi)


def radar_axis_count(difficulty: str, max_axes: int) -> int:
    target = {
        'baixa': 4,
        'mitjana': 5,
        'alta': 6,
        'molt alta': 8,
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


def choose_categories(cfg: Dict[str, Any], structure: str, count: int) -> List[str]:
    if structure == 'categòric amb etiquetes llargues':
        return cfg['long_categories'][:count]
    if structure == 'categòric ordenat':
        return ['Molt baix', 'Baix', 'Mitjà', 'Alt', 'Molt alt', 'Molt alt +'][0:count]
    if structure == 'geogràfica':
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
        'baixa': 3,
        'mitjana': 5,
        'alta': 8,
        'molt alta': 12,
    }.get(normalize_difficulty(difficulty), 5)
    return max(3, min(target, max_segments))


def difficulty_stacked_series_count(difficulty: str, max_series: int) -> int:
    target = {
        'baixa': 2,
        'mitjana': 3,
        'alta': 4,
        'molt alta': 5,
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
        'Altres',
        'Residual',
        'Complementària',
    ]
    series_pool = unique_preserve_order(base_series + extras)
    count = difficulty_stacked_series_count(difficulty, len(series_pool))
    return series_pool[:count]


def choose_series(cfg: Dict[str, Any], row: Dict[str, Any], structure: str) -> List[str]:
    subtype = str(row['subtipus_excel'])
    family = str(row['familia_excel'])
    if structure in {'sèrie temporal', 'sèrie temporal discreta', 'categòric simple', 'categòric ordenat', 'geogràfica'}:
        if family in {'Pie', 'Doughnut'}:
            return ['Pes']
        if '100% Stacked' in subtype and family in {'Bar', 'Column', 'Area', 'Line'}:
            if structure == 'categòric simple':
                return ['Pes relatiu']
        return ['Valor']
    if structure == 'parts del tot multianell':
        return ['Actual', 'Objectiu']
    if family == 'Area' and 'stacked' in subtype.lower():
        return choose_stacked_area_series(cfg, str(row['dificultat']))
    if 'multiserie' in structure or structure in {'multivariable sobre eixos comuns', 'temporal amb eix secundari'}:
        return cfg['series']
    return cfg['series']


def line_pattern(values_n: int, pattern: str, base: float, spread: float, rnd: random.Random) -> List[float]:
    xs = list(range(values_n))
    vals: List[float] = []
    for i in xs:
        if pattern in {'creixement', 'creixement amb taxa'}:
            v = base + spread * (i / max(1, values_n - 1)) * (1.1 if pattern == 'creixement amb taxa' else 1.0)
            v += rnd.uniform(-0.06, 0.06) * spread
        elif pattern == 'decreixement':
            v = base + spread * (1 - i / max(1, values_n - 1))
            v += rnd.uniform(-0.06, 0.06) * spread
        elif pattern == 'estacionalitat':
            v = base + spread * (0.5 + 0.45 * math.sin(2 * math.pi * i / max(3, values_n)))
            v += rnd.uniform(-0.04, 0.04) * spread
        elif pattern == 'pic local':
            peak = values_n // 2
            v = base + spread * (0.25 + 0.75 * math.exp(-((i - peak) ** 2) / max(1, values_n / 3)))
        elif pattern == 'pic sobtat':
            peak = rnd.randint(max(1, values_n // 3), max(1, values_n - 2))
            v = base + spread * (0.25 + (1.2 if i == peak else 0.15))
        elif pattern == 'vall sobtada':
            dip = rnd.randint(max(1, values_n // 3), max(1, values_n - 2))
            v = base + spread * (0.8 if i != dip else 0.15)
        elif pattern == 'canvi de règim':
            cut = max(1, values_n // 2)
            v = base + (0.35 if i < cut else 0.85) * spread + rnd.uniform(-0.05, 0.05) * spread
        elif pattern == 'baixa volatilitat':
            v = base + 0.1 * spread + rnd.uniform(-0.03, 0.03) * spread
        elif pattern == 'soroll alt':
            v = base + 0.5 * spread + rnd.uniform(-0.4, 0.4) * spread
        elif pattern == 'valors negatius':
            v = base - spread / 2 + spread * (i / max(1, values_n - 1)) + rnd.uniform(-0.12, 0.12) * spread
        elif pattern == 'valors molt propers':
            center = base + 0.5 * spread
            v = center + rnd.uniform(-0.05, 0.05) * center
        elif pattern == 'diferències clares':
            v = base + (0.15 + 0.8 * i / max(1, values_n - 1)) * spread
        else:
            v = base + spread * (0.45 + 0.12 * math.sin(i))
        vals.append(round(v, 1))
    return vals


def proportions(n: int, pattern: str, rnd: random.Random) -> List[float]:
    if pattern in {'parts equilibrades', 'perfil equilibrat', 'valors molt propers', 'segments molt propers', 'distribució homogènia'}:
        raw = [1 + rnd.uniform(-0.06, 0.06) for _ in range(n)]
    elif pattern in {'part dominant', 'cua llarga'}:
        raw = [n * 2.5] + [max(0.2, 1 / (i + 1) + rnd.uniform(0, 0.2)) for i in range(1, n)]
    elif pattern == 'dues regions dominants':
        raw = [n * 1.8, n * 1.6] + [max(0.3, 0.8 + rnd.uniform(-0.1, 0.15)) for _ in range(max(0, n - 2))]
    elif pattern in {'perfil espigat'}:
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
            axis_unit = axis_unit or 'data'
        elif axis_type == 'nominal':
            axis_unit = axis_unit or 'categoria'

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
            'inici': axis.get('min'),
            'final': axis.get('max'),
            'unitat': axis.get('unit'),
            'interval': axis.get('interval'),
        }
    return {
        'titol_grafic': title,
        'eixos': axes_payload,
    }


def count_series_in_records(records: List[Dict[str, Any]]) -> int:
    if not records:
        return 0
    if 'serie' in records[0]:
        return len(OrderedDict((r['serie'], None) for r in records))
    if 'grup' in records[0]:
        return len(OrderedDict((r['grup'], None) for r in records))
    return 1


def generate_time_long(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    structure = row['estructura_dades']
    family = row['familia_excel']
    pattern = row['patro_estadistic']
    count = difficulty_count(row['dificultat'], 'time', rnd)
    categories = choose_categories(cfg, structure, count)
    series = choose_series(cfg, row, structure)
    share_like = ('100% Stacked' in row['subtipus_excel']) or family in {'Pie', 'Doughnut'} or 'parts del tot' in structure

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
        value_range = choose_value_range(cfg, row['dificultat'], 'category_value_ranges', rnd)
        for s_idx, ser in enumerate(series):
            lo, hi, base, spread = base_spread_for_value_range(value_range, rnd, s_idx, len(series))
            vals = line_pattern(len(categories), pattern, base, spread, rnd)
            if '100% Stacked' in row['subtipus_excel'] and len(series) == 1:
                vals = scale_to_percent(vals)
                unit = cfg['unit_pct']
                y_title = cfg['unit_pct']
            else:
                vals = clamp_series_values(vals, lo, hi, allow_negative=pattern == 'valors negatius')
            for cat, v in zip(categories, vals):
                records.append({'categoria': cat, 'serie': ser, 'valor': float(v)})
    meta = {
        'x_title': 'Període' if 'temporal' in structure else 'Categoria',
        'y_title': y_title,
        'unit': unit,
        'series': series,
        'categories': categories,
    }
    return records, meta


def generate_parts(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    if row['estructura_dades'] == 'parts del tot multianell':
        categories = choose_parts_categories(cfg, str(row['dificultat']))
        series = ['Actual', 'Objectiu']
        records = []
        for ser in series:
            vals = proportions(len(categories), row['patro_estadistic'], rnd)
            for cat, v in zip(categories, vals):
                records.append({'categoria': cat, 'serie': ser, 'valor': v})
        return records, {
            'x_title': 'Categoria',
            'y_title': cfg['unit_pct'],
            'unit': cfg['unit_pct'],
            'series': series,
            'categories': categories,
        }
    else:
        categories = choose_parts_categories(cfg, str(row['dificultat']))
        vals = proportions(len(categories), row['patro_estadistic'], rnd)
        records = [{'categoria': cat, 'serie': 'Pes', 'valor': v} for cat, v in zip(categories, vals)]
        return records, {
            'x_title': 'Categoria',
            'y_title': cfg['unit_pct'],
            'unit': cfg['unit_pct'],
            'series': ['Pes'],
            'categories': categories,
        }


def generate_geo(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    categories = cfg['geo_categories'][:max(5, difficulty_count(row['dificultat'], 'geo', rnd))]
    if row['familia_excel'] == 'Map' or row['estructura_dades'] == 'geogràfica':
        pattern = row['patro_estadistic']
        lo, hi = choose_value_range(cfg, row['dificultat'], 'map_value_ranges', rnd)
        if pattern == 'gradient nord-sud':
            vals = list(reversed(clamp_series_values(line_pattern(len(categories), 'creixement', lo, hi - lo, rnd), lo, hi)))
        elif pattern == 'hotspot regional':
            vals = [round(rnd.uniform(lo, lo + 0.55 * (hi - lo)), 1) for _ in categories]
            vals[rnd.randrange(len(vals))] = round(rnd.uniform(lo + 0.78 * (hi - lo), hi), 1)
        else:
            vals = [round(rnd.uniform(lo, hi), 1) for _ in categories]
        records = [{'regio': cat, 'valor': v} for cat, v in zip(categories, vals)]
        return records, {
            'x_title': 'Regió',
            'y_title': cfg['map_metric'],
            'unit': cfg['map_metric'],
            'series': ['Valor'],
            'categories': categories,
        }
    return generate_time_long(row, cfg, rnd)


def generate_scatter(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = difficulty_count(row['dificultat'], 'scatter', rnd)
    x_name, x_min, x_max = cfg['scatter_x']
    y_name, y_min, y_max = cfg['scatter_y']
    pattern = row['patro_estadistic']
    groups = ['Grup 1', 'Grup 2'] if pattern == 'clústers' else ['Mostra']
    records = []
    for i in range(n):
        if pattern == 'correlació positiva':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_min + ratio * (y_max - y_min) + rnd.uniform(-0.08, 0.08) * (y_max - y_min)
        elif pattern == 'correlació negativa':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_max - ratio * (y_max - y_min) + rnd.uniform(-0.08, 0.08) * (y_max - y_min)
        elif pattern == 'correlació nul·la':
            x = rnd.uniform(x_min, x_max)
            y = rnd.uniform(y_min, y_max)
        elif pattern == 'relació corba':
            x = rnd.uniform(x_min, x_max)
            ratio = (x - x_min) / (x_max - x_min)
            y = y_min + (ratio ** 2) * (y_max - y_min) + rnd.uniform(-0.04, 0.04) * (y_max - y_min)
        elif pattern == 'clústers':
            grp = groups[i % 2]
            if grp == 'Grup 1':
                x = rnd.uniform(x_min, x_min + 0.35 * (x_max - x_min))
                y = rnd.uniform(y_min, y_min + 0.35 * (y_max - y_min))
            else:
                x = rnd.uniform(x_min + 0.55 * (x_max - x_min), x_max)
                y = rnd.uniform(y_min + 0.55 * (y_max - y_min), y_max)
            records.append({'x': round(x, 2), 'y': round(y, 2), 'grup': grp})
            continue
        elif pattern == 'outlier clar':
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
    cats = cfg['radar_axes'][:radar_axis_count(row['dificultat'], len(cfg['radar_axes']))]
    series = cfg['series'][:3]
    records = []
    for s_idx, ser in enumerate(series):
        vals = radar_series_profile(cats, row['patro_estadistic'], s_idx, len(series), rnd)
        for cat, v in zip(cats, vals):
            records.append({'categoria': cat, 'serie': ser, 'valor': round(v, 1)})
    return records, {
        'x_title': 'Indicador',
        'y_title': 'Puntuació (0-100)',
        'unit': 'Puntuació (0-100)',
        'series': series,
        'categories': cats,
    }


def generate_combo(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = difficulty_count(row['dificultat'], 'combo', rnd)
    cats = MONTHS[:n] if n <= 6 else YEARS[-n:]
    pattern = row['patro_estadistic']
    lo, hi, base, spread = base_spread_for_value_range(choose_value_range(cfg, row['dificultat'], 'category_value_ranges', rnd), rnd)
    primary = line_pattern(
        len(cats),
        'creixement' if pattern in {'dues escales diferents', 'sèrie principal + objectiu', 'creixement amb taxa'} else pattern,
        base,
        spread,
        rnd,
    )
    primary = clamp_series_values(primary, lo, hi, allow_negative=pattern == 'valors negatius')
    if pattern == 'dues escales diferents':
        secondary = [round(10 + 8 * math.sin(i), 1) for i in range(len(cats))]
    elif pattern == 'sèrie principal + objectiu':
        secondary = [round(sum(primary) / len(primary) * 0.95, 1) for _ in cats]
    else:
        secondary = line_pattern(len(cats), 'baixa volatilitat', 25, 12, rnd)
    records = []
    for cat, p, s in zip(cats, primary, secondary):
        records.append({'categoria': cat, 'serie': 'Primària', 'valor': float(p), 'eix': 'primari'})
        records.append({'categoria': cat, 'serie': 'Secundària', 'valor': float(s), 'eix': 'secundari'})
    return records, {
        'x_title': 'Període',
        'y_title': cfg['combo_primary'],
        'unit': cfg['combo_primary'],
        'secondary_title': cfg['combo_secondary'],
        'series': ['Primària', 'Secundària'],
        'categories': cats,
    }


def generate_stock(row: Dict[str, Any], cfg: Dict[str, Any], rnd: random.Random) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    n = max(10, min(20, difficulty_count(row['dificultat'], 'stock', rnd) + 8))
    dates = [f'2024-{m:02d}-{d:02d}' for m, d in zip(([1] * n), range(1, n + 1))]
    pattern = row['patro_estadistic']
    if pattern in {'creixement', 'baixa volatilitat'}:
        closes = line_pattern(n, pattern, 70, 18, rnd)
    else:
        closes = line_pattern(n, 'baixa volatilitat', 70, 18, rnd)
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
        'x_title': 'Data',
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
            if row['patro_estadistic'] == 'superfície rugosa':
                z += rnd.uniform(-7, 7)
            records.append({'x': x, 'y': y, 'z': round(z, 1)})
    return records, {
        'x_title': 'Eix X',
        'y_title': 'Eix Y',
        'unit': 'Intensitat',
        'series': None,
        'categories': None,
    }


def title_for(row: Dict[str, Any], cfg: Dict[str, Any], meta: Dict[str, Any]) -> str:
    domain = row['domini_semantic']
    structure = row['estructura_dades']
    family = row['familia_excel']
    case_label = str(row['case_id']).removeprefix('CASE_')
    if family == 'Scatter':
        return f"Relació entre indicadors de {domain} ({case_label})"
    if family == 'Stock':
        return f"Evolució diària de {cfg['stock_label'].lower()} ({case_label})"
    if family == 'Map':
        return f"Distribució territorial de {domain} ({case_label})"
    if family == 'Surface':
        return f"Superfície de resposta energètica ({case_label})"
    if structure == 'parts del tot multianell':
        return f"Composició comparada per categoria en {domain} ({case_label})"
    if 'parts del tot' in structure:
        return f"Pes relatiu per categoria en {domain} ({case_label})"
    if 'temporal' in structure:
        return f"Evolució de {domain} per període ({case_label})"
    if 'multivariable' in structure:
        return f"Perfil comparat d'indicadors de {domain} ({case_label})"
    if structure == 'geogràfica':
        return f"Indicador territorial de {domain} ({case_label})"
    return f"Indicadors de {domain} per categoria ({case_label})"


def make_canonical_instance(row: Dict[str, Any]) -> Dict[str, Any]:
    cfg = DOMAIN_CFG[str(row['domini_semantic']).strip().lower()]
    rnd = rng_for(row['case_id'])
    structure = row['estructura_dades']
    family = row['familia_excel']

    if family == 'Scatter' or structure == 'bivariant numèrica':
        records, meta = generate_scatter(row, cfg, rnd)
        data_format = 'scatter'
        encoding = {
            'x': {'field': 'x', 'type': 'quantitative', 'title': meta['x_title']},
            'y': {'field': 'y', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'grup', 'type': 'nominal', 'title': 'Grup'}
        }
    elif family == 'Stock' or structure == 'OHLC + volum':
        records, meta = generate_stock(row, cfg, rnd)
        data_format = 'stock'
        encoding = {
            'x': {'field': 'data', 'type': 'temporal', 'title': meta['x_title']},
            'y': {'field': 'tancament', 'type': 'quantitative', 'title': meta['y_title']},
        }
    elif family == 'Surface' or structure == 'graella numèrica 2D':
        records, meta = generate_surface(row, cfg, rnd)
        data_format = 'surface'
        encoding = {
            'x': {'field': 'x', 'type': 'quantitative', 'title': meta['x_title']},
            'y': {'field': 'y', 'type': 'quantitative', 'title': meta['y_title']},
            'z': {'field': 'z', 'type': 'quantitative', 'title': meta['unit']}
        }
    elif family == 'Radar' or structure == 'multivariable sobre eixos comuns':
        records, meta = generate_radar(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Sèrie'}
        }
    elif family == 'Combo' or structure == 'temporal amb eix secundari':
        records, meta = generate_combo(row, cfg, rnd)
        data_format = 'combo'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Sèrie'},
            'secondary_y': {'field': 'valor', 'type': 'quantitative', 'title': meta['secondary_title']}
        }
    elif family == 'Map' or structure == 'geogràfica':
        records, meta = generate_geo(row, cfg, rnd)
        data_format = 'map'
        encoding = {
            'x': {'field': 'regio', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
        }
    elif 'parts del tot' in structure or family in {'Pie', 'Doughnut'}:
        records, meta = generate_parts(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Sèrie'}
        }
    else:
        records, meta = generate_time_long(row, cfg, rnd)
        data_format = 'category_long'
        encoding = {
            'x': {'field': 'categoria', 'type': 'nominal', 'title': meta['x_title']},
            'y': {'field': 'valor', 'type': 'quantitative', 'title': meta['y_title']},
            'color': {'field': 'serie', 'type': 'nominal', 'title': 'Sèrie'}
        }

    title = title_for(row, cfg, meta)

    exact_excel_supported = family not in {'Map', 'Surface'} and row['subtipus_excel'] not in {'Pie of Pie', 'Bar of Pie', 'Filled Map', 'Contour'}
    render_mode = 'exacte' if exact_excel_supported else 'aproximat_o_metadades'

    provisional_inst = {
        'chart': {
            'stack_mode': 'percent' if '100% Stacked' in row['subtipus_excel'] else ('normal' if 'Stacked' in row['subtipus_excel'] else None),
        }
    }
    axes = build_axes_metadata(provisional_inst, records, meta, encoding)
    excel_metadata = build_excel_metadata(title, axes)
    series_count = count_series_in_records(records)

    return {
        '$schema': 'https://example.org/chart-canonical.schema.json',
        'version': '1.0-exploratoria',
        'id': row['case_id'],
        'title': title,
        'description': f"Instància canònica sintètica del gràfic {row['subtipus_excel']} en el domini {row['domini_semantic']}.",
        'style_ref': STYLE_REF,
        'source_case': row,
        'data': {
            'source': {'kind': 'inline', 'values': records},
            'structure': row['estructura_dades'],
            'statistical_pattern': row['patro_estadistic'],
            'semantic_domain': row['domini_semantic'],
            'unit': meta['unit'],
            'format': data_format,
        },
        'chart': {
            'family': row['familia_excel'].lower(),
            'excel_family': row['familia_excel'],
            'excel_subtype': row['subtipus_excel'],
            'excel_subtype_code': row['codi_subtipus'],
            'stack_mode': 'percent' if '100% Stacked' in row['subtipus_excel'] else ('normal' if 'Stacked' in row['subtipus_excel'] else None),
            'orientation': 'horizontal' if row['familia_excel'] == 'Bar' else 'vertical',
            'render_mode': render_mode,
            'exact_excel_supported': exact_excel_supported,
        },
        'excel_metadata': excel_metadata,
        'axes': axes,
        'encoding': encoding,
        'labels': {
            'title_visible': True,
            'legend_visible': series_count > 1,
            'data_labels': 'etiquetes de dades' in str(row.get('variant_estil', '')).lower(),
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
        ('Família Excel', inst['chart']['excel_family']),
        ('Subtipus Excel', inst['chart']['excel_subtype']),
        ('Estructura', inst['source_case']['estructura_dades']),
        ('Patró', inst['source_case']['patro_estadistic']),
        ('Dificultat', inst['source_case']['dificultat']),
        ('Domini', inst['source_case']['domini_semantic']),
        ('Style ref', inst['style_ref']['href']),
        ('Render Excel', inst['chart']['render_mode']),
    ]

    axes = inst.get('axes', {})
    axis_labels = [('x', 'X'), ('y', 'Y'), ('secondary_y', 'Y2'), ('z', 'Z')]
    for axis_key, axis_label in axis_labels:
        axis = axes.get(axis_key)
        if not axis:
            continue
        meta_rows.extend([
            (f'Eix {axis_label} - Mínim', axis.get('min')),
            (f'Eix {axis_label} - Màxim', axis.get('max')),
            (f'Eix {axis_label} - Interval', axis.get('interval')),
            (f'Eix {axis_label} - Unitat', axis.get('unit')),
        ])

    r = 2
    for k, v in meta_rows:
        ws.write(r, 0, k, formats['label'])
        ws.write(r, 1, '' if v is None else v, formats['text'])
        r += 1
    return r


def add_summary_sheet(workbook, instances, sheet_map, render_stats):
    ws = workbook.add_worksheet('Resum')
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9E2F3', 'border': 1})
    fmt_link = workbook.add_format({'font_color': 'blue', 'underline': 1})
    headers = ['case_id', 'full', 'familia_excel', 'subtipus_excel', 'estructura_dades', 'patro_estadistic', 'dificultat', 'domini_semantic', 'render_excel']
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt_header)
    for r, inst in enumerate(instances, start=1):
        src = inst['source_case']
        ws.write_url(r, 0, f"internal:'{sheet_map[inst['id']]}'!A1", fmt_link, string=inst['id'])
        ws.write(r, 1, inst['title'])
        ws.write(r, 2, src['familia_excel'])
        ws.write(r, 3, src['subtipus_excel'])
        ws.write(r, 4, src['estructura_dades'])
        ws.write(r, 5, src['patro_estadistic'])
        ws.write(r, 6, src['dificultat'])
        ws.write(r, 7, src['domini_semantic'])
        ws.write(r, 8, inst['chart']['render_mode'])
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(instances), len(headers) - 1)
    ws.set_column('A:A', 14)
    ws.set_column('B:B', 55)
    ws.set_column('C:I', 22)
    start = len(instances) + 3
    ws.write(start, 0, 'Estadístiques de render', fmt_header)
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
        return False, 'sense suport exacte'
    exact_subtypes_unsupported = {'Pie of Pie', 'Bar of Pie', 'Filled Map', 'Contour'}
    if inst['chart']['excel_subtype'] in exact_subtypes_unsupported:
        if inst['chart']['excel_family'] == 'Pie':
            chart_args = {'type': 'pie'}
        else:
            return False, 'subtipus no suportat exactament'
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
    return True, 'render exacte o proper'


def render_scatter(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    headers = ['x', 'y', 'grup']
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
    return True, 'render exacte'


def render_combo(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    categories = list(OrderedDict((r['categoria'], None) for r in records).keys())
    ws.write_row(start_row, 0, ['Període', 'Primària', 'Secundària'], formats['header'])
    prim = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Primària'}
    sec = {r['categoria']: r['valor'] for r in records if r['serie'] == 'Secundària'}
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
    return True, 'render exacte (combo)'


def render_stock(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    headers = ['Data', 'Obertura', 'Màxim', 'Mínim', 'Tancament', 'Volum']
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
        'name': 'Tancament',
        'categories': [sheet_name, start_row + 1, 0, start_row + len(records), 0],
        'values': [sheet_name, start_row + 1, 4, start_row + len(records), 4],
        **series_style('line', 0),
    })
    vol_chart = workbook.add_chart({'type': 'column'})
    vol_chart.add_series({
        'name': 'Volum',
        'categories': [sheet_name, start_row + 1, 0, start_row + len(records), 0],
        'values': [sheet_name, start_row + 1, 5, start_row + len(records), 5],
        'y2_axis': True,
        **series_style('column', 1),
    })
    line_chart.combine(vol_chart)
    line_chart.set_title({'name': inst['title']})
    line_chart.set_x_axis(excel_axis_options('Data'))
    line_chart.set_y_axis(excel_axis_options(inst['encoding']['y']['title']))
    line_chart.set_y2_axis(excel_axis_options('Volum'))
    apply_legend_for_series_count(line_chart, inst['xlsxwriter']['legend_position'], 2)
    line_chart.set_style(inst['xlsxwriter']['style'])
    line_chart.set_size(inst['xlsxwriter']['size'])
    ws.insert_chart(f'H{chart_row + 1}', line_chart)
    return True, 'render aproximat (línea + volum)'


def render_map(workbook, ws, formats, inst, sheet_name, start_row, chart_row):
    records = inst['data']['source']['values']
    ws.write_row(start_row, 0, ['Regió', 'Valor'], formats['header'])
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
    chart.set_title({'name': inst['title'] + ' - vista prèvia en barres'})
    chart.set_x_axis(excel_axis_options(inst['encoding']['y']['title']))
    chart.set_y_axis(excel_axis_options('Regió'))
    chart.set_style(inst['xlsxwriter']['style'])
    chart.set_size(inst['xlsxwriter']['size'])
    ws.write(start_row - 1, 0, 'Nota: Excel/XlsxWriter no suporta Filled Map; es mostra una vista prèvia en barres.', formats['note'])
    ws.insert_chart(f'G{chart_row + 1}', chart)
    return True, 'render aproximat (barres)'


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
    ws.write(start_row - 1, 0, 'Nota: Excel/XlsxWriter no suporta gràfics Surface/Contour; es mostra la graella amb escala de color.', formats['note'])
    return True, 'render aproximat (heatmap de cel·les)'


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

    with OUT_JSON.open('w', encoding='utf-8') as f:
        json.dump(instances, f, ensure_ascii=False, indent=2)

    workbook = xlsxwriter.Workbook(str(OUT_XLSX))
    used = set()
    sheet_map = {}
    render_stats = {'exacte': 0, 'aproximat_o_metadades': 0}
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
        'domains': sorted(df['domini_semantic'].dropna().unique().tolist()),
        'families': df['familia_excel'].value_counts().to_dict(),
    }
    with OUT_MANIFEST.open('w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print('Generat:')
    print(OUT_JSON)
    print(OUT_XLSX)
    print(OUT_MANIFEST)
    print(render_stats)


if __name__ == '__main__':
    main()
