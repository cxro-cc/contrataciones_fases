#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
 GENERADOR DE REPORTE DE CANDIDATOS POR FASE — UAS / UPE
====================================================================
Cruza los archivos de candidatos (UAS + UPE) con el catálogo oficial
de fases y produce un reporte HTML interactivo, navegable por fase.

USO:
    python generar_reporte.py \
        --uas      Candidatos_UAS.xlsx \
        --upe      Candidatos_UPE.xlsx \
        --catalogo Unidades_Inauguradas_fases.xlsx \
        --especial Fase_Especial.xlsx \
        --plantilla plantilla_reporte.html \
        --salida   index.html 

REQUISITOS:  pip install pandas openpyxl
====================================================================
"""

import argparse
import json
import re
import sys
import pandas as pd
from datetime import datetime


# --------------------------------------------------------------------
# CONFIGURACIÓN  (editable)
# --------------------------------------------------------------------

PHASE1_CLUES = [
    "BCIMB000623", "BCIMB001796", "BSIMB000812", "CCIMB001555",
    "DFIMB002353", "GRIMB010536", "MCIMB012476", "MNIMB004751", "OCIMB009626",
    "OCIMB009631", "OCIMB009643", "OCIMB009655", "OCIMB009660", "SLIMB002230",
    "SPIMB000585", "SRIMB000016", "VZIMB008946", "VZIMB008963", "VZIMB008980",
    "ZSIMB002650",
]
PHASE1_NAMES = {
    'BCIMB000623': 'Hospital Comunitario San Felipe',
    'BCIMB001796': 'Hospital General Tijuana Zona Este',
    'BSIMB000812': 'Hospital General de Santa Rosalía',
    'CCIMB001555': 'Centro de Salud El Naranjo',
    'DFIMB002353': 'Hospital Oncológico para la Mujer de la CDMX',
    'GRIMB010536': 'Hospital IMSS-Bienestar Tlapa',
    'MCIMB012476': 'Hospital General Atenco "Francisco Altamirano Núñez"',
    'MNIMB004751': 'Hospital Comunitario Maruata',
    'OCIMB009626': 'Centro de Salud de Servicios Ampliados (CESSA) Santiago Astata',
    'OCIMB009631': 'Hospital General de la Mujer y la Niñez Oaxaqueña',
    'OCIMB009643': 'Ixtlán de Juárez',
    'OCIMB009655': 'Centro de Salud Urbano 05 NB San Pablo Villa de Mitla',
    'OCIMB009660': 'Hospital General de San Juan Bautista Tuxtepec',
    'SLIMB002230': 'Hospital Pediátrico de Sinaloa',
    'SPIMB000585': 'Hospital General de Ríoverde',
    'SRIMB000016': 'Hospital Comunitario Vícam Switch',
    'VZIMB008946': 'Hospital Materno Infantil de Coatzacoalcos',
    'VZIMB008963': 'Hospital de Salud Mental Orizaba "Dr. Víctor M. Concha Vásquez"',
    'VZIMB008980': 'Hospital de la Comunidad de Nautla',
    'ZSIMB002650': 'Centro de Salud Jerez de García Salinas',
}

## FASE ESPECIAL
FASE_ESPECIAL = [  "MCIMB012295", "QRIMB001956", "DFIMB001822", "SPIMB002574",
  "DFIMB002341", "CSIMB001732", "DFIMB002435", "DFIMB002674", "TSIMB002865"]

## HOSPITALES ANCLA
CLUES_ANCLA = ['BCIMB000355',  'MNIMB003993',  'CSIMB003680',  'CSIMB001732',
    'CSIMB003622', 'CSIMB005500', 'CSIMB005582', 'GRIMB004096', 'GRIMB005163',
    'GRIMB008926', 'GRIMB008931', 'GRIMB010536', 'HGIMB001481', 'HGIMB004812',
    'HGIMB005005', 'MCIMB005225', 'MCIMB009063', 'MCIMB009174', 'MCIMB009425', 
    'MCIMB012073', 'MCIMB012476', 'MNIMB001292', 'MNIMB003940', 'MNIMB004734',
    'MSIMB000292', 'NTIMB000551', 'NTIMB001654', 'OCIMB008815', 'OCIMB000881',
    'OCIMB003133', 'OCIMB007625', 'OCIMB009660', 'PLIMB002516', 'PLIMB007083',
    'PLIMB005910', 'PLIMB006016', 'PLIMB006103', 'PLIMB006885', 'PLIMB007112',
    'QRIMB001606', 'QRIMB001623', 'QRIMB001973', 'SLIMB000014', 'SLIMB000031',
    'SLIMB002254', 'SLIMB002481', 'SLIMB002930', 'SPIMB000240', 'SPIMB002970',
    'SRIMB001404', 'SRIMB002203', 'TCIMB000786', 'TCIMB001841', 'TCIMB003446',
    'TSIMB001955', 'TSIMB002520', 'VZIMB000826', 'VZIMB000983', 'VZIMB003491',
    'VZIMB003911', 'VZIMB006315', 'VZIMB007575', 'VZIMB008053', 'ZSIMB000113',
    'ZSIMB000492', 'ZSIMB002481', 'BSIMB000754', 'OCIMB009071', 'TSIMB001260',
    'CCIMB001531', 'GRIMB001436', 'CSIMB000460']

# Prefijo de CLUES (2 letras) -> entidad federativa
CLUES_STATE = {
    'AS': 'Aguascalientes', 'BC': 'Baja California', 'BS': 'Baja California Sur',
    'CC': 'Campeche', 'CL': 'Coahuila', 'CO': 'Colima', 'CS': 'Chiapas',
    'CH': 'Chihuahua', 'DF': 'Ciudad de México', 'DG': 'Durango',
    'GT': 'Guanajuato', 'GR': 'Guerrero', 'HG': 'Hidalgo', 'JC': 'Jalisco',
    'MC': 'Estado de México', 'MN': 'Michoacán', 'MS': 'Morelos',
    'NT': 'Nayarit', 'NL': 'Nuevo León', 'OC': 'Oaxaca', 'PL': 'Puebla',
    'QT': 'Querétaro', 'QR': 'Quintana Roo', 'SL': 'Sinaloa',
    'SP': 'San Luis Potosí', 'SR': 'Sonora', 'TC': 'Tabasco',
    'TS': 'Tamaulipas', 'TL': 'Tlaxcala', 'VZ': 'Veracruz', 'YN': 'Yucatán',
    'ZS': 'Zacatecas',
}

UAS_CAT_ORDER = ['Enfermería', 'Médicos especialistas', 'Médicos generales',
                 'Otros profesionales de salud', 'Administrativos y apoyo', 'Otros']
UPE_CAT_ORDER = ['Médicos especialistas', 'Enfermería', 'Médicos generales',
                 'Otros profesionales de salud', 'Administrativos y apoyo', 'Otros']

PLACEHOLDER = '__PAYLOAD__'


# --------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------

def norm_fase(x):
    if pd.notna(x) and str(x).strip().replace('.', '').isdigit():
        return str(int(float(x)))
    return None


def clean_clues(c):
    m = re.match(r'([A-Z]{2}IMB\d+)', str(c))
    return m.group(1) if m else str(c).strip()


def title_case_es(s):
    s = str(s).strip()
    if not s or s.lower() == 'nan':
        return ''
    s = s.lower()
    smalls = {'de', 'del', 'la', 'las', 'los', 'y', 'el', 'en', 'con',
              'para', 'por', 'a', 'al', 'o', 'u', 'e'}
    return ' '.join(p if (p in smalls and i != 0) else p.capitalize()
                    for i, p in enumerate(s.split()))


def short_name(full):
    if not full:
        return ''
    s = re.sub(r'^Hospital\s+(de|del)?\s*', '', full, flags=re.I)
    s = re.sub(r'^Centro\s+de\s+Salud\s+(de|del)?\s*', '', s, flags=re.I)
    s = re.sub(r'"[^"]+"', '', s).strip()
    s = re.sub(r'\(.*?\)', '', s).strip()
    s = re.sub(r'\s+', ' ', s)
    return s[:30] + '…' if len(s) > 32 else s


def state_from_clues(c):
    return CLUES_STATE.get(str(c)[:2], '')


def categorize(row):
    cnpm = str(row.get('CNPM', '')).strip()
    puesto = str(row.get('clave_puesto', '')).upper()
    if cnpm.startswith('ME'):
        return 'Médicos especialistas'
    if cnpm.startswith('MG'):
        return 'Médicos generales'
    if cnpm.startswith('EN'):
        return 'Enfermería'
    if 'ENFERMER' in puesto or ('AUX' in puesto and 'ENFERM' in puesto):
        return 'Enfermería'
    if any(k in puesto for k in ['PSICOLOG', 'FISIOTERAP', 'QUIMIC', 'NUTRIOLOG',
                                 'LABORATORIST', 'CITOTECN', 'TECNICO RADIOLOG',
                                 'TRABAJADOR', 'INGENIER', 'FARMAC', 'HISTOPATOLOG']):
        return 'Otros profesionales de salud'
    if any(k in puesto for k in ['AUXILIAR ADMIN', 'COCIN', 'CAMILL', 'OPERADOR',
                                 'JEFE', 'JEFA', 'SUBDIRECTOR', 'COORDINADOR', 'AUXLIAR']):
        return 'Administrativos y apoyo'
    if any(k in puesto for k in ['MEDIC', 'CIRUG', 'IMAGEN']):
        return 'Médicos especialistas'
    return 'Otros'


def uas_estado(v):
    if pd.isna(v) or str(v).strip() == '':
        return 'Pendiente'
    s = str(v).strip().upper()
    return {'APROBADO': 'Aprobado', 'NO APROBADO': 'No aprobado'}.get(s, s.title())


def upe_estado(v):
    if pd.isna(v) or str(v).strip() == '':
        return 'Pendiente'
    s = str(v).strip()
    low = s.lower()
    if low.startswith('acept'):
        return 'Aceptado'
    if low.startswith('rechaz'):
        return 'Rechazado'
    if low.startswith('revis'):
        return 'Revisión'
    return s


def unidad_lookup(clues, *dfs):
    for df in dfs:
        s = df[df['clues'] == clues]['unidad_medica']
        if len(s) > 0 and pd.notna(s.iloc[0]):
            return str(s.iloc[0])
    return ''


def _clean_value(v):
    if pd.isna(v):
        return v
    s = str(v).strip()
    return s if s else None


def completar_cnpm_y_puesto(uas, upe):
    for df in (uas, upe):
        df['CNPM'] = df['CNPM'].apply(_clean_value)
        df['clave_puesto'] = df['clave_puesto'].apply(_clean_value)

    cnpm_prefix = re.compile(r'^([A-Z]{2,6}\d{3,5})[\s\.]')

    def extract_and_strip(row):
        kp = row['clave_puesto']
        cn = row['CNPM']
        if pd.isna(kp):
            return cn, kp
        m = cnpm_prefix.match(str(kp))
        if m:
            extracted = m.group(1)
            kp_clean = str(kp)[m.end():].strip()
            if pd.isna(cn) or not cn:
                cn = extracted
            return cn, kp_clean
        return cn, kp

    recovered_uas = 0
    recovered_upe = 0
    for df, _ in ((uas, 'uas'), (upe, 'upe')):
        before = df['CNPM'].isna().sum()
        results = df.apply(extract_and_strip, axis=1, result_type='expand')
        df['CNPM'] = results[0]
        df['clave_puesto'] = results[1]
        after = df['CNPM'].isna().sum()
        if df is uas:
            recovered_uas = before - after
        else:
            recovered_upe = before - after
    print(f"  Extracción de CNPM desde prefijo de clave_puesto: UAS +{recovered_uas} · UPE +{recovered_upe}")

    both = pd.concat([uas[['clave_puesto', 'CNPM']],
                      upe[['clave_puesto', 'CNPM']]], ignore_index=True)
    known = both.dropna(subset=['clave_puesto', 'CNPM'])
    puesto_to_cnpm = (known.groupby('clave_puesto')['CNPM']
                      .agg(lambda s: s.value_counts().idxmax()).to_dict())
    cnpm_to_puesto = (known.groupby('CNPM')['clave_puesto']
                      .agg(lambda s: s.value_counts().idxmax()).to_dict())

    def fill(df, label):
        before_c = df['CNPM'].isna().sum()
        before_p = df['clave_puesto'].isna().sum()
        mask_c = df['CNPM'].isna() & df['clave_puesto'].notna()
        df.loc[mask_c, 'CNPM'] = df.loc[mask_c, 'clave_puesto'].map(puesto_to_cnpm)
        mask_p = df['clave_puesto'].isna() & df['CNPM'].notna()
        df.loc[mask_p, 'clave_puesto'] = df.loc[mask_p, 'CNPM'].map(cnpm_to_puesto)
        after_c = df['CNPM'].isna().sum()
        after_p = df['clave_puesto'].isna().sum()
        print(f"  {label} CNPM         : {before_c} -> {after_c} (resueltos {before_c - after_c})")
        print(f"  {label} clave_puesto : {before_p} -> {after_p} (resueltos {before_p - after_p})")

    fill(uas, 'UAS')
    fill(upe, 'UPE')
    return uas, upe


# --------------------------------------------------------------------
# CARGA Y RESOLUCIÓN DE FASE
# --------------------------------------------------------------------

def cargar_y_resolver(uas_path, upe_path, catalogo_path):
    uas = pd.read_excel(uas_path)
    upe = pd.read_excel(upe_path)

    # Columnas mínimas
    for col in ['clues', 'fase', 'clave_puesto', 'CNPM', 'unidad_medica', 'turno']:
        if col not in uas.columns:
            uas[col] = None
        if col not in upe.columns:
            upe[col] = None
    if 'REVISION_UAS' not in uas.columns:
        uas['REVISION_UAS'] = None
    if 'UPE' not in upe.columns:
        upe['UPE'] = None

    uas['fase_norm'] = uas['fase'].apply(norm_fase)
    upe['fase_norm'] = upe['fase'].apply(norm_fase)

    cat_earliest = {}
    if catalogo_path:
        try:
            cat = pd.read_excel(catalogo_path, sheet_name='brecha')
        except Exception:
            xl = pd.ExcelFile(catalogo_path)
            cat = None
            for sh in xl.sheet_names:
                tmp = pd.read_excel(catalogo_path, sheet_name=sh)
                if {'fase', 'clues_imb'}.issubset(set(tmp.columns)):
                    cat = tmp
                    break
            if cat is None:
                print("AVISO: no se encontró hoja con columnas 'fase'/'clues_imb' en el catálogo.")
        if cat is not None:
            cat = cat[cat['clues_imb'].notna()].copy()
            cat['fase_num'] = cat['fase'].astype(str).str.extract(r'(\d+)')[0]
            grouped = cat.groupby('clues_imb')['fase_num'].agg(
                lambda x: sorted(set(x.dropna()), key=int))
            cat_earliest = {c: fs[0] for c, fs in grouped.items() if fs}

    combined = {}
    for df in (uas, upe):
        for clues, grp in df[df['fase_norm'].notna()].groupby('clues'):
            for f in grp['fase_norm'].unique():
                combined.setdefault(clues, set()).add(f)
    sibling_map = {c: list(fs)[0] for c, fs in combined.items() if len(fs) == 1}

    def resolve(row):
        if pd.notna(row['fase_norm']) and row['fase_norm'] is not None:
            return row['fase_norm']
        c = row['clues']
        if c in sibling_map:
            return sibling_map[c]
        cc = clean_clues(c)
        if cc in cat_earliest:
            return cat_earliest[cc]
        return None

    before = uas['fase_norm'].isna().sum() + upe['fase_norm'].isna().sum()
    uas['fase_norm'] = uas.apply(resolve, axis=1)
    upe['fase_norm'] = upe.apply(resolve, axis=1)
    after = uas['fase_norm'].isna().sum() + upe['fase_norm'].isna().sum()
    print(f"  Registros sin fase: {before} -> {after} (resueltos {before - after})")

    print("Completando CNPM y clave_puesto…")
    uas, upe = completar_cnpm_y_puesto(uas, upe)

    uas['categoria'] = uas.apply(categorize, axis=1)
    upe['categoria'] = upe.apply(categorize, axis=1)
    uas['estado_rev'] = uas['REVISION_UAS'].apply(uas_estado)
    upe['estado_rev'] = upe['UPE'].apply(upe_estado)
    return uas, upe


# --------------------------------------------------------------------
# CÓMPUTO DE VISTAS
# --------------------------------------------------------------------

def compute_view(uas_df, upe_df, expected_clues=None, expected_names=None):
    view = {'totales': {
        'uas': int(len(uas_df)),
        'upe': int(len(upe_df)),
        'aprobados': int((upe_df['estado_rev'] == 'Aceptado').sum()),
    }}

    def cat_table(df, order, statuses):
        cc = df['categoria'].value_counts().to_dict()
        cats = [c for c in order if c in cc] + [c for c in cc if c not in order]
        rows = []
        for cat in cats:
            sub = df[df['categoria'] == cat]
            r = {'cat': cat, 'total': int(len(sub))}
            for st in statuses:
                r[st] = int((sub['estado_rev'] == st).sum())
            rows.append(r)
        return rows

    view['uas_categorias'] = cat_table(uas_df, UAS_CAT_ORDER, ['Aprobado', 'No aprobado', 'Pendiente'])
    view['upe_categorias'] = cat_table(upe_df, UPE_CAT_ORDER, ['Aceptado', 'Rechazado', 'Revisión', 'Pendiente'])
    view['uas_status_totals'] = {st: int((uas_df['estado_rev'] == st).sum())
                                 for st in ['Aprobado', 'No aprobado', 'Pendiente']}
    view['upe_status_totals'] = {st: int((upe_df['estado_rev'] == st).sum())
                                 for st in ['Aceptado', 'Rechazado', 'Revisión', 'Pendiente']}

    all_clues = set(uas_df['clues'].dropna().unique()) | set(upe_df['clues'].dropna().unique())
    cd = {}
    for c in all_clues:
        su = uas_df[uas_df['clues'] == c]
        sp = upe_df[upe_df['clues'] == c]
        nombre = (expected_names.get(c) if (expected_names and c in expected_names)
                  else unidad_lookup(c, uas_df, upe_df))
        e = {
            'nombre': title_case_es(nombre),
            'short': short_name(title_case_es(nombre)),
            'estado': state_from_clues(c),
            'uas': {
                'total': int(len(su)),
                'Aprobado': int((su['estado_rev'] == 'Aprobado').sum()),
                'No aprobado': int((su['estado_rev'] == 'No aprobado').sum()),
                'Pendiente': int((su['estado_rev'] == 'Pendiente').sum()),
                'cats': {k: int(v) for k, v in su['categoria'].value_counts().items()},
            },
            'upe': {
                'total': int(len(sp)),
                'Aceptado': int((sp['estado_rev'] == 'Aceptado').sum()),
                'Rechazado': int((sp['estado_rev'] == 'Rechazado').sum()),
                'Revisión': int((sp['estado_rev'] == 'Revisión').sum()),
                'Pendiente': int((sp['estado_rev'] == 'Pendiente').sum()),
                'cats': {k: int(v) for k, v in sp['categoria'].value_counts().items()},
            },
            'detail_uas': {}, 'detail_upe': {},
        }
        for cat, gc in su.groupby('categoria'):
            arr = [{'puesto': title_case_es(p), 'total': int(len(g)),
                    'status': {k: int(v) for k, v in g['estado_rev'].value_counts().items()}}
                   for p, g in gc.groupby('clave_puesto')]
            arr.sort(key=lambda x: -x['total'])
            e['detail_uas'][cat] = arr
        for cat, gc in sp.groupby('categoria'):
            arr = [{'puesto': title_case_es(p), 'total': int(len(g)),
                    'status': {k: int(v) for k, v in g['estado_rev'].value_counts().items()}}
                   for p, g in gc.groupby('clave_puesto')]
            arr.sort(key=lambda x: -x['total'])
            e['detail_upe'][cat] = arr
        cd[c] = e

    empty = []
    if expected_clues:
        for c in expected_clues:
            if c not in all_clues:
                empty.append({'code': c,
                              'nombre': title_case_es((expected_names or {}).get(c, '')),
                              'estado': state_from_clues(c)})
    view['empty_clues'] = empty
    view['clues_order'] = sorted(all_clues, key=lambda c: (
        -cd[c]['uas']['total'], -cd[c]['upe']['total'], c))
    view['clues_data'] = cd
    return view


# El parámetro fase_especial ahora se recibe dinámicamente aquí abajo
def construir_payload(uas, upe, corte, fase_especial):
    especial_set = set(fase_especial)
    esp_mask_uas = uas['clues'].apply(clean_clues).isin(especial_set)
    esp_mask_upe = upe['clues'].apply(clean_clues).isin(especial_set)
    uas_esp = uas[esp_mask_uas]
    upe_esp = upe[esp_mask_upe]
    uas_base = uas[~esp_mask_uas]
    upe_base = upe[~esp_mask_upe]

    phases = sorted([p for p in set(uas_base['fase_norm'].dropna().tolist() +
                                    upe_base['fase_norm'].dropna().tolist()) if p],
                    key=lambda x: int(x))

    views = {'general': compute_view(uas, upe)}
    ancla_set = set(CLUES_ANCLA)
    for p in phases:
        if p == '1':
            views['1'] = compute_view(uas_base[uas_base['fase_norm'] == '1'],
                                      upe_base[upe_base['fase_norm'] == '1'],
                                      PHASE1_CLUES, PHASE1_NAMES)
        elif p == '3':
            uas_f3 = uas_base[uas_base['fase_norm'] == '3']
            upe_f3 = upe_base[upe_base['fase_norm'] == '3']
            mask_uas3 = (uas_f3['clues'].apply(clean_clues).isin(ancla_set) &
                         uas_f3['turno'].fillna('').str.strip().str.lower().eq('equipo itinerante'))
            mask_upe3 = (upe_f3['clues'].apply(clean_clues).isin(ancla_set) &
                         upe_f3['turno'].fillna('').str.strip().str.lower().eq('equipo itinerante'))
            views['3'] = compute_view(uas_f3[mask_uas3], upe_f3[mask_upe3])
            views['3b'] = compute_view(uas_f3[~mask_uas3], upe_f3[~mask_upe3])
        else:
            views[p] = compute_view(uas_base[uas_base['fase_norm'] == p],
                                    upe_base[upe_base['fase_norm'] == p])

    views['especial'] = compute_view(uas_esp, upe_esp)

    uas_sf = uas_base[uas_base['fase_norm'].isna()]
    upe_sf = upe_base[upe_base['fase_norm'].isna()]
    has_sf = len(uas_sf) > 0 or len(upe_sf) > 0
    if has_sf:
        views['sinfase'] = compute_view(uas_sf, upe_sf)

    meta = [{'id': 'general', 'label': 'General', 'sub': 'todas las fases'}]
    for p in phases:
        v = views[p]
        meta.append({'id': p, 'label': f'Fase {p}',
                     'sub': f"{v['totales']['uas']} UAS · {v['totales']['upe']} UPE"})
        if p == '3' and '3b' in views:
            v3b = views['3b']
            meta.append({'id': '3b', 'label': 'Fase 3 brecha específica',
                         'sub': f"{v3b['totales']['uas']} UAS · {v3b['totales']['upe']} UPE"})
    v_esp = views['especial']
    meta.append({'id': 'especial', 'label': 'Especial',
                 'sub': f"{v_esp['totales']['uas']} UAS · {v_esp['totales']['upe']} UPE"})
    if has_sf:
        v = views['sinfase']
        meta.append({'id': 'sinfase', 'label': 'Sin fase',
                     'sub': f"{v['totales']['uas']} UAS · {v['totales']['upe']} UPE"})

    return {'phases': meta, 'views': views, 'corte': corte}


# --------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Genera el reporte HTML de candidatos por fase.")
    ap.add_argument('--uas', required=True, help="Excel de candidatos UAS")
    ap.add_argument('--upe', required=True, help="Excel de candidatos UPE")
    ap.add_argument('--catalogo', default=None, help="Excel catálogo de fases (hoja 'brecha').")
    
    # LA SOLUCIÓN INICIA AQUÍ: Agregamos el parámetro para el archivo de OneDrive
    ap.add_argument('--especial', required=True, help="Excel de la Fase Especial proveniente de OneDrive")
    
    ap.add_argument('--plantilla', required=True, help="Plantilla HTML con el marcador __PAYLOAD__")
    ap.add_argument('--salida', default='index.html', help="Archivo HTML de salida")
    args = ap.parse_args()

    print("Cargando y resolviendo fases…")
    uas, upe = cargar_y_resolver(args.uas, args.upe, args.catalogo)

    # LA SOLUCIÓN CONTINÚA AQUÍ: Leemos el archivo descargado desde OneDrive dentro de main()
    print("Cargando catálogo de la Fase Especial desde OneDrive…")
    df_especial = pd.read_excel(args.especial)
    fase_especial_list = df_especial['clues'].dropna().unique().tolist()

    print("Calculando vistas…")
    corte = datetime.now().strftime('%d / %m / %Y a las %H:%M')
    
    # Pasamos la lista leída a la función constructora del JSON
    payload = construir_payload(uas, upe, corte, fase_especial_list)

    print("Inyectando en la plantilla…")
    with open(args.plantilla, 'r', encoding='utf-8') as f:
        html = f.read()
    if PLACEHOLDER not in html:
        html = re.sub(r'const DATA = \{.*?\};\n',
                      'const DATA = ' + json.dumps(payload, ensure_ascii=False,
                                                   separators=(',', ':')) + ';\n',
                      html, count=1, flags=re.DOTALL)
    else:
        html = html.replace(PLACEHOLDER,
                            json.dumps(payload, ensure_ascii=False, separators=(',', ':')))

    with open(args.salida, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n✓ Reporte generado: {args.salida}")
    print("\nResumen por vista:")
    for p in payload['phases']:
        v = payload['views'][p['id']]
        print(f"  {p['label']:10s} — UAS {v['totales']['uas']:5d} · "
              f"UPE {v['totales']['upe']:5d} · Aprob {v['totales']['aprobados']:5d} · "
              f"{len(v['clues_order'])} CLUES")


if __name__ == '__main__':
    main()