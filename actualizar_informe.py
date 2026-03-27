"""
actualizar_informe.py
=====================
Lee los Excel de OneDrive y regenera el HTML del informe de cosecha.

USO:
  1. Doble click en este archivo  (o: python actualizar_informe.py)
  2. El script genera: informe_cosecha.html en la misma carpeta

REQUISITOS:
  pip install pandas openpyxl
"""

import sys
import os
import json
import re
import pandas as pd
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────
#  CONFIGURACIÓN — ajustá estas rutas si cambian
# ─────────────────────────────────────────────────────────────
COSECHA_EXCEL    = r"C:\Users\Joaquin\OneDrive - ESPARTINA S.A\DocumentacionEspartina\COMERCIAL\Tablas PBI\Avance Cosecha y Produccion 25-26.xlsx"
PARTICIPACION_EXCEL = r"C:\Users\Joaquin\OneDrive - ESPARTINA S.A\DocumentacionEspartina\COMERCIAL\Tablas PBI\Participacion Socios 25-26.xlsx"

# Archivo HTML de salida (se genera en la misma carpeta que este script)
OUTPUT_HTML = Path(__file__).parent / "informe_cosecha.html"

# Archivo template (se guarda junto a este script la primera vez)
TEMPLATE_FILE = Path(__file__).parent / "template.html"

# ─────────────────────────────────────────────────────────────
#  SOCIOS — contraseñas (no tocar)
# ─────────────────────────────────────────────────────────────
SOCIOS = [
    {'name': 'AGROCAFE SA',                    'pwd': 'Agrocafe2026',      'col': 23},
    {'name': 'AGROPECUARIA LOS GROBITOS SA',   'pwd': 'Grobitos2026',      'col': 42},
    {'name': 'Bioceres Semillas SAU',           'pwd': 'Bioceres2026',      'col': 61},
    {'name': 'CAPELLINO AGROPECUARIA SA',       'pwd': 'Capellino2026',     'col': 80},
    {'name': 'CENTRO AGROPECUARIO MODELO S.A.','pwd': 'Centro2026',         'col': 99},
    {'name': 'Cereales Quemu SA',               'pwd': 'Quemu2026',         'col': 118},
    {'name': 'Enrique M Baya Casal SA',         'pwd': 'BayaCasal2026',     'col': 156},
    {'name': 'ITIN SRL',                        'pwd': 'Itin2026',          'col': 175},
    {'name': 'Kerube SA',                       'pwd': 'Kerube2026',        'col': 194},
    {'name': 'Lartirigoyen y Cia. S.A.',        'pwd': 'Lartirigoyen2026',  'col': 213},
    {'name': 'NVD PARTICIPACIONES S.A',         'pwd': 'NVD2026',           'col': 232},
    {'name': 'Partes Iguales SH',               'pwd': 'PartesIguales2026', 'col': 251},
    {'name': 'Rizobacter Argentina SA',         'pwd': 'Rizobacter2026',    'col': 289},
    {'name': 'Siner',                           'pwd': 'Siner2026',         'col': 308},
    {'name': 'SNACK CROPS SA',                  'pwd': 'SnackCrops2026',    'col': 327},
    {'name': 'Sovoilar SA',                     'pwd': 'Sovoilar2026',      'col': 346},
]
DEPOSITO_COL = 137


# ─────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────
def safe_num(v, default=0):
    try:
        f = float(v)
        return default if (f != f) else f
    except:
        return default

def is_skip(val):
    return not val or str(val).strip() in ['nan', 'Totales', 'Campo', 'NaN', '']


# ─────────────────────────────────────────────────────────────
#  PASO 1 — Leer Excel de cosecha
# ─────────────────────────────────────────────────────────────
def leer_cosecha(path):
    print(f"  Leyendo cosecha: {Path(path).name}")
    df = pd.read_excel(path)

    campos = df[
        df['Campo'].notna() &
        (~df['Campo'].astype(str).isin(['Totales', 'nan', 'NaN'])) &
        df['Actividad'].notna() &
        (~df['Actividad'].astype(str).isin(['Totales', 'nan']))
    ].copy()

    records = []
    for _, r in campos.iterrows():
        avance = safe_num(r['% de Avance'])
        if avance <= 1.0:
            avance = round(avance * 100, 2)
        records.append({
            'act':        str(r['Actividad']).strip(),
            'zona':       str(r['Zona']).strip(),
            'loc':        str(r['Localidad']).strip() if pd.notna(r['Localidad']) else '',
            'campo':      str(r['Campo']).strip(),
            'ppto':       round(safe_num(r['Rinde Presupuestado (Prm Pond)']), 1),
            'actual':     round(safe_num(r['Rinde Actual']), 1),
            'final':      round(safe_num(r['Rinde Final']), 1),
            'sembradas':  round(safe_num(r['Has Sembradas'])),
            'cosechadas': round(safe_num(r['Has. Cosechadas'])),
            'perdidas':   round(safe_num(r['Has Perdidas'])),
            'avance':     round(avance, 2),
        })

    tot_s = sum(r['sembradas'] for r in records)
    tot_c = sum(r['cosechadas'] for r in records)
    avance_global = round(tot_c / tot_s * 100, 1) if tot_s > 0 else 0
    print(f"  → {len(records)} campos | {tot_s:,} ha sembradas | avance {avance_global}%")
    return records, avance_global


# ─────────────────────────────────────────────────────────────
#  PASO 2 — Leer Excel de participación
# ─────────────────────────────────────────────────────────────
def leer_participacion(path):
    print(f"  Leyendo participación: {Path(path).name}")
    df = pd.read_excel(path, header=None)

    # Construir mapa de depósito/bolsa
    deposito = {}
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        act   = str(row[0]).strip() if pd.notna(row[0]) else ''
        zona  = str(row[1]).strip() if pd.notna(row[1]) else ''
        loc   = str(row[2]).strip() if pd.notna(row[2]) else ''
        campo = str(row[3]).strip() if pd.notna(row[3]) else ''
        if not act or act in ['nan', 'Totales']: continue
        if is_skip(campo): continue
        bolsa = safe_num(row[DEPOSITO_COL + 15])
        if bolsa > 0:
            deposito[f"{campo}||{act}"] = {'tn_bolsa': bolsa, 'zona': zona, 'loc': loc}

    socios_data = {}
    retiro_data = {}

    for s in SOCIOS:
        c = s['col']
        campos_pct = {}
        retiros    = []

        for idx in range(2, len(df)):
            row = df.iloc[idx]
            act   = str(row[0]).strip() if pd.notna(row[0]) else ''
            zona  = str(row[1]).strip() if pd.notna(row[1]) else ''
            loc   = str(row[2]).strip() if pd.notna(row[2]) else ''
            campo = str(row[3]).strip() if pd.notna(row[3]) else ''
            if not act or act in ['nan', 'Totales']: continue
            if is_skip(campo): continue
            if str(row[c]).strip() in ['-', 'nan', '']: continue

            pct = safe_num(row[c])
            if pct == 0: continue

            key = f"{campo}||{act}"
            campos_pct[key] = round(pct, 6)

            tn_bolsa_socio = round(deposito.get(key, {}).get('tn_bolsa', 0) * pct, 2)
            retiros.append({
                'act': act, 'zona': zona, 'loc': loc, 'campo': campo,
                'pct':          round(pct, 4),
                'tn_teoricas':  round(safe_num(row[c + 1]), 2),
                'tn_retiradas': round(safe_num(row[c + 2]), 2),
                'tn_a_retirar': round(safe_num(row[c + 3]), 2),
                'tn_estimadas': round(safe_num(row[c + 7]), 2),
                'tn_bolsa':     tn_bolsa_socio,
            })

        socios_data[s['pwd']] = {'name': s['name'], 'campos': campos_pct}
        retiro_data[s['pwd']] = retiros

    print(f"  → {len(SOCIOS)} socios procesados | {len(deposito)} campos con bolsa")
    return socios_data, retiro_data


# ─────────────────────────────────────────────────────────────
#  PASO 3 — Construir bloques JS
# ─────────────────────────────────────────────────────────────
def build_data_js(records):
    lines = ['// ===== DATA =====\nconst DATA = [']
    act_groups = {}
    for r in records:
        act_groups.setdefault(r['act'], []).append(r)
    for act, recs in act_groups.items():
        lines.append(f'  // {act}')
        for r in recs:
            lines.append(
                f'  {{act:{json.dumps(r["act"])},zona:{json.dumps(r["zona"])},'
                f'loc:{json.dumps(r["loc"])},campo:{json.dumps(r["campo"])},'
                f'ppto:{r["ppto"]},actual:{r["actual"]},final:{r["final"]},'
                f'sembradas:{r["sembradas"]},cosechadas:{r["cosechadas"]},'
                f'perdidas:{r["perdidas"]},avance:{r["avance"]}}},'
            )
    lines.append('];')
    return '\n'.join(lines)

def build_socios_js(socios_data):
    lines = ['const SOCIOS_DATA = {']
    for pwd, info in socios_data.items():
        v = json.dumps(info, ensure_ascii=False, separators=(',', ':'))
        lines.append(f'  {json.dumps(pwd)}:{v},')
    lines.append('};')
    return '\n'.join(lines)

def build_retiro_js(retiro_data):
    lines = ['const RETIRO_DATA = {']
    for pwd, campos in retiro_data.items():
        v = json.dumps(campos, ensure_ascii=False, separators=(',', ':'))
        lines.append(f'  {json.dumps(pwd)}:{v},')
    lines.append('};')
    return '\n'.join(lines)


# ─────────────────────────────────────────────────────────────
#  PASO 4 — Inyectar en template y guardar HTML
# ─────────────────────────────────────────────────────────────
def generar_html(data_js, socios_js, retiro_js, avance_global):
    if not TEMPLATE_FILE.exists():
        print(f"\n  ERROR: No encontré el archivo template.html")
        print(f"  Debe estar en la misma carpeta que este script: {TEMPLATE_FILE}")
        return False

    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    # Inyectar bloques
    html = template
    html = html.replace('@@DATA_BLOCK@@',   data_js)
    html = html.replace('@@SOCIOS_BLOCK@@', socios_js)
    html = html.replace('@@RETIRO_BLOCK@@', retiro_js)

    # Actualizar fecha y avance en el header
    fecha_hoy = datetime.now().strftime('%d/%m/%Y')
    html = re.sub(
        r'Estado al \d+ de \w+ de \d+ · Todas las zonas',
        f'Estado al {fecha_hoy} · Todas las zonas',
        html
    )
    html = re.sub(
        r'⚡ Avance general \d+\.?\d*%',
        f'⚡ Avance general {avance_global}%',
        html
    )

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = OUTPUT_HTML.stat().st_size // 1024
    print(f"  → Guardado: {OUTPUT_HTML.name} ({size_kb} KB)")
    return True


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  INFORME DE COSECHA — Espartina SA")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 55)

    # Verificar que existen los archivos
    for path, nombre in [(COSECHA_EXCEL, 'Cosecha'), (PARTICIPACION_EXCEL, 'Participación')]:
        if not Path(path).exists():
            print(f"\n  ERROR: No encontré el Excel de {nombre}:")
            print(f"  {path}")
            print(f"\n  Verificá que el archivo existe y que OneDrive está sincronizado.")
            input("\n  Presioná Enter para cerrar...")
            return

    try:
        print("\n[1/4] Leyendo Excel de cosecha...")
        records, avance_global = leer_cosecha(COSECHA_EXCEL)

        print("\n[2/4] Leyendo Excel de participación...")
        socios_data, retiro_data = leer_participacion(PARTICIPACION_EXCEL)

        print("\n[3/4] Construyendo bloques de datos...")
        data_js   = build_data_js(records)
        socios_js = build_socios_js(socios_data)
        retiro_js = build_retiro_js(retiro_data)
        print(f"  → DATA: {len(records)} registros")
        print(f"  → Socios: {len(socios_data)}")

        print("\n[4/4] Generando HTML...")
        ok = generar_html(data_js, socios_js, retiro_js, avance_global)

        if ok:
            print("\n" + "=" * 55)
            print("  ✅ INFORME ACTUALIZADO CORRECTAMENTE")
            print(f"  Archivo: {OUTPUT_HTML}")
            print("=" * 55)

    except Exception as e:
        print(f"\n  ERROR inesperado: {e}")
        import traceback
        traceback.print_exc()

    input("\n  Presioná Enter para cerrar...")


if __name__ == '__main__':
    main()
