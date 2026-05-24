#!/usr/bin/env python3
"""
actualizar_dashboard.py
=======================
Genera dashboard_institucional.html leyendo:
  napsis/{año}/XX_Notas.xlsx        — notas NAPSIS por curso y año
  DIA/DIA {año}/DIA Diagnóstico {año}/{curso}/*.xls(x)  — evaluación diagnóstica
  SEPA/Reportes progreso alumnos */*.pdf                 — progreso SEPA individual

Uso:
  python actualizar_dashboard.py
"""

import sys, re, json, unicodedata, difflib, io
from pathlib import Path
from collections import defaultdict
from datetime import date

from bs4 import BeautifulSoup
import pandas as pd
import pdfplumber

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════
BASE_DIR     = Path(__file__).parent
NAPSIS_DIR   = BASE_DIR / "napsis"
ACTAS_DIR    = NAPSIS_DIR / "actas"        # TXT descargados desde sección Actas de NAPSIS
DIA_DIR      = BASE_DIR / "DIA"
SEPA_DIR     = BASE_DIR / "SEPA"
TEMPLATE_DIR = BASE_DIR / "template"
OUTPUT_FILE  = BASE_DIR / "dashboard_institucional.html"

CURRENT_YEAR = date.today().year
SCHOOL_NAME  = "Liceo Bicentenario María Luisa Bombal"

GRADE_NAMES = {
    1: "1° Básico",  2: "2° Básico",  3: "3° Básico",  4: "4° Básico",
    5: "5° Básico",  6: "6° Básico",  7: "7° Básico",  8: "8° Básico",
    9: "I° Medio",  10: "II° Medio", 11: "III° Medio", 12: "IV° Medio",
}

NP_KEYWORDS = ['orientaci', 'taller', 'biblio', 'paes']


# ══════════════════════════════════════════════════════════════
# UTILIDADES GENERALES
# ══════════════════════════════════════════════════════════════

def fix_enc(t):
    """Repara texto latin-1 leído como UTF-8 (archivos NAPSIS)."""
    try:
        return str(t).encode('latin-1').decode('utf-8')
    except Exception:
        return str(t)


_GRAVE_TO_ACUTE = str.maketrans('ÀÈÌÒÙàèìòù', 'ÁÉÍÓÚáéíóú')

def fix_accents(text):
    """Reemplaza tildes graves (Ì, À…) por agudas (Í, Á…) — español sólo usa agudas."""
    return str(text).translate(_GRAVE_TO_ACUTE)


def clean_nombre(text):
    """Nombre de persona: corrige encoding, corrige tildes, mayúsculas."""
    return fix_accents(fix_enc(str(text))).upper()


def norm_key(nombre):
    """
    Clave de matching entre fuentes: tokens del nombre ordenados
    alfabéticamente, sin acentos, en minúsculas.
    Ej: "PARGA GONZÁLEZ, OLIVIA FLORENCIA" → "florencia_gonzalez_olivia_parga"
    """
    s = fix_enc(str(nombre))
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    tokens = sorted(
        t.lower() for t in re.sub(r'[^a-zA-Z ]', ' ', s).split()
        if len(t) > 1
    )
    return '_'.join(tokens)


# Aliases nombre oficial → nombre social aceptado en NAPSIS.
# Se usan cuando un estudiante aparece con un nombre distinto en SEPA/DIA
# (nombre legal) que en NAPSIS (nombre social).
# Formato: {norm_key(nombre_legal): norm_key(nombre_social_napsis)}
_NAME_ALIASES: dict[str, str] = {
    'luz_munoz_sotomayor': 'lucas_munoz_sotomayor',
}


def gen_from_grade(grade_num, year):
    """Generación (año de egreso) dado grado 1-12 y año del archivo."""
    return year + (12 - grade_num)


def grade_from_gen(gen, year):
    """Grado 1-12 para una generación en un año dado. None si fuera de rango."""
    g = 12 - (gen - year)
    return g if 1 <= g <= 12 else None


def is_np(name):
    """True si la asignatura es No Promocional."""
    n = name.lower()
    return any(kw in n for kw in NP_KEYWORDS) or '(n.p' in n


# Abreviaciones para la pestaña Cursos (cabeceras de columna, espacio reducido).
# IMPORTANTE: 'educación física' debe ir ANTES de 'física' para no capturarse mutuamente.
_SHORT_MAP = [
    ('lengua y lit',        'L.Y LIT.'),
    ('lenguaje',            'LENG.'),
    ('educación física',    'EFI'),
    ('educacion fisica',    'EFI'),
    ('idioma',              'INGLÉS'),
    ('inglés',              'INGLÉS'),
    ('ingles',              'INGLÉS'),
    ('matemát',             'MAT.'),
    ('matemat',             'MAT.'),
    ('biolog',              'BIO.'),
    ('física',              'FÍS.'),
    ('fisica',              'FÍS.'),
    ('química',             'QUÍM.'),
    ('quimica',             'QUÍM.'),
    ('historia',            'HIST.'),
    ('filosofía',           'FIL.'),
    ('filosofia',           'FIL.'),
    ('tecnolog',            'TEC.'),
    ('música',              'MÚS.'),
    ('musica',              'MÚS.'),
    ('artes vis',           'A.VIS.'),
    ('artes',               'ARTES'),
    ('ciencias naturales',  'C.NAT.'),
    ('ciencias para',       'C.CIU.'),
    ('ciencias de la sal',  'C.SAL.'),
    ('educación ciudadana', 'ED.CIU.'),
    ('educacion ciudadana', 'ED.CIU.'),
    ('orientac',            'ORI.'),
]


# Variantes de nombre completo que deben unificarse a una forma canónica.
# El matching es por subcadena (lowercase, sin tildes graves).
_LABEL_CANONICAL = [
    ('idioma extranjero',  'INGLÉS'),
    ('idioma',             'INGLÉS'),
    ('educación física',   'EDUCACIÓN FÍSICA Y SALUD'),   # debe ir ANTES que 'física'
    ('educacion fisica',   'EDUCACIÓN FÍSICA Y SALUD'),
    ('física',             'FÍSICA'),                      # unifica física y diferenciado de física
    ('fisica',             'FÍSICA'),
    ('química',            'QUÍMICA'),                     # unifica química y química electivo
    ('quimica',            'QUÍMICA'),
    ('música',             'MÚSICA'),                      # unifica música con y sin tilde
    ('musica',             'MÚSICA'),
    # NOTA: Lenguaje (básica) vs Lengua y Literatura (media) se unifica en JS,
    # no aquí, porque la decisión depende del historial individual del estudiante.
]


def normalize_label(label):
    """Normaliza el nombre completo de una asignatura a su forma canónica."""
    l = fix_accents(label).lower()
    for k, v in _LABEL_CANONICAL:
        if k in l:
            return v
    return label


def short_label(name):
    n = fix_accents(name).lower()
    for k, v in _SHORT_MAP:
        if k in n:
            return v
    return name[:6].upper() + '.'


def _partial_match(key1, key2):
    """True si los tokens del nombre más corto están contenidos en el más largo."""
    t1, t2 = set(key1.split('_')), set(key2.split('_'))
    shorter, longer = (t1, t2) if len(t1) <= len(t2) else (t2, t1)
    return len(shorter) >= 2 and shorter.issubset(longer)


def _token_pair_similar(a, b, threshold=0.82):
    """True si dos tokens de nombre son el mismo nombre con variación ortográfica."""
    if a == b:
        return True
    if len(a) < 3 or len(b) < 3:
        return False
    return difflib.SequenceMatcher(None, a, b).ratio() >= threshold


def _profile_similarity(key1, key2):
    """
    Similaridad entre dos claves de nombre (0.0–1.0).
    Asigna cada token del más corto al token más similar del más largo;
    devuelve fracción de tokens que se emparejan con similitud >= threshold.
    """
    t1, t2 = key1.split('_'), key2.split('_')
    shorter, longer = (t1, t2) if len(t1) <= len(t2) else (t2, t1)
    if len(shorter) < 2:
        return 0.0
    matched = 0
    used = set()
    for tok in shorter:
        best_i, best_r = -1, 0.0
        for i, lt in enumerate(longer):
            if i in used:
                continue
            r = difflib.SequenceMatcher(None, tok, lt).ratio()
            if r > best_r:
                best_r, best_i = r, i
        if best_r >= 0.82:
            matched += 1
            used.add(best_i)
    return matched / len(shorter)


def _merge_two_profiles(keep, drop):
    """Fusiona 'drop' dentro de 'keep' sin sobreescribir datos existentes."""
    # NAPSIS: años distintos se unen; el mismo año se deja al que ya existe
    for year, data in drop['napsis'].items():
        if year not in keep['napsis']:
            keep['napsis'][year] = data
    # DIA  (estructura: dia[year][periodo][materia])
    for year, periodos in drop['dia'].items():
        keep['dia'].setdefault(year, {})
        for periodo, mats in periodos.items():
            keep['dia'][year].setdefault(periodo, {})
            for mat, ejes in mats.items():
                if mat not in keep['dia'][year][periodo]:
                    keep['dia'][year][periodo][mat] = ejes
    # SEPA
    for area in ('lenguaje', 'matematica'):
        for period, val in drop['sepa'][area].items():
            if period not in keep['sepa'][area]:
                keep['sepa'][area][period] = val
    # Preferir el nombre que tenga más datos NAPSIS (tiende a ser el oficial)
    if len(drop['napsis']) > len(keep['napsis']):
        keep['nombre'] = drop['nombre']
        keep['key']    = drop['key']


def merge_duplicate_profiles(profiles, verbose=True):
    """
    Pasa por cada generación y fusiona perfiles que parezcan el mismo estudiante
    con distinta ortografía del nombre (ej. 'Dominikovski' vs 'Dominikovskaia').
    Criterio: similaridad de tokens ≥ 0.80 Y al menos 2 tokens coincidentes.
    """
    total_merged = 0
    for gen, gen_profiles in profiles.items():
        keys = list(gen_profiles.keys())
        to_delete = set()

        for i in range(len(keys)):
            if keys[i] in to_delete:
                continue
            for j in range(i + 1, len(keys)):
                if keys[j] in to_delete:
                    continue
                sim = _profile_similarity(keys[i], keys[j])
                if sim >= 0.80:
                    # Verificar que al menos 2 tokens sean idénticos (mismo
                    # apellido o mismo primer nombre) para evitar falsos positivos
                    t1 = set(keys[i].split('_'))
                    t2 = set(keys[j].split('_'))
                    if len(t1 & t2) < 2:
                        continue
                    # Decide cuál es el perfil "principal" (el que tiene más NAPSIS)
                    pi = gen_profiles[keys[i]]
                    pj = gen_profiles[keys[j]]
                    if len(pj['napsis']) > len(pi['napsis']):
                        keep_k, drop_k = keys[j], keys[i]
                    else:
                        keep_k, drop_k = keys[i], keys[j]
                    _merge_two_profiles(gen_profiles[keep_k], gen_profiles[drop_k])
                    to_delete.add(drop_k)
                    total_merged += 1
                    if verbose:
                        print(f"  [fusión] '{gen_profiles[keep_k]['nombre']}' "
                              f"← '{gen_profiles[drop_k]['nombre']}' (gen {gen})")

        for k in to_delete:
            del gen_profiles[k]

    if total_merged and verbose:
        print(f"  → {total_merged} perfiles duplicados fusionados")


def merge_repeaters(profiles, verbose=True):
    """
    Detecta y fusiona estudiantes que repitieron año.
    Criterio: mismo nombre (exacto o fuzzy) en generaciones CONSECUTIVAS (gen G y G+1)
    Y mismo grado en dos años seguidos (grado[año Y] == grado[año Y+1]).
    El perfil resultante queda en la generación más reciente (año de egreso real).
    Soporta repeticiones múltiples por orden ascendente de generaciones.
    """
    total_merged = 0
    gens = sorted(profiles.keys())

    for idx in range(len(gens) - 1):
        gen      = gens[idx]
        next_gen = gen + 1
        if next_gen not in profiles:
            continue

        to_delete = []   # keys a eliminar de profiles[gen] al final del bloque

        for key, p_curr in list(profiles[gen].items()):

            # ── 1. Buscar contraparte en la generación siguiente ──────────
            p_next = profiles[next_gen].get(key)   # coincidencia exacta
            match_key = key

            if p_next is None:                     # coincidencia fuzzy
                best_key, best_sim = None, 0.0
                for k2, p2 in profiles[next_gen].items():
                    sim = _profile_similarity(key, k2)
                    t1  = set(key.split('_'))
                    t2  = set(k2.split('_'))
                    if sim >= 0.85 and len(t1 & t2) >= 2 and sim > best_sim:
                        best_sim = sim
                        best_key = k2
                if best_key:
                    p_next    = profiles[next_gen][best_key]
                    match_key = best_key

            if p_next is None:
                continue

            # ── 2. Confirmar repetición: mismo grado en años consecutivos ─
            is_repeater = False
            for yr, data in p_curr['napsis'].items():
                nxt = p_next['napsis'].get(yr + 1)
                if nxt and data['grade'] == nxt['grade']:
                    is_repeater = True
                    break

            if not is_repeater:
                continue

            # ── 3. Fusionar p_curr → p_next (quedarse en la gen más reciente)
            _merge_two_profiles(p_next, p_curr)
            p_next['gen'] = next_gen          # asegurar gen correcta
            to_delete.append(key)
            total_merged += 1
            if verbose:
                print(f"  [repitente] '{p_next['nombre']}' "
                      f"gen {gen}→{next_gen}")

        for k in to_delete:
            del profiles[gen][k]

    if total_merged and verbose:
        print(f"  → {total_merged} estudiantes repitentes unificados")


# ══════════════════════════════════════════════════════════════
# LECTOR NAPSIS
# ══════════════════════════════════════════════════════════════

def _parse_napsis_file(path, year, grade_num):
    try:
        raw = _read_file_bytes(path)
    except OSError as e:
        print(f'    ✗ {path.name}: no se pudo abrir ({e})')
        return None
    soup = BeautifulSoup(raw.decode('latin-1', errors='replace'), 'html.parser')

    tables = soup.find_all('table')
    if len(tables) < 3:
        return None

    rows = tables[2].find_all('tr')
    if len(rows) < 4:
        return None

    # Cabecera de asignaturas
    headers = [fix_enc(td.get_text(strip=True))
               for td in rows[0].find_all(['td', 'th'])]

    col_map, used_keys = [], {'n', 'pf'}
    for h in headers[2:-1]:
        if not h:
            continue
        base = re.sub(r'[^a-z0-9]', '',
                      unicodedata.normalize('NFD', h.lower())
                      .encode('ascii', 'ignore').decode())[:6] or 'subj'
        key = base
        i = 2
        while key in used_keys:
            key = base[:5] + str(i)
            i += 1
        used_keys.add(key)
        label_clean = normalize_label(fix_accents(h).upper())  # mayúsculas, tildes, variantes
        col_map.append({
            'key': key, 'label': label_clean,
            'short': short_label(label_clean), 'np': is_np(h),
        })

    # Profesor jefe
    teacher = ''
    hdr_text = fix_enc(tables[0].get_text())
    m = re.search(r'Profesor Jefe[:\s]+([^|]+)', hdr_text)
    if m:
        teacher = clean_nombre(m.group(1).strip())

    # Filas de estudiantes
    students = []
    for row in rows[3:]:
        cells = [fix_enc(td.get_text(strip=True))
                 for td in row.find_all(['td', 'th'])]
        if not cells or not cells[0].strip().isdigit():
            break
        nombre = clean_nombre(cells[1]) if len(cells) > 1 else ''
        try:
            pf = float(cells[-1].replace(',', '.'))
        except Exception:
            pf = None

        asig = {}
        for i, col in enumerate(col_map):
            ci = i + 2
            if ci < len(cells) - 1:
                try:
                    asig[col['key']] = float(cells[ci].replace(',', '.'))
                except Exception:
                    asig[col['key']] = None

        students.append({
            'nombre': nombre,
            'key': norm_key(nombre),
            'pf': pf,
            'asignaturas': asig,
        })

    # Vigentes
    vigentes = 0
    for row in rows:
        t = fix_enc(row.get_text())
        mv = re.search(r'Vigentes:(\d+)', t)
        if mv:
            vigentes = int(mv.group(1))
            break
    if not vigentes:
        vigentes = len(students)

    pf_vals = [s['pf'] for s in students if s['pf'] is not None]
    pf_prom = round(sum(pf_vals) / len(pf_vals), 2) if pf_vals else None

    return {
        'gen': gen_from_grade(grade_num, year),
        'year': year,
        'grade': grade_num,
        'grade_name': GRADE_NAMES.get(grade_num, f'{grade_num}°'),
        'teacher': teacher,
        'vigentes': vigentes,
        'pf_prom': pf_prom,
        'col_map': col_map,
        'students': students,
    }


def _parse_actas(actas_dir, year):
    """
    Lee los archivos TXT de la sección Actas de NAPSIS y devuelve una lista de
    records en el mismo formato que _parse_napsis_file(), uno por curso.

    Archivos requeridos en actas_dir:
      a<rdb>_1.TXT  — Nómina de estudiantes (RUN → nombre)
      a<rdb>_4_*.txt — Antecedentes académicos (notas por asignatura)
      a<rdb>_5_*.txt — Situación de promoción (promedio final, asistencia)
      a<rdb>_7.TXT  — Acta del curso (vigentes, profesor jefe)
      a<rdb>_8.TXT  — Catálogo de asignaturas (código → nombre)
    """
    # ── Mapa tipo_enseñanza + grado → grade_num (1-12) ─────────────────
    def to_grade(tipo, grado):
        if tipo == 110:
            return grado        # 1-8 básico
        if tipo == 310:
            return grado + 8    # 1-4 medio → 9-12
        return None

    # ── Archivo 8: código → nombre canónico ─────────────────────────────
    catalogo = {}
    for f8 in sorted(actas_dir.glob('a*_8.TXT')):
        with open(f8, encoding='latin-1', errors='replace') as fh:
            for line in fh:
                p = line.rstrip('\n').split('\t')
                if len(p) < 6:
                    continue
                cod  = p[4].strip()
                nombre = fix_accents(fix_enc(p[5].strip())).upper()
                if cod and nombre:
                    catalogo[cod] = normalize_label(nombre)

    # ── Archivo 1: RUN → nombre completo ────────────────────────────────
    run_to_nombre = {}
    for f1 in sorted(actas_dir.glob('a*_1.TXT')):
        with open(f1, encoding='latin-1', errors='replace') as fh:
            for line in fh:
                p = line.rstrip('\n').split('\t')
                if len(p) < 9:
                    continue
                run    = p[3].strip()
                ap_pat = fix_accents(fix_enc(p[5].strip())).upper()
                ap_mat = fix_accents(fix_enc(p[6].strip())).upper()
                noms   = fix_accents(fix_enc(p[7].strip())).upper()
                run_to_nombre[run] = f'{ap_pat} {ap_mat} {noms}'.strip()

    # ── Archivo 7: (tipo, grado) → {vigentes, teacher} ──────────────────
    curso_meta = {}   # key: (tipo_int, grado_int)
    for f7 in sorted(actas_dir.glob('a*_7.TXT')):
        with open(f7, encoding='latin-1', errors='replace') as fh:
            for line in fh:
                p = line.rstrip('\n').split('\t')
                if len(p) < 18:
                    continue
                try:
                    tipo  = int(p[3].strip())
                    grado = int(p[4].strip())
                    vigentes = int(p[8].strip())    # matrícula final
                except ValueError:
                    continue
                teacher = fix_accents(fix_enc(p[17].strip())).upper()
                curso_meta[(tipo, grado)] = {'vigentes': vigentes, 'teacher': teacher}

    # ── Archivo 5: (RUN, tipo, grado) → pf ──────────────────────────────
    pf_map = {}
    for f5 in sorted(actas_dir.glob('a*_5_*.txt')):
        with open(f5, encoding='latin-1', errors='replace') as fh:
            for line in fh:
                p = line.rstrip('\n').split('\t')
                if len(p) < 13:
                    continue
                try:
                    tipo  = int(p[3].strip())
                    grado = int(p[4].strip())
                    run   = p[7].strip()
                    pf    = float(p[9].strip().replace(',', '.')) if p[9].strip() else None
                    asist = p[10].strip()
                except (ValueError, IndexError):
                    continue
                # pf=0 con asistencia=0 → alumno retirado
                if pf == 0.0 and asist in ('0', ''):
                    pf = None
                pf_map[(run, tipo, grado)] = pf

    # ── Archivo 4: (tipo, grado) → { RUN: { cod: nota } } ───────────────
    notas_por_curso = {}   # (tipo, grado) → { run: { cod: nota } }
    asigs_por_curso = {}   # (tipo, grado) → set de códigos presentes
    for f4 in sorted(actas_dir.glob('a*_4_*.txt')):
        with open(f4, encoding='latin-1', errors='replace') as fh:
            for line in fh:
                p = line.rstrip('\n').split('\t')
                if len(p) < 13:
                    continue
                try:
                    tipo  = int(p[3].strip())
                    grado = int(p[4].strip())
                    run   = p[7].strip()
                    cod   = p[11].strip()
                    nota  = float(p[12].strip().replace(',', '.'))
                except (ValueError, IndexError):
                    continue
                ck = (tipo, grado)
                notas_por_curso.setdefault(ck, {}).setdefault(run, {})[cod] = nota
                asigs_por_curso.setdefault(ck, set()).add(cod)

    # ── Construir un record por curso ────────────────────────────────────
    records = []
    for (tipo, grado), notas_alumnos in sorted(notas_por_curso.items()):
        grade_num = to_grade(tipo, grado)
        if grade_num is None:
            continue

        meta = curso_meta.get((tipo, grado), {})

        # col_map: una entrada por asignatura presente en este curso
        codigos_curso = sorted(asigs_por_curso[(tipo, grado)],
                               key=lambda c: catalogo.get(c, c))
        col_map = []
        for cod in codigos_curso:
            label = catalogo.get(cod, f'COD_{cod}')
            col_map.append({
                'key':   f'a{cod}',
                'label': label,
                'short': short_label(label),
                'np':    is_np(label),
            })

        # Estudiantes
        students = []
        for run, asig_dict in notas_alumnos.items():
            nombre = run_to_nombre.get(run, run)
            pf     = pf_map.get((run, tipo, grado))
            asignaturas = {f'a{cod}': nota for cod, nota in asig_dict.items()}
            students.append({
                'nombre':      nombre,
                'key':         norm_key(nombre),
                'pf':          pf,
                'asignaturas': asignaturas,
            })

        pf_vals  = [s['pf'] for s in students if s['pf'] is not None]
        pf_prom  = round(sum(pf_vals) / len(pf_vals), 2) if pf_vals else None
        vigentes = meta.get('vigentes', len(students))
        teacher  = meta.get('teacher', '')

        records.append({
            'gen':        gen_from_grade(grade_num, year),
            'year':       year,
            'grade':      grade_num,
            'grade_name': GRADE_NAMES.get(grade_num, f'{grade_num}°'),
            'teacher':    teacher,
            'vigentes':   vigentes,
            'pf_prom':    pf_prom,
            'col_map':    col_map,
            'students':   students,
        })

    return records


def read_all_napsis():
    records = []
    if not NAPSIS_DIR.exists():
        print('  ⚠  napsis/ no encontrada')
        return records

    # Si existen archivos de Actas para el año actual, se usan en lugar de HTML
    actas_disponibles = ACTAS_DIR.is_dir() and any(ACTAS_DIR.glob('a*_4_*.txt'))

    for year_dir in sorted(NAPSIS_DIR.iterdir()):
        if not year_dir.is_dir() or not year_dir.name.isdigit():
            continue
        year = int(year_dir.name)

        if year == CURRENT_YEAR and actas_disponibles:
            print(f'  NAPSIS {year} — leyendo desde Actas (napsis/actas/)...')
            for rec in _parse_actas(ACTAS_DIR, year):
                records.append(rec)
                print(f'  NAPSIS {year} — {rec["grade_name"]:12s} '
                      f'{rec["vigentes"]:3d} est.  prom {rec["pf_prom"]}')
            continue

        for f in sorted(year_dir.glob('*.xlsx')):
            m = re.match(r'^(\d{2})A_Notas', f.name)
            if not m:
                continue
            grade_num = int(m.group(1))
            rec = _parse_napsis_file(f, year, grade_num)
            if rec:
                records.append(rec)
                print(f'  NAPSIS {year} — {rec["grade_name"]:12s} '
                      f'{rec["vigentes"]:3d} est.  prom {rec["pf_prom"]}')
    return records


# ══════════════════════════════════════════════════════════════
# LECTOR DIA
# ══════════════════════════════════════════════════════════════

def _grade_from_dia_folder(name):
    n = name.lower().strip()
    # Básico numérico: "2° básico", "8° básico"
    m = re.search(r'(\d+)[°\s]', n)
    if m and ('bás' in n or 'bas' in n):
        return int(m.group(1))
    # Enseñanza media con romano
    for roman, grade in [('iv', 12), ('iii', 11), ('ii', 10), ('i', 9)]:
        if re.search(rf'\b{roman}[°\s]', n):
            return grade
    return None


def _read_file_bytes(path):
    """
    Lee un archivo en modo binario de forma robusta en Windows.
    En rutas largas o con caracteres Unicode (Google Drive, tildes, grados, etc.)
    usa el prefijo \\\\?\\ que bypasea el limite MAX_PATH (260 chars) de Windows.
    """
    try:
        return path.read_bytes()
    except OSError:
        pass
    if sys.platform == 'win32':
        abs_str = str(path.absolute())
        prefix  = '\\\\?\\'
        if not abs_str.startswith(prefix):
            abs_str = prefix + abs_str
        try:
            with open(abs_str, 'rb') as f:
                return f.read()
        except OSError:
            pass
    raise OSError(f'No se pudo leer: {path}')


def _parse_dia_excel(path):
    try:
        data = _read_file_bytes(path)
    except OSError as e:
        print(f'    ✗ {path.name}: no se pudo abrir ({e})')
        return []
    engine = 'xlrd' if data[:4] == b'\xd0\xcf\x11\xe0' else 'openpyxl'
    try:
        df = pd.read_excel(io.BytesIO(data), engine=engine, header=None)
    except Exception as e:
        print(f'    ✗ {path.name}: {e}')
        return []

    # Fila de encabezado: contiene "Nombre" o "Lista"
    hdr_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row
                if str(v).strip() not in ('', 'nan')]
        if any('nombre' in v.lower() or 'lista' in v.lower() for v in vals):
            hdr_row = i
            break
    if hdr_row is None:
        return []

    headers = [str(v).strip() for v in df.iloc[hdr_row]]

    results = []
    for i in range(hdr_row + 1, len(df)):
        row = df.iloc[i]
        try:
            int(float(str(row.iloc[0]).strip()))
        except Exception:
            break
        nombre = clean_nombre(str(row.iloc[1]))
        if nombre in ('', 'NAN'):
            continue
        ejes = {}
        for j, h in enumerate(headers[2:], start=2):
            if h in ('', 'nan'):
                continue
            try:
                ejes[h] = round(float(str(row.iloc[j]).replace(',', '.')), 1)
            except Exception:
                pass
        results.append({'nombre': nombre, 'key': norm_key(nombre), 'ejes': ejes})
    return results


# Períodos DIA en orden cronológico: diagnóstico → monitoreo → cierre
_DIA_PERIODOS = [
    ('diagn',   'diagnostico', 'Diagnóstico', 1),
    ('monitor', 'monitoreo',   'Monitoreo',   2),
    ('cierr',   'cierre',      'Cierre',      3),
]

DIA_PERIODO_LABEL = {key: label for _, key, label, _ in _DIA_PERIODOS}
DIA_PERIODO_ORDER = {key: order for _, key, _, order in _DIA_PERIODOS}


def _periodo_from_folder(name):
    """Detecta el período DIA (diagnostico/monitoreo/cierre) a partir del nombre de carpeta."""
    n = name.lower()
    for kw, key, label, order in _DIA_PERIODOS:
        if kw in n:
            return key, order
    return None, 99


def read_all_dia():
    records = []
    if not DIA_DIR.exists():
        return records

    for year_dir in sorted(DIA_DIR.iterdir()):
        if not year_dir.is_dir():
            continue
        m = re.search(r'(\d{4})', year_dir.name)
        if not m:
            continue
        year = int(m.group(1))

        # Recopilar subcarpetas de período (Diagnóstico, Monitoreo, Cierre)
        # y ordenarlas cronológicamente.
        period_dirs = []
        for d in year_dir.iterdir():
            if not d.is_dir():
                continue
            pkey, porder = _periodo_from_folder(d.name)
            if pkey:
                period_dirs.append((porder, pkey, d))

        # Si no hay subcarpetas de período, tratar el directorio año como raíz
        if not period_dirs:
            period_dirs = [(1, 'diagnostico', year_dir)]

        period_dirs.sort(key=lambda x: x[0])

        for porder, pkey, search_root in period_dirs:
            for curso_dir in sorted(search_root.iterdir()):
                if not curso_dir.is_dir():
                    continue
                grade = _grade_from_dia_folder(curso_dir.name)
                if not grade:
                    continue
                gen = gen_from_grade(grade, year)

                for f in curso_dir.iterdir():
                    if f.suffix.lower() not in ('.xls', '.xlsx'):
                        continue
                    fn = f.name.upper()
                    if 'LECTURA' in fn:
                        materia = 'lectura'
                    elif 'MATEM' in fn:
                        materia = 'matematica'
                    else:
                        continue

                    students = _parse_dia_excel(f)
                    if students:
                        records.append({
                            'gen': gen, 'year': year, 'grade': grade,
                            'grade_name': GRADE_NAMES.get(grade, f'{grade}°'),
                            'periodo': pkey,
                            'materia': materia, 'students': students,
                        })
                        print(f'  DIA  {year} [{DIA_PERIODO_LABEL[pkey][:5]}] '
                              f'— {GRADE_NAMES.get(grade):12s} '
                              f'{materia:10s}  {len(students)} est.')
    return records


# ══════════════════════════════════════════════════════════════
# LECTOR SEPA
# ══════════════════════════════════════════════════════════════

def _grade_from_sepa_folder(name):
    n = name.lower()
    m = re.search(r'(\d+)\s+b', n)
    if m:
        return int(m.group(1))
    for roman, grade in [('iii', 11), ('ii', 10), ('iv', 12), ('i', 9)]:
        if f' {roman} ' in n or f' {roman}m' in n or n.endswith(f' {roman}'):
            return grade
    return None


def _parse_col_key(header):
    """'2022\n(Principios\ndeaño)' → '2022_P',  '2023\n(Finales...' → '2023_F'"""
    clean = header.replace('\n', ' ')
    ym = re.search(r'(\d{4})', clean)
    if not ym:
        return None
    yr = ym.group(1)
    period = 'P' if 'rincipio' in clean or 'nicio' in clean.lower() else 'F'
    return f'{yr}_{period}'


def _parse_sepa_pdf(path):
    try:
        data = _read_file_bytes(path)
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            nombre_raw = path.stem      # "Olivia Florencia Parga Gonzalez"
            nombre = clean_nombre(nombre_raw)
            result = {
                'nombre': nombre,
                'key': norm_key(nombre_raw),   # norm_key opera sobre el texto original
                'lenguaje': {},
                'matematica': {},
            }
            for page in pdf.pages:
                text = page.extract_text() or ''
                # Detectar sección por primera palabra del cuerpo
                is_leng = bool(re.search(r'^Lenguaje', text.strip()))
                is_mat  = bool(re.search(r'^Matemática', text.strip()))
                if not is_leng and not is_mat:
                    continue

                for tbl in page.extract_tables():
                    if not tbl or len(tbl) < 2:
                        continue
                    headers = tbl[0]

                    # Fila del estudiante (la primera que no es referencia)
                    ref_keys = {'establecimiento', 'mun', 'ps', 'pp', ''}
                    student_row = next(
                        (r for r in tbl[1:]
                         if r and r[0] and r[0].strip().lower() not in ref_keys),
                        None
                    )
                    if not student_row:
                        continue

                    scores = {}
                    for j, h in enumerate(headers):
                        if not h or j == 0:
                            continue
                        col_key = _parse_col_key(h)
                        if not col_key:
                            continue
                        try:
                            scores[col_key] = int(float(
                                str(student_row[j]).replace(',', '.')))
                        except Exception:
                            pass

                    if scores:
                        if is_leng:
                            result['lenguaje'].update(scores)
                        else:
                            result['matematica'].update(scores)

            return result if (result['lenguaje'] or result['matematica']) else None
    except Exception as e:
        print(f'    ✗ {path.name}: {e}')
        return None


def read_all_sepa():
    records = []
    if not SEPA_DIR.exists():
        return records

    for folder in sorted(SEPA_DIR.iterdir()):
        if not folder.is_dir():
            continue
        grade = _grade_from_sepa_folder(folder.name)
        if not grade:
            continue

        pdfs = sorted(folder.glob('*.pdf'))
        if not pdfs:
            continue

        # Año del reporte desde el primer PDF
        sepa_year = None
        try:
            data = _read_file_bytes(pdfs[0])
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                t = pdf.pages[0].extract_text() or ''
                m = re.search(r'Año\s+(\d{4})', t)
                if m:
                    sepa_year = int(m.group(1))
        except Exception:
            pass
        if not sepa_year:
            sepa_year = CURRENT_YEAR - 1

        gen = gen_from_grade(grade, sepa_year)

        ok = 0
        for pdf_path in pdfs:
            r = _parse_sepa_pdf(pdf_path)
            if r:
                r['gen'] = gen
                r['grade'] = grade
                r['sepa_year'] = sepa_year
                records.append(r)
                ok += 1

        print(f'  SEPA {sepa_year} — {GRADE_NAMES.get(grade):12s} '
              f'{ok}/{len(pdfs)} PDFs leídos')
    return records


# ══════════════════════════════════════════════════════════════
# VINCULADOR DE PERFILES
# ══════════════════════════════════════════════════════════════

def build_profiles(napsis_records, dia_records, sepa_records):
    """
    Devuelve dict { gen: { name_key: profile } }
    Cada perfil contiene toda la información disponible del estudiante.
    """
    profiles = defaultdict(dict)

    # 1. Sembrar desde NAPSIS
    for rec in napsis_records:
        gen = rec['gen']
        for s in rec['students']:
            key = s['key']
            if key not in profiles[gen]:
                profiles[gen][key] = {
                    'nombre': s['nombre'],
                    'key': key,
                    'gen': gen,
                    'napsis': {},
                    'dia': {},
                    'sepa': {'lenguaje': {}, 'matematica': {}},
                }
            profiles[gen][key]['napsis'][rec['year']] = {
                'grade': rec['grade'],
                'grade_name': rec['grade_name'],
                'pf': s['pf'],
                'asignaturas': s['asignaturas'],
                'col_map': rec['col_map'],
            }

    def _find_or_create(gen, key, nombre):
        # Redirigir alias de nombre oficial → nombre social (ej. Luz → Lucas)
        if key in _NAME_ALIASES:
            alias_target = _NAME_ALIASES[key]
            if alias_target in profiles[gen]:
                print(f'  [alias] {key!r} → {alias_target!r} (gen {gen})')
                return profiles[gen][alias_target]
        if key in profiles[gen]:
            return profiles[gen][key]
        # Fuzzy: buscar token subset
        match = next(
            (p for k, p in profiles[gen].items() if _partial_match(key, k)),
            None
        )
        if match:
            return match
        # Crear perfil nuevo
        profiles[gen][key] = {
            'nombre': nombre, 'key': key, 'gen': gen,
            'napsis': {}, 'dia': {},
            'sepa': {'lenguaje': {}, 'matematica': {}},
        }
        return profiles[gen][key]

    # 2. Agregar DIA  (estructura: dia[year][periodo][materia])
    for rec in dia_records:
        gen = rec['gen']
        for s in rec['students']:
            p = _find_or_create(gen, s['key'], s['nombre'])
            p['dia'].setdefault(rec['year'], {})
            p['dia'][rec['year']].setdefault(rec['periodo'], {})
            p['dia'][rec['year']][rec['periodo']][rec['materia']] = s['ejes']

    # 3. Agregar SEPA
    for s in sepa_records:
        p = _find_or_create(s['gen'], s['key'], s['nombre'])
        p['sepa']['lenguaje'].update(s['lenguaje'])
        p['sepa']['matematica'].update(s['matematica'])

    # 4. Fusionar perfiles duplicados (variaciones ortográficas del mismo nombre)
    merge_duplicate_profiles(profiles)

    # 5. Fusionar repitentes (mismo estudiante en generaciones consecutivas)
    merge_repeaters(profiles)

    return profiles


# ══════════════════════════════════════════════════════════════
# CÁLCULO DE RIESGO
# ══════════════════════════════════════════════════════════════

def _risk_level(pf, reds):
    if pf is None:
        return None
    if pf < 4.0:                  return 'alerta'
    if reds >= 3:                  return 'alerta'
    if reds == 2 and pf < 5.0:    return 'alerta'
    if reds == 1 and pf < 4.5:    return 'alerta'
    if reds >= 1:                  return 'activo'
    if pf <= 5.5:                  return 'preventivo'
    return None


def calc_risk(profile, year):
    napsis = profile['napsis'].get(year)
    if not napsis:
        return None
    pf = napsis['pf']
    reds = sum(
        1 for c in napsis['col_map']
        if not c['np']
        and napsis['asignaturas'].get(c['key']) is not None
        and napsis['asignaturas'][c['key']] < 4.0
    )
    return _risk_level(pf, reds)


# ══════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL OBJETO DE DATOS PARA EL HTML
# ══════════════════════════════════════════════════════════════

def build_js_data(napsis_records, dia_records, sepa_records, profiles):
    """
    Construye el dict que se serializa como JSON en el HTML.
    """

    # ── Resumen de cursos del año actual ──────────────────────
    cursos_actuales = []
    for rec in sorted(
        (r for r in napsis_records if r['year'] == CURRENT_YEAR),
        key=lambda r: r['grade']
    ):
        gen = rec['gen']
        gen_profiles = list(profiles[gen].values())

        counts = {'alerta': 0, 'activo': 0, 'preventivo': 0}
        for p in gen_profiles:
            lv = calc_risk(p, CURRENT_YEAR)
            if lv:
                counts[lv] += 1

        # DIA del año actual para este curso — agrupado por período
        # Estructura: {pkey: {materia: {eje: promedio}}}
        dia_periodos: dict = {}
        for drec in dia_records:
            if drec['gen'] == gen and drec['year'] == CURRENT_YEAR:
                pkey = drec['periodo']
                ejes_avg: dict = {}
                for s in drec['students']:
                    for eje, val in s['ejes'].items():
                        ejes_avg.setdefault(eje, []).append(val)
                avg = {e: round(sum(v) / len(v), 1)
                       for e, v in ejes_avg.items() if v}
                dia_periodos.setdefault(pkey, {})[drec['materia']] = avg

        cursos_actuales.append({
            'gen': gen,
            'grade': rec['grade'],
            'grade_name': rec['grade_name'],
            'teacher': rec['teacher'],
            'vigentes': rec['vigentes'],
            'pf_prom': rec['pf_prom'],
            'alerta': counts['alerta'],
            'activo': counts['activo'],
            'preventivo': counts['preventivo'],
            'dia': dia_periodos,
        })

    # ── Trayectorias longitudinales por generación ────────────
    # Para cada generación, promedio final por año
    trayectorias = {}
    years_available = sorted({r['year'] for r in napsis_records})
    for rec in napsis_records:
        gen = rec['gen']
        if gen not in trayectorias:
            trayectorias[gen] = {}
        trayectorias[gen][rec['year']] = rec['pf_prom']

    # ── Perfiles completos de estudiantes ─────────────────────
    estudiantes = []
    for gen, gen_dict in sorted(profiles.items()):
        for key, p in sorted(gen_dict.items(),
                              key=lambda kv: kv[1]['nombre']):
            # Serializar col_map dentro de napsis (sólo label/short/np)
            napsis_serial = {}
            for yr, nd in p['napsis'].items():
                napsis_serial[yr] = {
                    'grade': nd['grade'],
                    'grade_name': nd['grade_name'],
                    'pf': nd['pf'],
                    'asignaturas': nd['asignaturas'],
                    'col_map': [
                        {'key': c['key'], 'label': c['label'],
                         'short': c['short'], 'np': c['np']}
                        for c in nd['col_map']
                    ],
                    'risk': calc_risk(p, yr),
                }

            estudiantes.append({
                'nombre': p['nombre'],
                'key': key,
                'gen': gen,
                'napsis': napsis_serial,
                'dia': p['dia'],
                'sepa': p['sepa'],
            })

    return {
        'school': SCHOOL_NAME,
        'current_year': CURRENT_YEAR,
        'years': years_available,
        'cursos': cursos_actuales,
        'trayectorias': trayectorias,
        'estudiantes': estudiantes,
    }


# ══════════════════════════════════════════════════════════════
# GENERADOR HTML
# ══════════════════════════════════════════════════════════════

def _get_css():
    """Lee el CSS del template existente."""
    tmpl = TEMPLATE_DIR / 'dashboard_ejemplo.html'
    if not tmpl.exists():
        return ''
    try:
        raw = _read_file_bytes(tmpl)
    except OSError:
        return ''
    soup = BeautifulSoup(raw.decode('utf-8', errors='replace'), 'html.parser')
    style = soup.find('style')
    return style.string.strip() if style else ''


def generate_html(js_data):
    css = _get_css()
    data_json = json.dumps(js_data, ensure_ascii=False)
    today = date.today().strftime('%d/%m/%Y')

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Institucional · {SCHOOL_NAME} · {CURRENT_YEAR}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
<style>
{css}

/* ── Extras dashboard institucional ── */
.gen-badge {{
  display:inline-block;background:var(--surface);border:1px solid var(--border);
  border-radius:4px;padding:1px 7px;font-size:10px;font-family:var(--mono);
  color:var(--muted);text-transform:uppercase;letter-spacing:.5px;
}}
.course-select {{
  background:var(--card);border:1px solid var(--border);border-radius:6px;
  padding:9px 14px;color:var(--text);font-family:var(--sans);font-size:13px;
  outline:none;cursor:pointer;min-width:220px;
}}
.course-select:focus {{ border-color:var(--accent); }}
.summary-table {{ width:100%;border-collapse:collapse; }}
.summary-table th {{
  text-align:left;font-family:var(--mono);font-size:10px;
  text-transform:uppercase;letter-spacing:1px;color:var(--muted);
  padding:8px 10px;border-bottom:2px solid var(--border);
}}
.summary-table td {{ padding:9px 10px;border-bottom:1px solid var(--surface);font-size:13px; }}
.summary-table tr:hover td {{ background:var(--surface); }}
.dia-bar-wrap {{ display:flex;flex-direction:column;gap:6px;margin-top:8px; }}
.dia-bar-row {{ display:flex;align-items:center;gap:8px;font-size:11px; }}
.dia-bar-label {{ width:160px;color:var(--muted);font-family:var(--mono);
  font-size:10px;text-align:right;flex-shrink:0; }}
.dia-bar-outer {{ flex:1;background:var(--bg);border-radius:3px;height:10px; }}
.dia-bar-fill {{ height:10px;border-radius:3px;transition:width .4s; }}
.dia-val {{ width:38px;text-align:right;font-family:var(--mono);font-size:11px;
  color:var(--text); }}
.tray-chart-wrap {{ position:relative;height:260px; }}
.student-profile {{ display:none; }}
.student-profile.active {{ display:block; }}
.profile-section {{ margin-bottom:24px; }}
.sepa-chart-wrap {{ position:relative;height:220px; }}
.years-table {{ width:100%;border-collapse:collapse;font-size:12px;margin-top:12px; }}
.years-table th {{
  font-size:10px;font-family:var(--mono);letter-spacing:1px;color:var(--muted);
  text-align:center;padding:6px 8px;border-bottom:2px solid var(--border);
}}
.years-table th:first-child {{ text-align:left; }}
.years-table td {{
  padding:8px;text-align:center;border-bottom:1px solid rgba(184,208,232,.6);
  font-family:var(--mono);font-size:12px;
}}
.years-table td:first-child {{ text-align:left;font-family:var(--sans);font-weight:500; }}
.no-data {{ color:var(--muted);font-style:italic;font-size:13px;padding:16px 0; }}
.risk-dot-sm {{ display:inline-block;width:8px;height:8px;border-radius:50%;
  margin-right:4px;vertical-align:middle; }}
.dia-toggle {{
  background:var(--surface);border:1px solid var(--border);border-radius:4px;
  padding:4px 11px;font-size:10px;font-family:var(--mono);color:var(--muted);
  cursor:pointer;text-transform:uppercase;letter-spacing:.5px;transition:all .15s;
}}
.dia-toggle.active {{
  background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600;
}}
</style>
</head>
<body>

<header>
  <div class="hdr-left">
    <span class="chip">{SCHOOL_NAME}</span>
    <h1>Dashboard Institucional · {CURRENT_YEAR}</h1>
    <span class="hdr-meta">Actualizado: {today}</span>
  </div>
  <div class="hdr-stats" id="hdr-stats"></div>
</header>

<nav>
  <button class="active" onclick="showTab('colegio',this)">Colegio</button>
  <button onclick="showTab('cursos',this)">Cursos</button>
  <button onclick="showTab('estudiantes',this)">Estudiantes</button>
</nav>

<main>

<!-- ══════════════════ COLEGIO ══════════════════ -->
<div class="tab-content active" id="tab-colegio">
  <div class="kpi-row" id="kpi-colegio"></div>

  <div class="card mb24">
    <div class="card-title">Situación por Curso · {CURRENT_YEAR}</div>
    <div style="overflow-x:auto">
      <table class="summary-table" id="tbl-cursos"></table>
    </div>
  </div>

  <div class="g2 mb24">
    <div class="card">
      <div class="card-title">DIA 2026 — Lectura · Promedio por eje por curso</div>
      <div id="dia-lec-bars"></div>
    </div>
    <div class="card">
      <div class="card-title">DIA 2026 — Matemática · Promedio por eje por curso</div>
      <div id="dia-mat-bars"></div>
    </div>
  </div>

  <div class="card mb24">
    <div class="card-title">Trayectoria longitudinal · Promedio final por generación</div>
    <div class="tray-chart-wrap"><canvas id="chart-trayectoria"></canvas></div>
  </div>
</div>

<!-- ══════════════════ CURSOS ══════════════════ -->
<div class="tab-content" id="tab-cursos">
  <div style="margin-bottom:20px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
    <span style="font-size:11px;font-family:var(--mono);color:var(--muted);
      letter-spacing:1px;text-transform:uppercase;">Seleccionar curso</span>
    <select class="course-select" id="curso-select" onchange="renderCurso()"></select>
  </div>
  <div id="curso-content"></div>
</div>

<!-- ══════════════════ ESTUDIANTES ══════════════════ -->
<div class="tab-content" id="tab-estudiantes">
  <div class="student-search mb24">
    <input type="text" class="search-input" id="est-search"
      placeholder="Buscar estudiante por nombre…"
      oninput="searchStudent(this.value)" autocomplete="off"
      style="width:340px">
    <span style="font-size:11px;color:var(--muted);font-family:var(--mono)" id="est-count"></span>
  </div>
  <div id="est-results"></div>
  <div id="est-profile"></div>
</div>

</main>

<script>
// ════════════════════════════════════════════════════════════
// DATOS
// ════════════════════════════════════════════════════════════
const DATA = {data_json};

const CURSOS    = DATA.cursos;
const TRAY      = DATA.trayectorias;
const EST       = DATA.estudiantes;
const CUR_YEAR  = DATA.current_year;
const YEARS     = DATA.years;

// ════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════
const nc  = v => v===null?'':v<4?'nf':v<5?'nw':v<6?'no':'ng';
const ns  = v => v===null?'—':v.toFixed(1);
const pct = v => v===null?'—':(+v).toFixed(1)+'%';

const RISK_COLOR = {{alerta:'var(--red)',activo:'var(--yellow)',preventivo:'var(--accent)'}};
const RISK_LABEL = {{alerta:'Alerta',activo:'Riesgo activo',preventivo:'Seg. preventivo'}};

function diaColor(v) {{
  if(v===null||v===undefined) return 'var(--muted)';
  if(v<40) return 'var(--red)';
  if(v<60) return 'var(--yellow)';
  if(v<80) return 'var(--accent)';
  return 'var(--green)';
}}

function sepaColor(v, base) {{
  // base ≈ 650 promedio establecimiento
  if(v===null||v===undefined) return 'var(--muted)';
  if(v < base - 40) return 'var(--red)';
  if(v < base - 10) return 'var(--yellow)';
  if(v < base + 20) return 'var(--accent)';
  return 'var(--green)';
}}

// Paleta para líneas de trayectoria
const LINE_COLORS = [
  '#1a5276','#c0392b','#1e8449','#b7770d','#6c3483',
  '#117a65','#784212','#1a5276','#922b21','#1d8348',
  '#9b59b6','#2e86c1','#d35400','#148f77',
];

// ════════════════════════════════════════════════════════════
// TABS
// ════════════════════════════════════════════════════════════
function showTab(id, btn) {{
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  if(btn) btn.classList.add('active');
  if(id==='colegio' && !window._colegioDrawn) {{ renderColegio(); window._colegioDrawn=true; }}
  if(id==='cursos'  && !window._cursosInit)   {{ initCursos();  window._cursosInit=true; }}
}}

// ════════════════════════════════════════════════════════════
// PESTAÑA COLEGIO
// ════════════════════════════════════════════════════════════
function renderColegio() {{
  // KPIs
  const total    = CURSOS.reduce((a,c)=>a+c.vigentes,0);
  const alerta   = CURSOS.reduce((a,c)=>a+c.alerta,0);
  const activo   = CURSOS.reduce((a,c)=>a+c.activo,0);
  const prev     = CURSOS.reduce((a,c)=>a+c.preventivo,0);
  const pfs      = CURSOS.map(c=>c.pf_prom).filter(v=>v!==null);
  const promGen  = pfs.length ? (pfs.reduce((a,b)=>a+b)/pfs.length).toFixed(1) : '—';

  document.getElementById('hdr-stats').innerHTML = `
    <div class="stat-pill"><div class="num ca">${{total}}</div><div class="lab">Vigentes</div></div>
    <div class="stat-pill"><div class="num cr">${{alerta}}</div><div class="lab">Alerta</div></div>
    <div class="stat-pill"><div class="num cy">${{activo}}</div><div class="lab">Riesgo activo</div></div>
    <div class="stat-pill"><div class="num ca">${{CURSOS.length}}</div><div class="lab">Cursos</div></div>`;

  document.getElementById('kpi-colegio').innerHTML = `
    <div class="kpi"><div class="value ca">${{total}}</div><div class="desc">Total vigentes</div></div>
    <div class="kpi"><div class="value cr">${{alerta}}</div><div class="desc">En alerta</div></div>
    <div class="kpi"><div class="value cy">${{activo}}</div><div class="desc">Riesgo activo</div></div>
    <div class="kpi"><div class="value ca">${{prev}}</div><div class="desc">Seg. preventivo</div></div>
    <div class="kpi"><div class="value cg">${{promGen}}</div><div class="desc">Prom. general</div></div>`;

  // Tabla resumen por curso
  const tbl = document.getElementById('tbl-cursos');
  tbl.innerHTML = `<thead><tr>
    <th>Curso</th><th>Generación</th><th>Profesor/a acompañante</th>
    <th style="text-align:center">Vigentes</th>
    <th style="text-align:center;color:var(--red)">Alerta</th>
    <th style="text-align:center;color:var(--yellow)">Riesgo</th>
    <th style="text-align:center;color:var(--accent)">Prev.</th>
    <th style="text-align:center">Prom.</th>
    <th style="text-align:center">DIA Lect.</th>
    <th style="text-align:center">DIA Mat.</th>
  </tr></thead>`;
  // Devuelve el promedio de ejes del período más avanzado disponible
  function diaLatestAvg(curso, materia) {{
    const order = ['cierre','monitoreo','diagnostico'];
    for(const p of order) {{
      const d = curso.dia?.[p]?.[materia];
      if(d && Object.keys(d).length) {{
        const vals = Object.values(d);
        return vals.reduce((a,b)=>a+b,0)/vals.length;
      }}
    }}
    return null;
  }}

  const tbody = document.createElement('tbody');
  CURSOS.forEach(c => {{
    const diaL = diaLatestAvg(c,'lectura');
    const diaM = diaLatestAvg(c,'matematica');
    const pfCol = c.pf_prom ? nc(c.pf_prom) : '';
    tbody.innerHTML += `<tr>
      <td><strong>${{c.grade_name}}</strong></td>
      <td><span class="gen-badge">Gen. ${{c.gen}}</span></td>
      <td style="font-size:12px;color:var(--muted)">${{c.teacher}}</td>
      <td style="text-align:center;font-family:var(--mono)">${{c.vigentes}}</td>
      <td style="text-align:center;font-family:var(--mono);color:var(--red);font-weight:${{c.alerta?700:400}}">${{c.alerta||'—'}}</td>
      <td style="text-align:center;font-family:var(--mono);color:var(--yellow);font-weight:${{c.activo?700:400}}">${{c.activo||'—'}}</td>
      <td style="text-align:center;font-family:var(--mono);color:var(--accent)">${{c.preventivo||'—'}}</td>
      <td style="text-align:center;font-family:var(--mono)" class="${{pfCol}}">${{ns(c.pf_prom)}}</td>
      <td style="text-align:center;font-family:var(--mono);color:${{diaColor(diaL)}}">${{diaL!==null?diaL.toFixed(1)+'%':'—'}}</td>
      <td style="text-align:center;font-family:var(--mono);color:${{diaColor(diaM)}}">${{diaM!==null?diaM.toFixed(1)+'%':'—'}}</td>
    </tr>`;
  }});
  tbl.appendChild(tbody);

  // Barras DIA
  renderDiaBars('dia-lec-bars', 'lectura');
  renderDiaBars('dia-mat-bars', 'matematica');

  // Gráfico trayectorias
  drawTrayectoria();
}}

const _DIA_PERIODO_KEYS   = ['diagnostico','monitoreo','cierre'];
const _DIA_PERIODO_LABELS = {{diagnostico:'Diagnóstico',monitoreo:'Monitoreo',cierre:'Cierre'}};
const _DIA_SHORT          = {{diagnostico:'Diagn.',monitoreo:'Monit.',cierre:'Cierre'}};

// Alterna entre la vista de trayectoria (gráfico) y el detalle (barras por eje)
function diaToggle(view) {{
  const traj = document.getElementById('dia-traj-view');
  const det  = document.getElementById('dia-det-view');
  const btnT = document.getElementById('btn-dia-traj');
  const btnD = document.getElementById('btn-dia-det');
  if(!traj) return;
  traj.style.display = view === 'traj' ? '' : 'none';
  det.style.display  = view === 'det'  ? '' : 'none';
  btnT.classList.toggle('active', view === 'traj');
  btnD.classList.toggle('active', view === 'det');
}}

function renderDiaBars(containerId, materia) {{
  const container = document.getElementById(containerId);
  container.innerHTML = '';

  // Períodos con datos para esta materia
  const activePeriodos = _DIA_PERIODO_KEYS.filter(p =>
    CURSOS.some(c => c.dia?.[p]?.[materia] && Object.keys(c.dia[p][materia]).length)
  );
  if(!activePeriodos.length) {{
    container.innerHTML = '<div class="no-data">Sin datos DIA disponibles</div>';
    return;
  }}

  activePeriodos.forEach(pkey => {{
    // Subtítulo de período
    const hdr = document.createElement('div');
    hdr.style.cssText = 'font-size:10px;font-family:var(--mono);color:var(--muted);' +
      'text-transform:uppercase;letter-spacing:1px;margin:10px 0 4px';
    hdr.textContent = _DIA_PERIODO_LABELS[pkey];
    container.appendChild(hdr);

    // Recolectar ejes para este período/materia
    const allEjes = new Set();
    CURSOS.forEach(c => {{
      const d = c.dia?.[pkey]?.[materia];
      if(d) Object.keys(d).forEach(e => allEjes.add(e));
    }});

    allEjes.forEach(eje => {{
      const row = document.createElement('div');
      row.className = 'dia-bar-row';
      row.innerHTML = `<span class="dia-bar-label">${{eje}}</span>`;
      CURSOS.forEach(c => {{
        const d = c.dia?.[pkey]?.[materia];
        if(!d) return;
        const val = d[eje];
        if(val === undefined) return;
        const col = diaColor(val);
        row.innerHTML += `
          <div style="display:flex;align-items:center;gap:4px;flex:1;min-width:80px">
            <span style="font-size:9px;font-family:var(--mono);color:var(--muted);
              width:52px;text-align:right">${{c.grade_name}}</span>
            <div class="dia-bar-outer">
              <div class="dia-bar-fill" style="width:${{val}}%;background:${{col}}"></div>
            </div>
            <span class="dia-val" style="color:${{col}}">${{val}}%</span>
          </div>`;
      }});
      container.appendChild(row);
    }});
  }});
}}

let _trayChart = null;
function drawTrayectoria() {{
  if(_trayChart) {{ _trayChart.destroy(); _trayChart=null; }}
  const gens = Object.keys(TRAY).map(Number).sort((a,b)=>a-b);
  const datasets = gens.map((gen, idx) => {{
    const data = YEARS.map(y => TRAY[gen][y] ?? null);
    const col = LINE_COLORS[idx % LINE_COLORS.length];
    return {{
      label: `Gen. ${{gen}}`,
      data, borderColor: col, backgroundColor: col+'22',
      borderWidth: 2, pointRadius: 4, pointHoverRadius: 6,
      tension: 0.3, spanGaps: false,
    }};
  }});
  _trayChart = new Chart(document.getElementById('chart-trayectoria'), {{
    type: 'line',
    data: {{ labels: YEARS, datasets }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ position:'right', labels:{{ font:{{family:'IBM Plex Mono',size:10}},
          color:'#5a7a96', boxWidth:12 }} }},
      }},
      scales: {{
        x: {{ grid:{{ color:'rgba(184,208,232,.5)' }},
              ticks:{{ color:'#5a7a96',font:{{family:'IBM Plex Mono',size:10}} }} }},
        y: {{ min:1, max:7,
              grid:{{ color:'rgba(184,208,232,.5)' }},
              ticks:{{ color:'#5a7a96',font:{{family:'IBM Plex Mono',size:10}},stepSize:.5 }} }},
      }},
    }},
  }});
}}

// ════════════════════════════════════════════════════════════
// PESTAÑA CURSOS
// ════════════════════════════════════════════════════════════
function initCursos() {{
  const sel = document.getElementById('curso-select');
  CURSOS.forEach(c => {{
    const opt = document.createElement('option');
    opt.value = c.gen;
    opt.textContent = `${{c.grade_name}} — Gen. ${{c.gen}}`;
    sel.appendChild(opt);
  }});
  renderCurso();
}}

let _cursoCharts = [];
function renderCurso() {{
  _cursoCharts.forEach(ch => ch.destroy());
  _cursoCharts = [];

  const gen = parseInt(document.getElementById('curso-select').value);
  const curso = CURSOS.find(c => c.gen === gen);
  if(!curso) return;

  // Estudiantes del curso en año actual
  const students = EST.filter(e => e.gen === gen && e.napsis[CUR_YEAR]);
  students.sort((a,b) => {{
    const pa = a.napsis[CUR_YEAR]?.pf ?? 99;
    const pb = b.napsis[CUR_YEAR]?.pf ?? 99;
    return pa - pb;
  }});

  const col_map = students[0]?.napsis[CUR_YEAR]?.col_map ?? [];
  const active_cols = col_map.filter(c => {{
    if(c.np) return false;
    return students.some(s => {{
      const v = s.napsis[CUR_YEAR]?.asignaturas?.[c.key];
      return v !== null && v !== undefined;
    }});
  }});

  // Calcular estadísticas de asignatura
  const subjStats = active_cols.map(col => {{
    const vals = students.map(s => s.napsis[CUR_YEAR]?.asignaturas?.[col.key])
      .filter(v => v !== null && v !== undefined);
    const prom = vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
    const reprob = vals.filter(v=>v<4).length;
    return {{ ...col, prom: prom?+prom.toFixed(2):null, reprob, n: vals.length }};
  }});

  // Riesgo
  const riskGroups = {{ alerta:[], activo:[], preventivo:[] }};
  students.forEach(s => {{
    const lv = s.napsis[CUR_YEAR]?.risk;
    if(lv) riskGroups[lv].push(s);
  }});

  const html = `
  <div class="kpi-row">
    <div class="kpi"><div class="value ca">${{curso.vigentes}}</div><div class="desc">Vigentes</div></div>
    <div class="kpi"><div class="value ${{nc(curso.pf_prom)}}">${{ns(curso.pf_prom)}}</div><div class="desc">Promedio curso</div></div>
    <div class="kpi"><div class="value cr">${{curso.alerta}}</div><div class="desc">Alerta</div></div>
    <div class="kpi"><div class="value cy">${{curso.activo}}</div><div class="desc">Riesgo activo</div></div>
    <div class="kpi"><div class="value ca">${{curso.preventivo}}</div><div class="desc">Seg. preventivo</div></div>
  </div>

  <div class="card mb24">
    <div class="card-title">Estudiantes · ${{curso.grade_name}} · ${{CUR_YEAR}}</div>
    <div class="student-search" style="margin-bottom:12px">
      <input type="text" class="search-input" id="curso-search"
        placeholder="Filtrar…" oninput="filterCursoTable(this.value)">
    </div>
    <div style="overflow-x:auto">
      <table class="std-table" id="curso-std-table">
        <thead><tr>
          <th onclick="sortCurso('nombre',this)">Estudiante</th>
          ${{active_cols.map(c=>`<th title="${{c.label}}" onclick="sortCurso('${{c.key}}',this)">${{c.short}}</th>`).join('')}}
          <th onclick="sortCurso('pf',this)">PROM</th>
          <th>Estado</th>
        </tr></thead>
        <tbody id="curso-tbody"></tbody>
      </table>
    </div>
    <div class="table-note">Clic en encabezado para ordenar · — = sin nota registrada</div>
  </div>

  <div class="card mb24">
    <div class="card-title">Seguimiento · ${{riskGroups.alerta.length + riskGroups.activo.length + riskGroups.preventivo.length}} estudiantes</div>
    ${{['alerta','activo','preventivo'].map(lv => `
      <div class="risk-section">
        <div class="risk-section-label">
          <div class="risk-dot" style="background:${{RISK_COLOR[lv]}}"></div>
          <span style="color:${{RISK_COLOR[lv]}}">${{RISK_LABEL[lv]}}</span>
          <span style="color:var(--muted)">· ${{riskGroups[lv].length}} estudiantes</span>
        </div>
        <div class="risk-grid">
          ${{riskGroups[lv].length === 0
            ? '<div class="empty-risk">Ningún estudiante en esta categoría.</div>'
            : riskGroups[lv].map(s => buildRiskCard(s, lv, CUR_YEAR, active_cols)).join('')
          }}
        </div>
      </div>`).join('')}}
  </div>

  ${{buildDiaCurso(curso)}}
  `;

  document.getElementById('curso-content').innerHTML = html;

  // Render tabla de estudiantes
  window._cursoStudents = students;
  window._cursoSort = {{ key:'nombre', dir:1 }};
  renderCursoTable(students, active_cols);
}}

function buildRiskCard(s, lv, year, active_cols) {{
  const nd = s.napsis[year];
  if(!nd) return '';
  const reds  = active_cols.filter(c=>!c.np&&nd.asignaturas[c.key]!==null&&nd.asignaturas[c.key]!==undefined&&nd.asignaturas[c.key]<4).map(c=>c.short);
  const warns = active_cols.filter(c=>!c.np&&nd.asignaturas[c.key]!==null&&nd.asignaturas[c.key]!==undefined&&nd.asignaturas[c.key]>=4&&nd.asignaturas[c.key]<5).map(c=>c.short);
  const tags = [...reds.map(l=>`<span class="rtag rtag-red">${{l}}</span>`),
                ...warns.map(l=>`<span class="rtag rtag-yellow">${{l}}</span>`)].join('');
  return `<div class="risk-card ${{lv}}">
    <div>
      <div class="risk-name" style="cursor:pointer" onclick="openStudent('${{s.key}}')">${{s.nombre}}</div>
      <div class="risk-detail">Gen. ${{s.gen}}</div>
      <div class="risk-tags">${{tags||'<span style="font-size:11px;color:var(--muted);font-style:italic">Sin asignaturas críticas</span>'}}</div>
    </div>
    <div class="risk-prom" style="color:${{RISK_COLOR[lv]}}">${{ns(nd.pf)}}</div>
  </div>`;
}}

function buildDiaCurso(curso) {{
  const hasDia = _DIA_PERIODO_KEYS.some(p =>
    curso.dia?.[p]?.lectura || curso.dia?.[p]?.matematica
  );
  if(!hasDia) return '';

  const barHtml = (ejesObj, title) => {{
    if(!ejesObj || !Object.keys(ejesObj).length)
      return `<div class="card"><div class="card-title">${{title}}</div><div class="no-data">Sin datos</div></div>`;
    const bars = Object.entries(ejesObj).map(([eje,val]) => {{
      const col = diaColor(val);
      return `<div class="dia-bar-row">
        <span class="dia-bar-label">${{eje}}</span>
        <div class="dia-bar-outer"><div class="dia-bar-fill" style="width:${{val}}%;background:${{col}}"></div></div>
        <span class="dia-val" style="color:${{col}}">${{val}}%</span>
      </div>`;
    }}).join('');
    return `<div class="card"><div class="card-title">${{title}}</div><div class="dia-bar-wrap">${{bars}}</div></div>`;
  }};

  return _DIA_PERIODO_KEYS
    .filter(p => curso.dia?.[p]?.lectura || curso.dia?.[p]?.matematica)
    .map(p => `<div class="g2 mb24">
      ${{barHtml(curso.dia[p]?.lectura,  `DIA ${{_DIA_PERIODO_LABELS[p]}} · Lectura · promedio por eje`)}}
      ${{barHtml(curso.dia[p]?.matematica, `DIA ${{_DIA_PERIODO_LABELS[p]}} · Matemática · promedio por eje`)}}
    </div>`)
    .join('');
}}

let _cursoSort = {{ key:'nombre', dir:1 }};
function sortCurso(key, th) {{
  if(_cursoSort.key===key) _cursoSort.dir*=-1;
  else _cursoSort = {{key,dir:key==='nombre'?1:-1}};
  const active_cols = window._cursoStudents[0]?.napsis[CUR_YEAR]?.col_map?.filter(c=>!c.np)??[];
  renderCursoTable(window._cursoStudents, active_cols);
}}
function filterCursoTable(q) {{
  const active_cols = window._cursoStudents[0]?.napsis[CUR_YEAR]?.col_map?.filter(c=>!c.np)??[];
  const filtered = window._cursoStudents.filter(s=>s.nombre.toLowerCase().includes(q.toLowerCase()));
  renderCursoTable(filtered, active_cols);
}}
function renderCursoTable(students, active_cols) {{
  const sorted = [...students].sort((a,b) => {{
    const k = _cursoSort.key;
    const av = k==='nombre' ? a.nombre : (a.napsis[CUR_YEAR]?.asignaturas?.[k] ?? a.napsis[CUR_YEAR]?.pf ?? 99);
    const bv = k==='nombre' ? b.nombre : (b.napsis[CUR_YEAR]?.asignaturas?.[k] ?? b.napsis[CUR_YEAR]?.pf ?? 99);
    if(av===null||av===99) return 1; if(bv===null||bv===99) return -1;
    return typeof av==='string' ? av.localeCompare(bv)*_cursoSort.dir : (av-bv)*_cursoSort.dir;
  }});
  const tbody = document.getElementById('curso-tbody');
  if(!tbody) return;
  tbody.innerHTML = sorted.map(s => {{
    const nd = s.napsis[CUR_YEAR];
    const cells = active_cols.map(c => {{
      const v = nd?.asignaturas?.[c.key];
      return `<td class="${{nc(v??null)}}">${{ns(v??null)}}</td>`;
    }}).join('');
    const lv = nd?.risk;
    const dot = lv ? `<span class="risk-dot-sm" style="background:${{RISK_COLOR[lv]}}"></span>` : '';
    return `<tr onclick="openStudent('${{s.key}}')" style="cursor:pointer">
      <td title="${{s.nombre}}">${{s.nombre}}</td>
      ${{cells}}
      <td class="pf-cell ${{nc(nd?.pf??null)}}">${{ns(nd?.pf??null)}}</td>
      <td>${{dot}}${{lv?RISK_LABEL[lv]:''}}</td>
    </tr>`;
  }}).join('');
}}

// ════════════════════════════════════════════════════════════
// PESTAÑA ESTUDIANTES
// ════════════════════════════════════════════════════════════
let _estCharts = [];
function searchStudent(q) {{
  const count = document.getElementById('est-count');
  const results = document.getElementById('est-results');
  document.getElementById('est-profile').innerHTML = '';
  if(q.length < 2) {{
    results.innerHTML = '';
    count.textContent = '';
    return;
  }}
  const matches = EST.filter(s => s.nombre.toLowerCase().includes(q.toLowerCase())).slice(0,20);
  count.textContent = matches.length ? `${{matches.length}} resultado${{matches.length>1?'s':''}}` : 'Sin resultados';
  results.innerHTML = matches.map(s => {{
    const nd = s.napsis[CUR_YEAR];
    const grade = nd?.grade_name ?? Object.values(s.napsis).slice(-1)[0]?.grade_name ?? '—';
    const pfStr = nd ? ns(nd.pf) : '—';
    const lv = nd?.risk;
    const dot = lv ? `<span class="risk-dot-sm" style="background:${{RISK_COLOR[lv]}}"></span>` : '';
    return `<div onclick="openStudent('${{s.key}}')" style="cursor:pointer;padding:10px 16px;
      border:1px solid var(--border);border-radius:6px;margin-bottom:6px;
      background:var(--card);display:flex;justify-content:space-between;align-items:center">
      <div>
        <div style="font-weight:600;font-size:13px">${{s.nombre}}</div>
        <div style="font-size:11px;color:var(--muted);font-family:var(--mono)">
          ${{grade}} · Gen. ${{s.gen}} ${{dot}}${{lv?RISK_LABEL[lv]:''}}
        </div>
      </div>
      <div style="font-family:var(--mono);font-size:18px;font-weight:700" class="${{nc(nd?.pf??null)}}">${{pfStr}}</div>
    </div>`;
  }}).join('');
}}

function openStudent(key) {{
  _estCharts.forEach(ch=>ch.destroy());
  _estCharts = [];

  const s = EST.find(e=>e.key===key);
  if(!s) return;

  // Ir a la pestaña estudiantes si no estamos ahí
  showTab('estudiantes', document.querySelector('nav button:nth-child(3)'));
  document.getElementById('est-search').value = s.nombre.split(',')[0];
  document.getElementById('est-results').innerHTML = '';
  document.getElementById('est-count').textContent = '';

  const profile = document.getElementById('est-profile');

  const napYears = Object.keys(s.napsis).map(Number).sort();
  const nd_curr  = s.napsis[CUR_YEAR];

  // ── Cabecera del perfil ───────────────────────────────────
  const grade_curr = nd_curr?.grade_name ?? Object.values(s.napsis).slice(-1)[0]?.grade_name ?? '—';
  const lv = nd_curr?.risk;

  let html = `
  <div class="card mb24">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px">
      <div>
        <div style="font-size:18px;font-weight:700;margin-bottom:4px">${{s.nombre}}</div>
        <div style="font-size:12px;color:var(--muted);font-family:var(--mono)">
          Generación ${{s.gen}} &nbsp;·&nbsp; ${{grade_curr}} &nbsp;·&nbsp; ${{CUR_YEAR}}
          ${{lv ? `&nbsp;·&nbsp;<span style="color:${{RISK_COLOR[lv]}};font-weight:600">${{RISK_LABEL[lv]}}</span>` : ''}}
        </div>
      </div>
      ${{nd_curr ? `<div style="font-family:var(--mono);font-size:36px;font-weight:700" class="${{nc(nd_curr.pf)}}">${{ns(nd_curr.pf)}}</div>` : ''}}
    </div>
  </div>`;

  // ── Trayectoria NAPSIS ────────────────────────────────────
  html += `<div class="card mb24">
    <div class="card-title">Trayectoria NAPSIS · Promedio final por año</div>
    <div style="position:relative;height:200px"><canvas id="chart-est-napsis"></canvas></div>
    <div style="overflow-x:auto;margin-top:16px">
      <table class="years-table">
        <thead><tr>
          <th>Asignatura</th>
          ${{napYears.map(y=>`<th>${{y}}<br><span style="font-weight:400;color:var(--muted)">${{s.napsis[y].grade_name}}</span></th>`).join('')}}
        </tr></thead>
        <tbody id="est-asig-tbody"></tbody>
      </table>
    </div>
  </div>`;

  // ── DIA ───────────────────────────────────────────────────
  const diaYears = Object.keys(s.dia).map(Number).sort();
  const hasDiaData = diaYears.some(yr =>
    _DIA_PERIODO_KEYS.some(p =>
      ['lectura','matematica'].some(m => s.dia[yr]?.[p]?.[m])
    )
  );

  // Puntos de trayectoria separados por materia (una entrada por eval con datos)
  const _diaAvg = obj => obj && Object.keys(obj).length
    ? +(Object.values(obj).reduce((a,b)=>a+b,0)/Object.values(obj).length).toFixed(1)
    : null;
  window._diaLecPts = [];
  window._diaMatPts = [];
  diaYears.forEach(yr => {{
    _DIA_PERIODO_KEYS.forEach(pkey => {{
      const lec = s.dia[yr]?.[pkey]?.lectura;
      const mat = s.dia[yr]?.[pkey]?.matematica;
      const lbl = `${{_DIA_SHORT[pkey]}} ${{yr}}`;
      if(lec && Object.keys(lec).length) window._diaLecPts.push({{label:lbl, value:_diaAvg(lec)}});
      if(mat && Object.keys(mat).length) window._diaMatPts.push({{label:lbl, value:_diaAvg(mat)}});
    }});
  }});

  if(hasDiaData) {{
    // ── Vista detalle: siempre lectura izq / matemática der ──
    // Si falta una materia en un período, se inserta <div></div> vacío
    // para que la grilla de 2 columnas no mezcle materias.
    let detHtml = '<div class="g2">';
    diaYears.forEach(yr => {{
      _DIA_PERIODO_KEYS.forEach(pkey => {{
        const lecEjes = s.dia[yr]?.[pkey]?.lectura;
        const matEjes = s.dia[yr]?.[pkey]?.matematica;
        const hasLec  = lecEjes && Object.keys(lecEjes).length;
        const hasMat  = matEjes && Object.keys(matEjes).length;
        if(!hasLec && !hasMat) return;
        ['lectura','matematica'].forEach(matNm => {{
          const ejes = matNm === 'lectura' ? lecEjes : matEjes;
          if(!ejes || !Object.keys(ejes).length) {{
            detHtml += '<div></div>';   // celda vacía — mantiene posición en grilla
            return;
          }}
          detHtml += `<div>
            <div style="font-size:11px;font-family:var(--mono);color:var(--muted);
              margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">
              ${{matNm==='lectura'?'Lectura':'Matemática'}} · ${{_DIA_PERIODO_LABELS[pkey]}} · ${{yr}}
            </div>
            <div class="dia-bar-wrap">
              ${{Object.entries(ejes).map(([eje,val])=>{{
                const col=diaColor(val);
                return `<div class="dia-bar-row">
                  <span class="dia-bar-label">${{eje}}</span>
                  <div class="dia-bar-outer"><div class="dia-bar-fill"
                    style="width:${{val}}%;background:${{col}}"></div></div>
                  <span class="dia-val" style="color:${{col}}">${{val}}%</span>
                </div>`;
              }}).join('')}}
            </div>
          </div>`;
        }});
      }});
    }});
    detHtml += '</div>';

    // Subtítulo de materia (idéntico al estilo de SEPA)
    const _diaSub = t => `<div style="font-size:11px;font-family:var(--mono);color:var(--muted);
      margin-bottom:6px;text-transform:uppercase;letter-spacing:1px">${{t}}</div>`;
    const _noD = '<div class="no-data">Sin datos</div>';

    html += `<div class="card mb24">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
        <span class="card-title" style="margin:0">DIA · Resultados</span>
        <div style="display:flex;gap:4px">
          <button class="dia-toggle active" id="btn-dia-traj"
            onclick="diaToggle('traj')">Trayectoria</button>
          <button class="dia-toggle" id="btn-dia-det"
            onclick="diaToggle('det')">Detalle</button>
        </div>
      </div>
      <div id="dia-traj-view">
        <div class="g2">
          <div>
            ${{_diaSub('Lectura')}}
            ${{window._diaLecPts.length
              ? '<div style="position:relative;height:180px"><canvas id="chart-dia-lec"></canvas></div>'
              : _noD}}
          </div>
          <div>
            ${{_diaSub('Matemática')}}
            ${{window._diaMatPts.length
              ? '<div style="position:relative;height:180px"><canvas id="chart-dia-mat"></canvas></div>'
              : _noD}}
          </div>
        </div>
      </div>
      <div id="dia-det-view" style="display:none">
        ${{detHtml}}
      </div>
    </div>`;
  }}

  // ── SEPA ──────────────────────────────────────────────────
  const hasLeng = Object.keys(s.sepa.lenguaje).length > 0;
  const hasMat  = Object.keys(s.sepa.matematica).length > 0;
  if(hasLeng || hasMat) {{
    html += `<div class="card mb24">
      <div class="card-title">SEPA · Progreso longitudinal (puntaje)</div>
      <div class="g2">
        ${{hasLeng ? '<div><div style="font-size:11px;font-family:var(--mono);color:var(--muted);margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">Lenguaje</div><div style="position:relative;height:180px"><canvas id="chart-sepa-leng"></canvas></div></div>' : ''}}
        ${{hasMat  ? '<div><div style="font-size:11px;font-family:var(--mono);color:var(--muted);margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">Matemática</div><div style="position:relative;height:180px"><canvas id="chart-sepa-mat"></canvas></div></div>' : ''}}
      </div>
    </div>`;
  }}

  profile.innerHTML = html;

  // ── Gráfico NAPSIS ────────────────────────────────────────
  const pfByYear = napYears.map(y => s.napsis[y].pf);
  const ch1 = new Chart(document.getElementById('chart-est-napsis'), {{
    type:'line',
    data:{{
      labels: napYears.map(y=>`${{y}}\\n${{s.napsis[y].grade_name}}`),
      datasets:[{{
        label:'Promedio final', data: pfByYear,
        borderColor:'#1a5276', backgroundColor:'rgba(26,82,118,.12)',
        borderWidth:2, pointRadius:5, pointHoverRadius:7, tension:.3,
      }}]
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{ display:false }} }},
      scales:{{
        x:{{ grid:{{ color:'rgba(184,208,232,.5)' }}, ticks:{{ color:'#5a7a96',font:{{size:10,family:'IBM Plex Mono'}} }} }},
        y:{{ min:1, max:7,
          grid:{{ color:'rgba(184,208,232,.5)' }},
          ticks:{{ color:'#5a7a96',font:{{size:10,family:'IBM Plex Mono'}},stepSize:.5 }} }},
      }},
    }},
  }});
  _estCharts.push(ch1);

  // ── Tabla de asignaturas por año ──────────────────────────
  // Indexar por LABEL (no por key) para unificar la misma asignatura
  // que puede tener nombres distintos en NAPSIS según el año.
  // keyByYear guarda el key real de cada año para recuperar la nota correcta.

  // Lenguaje: si el estudiante tiene registros desde 7° básico (grado ≥ 7),
  // unificamos "LENGUAJE Y COMUNICACIÓN" (básica) con "LENGUA Y LITERATURA" (media)
  // bajo el nombre de media. Si solo llega hasta 6°, conservamos el nombre de básica.
  const maxGrade = Math.max(...napYears.map(y => s.napsis[y]?.grade || 0));
  function lenguajeLabel(lbl) {{
    const l = lbl.toUpperCase();
    if(maxGrade >= 7 && (l.includes('LENGUAJE Y COMUNICAC') || l.includes('LENGUAJE Y COMUNICACI'))) {{
      return 'LENGUA Y LITERATURA';
    }}
    return lbl;
  }}

  const allSubjs = new Map(); // label → {{label, short, np, keyByYear}}
  napYears.forEach(y => {{
    (s.napsis[y].col_map||[]).forEach(c => {{
      const lbl = lenguajeLabel(c.label);
      if(!allSubjs.has(lbl)) {{
        allSubjs.set(lbl, {{label:lbl, short:c.short, np:c.np, keyByYear:{{}}}});
      }}
      allSubjs.get(lbl).keyByYear[y] = c.key;
    }});
  }});
  const tbody = document.getElementById('est-asig-tbody');
  // Fila de promedio primero
  tbody.innerHTML = `<tr style="border-top:2px solid var(--border)">
    <td style="font-weight:700">Promedio final</td>
    ${{napYears.map(y=>{{
      const v=s.napsis[y].pf;
      return `<td class="pf-cell ${{nc(v)}}" style="font-weight:700">${{ns(v)}}</td>`;
    }}).join('')}}
  </tr>`;
  allSubjs.forEach((col) => {{
    if(col.np) return;
    // Omitir si no hay ninguna nota en ningún año (asignatura que no tuvo el estudiante)
    const hasAnyGrade = napYears.some(y => {{
      const k = col.keyByYear[y];
      const v = k ? s.napsis[y]?.asignaturas?.[k] : undefined;
      return v !== null && v !== undefined;
    }});
    if(!hasAnyGrade) return;
    const cells = napYears.map(y => {{
      const k = col.keyByYear[y];
      const v = k ? s.napsis[y]?.asignaturas?.[k] : undefined;
      if(v===null||v===undefined) return `<td style="color:var(--muted)">—</td>`;
      return `<td class="${{nc(v)}}">${{ns(v)}}</td>`;
    }}).join('');
    tbody.innerHTML += `<tr><td>${{col.label}}</td>${{cells}}</tr>`;
  }});

  // ── Gráficos SEPA ─────────────────────────────────────────
  const drawSepa = (canvasId, scores) => {{
    if(!document.getElementById(canvasId)) return;
    const entries = Object.entries(scores).sort((a,b)=>{{
      const ka = a[0].replace('_P','_0').replace('_F','_1');
      const kb = b[0].replace('_P','_0').replace('_F','_1');
      return ka.localeCompare(kb);
    }});
    const labels = entries.map(([k])=>k.replace('_P',' (inicio)').replace('_F',' (fin)'));
    const data   = entries.map(([,v])=>v);
    const ch = new Chart(document.getElementById(canvasId), {{
      type:'line',
      data:{{
        labels,
        datasets:[{{
          label:'Estudiante', data,
          borderColor:'#1a5276', backgroundColor:'rgba(26,82,118,.12)',
          borderWidth:2, pointRadius:4, tension:.3,
        }}]
      }},
      options:{{
        responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales:{{
          x:{{ grid:{{ color:'rgba(184,208,232,.5)' }}, ticks:{{ color:'#5a7a96',font:{{size:9,family:'IBM Plex Mono'}},maxRotation:30 }} }},
          y:{{ min:450, max:850,
            grid:{{ color:'rgba(184,208,232,.5)' }},
            ticks:{{ color:'#5a7a96',font:{{size:10,family:'IBM Plex Mono'}},stepSize:50 }} }},
        }},
      }},
    }});
    _estCharts.push(ch);
  }};
  if(hasLeng) drawSepa('chart-sepa-leng', s.sepa.lenguaje);
  if(hasMat)  drawSepa('chart-sepa-mat',  s.sepa.matematica);

  // ── Gráficos DIA · trayectoria (lectura y matemática separados) ──
  const _diaScales = {{
    x:{{ grid:{{color:'rgba(184,208,232,.5)'}},
         ticks:{{color:'#5a7a96',font:{{size:9,family:'IBM Plex Mono'}},maxRotation:30}} }},
    y:{{ min:0, max:100,
         grid:{{color:'rgba(184,208,232,.5)'}},
         ticks:{{color:'#5a7a96',font:{{size:10,family:'IBM Plex Mono'}},
                 stepSize:20,callback:v=>v+'%'}} }},
  }};
  const _drawDiaLine = (canvasId, pts, color, bgColor) => {{
    const cvs = document.getElementById(canvasId);
    if(!cvs || !pts || !pts.length) return;
    const ch = new Chart(cvs, {{
      type:'line',
      data:{{
        labels: pts.map(p=>p.label),
        datasets:[{{
          data: pts.map(p=>p.value),
          borderColor:color, backgroundColor:bgColor,
          borderWidth:2, pointRadius:4, tension:.3,
        }}]
      }},
      options:{{
        responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales: _diaScales,
      }},
    }});
    _estCharts.push(ch);
  }};
  _drawDiaLine('chart-dia-lec', window._diaLecPts, '#1a5276', 'rgba(26,82,118,.12)');
  _drawDiaLine('chart-dia-mat', window._diaMatPts, '#1a5276', 'rgba(26,82,118,.12)');
}}

// ════════════════════════════════════════════════════════════
// ARRANQUE
// ════════════════════════════════════════════════════════════
renderColegio();
window._colegioDrawn = true;
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    print()
    print('═' * 62)
    print(f'  DASHBOARD INSTITUCIONAL — {SCHOOL_NAME}')
    print('═' * 62)

    print('\n── Leyendo NAPSIS ──')
    napsis_records = read_all_napsis()

    print('\n── Leyendo DIA ──')
    dia_records = read_all_dia()

    print('\n── Leyendo SEPA ──')
    sepa_records = read_all_sepa()

    print('\n── Vinculando perfiles ──')
    profiles = build_profiles(napsis_records, dia_records, sepa_records)
    total_students = sum(len(v) for v in profiles.values())
    total_gens = len(profiles)
    print(f'  → {total_students} perfiles de estudiantes en {total_gens} generaciones')

    print('\n── Construyendo datos ──')
    js_data = build_js_data(napsis_records, dia_records, sepa_records, profiles)
    print(f'  → {len(js_data["cursos"])} cursos en {CURRENT_YEAR}')
    print(f'  → {len(js_data["estudiantes"])} perfiles totales')

    print('\n── Generando HTML ──')
    html = generate_html(js_data)
    OUTPUT_FILE.write_text(html, encoding='utf-8')
    print(f'  → {OUTPUT_FILE.name}  ({len(html)//1024} KB)')

    print()
    print('═' * 62)
    print(f'  ✓  Listo: {OUTPUT_FILE}')
    print('═' * 62)
    print()


if __name__ == '__main__':
    main()
