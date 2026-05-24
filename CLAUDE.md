# Dashboard Institucional — Liceo Bicentenario María Luisa Bombal

> **Instrucción para Claude:** Este archivo debe mantenerse actualizado a lo largo de las conversaciones. Al iniciar una sesión, léelo completo. Al terminar cambios significativos (nuevas funciones, correcciones de lógica, cambios en la estructura de datos, nuevas unificaciones de asignaturas, nuevos aliases), edita las secciones relevantes y actualiza la fecha en "Estado actual". El usuario trabaja en dos computadores distintos — este archivo es el único contexto compartido entre conversaciones.
>
> **Al iniciar sesión:** verificar que los paquetes requeridos estén instalados ejecutando `pip install -r requirements.txt` (es instantáneo si ya están instalados). Esto garantiza compatibilidad al cambiar de computador.

---

## Estado actual — última actualización: 2026-05-24

### Funcionalidades implementadas
- ✅ Lectura NAPSIS 2021–2026 (HTML disfrazado de .xlsx)
- ✅ Lectura DIA multi-período (Diagnóstico / Monitoreo / Cierre) por año
- ✅ Lectura SEPA PDFs con puntajes longitudinales
- ✅ Vinculación de perfiles entre las tres fuentes (por `norm_key`)
- ✅ Fusión de perfiles duplicados con nombres similares (fuzzy matching)
- ✅ Detección y fusión de repitentes (mismo nombre en generaciones consecutivas)
- ✅ Aliases nombre social / nombre legal (`_NAME_ALIASES`)
- ✅ Normalización de asignaturas (`_LABEL_CANONICAL`): INGLÉS, EFI, FÍSICA, QUÍMICA, MÚSICA
- ✅ Lenguaje vs Lengua y Literatura: unificación por historial individual (JS)
- ✅ Pestaña Colegio: KPIs, tabla resumen, barras DIA por período, trayectoria longitudinal
- ✅ Pestaña Cursos: tabla ordenada por nombre, columnas vacías ocultas, DIA por período, seguimiento de riesgo
- ✅ Pestaña Estudiantes: NAPSIS + DIA (trayectoria/detalle toggle) + SEPA
- ✅ Nombres en mayúsculas con tildes correctas (personas y asignaturas)
- ✅ Sin letra "A" en identificación de cursos (colegio con un curso por nivel)

### Aliases activos
| Nombre legal (SEPA/DIA) | Nombre social (NAPSIS) | Contexto |
|---|---|---|
| Luz Muñoz Sotomayor | Lucas Muñoz Sotomayor | II° Medio, gen 2028 |

### Unificaciones de asignaturas activas
INGLÉS (idioma extranjero inglés), EDUCACIÓN FÍSICA Y SALUD (educación física), FÍSICA (física diferenciado), QUÍMICA (química electivo), MÚSICA (musica sin tilde)

---

## Contexto

Este proyecto genera `dashboard_institucional.html` a partir de tres fuentes de datos académicos. El script principal es `actualizar_dashboard.py`. Se ejecuta con `python actualizar_dashboard.py` desde esta carpeta y produce un único archivo HTML autocontenido (~4.5 MB).

El colegio tiene **un solo curso por nivel**, por lo que nunca se muestra la letra "A" en las identificaciones de curso.

---

## Estructura de carpetas

```
Seguimiento de notas/                   ← carpeta raíz en Google Drive
│                                          (solo contiene acceso directo al dashboard)
└── Seguimiento institucional/          ← carpeta de trabajo principal (aquí se ejecuta todo)
    ├── actualizar_dashboard.py         ← script principal (todo el código aquí)
    ├── dashboard_institucional.html    ← output generado
    ├── CLAUDE.md                       ← este archivo
    ├── requirements.txt                ← paquetes Python requeridos (pip install -r requirements.txt)
    ├── napsis/
    │   ├── 2021/  XX_Notas.xlsx        ← archivos HTML disfrazados de .xlsx (magic bytes <!DOCTYPE)
    │   ├── 2022/  ...
    │   ├── 2025/  ...
    │   ├── 2026/  XX_Notas.xlsx        ← solo si no hay actas (fallback)
    │   └── actas/ a8833_*.TXT/.txt     ← ★ fuente principal año actual (sección Actas NAPSIS)
    ├── DIA/
    │   └── DIA {año}/
    │       ├── DIA Diagnóstico {año}/{curso}/*.xls(x)
    │       ├── DIA Monitoreo {año}/{curso}/*.xls(x)
    │       └── DIA Cierre {año}/{curso}/*.xls(x)
    ├── SEPA/
    │   └── Reportes progreso alumnos {curso}/  *.pdf   (un PDF por estudiante)
    ├── template/
    │   └── dashboard_ejemplo.html      ← fuente del CSS base
    └── proyectos/                      ← herramientas y proyectos complementarios
        │                                  (cada uno se gestiona en su propia sesión de Claude Code)
        ├── descargar_napsis_historico.py  ← descarga masiva de NAPSIS con Playwright
        │                                     (herramienta separada: puebla napsis/{año}/ automaticamente)
        │                                     NO está integrada en actualizar_dashboard.py
        ├── dashboard/                  ← proyecto de dashboard alternativo/reportes PDF
        │   ├── generate_all.py
        │   ├── generate_unified.py
        │   ├── run_pipeline.py
        │   ├── export_pdf.py
        │   ├── descargar_notas_napsis.py
        │   ├── input/, output/, template/, pdf/
        │   └── Reporte_Notas_2026.pdf
        ├── actualizar_dashboard_prueba.py  ← wrapper de prueba (ver sección Flujo prueba→producción)
        ├── dashboard_institucional_prueba.html  ← output del script de prueba (generado)
        └── 05. mayo/                   ← proyectos de reuniones por fecha
            └── 01. mayo 11/ ... 07. mayo 20/
```

---

## Fuentes de datos

### NAPSIS (notas finales por asignatura)
- **Año actual**: se leen desde `napsis/actas/` (archivos TXT de la sección Actas de NAPSIS). Ver flujo de actualización abajo.
- **Años históricos (2021–2025)**: archivos `.xlsx` que son en realidad HTML (`<!DOCTYPE`), leídos con BeautifulSoup + `encoding='latin-1'`
- Nombres con encoding roto: se corrigen con `fix_enc()` → `fix_accents()` → `.upper()`
- La detección es automática: si `napsis/actas/a*_4_*.txt` existe → se usan las actas; si no → se usan los HTML del año actual

### Flujo de actualización rápida (Actas NAPSIS)
En lugar de descargar 12 archivos HTML desde la sección de notas:
1. En NAPSIS → sección **Actas** → descargar los archivos del establecimiento
2. Reemplazar los archivos en `napsis/actas/` (sobreescribir los TXT anteriores)
3. Ejecutar `python actualizar_dashboard.py`

Archivos que usa el lector de actas:
| Archivo | Contenido |
|---|---|
| `a8833_1.TXT` | Nómina: RUN → nombre completo |
| `a8833_4_*.txt` | Notas por alumno y asignatura |
| `a8833_5_*.txt` | Promedio final y asistencia oficial |
| `a8833_7.TXT` | Matrícula final y profesor jefe por curso |
| `a8833_8.TXT` | Catálogo código → nombre de asignatura |

### DIA (evaluación diagnóstica)
- Archivos Excel reales: `.xls` (xlrd) o `.xlsx` (openpyxl), detectados por magic bytes
- Tres períodos por año en orden: **Diagnóstico → Monitoreo → Cierre**
- Estructura de datos del perfil: `dia[año][periodo][materia]` donde `periodo` ∈ `{'diagnostico','monitoreo','cierre'}` y `materia` ∈ `{'lectura','matematica'}`
- Los ejes de matemática difieren entre períodos (OA en Diagnóstico; Números/Álgebra/Geometría/Estadística en Monitoreo y Cierre)
- Rutas largas de Google Drive: se leen con prefijo `\\?\` via `_read_file_bytes()` + `io.BytesIO`

### SEPA (progreso individual longitudinal)
- Un PDF por estudiante, puntajes en escala 600–800
- Columnas clave: `2022_P` (inicio), `2022_F` (fin), `2023_F`, etc.
- Orden cronológico en gráfico: `_P → _0`, `_F → _1` para el sort

---

## Fórmula de generación (gen)

```python
gen = año + (12 - grado_numérico)
```

Estable por cohorte: el mismo estudiante tiene el mismo `gen` en cualquier año.
Grados: 1–8 = Básico, 9 = I° Medio, 10 = II° Medio, 11 = III° Medio, 12 = IV° Medio.

---

## Vinculación de perfiles entre fuentes

### Normalización de nombre → clave
```python
norm_key(nombre) → tokens ordenados alfabéticamente, sin tildes, minúsculas, unidos por _
# Ej: "PARGA GONZÁLEZ, OLIVIA" → "gonzalez_olivia_parga"
```

### Fusión de duplicados (`merge_duplicate_profiles`)
- `difflib.SequenceMatcher` por token, umbral ≥ 0.82
- Requiere ≥ 2 tokens exactos compartidos

### Detección de repitentes (`merge_repeaters`)
- Mismo nombre en generaciones consecutivas (gen G y G+1)
- Mismo grado en dos años seguidos
- Se fusiona en la generación mayor (año real de egreso)

### Aliases nombre social / nombre legal (`_NAME_ALIASES`)
Para estudiantes cuyo nombre social (aceptado en NAPSIS) difiere del nombre legal (en SEPA/DIA):

```python
_NAME_ALIASES = {
    'luz_munoz_sotomayor': 'lucas_munoz_sotomayor',
    # Agregar más casos con el mismo formato:
    # norm_key(nombre_legal): norm_key(nombre_social_napsis)
}
```

---

## Normalización de asignaturas

### `_LABEL_CANONICAL` — unifica variantes al mismo nombre completo
```python
('idioma extranjero', 'INGLÉS')
('idioma',            'INGLÉS')
('educación física',  'EDUCACIÓN FÍSICA Y SALUD')   # ANTES que 'física'
('educacion fisica',  'EDUCACIÓN FÍSICA Y SALUD')
('física',            'FÍSICA')                      # unifica con diferenciado de física
('fisica',            'FÍSICA')
('química',           'QUÍMICA')                     # unifica con química electivo
('quimica',           'QUÍMICA')
('música',            'MÚSICA')                      # unifica con y sin tilde
('musica',            'MÚSICA')
```
Para agregar una nueva unificación: añadir par `(subcadena_lowercase, 'NOMBRE CANÓNICO')`.
**Importante:** colocar variantes más específicas ANTES que las generales (ej. 'educación física' antes que 'física').

### `_SHORT_MAP` — abreviaciones para cabeceras de la pestaña Cursos
EFI = Educación Física y Salud, MAT. = Matemática, etc.

### Lenguaje vs Lengua y Literatura
Se unifica **en JS**, no en Python, porque depende del historial individual:
- Si el estudiante tiene algún registro desde 7° básico (grado ≥ 7) → "LENGUA Y LITERATURA"
- Si solo llega hasta 6° básico → "LENGUAJE Y COMUNICACIÓN"

---

## Estructura del HTML generado

Tres pestañas:

### Pestaña Colegio
- KPIs generales (vigentes, alertas, promedio)
- Tabla resumen por curso con columnas: Curso, Generación, Profesor/a acompañante, Vigentes, Alerta, Riesgo, Prev., Prom., DIA Lect., DIA Mat.
- DIA 2026 — barras por eje por curso, agrupadas por período
- Gráfico de trayectoria longitudinal por generación

### Pestaña Cursos
- Selector de curso por generación
- Tabla de estudiantes ordenada por nombre por defecto (clic en columna para reordenar)
- Solo muestra columnas donde al menos un estudiante tiene nota
- Sección de seguimiento por nivel de riesgo (alerta / riesgo activo / seg. preventivo)
- Barras DIA del año actual, separadas por período (Diagnóstico → Monitoreo → Cierre)

### Pestaña Estudiantes
- Buscador por nombre
- Trayectoria NAPSIS (gráfico de línea + tabla de asignaturas por año)
  - Solo muestra asignaturas con al menos una nota en algún año
- DIA — tarjeta con dos vistas alternables:
  - **Trayectoria**: dos gráficos separados (Lectura y Matemática), promedio de ejes por evaluación
  - **Detalle**: barras por eje; lectura siempre izquierda, matemática siempre derecha; celda vacía si falta una materia
- SEPA — dos gráficos separados (Lenguaje y Matemática), escala 450–850

---

## Niveles de riesgo académico (`calc_risk`)
- **Alerta**: promedio final < 4.5 o ≥ 2 asignaturas reprobadas
- **Riesgo activo**: promedio 4.5–5.0 o 1 asignatura reprobada
- **Seg. preventivo**: promedio 5.0–5.5

---

## Convenciones de código

- Todo el HTML/JS se genera como f-string en `generate_html()`. Las llaves literales JS usan `{{` y `}}`.
- `${{variable_js}}` en Python → `${variable_js}` en el HTML resultante.
- Los nombres de personas siempre en MAYÚSCULAS con tildes agudas (Í no Ì).
- Profesores/as: se aplica `clean_nombre()` al extraer de NAPSIS.
- Rutas Windows con Unicode largo: usar `_read_file_bytes()` que aplica prefijo `\\?\`.

---

## Flujo prueba → producción (nuevas herramientas)

Las integraciones nuevas se desarrollan en `proyectos/actualizar_dashboard_prueba.py`, que es un wrapper que importa `actualizar_dashboard.py` y solo cambia el archivo de salida a `proyectos/dashboard_institucional_prueba.html`.

**Para probar:**
```
python proyectos/actualizar_dashboard_prueba.py
```

**Para promover a producción:** cuando la herramienta funciona correctamente, decirle a Claude:
> "ya estamos listos, incorporar al script principal"

Claude trasladará los cambios de `actualizar_dashboard_prueba.py` → `actualizar_dashboard.py`.
No copiar manualmente entre archivos.

---

## Cómo agregar una nueva unificación de asignaturas

1. Abrir `actualizar_dashboard.py`
2. Localizar `_LABEL_CANONICAL` (~línea 150)
3. Agregar `('subcadena_en_minúsculas', 'NOMBRE CANÓNICO EN MAYÚSCULAS')`
4. Si hay versión sin tilde, agregar también esa variante
5. Ejecutar `python actualizar_dashboard.py`

## Cómo agregar un alias de nombre social

1. Localizar `_NAME_ALIASES` (~línea 95)
2. Agregar `'norm_key_nombre_legal': 'norm_key_nombre_social_napsis'`
   - `norm_key` = tokens ordenados alfabéticamente, sin tildes, en minúsculas, separados por `_`
   - Ejemplo: "Luz Muñoz Sotomayor" → `'luz_munoz_sotomayor'`
3. Ejecutar el script; el log mostrará líneas `[alias] ... → ...` confirmando la redirección
