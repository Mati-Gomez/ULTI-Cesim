"""
cesim_parser.py
Modulo de parseo de exports de Cesim. Se importa directo desde app.py:
no requiere correrse por separado ni generar archivos intermedios.
"""
import re
import pandas as pd
import xlrd

COMPANIES = ['LUCKY', 'MECHA', 'PRELA', 'NEXUS', 'ULTI', 'NOSRE', 'NODO']
N_PRACTICE_ROUNDS = 3  # Cesim: 3 rondas de practica + 12 oficiales

# Alias directos para mantener siempre los mismos nombres canónicos de la industria de ULTI
# en todos los años, incluso si en alguna ronda se renombra el equipo en Cesim (ej. ULTI -> CADIZ).
COMPANY_ALIASES = {
    'CADIZ': 'ULTI',
    'CÁDIZ': 'ULTI',
}


def normalizar_equipos(vals):
    """Normaliza una fila de encabezado de equipos del .xls a los 7 nombres canónicos de COMPANIES:
    1) Si los nombres (tras aplicar COMPANY_ALIASES) coinciden con el conjunto COMPANIES, devuelve
       esos nombres normalizados respetando su orden.
    2) Si cambiaron los nombres de la industria en el .xls (ej. traspaso de cuentas a nombres de
       otra comisión), mapea 1:1 por la posición fija de columna de Cesim (cols 1..7 -> COMPANIES).
    """
    limpios = [COMPANY_ALIASES.get(str(v).strip().upper(), str(v).strip().upper()) for v in vals]
    if set(limpios) == set(COMPANIES):
        return limpios
    return list(COMPANIES)


MODULE_MAP = {
    'Cuenta de resultados, miles USD, Global': 'Estados financieros',
    'Hoja de Balance, miles USD, Global': 'Estados financieros',
    'Cuenta de resultados, miles USD, EE.UU.': 'Estados financieros',
    'Hoja de Balance, miles USD, EE.UU.': 'Estados financieros',
    'Flujo de efectivo de casa matriz, miles USD': 'Estados financieros',
    'Cuenta de resultados, miles USD, Asia': 'Estados financieros',
    'Hoja de Balance, miles USD, Asia': 'Estados financieros',
    'Estado de flujo de efectivo, miles USD, Asia': 'Estados financieros',
    'Cuenta de resultados, miles USD, Europa': 'Estados financieros',
    'Hoja de Balance, miles USD, Europa': 'Estados financieros',
    'Estado de flujo de efectivo, miles USD, Europa': 'Estados financieros',
    'Ratios e indicadores financieros clave': 'Ratios',
    'Informe de mercado, global': 'Informes de mercado',
    'Informe de mercado, EE.UU.': 'Informes de mercado',
    'Informe de mercado, Asia': 'Informes de mercado',
    'Informe de mercado, Europa': 'Informes de mercado',
    'Informe de RRHH': 'Informe de RRHH',
    'Informe ESG': 'Sostenibilidad',
    'Informe del proveedor de componentes': 'Informes de producción',
    'Detalles de fabricación': 'Informes de producción',
    'Detalles de logística': 'Informes de producción',
    'Informe de costos': 'Informes de costos',
    'Desglose de margen por tec, miles USD, EE.UU.': 'Informes de costos',
    'Desglose de margen por tec, miles USD, Asia': 'Informes de costos',
    'Desglose de margen por tec, miles USD, Europa': 'Informes de costos',
    'Valuación - Global': 'Valuación',
    'Valuación - EE.UU.': 'Valuación',
    'Valuación - Asia': 'Valuación',
    'Valuación - Europa': 'Valuación',
    'Creación de valor, miles USD': 'Creación de valor',
    'Demanda estimada, miles unidades': 'Informes de mercado',
    'Precio de venta': 'Informes de mercado',
}


def detect_round(title: str):
    title = title.strip()
    m = re.search(r'Ronda de pr[aá]ctica\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f'Práctica {n}', 'Práctica', n, n - N_PRACTICE_ROUNDS
    m = re.search(r'Ronda\s+inicial\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f'Año {n}', 'Oficial', n, n
    m = re.search(r'Ronda\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f'Año {n}', 'Oficial', n, n
    return title, 'Desconocido', None, None


def parse_cesim_xls(path_or_buffer) -> pd.DataFrame:
    """Parsea UN archivo .xls de Cesim (ruta o file-like buffer) a formato tidy."""
    book = xlrd.open_workbook(file_contents=path_or_buffer.read()) if hasattr(path_or_buffer, 'read') \
        else xlrd.open_workbook(path_or_buffer, formatting_info=True)
    if hasattr(path_or_buffer, 'read'):
        path_or_buffer.seek(0)
        book = xlrd.open_workbook(file_contents=path_or_buffer.read(), formatting_info=True)

    sheet = book.sheet_by_index(0)
    xf_list = book.xf_list
    font_list = book.font_list

    title_row = sheet.cell(0, 0).value
    ronda, tipo_ronda, numero, orden = detect_round(title_row)

    raw = []
    for r in range(1, sheet.nrows):
        label = sheet.cell(r, 0).value
        vals = [sheet.cell(r, c).value for c in range(1, 8)]
        if label == '' and all(v == '' for v in vals):
            continue
        xf = xf_list[sheet.cell_xf_index(r, 0)]
        font = font_list[xf.font_index]
        size = int(font.height / 20)
        weight = font.weight
        raw.append(dict(label=label.strip(), size=size, weight=weight, vals=vals))

    records = []
    cur_statement = cur_section = cur_subgroup = None
    active_companies = list(COMPANIES)

    def is_company_header(label, vals):
        return label == '' and all(isinstance(v, str) and v.strip() != '' for v in vals)

    for item in raw:
        label, size, weight, vals = item['label'], item['size'], item['weight'], item['vals']
        if is_company_header(label, vals):
            active_companies = normalizar_equipos(vals)
            continue
        all_empty = all(v == '' for v in vals)

        if all_empty:
            if size == 14:
                cur_statement, cur_section, cur_subgroup = label, None, None
            elif size == 12:
                cur_section, cur_subgroup = label, None
            elif weight >= 700 or font_list[0].weight >= 700:
                cur_subgroup = label
            else:
                for company, v in zip(active_companies, vals):
                    records.append({
                        'Ronda': ronda,
                        'Tipo_Ronda': tipo_ronda,
                        'Ronda_Numero': numero,
                        'Ronda_Orden': orden,
                        'Modulo': MODULE_MAP.get(cur_statement, 'Sin clasificar'),
                        'Estado': cur_statement or '',
                        'Seccion': cur_section or '',
                        'Subgrupo': cur_subgroup or '',
                        'Metrica': label,
                        'Empresa': company,
                        'Valor': v,
                    })
            continue

        for company, v in zip(active_companies, vals):
            records.append({
                'Ronda': ronda,
                'Tipo_Ronda': tipo_ronda,
                'Ronda_Numero': numero,
                'Ronda_Orden': orden,
                'Modulo': MODULE_MAP.get(cur_statement, 'Sin clasificar'),
                'Estado': cur_statement or '',
                'Seccion': cur_section or '',
                'Subgrupo': cur_subgroup or '',
                'Metrica': label,
                'Empresa': company,
                'Valor': v,
            })

    return pd.DataFrame(records)


_PATRON_ESCALA_MILES = re.compile(r'miles (?:USD|RMB|EUR)', re.IGNORECASE)
_EXCEPCIONES_ESCALA_MONETARIA = {'Costos totales mensuales por empleado, USD'}


def _corregir_escala_monetaria(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica el x1000 verificado a toda fila cuyo Estado, Sección o Métrica
    contenga literalmente "miles USD/RMB/EUR", excepto la única excepción confirmada."""
    if df.empty:
        return df
    texto = (df['Estado'].astype(str) + ' ' + df['Seccion'].astype(str) + ' ' + df['Metrica'].astype(str))
    necesita_escala = texto.str.contains(_PATRON_ESCALA_MILES, regex=True, na=False)
    necesita_escala &= ~df['Metrica'].isin(_EXCEPCIONES_ESCALA_MONETARIA)
    valores = pd.to_numeric(df.loc[necesita_escala, 'Valor'], errors='coerce')
    idx_numericos = valores.dropna().index
    df.loc[idx_numericos, 'Valor'] = valores.loc[idx_numericos] * 1000.0
    return df


def build_historico(file_paths) -> pd.DataFrame:
    """Parsea una lista de rutas .xls y devuelve el dataset historico consolidado,
    ordenado cronologicamente, con nombres de equipos normalizados al universo ULTI."""
    por_ronda = {}
    for p in file_paths:
        frame = parse_cesim_xls(p)
        if frame.empty:
            continue
        ronda = frame['Ronda'].iloc[0]
        por_ronda[ronda] = frame

    if not por_ronda:
        return pd.DataFrame()

    df = pd.concat(por_ronda.values(), ignore_index=True)
    df = df.sort_values(['Ronda_Orden', 'Modulo', 'Estado', 'Seccion', 'Metrica', 'Empresa'])
    df = _corregir_escala_monetaria(df)
    return df
