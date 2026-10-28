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
    # Bug real verificado (no un supuesto): el RDOS oficial de la Ronda 0 de ULTI titula la hoja
    # "Resultados, Ronda inicial 0" (confirmado contra el .xls real), no "Ronda 0". La regex genérica
    # de abajo NO matchea por la palabra "inicial" interpuesta -> sin este caso especial, Ronda 0
    # quedaba con Tipo_Ronda='Desconocido' y Ronda_Orden=None, y como app.py filtra Tipo_Ronda=='Oficial'
    # y ordena por Ronda_Orden, Ronda 0 desaparecía SILENCIOSAMENTE de todo el tablero (incluidas
    # evoluciones ya entregadas, ej. Capitalización de Mercado). Se resuelve ANTES de la regex genérica.
    m = re.search(r'Ronda\s+inicial\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        # Rebranding a pedido del equipo (ver Informe de ajustes del Tablero): la etiqueta que
        # se MUESTRA cambia de "Ronda N" a "Año N" en TODO el tablero. El valor que cambia es
        # el STRING DE SALIDA de detect_round() -- este es el único lugar donde se genera --, no
        # el nombre de la columna 'Ronda' del DataFrame (sigue llamándose 'Ronda' internamente,
        # es solo un identificador de columna, nunca se le muestra al usuario) ni las regex de
        # ENTRADA de arriba (matchean el título real que exporta CESIM, que sigue diciendo
        # "Ronda..." -- eso es dato de origen de CESIM, no texto de nuestra UI).
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
        # necesitamos formatting_info tambien en modo buffer
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
            active_companies = [v.strip() for v in vals]
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
                # Fila de datos vacia (ej. tecnologia sin produccion/ventas en la ronda):
                # NO es un subgrupo (su fuente es normal, weight 400), pero registramos
                # las celdas vacias para mantener la estructura completa y no ensuciar cur_subgroup.
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


# ----------------------------------------------------------------------------------------------
# Corrección de escala monetaria (x1000 de "miles USD/RMB/EUR") -- CENTRALIZADA acá.
#
# Regla CESIM verificada (no un supuesto): el RDOS rotula "miles USD" (o "miles RMB"/"miles EUR")
# en el nombre del Estado, de la Sección o de la Métrica -- según la hoja -- y el valor de la celda
# es el crudo EN MILES, no el absoluto. Evidencia ya verificada de forma independiente (bottom-up,
# 0,003% de diferencia): precio de venta x volumen vendido x tipo de cambio, sumado en las 3 áreas y
# las 2 tecnologías activas de ULTI (Ronda 2) = USD 73.716.619.198 contra "Ingresos por ventas"
# Global publicado (73.714.297,44 "miles USD") x1000 = USD 73.714.297.440,62. Ver el detalle completo
# (con el historial de idas y vueltas -- Adendas 23 a 27) en gap_analysis.py, docstring de
# `_to_absoluto()`.
#
# Antes de esta corrección, el dataset que devuelve build_historico() traía el valor CRUDO tal cual
# la celda, y cada consumidor tenía que acordarse de aplicar el x1000 por su cuenta -- lo que llevó a
# que la corrección quedara DUPLICADA y DISPERSA (gap_analysis.py: `_to_absoluto()` para tipo="usd" y
# tipo="unidades" de la Comparativa Plan vs. Real, y `flujo_caja_plan_real_global()` a mano para el
# Flujo de Caja; app.py: parches puntuales en Capitalización de Mercado, Beneficio Neto Acumulado y
# Dinámica de Mercado agregados en las últimas dos tandas) y a la vez INCOMPLETA (el resto del
# tablero nativo -- Estado de Resultados, Balance, waterfall de EBITDA, Operaciones, RRHH -- seguía
# mostrando el valor crudo sin corregir, la "escala híbrida" reportada por el equipo). Se centraliza
# acá, en el ÚNICO lugar que arma el dataset real, para que CUALQUIER consumidor (app.py y
# gap_analysis.py por igual) reciba directamente el valor en USD/unidades absolutas -- los parches
# puntuales que compensaban el crudo se retiran en el mismo cambio (ver Informe de esta refactorización
# para el detalle de qué se retiró y dónde).
#
# Excepción verificada (Categoría: dato real, confirmado por orden de magnitud, no un supuesto):
# 'Costos totales mensuales por empleado, USD' (Informe de RRHH) vive DENTRO de la Sección "Desglose
# de los costos, miles USD" -- junto a líneas que sí son agregados (Costos totales, Salarios y costos
# laborales, etc.) -- pero es un costo POR EMPLEADO, ya absoluto (~USD 10.783/mes/empleado en los RDOS
# reales de ULTI; x1000 sería ~USD 10,78 millones/mes/empleado, no plausible). Auditado exhaustivamente
# contra los RDOS reales de ULTI R0-R3 (los 7 equipos): es el ÚNICO caso de una métrica por-unidad/
# por-empleado anidada dentro de un Estado/Sección/Métrica rotulado "miles" -- no hay otros.
_PATRON_ESCALA_MILES = re.compile(r'miles (?:USD|RMB|EUR)', re.IGNORECASE)
_EXCEPCIONES_ESCALA_MONETARIA = {'Costos totales mensuales por empleado, USD'}


def _corregir_escala_monetaria(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica el x1000 verificado (ver nota arriba) a toda fila cuyo Estado, Sección o Métrica
    contenga literalmente "miles USD/RMB/EUR", excepto la única excepción confirmada. No toca
    unidades ("miles unidades", ya se maneja aparte donde corresponde) ni valores no numéricos
    (se dejan tal cual, ej. el placeholder '-' de una Valuación no calculada)."""
    if df.empty:
        return df
    texto = (df['Estado'].astype(str) + ' ' + df['Seccion'].astype(str) + ' ' + df['Metrica'].astype(str))
    necesita_escala = texto.str.contains(_PATRON_ESCALA_MILES, regex=True, na=False)
    necesita_escala &= ~df['Metrica'].isin(_EXCEPCIONES_ESCALA_MONETARIA)
    valores = pd.to_numeric(df.loc[necesita_escala, 'Valor'], errors='coerce')
    idx_numericos = valores.dropna().index
    df.loc[idx_numericos, 'Valor'] = valores.loc[idx_numericos] * 1000.0
    return df
# ----------------------------------------------------------------------------------------------


def build_historico(file_paths) -> pd.DataFrame:
    """Parsea una lista de rutas .xls y devuelve el dataset historico consolidado,
    ordenado cronologicamente. Si dos archivos corresponden a la misma Ronda
    (ej. se resubio corregido), se queda con el ULTIMO archivo de esa ronda
    completo (no mezcla filas fila-por-fila entre ambos, evita perder datos
    legitimos con nombres de metrica repetidos dentro de una misma seccion).

    Aplica también la corrección de escala monetaria (x1000 de "miles USD/RMB/EUR", ver nota arriba
    de _corregir_escala_monetaria) antes de devolver el dataset -- así TODO consumidor (app.py,
    gap_analysis.py) recibe el valor absoluto ya corregido, sin tener que acordarse de escalarlo."""
    por_ronda = {}  # Ronda -> DataFrame (el ultimo archivo visto para esa ronda gana)
    for p in file_paths:
        frame = parse_cesim_xls(p)
        if frame.empty:
            continue
        ronda = frame['Ronda'].iloc[0]
        por_ronda[ronda] = frame  # sobreescribe si la ronda ya existia

    if not por_ronda:
        return pd.DataFrame()

    df = pd.concat(por_ronda.values(), ignore_index=True)
    df = df.sort_values(['Ronda_Orden', 'Modulo', 'Estado', 'Seccion', 'Metrica', 'Empresa'])
    df = _corregir_escala_monetaria(df)
    return df
