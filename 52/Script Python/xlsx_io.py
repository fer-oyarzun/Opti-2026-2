# -*- coding: utf-8 -*-
"""Lectura y escritura de planillas .xlsx usando solo la librería estándar (zipfile + xml).

Evita depender de openpyxl/xlsxwriter, que no están en la lista de librerías permitidas
del curso (gurobipy, pandas, numpy, scipy, sympy, matplotlib, seaborn).
"""

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import pandas as pd

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


# =============================================================================
# LECTURA
# =============================================================================

def _col_index(ref):
    """'C12' -> 2 (columna base 0)."""
    letras = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _texto(nodo):
    """Concatena todos los <t> de un nodo (maneja texto enriquecido)."""
    return "".join(t.text or "" for t in nodo.iter(f"{{{_NS['m']}}}t"))


class LibroXlsx:
    """Libro .xlsx de solo lectura. Uso: libro.sheet_names, libro.hoja(nombre) -> DataFrame."""

    def __init__(self, ruta):
        self.ruta = ruta
        with zipfile.ZipFile(ruta) as z:
            nombres = set(z.namelist())
            self._compartidos = []
            if "xl/sharedStrings.xml" in nombres:
                raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
                self._compartidos = [_texto(si) for si in raiz.findall("m:si", _NS)]

            libro = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            destino = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", _NS)}

            self._rutas = {}
            for hoja in libro.find("m:sheets", _NS).findall("m:sheet", _NS):
                objetivo = destino[hoja.get(_RID)].lstrip("/")
                if not objetivo.startswith("xl/"):
                    objetivo = "xl/" + objetivo
                self._rutas[hoja.get("name")] = objetivo
            self._xml = {nombre: z.read(r) for nombre, r in self._rutas.items()}

        self.sheet_names = list(self._rutas)

    def _valor(self, c):
        tipo = c.get("t")
        if tipo == "inlineStr":
            nodo = c.find("m:is", _NS)
            return _texto(nodo) if nodo is not None else None
        v = c.find("m:v", _NS)
        if v is None or v.text is None:
            return None
        if tipo == "s":
            return self._compartidos[int(v.text)]
        if tipo in ("str", "e"):
            return v.text
        if tipo == "b":
            return v.text == "1"
        num = float(v.text)
        return int(num) if num.is_integer() and abs(num) < 2**53 else num

    def filas(self, nombre):
        """Devuelve la hoja como lista de listas (celdas vacías = None)."""
        if nombre not in self._xml:
            raise ValueError(f"La planilla no contiene la hoja '{nombre}'. Hojas: {self.sheet_names}")
        raiz = ET.fromstring(self._xml[nombre])
        datos = raiz.find("m:sheetData", _NS)
        salida = []
        for fila in datos.findall("m:row", _NS):
            idx_fila = int(fila.get("r")) - 1
            while len(salida) < idx_fila:
                salida.append([])
            valores = []
            for j, c in enumerate(fila.findall("m:c", _NS)):
                col = _col_index(c.get("r")) if c.get("r") else j
                while len(valores) < col:
                    valores.append(None)
                valores.append(self._valor(c))
            salida.append(valores)
        return salida

    def hoja(self, nombre):
        """Devuelve la hoja como DataFrame usando la primera fila como encabezado."""
        filas = self.filas(nombre)
        if not filas:
            return pd.DataFrame()
        encabezado = [str(h).strip() if h is not None else f"Unnamed_{i}" for i, h in enumerate(filas[0])]
        ancho = len(encabezado)
        cuerpo = [(f + [None] * ancho)[:ancho] for f in filas[1:] if any(v is not None for v in f)]
        return pd.DataFrame(cuerpo, columns=encabezado)


# =============================================================================
# ESCRITURA
# =============================================================================

def _letra(col):
    s = ""
    col += 1
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _celda(ref, valor, estilo_encabezado=False):
    st = ' s="1"' if estilo_encabezado else ""
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return f'<c r="{ref}" t="b"{st}><v>{int(valor)}</v></c>'
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if isinstance(valor, float) and not math.isfinite(valor):
            return ""
        return f'<c r="{ref}"{st}><v>{repr(float(valor)) if isinstance(valor, float) else valor}</v></c>'
    if hasattr(valor, "item"):  # escalares numpy
        return _celda(ref, valor.item(), estilo_encabezado)
    texto = escape(str(valor))
    return f'<c r="{ref}" t="inlineStr"{st}><is><t xml:space="preserve">{texto}</t></is></c>'


def _hoja_xml(df):
    filas = []
    cols = list(df.columns)
    celdas = "".join(_celda(f"{_letra(j)}1", str(c), True) for j, c in enumerate(cols))
    filas.append(f'<row r="1">{celdas}</row>')
    for i, registro in enumerate(df.itertuples(index=False), start=2):
        celdas = []
        for j, v in enumerate(registro):
            if v is not None and not isinstance(v, str) and pd.isna(v):
                v = None
            celdas.append(_celda(f"{_letra(j)}{i}", v))
        filas.append(f'<row r="{i}">{"".join(celdas)}</row>')
    anchos = "".join(
        f'<col min="{j + 1}" max="{j + 1}" width="{min(60, max(10, len(str(c)) + 2))}" customWidth="1"/>'
        for j, c in enumerate(cols)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_NS["m"]}">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        + (f"<cols>{anchos}</cols>" if cols else "")
        + f'<sheetData>{"".join(filas)}</sheetData></worksheet>'
    )


def escribir_xlsx(ruta, hojas):
    """Escribe un .xlsx con varias hojas. `hojas` es un dict {nombre_hoja: DataFrame}."""
    nombres = []
    for nombre in hojas:
        limpio = re.sub(r"[\[\]\*\?/\\:]", "_", str(nombre))[:31]
        if limpio in nombres:
            raise ValueError(f"Nombre de hoja duplicado tras truncar a 31 caracteres: '{limpio}'")
        nombres.append(limpio)

    tipos = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(nombres) + 1)
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{tipos}</Types>"
    )
    rels_raiz = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS["pr"]}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    hojas_xml = "".join(
        f'<sheet name="{escape(n, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
        for i, n in enumerate(nombres, start=1)
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_NS["m"]}" xmlns:r="{_NS["r"]}"><sheets>{hojas_xml}</sheets></workbook>'
    )
    rels_libro = "".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(nombres) + 1)
    )
    rels_libro = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS["pr"]}">{rels_libro}'
        f'<Relationship Id="rId{len(nombres) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    estilos = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{_NS["m"]}">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_raiz)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", rels_libro)
        z.writestr("xl/styles.xml", estilos)
        for i, df in enumerate(hojas.values(), start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _hoja_xml(df))
