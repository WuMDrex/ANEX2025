# -*- coding: utf-8 -*-
"""Verificación temporal (se puede borrar) de la regla 'ATE solo si el código
es numérico y mayor al umbral estricto'."""
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
import main as M

fallos = []


def chk(nombre, obtenido, esperado):
    ok = obtenido == esperado
    print(f"{'PASA' if ok else 'FALLA'}  {nombre}: {obtenido!r} (esperado {esperado!r})")
    if not ok:
        fallos.append(nombre)


print("=== Estado._tipo_documento ===")
casos = {"2500": "ATE", "3000": "ATE", "999999": "ATE", "02000": None,
         "2000": None, "1999": None, "1083": None, "P001": None, "E051": None,
         "": None, " ": None, "2000.0": None}
for cod, forzado in casos.items():
    for dif in (5, 0, -5):
        signo = "OTS" if dif > 0 else "OC" if dif < 0 else "---"
        chk(f"codigo={cod!r} dif={dif}", M.Estado._tipo_documento(None, cod, dif),
            forzado or signo)

print("\n=== _xl_formulas_base_tabla / f_j ===")
formulas = M._xl_formulas_base_tabla("ANE_T", "AV_T", "ROUND", "ROUND", "CANTIDAD VARIADA")
print(f"f_j = {formulas[8]}")
chk("f_j exacta", formulas[8],
    f'=IF(AND(ISNUMBER([@[No.]]),[@[No.]]>{M.UMBRAL_CODIGO_ATE}),"ATE",'
    'IF([@[CANTIDAD VARIADA]]>0,"OTS",'
    'IF([@[CANTIDAD VARIADA]]<0,"OC","---")))')
chk("f_j sin Y/IFERROR", ("Y(" not in formulas[8] and "IFERROR" not in formulas[8]), True)
chk("11 formulas", len(formulas), 11)
chk("f_b intacta", formulas[0].startswith('=_xlfn.XLOOKUP([@[No.]],CODIGOS_SICOA['), True)
chk("f_k intacta", formulas[9], "=[@[CANTIDAD - V]]-[@[CANTIDAD]]")

print("\n=== patrón de la regla vieja ===")
# La app normaliza sin espacios, en mayúsculas y sin el "=" inicial; la
# retroreferencia \1 obliga a que el +0= hable de la MISMA celda.
for muestra, esperado in (("=A7+0=A7", True), ("=AB12+0=AB12", True),
                          ("=A7+0=A8", False),
                          ("=AND(A7+0=A7,A7>2000)", False), ("=A7>2000", False)):
    chk(f"viejo en {muestra}",
        bool(re.fullmatch(r"=?([A-Z]{1,3}\d+)\+0=\1$",
                          muestra.replace(" ", "").upper().lstrip("="))), esperado)

print("\n=== Excel real: formato condicional ===")
wb = app = None
try:
    import xlwings as xw
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.add()
    ws = wb.sheets[0]
    ws.range("A6").value = [["CODIGO COMPONENTE", "CANTIDAD"],
                            [1999, 1], [2000, 1], [2500, 1], [3000, 1]]
    tabla = ws.tables.add(ws.range("A6:B10"), name="T_PRUEBA_ATE")
    rango = tabla.api.ListColumns(1).DataBodyRange
    rango.FormatConditions.Add(Type=2, Formula1="=A7+0=A7").Font.Color = 16777215
    print("reglas antes  :", int(rango.FormatConditions.Count),
          "->", rango.FormatConditions.Item(1).Formula1)
    M._xl_formato_condicional_codigo_numerico(tabla)
    print("reglas después:", int(rango.FormatConditions.Count),
          "->", rango.FormatConditions.Item(1).Formula1)
    chk("una sola regla", int(rango.FormatConditions.Count), 1)
    colores = {}
    for f in range(7, 11):
        celda = ws.range(f"A{f}")
        colores[int(celda.value)] = int(celda.api.DisplayFormat.Font.Color)
        print(f"  A{f} = {celda.value}  Font.Color = {colores[int(celda.value)]}")
    chk("1999 visible", colores[1999] != 16777215, True)
    chk("2000 visible", colores[2000] != 16777215, True)
    chk("2500 oculto", colores[2500], 16777215)
    chk("3000 oculto", colores[3000], 16777215)
    M._xl_formato_condicional_codigo_numerico(tabla)
    chk("sin duplicar al repetir", int(rango.FormatConditions.Count), 1)
except Exception as e:
    print("NO VERIFICADO en Excel:", type(e).__name__, e)
finally:
    try:
        if wb is not None:
            wb.close()
    except Exception:
        pass
    try:
        if app is not None:
            app.quit()
    except Exception:
        pass

print("\nRESULTADO:", "TODO PASA" if not fallos else f"FALLAN {fallos}")
