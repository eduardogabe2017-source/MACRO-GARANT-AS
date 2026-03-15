import xlwings as xw
from collections import defaultdict


# ─────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────

def _norm_trx(x):
    if x is None:
        return None
    if isinstance(x, float):
        if x != x:
            return None
        if x == int(x):
            return str(int(x))
        return str(x).strip()
    if isinstance(x, int):
        return str(x)
    s = str(x).strip()
    if s.endswith(".0"):
        base = s[:-2]
        if base.replace(",", "").isdigit():
            return base
    return s if s else None


def _norm_str(x):
    return str(x).strip().upper() if x is not None else ""


def _last_row(sheet, col_letter: str, start: int = 1) -> int:
    last = sheet.cells.last_cell.row
    return sheet.range(f"{col_letter}{last}").end("up").row


def _read_col(sheet, col_letter: str, start_row: int, end_row: int):
    if end_row < start_row:
        return []
    vals = sheet.range(f"{col_letter}{start_row}:{col_letter}{end_row}").value
    return vals if isinstance(vals, list) else [vals]


def _pad(lst, length):
    lst = lst if isinstance(lst, list) else [lst]
    return lst + [None] * (length - len(lst))


def _build_concatenado(serie, numero):
    if not serie or not numero:
        return ""
    serie_str = _norm_str(serie)
    if not serie_str:
        return ""
    primera_letra = serie_str[0]
    if primera_letra == "F":
        tipo = "01"
    elif primera_letra == "B":
        tipo = "03"
    else:
        tipo = "00"
    if isinstance(numero, float):
        numero = int(numero) if numero == int(numero) else numero
    numero_str = str(numero).strip()
    if numero_str.endswith(".0"):
        numero_str = numero_str[:-2]
    numero_str = numero_str.zfill(8)
    return f"{tipo}-{serie_str}-{numero_str}"


# ─────────────────────────────────────────────
# Proceso principal
# ─────────────────────────────────────────────

def main():
    """
    PASO 1 - Cruce RG (col F) vs RGC (col L) → 'Pendiente' en RGC col AB
    PASO 2 - Armar CONCATENADO EY en RGC col X (TT-SSSS-NNNNNNNN)
    PASO 3 - Cruzar concatenado vs CXC col F:
             Longitud != 16 → Y='SIN', Z='SIN', AA='SIN DETALLE'
             No encontrado  → Y='NU',  Z='NU',  AA='NO UBICADO'
             Encontrado     → Z=moneda, Y=monto soles, AA='VALIDO'
    PASO 4 - Mapeo de saldos en col AC (solo filas AB='Pendiente'):
             Alguna fila NU  → 'OBSERVADO'
             Todas numericas → |suma Y - col O| <= 1.5 → 'APLICAR' sino 'DIFERENCIA'
             Filas SIN       → no reciben valor en AC
    """
    try:
        wb = xw.Book.caller()
    except Exception:
        raise RuntimeError("El script debe ejecutarse desde Excel mediante RunPython.")

    sheet_names = [s.name for s in wb.sheets]
    for required in ("RG", "RGC", "CXC"):
        if required not in sheet_names:
            raise RuntimeError(f"No se encontró la hoja obligatoria: {required}")

    sh_rg  = wb.sheets["RG"]
    sh_rgc = wb.sheets["RGC"]
    sh_cxc = wb.sheets["CXC"]

    # Leer RG col F desde fila 4
    last_rg_F = _last_row(sh_rg, "F", start=4)
    if last_rg_F < 4:
        raise RuntimeError("No se encontró data en RG (col F).")

    rg_trx_set = set()
    for v in _read_col(sh_rg, "F", 4, last_rg_F):
        t = _norm_trx(v)
        if t:
            rg_trx_set.add(t)

    if not rg_trx_set:
        raise RuntimeError("No se encontraron TRX en RG col F.")

    # Leer RGC desde fila 4
    last_rgc = _last_row(sh_rgc, "L", start=4)
    if last_rgc < 4:
        raise RuntimeError("No se encontró data en RGC (col L).")

    n_rows = last_rgc - 4 + 1
    rgc_L = _pad(_read_col(sh_rgc, "L", 4, last_rgc), n_rows)
    rgc_Q = _pad(_read_col(sh_rgc, "Q", 4, last_rgc), n_rows)
    rgc_R = _pad(_read_col(sh_rgc, "R", 4, last_rgc), n_rows)
    rgc_D = _pad(_read_col(sh_rgc, "D", 4, last_rgc), n_rows)
    rgc_O = _pad(_read_col(sh_rgc, "O", 4, last_rgc), n_rows)

    # Limpiar columnas X a AC desde fila 4
    sh_rgc.range(f"X4:AC{last_rgc}").clear_contents()

    # Leer CXC col F, H, I desde fila 3
    last_cxc_F = _last_row(sh_cxc, "F", start=3)
    cxc_dict = {}
    if last_cxc_F >= 3:
        n_cxc = last_cxc_F - 3 + 1
        cxc_F_raw = _pad(_read_col(sh_cxc, "F", 3, last_cxc_F), n_cxc)
        cxc_H_raw = _pad(_read_col(sh_cxc, "H", 3, last_cxc_F), n_cxc)
        cxc_I_raw = _pad(_read_col(sh_cxc, "I", 3, last_cxc_F), n_cxc)
        for ref, mon, monto in zip(cxc_F_raw, cxc_H_raw, cxc_I_raw):
            ref_n = _norm_str(ref)
            if ref_n:
                cxc_dict[ref_n] = (_norm_str(mon), monto)

    # PASOS 1, 2 y 3
    col_Y_results = []
    col_AB_results = []

    for i in range(n_rows):
        row_excel = i + 4

        # PASO 1: cruce RG vs RGC
        trx = _norm_trx(rgc_L[i])
        ab_val = "Pendiente" if (trx and trx in rg_trx_set) else ""
        sh_rgc.range(f"AB{row_excel}").value = ab_val
        col_AB_results.append(ab_val)

        # PASO 2: concatenado EY
        concat = _build_concatenado(rgc_Q[i], rgc_R[i])
        sh_rgc.range(f"X{row_excel}").value = concat

        # PASO 3: cruce con CXC
        if len(concat) != 16:
            sh_rgc.range(f"Y{row_excel}").value = "SIN"
            sh_rgc.range(f"Z{row_excel}").value = "SIN"
            sh_rgc.range(f"AA{row_excel}").value = "SIN DETALLE"
            col_Y_results.append("SIN")
            continue

        cxc_key = concat.upper()
        if cxc_key not in cxc_dict:
            sh_rgc.range(f"Y{row_excel}").value = "NU"
            sh_rgc.range(f"Z{row_excel}").value = "NU"
            sh_rgc.range(f"AA{row_excel}").value = "NO UBICADO"
            col_Y_results.append("NU")
            continue

        moneda, monto_cxc = cxc_dict[cxc_key]
        sh_rgc.range(f"Z{row_excel}").value = moneda
        sh_rgc.range(f"AA{row_excel}").value = "VÁLIDO"

        if monto_cxc is None:
            sh_rgc.range(f"Y{row_excel}").value = "NU"
            col_Y_results.append("NU")
        elif moneda in ("USS", "US$", "USD"):
            try:
                tc_float = float(rgc_D[i]) if rgc_D[i] is not None else None
            except (ValueError, TypeError):
                tc_float = None
            if tc_float:
                monto_soles = round(float(monto_cxc) * tc_float, 2)
                sh_rgc.range(f"Y{row_excel}").value = monto_soles
                col_Y_results.append(monto_soles)
            else:
                sh_rgc.range(f"Y{row_excel}").value = "NU"
                col_Y_results.append("NU")
        else:
            monto_soles = round(float(monto_cxc), 2)
            sh_rgc.range(f"Y{row_excel}").value = monto_soles
            col_Y_results.append(monto_soles)

    # PASO 4: mapeo de saldos en col AC
    trx_groups = defaultdict(list)
    for i in range(n_rows):
        if col_AB_results[i] == "Pendiente":
            trx = _norm_trx(rgc_L[i])
            if trx:
                trx_groups[trx].append(i)

    for trx, indices in trx_groups.items():
        y_vals = [col_Y_results[i] for i in indices]
        has_nu  = any(v == "NU"  for v in y_vals)
        has_sin = any(v == "SIN" for v in y_vals)

        if has_sin and not has_nu:
            continue

        if has_nu:
            ac_val = "OBSERVADO"
        else:
            monto_o = rgc_O[indices[0]]
            try:
                monto_o_float = float(monto_o) if monto_o is not None else None
            except (ValueError, TypeError):
                monto_o_float = None

            if monto_o_float is None:
                ac_val = "OBSERVADO"
            else:
                try:
                    suma_y = sum(float(v) for v in y_vals if isinstance(v, (int, float)))
                except (ValueError, TypeError):
                    ac_val = "OBSERVADO"
                else:
                    ac_val = "APLICAR" if abs(suma_y - monto_o_float) <= 1.5 else "DIFERENCIA"

        for i in indices:
            sh_rgc.range(f"AC{i + 4}").value = ac_val

    wb.save()
