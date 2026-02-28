import xlwings as xw


# =========================
# Utilidades de normalización
# =========================
def _norm_text(x):
    return str(x).strip().upper() if x is not None else ""


def _norm_trx(x):
    if x is None:
        return None

    if isinstance(x, int):
        return str(x)

    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return str(x).strip()

    s = str(x).strip()
    if s.endswith(".0"):
        base = s[:-2]
        if base.replace(",", "").isdigit():
            return base
    return s


def _last_row(sheet, col_letter: str) -> int:
    last = sheet.cells.last_cell.row
    return sheet.range(f"{col_letter}{last}").end("up").row


def _read_col(sheet, col_letter: str, start_row: int, end_row: int):
    if end_row < start_row:
        return []
    vals = sheet.range(f"{col_letter}{start_row}:{col_letter}{end_row}").value
    return vals if isinstance(vals, list) else [vals]


# =========================
# Proceso principal
# =========================
def main():
    """
    Valida TRX cruzando RGC vs RG y escribe el resultado en VALIDACIÓN.
    Se ejecuta EXCLUSIVAMENTE desde Excel vía RunPython.
    """

    # --- Obtener workbook caller ---
    try:
        wb = xw.Book.caller()
    except Exception:
        raise RuntimeError(
            "El script debe ejecutarse desde Excel mediante RunPython."
        )

    # --- Validar hojas obligatorias ---
    required_sheets = ("RGC", "RG", "VALIDACIÓN")
    for sh in required_sheets:
        if sh not in [s.name for s in wb.sheets]:
            raise RuntimeError(f"No se encontró la hoja obligatoria: {sh}")

    sh_rgc = wb.sheets["RGC"]
    sh_rg = wb.sheets["RG"]
    sh_val = wb.sheets["VALIDACIÓN"]

    # --- Determinar últimas filas ---
    last_rgc_L = _last_row(sh_rgc, "L")
    last_rgc_AA = _last_row(sh_rgc, "AA")
    last_rg_F = _last_row(sh_rg, "F")

    if last_rgc_L < 3 or last_rg_F < 3:
        raise RuntimeError("No se encontró data suficiente para procesar.")

    # --- Lectura de columnas ---
    rgc_trx = _read_col(sh_rgc, "L", 3, last_rgc_L)
    rgc_status = _read_col(sh_rgc, "AA", 3, last_rgc_AA)
    rg_trx = _read_col(sh_rg, "F", 3, last_rg_F)

    # --- Alinear longitudes (RGC es matriz cerrada) ---
    n = min(len(rgc_trx), len(rgc_status))
    rgc_trx = rgc_trx[:n]
    rgc_status = rgc_status[:n]

    # --- TRX válidas desde RGC ---
    valid_trx = set()
    for trx, st in zip(rgc_trx, rgc_status):
        trx_n = _norm_trx(trx)
        if not trx_n:
            continue
        if _norm_text(st) in ("VÁLIDO", "VALIDO"):
            valid_trx.add(trx_n)

    if not valid_trx:
        raise RuntimeError("No se encontraron TRX válidas en RGC.")

    # --- Cruce con RG + deduplicación ---
    out = []
    seen = set()

    for trx in rg_trx:
        trx_n = _norm_trx(trx)
        if not trx_n:
            continue
        if trx_n in valid_trx and trx_n not in seen:
            out.append(trx_n)
            seen.add(trx_n)

    if not out:
        raise RuntimeError("No hubo coincidencias entre RGC y RG.")

    # --- Escritura en VALIDACIÓN ---
    start_cell = "B14"
    col_letter = "".join(c for c in start_cell if c.isalpha())
    row_num = int("".join(c for c in start_cell if c.isdigit()))

    sh_val.range(f"{col_letter}{row_num}:{col_letter}1048576").clear_contents()
    sh_val.range(f"{col_letter}{row_num - 1}").value = "TRX_VALIDOS"

    sh_val.range(start_cell).options(transpose=True).value = out

    wb.save()