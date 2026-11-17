"""
Processador CEMIG — Seqüência de Leitura Executada -> Matriz_ControleDeSequenciaLeitura.xlsx
Replica fielmente as 3 macros VBA:
  - Etapa 1: ParteUm_Tratamento (Ctrl+Shift+F)
  - Etapa 2: ParteDois_UlMatr (Ctrl+Shift+T)
  - Etapa 3: ParteTrês_RelatórioFinal (Ctrl+Shift+V)

Saída com as 17 colunas exatas:
  INSTALAÇÃO, MEDIDOR, REGISTRO, RZ_UL_LV, MATRÍCULA, COD, DATA (dd/mm/aaaa),
  HORA, ENDEREÇO, Leit. Anterior, Leit. Executada, CONSUMO,
  NOSB´IMPEDIMENTO, DIGITAÇÃO, CONSISTÊNCIA, NOSB´SIMULAÇÃO, FOTOS.
Ordenado por Leiturista (MATRÍCULA) -> DATA -> HORA crescentes.
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

DEFAULT_ALLOWED_LEITURISTAS: list[str] = [
    "Leiturista: 4MW",
    "Leiturista: 4NI",
    "Leiturista: 51O",
    "Leiturista: 56A",
    "Leiturista: 58A",
    "Leiturista: 5GL",
    "Leiturista: 5HL",
    "Leiturista: 5IL",
    "Leiturista: 5JL",
    "Leiturista: 5LL",
    "Leiturista: 5UL",
    "Leiturista: 5ZP",
    "Leiturista: 80J",
    "Leiturista: 80L",
    "Leiturista: 89O",
]

FINAL_HEADERS: list[str] = [
    "INSTALAÇÃO",
    "MEDIDOR",
    "REGISTRO",
    "RZ_UL_LV",
    "MATRÍCULA",
    "COD",
    "DATA",
    "HORA",
    "ENDEREÇO",
    "Leit. Anterior",
    "Leit. Executada",
    "CONSUMO",
    "NOSB´IMPEDIMENTO",
    "DIGITAÇÃO",
    "CONSISTÊNCIA",
    "NOSB´SIMULAÇÃO",
    "FOTOS",
]


class CemigValidationError(ValueError):
    """Erro claro de validação quando a planilha de entrada não segue o layout CEMIG."""


@dataclass
class ProcessingReport:
    raw_rows: int
    raw_cols: int
    after_deduplication_rows: int
    after_filter_rows: int
    final_rows: int
    unique_units_count: int
    warnings: list[str] = field(default_factory=list)


def _norm(text: str) -> str:
    nfkd = unicodedata.normalize("NFD", str(text or ""))
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch)).lower().strip()


def _format_cell_as_text(cell: openpyxl.cell.Cell) -> str:
    val = cell.value
    if val is None:
        return ""
    if isinstance(val, str):
        return val
    fmt = (cell.number_format or "").strip()
    if isinstance(val, (int, float)) and fmt and set(fmt) == {"0"}:
        return str(int(val)).zfill(len(fmt))
    if hasattr(val, "strftime"):
        if "AM/PM" in fmt.upper():
            return val.strftime("%I:%M:%S %p").lstrip("0")
        return f"{val.month}/{val.day}/{val.year}"
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val)


def read_sheet_as_text_grid(
    file_bytes: bytes,
    filename: str = "Matriz.xlsx",
    preferred_sheet: Optional[str] = None,
) -> tuple[str, list[list[str]], Optional[list[list[str]]]]:
    lower = filename.lower()
    if lower.endswith(".xls") and not lower.endswith(".xlsx"):
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="xlrd")
        target = preferred_sheet if preferred_sheet in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=target, header=None, dtype=str).fillna("")
        grid = df.values.tolist()
        p3_grid = None
        for s in xls.sheet_names:
            if _norm(s) == "planilha3" and s != target:
                df_p3 = pd.read_excel(xls, sheet_name=s, header=None, dtype=str).fillna("")
                p3_grid = df_p3.values.tolist()
                break
        return target, grid, p3_grid

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    if not wb.sheetnames:
        raise CemigValidationError("O arquivo Excel não possui abas.")

    target = preferred_sheet if preferred_sheet in wb.sheetnames else wb.sheetnames[0]

    def extract_ws(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[list[str]]:
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        if max_r == 0 or max_c == 0:
            return []

        merged_ranges = list(ws.merged_cells.ranges)
        for m_range in merged_ranges:
            top_left_val = _format_cell_as_text(ws.cell(m_range.min_row, m_range.min_col))
            ws.unmerge_cells(str(m_range))
            ws.cell(m_range.min_row, m_range.min_col).value = top_left_val
            for r in range(m_range.min_row, m_range.max_row + 1):
                for c in range(m_range.min_col, m_range.max_col + 1):
                    if not (r == m_range.min_row and c == m_range.min_col):
                        ws.cell(r, c).value = ""

        rows: list[list[str]] = []
        for r in range(1, max_r + 1):
            rows.append([_format_cell_as_text(ws.cell(r, c)) for c in range(1, max_c + 1)])

        while rows and all(cell == "" for cell in rows[-1]):
            rows.pop()
        return rows

    main_grid = extract_ws(wb[target])
    p3_grid = None
    for s in wb.sheetnames:
        if _norm(s) == "planilha3" and s != target:
            p3_grid = extract_ws(wb[s])
            break

    return target, main_grid, p3_grid


def validate_cemig_layout(grid: list[list[str]]) -> None:
    if len(grid) < 12:
        raise CemigValidationError(
            f"Layout inválido: apenas {len(grid)} linha(s) encontrada(s). Mínimo esperado: 12."
        )
    max_cols = max((len(r) for r in grid), default=0)
    if max_cols < 10:
        raise CemigValidationError(
            f"Layout inválido: apenas {max_cols} coluna(s). Esperado >= 10."
        )

    markers = set()
    for row in grid[:250]:
        for cell in row:
            n = _norm(cell)
            if "cemig" in n:
                markers.add("cemig")
            if "sequencia de leitura" in n:
                markers.add("sequencia")
            if "unidade de leitura" in n:
                markers.add("unidade")
            if "leiturista:" in n:
                markers.add("leiturista")
            if "medidor" in n:
                markers.add("medidor")
            if "leit. ant" in n or "leit. exec" in n:
                markers.add("leituras")

    if len(markers) < 2:
        raise CemigValidationError(
            "A planilha enviada não corresponde ao relatório bruto CEMIG 'Seqüência de Leitura Executada'."
        )


def _ensure_rect(grid: list[list[str]], min_cols: int) -> None:
    width = max(min_cols, max((len(r) for r in grid), default=0))
    for row in grid:
        if len(row) < width:
            row.extend([""] * (width - len(row)))


def _delete_cols(grid: list[list[str]], start_col1: int, count: int = 1) -> None:
    idx = start_col1 - 1
    for row in grid:
        if idx < len(row):
            del row[idx : idx + count]


def _insert_cols(grid: list[list[str]], col1: int, count: int = 1) -> None:
    idx = col1 - 1
    for row in grid:
        while len(row) < idx:
            row.append("")
        for _ in range(count):
            row.insert(idx, "")


def _copy_col(grid: list[list[str]], src_col1: int, dst_col1: int) -> None:
    _ensure_rect(grid, max(src_col1, dst_col1))
    s_idx, d_idx = src_col1 - 1, dst_col1 - 1
    for row in grid:
        row[d_idx] = row[s_idx]


def _delete_cells_shift_up(grid: list[list[str]], col1: int, start_row1: int, count: int) -> None:
    _ensure_rect(grid, col1)
    c_idx = col1 - 1
    r_idx = start_row1 - 1
    n_rows = len(grid)
    for r in range(r_idx, n_rows):
        src_r = r + count
        grid[r][c_idx] = grid[src_r][c_idx] if src_r < n_rows else ""


def _insert_cells_shift_down(grid: list[list[str]], col1: int, start_row1: int, count: int) -> None:
    _ensure_rect(grid, col1)
    width = len(grid[0]) if grid else col1
    for _ in range(count):
        grid.append([""] * width)
    c_idx = col1 - 1
    r_idx = start_row1 - 1
    for r in range(len(grid) - 1, r_idx + count - 1, -1):
        grid[r][c_idx] = grid[r - count][c_idx]
    for r in range(r_idx, min(r_idx + count, len(grid))):
        grid[r][c_idx] = ""


def _detect_prefer_mdy(raw_dates: list[str]) -> bool:
    has_p1_over_12 = False
    has_p2_over_12 = False
    for raw in raw_dates:
        m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", raw.strip())
        if not m:
            continue
        p1, p2 = int(m.group(1)), int(m.group(2))
        if p1 > 12:
            has_p1_over_12 = True
        if p2 > 12:
            has_p2_over_12 = True
    if has_p2_over_12:
        return True
    if has_p1_over_12:
        return False
    return True


def _format_date_ddmmyyyy(raw: str, prefer_mdy: bool) -> str:
    trimmed = raw.strip()
    if not trimmed:
        return ""
    m_iso = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", trimmed)
    if m_iso:
        return f"{int(m_iso.group(3)):02d}/{int(m_iso.group(2)):02d}/{m_iso.group(1)}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", trimmed)
    if not m:
        return trimmed
    p1, p2 = int(m.group(1)), int(m.group(2))
    yyyy = f"20{m.group(3)}" if len(m.group(3)) == 2 else m.group(3).zfill(4)
    if p1 > 12:
        day, month = p1, p2
    elif p2 > 12:
        month, day = p1, p2
    elif prefer_mdy:
        month, day = p1, p2
    else:
        day, month = p1, p2
    return f"{day:02d}/{month:02d}/{yyyy}"


def _date_sort_key(ddmmyyyy: str) -> int:
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", ddmmyyyy.strip())
    if not m:
        return 99999999
    return int(m.group(3)) * 10000 + int(m.group(2)) * 100 + int(m.group(1))


def _time_sort_seconds(raw_time: str) -> int:
    trimmed = raw_time.strip().upper()
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$", trimmed)
    if not m:
        return 999999
    hours = int(m.group(1))
    minutes = int(m.group(2))
    seconds = int(m.group(3)) if m.group(3) else 0
    meridiem = m.group(4)
    if meridiem == "AM" and hours == 12:
        hours = 0
    elif meridiem == "PM" and hours < 12:
        hours += 12
    return hours * 3600 + minutes * 60 + seconds


def parte_um_tratamento(
    grid: list[list[str]],
    allowed_leituristas: list[str],
) -> tuple[list[str], list[list[str]], list[tuple[str, str]], int, int]:
    _ensure_rect(grid, 16)

    _delete_cols(grid, 5, 1)
    _delete_cols(grid, 7, 2)
    _copy_col(grid, 10, 11)
    _delete_cells_shift_up(grid, 10, 3, 2)
    _insert_cols(grid, 6, 1)
    _copy_col(grid, 7, 6)
    _delete_cells_shift_up(grid, 7, 5, 2)
    _insert_cols(grid, 4, 2)
    _copy_col(grid, 1, 4)
    _copy_col(grid, 1, 5)
    _insert_cells_shift_down(grid, 4, 1, 5)
    _insert_cells_shift_down(grid, 5, 2, 4)

    dedup_cols = max(16, max((len(r) for r in grid), default=16))
    _ensure_rect(grid, dedup_cols)
    df = pd.DataFrame([r[:dedup_cols] for r in grid], dtype=str)
    df_dedup = df.drop_duplicates(keep="first").reset_index(drop=True)
    dedup_grid = df_dedup.values.tolist()
    after_dedup = len(dedup_grid)

    allowed_set = {x.strip() for x in allowed_leituristas}

    def is_header_candidate(row: list[str]) -> bool:
        joined = " | ".join(_norm(c) for c in row)
        return any(k in joined for k in ("medidor", "leit. exec", "leit exec", "consumidora"))

    header_row_1based = 11
    matches_after_11 = sum(
        1 for row in dedup_grid[11:] if row[4].strip() in allowed_set
    )
    if matches_after_11 == 0 or (len(dedup_grid) > 10 and not is_header_candidate(dedup_grid[10])):
        for idx, row in enumerate(dedup_grid):
            if is_header_candidate(row):
                header_row_1based = idx + 1
                break

    h_idx = header_row_1based - 1
    header_row = list(dedup_grid[h_idx]) if h_idx < len(dedup_grid) else [""] * dedup_cols
    filtered_rows = [
        list(row)
        for row in dedup_grid[h_idx + 1 :]
        if row[4].strip() in allowed_set and "cemig distribuicao" not in " ".join(_norm(c) for c in row)
    ]

    if not filtered_rows:
        raise CemigValidationError(
            "Nenhuma linha restou após o filtro de leituristas na Etapa 1 (Coluna E)."
        )

    intermediate_de = [(row[3], row[4]) for row in filtered_rows]
    return header_row, filtered_rows, intermediate_de, after_dedup, len(filtered_rows)


def parte_dois_ul_matr(intermediate_de: list[tuple[str, str]]) -> list[tuple[str, str]]:
    unique_pairs: list[tuple[str, str]] = []
    seen_units: set[str] = set()

    for raw_d, raw_e in intermediate_de:
        unidade = (raw_d[19:] if len(raw_d) > 19 else "").strip()
        matricula = (raw_e[11:] if len(raw_e) > 11 else "").strip()
        if unidade and unidade not in seen_units:
            seen_units.add(unidade)
            unique_pairs.append((unidade, matricula))

    return unique_pairs


def parte_tres_relatorio_final(
    header_row: list[str],
    filtered_rows: list[list[str]],
    ul_matr_pairs: list[tuple[str, str]],
    planilha3_grid: Optional[list[list[str]]],
    warnings: list[str],
) -> bytes:
    grid = [list(header_row), *[list(r) for r in filtered_rows]]
    _ensure_rect(grid, max(16, max((len(r) for r in grid), default=16)))

    _insert_cols(grid, 5, 1)
    for r in range(len(grid)):
        if r == 0:
            grid[r][4] = "Unidade_2"
            grid[r][3] = ""
        else:
            raw_d = grid[r][3]
            grid[r][4] = (raw_d[20:] if len(raw_d) > 20 else "").strip()
            grid[r][3] = raw_d[:20]
    _delete_cols(grid, 4, 1)

    _insert_cols(grid, 5, 2)
    grid[0][4] = "Coluna_Aux"
    grid[0][5] = "Matrícula"
    if len(grid) > 1 and grid[1][6].strip().startswith("Leiturista:"):
        grid[0][6] = "Leiturista"

    _insert_cols(grid, 13, 1)
    grid[0][12] = "Consumo"

    def find_col(primary: list[str], secondary: list[str], fallback: int) -> int:
        for group in (primary, secondary):
            for c, h_val in enumerate(grid[0]):
                nh = _norm(h_val)
                if nh and any(cand in nh for cand in group):
                    return c
        return fallback

    src_medidor = find_col(["medidor"], [], 1)
    src_registro = find_col(["t. reg", "t reg", "registro"], [], 2)
    src_cod = find_col(["oc", "cod"], ["sequencia"], 7)
    src_data = find_col(["data"], [], 8)
    src_hora = find_col(["hora"], [], 9)
    src_endereco = find_col(["endereco", "endereço", "apres"], [], 15)
    src_ant = find_col(["leit. ant", "leit ant"], ["coluna6"], 10)
    src_exec = find_col(["leit. exec", "leit exec"], ["coluna7"], 11)
    src_imp = find_col(["imp.", "imp"], ["coluna8"], 13)
    src_dig = find_col(["qtd dig", "qtd. dig", "digitacao"], [], 14)

    p3_ab: dict[str, str] = {}
    p3_de: dict[str, str] = {}
    p3_gh: dict[str, str] = {}
    p3_jk: dict[str, str] = {}

    if planilha3_grid is not None:
        for row in planilha3_grid[:50000]:
            if len(row) >= 2 and row[0].strip() and row[0].strip() not in p3_ab:
                p3_ab[row[0].strip()] = row[1].strip()
            if len(row) >= 5 and row[3].strip() and row[3].strip() not in p3_de:
                p3_de[row[3].strip()] = row[4].strip()
            if len(row) >= 8 and row[6].strip() and row[6].strip() not in p3_gh:
                p3_gh[row[6].strip()] = row[7].strip()
            if len(row) >= 11 and row[9].strip() and row[9].strip() not in p3_jk:
                p3_jk[row[9].strip()] = row[10].strip()
    else:
        for u, m in ul_matr_pairs:
            if u not in p3_ab:
                p3_ab[u] = m
        warnings.append("Tabela completa do tratamento.")

    prefer_mdy = _detect_prefer_mdy([r[src_data] for r in grid[1:] if r[src_data].strip()])

    records = []
    for idx, row in enumerate(grid[1:]):
        inst = row[0].strip()
        medidor = row[src_medidor].strip()
        registro = row[src_registro].strip()
        rz = row[3].strip()
        leit_fallback = row[6].replace("Leiturista:", "").strip()
        matr = p3_ab.get(rz) or leit_fallback
        cod = row[src_cod].strip()
        data = _format_date_ddmmyyyy(row[src_data], prefer_mdy)
        hora = row[src_hora].strip()
        endereco = row[src_endereco].strip()
        ant_raw = row[src_ant].strip()
        exec_raw = row[src_exec].strip()
        imp = row[src_imp].strip() or "Faturada"
        dig = row[src_dig].strip() or "N00"

        records.append(
            {
                "inst": inst,
                "medidor": medidor,
                "registro": registro,
                "rz": rz,
                "matr": matr,
                "cod": cod,
                "data": data,
                "hora": hora,
                "endereco": endereco,
                "ant_raw": ant_raw,
                "exec_raw": exec_raw,
                "imp": imp,
                "dig": dig,
                "sort_key": (matr.lower(), _date_sort_key(data), _time_sort_seconds(hora), idx),
            }
        )

    records.sort(key=lambda x: x["sort_key"])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ControleLeitura"
    ws.freeze_panes = "A2"

    ws_p3 = wb.create_sheet("Planilha3")
    if planilha3_grid is not None:
        for r_idx, row in enumerate(planilha3_grid, start=1):
            for c_idx, val in enumerate(row, start=1):
                if val != "":
                    c = ws_p3.cell(r_idx, c_idx, val)
                    c.number_format = "@"
    else:
        for r_idx, (unid, matr) in enumerate(ul_matr_pairs, start=1):
            c1 = ws_p3.cell(r_idx, 1, unid)
            c2 = ws_p3.cell(r_idx, 2, matr)
            c1.number_format = "@"
            c2.number_format = "@"

    align = Alignment(horizontal="left", vertical="bottom", wrap_text=False)

    for c_idx, h_name in enumerate(FINAL_HEADERS, start=1):
        cell = ws.cell(1, c_idx, h_name)
        cell.alignment = align
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0F766E" if c_idx >= 15 else "0F172A")

    for idx, rec in enumerate(records, start=2):
        ant_val = int(rec["ant_raw"]) if rec["ant_raw"].isdigit() and not (len(rec["ant_raw"]) > 1 and rec["ant_raw"].startswith("0")) else rec["ant_raw"]
        exec_val = int(rec["exec_raw"]) if rec["exec_raw"].isdigit() and not (len(rec["exec_raw"]) > 1 and rec["exec_raw"].startswith("0")) else rec["exec_raw"]

        row_values = [
            rec["inst"],
            rec["medidor"],
            rec["registro"],
            rec["rz"],
            f"=VLOOKUP(D{idx},Planilha3!$A$1:$B$50000,2,0)",
            rec["cod"],
            rec["data"],
            rec["hora"],
            rec["endereco"],
            ant_val,
            exec_val,
            f"=SUM(K{idx}-J{idx})",
            rec["imp"],
            rec["dig"],
            f'=IFERROR(VLOOKUP(A{idx},Planilha3!$D$1:$E$50000,2,0),"Faturada")',
            f'=IFERROR(VLOOKUP(A{idx},Planilha3!$G$1:$H$50000,2,0),"Faturada")',
            f'=IFERROR(VLOOKUP(A{idx},Planilha3!$J$1:$K$50000,2,0),"N-OK")',
        ]
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(idx, c_idx, val)
            cell.alignment = align
            if isinstance(val, str) and ((val.isdigit() and len(val) > 1 and val.startswith("0")) or c_idx == 7):
                cell.number_format = "@"

    last_row = len(records) + 1
    tab = Table(displayName="Tabela1", ref=f"A1:Q{last_row}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    for c_idx in range(1, 18):
        ws.column_dimensions[get_column_letter(c_idx)].width = max(len(FINAL_HEADERS[c_idx - 1]) + 4, 14)

    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()


def process_matriz_bytes(
    matriz_bytes: bytes,
    matriz_filename: str = "Matriz.xlsx",
    planilha3_bytes: Optional[bytes] = None,
    planilha3_filename: str = "Planilha3.xlsx",
    allowed_leituristas: Optional[list[str]] = None,
) -> tuple[bytes, ProcessingReport]:
    warnings: list[str] = []
    allowed = allowed_leituristas or DEFAULT_ALLOWED_LEITURISTAS

    _, raw_grid, embedded_p3 = read_sheet_as_text_grid(matriz_bytes, matriz_filename)
    validate_cemig_layout(raw_grid)

    raw_rows = len(raw_grid)
    raw_cols = max((len(r) for r in raw_grid), default=0)

    header_row, filtered_rows, intermediate_de, after_dedup, after_filter = parte_um_tratamento(
        raw_grid, allowed
    )
    ul_matr_pairs = parte_dois_ul_matr(intermediate_de)

    p3_grid = embedded_p3
    if planilha3_bytes is not None:
        _, p3_grid, _ = read_sheet_as_text_grid(planilha3_bytes, planilha3_filename, "Planilha3")

    out_bytes = parte_tres_relatorio_final(
        header_row, filtered_rows, ul_matr_pairs, p3_grid, warnings
    )

    report = ProcessingReport(
        raw_rows=raw_rows,
        raw_cols=raw_cols,
        after_deduplication_rows=after_dedup,
        after_filter_rows=after_filter,
        final_rows=len(filtered_rows),
        unique_units_count=len(ul_matr_pairs),
        warnings=list(dict.fromkeys(warnings)),
    )
    return out_bytes, report
