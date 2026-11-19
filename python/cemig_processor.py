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
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any, Optional
from xml.sax.saxutils import escape as xml_escape

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
    "UL",
    "MATRÍCULA",
    "TIPO",
    "COD",
    "DATA",
    "HORA",
    "ENDEREÇO",
    "Leit. Anterior",
    "Leit. Executada",
    "CONSUMO",
    "NOSB´IMPEDIMENTO",
    "DIGITAÇÃO",
    "Apres.",
    "CONSISTÊNCIA",
    "NOSB´SIMULAÇÃO",
    "FOTOS",
]

CONTROLE_HORARIO_HEADERS: list[str] = [
    "Matricula",
    "Data",
    "Hora inicial",
    "Hora Final",
    "Tempo de Serviço",
]

RELATORIO_GERAL_HEADERS: list[str] = [
    "INSTALAÇÃO",
    "MEDIDOR",
    "REG.",
    "RZ UL LV",
    "UL",
    "Foto?",
    "Tipo",
    "MATR",
    "NL",
    "DATA",
    "HORA",
    "ENDEREÇO",
    "ANTERIOR",
    "ATUAL",
    "CONSUMO",
    "IMPEDIMENTO",
    "SIMULAÇÃO",
    "DIG",
    "CNA",
    "CONSISTENCIA",
]


def _is_excluded_extra_col(header_name: str) -> bool:
    norm = _norm(header_name).rstrip(".").strip()
    return norm in ("c", "a exec", "exec", "leiturista")


def normalize_apres_value(raw: str) -> str:
    return "00" if str(raw or "").strip() == "00" else "N00"


def extract_ul_from_rz_ul_lv(rz_ul_lv: str) -> str:
    """
    Regra da coluna UL (a partir de RZ_UL_LV):
    - Extrair do 3º ao 6º caractere contando da esquerda para a direita:
      substring(RZ_UL_LV, 3, 4) em base 1 (índices 2..5 em base 0);
    - Contabilizar zeros à esquerda normalmente (não ignorar zeros iniciais);
    - Se tiver menos de 6 caracteres, pegar do 3º caractere até o final.
    """
    clean = str(rz_ul_lv or "").strip()
    if len(clean) < 3:
        return ""
    return clean[2:6]


def classify_tipo_from_rz_ul_lv(rz_ul_lv: str) -> str:
    """
    Classificação da coluna TIPO pelos 2 últimos dígitos de RZ_UL_LV:
    - 00 a 79 -> "Urbano"
    - 80 a 89 -> "Povoado"
    - 90 a 99 -> "Rural"
    - Vazio ou inválido -> ""
    """
    digits = re.sub(r"\D", "", rz_ul_lv or "")
    if len(digits) < 2:
        return ""
    last_two = int(digits[-2:])
    if last_two <= 79:
        return "Urbano"
    if last_two <= 89:
        return "Povoado"
    return "Rural"


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
    unique_instalacoes_count: int = 0
    rural_count: int = 0
    povoado_count: int = 0
    urbano_count: int = 0
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
            if _norm(s) in ("filtros", "planilha3") and s != target:
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
        if _norm(s) in ("filtros", "planilha3") and s != target:
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
    if re.match(r"^\d{5}(\.\d+)?$", trimmed):
        serial = int(float(trimmed))
        if 25000 <= serial <= 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=serial)
            return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
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


def _format_hora_hhmm(raw: str) -> str:
    trimmed = raw.strip()
    if not trimmed:
        return ""
    if re.match(r"^0?\.\d+$", trimmed):
        frac = float(trimmed)
        if 0 <= frac < 1:
            total_mins = round(frac * 24 * 60) % (24 * 60)
            return f"{total_mins // 60:02d}:{total_mins % 60:02d}"
    return trimmed


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


def _parse_time_to_minutes_hhmm(raw_time: str) -> Optional[int]:
    trimmed = (raw_time or "").strip().upper()
    if not trimmed:
        return None
    if re.match(r"^0?\.\d+$", trimmed):
        frac = float(trimmed)
        if 0 <= frac < 1:
            return round(frac * 24 * 60) % (24 * 60)
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$", trimmed)
    if not m:
        return None
    hours = int(m.group(1))
    minutes = int(m.group(2))
    meridiem = m.group(4)
    if minutes < 0 or minutes > 59:
        return None
    if meridiem == "AM" and hours == 12:
        hours = 0
    elif meridiem == "PM" and hours < 12:
        hours += 12
    if hours < 0 or hours > 23:
        return None
    return hours * 60 + minutes


def _format_minutes_to_hhmm(total_minutes: int) -> str:
    safe = max(0, int(total_minutes))
    return f"{safe // 60:02d}:{safe % 60:02d}"


def build_controle_horario_rows(records: list[dict[str, Any]]) -> list[dict[str, str]]:
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for rec in records:
        matr = str(rec.get("matr") or "").strip() or "Sem Leiturista"
        data = str(rec.get("data") or "").strip()
        key = (matr, data)
        mins = _parse_time_to_minutes_hhmm(str(rec.get("hora") or ""))
        grp = groups.get(key)
        if grp is None:
            groups[key] = {
                "matricula": matr,
                "data": data,
                "date_sort_key": _date_sort_key(data),
                "min_mins": mins,
                "max_mins": mins,
            }
        elif mins is not None:
            if grp["min_mins"] is None or mins < grp["min_mins"]:
                grp["min_mins"] = mins
            if grp["max_mins"] is None or mins > grp["max_mins"]:
                grp["max_mins"] = mins

    sorted_groups = sorted(
        groups.values(),
        key=lambda g: (g["matricula"].lower(), g["date_sort_key"]),
    )
    rows: list[dict[str, str]] = []
    for grp in sorted_groups:
        min_m = grp["min_mins"]
        max_m = grp["max_mins"]
        has_times = min_m is not None and max_m is not None
        hora_ini = _format_minutes_to_hhmm(min_m) if has_times else ""
        hora_fim = _format_minutes_to_hhmm(max_m) if has_times else ""
        diff = (max_m - min_m) if (has_times and max_m > min_m) else 0
        tempo = _format_minutes_to_hhmm(diff)
        rows.append(
            {
                "matricula": grp["matricula"],
                "data": grp["data"],
                "hora_inicial": hora_ini,
                "hora_final": hora_fim,
                "tempo_servico": tempo,
            }
        )
    return rows


def parte_um_tratamento(
    grid: list[list[str]],
    allowed_leituristas: list[str],
) -> tuple[
    list[str],
    list[list[str]],
    list[tuple[str, str]],
    int,
    int,
    list[list[str]],
    list[int],
    int,
]:
    _ensure_rect(grid, 16)
    raw_original_grid = [list(r) for r in grid]

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

    seen_keys: set[tuple[str, ...]] = set()
    dedup_grid: list[list[str]] = []
    dedup_orig_indices: list[int] = []
    for r_idx, row in enumerate(grid):
        key = tuple(row[:dedup_cols])
        if key not in seen_keys:
            seen_keys.add(key)
            dedup_grid.append(list(row))
            dedup_orig_indices.append(r_idx if r_idx < len(raw_original_grid) else -1)
    after_dedup = len(dedup_grid)
    print(f"[DIAG] Após deduplicação: {after_dedup} linhas")

    def _norm_leiturista(v: str) -> str:
        collapsed = " ".join(v.strip().split())
        if collapsed.lower().startswith("leiturista:"):
            return "Leiturista: " + collapsed[11:].strip()
        return collapsed

    allowed_set = {_norm_leiturista(x) for x in allowed_leituristas}

    def is_header_candidate(row: list[str]) -> bool:
        norm_cells = [_norm(c) for c in row]
        has_medidor = any("medidor" in c for c in norm_cells)
        has_second = any(
            ("t. reg" in c or "t reg" in c or "leit. exec" in c or "leit exec" in c or c == "consumidora")
            for c in norm_cells
        )
        return has_medidor and has_second

    header_row_1based = 11
    matches_after_11 = sum(
        1 for row in dedup_grid[11:] if _norm_leiturista(row[4]) in allowed_set
    )
    if matches_after_11 == 0 or (len(dedup_grid) > 10 and not is_header_candidate(dedup_grid[10])):
        for idx, row in enumerate(dedup_grid):
            if is_header_candidate(row):
                header_row_1based = idx + 1
                break

    h_idx = header_row_1based - 1
    header_row = list(dedup_grid[h_idx]) if h_idx < len(dedup_grid) else [""] * dedup_cols
    orig_header_idx = dedup_orig_indices[h_idx] if h_idx < len(dedup_orig_indices) else -1

    filtered_rows: list[list[str]] = []
    filtered_orig_indices: list[int] = []
    for r_idx, row in enumerate(dedup_grid[h_idx + 1 :], start=h_idx + 1):
        if _norm_leiturista(row[4]) in allowed_set and "cemig distribuicao" not in " ".join(_norm(c) for c in row):
            filtered_rows.append(list(row))
            filtered_orig_indices.append(dedup_orig_indices[r_idx])

    if not filtered_rows:
        raise CemigValidationError(
            "Nenhuma linha restou após o filtro de leituristas na Etapa 1 (Coluna E)."
        )

    print(f"[DIAG] Após filtro leiturista: {len(filtered_rows)} linhas")
    print(f"[DIAG] Etapa 1 concluída: {len(filtered_rows)} linhas")

    intermediate_de = [(row[3], row[4]) for row in filtered_rows]
    return (
        header_row,
        filtered_rows,
        intermediate_de,
        after_dedup,
        len(filtered_rows),
        raw_original_grid,
        filtered_orig_indices,
        orig_header_idx,
    )


def parte_dois_ul_matr(intermediate_de: list[tuple[str, str]]) -> list[tuple[str, str]]:
    unique_pairs: list[tuple[str, str]] = []
    seen_units: set[str] = set()

    for raw_d, raw_e in intermediate_de:
        unidade = (raw_d[19:] if len(raw_d) > 19 else "").strip()
        matricula = (raw_e[11:] if len(raw_e) > 11 else "").strip()
        if unidade and unidade not in seen_units:
            seen_units.add(unidade)
            unique_pairs.append((unidade, matricula))

    print(f"[DIAG] Etapa 2 concluída: {len(unique_pairs)} linhas")
    return unique_pairs


def parte_tres_relatorio_final(
    header_row: list[str],
    filtered_rows: list[list[str]],
    ul_matr_pairs: list[tuple[str, str]],
    planilha3_grid: Optional[list[list[str]]],
    warnings: list[str],
    raw_original_grid: Optional[list[list[str]]] = None,
    filtered_orig_indices: Optional[list[int]] = None,
    orig_header_idx: int = -1,
) -> tuple[bytes, dict[str, int]]:
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
    src_endereco = find_col(["endereco", "endereço"], ["apres"], 15)
    src_ant = find_col(["leit. ant", "leit ant"], ["coluna6"], 10)
    src_exec = find_col(["leit. exec", "leit exec"], ["coluna7"], 11)
    src_imp = find_col(["imp.", "imp"], ["coluna8"], 13)
    src_dig = find_col(["qtd dig", "qtd. dig", "digitacao"], [], 14)
    src_apres = find_col(["apres.", "apres"], [], -1)

    consumed_cols = {
        0,
        src_medidor,
        src_registro,
        3,
        4,
        5,
        6,  # Leiturista excluída da tabela final (mantida apenas na ordenação)
        src_cod,
        src_data,
        src_hora,
        src_endereco,
        src_ant,
        src_exec,
        12,
        src_imp,
        src_dig,
    }
    if src_apres >= 0:
        consumed_cols.add(src_apres)

    used_headers_lower = {h.lower() for h in FINAL_HEADERS}

    def make_unique_header(candidate: str) -> str:
        clean = re.sub(r"[\r\n\t]+", " ", candidate).strip()[:255] or "Coluna_Extra"
        if clean.lower() not in used_headers_lower:
            used_headers_lower.add(clean.lower())
            return clean
        suffix = 2
        while f"{clean}_{suffix}".lower() in used_headers_lower:
            suffix += 1
        final_name = f"{clean}_{suffix}"[:255]
        used_headers_lower.add(final_name.lower())
        return final_name

    extra_specs: list[tuple[str, str, int]] = []
    for c, raw_h in enumerate(grid[0]):
        if c in consumed_cols:
            continue
        if raw_h.strip() and _is_excluded_extra_col(raw_h):
            continue
        has_data = any((row[c] or "").strip() != "" for row in grid[1:])
        if not raw_h.strip() and not has_data:
            continue
        extra_specs.append((make_unique_header(raw_h or f"Coluna_Bruta_{c + 1}"), "stage3", c))

    if raw_original_grid and filtered_orig_indices:
        orig_cols_count = max((len(r) for r in raw_original_grid), default=0)
        raw_h_row = (
            raw_original_grid[orig_header_idx]
            if 0 <= orig_header_idx < len(raw_original_grid)
            else []
        )
        for orig_c in (4, 7, 8):
            if orig_c >= orig_cols_count:
                continue
            raw_h = raw_h_row[orig_c].strip() if orig_c < len(raw_h_row) else ""
            if raw_h and _is_excluded_extra_col(raw_h):
                continue
            has_data = any(
                orig_r >= 0
                and orig_c < len(raw_original_grid[orig_r])
                and raw_original_grid[orig_r][orig_c].strip() != ""
                for orig_r in filtered_orig_indices
            )
            if not raw_h and not has_data:
                continue
            extra_specs.append(
                (make_unique_header(raw_h or f"Coluna_Original_{orig_c + 1}"), "raw", orig_c)
            )

    final_headers = [*FINAL_HEADERS, *[s[0] for s in extra_specs]]

    p3_ab: dict[str, str] = {}
    p3_de: dict[str, str] = {}
    p3_gh: dict[str, str] = {}
    p3_jk: dict[str, str] = {}

    if planilha3_grid is not None:
        for row in planilha3_grid[:50000]:
            if len(row) >= 2 and row[0].strip() and _norm(row[0]) != "rz_ul_lv" and row[0].strip() not in p3_ab:
                p3_ab[row[0].strip()] = row[1].strip()
            if len(row) >= 5 and row[3].strip() and _norm(row[3]) != "consistencia" and row[3].strip() not in p3_de:
                p3_de[row[3].strip()] = row[4].strip()
            if len(row) >= 8 and row[6].strip() and _norm(row[6]) != "nosb" and row[6].strip() not in p3_gh:
                p3_gh[row[6].strip()] = row[7].strip()
            if len(row) >= 11 and row[9].strip() and _norm(row[9]) != "fotos" and row[9].strip() not in p3_jk:
                p3_jk[row[9].strip()] = row[10].strip()
    else:
        for u, m in ul_matr_pairs:
            if u not in p3_ab:
                p3_ab[u] = m
        warnings.append("Tabela completa do tratamento.")

    prefer_mdy = _detect_prefer_mdy([r[src_data] for r in grid[1:] if r[src_data].strip()])

    records = []
    for idx, row in enumerate(grid[1:]):
        orig_r = filtered_orig_indices[idx] if filtered_orig_indices and idx < len(filtered_orig_indices) else -1
        inst = row[0].strip()
        medidor = row[src_medidor].strip()
        registro = row[src_registro].strip()
        rz = row[3].strip()
        ul = extract_ul_from_rz_ul_lv(rz)
        leit_fallback = row[6].replace("Leiturista:", "").strip()
        matr = p3_ab.get(rz) or leit_fallback
        tipo = classify_tipo_from_rz_ul_lv(rz)
        cod = row[src_cod].strip()
        data = _format_date_ddmmyyyy(row[src_data], prefer_mdy)
        hora = row[src_hora].strip()
        endereco = row[src_endereco].strip()
        ant_raw = row[src_ant].strip()
        exec_raw = row[src_exec].strip()
        imp = row[src_imp].strip() or "Faturada"
        dig = row[src_dig].strip() or "N00"
        apres = normalize_apres_value(row[src_apres].strip() if src_apres >= 0 else "")

        extra_vals: list[str] = []
        for _, src_kind, col_i in extra_specs:
            if src_kind == "stage3":
                extra_vals.append(row[col_i].strip() if col_i < len(row) else "")
            else:
                if raw_original_grid and orig_r >= 0 and col_i < len(raw_original_grid[orig_r]):
                    extra_vals.append(raw_original_grid[orig_r][col_i].strip())
                else:
                    extra_vals.append("")

        records.append(
            {
                "inst": inst,
                "medidor": medidor,
                "registro": registro,
                "rz": rz,
                "ul": ul,
                "matr": matr,
                "tipo": tipo,
                "cod": cod,
                "data": data,
                "hora": hora,
                "endereco": endereco,
                "ant_raw": ant_raw,
                "exec_raw": exec_raw,
                "imp": imp,
                "dig": dig,
                "apres": apres,
                "extra_vals": extra_vals,
                "sort_key": (matr.lower(), _date_sort_key(data), _time_sort_seconds(hora), idx),
            }
        )

    records.sort(key=lambda x: x["sort_key"])
    print(f"[DIAG] Etapa 3 concluída: {len(records)} linhas")
    print(f"[DIAG] Tabela final montada: {len(records)} linhas x {len(final_headers)} colunas")

    return _build_output_workbook(
        final_headers=final_headers,
        records=records,
        ul_matr_pairs=ul_matr_pairs,
        planilha3_grid=planilha3_grid,
        imported_meters_count=len(records),
    )


def _is_aux_header_row(col_a: str, col_b: str) -> bool:
    a_trim = col_a.strip()
    b_trim = col_b.strip()
    a_norm = _norm(a_trim)
    b_norm = _norm(b_trim)
    if not a_norm and not b_norm:
        return False
    if re.search(r"\d", a_trim):
        return False
    keywords = (
        "instalacao", "instalacoes", "unidade", "chave", "codigo", "cod",
        "consistencia", "nosb", "simulacao", "fotos", "foto",
        "valor", "resultado", "status", "descricao", "situacao",
        "medidor", "cliente", "observacao",
    )
    if any(kw == a_norm or kw in a_norm or kw == b_norm or kw in b_norm for kw in keywords):
        return True
    known_vals = {"ok", "n-ok", "nok", "faturada", "nao faturada", "sim", "nao", "pendente", "divergente"}
    a_words = [w for w in a_norm.split() if w]
    if re.search(r"[a-zA-Z]", a_norm) and not re.search(r"\d", a_norm) and len(a_words) <= 2 and b_norm not in known_vals:
        return True
    return False


def _read_aux_pair_columns(aux_bytes: bytes, aux_filename: str, aux_name: str) -> tuple[Optional[str], list[tuple[str, str]]]:
    _, grid, _ = read_sheet_as_text_grid(aux_bytes, aux_filename, preferred_sheet=None)
    non_empty = [r for r in grid if (len(r) > 0 and r[0].strip() != "") or (len(r) > 1 and r[1].strip() != "")]
    if not non_empty:
        return None, []
    first_a = non_empty[0][0].strip() if len(non_empty[0]) > 0 else ""
    first_b = non_empty[0][1].strip() if len(non_empty[0]) > 1 else ""
    has_header = _is_aux_header_row(first_a, first_b)
    start_idx = 1 if has_header else 0
    col_b_header = (first_b or "Valor") if has_header else "Valor"
    pairs: list[tuple[str, str]] = []
    for row in non_empty[start_idx:]:
        a = row[0].strip() if len(row) > 0 else ""
        b = row[1].strip() if len(row) > 1 else ""
        if a or b:
            pairs.append((a, b))
    print(f"[DIAG] Planilha auxiliar {aux_name}: {len(pairs)} linhas carregadas")
    return col_b_header, pairs


def _build_filtros_grid_with_aux(
    base_grid: Optional[list[list[str]]],
    ul_matr_pairs: list[tuple[str, str]],
    consistencia_bytes: Optional[bytes] = None,
    nosb_bytes: Optional[bytes] = None,
    fotos_bytes: Optional[bytes] = None,
) -> list[list[str]]:
    if base_grid is not None:
        p3_grid = [list(r) for r in base_grid]
    else:
        p3_grid = [
            ["RZ_UL_LV", "MATRÍCULA", "", "Consistencia", "", "", "Nosb", "", "", "Fotos", ""],
            *[[u, m] for u, m in ul_matr_pairs],
        ]
    _ensure_rect(p3_grid, 11)
    if not p3_grid:
        p3_grid.append([""] * 11)
    if not p3_grid[0][3].strip():
        p3_grid[0][3] = "Consistencia"
    if not p3_grid[0][6].strip():
        p3_grid[0][6] = "Nosb"
    if not p3_grid[0][9].strip():
        p3_grid[0][9] = "Fotos"

    def write_aux(col_b_header: Optional[str], rows: list[tuple[str, str]], key_col: int, val_col: int, fixed_hdr: str) -> None:
        p3_grid[0][key_col] = fixed_hdr
        if col_b_header:
            p3_grid[0][val_col] = col_b_header
        for r in range(1, len(p3_grid)):
            p3_grid[r][key_col] = ""
            p3_grid[r][val_col] = ""
        width = max(len(p3_grid[0]), 11)
        while len(p3_grid) < 1 + len(rows):
            p3_grid.append([""] * width)
        for i, (k, v) in enumerate(rows):
            p3_grid[1 + i][key_col] = k
            p3_grid[1 + i][val_col] = v

    if consistencia_bytes:
        hdr, rows = _read_aux_pair_columns(consistencia_bytes, "Consistencia.xlsx", "Consistencia")
        write_aux(hdr, rows, 3, 4, "Consistencia")
    if nosb_bytes:
        hdr, rows = _read_aux_pair_columns(nosb_bytes, "Nosb.xlsx", "Nosb")
        write_aux(hdr, rows, 6, 7, "Nosb")
    if fotos_bytes:
        hdr, rows = _read_aux_pair_columns(fotos_bytes, "Fotos.xlsx", "Fotos")
        write_aux(hdr, rows, 9, 10, "Fotos")

    return p3_grid


def _set_xml_attr(attrs: str, name: str, value: str) -> str:
    pattern = re.compile(rf'(\s|^){re.escape(name)}\s*=\s*(?:"[^"]*"|\'[^\']*\')')
    if pattern.search(attrs):
        return pattern.sub(rf'\1{name}="{value}"', attrs, count=1)
    trimmed = attrs.strip()
    return f'{trimmed} {name}="{value}"' if trimmed else f'{name}="{value}"'


def _patch_xlsm_zip_rgeral(
    modelo_bytes: bytes,
    records: list[dict[str, Any]],
) -> tuple[bytes, bool, list[str], list[str], int, int, bool]:
    text_preserving_cols = {1, 2, 3, 4, 5, 9, 10, 11, 19}
    with zipfile.ZipFile(io.BytesIO(modelo_bytes), "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    input_has_vba = "xl/vbaProject.bin" in entries
    input_pivot_count = sum(1 for n in entries if re.match(r"^xl/pivotTables/pivotTable\d+\.xml$", n, re.I))
    input_chart_count = sum(1 for n in entries if re.match(r"^xl/charts/chart\d+\.xml$", n, re.I))

    wb_xml = entries["xl/workbook.xml"].decode("utf-8")
    rels_xml = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct_xml = entries["[Content_Types].xml"].decode("utf-8")

    rel_map: dict[str, str] = {}
    for m in re.finditer(r"<Relationship\b([^>]*?)/?>", rels_xml):
        attrs = m.group(1)
        rid_m = re.search(r'\bId="([^"]*)"', attrs)
        tgt_m = re.search(r'\bTarget="([^"]*)"', attrs)
        if rid_m and tgt_m:
            rel_map[rid_m.group(1)] = tgt_m.group(1)

    original_sheets: list[str] = []
    sheet_path: Optional[str] = None
    target_sheet_id: Optional[str] = None
    max_sid = 0
    for m in re.finditer(r"<sheet\b([^>]*?)/?>", wb_xml):
        attrs = m.group(1)
        name_m = re.search(r'\bname="([^"]*)"', attrs)
        sid_m = re.search(r'\bsheetId="(\d+)"', attrs)
        rid_m = re.search(r'\br:id="([^"]*)"', attrs)
        if name_m and name_m.group(1):
            original_sheets.append(name_m.group(1))
        if sid_m:
            max_sid = max(max_sid, int(sid_m.group(1)))
        if name_m and name_m.group(1) == "RGeral" and rid_m:
            tgt = rel_map.get(rid_m.group(1), "").lstrip("/")
            sheet_path = tgt if tgt.startswith("xl/") else f"xl/{tgt}"
            if sid_m:
                target_sheet_id = sid_m.group(1)

    sheet_found = bool(sheet_path and sheet_path in entries)
    if sheet_found and sheet_path:
        sheet_xml = entries[sheet_path].decode("utf-8")
    else:
        n = 1
        while f"xl/worksheets/sheet{n}.xml" in entries:
            n += 1
        k = 1
        while f"rId{k}" in rel_map:
            k += 1
        next_rid = f"rId{k}"
        next_sid = str(max_sid + 1)
        target_sheet_id = next_sid
        sheet_path = f"xl/worksheets/sheet{n}.xml"
        wb_xml = wb_xml.replace("</sheets>", f'<sheet name="RGeral" sheetId="{next_sid}" r:id="{next_rid}"/></sheets>')
        rels_xml = rels_xml.replace(
            "</Relationships>",
            f'<Relationship Id="{next_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>',
        )
        if f'PartName="/{sheet_path}"' not in ct_xml:
            ct_xml = ct_xml.replace(
                "</Types>",
                f'<Override PartName="/{sheet_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
            )
        entries["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
        entries["[Content_Types].xml"] = ct_xml.encode("utf-8")
        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="A1:T1"/>'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews><sheetData/></worksheet>'
        )

    sd_m = re.search(r"(<sheetData\b[^>]*>[\s\S]*?</sheetData>|<sheetData\b[^>]*/>)", sheet_xml)
    if not sd_m:
        raise CemigValidationError("Tag <sheetData> não encontrada na aba RGeral.")

    existing_rows: dict[int, str] = {}
    for rm in re.finditer(r"<row\b([^>]*?)(?:/>|>[\s\S]*?</row>)", sd_m.group(0)):
        r_m = re.search(r'\br="(\d+)"', rm.group(1))
        if r_m:
            existing_rows[int(r_m.group(1))] = rm.group(0)

    def _extract_row_cells(row_xml: str) -> dict[int, tuple[Optional[str], str]]:
        out: dict[int, tuple[Optional[str], str]] = {}
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>[\s\S]*?</c>)", row_xml):
            cattrs = cm.group(1)
            ref_m = re.search(r'\br="([A-Za-z]+)\d+"', cattrs)
            s_m = re.search(r'\bs="(\d+)"', cattrs)
            if ref_m:
                c_idx = openpyxl.utils.column_index_from_string(ref_m.group(1).upper())
                out[c_idx] = (s_m.group(1) if s_m else None, cm.group(0))
        return out

    rebuilt: list[str] = []
    if 1 in existing_rows and re.search(r"<(?:v|t)\b", existing_rows[1]):
        rebuilt.append(existing_rows[1])
    else:
        hdr_cells = "".join(
            f'<c r="{get_column_letter(c)}1" t="inlineStr"><is><t>{xml_escape(h)}</t></is></c>'
            for c, h in enumerate(RELATORIO_GERAL_HEADERS, start=1)
        )
        rebuilt.append(f'<row r="1" spans="1:20">{hdr_cells}</row>')

    if 2 in existing_rows:
        rebuilt.append(existing_rows[2])

    max_row_written = 2 if 2 in existing_rows else 1
    max_col_used = 20

    for i, rec in enumerate(records, start=3):
        max_row_written = i
        ant_val = (
            int(rec["ant_raw"])
            if rec["ant_raw"].isdigit() and not (len(rec["ant_raw"]) > 1 and rec["ant_raw"].startswith("0"))
            else rec["ant_raw"]
        )
        exec_val = (
            int(rec["exec_raw"])
            if rec["exec_raw"].isdigit() and not (len(rec["exec_raw"]) > 1 and rec["exec_raw"].startswith("0"))
            else rec["exec_raw"]
        )
        consumo_val: Any = exec_val - ant_val if isinstance(ant_val, int) and isinstance(exec_val, int) else ""
        row_values = [
            rec["inst"], rec["medidor"], rec["registro"], rec["rz"], rec["ul"],
            rec.get("fotos", "N-OK"), rec["tipo"], rec["matr"], rec["cod"], rec["data"],
            rec["hora"], rec["endereco"], ant_val, exec_val, consumo_val,
            rec["imp"], rec.get("simulacao", "Faturada"), rec["dig"], rec["apres"],
            rec.get("consistencia", "Faturada"),
        ]
        prev_cells = _extract_row_cells(existing_rows[i]) if i in existing_rows else {}
        cells_xml: list[str] = []
        for c_idx, val in enumerate(row_values, start=1):
            ref = f"{get_column_letter(c_idx)}{i}"
            orig_s = prev_cells.get(c_idx, (None, ""))[0]
            s_attr = f' s="{orig_s}"' if orig_s is not None else ""
            if isinstance(val, int):
                cells_xml.append(f'<c r="{ref}"{s_attr}><v>{val}</v></c>')
            else:
                sval = str(val or "")
                if sval != "":
                    _ = c_idx in text_preserving_cols
                    cells_xml.append(
                        f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{xml_escape(sval)}</t></is></c>'
                    )
                elif orig_s is not None:
                    cells_xml.append(f'<c r="{ref}"{s_attr}/>')
        for c_idx in sorted(prev_cells):
            if c_idx > 20:
                max_col_used = max(max_col_used, c_idx)
                cells_xml.append(prev_cells[c_idx][1])
        rebuilt.append(f'<row r="{i}" spans="1:20">{"".join(cells_xml)}</row>')

    # Preservar células em colunas > 20 nas linhas excedentes (limpando A..T)
    for r_num in sorted(r for r in existing_rows if r >= 3 + len(records)):
        prev_cells = _extract_row_cells(existing_rows[r_num])
        extra = [(c, prev_cells[c][1]) for c in sorted(prev_cells) if c > 20]
        if extra:
            max_row_written = max(max_row_written, r_num)
            max_col_used = max(max_col_used, extra[-1][0])
            rebuilt.append(f'<row r="{r_num}">{"".join(x[1] for x in extra)}</row>')

    new_sd = f'<sheetData>{"".join(rebuilt)}</sheetData>'
    sheet_xml = sheet_xml[: sd_m.start()] + new_sd + sheet_xml[sd_m.end() :]
    new_dim = f"A1:{get_column_letter(max_col_used)}{max(max_row_written, 3)}"
    if re.search(r"<dimension\b[^>]*/>", sheet_xml):
        sheet_xml = re.sub(
            r"<dimension\b([^>]*?)/>",
            lambda dm: f'<dimension {_set_xml_attr(dm.group(1), "ref", new_dim).strip()}/>',
            sheet_xml,
            count=1,
        )
    entries[sheet_path] = sheet_xml.encode("utf-8")

    # Remover entradas de calcChain apenas para células sobrescritas em RGeral (A3:T...)
    if "xl/calcChain.xml" in entries and target_sheet_id:
        cc_xml = entries["xl/calcChain.xml"].decode("utf-8")
        changed_cc = False

        def _filter_cc(match: re.Match[str]) -> str:
            nonlocal changed_cc
            attrs = match.group(1)
            i_m = re.search(r'\bi="([^"]*)"', attrs)
            r_m = re.search(r'\br="([A-Za-z]+)(\d+)"', attrs)
            if i_m and r_m and i_m.group(1) == target_sheet_id:
                c_idx = openpyxl.utils.column_index_from_string(r_m.group(1).upper())
                r_idx = int(r_m.group(2))
                if r_idx >= 3 and c_idx <= 20:
                    changed_cc = True
                    return ""
            return match.group(0)

        cc_xml = re.sub(r"<c\b([^>]*?)/>", _filter_cc, cc_xml)
        if changed_cc:
            if not re.search(r"<c\b", cc_xml):
                entries.pop("xl/calcChain.xml", None)
                rels_xml = re.sub(
                    r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>',
                    "",
                    entries["xl/_rels/workbook.xml.rels"].decode("utf-8"),
                )
                entries["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
                ct_xml = re.sub(
                    r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>',
                    "",
                    entries["[Content_Types].xml"].decode("utf-8"),
                )
                entries["[Content_Types].xml"] = ct_xml.encode("utf-8")
            else:
                entries["xl/calcChain.xml"] = cc_xml.encode("utf-8")

    # Garantir recálculo automático nas demais abas (Resumo, Dashboard, etc.) ao abrir no Excel
    if re.search(r"<calcPr\b[^>]*?/?>", wb_xml):
        wb_xml = re.sub(
            r"<calcPr\b([^>]*?)(/?)>",
            lambda cm: f'<calcPr {_set_xml_attr(cm.group(1), "fullCalcOnLoad", "1").strip()}{cm.group(2)}>',
            wb_xml,
            count=1,
        )
    elif "</workbook>" in wb_xml:
        if "<extLst" in wb_xml:
            wb_xml = wb_xml.replace("<extLst", '<calcPr fullCalcOnLoad="1"/><extLst', 1)
        else:
            wb_xml = wb_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1)
    entries["xl/workbook.xml"] = wb_xml.encode("utf-8")

    # Se existirem caches de PivotTable, marcar refreshOnLoad="1"
    for fname in list(entries.keys()):
        if re.match(r"^xl/pivotCache/pivotCacheDefinition\d+\.xml$", fname, re.I):
            pc_xml = entries[fname].decode("utf-8")
            updated_pc = re.sub(
                r"<pivotCacheDefinition\b([^>]*?)>",
                lambda pm: f'<pivotCacheDefinition {_set_xml_attr(pm.group(1), "refreshOnLoad", "1").strip()}>',
                pc_xml,
                count=1,
            )
            entries[fname] = updated_pc.encode("utf-8")

    out_stream = io.BytesIO()
    with zipfile.ZipFile(out_stream, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    return (
        out_stream.getvalue(),
        sheet_found,
        list(RELATORIO_GERAL_HEADERS),
        original_sheets,
        input_pivot_count,
        input_chart_count,
        input_has_vba,
    )


def fill_relatorio_geral_template_bytes(
    modelo_bytes: bytes,
    records: list[dict[str, Any]],
    modelo_filename: str = "Relatorio Geral.xlsx",
) -> bytes:
    is_xlsm = modelo_filename.lower().endswith(".xlsm")
    ext_label = ".xlsm" if is_xlsm else ".xlsx"
    out_name = f"Relatorio Geral{ext_label}"
    print(f"[RGeral] Modelo recebido: {modelo_filename}, {len(modelo_bytes)} bytes, extensão: {ext_label}")

    (
        out_bytes,
        sheet_found,
        hdrs,
        original_sheets,
        input_pivot_count,
        input_chart_count,
        input_has_vba,
    ) = _patch_xlsm_zip_rgeral(modelo_bytes, records)

    print(f"[RGeral] Abas do modelo original: {original_sheets}")
    print(f"[RGeral] Aba 'RGeral' encontrada: {'sim' if sheet_found else 'não'}")
    print(f"[RGeral] Cabeçalhos na linha 1: {hdrs}")
    print(f"[RGeral] Linhas de dados a colar: {len(records)}")
    print(f"[RGeral] Dados colados como VALORES (não fórmulas): {len(records)} linhas x 20 colunas")
    if records:
        print(f"[RGeral] Dados colados a partir de A3 até T{2 + len(records)}")
    print(f"[RGeral] Arquivo gerado: {out_name}, {len(out_bytes)} bytes (modelo enviado: {len(modelo_bytes)} bytes)")

    try:
        with zipfile.ZipFile(io.BytesIO(out_bytes), "r") as zchk:
            out_files = zchk.namelist()
            has_vba = "xl/vbaProject.bin" in out_files
            out_pivot_count = sum(1 for n in out_files if re.match(r"^xl/pivotTables/pivotTable\d+\.xml$", n, re.I))
            out_chart_count = sum(1 for n in out_files if re.match(r"^xl/charts/chart\d+\.xml$", n, re.I))
            print(f"[RGeral] PivotTables preservadas: {out_pivot_count}")
            print(f"[RGeral] Gráficos preservados: {out_chart_count}")
            if out_pivot_count < input_pivot_count:
                raise RuntimeError(f"PivotTables perdidas na gravação ({out_pivot_count} < {input_pivot_count})")
            if out_chart_count < input_chart_count:
                raise RuntimeError(f"Gráficos perdidos na gravação ({out_chart_count} < {input_chart_count})")
            if is_xlsm:
                print(f"[RGeral] vbaProject.bin preservado: {'sim' if has_vba else 'não'}")
                if input_has_vba and not has_vba:
                    raise RuntimeError("xl/vbaProject.bin ausente após gravação do .xlsm")
            else:
                print("[RGeral] vbaProject.bin preservado: não-aplicável")

        chk_wb = openpyxl.load_workbook(io.BytesIO(out_bytes), read_only=True, keep_vba=is_xlsm)
        final_sheets = list(chk_wb.sheetnames)
        missing_sheets = [s for s in original_sheets if s not in final_sheets]
        print(f"[RGeral] Abas preservadas no arquivo final: {final_sheets}")
        print(f"[RGeral] Match de abas (original = final): {'sim' if not missing_sheets else 'não'}")
        if missing_sheets:
            print(f"[RGeral] Abas perdidas: {missing_sheets}")
            for missing in missing_sheets:
                print(f"[ERRO] Aba {missing} perdida na gravação")
            chk_wb.close()
            raise RuntimeError(f"Abas originais perdidas na gravação: {', '.join(missing_sheets)}")
        if "RGeral" not in final_sheets:
            chk_wb.close()
            print("[ERRO] Aba RGeral perdida na gravação")
            raise RuntimeError("Aba RGeral ausente ao reabrir arquivo gerado.")
        ws_chk = chk_wb["RGeral"]
        for r_idx, row_vals in enumerate(
            ws_chk.iter_rows(min_row=3, max_row=2 + min(len(records), 50), min_col=1, max_col=20, values_only=True),
            start=3,
        ):
            for c_idx, cell_val in enumerate(row_vals, start=1):
                if isinstance(cell_val, str) and cell_val.startswith("="):
                    chk_wb.close()
                    raise RuntimeError(
                        f"Célula {get_column_letter(c_idx)}{r_idx} em RGeral contém fórmula ({cell_val}) em vez de valor estático."
                    )
        chk_wb.close()
        print("[RGeral] Reabertura pós-gravação: sucesso")
    except Exception as exc:
        print("[ERRO] Arquivo gerado está corrompido")
        print(f"[RGeral] ERRO: {exc}")
        print("[RGeral] Reabertura pós-gravação: falha")
        raise
    return out_bytes


def _build_output_workbook(
    final_headers: list[str],
    records: list[dict[str, Any]],
    ul_matr_pairs: list[tuple[str, str]],
    planilha3_grid: Optional[list[list[str]]],
    imported_meters_count: int,
) -> tuple[bytes, dict[str, int]]:
    kpi_counts = {
        "unique_instalacoes_count": sum(1 for r in records if str(r.get("inst") or "").strip() != ""),
        "rural_count": sum(1 for r in records if r["tipo"] == "Rural"),
        "povoado_count": sum(1 for r in records if r["tipo"] == "Povoado"),
        "urbano_count": sum(1 for r in records if r["tipo"] == "Urbano"),
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ControleLeitura"
    ws.freeze_panes = "A2"

    ws_filtros = wb.create_sheet("Filtros")
    p3_rows = planilha3_grid if planilha3_grid is not None else [
        ["RZ_UL_LV", "MATRÍCULA", "", "Consistencia", "", "", "Nosb", "", "", "Fotos", ""],
        *[[u, m] for u, m in ul_matr_pairs],
    ]
    _ensure_rect(p3_rows, 11)
    for r_idx, row in enumerate(p3_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            if val != "":
                c = ws_filtros.cell(r_idx, c_idx, val)
                c.number_format = "@"
    filtros_rows = min(max(len(p3_rows), 50), 20000)

    if not ws_filtros.cell(1, 4).value:
        ws_filtros.cell(1, 4, "Consistencia")
    if not ws_filtros.cell(1, 7).value:
        ws_filtros.cell(1, 7, "Nosb")
    if not ws_filtros.cell(1, 10).value:
        ws_filtros.cell(1, 10, "Fotos")

    filtros_fills = {
        4: PatternFill("solid", fgColor="DCE6F1"),
        5: PatternFill("solid", fgColor="DCE6F1"),
        7: PatternFill("solid", fgColor="E2EFDA"),
        8: PatternFill("solid", fgColor="E2EFDA"),
        10: PatternFill("solid", fgColor="FFF2CC"),
        11: PatternFill("solid", fgColor="FFF2CC"),
    }
    for r_idx in range(1, filtros_rows + 1):
        for col_num, fill_obj in filtros_fills.items():
            cell = ws_filtros.cell(r_idx, col_num)
            cell.fill = fill_obj
            if r_idx == 1:
                cell.font = Font(name="Calibri", size=11, bold=True, color="0F172A")

    vlookup_end_row = max(50000, len(p3_rows))
    align = Alignment(horizontal="left", vertical="bottom", wrap_text=False)

    for c_idx, h_name in enumerate(final_headers, start=1):
        cell = ws.cell(1, c_idx, h_name)
        cell.alignment = align
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_hex = "0F766E" if 18 <= c_idx <= 20 else ("1E293B" if c_idx > 20 else "0F172A")
        cell.fill = PatternFill("solid", fgColor=fill_hex)

    print(f"[DIAG] Total de linhas no DataFrame tratado (antes de gravar): {len(records)}")
    print(f"[DIAG] Total de colunas no DataFrame tratado: {len(final_headers)} — colunas: {final_headers}")
    print(f"[DIAG] Shape do DataFrame antes de gravar: {len(records)} x {len(final_headers)}")
    if records:
        print(f"[DIAG] Primeira linha: {records[0]!r}")
        print(f"[DIAG] Última linha: {records[-1]!r}")

    for idx, rec in enumerate(records, start=2):
        ant_val = (
            int(rec["ant_raw"])
            if rec["ant_raw"].isdigit() and not (len(rec["ant_raw"]) > 1 and rec["ant_raw"].startswith("0"))
            else rec["ant_raw"]
        )
        exec_val = (
            int(rec["exec_raw"])
            if rec["exec_raw"].isdigit() and not (len(rec["exec_raw"]) > 1 and rec["exec_raw"].startswith("0"))
            else rec["exec_raw"]
        )

        row_values = [
            rec["inst"],
            rec["medidor"],
            rec["registro"],
            rec["rz"],
            rec["ul"],
            f"=VLOOKUP(D{idx},Filtros!$A$1:$B${vlookup_end_row},2,0)",
            rec["tipo"],
            rec["cod"],
            rec["data"],
            rec["hora"],
            rec["endereco"],
            ant_val,
            exec_val,
            f"=SUM(M{idx}-L{idx})",
            rec["imp"],
            rec["dig"],
            rec["apres"],
            f'=IFERROR(VLOOKUP(B{idx},Filtros!$D$1:$E${vlookup_end_row},2,0),"Faturada")',
            f'=IFERROR(VLOOKUP(B{idx},Filtros!$G$1:$H${vlookup_end_row},2,0),"Faturada")',
            f'=IFERROR(VLOOKUP(B{idx},Filtros!$J$1:$K${vlookup_end_row},2,0),"N-OK")',
            *rec["extra_vals"],
        ]
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(idx, c_idx, val)
            cell.alignment = align
            if isinstance(val, str) and (
                (val.isdigit() and len(val) > 1 and val.startswith("0")) or c_idx in (5, 9)
            ):
                cell.number_format = "@"

    last_row = len(records) + 1
    last_col_letter = get_column_letter(len(final_headers))
    tab = Table(displayName="Tabela1", ref=f"A1:{last_col_letter}{last_row}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    for c_idx in range(1, len(final_headers) + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = max(len(final_headers[c_idx - 1]) + 4, 14)

    controle_horario = build_controle_horario_rows(records)
    ws_ch = wb.create_sheet("Controle Horario")
    ws_ch.freeze_panes = "A2"
    ws_ch.auto_filter.ref = "A1:E1"
    for c_idx, h_name in enumerate(CONTROLE_HORARIO_HEADERS, start=1):
        cell = ws_ch.cell(1, c_idx, h_name)
        cell.alignment = align
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0F172A")
        ws_ch.column_dimensions[get_column_letter(c_idx)].width = max(len(h_name) + 6, 15)

    for r_idx, ch in enumerate(controle_horario, start=2):
        vals = [ch["matricula"], ch["data"], ch["hora_inicial"], ch["hora_final"], ch["tempo_servico"]]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws_ch.cell(r_idx, c_idx, val)
            cell.alignment = align
            cell.number_format = "@"

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_bytes = out_stream.getvalue()

    verify_wb = openpyxl.load_workbook(io.BytesIO(out_bytes), read_only=True)
    verify_ws = verify_wb["ControleLeitura"]
    xlsx_rows = list(verify_ws.iter_rows(values_only=True))
    xlsx_headers = [str(v or "") for v in (xlsx_rows[0] if xlsx_rows else [])]
    xlsx_data_rows = max(0, len(xlsx_rows) - 1)
    xlsx_cols = len(xlsx_headers)
    match_post_write = xlsx_data_rows == len(records) and xlsx_cols == len(final_headers)
    match_meters = imported_meters_count == len(records) == xlsx_data_rows
    print(f"[DIAG] Total de linhas efetivamente gravadas no .xlsx: {xlsx_data_rows}")
    print(f"[DIAG] Total de colunas efetivamente gravadas: {xlsx_cols} — colunas: {xlsx_headers}")
    print(f"[DIAG] Shape do .xlsx após gravação: {xlsx_data_rows} x {xlsx_cols}")
    print(f"[DIAG] Tamanho do .xlsx gerado: {len(out_bytes)} bytes")
    print(f"[DIAG] Verificação pós-gravação: .xlsx tem {xlsx_data_rows} linhas; DataFrame tem {len(records)} linhas. Match: {'sim' if match_post_write else 'não'}")
    print(f"[DIAG] Total de medidores no arquivo importado: {imported_meters_count}")
    print(f"[DIAG] Total de medidores na tabela final: {len(records)}")
    print(f"[DIAG] Total de medidores no .xlsx exportado: {xlsx_data_rows}")
    print(f"[DIAG] Match (importado = final = exportado): {'sim' if match_meters else 'não'}")
    if not match_meters or not match_post_write:
        print(f"[DIAG] TRUNCAMENTO/OMISSÃO DETECTADO: importado={imported_meters_count}, final={len(records)}, exportado={xlsx_data_rows}, diferença={imported_meters_count - xlsx_data_rows}")
    verify_wb.close()

    return out_bytes, kpi_counts


def is_leitura_3line_layout(grid: list[list[str]]) -> bool:
    for row in grid[:300]:
        if len(row) < 4:
            continue
        norm_a = _norm(row[0])
        norm_b = _norm(row[1])
        norm_c = _norm(row[2])
        norm_d = _norm(row[3])
        norm_joined = " | ".join(_norm(c) for c in row)

        if norm_a == "unidade" and norm_b == "medidor" and ("t. reg" in norm_c or "t.reg" in norm_c or "t reg" in norm_c):
            return True
        if norm_b == "medidor" and norm_d == "oc" and ("a exec" in norm_joined or "exec" in norm_joined):
            return True
        if norm_a == "cliente" and "endereco" in norm_joined and ("qtd dig" in norm_joined or "qtd. dig" in norm_joined):
            return True
    return False


def _is_time_like(val: str) -> bool:
    return bool(re.match(r"^\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)?$", val.strip(), re.I))


def _is_cpf_cnpj_or_id(val: str) -> bool:
    clean = val.strip()
    if not clean or re.search(r"[a-zA-Z]", clean):
        return False
    digits = re.sub(r"\D", "", clean)
    return len(digits) >= 6 and bool(re.match(r"^[\d.\-/\s]+$", clean))


def process_leitura_3line_grid(
    raw_grid: list[list[str]],
    p3_grid: Optional[list[list[str]]],
    warnings: list[str],
    consistencia_bytes: Optional[bytes] = None,
    nosb_bytes: Optional[bytes] = None,
    fotos_bytes: Optional[bytes] = None,
) -> tuple[bytes, ProcessingReport]:
    raw_rows = len(raw_grid)
    raw_cols = max((len(r) for r in raw_grid), default=0)
    print(f"[DIAG] Linhas lidas no arquivo bruto: {raw_rows}")

    header_rows_ignored = 0
    aux_rows_merged = 0
    current_rz_ul_lv = ""
    current_leiturista = ""
    blocks: list[dict[str, Any]] = []
    active_block: Optional[dict[str, Any]] = None

    def flush_active() -> None:
        nonlocal active_block
        if active_block is not None:
            blocks.append(active_block)
            active_block = None

    def extract_page_context(row: list[str]) -> None:
        nonlocal current_rz_ul_lv, current_leiturista
        for c_idx, raw_cell in enumerate(row):
            cell = raw_cell.strip()
            if not cell:
                continue
            n = _norm(cell)
            if n.startswith("unidade de leitura") and "inicial" not in n and "final" not in n:
                after_colon = cell.split(":", 1)[1].strip() if ":" in cell else ""
                if after_colon:
                    current_rz_ul_lv = after_colon
                else:
                    for nxt in row[c_idx + 1 :]:
                        nxt_val = nxt.strip()
                        if nxt_val and "data" not in _norm(nxt_val):
                            current_rz_ul_lv = nxt_val
                            break
            if n.startswith("leiturista"):
                code = cell.split(":", 1)[1].strip() if ":" in cell else ""
                if not code:
                    for nxt in row[c_idx + 1 :]:
                        nxt_val = nxt.strip()
                        if nxt_val and "tipo" not in _norm(nxt_val):
                            code = nxt_val
                            break
                if code:
                    current_leiturista = code

    def is_header_or_banner(row: list[str]) -> bool:
        if not any(c.strip() for c in row):
            return False
        norm_a = _norm(row[0]) if len(row) > 0 else ""
        norm_b = _norm(row[1]) if len(row) > 1 else ""
        norm_c = _norm(row[2]) if len(row) > 2 else ""
        norm_joined = " ".join(_norm(c) for c in row)

        if (
            "cemig distribuicao" in norm_a
            or norm_a == "sequencia de leitura executada"
            or bool(re.search(r"\bpag\.\s*\d+\s*/", norm_joined))
            or "unidade de leitura final" in norm_joined
            or "unidade de leitura inicial" in norm_joined
            or norm_a.startswith("unidade de leitura")
            or "unidade de leitura:" in norm_joined
            or norm_a.startswith("leiturista:")
            or "leiturista:" in norm_joined
            or norm_a == "leiturista"
            or "tipo de leitura:" in norm_joined
            or norm_a.startswith("data inicio")
            or "data termino" in norm_joined
            or (norm_a == "contrato" and "mes/ano" in norm_joined)
            or bool(re.match(r"^[-=_]{5,}$", norm_a))
        ):
            return True

        if re.match(r"^\d{8,}$", norm_a) and re.match(r"^\d{1,2}/\d{4}$", norm_b) and not norm_c:
            return True

        if norm_a == "unidade" and norm_b == "medidor":
            return True
        if norm_b == "medidor" and ("t. reg" in norm_c or "t.reg" in norm_c or "t reg" in norm_c):
            return True
        if norm_a == "consumidora" and not norm_b and not norm_c:
            return True
        if norm_a == "cliente" and not norm_b and not norm_c and (
            "endereco" in norm_joined or "qtd dig" in norm_joined or "qtd. dig" in norm_joined
        ):
            return True
        return False

    for row in raw_grid:
        if not any(c.strip() for c in row):
            continue
        if is_header_or_banner(row):
            flush_active()
            extract_page_context(row)
            header_rows_ignored += 1
            continue

        extract_page_context(row)
        col_a = row[0].strip() if len(row) > 0 else ""
        col_b = row[1].strip() if len(row) > 1 else ""
        col_c = row[2].strip() if len(row) > 2 else ""

        if col_b != "" and col_c != "":
            flush_active()
            active_block = {
                "inst": col_a,
                "medidor": col_b,
                "registro": col_c,
                "cod": row[3].strip() if len(row) > 3 else "",
                "raw_data": row[5].strip() if len(row) > 5 else "",
                "hora": _format_hora_hhmm(row[6].strip() if len(row) > 6 else ""),
                "a_exec": row[7].strip() if len(row) > 7 else "",
                "exec": row[8].strip() if len(row) > 8 else "",
                "ant_raw": row[9].strip() if len(row) > 9 else "",
                "exec_raw": row[10].strip() if len(row) > 10 else "",
                "imp_raw": row[11].strip() if len(row) > 11 else "",
                "apres": row[12].strip() if len(row) > 12 else "",
                "rz": current_rz_ul_lv,
                "leiturista": current_leiturista,
                "aux_full_time": "",
                "endereco": "",
                "qtd_dig": "",
                "consumidora": "",
                "cliente": "",
                "orig_order": len(blocks) + 1,
            }
            continue

        if active_block is not None:
            aux_rows_merged += 1
            for t_idx in (5, 6, 4, 7):
                if t_idx < len(row):
                    cand = row[t_idx].strip()
                    if cand and _is_time_like(cand):
                        if not active_block["aux_full_time"]:
                            active_block["aux_full_time"] = cand
                        break

            for d_idx in (12, 13, 14):
                if d_idx < len(row):
                    cand = row[d_idx].strip()
                    if cand != "":
                        active_block["qtd_dig"] = cand
                        break

            for e_idx in (6, 7):
                if e_idx < len(row):
                    cand = row[e_idx].strip()
                    if cand and not _is_time_like(cand):
                        if len(cand) > len(active_block["endereco"]):
                            active_block["endereco"] = cand

            col_d = row[3].strip() if len(row) > 3 else ""
            if col_d != "" and not active_block["cliente"]:
                active_block["cliente"] = col_d

            if col_a != "":
                if _is_cpf_cnpj_or_id(col_a):
                    if not active_block["cliente"]:
                        active_block["cliente"] = col_a
                elif not active_block["consumidora"]:
                    active_block["consumidora"] = col_a
                elif not active_block["cliente"]:
                    active_block["cliente"] = col_a

    flush_active()
    detected_meters_count = len(blocks)
    print(f"[DIAG] Linhas de cabeçalho ignoradas: {header_rows_ignored}")
    print(f"[DIAG] Medidores detectados (linhas principais): {detected_meters_count}")
    print(f"[DIAG] Linhas auxiliares mescladas: {aux_rows_merged}")

    if detected_meters_count == 0:
        raise CemigValidationError(
            "Nenhum medidor foi detectado no relatório CEMIG (nenhuma linha principal com Medidor na coluna B e T.Reg na coluna C)."
        )

    ul_matr_pairs: list[tuple[str, str]] = []
    seen_units: set[str] = set()
    for blk in blocks:
        unid = blk["rz"].strip()
        matr = (blk["leiturista"] or unid).strip()
        if unid and unid not in seen_units:
            seen_units.add(unid)
            ul_matr_pairs.append((unid, matr))

    merged_p3 = _build_filtros_grid_with_aux(
        p3_grid,
        ul_matr_pairs,
        consistencia_bytes=consistencia_bytes,
        nosb_bytes=nosb_bytes,
        fotos_bytes=fotos_bytes,
    )

    p3_ab: dict[str, str] = {}
    for row in merged_p3:
        if len(row) >= 2 and row[0].strip() and _norm(row[0]) != "rz_ul_lv" and row[0].strip() not in p3_ab:
            p3_ab[row[0].strip()] = row[1].strip()
    if p3_grid is None:
        warnings.append("Tabela completa do tratamento.")

    prefer_mdy = _detect_prefer_mdy([b["raw_data"] for b in blocks if b["raw_data"]])

    empty_endereco = 0
    empty_qtd_dig = 0
    empty_consumidora = 0
    empty_cliente = 0
    records: list[dict[str, Any]] = []

    for blk in blocks:
        if not blk["endereco"]:
            empty_endereco += 1
        if not blk["qtd_dig"]:
            empty_qtd_dig += 1
        if not blk["consumidora"]:
            empty_consumidora += 1
        if not blk["cliente"]:
            empty_cliente += 1

        rz = blk["rz"].strip()
        ul = extract_ul_from_rz_ul_lv(rz)
        matr = p3_ab.get(rz) or blk["leiturista"] or rz
        tipo = classify_tipo_from_rz_ul_lv(rz)
        data = _format_date_ddmmyyyy(blk["raw_data"], prefer_mdy)
        hora = blk["hora"] or _format_hora_hhmm(blk["aux_full_time"])
        imp = blk["imp_raw"] or "Faturada"
        dig = blk["qtd_dig"] or "N00"
        apres = normalize_apres_value(blk["apres"])

        a_exec_k = int(blk["a_exec"]) if blk["a_exec"].isdigit() else 99999999
        exec_k = int(blk["exec"]) if blk["exec"].isdigit() else 99999999

        records.append(
            {
                "inst": blk["inst"],
                "medidor": blk["medidor"],
                "registro": blk["registro"],
                "rz": rz,
                "ul": ul,
                "matr": matr,
                "tipo": tipo,
                "cod": blk["cod"],
                "data": data,
                "hora": hora,
                "endereco": blk["endereco"],
                "ant_raw": blk["ant_raw"],
                "exec_raw": blk["exec_raw"],
                "imp": imp,
                "dig": dig,
                "apres": apres,
                "extra_vals": [],
                "sort_key": (
                    matr.lower(),
                    _date_sort_key(data),
                    _time_sort_seconds(blk["aux_full_time"] or hora),
                    a_exec_k,
                    exec_k,
                    blk["orig_order"],
                ),
            }
        )

    print(f"[DIAG] Linhas com Endereço vazio: {empty_endereco} (deve ser 0)")
    print(f"[DIAG] Linhas com Qtd Dig. vazio: {empty_qtd_dig} (deve ser 0)")
    print(f"[DIAG] Linhas com Consumidora vazio: {empty_consumidora}")
    print(f"[DIAG] Linhas com Cliente vazio: {empty_cliente}")

    records.sort(key=lambda x: x["sort_key"])
    final_headers = [*FINAL_HEADERS]

    out_bytes, kpi_counts = _build_output_workbook(
        final_headers=final_headers,
        records=records,
        ul_matr_pairs=ul_matr_pairs,
        planilha3_grid=merged_p3,
        imported_meters_count=detected_meters_count,
    )

    report = ProcessingReport(
        raw_rows=raw_rows,
        raw_cols=raw_cols,
        after_deduplication_rows=detected_meters_count,
        after_filter_rows=detected_meters_count,
        final_rows=len(records),
        unique_units_count=len(ul_matr_pairs),
        unique_instalacoes_count=kpi_counts["unique_instalacoes_count"],
        rural_count=kpi_counts["rural_count"],
        povoado_count=kpi_counts["povoado_count"],
        urbano_count=kpi_counts["urbano_count"],
        warnings=list(dict.fromkeys(warnings)),
    )
    return out_bytes, report


def process_matriz_bytes(
    matriz_bytes: bytes,
    matriz_filename: str = "Matriz.xlsx",
    planilha3_bytes: Optional[bytes] = None,
    planilha3_filename: str = "Planilha3.xlsx",
    allowed_leituristas: Optional[list[str]] = None,
    consistencia_bytes: Optional[bytes] = None,
    nosb_bytes: Optional[bytes] = None,
    fotos_bytes: Optional[bytes] = None,
) -> tuple[bytes, ProcessingReport]:
    warnings: list[str] = []
    allowed = allowed_leituristas or DEFAULT_ALLOWED_LEITURISTAS

    _, raw_grid, embedded_p3 = read_sheet_as_text_grid(matriz_bytes, matriz_filename)
    validate_cemig_layout(raw_grid)

    p3_grid = embedded_p3
    if planilha3_bytes is not None:
        _, p3_grid, _ = read_sheet_as_text_grid(planilha3_bytes, planilha3_filename, "Filtros")

    if is_leitura_3line_layout(raw_grid):
        return process_leitura_3line_grid(
            raw_grid,
            p3_grid,
            warnings,
            consistencia_bytes=consistencia_bytes,
            nosb_bytes=nosb_bytes,
            fotos_bytes=fotos_bytes,
        )

    raw_rows = len(raw_grid)
    raw_cols = max((len(r) for r in raw_grid), default=0)

    (
        header_row,
        filtered_rows,
        intermediate_de,
        after_dedup,
        after_filter,
        raw_original_grid,
        filtered_orig_indices,
        orig_header_idx,
    ) = parte_um_tratamento(raw_grid, allowed)
    ul_matr_pairs = parte_dois_ul_matr(intermediate_de)
    merged_p3 = _build_filtros_grid_with_aux(
        p3_grid,
        ul_matr_pairs,
        consistencia_bytes=consistencia_bytes,
        nosb_bytes=nosb_bytes,
        fotos_bytes=fotos_bytes,
    )

    out_bytes, kpi_counts = parte_tres_relatorio_final(
        header_row,
        filtered_rows,
        ul_matr_pairs,
        merged_p3,
        warnings,
        raw_original_grid=raw_original_grid,
        filtered_orig_indices=filtered_orig_indices,
        orig_header_idx=orig_header_idx,
    )

    report = ProcessingReport(
        raw_rows=raw_rows,
        raw_cols=raw_cols,
        after_deduplication_rows=after_dedup,
        after_filter_rows=after_filter,
        final_rows=len(filtered_rows),
        unique_units_count=len(ul_matr_pairs),
        unique_instalacoes_count=kpi_counts["unique_instalacoes_count"],
        rural_count=kpi_counts["rural_count"],
        povoado_count=kpi_counts["povoado_count"],
        urbano_count=kpi_counts["urbano_count"],
        warnings=list(dict.fromkeys(warnings)),
    )
    return out_bytes, report
