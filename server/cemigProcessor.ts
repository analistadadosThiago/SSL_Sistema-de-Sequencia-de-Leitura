import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { fillRGeralTemplateOOXML } from './rGeralTemplateFiller.ts';

export const DEFAULT_ALLOWED_LEITURISTAS: readonly string[] = [
  'Leiturista: 4MW',
  'Leiturista: 4NI',
  'Leiturista: 51O',
  'Leiturista: 56A',
  'Leiturista: 58A',
  'Leiturista: 5GL',
  'Leiturista: 5HL',
  'Leiturista: 5IL',
  'Leiturista: 5JL',
  'Leiturista: 5LL',
  'Leiturista: 5UL',
  'Leiturista: 5ZP',
  'Leiturista: 80J',
  'Leiturista: 80L',
  'Leiturista: 89O',
];

export const FINAL_HEADERS: readonly string[] = [
  'INSTALAÇÃO',
  'MEDIDOR',
  'REGISTRO',
  'RZ_UL_LV',
  'UL',
  'MATRÍCULA',
  'TIPO',
  'COD',
  'DATA',
  'HORA',
  'ENDEREÇO',
  'Leit. Anterior',
  'Leit. Executada',
  'CONSUMO',
  'NOSB´IMPEDIMENTO',
  'DIGITAÇÃO',
  'Apres.',
  'CONSISTÊNCIA',
  'NOSB´SIMULAÇÃO',
  'FOTOS',
];

/**
 * Regra da coluna UL (a partir de RZ_UL_LV):
 * - Extrair do 3º ao 6º caractere contando da esquerda para a direita:
 *   substring(RZ_UL_LV, 3, 4) em base 1 (índices 2 a 5 em base 0);
 * - Contabilizar zeros à esquerda normalmente (não remover zeros iniciais);
 * - Se RZ_UL_LV tiver menos de 6 caracteres, pegar do 3º caractere até o final;
 * - Se tiver menos de 3 caracteres, retornar "".
 */
export function extractULFromRzUlLv(rzUlLv: string): string {
  const clean = (rzUlLv ?? '').trim();
  if (clean.length < 3) return '';
  return clean.slice(2, 6);
}

/**
 * Classificação da coluna TIPO pelos 2 últimos dígitos de RZ_UL_LV:
 * - 00 a 79 -> "Urbano"
 * - 80 a 89 -> "Povoado"
 * - 90 a 99 -> "Rural"
 * - Vazio ou inválido -> ""
 */
export function classifyTipoFromRzUlLv(rzUlLv: string): string {
  const digits = rzUlLv.replace(/\D/g, '');
  if (digits.length < 2) return '';
  const lastTwo = Number(digits.slice(-2));
  if (!Number.isFinite(lastTwo)) return '';
  if (lastTwo <= 79) return 'Urbano';
  if (lastTwo <= 89) return 'Povoado';
  return 'Rural';
}

export interface ProcessOptions {
  allowedLeituristas?: string[];
  populatePlanilha3FromStage2?: boolean;
  consistenciaBuffer?: Buffer | null;
  nosbBuffer?: Buffer | null;
  fotosBuffer?: Buffer | null;
  modeloRGeralBuffer?: Buffer | null;
  modeloRGeralFileName?: string | null;
}

export function normalizeApresValue(raw: string): string {
  return (raw ?? '').trim() === '00' ? '00' : 'N00';
}

export interface StageStepLog {
  stage: 'Etapa 1' | 'Etapa 2' | 'Etapa 3';
  step: string;
  detail: string;
}

export interface ProcessingStats {
  sheetName: string;
  rawRows: number;
  rawCols: number;
  afterDeduplicationRows: number;
  afterFilterRows: number;
  finalRows: number;
  finalCols: number;
  uniqueUnitsCount: number;
  uniqueInstalacoesCount: number;
  ruralCount: number;
  povoadoCount: number;
  urbanoCount: number;
  planilha3Provided: boolean;
  planilha3Rows: number;
  leituristaBreakdown: Record<string, number>;
  foundLeituristasInFile: string[];
  durationMs: number;
}

export const RELATORIO_GERAL_HEADERS: readonly string[] = [
  'INSTALAÇÃO',
  'MEDIDOR',
  'REG.',
  'RZ UL LV',
  'UL',
  'Foto?',
  'Tipo',
  'MATR',
  'NL',
  'DATA',
  'HORA',
  'ENDEREÇO',
  'ANTERIOR',
  'ATUAL',
  'CONSUMO',
  'IMPEDIMENTO',
  'SIMULAÇÃO',
  'DIG',
  'CNA',
  'CONSISTENCIA',
];

export const CONTROLE_HORARIO_HEADERS: readonly string[] = [
  'Matricula',
  'Data',
  'Hora inicial',
  'Hora Final',
  'Tempo de Serviço',
];

export interface ControleHorarioRow {
  matricula: string;
  data: string;
  horaInicial: string;
  horaFinal: string;
  tempoServico: string;
}

export interface PreviewData {
  headers: string[];
  rows: string[][];
  totalDataRows: number;
}

export interface ProcessResult {
  outputBuffer: Buffer;
  fileName: string;
  relatorioGeralBuffer: Buffer | null;
  relatorioGeralFileName: string | null;
  stats: ProcessingStats;
  logs: StageStepLog[];
  warnings: string[];
  preview: PreviewData;
  controleHorario: ControleHorarioRow[];
  stage2Sample: Array<{ unidade: string; matricula: string }>;
}

export class CemigValidationError extends Error {
  code: 'VALIDATION' | 'CORRUPT_FILE';

  constructor(
    message: string,
    code: 'VALIDATION' | 'CORRUPT_FILE' = 'VALIDATION',
  ) {
    super(message);
    this.name = 'CemigValidationError';
    this.code = code;
  }
}

function normalizeLeituristaCode(value: string): string {
  return value
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/^leiturista\s*:\s*/i, 'Leiturista: ');
}

type Grid = string[][];

function normalizeText(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function isExcludedExtraColumn(headerName: string): boolean {
  const norm = normalizeText(headerName).replace(/\.+$/, '').trim();
  return (
    norm === 'c' ||
    norm === 'a exec' ||
    norm === 'exec' ||
    norm === 'leiturista'
  );
}

function ensureGridRect(grid: Grid, minCols: number): void {
  let maxCols = minCols;
  for (const row of grid) {
    if (row.length > maxCols) maxCols = row.length;
  }
  for (const row of grid) {
    while (row.length < maxCols) {
      row.push('');
    }
  }
}

function trimTrailingEmptyRows(grid: Grid): Grid {
  let lastNonEmpty = -1;
  for (let r = grid.length - 1; r >= 0; r--) {
    if (grid[r].some((cell) => cell !== '')) {
      lastNonEmpty = r;
      break;
    }
  }
  return lastNonEmpty >= 0 ? grid.slice(0, lastNonEmpty + 1) : [];
}

function getGridMaxCols(grid: Grid): number {
  let maxCols = 0;
  for (const row of grid) {
    for (let c = row.length - 1; c >= 0; c--) {
      if (row[c] !== '') {
        if (c + 1 > maxCols) maxCols = c + 1;
        break;
      }
    }
  }
  return maxCols;
}

/**
 * Reads the first worksheet (and optional Planilha3) from an .xls or .xlsx buffer,
 * unmerging all merged cells (keeping top-left value) and preserving leading zeros
 * and original date/time strings.
 */
function readWorkbookSheetAsGrid(
  buffer: Buffer,
  preferredSheetName?: string,
  preferFirstSheet = false,
): { sheetName: string; grid: Grid; planilha3Grid: Grid | null } {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, {
      type: 'buffer',
      raw: false,
      cellDates: false,
      cellNF: true,
      cellText: true,
    });
  } catch (err) {
    throw new CemigValidationError(
      `Não foi possível ler o arquivo enviado. Verifique se é uma planilha Excel válida (.xlsx ou .xls). Detalhe: ${String(err instanceof Error ? err.message : err)}`,
      'CORRUPT_FILE',
    );
  }

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new CemigValidationError(
      'O arquivo Excel enviado não contém nenhuma aba.',
      'CORRUPT_FILE',
    );
  }

  const targetName = preferFirstSheet
    ? workbook.SheetNames[0]
    : (preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
        ? preferredSheetName
        : workbook.SheetNames.find((n) => normalizeText(n) === 'leituras')) ||
      workbook.SheetNames[0];

  const extractSheet = (name: string): Grid => {
    const ws = workbook.Sheets[name];
    if (!ws) return [];

    // Sparse scan avoids OOM when Excel sets huge !ref (e.g. A1:XFD1048576)
    const cellsByRow = new Map<number, Map<number, string>>();
    let maxRow = -1;
    let maxCol = -1;

    for (const key of Object.keys(ws)) {
      if (key.startsWith('!')) continue;
      const cell = ws[key] as XLSX.CellObject | undefined;
      if (!cell) continue;

      let strVal = '';
      if (typeof cell.w === 'string' && cell.w.length > 0) {
        strVal = cell.w;
      } else if (cell.v !== undefined && cell.v !== null) {
        strVal = String(cell.v);
      }
      if (strVal === '') continue;

      const { r, c } = XLSX.utils.decode_cell(key);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
      let rowMap = cellsByRow.get(r);
      if (!rowMap) {
        rowMap = new Map<number, string>();
        cellsByRow.set(r, rowMap);
      }
      rowMap.set(c, strVal);
    }

    if (maxRow < 0 || maxCol < 0) return [];

    const numRows = maxRow + 1;
    const numCols = maxCol + 1;
    const grid: Grid = Array.from({ length: numRows }, () =>
      Array.from({ length: numCols }, () => ''),
    );

    for (const [r, rowMap] of cellsByRow.entries()) {
      for (const [c, val] of rowMap.entries()) {
        grid[r][c] = val;
      }
    }

    // Unmerge cells: Excel UnMerge keeps top-left cell value and clears remaining cells
    const merges = ws['!merges'] ?? [];
    for (const m of merges) {
      if (m.s.r >= numRows || m.s.c >= numCols) continue;
      const topLeftVal = grid[m.s.r]?.[m.s.c] ?? '';
      const endR = Math.min(m.e.r, numRows - 1);
      const endC = Math.min(m.e.c, numCols - 1);
      for (let r = m.s.r; r <= endR; r++) {
        for (let c = m.s.c; c <= endC; c++) {
          if (r === m.s.r && c === m.s.c) {
            grid[r][c] = topLeftVal;
          } else {
            grid[r][c] = '';
          }
        }
      }
    }

    return trimTrailingEmptyRows(grid);
  };

  const mainGrid = extractSheet(targetName);

  let planilha3Grid: Grid | null = null;
  const p3Name = workbook.SheetNames.find((s) => {
    const norm = normalizeText(s);
    return (norm === 'filtros' || norm === 'planilha3') && s !== targetName;
  });
  if (p3Name) {
    planilha3Grid = extractSheet(p3Name);
  }

  return { sheetName: targetName, grid: mainGrid, planilha3Grid };
}

function validateCemigLayout(grid: Grid): void {
  if (grid.length < 12) {
    throw new CemigValidationError(
      `Layout inválido: a planilha possui apenas ${grid.length} linha(s). O relatório bruto "Seqüência de Leitura Executada" da CEMIG exige cabeçalho e blocos de leitura (mínimo 12 linhas).`,
    );
  }

  const maxCols = getGridMaxCols(grid);
  if (maxCols < 10) {
    throw new CemigValidationError(
      `Layout inválido: apenas ${maxCols} coluna(s) detectada(s). São necessárias pelo menos 10 colunas do relatório CEMIG.`,
    );
  }

  const markers = new Set<string>();
  const sampleLimit = Math.min(grid.length, 250);
  for (let r = 0; r < sampleLimit; r++) {
    for (const cell of grid[r]) {
      if (!cell) continue;
      const norm = normalizeText(cell);
      if (norm.includes('cemig')) markers.add('cemig');
      if (norm.includes('sequencia de leitura')) markers.add('sequencia');
      if (norm.includes('unidade de leitura')) markers.add('unidade_leitura');
      if (norm.includes('leiturista:')) markers.add('leiturista');
      if (norm.includes('medidor')) markers.add('medidor');
      if (norm.includes('leit. ant') || norm.includes('leit. exec')) markers.add('leituras');
    }
  }

  if (markers.size < 2) {
    throw new CemigValidationError(
      'A planilha enviada não corresponde ao layout CEMIG "Seqüência de Leitura Executada". Não foram encontrados marcadores esperados.',
    );
  }
}

// ============================================================================
// Grid Manipulation Primitives (1-based Excel column/row semantics)
// ============================================================================

function deleteColumn(grid: Grid, col1: number): void {
  const idx = col1 - 1;
  for (const row of grid) {
    if (idx < row.length) {
      row.splice(idx, 1);
    }
  }
}

function deleteColumnsRange(grid: Grid, startCol1: number, count: number): void {
  const idx = startCol1 - 1;
  for (const row of grid) {
    if (idx < row.length) {
      row.splice(idx, count);
    }
  }
}

function insertColumnsAt(grid: Grid, col1: number, count: number): void {
  const idx = col1 - 1;
  const blanks = Array.from({ length: count }, () => '');
  for (const row of grid) {
    while (row.length < idx) row.push('');
    row.splice(idx, 0, ...blanks);
  }
}

function copyColumn(grid: Grid, srcCol1: number, dstCol1: number): void {
  ensureGridRect(grid, Math.max(srcCol1, dstCol1));
  const srcIdx = srcCol1 - 1;
  const dstIdx = dstCol1 - 1;
  for (const row of grid) {
    row[dstIdx] = row[srcIdx] ?? '';
  }
}

function deleteCellsShiftUp(
  grid: Grid,
  col1: number,
  startRow1: number,
  count: number,
): void {
  ensureGridRect(grid, col1);
  const cIdx = col1 - 1;
  const rIdx = startRow1 - 1;
  for (let r = rIdx; r < grid.length; r++) {
    const srcRow = r + count;
    grid[r][cIdx] = srcRow < grid.length ? grid[srcRow][cIdx] : '';
  }
}

function insertCellsShiftDown(
  grid: Grid,
  col1: number,
  startRow1: number,
  count: number,
): void {
  ensureGridRect(grid, col1);
  const cIdx = col1 - 1;
  const width = grid[0]?.length ?? col1;
  for (let i = 0; i < count; i++) {
    grid.push(Array.from({ length: width }, () => ''));
  }
  const rIdx = startRow1 - 1;
  for (let r = grid.length - 1; r >= rIdx + count; r--) {
    grid[r][cIdx] = grid[r - count][cIdx];
  }
  for (let r = rIdx; r < rIdx + count && r < grid.length; r++) {
    grid[r][cIdx] = '';
  }
}

function splitFixedWidthCol(str: string, breakPos: number): [string, string] {
  if (str.length <= breakPos) {
    return [str, ''];
  }
  return [str.slice(0, breakPos), str.slice(breakPos)];
}

function parseNumericMaybe(val: string): number | null {
  const trimmed = val.trim();
  if (!trimmed) return null;
  if (/^0\d+$/.test(trimmed)) return null;
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

function parseNumberForCalc(val: string): number | null {
  const trimmed = val.trim().replace(/\./g, '').replace(',', '.');
  if (!trimmed) return null;
  const num = Number(trimmed);
  return Number.isFinite(num) ? num : null;
}

function isPageNoiseRow(row: string[]): boolean {
  const joined = row.map((c) => normalizeText(c)).join(' ');
  if (!joined.trim()) return true;
  if (
    joined.includes('cemig distribuicao') ||
    /\bpag\.\s*\d+\s*\//.test(joined) ||
    joined.includes('unidade de leitura final')
  ) {
    return true;
  }
  return false;
}

// ============================================================================
// Date & Time Conversion and Sorting Helpers
// ============================================================================

function detectPreferMDY(rawDates: string[]): boolean {
  let hasP1Over12 = false;
  let hasP2Over12 = false;
  for (const raw of rawDates) {
    const m = raw.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (!m) continue;
    const p1 = Number(m[1]);
    const p2 = Number(m[2]);
    if (p1 > 12) hasP1Over12 = true;
    if (p2 > 12) hasP2Over12 = true;
  }
  if (hasP2Over12) return true;
  if (hasP1Over12) return false;
  // Padrão CEMIG bruto (ex.: 11/6/2026 -> 06/11/2026, 11/12/2026 -> 12/11/2026)
  return true;
}

function formatDateToDDMMYYYY(raw: string, preferMDY: boolean): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';

  // Serial numérico do Excel (ex.: 46343 -> 17/11/2026)
  if (/^\d{5}(\.\d+)?$/.test(trimmed)) {
    const serial = Math.floor(Number(trimmed));
    if (serial >= 25000 && serial <= 80000) {
      const utcMs = Date.UTC(1899, 11, 30) + serial * 86400_000;
      const dt = new Date(utcMs);
      const dd = String(dt.getUTCDate()).padStart(2, '0');
      const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
      const yyyy = String(dt.getUTCFullYear());
      return `${dd}/${mm}/${yyyy}`;
    }
  }

  // ISO YYYY-MM-DD
  const isoMatch = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (isoMatch) {
    const yyyy = isoMatch[1];
    const mm = isoMatch[2].padStart(2, '0');
    const dd = isoMatch[3].padStart(2, '0');
    return `${dd}/${mm}/${yyyy}`;
  }

  const slashMatch = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!slashMatch) return trimmed;

  const p1 = Number(slashMatch[1]);
  const p2 = Number(slashMatch[2]);
  const rawYear = slashMatch[3];
  const yyyy = rawYear.length === 2 ? `20${rawYear}` : rawYear.padStart(4, '0');

  let day: number;
  let month: number;

  if (p1 > 12) {
    day = p1;
    month = p2;
  } else if (p2 > 12) {
    month = p1;
    day = p2;
  } else if (preferMDY) {
    month = p1;
    day = p2;
  } else {
    day = p1;
    month = p2;
  }

  return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${yyyy}`;
}

function formatHoraHHMM(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  if (/^0?\.\d+$/.test(trimmed)) {
    const frac = Number(trimmed);
    if (Number.isFinite(frac) && frac >= 0 && frac < 1) {
      const totalMins = Math.round(frac * 24 * 60) % (24 * 60);
      const hh = String(Math.floor(totalMins / 60)).padStart(2, '0');
      const mm = String(totalMins % 60).padStart(2, '0');
      return `${hh}:${mm}`;
    }
  }
  return trimmed;
}

function dateDDMMYYYYToSortKey(ddmmyyyy: string): number {
  const m = ddmmyyyy.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return Number.MAX_SAFE_INTEGER;
  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);
  return yyyy * 10000 + mm * 100 + dd;
}

function timeToSortSeconds(rawTime: string): number {
  const trimmed = rawTime.trim().toUpperCase();
  if (!trimmed) return Number.MAX_SAFE_INTEGER;

  const m = trimmed.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/);
  if (!m) return Number.MAX_SAFE_INTEGER;

  let hours = Number(m[1]);
  const minutes = Number(m[2]);
  const seconds = m[3] ? Number(m[3]) : 0;
  const meridiem = m[4];

  if (meridiem === 'AM') {
    if (hours === 12) hours = 0;
  } else if (meridiem === 'PM') {
    if (hours < 12) hours += 12;
  }

  return hours * 3600 + minutes * 60 + seconds;
}

function parseTimeToMinutesHHMM(rawTime: string): number | null {
  const trimmed = (rawTime ?? '').trim().toUpperCase();
  if (!trimmed) return null;

  if (/^0?\.\d+$/.test(trimmed)) {
    const frac = Number(trimmed);
    if (Number.isFinite(frac) && frac >= 0 && frac < 1) {
      return Math.round(frac * 24 * 60) % (24 * 60);
    }
  }

  const m = trimmed.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/);
  if (!m) return null;

  let hours = Number(m[1]);
  const minutes = Number(m[2]);
  const meridiem = m[4];

  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || minutes < 0 || minutes > 59) {
    return null;
  }

  if (meridiem === 'AM') {
    if (hours === 12) hours = 0;
  } else if (meridiem === 'PM') {
    if (hours < 12) hours += 12;
  }

  if (hours < 0 || hours > 23) return null;
  return hours * 60 + minutes;
}

function formatMinutesToHHMM(totalMinutes: number): string {
  const safeMins = Math.max(0, Math.floor(totalMinutes));
  const hh = String(Math.floor(safeMins / 60)).padStart(2, '0');
  const mm = String(safeMins % 60).padStart(2, '0');
  return `${hh}:${mm}`;
}

export function buildControleHorarioRows(records: NormalizedRecord[]): ControleHorarioRow[] {
  const groups = new Map<
    string,
    {
      matricula: string;
      data: string;
      dateSortKey: number;
      minMinutes: number | null;
      maxMinutes: number | null;
    }
  >();

  for (const rec of records) {
    const matricula = (rec.matricula ?? '').trim() || 'Sem Leiturista';
    const data = (rec.data ?? '').trim();
    const key = `${matricula}\u0001${data}`;
    const mins = parseTimeToMinutesHHMM(rec.hora);

    let grp = groups.get(key);
    if (!grp) {
      grp = {
        matricula,
        data,
        dateSortKey: dateDDMMYYYYToSortKey(data),
        minMinutes: mins,
        maxMinutes: mins,
      };
      groups.set(key, grp);
    } else if (mins !== null) {
      if (grp.minMinutes === null || mins < grp.minMinutes) {
        grp.minMinutes = mins;
      }
      if (grp.maxMinutes === null || mins > grp.maxMinutes) {
        grp.maxMinutes = mins;
      }
    }
  }

  const rows: ControleHorarioRow[] = [];
  const sortedGroups = Array.from(groups.values()).sort((a, b) => {
    const matrCmp = a.matricula.localeCompare(b.matricula, 'pt-BR', {
      numeric: true,
      sensitivity: 'base',
    });
    if (matrCmp !== 0) return matrCmp;
    return a.dateSortKey - b.dateSortKey;
  });

  for (const grp of sortedGroups) {
    const hasTimes = grp.minMinutes !== null && grp.maxMinutes !== null;
    const horaInicial = hasTimes ? formatMinutesToHHMM(grp.minMinutes!) : '';
    const horaFinal = hasTimes ? formatMinutesToHHMM(grp.maxMinutes!) : '';
    const diff =
      hasTimes && grp.maxMinutes! > grp.minMinutes!
        ? grp.maxMinutes! - grp.minMinutes!
        : 0;
    const tempoServico = formatMinutesToHHMM(diff);

    rows.push({
      matricula: grp.matricula,
      data: grp.data,
      horaInicial,
      horaFinal,
      tempoServico,
    });
  }

  return rows;
}

interface NormalizedRecord {
  instalacao: string;
  medidor: string;
  registro: string;
  rzUlLv: string;
  ul: string;
  matricula: string;
  tipo: string;
  cod: string;
  data: string;
  hora: string;
  endereco: string;
  leitAntRaw: string;
  leitExecRaw: string;
  consumoCalc: number | '';
  impedimento: string;
  digitacao: string;
  apres: string;
  consistencia: string;
  simulacao: string;
  fotos: string;
  extraValues: string[];
  dateSortKey: number;
  timeSortKey: number;
  aExecSortKey?: number;
  execSortKey?: number;
  originalOrder: number;
}

function isAuxHeaderRow(colA: string, colB: string): boolean {
  const aTrim = colA.trim();
  const bTrim = colB.trim();
  const aNorm = normalizeText(aTrim);
  const bNorm = normalizeText(bTrim);
  if (!aNorm && !bNorm) return false;

  // IDs numéricos / instalações / códigos de medidor (ex.: "2.796.450.018-02", "APG176044248") nunca são cabeçalho
  if (/\d/.test(aTrim)) {
    return false;
  }

  const headerKeywords = [
    'instalacao',
    'instalacoes',
    'unidade',
    'chave',
    'codigo',
    'cod',
    'consistencia',
    'nosb',
    'simulacao',
    'fotos',
    'foto',
    'valor',
    'resultado',
    'status',
    'descricao',
    'situacao',
    'medidor',
    'cliente',
    'observacao',
  ];
  if (
    headerKeywords.some(
      (kw) => aNorm === kw || aNorm.includes(kw) || bNorm === kw || bNorm.includes(kw),
    )
  ) {
    return true;
  }

  const knownDataValues = new Set([
    'ok',
    'n-ok',
    'nok',
    'faturada',
    'nao faturada',
    'sim',
    'nao',
    'pendente',
    'divergente',
  ]);
  const aWords = aNorm.split(/\s+/).filter(Boolean);
  if (
    /[a-zA-Z]/.test(aNorm) &&
    !/\d/.test(aNorm) &&
    aWords.length <= 2 &&
    !knownDataValues.has(bNorm)
  ) {
    return true;
  }
  return false;
}

function readAuxPairColumns(
  buffer: Buffer,
  auxName: string,
): { colBHeader: string | null; rows: Array<[string, string]> } {
  const { grid } = readWorkbookSheetAsGrid(buffer, undefined, true);
  const nonEmptyRows = grid.filter(
    (r) => (r[0] ?? '').trim() !== '' || (r[1] ?? '').trim() !== '',
  );
  if (nonEmptyRows.length === 0) {
    return { colBHeader: null, rows: [] };
  }

  const firstA = (nonEmptyRows[0][0] ?? '').trim();
  const firstB = (nonEmptyRows[0][1] ?? '').trim();
  const hasHeader = isAuxHeaderRow(firstA, firstB);

  const startIdx = hasHeader ? 1 : 0;
  const colBHeader = hasHeader ? firstB || 'Valor' : 'Valor';
  const rows: Array<[string, string]> = [];
  for (let i = startIdx; i < nonEmptyRows.length; i++) {
    const a = (nonEmptyRows[i][0] ?? '').trim();
    const b = (nonEmptyRows[i][1] ?? '').trim();
    if (a !== '' || b !== '') {
      rows.push([a, b]);
    }
  }
  console.log(
    `[DIAG] Planilha auxiliar ${auxName}: ${rows.length} linhas carregadas (cabeçalho=${hasHeader ? 'sim' : 'não'})`,
  );
  return { colBHeader, rows };
}

function prepareFiltrosContext(
  planilha3Buffer: Buffer | null | undefined,
  embeddedP3: Grid | null,
  populateP3FromStage2: boolean,
  stage2Pairs: Array<{ unidade: string; matricula: string }>,
  warningsSet: Set<string>,
  consistenciaBuffer?: Buffer | null,
  nosbBuffer?: Buffer | null,
  fotosBuffer?: Buffer | null,
) {
  let p3Grid: Grid = [];
  let planilha3Provided = false;

  if (planilha3Buffer) {
    const parsedP3 = readWorkbookSheetAsGrid(planilha3Buffer, 'Filtros');
    p3Grid = parsedP3.grid;
    planilha3Provided = true;
  } else if (embeddedP3 && embeddedP3.length > 0) {
    p3Grid = embeddedP3;
    planilha3Provided = true;
  } else {
    if (populateP3FromStage2) {
      p3Grid = [
        ['RZ_UL_LV', 'MATRÍCULA', '', 'Consistencia', '', '', 'Nosb', '', '', 'Fotos', ''],
        ...stage2Pairs.map((pair) => [pair.unidade, pair.matricula]),
      ];
    } else {
      p3Grid = [
        ['RZ_UL_LV', 'MATRÍCULA', '', 'Consistencia', '', '', 'Nosb', '', '', 'Fotos', ''],
      ];
    }
    warningsSet.add('Tabela completa do tratamento.');
  }

  ensureGridRect(p3Grid, 11);
  if (p3Grid.length === 0) {
    p3Grid.push(Array.from({ length: 11 }, () => ''));
  }
  if (!p3Grid[0][3]?.trim()) p3Grid[0][3] = 'Consistencia';
  if (!p3Grid[0][6]?.trim()) p3Grid[0][6] = 'Nosb';
  if (!p3Grid[0][9]?.trim()) p3Grid[0][9] = 'Fotos';

  const writeAuxToFiltrosCols = (
    auxData: { colBHeader: string | null; rows: Array<[string, string]> },
    keyColIdx: number,
    valColIdx: number,
    fixedKeyHeader: string,
  ) => {
    p3Grid[0][keyColIdx] = fixedKeyHeader;
    if (auxData.colBHeader) {
      p3Grid[0][valColIdx] = auxData.colBHeader;
    }
    for (let r = 1; r < p3Grid.length; r++) {
      p3Grid[r][keyColIdx] = '';
      p3Grid[r][valColIdx] = '';
    }
    const requiredRows = 1 + auxData.rows.length;
    while (p3Grid.length < requiredRows) {
      p3Grid.push(Array.from({ length: Math.max(p3Grid[0].length, 11) }, () => ''));
    }
    for (let i = 0; i < auxData.rows.length; i++) {
      p3Grid[1 + i][keyColIdx] = auxData.rows[i][0];
      p3Grid[1 + i][valColIdx] = auxData.rows[i][1];
    }
  };

  if (consistenciaBuffer && consistenciaBuffer.length > 0) {
    const aux = readAuxPairColumns(consistenciaBuffer, 'Consistencia');
    writeAuxToFiltrosCols(aux, 3, 4, 'Consistencia');
    planilha3Provided = true;
  }
  if (nosbBuffer && nosbBuffer.length > 0) {
    const aux = readAuxPairColumns(nosbBuffer, 'Nosb');
    writeAuxToFiltrosCols(aux, 6, 7, 'Nosb');
    planilha3Provided = true;
  }
  if (fotosBuffer && fotosBuffer.length > 0) {
    const aux = readAuxPairColumns(fotosBuffer, 'Fotos');
    writeAuxToFiltrosCols(aux, 9, 10, 'Fotos');
    planilha3Provided = true;
  }

  const p3LookupAB = new Map<string, string>();
  const p3LookupDE = new Map<string, string>();
  const p3LookupGH = new Map<string, string>();
  const p3LookupJK = new Map<string, string>();

  for (let r = 0; r < p3Grid.length; r++) {
    const row = p3Grid[r];
    const kA = (row[0] ?? '').trim();
    if (kA && normalizeText(kA) !== 'rz_ul_lv' && !p3LookupAB.has(kA)) {
      p3LookupAB.set(kA, (row[1] ?? '').trim());
    }

    const kD = (row[3] ?? '').trim();
    if (kD && normalizeText(kD) !== 'consistencia' && !p3LookupDE.has(kD)) {
      p3LookupDE.set(kD, (row[4] ?? '').trim());
    }

    const kG = (row[6] ?? '').trim();
    if (kG && normalizeText(kG) !== 'nosb' && !p3LookupGH.has(kG)) {
      p3LookupGH.set(kG, (row[7] ?? '').trim());
    }

    const kJ = (row[9] ?? '').trim();
    if (kJ && normalizeText(kJ) !== 'fotos' && !p3LookupJK.has(kJ)) {
      p3LookupJK.set(kJ, (row[10] ?? '').trim());
    }
  }

  return { p3Grid, planilha3Provided, p3LookupAB, p3LookupDE, p3LookupGH, p3LookupJK };
}

async function loadTemplateWorkbook(modeloBuffer: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(new Uint8Array(modeloBuffer).buffer);
    if (wb.worksheets.length > 0) {
      return wb;
    }
  } catch {
    // Fallback para arquivos .xls (BIFF) via SheetJS
  }

  let rawWb: XLSX.WorkBook;
  try {
    rawWb = XLSX.read(modeloBuffer, {
      type: 'buffer',
      cellDates: false,
      cellNF: true,
    });
  } catch {
    throw new CemigValidationError(
      'Arquivo da Planilha Modelo RGeral corrompido ou formato não reconhecido.',
      'CORRUPT_FILE',
    );
  }

  const converted = new ExcelJS.Workbook();
  for (const name of rawWb.SheetNames) {
    const srcWs = rawWb.Sheets[name];
    const dstWs = converted.addWorksheet(name);
    if (!srcWs) continue;
    const range = srcWs['!ref'] ? XLSX.utils.decode_range(srcWs['!ref']) : null;
    if (range) {
      for (let r = range.s.r; r <= range.e.r; r++) {
        const dstRow = dstWs.getRow(r + 1);
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cellObj = srcWs[addr] as XLSX.CellObject | undefined;
          if (!cellObj) continue;
          const dstCell = dstRow.getCell(c + 1);
          if (cellObj.f) {
            dstCell.value = { formula: cellObj.f, result: cellObj.v as string | number | undefined };
          } else if (cellObj.v !== undefined) {
            dstCell.value = cellObj.v as string | number | boolean;
          }
          if (typeof cellObj.z === 'string') {
            dstCell.numFmt = cellObj.z;
          }
        }
      }
    }
    const colsMeta = srcWs['!cols'] as Array<{ wch?: number }> | undefined;
    if (Array.isArray(colsMeta)) {
      colsMeta.forEach((colInfo, idx) => {
        if (colInfo?.wch) {
          dstWs.getColumn(idx + 1).width = colInfo.wch;
        }
      });
    }
  }
  return converted;
}

async function fillRelatorioGeralTemplateWorkbook(
  modeloBuffer: Buffer,
  modeloFileName: string | null | undefined,
  records: NormalizedRecord[],
): Promise<{ buffer: Buffer; fileName: string }> {
  let effectiveBuffer = modeloBuffer;
  let effectiveName = modeloFileName || 'Relatorio Geral.xlsx';
  const isZipOoxml =
    modeloBuffer.length >= 2 && modeloBuffer[0] === 0x50 && modeloBuffer[1] === 0x4b;

  if (!isZipOoxml) {
    // Fallback para arquivos .xls (BIFF) antigos: converter para .xlsx antes de preencher via OOXML
    const convertedWb = await loadTemplateWorkbook(modeloBuffer);
    effectiveBuffer = Buffer.from(await convertedWb.xlsx.writeBuffer());
    effectiveName = effectiveName.replace(/\.xls$/i, '.xlsx');
  }

  return fillRGeralTemplateOOXML({
    modeloBuffer: effectiveBuffer,
    modeloFileName: effectiveName,
    records: records.map((rec) => {
      const antNumeric = parseNumericMaybe(rec.leitAntRaw);
      const execNumeric = parseNumericMaybe(rec.leitExecRaw);
      return {
        instalacao: rec.instalacao,
        medidor: rec.medidor,
        registro: rec.registro,
        rzUlLv: rec.rzUlLv,
        ul: rec.ul,
        fotos: rec.fotos,
        tipo: rec.tipo,
        matricula: rec.matricula,
        cod: rec.cod,
        data: rec.data,
        hora: rec.hora,
        endereco: rec.endereco,
        leitAnterior: antNumeric !== null ? antNumeric : rec.leitAntRaw,
        leitAtual: execNumeric !== null ? execNumeric : rec.leitExecRaw,
        consumo: typeof rec.consumoCalc === 'number' ? rec.consumoCalc : '',
        impedimento: rec.impedimento,
        simulacao: rec.simulacao,
        digitacao: rec.digitacao,
        apres: rec.apres,
        consistencia: rec.consistencia,
      };
    }),
  });
}

async function buildOutputWorkbookAndResult(params: {
  sheetName: string;
  rawRows: number;
  rawCols: number;
  afterDeduplicationRows: number;
  afterFilterRows: number;
  importedMetersCount: number;
  records: NormalizedRecord[];
  finalHeaders: string[];
  p3Grid: Grid;
  planilha3Provided: boolean;
  stage2Pairs: Array<{ unidade: string; matricula: string }>;
  leituristaBreakdown: Record<string, number>;
  foundLeituristasInFile: string[];
  logs: StageStepLog[];
  warningsSet: Set<string>;
  startTime: number;
  modeloRGeralBuffer?: Buffer | null;
  modeloRGeralFileName?: string | null;
}): Promise<ProcessResult> {
  const {
    sheetName,
    rawRows,
    rawCols,
    afterDeduplicationRows,
    afterFilterRows,
    importedMetersCount,
    records,
    finalHeaders,
    p3Grid,
    planilha3Provided,
    stage2Pairs,
    leituristaBreakdown,
    foundLeituristasInFile,
    logs,
    warningsSet,
    startTime,
    modeloRGeralBuffer,
    modeloRGeralFileName,
  } = params;

  let totalInstalacoesCount = 0;
  let ruralCount = 0;
  let povoadoCount = 0;
  let urbanoCount = 0;
  for (const rec of records) {
    if ((rec.instalacao ?? '').trim() !== '') totalInstalacoesCount++;
    if (rec.tipo === 'Rural') ruralCount++;
    else if (rec.tipo === 'Povoado') povoadoCount++;
    else if (rec.tipo === 'Urbano') urbanoCount++;
  }

  logs.push({
    stage: 'Etapa 3',
    step: `Tabela1 (${finalHeaders.length} Colunas), Datas dd/mm/aaaa e Ordenação`,
    detail: `${records.length} registros ordenados por Leiturista → DATA (dd/mm/aaaa) → HORA crescente, com Apres., UL e TIPO.`,
  });

  const outWb = new ExcelJS.Workbook();
  outWb.creator = 'Controle de Sequenciamento de Leitura';
  outWb.created = new Date();

  const wsFinal = outWb.addWorksheet('ControleLeitura', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  const wsFiltros = outWb.addWorksheet('Filtros');

  for (let r = 0; r < p3Grid.length; r++) {
    const outRow = wsFiltros.getRow(r + 1);
    for (let c = 0; c < p3Grid[r].length; c++) {
      const val = p3Grid[r][c] ?? '';
      if (val !== '') {
        const cell = outRow.getCell(c + 1);
        cell.value = val;
        cell.numFmt = '@';
      }
    }
  }

  wsFiltros.getCell('D1').value = wsFiltros.getCell('D1').value || 'Consistencia';
  wsFiltros.getCell('G1').value = wsFiltros.getCell('G1').value || 'Nosb';
  wsFiltros.getCell('J1').value = wsFiltros.getCell('J1').value || 'Fotos';

  const filtrosColFills: Record<number, string> = {
    4: 'FFDCE6F1',
    5: 'FFDCE6F1',
    7: 'FFE2EFDA',
    8: 'FFE2EFDA',
    10: 'FFFFF2CC',
    11: 'FFFFF2CC',
  };
  const filtrosStyledRows = Math.min(Math.max(p3Grid.length, 50), 20000);

  for (const colNum of [4, 5, 7, 8, 10, 11]) {
    const colObj = wsFiltros.getColumn(colNum);
    colObj.numFmt = '@';
    colObj.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: filtrosColFills[colNum] },
    };
  }

  for (let r = 1; r <= filtrosStyledRows; r++) {
    const rowObj = wsFiltros.getRow(r);
    for (const colNum of [4, 5, 7, 8, 10, 11]) {
      const cell = rowObj.getCell(colNum);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: filtrosColFills[colNum] },
      };
      if (r === 1) {
        cell.font = {
          name: 'Calibri',
          size: 11,
          bold: true,
          color: { argb: 'FF0F172A' },
        };
      }
      cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
    }
  }
  for (const colNum of [1, 2]) {
    const c1 = wsFiltros.getRow(1).getCell(colNum);
    if (c1.value) {
      c1.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    }
  }

  const vlookupEndRow = Math.max(50000, p3Grid.length);
  const allPreviewRows: string[][] = [];
  const tableRowsForExcelJS: Array<
    Array<string | number | { formula: string; result?: string | number }>
  > = [];
  const colMaxWidths = finalHeaders.map((h) => h.length + 4);

  for (let i = 0; i < records.length; i++) {
    const rec = records[i];
    const excelRowNum = i + 2;

    const antNumeric = parseNumericMaybe(rec.leitAntRaw);
    const execNumeric = parseNumericMaybe(rec.leitExecRaw);
    const consumoResult = typeof rec.consumoCalc === 'number' ? rec.consumoCalc : 0;

    const rowCellsForTable: Array<
      string | number | { formula: string; result?: string | number }
    > = [
      rec.instalacao,
      rec.medidor,
      rec.registro,
      rec.rzUlLv,
      rec.ul,
      {
        formula: `VLOOKUP(D${excelRowNum},Filtros!$A$1:$B$${vlookupEndRow},2,0)`,
        result: rec.matricula || undefined,
      },
      rec.tipo,
      rec.cod,
      rec.data,
      rec.hora,
      rec.endereco,
      antNumeric !== null ? antNumeric : rec.leitAntRaw,
      execNumeric !== null ? execNumeric : rec.leitExecRaw,
      {
        formula: `SUM(M${excelRowNum}-L${excelRowNum})`,
        result: consumoResult,
      },
      rec.impedimento,
      rec.digitacao,
      rec.apres,
      {
        formula: `IFERROR(VLOOKUP(B${excelRowNum},Filtros!$D$1:$E$${vlookupEndRow},2,0),"Faturada")`,
        result: rec.consistencia,
      },
      {
        formula: `IFERROR(VLOOKUP(B${excelRowNum},Filtros!$G$1:$H$${vlookupEndRow},2,0),"Faturada")`,
        result: rec.simulacao,
      },
      {
        formula: `IFERROR(VLOOKUP(B${excelRowNum},Filtros!$J$1:$K$${vlookupEndRow},2,0),"N-OK")`,
        result: rec.fotos,
      },
      ...rec.extraValues,
    ];

    tableRowsForExcelJS.push(rowCellsForTable);

    const previewRow = [
      rec.instalacao,
      rec.medidor,
      rec.registro,
      rec.rzUlLv,
      rec.ul,
      rec.matricula,
      rec.tipo,
      rec.cod,
      rec.data,
      rec.hora,
      rec.endereco,
      rec.leitAntRaw,
      rec.leitExecRaw,
      typeof rec.consumoCalc === 'number' ? String(rec.consumoCalc) : '',
      rec.impedimento,
      rec.digitacao,
      rec.apres,
      rec.consistencia,
      rec.simulacao,
      rec.fotos,
      ...rec.extraValues,
    ];
    allPreviewRows.push(previewRow);

    for (let c = 0; c < previewRow.length; c++) {
      const len = (previewRow[c]?.length ?? 0) + 2;
      if (len > colMaxWidths[c]) colMaxWidths[c] = len;
    }
  }

  console.log(
    `[DIAG] Total de linhas no DataFrame tratado (antes de gravar): ${records.length}`,
  );
  console.log(
    `[DIAG] Total de colunas no DataFrame tratado: ${finalHeaders.length} — colunas: ${JSON.stringify(finalHeaders)}`,
  );
  console.log(
    `[DIAG] Shape do DataFrame antes de gravar: ${records.length} x ${finalHeaders.length}`,
  );
  console.log(
    `[DIAG] Primeira linha: ${JSON.stringify(allPreviewRows[0] ?? [])}`,
  );
  console.log(
    `[DIAG] Última linha: ${JSON.stringify(allPreviewRows[allPreviewRows.length - 1] ?? [])}`,
  );

  wsFinal.addTable({
    name: 'Tabela1',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: finalHeaders.map((name) => ({
      name,
      filterButton: true,
    })),
    rows: tableRowsForExcelJS,
  });

  for (let c = 1; c <= finalHeaders.length; c++) {
    const colObj = wsFinal.getColumn(c);
    colObj.width = Math.min(Math.max(colMaxWidths[c - 1], 11), 46);
    colObj.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
    if (c === 1 || c === 2 || c === 3 || c === 4 || c === 5 || c === 9 || c === 17) {
      colObj.numFmt = '@';
    }
  }

  const outHeaderRow = wsFinal.getRow(1);
  for (let c = 1; c <= finalHeaders.length; c++) {
    const cell = outHeaderRow.getCell(c);
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: c >= 18 && c <= 20 ? 'FF0F766E' : c > 20 ? 'FF1E293B' : 'FF0F172A',
      },
    };
    cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
  }

  // Preservar formatação por célula para arquivos até 15.000 linhas e zeros à esquerda sempre
  const totalSheetRows = records.length + 1;
  const applyPerCellAlignment = records.length <= 15000;
  for (let r = 2; r <= totalSheetRows; r++) {
    const excelRow = wsFinal.getRow(r);
    for (let c = 1; c <= finalHeaders.length; c++) {
      const cell = excelRow.getCell(c);
      if (applyPerCellAlignment) {
        cell.alignment = {
          horizontal: 'left',
          vertical: 'bottom',
          wrapText: false,
        };
      }
      if (typeof cell.value === 'string') {
        if (/^0\d+$/.test(cell.value) || c === 5 || c === 9) {
          cell.numFmt = '@';
        }
      }
    }
  }

  wsFiltros.getColumn(1).width = 18;
  wsFiltros.getColumn(2).width = 14;
  wsFiltros.getColumn(4).width = 18;
  wsFiltros.getColumn(5).width = 18;
  wsFiltros.getColumn(7).width = 18;
  wsFiltros.getColumn(8).width = 18;
  wsFiltros.getColumn(10).width = 18;
  wsFiltros.getColumn(11).width = 18;

  const controleHorario = buildControleHorarioRows(records);
  const wsControleHorario = outWb.addWorksheet('Controle Horario', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  const chHeaderRow = wsControleHorario.getRow(1);
  for (let c = 1; c <= CONTROLE_HORARIO_HEADERS.length; c++) {
    const cell = chHeaderRow.getCell(c);
    cell.value = CONTROLE_HORARIO_HEADERS[c - 1];
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
    cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
  }

  for (let i = 0; i < controleHorario.length; i++) {
    const ch = controleHorario[i];
    const rowObj = wsControleHorario.getRow(i + 2);
    const values = [
      ch.matricula,
      ch.data,
      ch.horaInicial,
      ch.horaFinal,
      ch.tempoServico,
    ];
    for (let c = 1; c <= values.length; c++) {
      const cell = rowObj.getCell(c);
      cell.value = values[c - 1];
      cell.numFmt = '@';
      cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
    }
  }

  wsControleHorario.autoFilter = 'A1:E1';
  wsControleHorario.getColumn(1).width = 16;
  wsControleHorario.getColumn(2).width = 15;
  wsControleHorario.getColumn(3).width = 15;
  wsControleHorario.getColumn(4).width = 15;
  wsControleHorario.getColumn(5).width = 20;
  for (let c = 1; c <= 5; c++) {
    wsControleHorario.getColumn(c).numFmt = '@';
  }

  const outArrayBuffer = await outWb.xlsx.writeBuffer();
  const outputBuffer = Buffer.from(outArrayBuffer);

  const relatorioGeralResult =
    modeloRGeralBuffer && modeloRGeralBuffer.length > 0
      ? await fillRelatorioGeralTemplateWorkbook(
          modeloRGeralBuffer,
          modeloRGeralFileName,
          records,
        )
      : null;
  const relatorioGeralBuffer = relatorioGeralResult?.buffer ?? null;
  const relatorioGeralFileName = relatorioGeralResult?.fileName ?? null;

  // Verificação pós-gravação com baixo consumo de memória (XLSX.read apenas na aba ControleLeitura)
  const verifyWb = XLSX.read(outputBuffer, {
    type: 'buffer',
    sheets: 'ControleLeitura',
  });
  const verifyWs = verifyWb.Sheets['ControleLeitura'];
  let xlsxDataRows = 0;
  let xlsxCols = 0;
  const xlsxHeaders: string[] = [];

  if (verifyWs) {
    const nonEmptyRows = new Set<number>();
    const headerCells = new Map<number, string>();
    for (const key of Object.keys(verifyWs)) {
      if (key.startsWith('!')) continue;
      const cell = verifyWs[key] as XLSX.CellObject | undefined;
      if (!cell || (cell.v === undefined && !cell.f && !cell.w)) continue;
      const { r, c } = XLSX.utils.decode_cell(key);
      if (r === 0) {
        headerCells.set(c, String(cell.w ?? cell.v ?? ''));
        if (c + 1 > xlsxCols) xlsxCols = c + 1;
      } else {
        nonEmptyRows.add(r);
      }
    }
    xlsxDataRows = nonEmptyRows.size;
    for (let c = 0; c < xlsxCols; c++) {
      xlsxHeaders.push(headerCells.get(c) ?? '');
    }
  }

  const matchPostWrite =
    xlsxDataRows === records.length && xlsxCols === finalHeaders.length;
  const matchMeters =
    importedMetersCount === records.length && records.length === xlsxDataRows;

  console.log(`[DIAG] Total de linhas efetivamente gravadas no .xlsx: ${xlsxDataRows}`);
  console.log(
    `[DIAG] Total de colunas efetivamente gravadas: ${xlsxCols} — colunas: ${JSON.stringify(xlsxHeaders)}`,
  );
  console.log(`[DIAG] Shape do .xlsx após gravação: ${xlsxDataRows} x ${xlsxCols}`);
  console.log(`[DIAG] Tamanho do .xlsx gerado: ${outputBuffer.length} bytes`);
  console.log(
    `[DIAG] Verificação pós-gravação: .xlsx tem ${xlsxDataRows} linhas; DataFrame tem ${records.length} linhas. Match: ${matchPostWrite ? 'sim' : 'não'}`,
  );
  console.log(`[DIAG] Total de medidores no arquivo importado: ${importedMetersCount}`);
  console.log(`[DIAG] Total de medidores na tabela final: ${records.length}`);
  console.log(`[DIAG] Total de medidores no .xlsx exportado: ${xlsxDataRows}`);
  console.log(
    `[DIAG] Match (importado = final = exportado): ${matchMeters ? 'sim' : 'não'}`,
  );

  if (!matchMeters || !matchPostWrite) {
    console.error(
      `[DIAG] TRUNCAMENTO/OMISSÃO DETECTADO: importado=${importedMetersCount}, final=${records.length}, exportado=${xlsxDataRows}, diferença=${importedMetersCount - xlsxDataRows}`,
    );
  }

  logs.push({
    stage: 'Etapa 3',
    step: 'Verificação pós-gravação',
    detail: `Processamento concluído: ${records.length} linhas; gravadas ${xlsxDataRows} linhas no .xlsx (${xlsxCols} colunas).`,
  });

  const durationMs = Date.now() - startTime;

  return {
    outputBuffer,
    fileName: 'Matriz_ControleDeSequenciaLeitura.xlsx',
    relatorioGeralBuffer,
    relatorioGeralFileName,
    stats: {
      sheetName,
      rawRows,
      rawCols,
      afterDeduplicationRows,
      afterFilterRows,
      finalRows: records.length,
      finalCols: finalHeaders.length,
      uniqueUnitsCount: stage2Pairs.length,
      uniqueInstalacoesCount: totalInstalacoesCount,
      ruralCount,
      povoadoCount,
      urbanoCount,
      planilha3Provided,
      planilha3Rows: p3Grid.length,
      leituristaBreakdown,
      foundLeituristasInFile,
      durationMs,
    },
    logs,
    warnings: Array.from(warningsSet),
    preview: {
      headers: finalHeaders,
      rows: allPreviewRows,
      totalDataRows: records.length,
    },
    controleHorario,
    stage2Sample: stage2Pairs.slice(0, 50),
  };
}

// ============================================================================
// Detecção e Parser do Molde Leitura.xlsx (Bloco de 3 linhas por medidor)
// ============================================================================

function isLeitura3LineLayout(grid: Grid): boolean {
  const sampleLimit = Math.min(grid.length, 300);
  for (let r = 0; r < sampleLimit; r++) {
    const row = grid[r];
    if (!row || row.length < 4) continue;
    const normA = normalizeText(row[0] ?? '');
    const normB = normalizeText(row[1] ?? '');
    const normC = normalizeText(row[2] ?? '');
    const normD = normalizeText(row[3] ?? '');
    const normJoined = row.map((c) => normalizeText(c)).join(' | ');

    if (
      normA === 'unidade' &&
      normB === 'medidor' &&
      (normC.includes('t. reg') || normC.includes('t.reg') || normC.includes('t reg'))
    ) {
      return true;
    }
    if (
      normB === 'medidor' &&
      normD === 'oc' &&
      (normJoined.includes('a exec') || normJoined.includes('exec'))
    ) {
      return true;
    }
    if (
      normA === 'cliente' &&
      normJoined.includes('endereco') &&
      (normJoined.includes('qtd dig') || normJoined.includes('qtd. dig'))
    ) {
      return true;
    }
  }
  return false;
}

interface RawLeituraBlock {
  instalacao: string;
  medidor: string;
  registro: string;
  cod: string;
  rawData: string;
  hora: string;
  aExec: string;
  exec: string;
  leitAntRaw: string;
  leitExecRaw: string;
  impRaw: string;
  apres: string;
  rzUlLv: string;
  leiturista: string;
  auxFullTime: string;
  endereco: string;
  qtdDig: string;
  consumidora: string;
  cliente: string;
  originalOrder: number;
}

function isTimeLikeValue(val: string): boolean {
  return /^\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)?$/i.test(val.trim());
}

function isCpfCnpjOrNumericId(val: string): boolean {
  const clean = val.trim();
  if (!clean) return false;
  if (/[a-zA-Z]/.test(clean)) return false;
  const digits = clean.replace(/\D/g, '');
  return digits.length >= 6 && /^[\d.\-\/\s]+$/.test(clean);
}

async function processLeitura3LineLayout(params: {
  sheetName: string;
  grid: Grid;
  rawRows: number;
  rawCols: number;
  embeddedP3: Grid | null;
  planilha3Buffer: Buffer | null | undefined;
  consistenciaBuffer?: Buffer | null;
  nosbBuffer?: Buffer | null;
  fotosBuffer?: Buffer | null;
  modeloRGeralBuffer?: Buffer | null;
  modeloRGeralFileName?: string | null;
  populateP3FromStage2: boolean;
  logs: StageStepLog[];
  warningsSet: Set<string>;
  startTime: number;
}): Promise<ProcessResult> {
  const {
    sheetName,
    grid,
    rawRows,
    rawCols,
    embeddedP3,
    planilha3Buffer,
    consistenciaBuffer,
    nosbBuffer,
    fotosBuffer,
    modeloRGeralBuffer,
    modeloRGeralFileName,
    populateP3FromStage2,
    logs,
    warningsSet,
    startTime,
  } = params;

  console.log(`[DIAG] Linhas lidas no arquivo bruto: ${rawRows}`);

  let headerRowsIgnored = 0;
  let auxRowsMerged = 0;
  let currentRzUlLv = '';
  let currentLeiturista = '';
  const blocks: RawLeituraBlock[] = [];
  let activeBlock: RawLeituraBlock | null = null;
  const foundLeituristasSet = new Set<string>();

  const flushActiveBlock = () => {
    if (activeBlock) {
      blocks.push(activeBlock);
      activeBlock = null;
    }
  };

  const extractPageContextFromRow = (row: string[]) => {
    for (let c = 0; c < row.length; c++) {
      const cell = (row[c] ?? '').trim();
      if (!cell) continue;
      const norm = normalizeText(cell);

      if (
        norm.startsWith('unidade de leitura') &&
        !norm.includes('inicial') &&
        !norm.includes('final')
      ) {
        const afterColon = cell.includes(':') ? cell.split(':').slice(1).join(':').trim() : '';
        if (afterColon) {
          currentRzUlLv = afterColon;
        } else {
          for (let k = c + 1; k < row.length; k++) {
            const nextVal = (row[k] ?? '').trim();
            if (nextVal && !normalizeText(nextVal).includes('data')) {
              currentRzUlLv = nextVal;
              break;
            }
          }
        }
      }

      if (norm.startsWith('leiturista')) {
        const afterColon = cell.includes(':') ? cell.split(':').slice(1).join(':').trim() : '';
        let code = afterColon;
        if (!code) {
          for (let k = c + 1; k < row.length; k++) {
            const nextVal = (row[k] ?? '').trim();
            if (nextVal && !normalizeText(nextVal).includes('tipo')) {
              code = nextVal;
              break;
            }
          }
        }
        if (code) {
          currentLeiturista = code;
          foundLeituristasSet.add(`Leiturista: ${code}`);
        }
      }
    }
  };

  const isHeaderOrBannerRow = (row: string[]): boolean => {
    const nonEmptyCount = row.reduce((acc, c) => (c.trim() ? acc + 1 : acc), 0);
    if (nonEmptyCount === 0) return false;

    const normA = normalizeText(row[0] ?? '');
    const normB = normalizeText(row[1] ?? '');
    const normC = normalizeText(row[2] ?? '');
    const normJoined = row.map((c) => normalizeText(c)).join(' ');

    if (
      normA.includes('cemig distribuicao') ||
      normA === 'sequencia de leitura executada' ||
      /\bpag\.\s*\d+\s*\//.test(normJoined) ||
      normJoined.includes('unidade de leitura final') ||
      normJoined.includes('unidade de leitura inicial') ||
      normA.startsWith('unidade de leitura') ||
      normJoined.includes('unidade de leitura:') ||
      normA.startsWith('leiturista:') ||
      normJoined.includes('leiturista:') ||
      normA === 'leiturista' ||
      normJoined.includes('tipo de leitura:') ||
      normA.startsWith('data inicio') ||
      normJoined.includes('data termino') ||
      (normA === 'contrato' && normJoined.includes('mes/ano')) ||
      /^[-=_]{5,}$/.test(normA)
    ) {
      return true;
    }

    // Linha de valores do cabeçalho do contrato: "<contrato> | <MM/YYYY> | ... | <UL_ini> | ... | <UL_fim>"
    if (/^\d{8,}$/.test(normA) && /^\d{1,2}\/\d{4}$/.test(normB) && !normC) {
      return true;
    }

    // Cabeçalho da tabela (3 linhas de rótulos de coluna)
    if (normA === 'unidade' && normB === 'medidor') return true;
    if (normB === 'medidor' && (normC.includes('t. reg') || normC.includes('t.reg') || normC.includes('t reg'))) {
      return true;
    }
    if (normA === 'consumidora' && !normB && !normC) return true;
    if (
      normA === 'cliente' &&
      !normB &&
      !normC &&
      (normJoined.includes('endereco') || normJoined.includes('qtd dig') || normJoined.includes('qtd. dig'))
    ) {
      return true;
    }

    return false;
  };

  // Iteração completa de TODAS as linhas do arquivo (sem limite ou passo fixo)
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.every((c) => !c.trim())) {
      continue;
    }

    if (isHeaderOrBannerRow(row)) {
      flushActiveBlock();
      extractPageContextFromRow(row);
      headerRowsIgnored++;
      continue;
    }

    // Também checar contexto caso apareça sem marcador padrão
    extractPageContextFromRow(row);

    const colA = (row[0] ?? '').trim();
    const colB = (row[1] ?? '').trim();
    const colC = (row[2] ?? '').trim();

    // Linha principal do medidor: Medidor em B não vazio E T.Reg em C preenchido
    if (colB !== '' && colC !== '') {
      flushActiveBlock();
      activeBlock = {
        instalacao: colA,
        medidor: colB,
        registro: colC,
        cod: (row[3] ?? '').trim(),
        rawData: (row[5] ?? '').trim(),
        hora: formatHoraHHMM((row[6] ?? '').trim()),
        aExec: (row[7] ?? '').trim(),
        exec: (row[8] ?? '').trim(),
        leitAntRaw: (row[9] ?? '').trim(),
        leitExecRaw: (row[10] ?? '').trim(),
        impRaw: (row[11] ?? '').trim(),
        apres: (row[12] ?? '').trim(),
        rzUlLv: currentRzUlLv,
        leiturista: currentLeiturista,
        auxFullTime: '',
        endereco: '',
        qtdDig: '',
        consumidora: '',
        cliente: '',
        originalOrder: blocks.length + 1,
      };
      continue;
    }

    // Linha auxiliar pertencente ao medidor ativo (Linha 2, Linha 3, ...)
    if (activeBlock) {
      auxRowsMerged++;

      // Hora completa (col F=5 ou col G=6)
      for (const tIdx of [5, 6, 4, 7]) {
        const candidate = (row[tIdx] ?? '').trim();
        if (candidate && isTimeLikeValue(candidate)) {
          if (!activeBlock.auxFullTime) {
            activeBlock.auxFullTime = candidate;
          }
          break;
        }
      }

      // Qtd Dig. (col M=12 ou N=13, ignorando o "00" da linha principal)
      let foundDig = '';
      for (const dIdx of [12, 13, 14]) {
        const candidate = (row[dIdx] ?? '').trim();
        if (candidate !== '') {
          foundDig = candidate;
          break;
        }
      }
      if (foundDig !== '') {
        activeBlock.qtdDig = foundDig;
      }

      // Endereço (col G=6 ou col H=7)
      for (const eIdx of [6, 7]) {
        const candidate = (row[eIdx] ?? '').trim();
        if (candidate && !isTimeLikeValue(candidate)) {
          if (candidate.length > activeBlock.endereco.length) {
            activeBlock.endereco = candidate;
          }
        }
      }

      // Cliente (CPF/CNPJ na col D=3 ou col A=0) e Consumidora (nome na col A=0)
      const colD = (row[3] ?? '').trim();
      if (colD !== '' && !activeBlock.cliente) {
        activeBlock.cliente = colD;
      }

      if (colA !== '') {
        if (isCpfCnpjOrNumericId(colA)) {
          if (!activeBlock.cliente) activeBlock.cliente = colA;
        } else if (!activeBlock.consumidora) {
          activeBlock.consumidora = colA;
        } else if (!activeBlock.cliente) {
          activeBlock.cliente = colA;
        }
      }
    }
  }

  flushActiveBlock();

  const detectedMetersCount = blocks.length;
  console.log(`[DIAG] Linhas de cabeçalho ignoradas: ${headerRowsIgnored}`);
  console.log(`[DIAG] Medidores detectados (linhas principais): ${detectedMetersCount}`);
  console.log(`[DIAG] Linhas auxiliares mescladas: ${auxRowsMerged}`);

  if (detectedMetersCount === 0) {
    throw new CemigValidationError(
      'Nenhum medidor foi detectado no relatório CEMIG (nenhuma linha principal com Medidor na coluna B e T.Reg na coluna C).',
    );
  }

  const ulMatrMap = new Map<string, string>();
  const stage2Pairs: Array<{ unidade: string; matricula: string }> = [];
  for (const blk of blocks) {
    const unid = blk.rzUlLv.trim();
    const matr = (blk.leiturista || unid).trim();
    if (unid && !ulMatrMap.has(unid)) {
      ulMatrMap.set(unid, matr);
      stage2Pairs.push({ unidade: unid, matricula: matr });
    }
  }
  console.log(`[DIAG] Etapa 2 concluída: ${stage2Pairs.length} linhas`);

  const {
    p3Grid,
    planilha3Provided,
    p3LookupAB,
    p3LookupDE,
    p3LookupGH,
    p3LookupJK,
  } = prepareFiltrosContext(
    planilha3Buffer,
    embeddedP3,
    populateP3FromStage2,
    stage2Pairs,
    warningsSet,
    consistenciaBuffer,
    nosbBuffer,
    fotosBuffer,
  );

  const preferMDY = detectPreferMDY(
    blocks.map((b) => b.rawData).filter(Boolean),
  );

  let emptyEnderecoCount = 0;
  let emptyQtdDigCount = 0;
  let emptyConsumidoraCount = 0;
  let emptyClienteCount = 0;

  const finalHeaders: string[] = [...FINAL_HEADERS];
  const records: NormalizedRecord[] = [];
  const leituristaBreakdown: Record<string, number> = {};

  for (const blk of blocks) {
    if (!blk.endereco) emptyEnderecoCount++;
    if (!blk.qtdDig) emptyQtdDigCount++;
    if (!blk.consumidora) emptyConsumidoraCount++;
    if (!blk.cliente) emptyClienteCount++;

    const rzUlLv = blk.rzUlLv.trim();
    const ul = extractULFromRzUlLv(rzUlLv);
    const matricula = p3LookupAB.get(rzUlLv) || blk.leiturista || rzUlLv;
    const tipo = classifyTipoFromRzUlLv(rzUlLv);
    const data = formatDateToDDMMYYYY(blk.rawData, preferMDY);
    const hora = blk.hora || formatHoraHHMM(blk.auxFullTime);

    const antNum = parseNumberForCalc(blk.leitAntRaw);
    const execNum = parseNumberForCalc(blk.leitExecRaw);
    const consumoCalc =
      antNum !== null && execNum !== null ? execNum - antNum : '';

    const impedimento = blk.impRaw || 'Faturada';
    const digitacao = blk.qtdDig || 'N00';
    const apres = normalizeApresValue(blk.apres);
    const consistencia = p3LookupDE.get(blk.medidor) || 'Faturada';
    const simulacao = p3LookupGH.get(blk.medidor) || 'Faturada';
    const fotos = p3LookupJK.get(blk.medidor) || 'N-OK';

    const shortCode = matricula || 'Sem Leiturista';
    leituristaBreakdown[shortCode] = (leituristaBreakdown[shortCode] ?? 0) + 1;

    const aExecNum = Number(blk.aExec);
    const execOrderNum = Number(blk.exec);

    records.push({
      instalacao: blk.instalacao,
      medidor: blk.medidor,
      registro: blk.registro,
      rzUlLv,
      ul,
      matricula,
      tipo,
      cod: blk.cod,
      data,
      hora,
      endereco: blk.endereco,
      leitAntRaw: blk.leitAntRaw,
      leitExecRaw: blk.leitExecRaw,
      consumoCalc,
      impedimento,
      digitacao,
      apres,
      consistencia,
      simulacao,
      fotos,
      extraValues: [],
      dateSortKey: dateDDMMYYYYToSortKey(data),
      timeSortKey: timeToSortSeconds(blk.auxFullTime || hora),
      aExecSortKey: Number.isFinite(aExecNum) ? aExecNum : Number.MAX_SAFE_INTEGER,
      execSortKey: Number.isFinite(execOrderNum) ? execOrderNum : Number.MAX_SAFE_INTEGER,
      originalOrder: blk.originalOrder,
    });
  }

  console.log(`[DIAG] Linhas com Endereço vazio: ${emptyEnderecoCount} (deve ser 0)`);
  console.log(`[DIAG] Linhas com Qtd Dig. vazio: ${emptyQtdDigCount} (deve ser 0)`);
  console.log(`[DIAG] Linhas com Consumidora vazio: ${emptyConsumidoraCount}`);
  console.log(`[DIAG] Linhas com Cliente vazio: ${emptyClienteCount}`);

  // Ordenação por Leiturista (MATRÍCULA) -> DATA crescente -> HORA crescente
  records.sort((a, b) => {
    const leitCmp = a.matricula.localeCompare(b.matricula, 'pt-BR', {
      sensitivity: 'base',
    });
    if (leitCmp !== 0) return leitCmp;
    if (a.dateSortKey !== b.dateSortKey) return a.dateSortKey - b.dateSortKey;
    if (a.timeSortKey !== b.timeSortKey) return a.timeSortKey - b.timeSortKey;
    if ((a.aExecSortKey ?? 0) !== (b.aExecSortKey ?? 0)) {
      return (a.aExecSortKey ?? 0) - (b.aExecSortKey ?? 0);
    }
    if ((a.execSortKey ?? 0) !== (b.execSortKey ?? 0)) {
      return (a.execSortKey ?? 0) - (b.execSortKey ?? 0);
    }
    return a.originalOrder - b.originalOrder;
  });

  console.log(`[DIAG] Etapa 1 concluída: ${detectedMetersCount} linhas`);
  console.log(`[DIAG] Etapa 3 concluída: ${records.length} linhas`);
  console.log(
    `[DIAG] Tabela final montada: ${records.length} linhas x ${finalHeaders.length} colunas`,
  );

  logs.push({
    stage: 'Etapa 1',
    step: 'Agrupamento 3 Linhas por Medidor',
    detail: `${detectedMetersCount} medidores consolidados (${auxRowsMerged} linhas auxiliares mescladas, ${headerRowsIgnored} linhas de cabeçalho ignoradas).`,
  });
  logs.push({
    stage: 'Etapa 2',
    step: 'ParteDois_UlMatr',
    detail: `${stage2Pairs.length} pares únicos de Unidade de Leitura ↔ Matrícula gerados.`,
  });

  return buildOutputWorkbookAndResult({
    sheetName,
    rawRows,
    rawCols,
    afterDeduplicationRows: detectedMetersCount,
    afterFilterRows: detectedMetersCount,
    importedMetersCount: detectedMetersCount,
    records,
    finalHeaders,
    p3Grid,
    planilha3Provided,
    stage2Pairs,
    leituristaBreakdown,
    foundLeituristasInFile: Array.from(foundLeituristasSet).sort(),
    logs,
    warningsSet,
    startTime,
    modeloRGeralBuffer,
    modeloRGeralFileName,
  });
}

// ============================================================================
// Main 3-Stage Pipeline
// ============================================================================

export async function processCemigWorkbook(
  matrizBuffer: Buffer,
  planilha3Buffer?: Buffer | null,
  options: ProcessOptions = {},
): Promise<ProcessResult> {
  const startTime = Date.now();
  const logs: StageStepLog[] = [];
  const warningsSet = new Set<string>();

  const allowedLeituristas =
    options.allowedLeituristas && options.allowedLeituristas.length > 0
      ? options.allowedLeituristas
      : [...DEFAULT_ALLOWED_LEITURISTAS];

  const allowedSet = new Set(
    allowedLeituristas.map((s) => normalizeLeituristaCode(s)),
  );
  const populateP3FromStage2 = options.populatePlanilha3FromStage2 ?? true;

  // Read Matriz workbook
  const { sheetName, grid, planilha3Grid: embeddedP3 } =
    readWorkbookSheetAsGrid(matrizBuffer);

  validateCemigLayout(grid);

  const rawRows = grid.length;
  const rawCols = getGridMaxCols(grid);
  ensureGridRect(grid, Math.max(rawCols, 16));
  const rawOriginalGrid: Grid = grid.map((r) => [...r]);

  logs.push({
    stage: 'Etapa 1',
    step: 'Leitura e Desmesclagem',
    detail: `Aba "${sheetName}" lida com ${rawRows} linhas e ${rawCols} colunas preservando zeros à esquerda e datas.`,
  });

  if (isLeitura3LineLayout(grid)) {
    return processLeitura3LineLayout({
      sheetName,
      grid,
      rawRows,
      rawCols,
      embeddedP3,
      planilha3Buffer,
      consistenciaBuffer: options.consistenciaBuffer,
      nosbBuffer: options.nosbBuffer,
      fotosBuffer: options.fotosBuffer,
      modeloRGeralBuffer: options.modeloRGeralBuffer,
      modeloRGeralFileName: options.modeloRGeralFileName,
      populateP3FromStage2,
      logs,
      warningsSet,
      startTime,
    });
  }

  // --------------------------------------------------------------------------
  // ETAPA 1 — ParteUm_Tratamento (Ctrl+Shift+F)
  // --------------------------------------------------------------------------
  deleteColumn(grid, 5);
  deleteColumnsRange(grid, 7, 2);
  copyColumn(grid, 10, 11);
  deleteCellsShiftUp(grid, 10, 3, 2);
  insertColumnsAt(grid, 6, 1);
  copyColumn(grid, 7, 6);
  deleteCellsShiftUp(grid, 7, 5, 2);
  insertColumnsAt(grid, 4, 2);
  copyColumn(grid, 1, 4);
  copyColumn(grid, 1, 5);
  insertCellsShiftDown(grid, 4, 1, 5);
  insertCellsShiftDown(grid, 5, 2, 4);

  logs.push({
    stage: 'Etapa 1',
    step: 'Normalização de Colunas e Alinhamento Vertical (Passos 1–8)',
    detail:
      'Normalização estrutural aplicada com alinhamento de Unidade de Leitura, Leiturista e campos da linha principal.',
  });

  // (9) Remover duplicatas considerando todas as 16+ colunas (Header:=xlNo)
  const dedupCols = Math.max(16, getGridMaxCols(grid));
  ensureGridRect(grid, dedupCols);
  const seenRows = new Set<string>();
  const dedupGrid: Grid = [];
  const dedupOrigIndices: number[] = [];
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    const key = row.slice(0, dedupCols).join('\u0001');
    if (!seenRows.has(key)) {
      seenRows.add(key);
      dedupGrid.push(row);
      dedupOrigIndices.push(r < rawOriginalGrid.length ? r : -1);
    }
  }
  const afterDeduplicationRows = dedupGrid.length;
  console.log(`[DIAG] Após deduplicação: ${afterDeduplicationRows} linhas`);

  logs.push({
    stage: 'Etapa 1',
    step: 'Remoção de Duplicatas (Passo 9)',
    detail: `Deduplicação concluída: ${rawRows} → ${afterDeduplicationRows} linhas.`,
  });

  const foundLeituristasSet = new Set<string>();
  for (let r = 0; r < dedupGrid.length; r++) {
    const val = normalizeLeituristaCode(dedupGrid[r][4] ?? '');
    if (val.startsWith('Leiturista:')) {
      foundLeituristasSet.add(val);
    }
  }
  const foundLeituristasInFile = Array.from(foundLeituristasSet).sort();

  // (10) AutoFilter na linha 11 (index 10) e filtrar coluna E (Field:=5)
  const isHeaderCandidate = (row: string[] | undefined): boolean => {
    if (!row) return false;
    const normCells = row.map((c) => normalizeText(c));
    const hasMedidor = normCells.some(
      (c) => c === 'medidor' || c.includes('medidor'),
    );
    const hasSecondHeaderCol = normCells.some(
      (c) =>
        c.includes('t. reg') ||
        c.includes('t reg') ||
        c.includes('leit. exec') ||
        c.includes('leit exec') ||
        c === 'consumidora',
    );
    return hasMedidor && hasSecondHeaderCol;
  };

  let headerRow1Based = 11;
  const rowsMatchingAfter11 = dedupGrid
    .slice(11)
    .filter((row) => allowedSet.has(normalizeLeituristaCode(row[4] ?? ''))).length;

  if (rowsMatchingAfter11 === 0 || !isHeaderCandidate(dedupGrid[10])) {
    const detectedHeaderIdx = dedupGrid.findIndex((row) => isHeaderCandidate(row));
    if (detectedHeaderIdx >= 0) {
      headerRow1Based = detectedHeaderIdx + 1;
    } else if (rowsMatchingAfter11 === 0) {
      const firstMatchIdx = dedupGrid.findIndex((row) =>
        allowedSet.has(normalizeLeituristaCode(row[4] ?? '')),
      );
      if (firstMatchIdx > 0) {
        headerRow1Based = firstMatchIdx;
      }
    }
  }

  const headerRowIdx = headerRow1Based - 1;
  const headerRow = dedupGrid[headerRowIdx]
    ? [...dedupGrid[headerRowIdx]]
    : Array.from({ length: dedupCols }, () => '');

  const filteredDataRows: Grid = [];
  const filteredOrigIndices: number[] = [];
  const leituristaBreakdown: Record<string, number> = {};

  for (let r = headerRowIdx + 1; r < dedupGrid.length; r++) {
    const row = dedupGrid[r];
    const colE = normalizeLeituristaCode(row[4] ?? '');
    if (allowedSet.has(colE) && !isPageNoiseRow(row)) {
      filteredDataRows.push([...row]);
      filteredOrigIndices.push(dedupOrigIndices[r] ?? -1);
      const shortCode = colE.replace(/^Leiturista:\s*/i, '').trim() || colE;
      leituristaBreakdown[shortCode] = (leituristaBreakdown[shortCode] ?? 0) + 1;
    }
  }

  if (filteredDataRows.length === 0) {
    const foundList =
      foundLeituristasInFile.length > 0
        ? `Leituristas encontrados na coluna E: ${foundLeituristasInFile.join(', ')}.`
        : 'Nenhum valor no padrão "Leiturista: XXX" alinhou com a coluna E após os deslocamentos da Etapa 1.';
    throw new CemigValidationError(
      `Nenhuma linha de leitura restou após o filtro de leituristas da Etapa 1. ${foundList}`,
    );
  }

  const afterFilterRows = filteredDataRows.length;
  console.log(`[DIAG] Após filtro leiturista: ${afterFilterRows} linhas`);
  console.log(`[DIAG] Etapa 1 concluída: ${afterFilterRows} linhas`);

  // (11) Selecionar D12:E<last>
  const stage1IntermediateDE: Array<[string, string]> = filteredDataRows.map(
    (row) => [row[3] ?? '', row[4] ?? ''],
  );

  logs.push({
    stage: 'Etapa 1',
    step: 'Filtro de Leituristas e Extração D:E (Passos 10–11)',
    detail: `${afterFilterRows} registros retidos para os leituristas selecionados.`,
  });

  // --------------------------------------------------------------------------
  // ETAPA 2 — ParteDois_UlMatr (Ctrl+Shift+T)
  // --------------------------------------------------------------------------
  const stage2Grid: Grid = stage1IntermediateDE.map(([d, e]) => [d, e]);
  insertColumnsAt(stage2Grid, 2, 1);
  for (const row of stage2Grid) {
    const [left, right] = splitFixedWidthCol(row[0] ?? '', 19);
    row[0] = left;
    row[1] = right;
  }
  deleteColumn(stage2Grid, 1);

  ensureGridRect(stage2Grid, 3);
  for (const row of stage2Grid) {
    const [left, right] = splitFixedWidthCol(row[1] ?? '', 11);
    row[1] = left;
    row[2] = right;
  }
  deleteColumn(stage2Grid, 2);

  const ulMatrMap = new Map<string, string>();
  const stage2Pairs: Array<{ unidade: string; matricula: string }> = [];
  for (const row of stage2Grid) {
    const unidade = (row[0] ?? '').trim();
    const matricula = (row[1] ?? '').trim();
    if (unidade && !ulMatrMap.has(unidade)) {
      ulMatrMap.set(unidade, matricula);
      stage2Pairs.push({ unidade, matricula });
    }
  }
  console.log(`[DIAG] Etapa 2 concluída: ${stage2Pairs.length} linhas`);

  logs.push({
    stage: 'Etapa 2',
    step: 'ParteDois_UlMatr',
    detail: `${stage2Pairs.length} pares únicos de Unidade de Leitura ↔ Matrícula gerados.`,
  });

  // --------------------------------------------------------------------------
  // Preparação da aba "Filtros" (antiga "Planilha3")
  // --------------------------------------------------------------------------
  const {
    p3Grid,
    planilha3Provided,
    p3LookupAB,
    p3LookupDE,
    p3LookupGH,
    p3LookupJK,
  } = prepareFiltrosContext(
    planilha3Buffer,
    embeddedP3,
    populateP3FromStage2,
    stage2Pairs,
    warningsSet,
    options.consistenciaBuffer,
    options.nosbBuffer,
    options.fotosBuffer,
  );

  // --------------------------------------------------------------------------
  // ETAPA 3 — ParteTrês_RelatórioFinal (Ctrl+Shift+V) + 20 Colunas + Extras Brutos
  // --------------------------------------------------------------------------
  const fullStage3Grid: Grid = [
    [...headerRow],
    ...filteredDataRows.map((r) => [...r]),
  ];
  ensureGridRect(fullStage3Grid, Math.max(dedupCols, 16));

  // (5)-(6) Split largura fixa 20 na coluna D (idx 3) para extrair RZ_UL_LV (Unidade_2)
  insertColumnsAt(fullStage3Grid, 5, 1);
  for (let r = 0; r < fullStage3Grid.length; r++) {
    const rawD = fullStage3Grid[r][3] ?? '';
    if (r === 0) {
      fullStage3Grid[r][4] = 'Unidade_2';
      fullStage3Grid[r][3] = '';
    } else {
      const [left, right] = splitFixedWidthCol(rawD, 20);
      fullStage3Grid[r][3] = left;
      fullStage3Grid[r][4] = right.trim();
    }
  }
  deleteColumn(fullStage3Grid, 4); // Coluna D (idx 3) agora é Unidade_2 limpa

  // (8) Inserir duas colunas em E (E e F)
  insertColumnsAt(fullStage3Grid, 5, 2);
  fullStage3Grid[0][4] = 'Coluna_Aux';
  fullStage3Grid[0][5] = 'Matrícula';
  if ((fullStage3Grid[1]?.[6] ?? '').trim().startsWith('Leiturista:')) {
    fullStage3Grid[0][6] = 'Leiturista';
  }

  // (10) Inserir coluna em M (idx 12)
  insertColumnsAt(fullStage3Grid, 13, 1);
  fullStage3Grid[0][12] = 'Consumo';

  // Localizar colunas de origem no grid tratado da Etapa 3
  const fullHeaders = fullStage3Grid[0];
  const findSourceCol = (
    primary: string[],
    secondary: string[],
    fallbackIdx0: number,
  ): number => {
    for (const group of [primary, secondary]) {
      for (let c = 0; c < fullHeaders.length; c++) {
        const nh = normalizeText(fullHeaders[c] ?? '');
        if (!nh) continue;
        if (group.some((cand) => nh === cand || nh.includes(cand))) {
          return c;
        }
      }
    }
    return fallbackIdx0;
  };

  const srcInstalacaoIdx = 0; // A: Unidade / Consumidora
  const srcMedidorIdx = findSourceCol(['medidor'], [], 1);
  const srcRegistroIdx = findSourceCol(['t. reg', 't reg', 'registro'], [], 2);
  const srcRzUlLvIdx = 3; // D: Unidade_2 extraída no passo 6
  const srcLeituristaRawIdx = 6; // G: "Leiturista: XXX" (para fallback de matrícula e ordenação)
  const srcCodIdx = findSourceCol(['oc', 'cod'], ['sequencia'], 7);
  const srcDataIdx = findSourceCol(['data'], [], 8);
  const srcHoraIdx = findSourceCol(['hora'], [], 9);
  const srcEnderecoIdx = findSourceCol(['endereco', 'endereço'], ['apres'], 15);
  const srcLeitAntIdx = findSourceCol(['leit. ant', 'leit ant'], ['coluna6'], 10);
  const srcLeitExecIdx = findSourceCol(['leit. exec', 'leit exec'], ['coluna7'], 11);
  const srcImpIdx = findSourceCol(['imp.', 'imp'], ['coluna8'], 13);
  const srcQtdDigIdx = findSourceCol(['qtd dig', 'qtd. dig', 'digitacao'], [], 14);
  const srcApresIdx = findSourceCol(['apres.', 'apres'], [], -1);

  // Preservar demais colunas presentes na planilha importada, exceto "c", "A Exec.", "Exec." e "Leiturista"
  const consumedStage3Cols = new Set<number>([
    srcInstalacaoIdx,
    srcMedidorIdx,
    srcRegistroIdx,
    srcRzUlLvIdx,
    4, // Coluna_Aux placeholder
    5, // Matrícula placeholder
    srcLeituristaRawIdx, // Leiturista excluída da tabela final (mantida apenas na ordenação)
    srcCodIdx,
    srcDataIdx,
    srcHoraIdx,
    srcEnderecoIdx,
    srcLeitAntIdx,
    srcLeitExecIdx,
    12, // Consumo placeholder
    srcImpIdx,
    srcQtdDigIdx,
  ]);
  if (srcApresIdx >= 0) consumedStage3Cols.add(srcApresIdx);

  const usedHeaderNamesLower = new Set<string>(
    FINAL_HEADERS.map((h) => h.toLowerCase()),
  );
  const makeUniqueHeader = (candidate: string): string => {
    const clean =
      candidate.replace(/[\r\n\t]+/g, ' ').trim().slice(0, 255) || 'Coluna_Extra';
    if (!usedHeaderNamesLower.has(clean.toLowerCase())) {
      usedHeaderNamesLower.add(clean.toLowerCase());
      return clean;
    }
    let suffix = 2;
    while (usedHeaderNamesLower.has(`${clean}_${suffix}`.toLowerCase())) {
      suffix++;
    }
    const finalName = `${clean}_${suffix}`.slice(0, 255);
    usedHeaderNamesLower.add(finalName.toLowerCase());
    return finalName;
  };

  interface ExtraColSpec {
    header: string;
    source: 'stage3' | 'rawOrig';
    colIdx: number;
  }

  const extraColSpecs: ExtraColSpec[] = [];

  for (let c = 0; c < fullHeaders.length; c++) {
    if (consumedStage3Cols.has(c)) continue;
    const rawHeader = (fullHeaders[c] ?? '').trim();
    if (rawHeader && isExcludedExtraColumn(rawHeader)) continue;
    const hasAnyData = fullStage3Grid
      .slice(1)
      .some((row) => (row[c] ?? '').trim() !== '');
    if (!rawHeader && !hasAnyData) continue;
    extraColSpecs.push({
      header: makeUniqueHeader(rawHeader || `Coluna_Bruta_${c + 1}`),
      source: 'stage3',
      colIdx: c,
    });
  }

  // Colunas originais removidas nos passos 1 e 2 da Etapa 1 (orig idx 4, 7, 8)
  const origHeaderIdx = dedupOrigIndices[headerRowIdx] ?? -1;
  const rawHeaderOrig =
    origHeaderIdx >= 0 && origHeaderIdx < rawOriginalGrid.length
      ? rawOriginalGrid[origHeaderIdx]
      : [];
  for (const origColIdx of [4, 7, 8]) {
    if (origColIdx >= rawCols) continue;
    const rawHeader = (rawHeaderOrig[origColIdx] ?? '').trim();
    if (rawHeader && isExcludedExtraColumn(rawHeader)) continue;
    const hasAnyData = filteredOrigIndices.some((origR) =>
      origR >= 0 ? ((rawOriginalGrid[origR]?.[origColIdx] ?? '').trim() !== '') : false,
    );
    if (!rawHeader && !hasAnyData) continue;
    extraColSpecs.push({
      header: makeUniqueHeader(rawHeader || `Coluna_Original_${origColIdx + 1}`),
      source: 'rawOrig',
      colIdx: origColIdx,
    });
  }

  const finalHeaders: string[] = [
    ...FINAL_HEADERS,
    ...extraColSpecs.map((spec) => spec.header),
  ];

  // Detectar ordem das datas no lote (M/D/YYYY vs D/M/YYYY)
  const rawDatesSample = fullStage3Grid
    .slice(1)
    .map((row) => row[srcDataIdx] ?? '')
    .filter(Boolean);
  const preferMDY = detectPreferMDY(rawDatesSample);

  const records: NormalizedRecord[] = [];
  for (let r = 1; r < fullStage3Grid.length; r++) {
    const row = fullStage3Grid[r];
    const origR = filteredOrigIndices[r - 1] ?? -1;
    const instalacao = (row[srcInstalacaoIdx] ?? '').trim();
    const medidor = (row[srcMedidorIdx] ?? '').trim();
    const registro = (row[srcRegistroIdx] ?? '').trim();
    const rzUlLv = (row[srcRzUlLvIdx] ?? '').trim();
    const ul = extractULFromRzUlLv(rzUlLv);

    const leituristaFallback = (row[srcLeituristaRawIdx] ?? '')
      .replace(/^Leiturista:\s*/i, '')
      .trim();
    const matricula = p3LookupAB.get(rzUlLv) || leituristaFallback;
    const tipo = classifyTipoFromRzUlLv(rzUlLv);

    const cod = (row[srcCodIdx] ?? '').trim();
    const data = formatDateToDDMMYYYY(row[srcDataIdx] ?? '', preferMDY);
    const hora = (row[srcHoraIdx] ?? '').trim();
    const endereco = (row[srcEnderecoIdx] ?? '').trim();
    const leitAntRaw = (row[srcLeitAntIdx] ?? '').trim();
    const leitExecRaw = (row[srcLeitExecIdx] ?? '').trim();

    const antNum = parseNumberForCalc(leitAntRaw);
    const execNum = parseNumberForCalc(leitExecRaw);
    const consumoCalc =
      antNum !== null && execNum !== null ? execNum - antNum : '';

    // Passo 12: preencher brancos com "Faturada" e "N00"
    const impedimento = (row[srcImpIdx] ?? '').trim() || 'Faturada';
    const digitacao = (row[srcQtdDigIdx] ?? '').trim() || 'N00';

    // Apres.: "00" se exatamente "00", caso contrário "N00"
    const rawApres = srcApresIdx >= 0 ? (row[srcApresIdx] ?? '').trim() : '';
    const apres = normalizeApresValue(rawApres);

    // Passo 13: enriquecimento Filtros com fallback IFERROR (chave = Coluna B / MEDIDOR)
    const consistencia = p3LookupDE.get(medidor) || 'Faturada';
    const simulacao = p3LookupGH.get(medidor) || 'Faturada';
    const fotos = p3LookupJK.get(medidor) || 'N-OK';

    const extraValues = extraColSpecs.map((spec) => {
      if (spec.source === 'stage3') {
        return (row[spec.colIdx] ?? '').trim();
      }
      return origR >= 0 ? (rawOriginalGrid[origR]?.[spec.colIdx] ?? '').trim() : '';
    });

    records.push({
      instalacao,
      medidor,
      registro,
      rzUlLv,
      ul,
      matricula,
      tipo,
      cod,
      data,
      hora,
      endereco,
      leitAntRaw,
      leitExecRaw,
      consumoCalc,
      impedimento,
      digitacao,
      apres,
      consistencia,
      simulacao,
      fotos,
      extraValues,
      dateSortKey: dateDDMMYYYYToSortKey(data),
      timeSortKey: timeToSortSeconds(hora),
      originalOrder: r,
    });
  }

  // ORDENAÇÃO OBRIGATÓRIA:
  // Agrupado/ordenado por LEITURISTA (MATRÍCULA) e, dentro de cada leiturista,
  // ordem crescente por DATA e depois por HORA.
  records.sort((a, b) => {
    const leitCmp = a.matricula.localeCompare(b.matricula, 'pt-BR', {
      sensitivity: 'base',
    });
    if (leitCmp !== 0) return leitCmp;
    if (a.dateSortKey !== b.dateSortKey) return a.dateSortKey - b.dateSortKey;
    if (a.timeSortKey !== b.timeSortKey) return a.timeSortKey - b.timeSortKey;
    return a.originalOrder - b.originalOrder;
  });

  console.log(`[DIAG] Etapa 3 concluída: ${records.length} linhas`);
  console.log(
    `[DIAG] Tabela final montada: ${records.length} linhas x ${finalHeaders.length} colunas`,
  );

  return buildOutputWorkbookAndResult({
    sheetName,
    rawRows,
    rawCols,
    afterDeduplicationRows,
    afterFilterRows,
    importedMetersCount: afterFilterRows,
    records,
    finalHeaders,
    p3Grid,
    planilha3Provided,
    stage2Pairs,
    leituristaBreakdown,
    foundLeituristasInFile,
    logs,
    warningsSet,
    startTime,
    modeloRGeralBuffer: options.modeloRGeralBuffer,
    modeloRGeralFileName: options.modeloRGeralFileName,
  });
}
