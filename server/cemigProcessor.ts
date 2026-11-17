import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

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
  'MATRÍCULA',
  'COD',
  'DATA',
  'HORA',
  'ENDEREÇO',
  'Leit. Anterior',
  'Leit. Executada',
  'CONSUMO',
  'NOSB´IMPEDIMENTO',
  'DIGITAÇÃO',
  'CONSISTÊNCIA',
  'NOSB´SIMULAÇÃO',
  'FOTOS',
];

export interface ProcessOptions {
  allowedLeituristas?: string[];
  populatePlanilha3FromStage2?: boolean;
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
  planilha3Provided: boolean;
  planilha3Rows: number;
  leituristaBreakdown: Record<string, number>;
  foundLeituristasInFile: string[];
  durationMs: number;
}

export interface PreviewData {
  headers: string[];
  rows: string[][];
  totalDataRows: number;
}

export interface ProcessResult {
  outputBuffer: Buffer;
  fileName: string;
  stats: ProcessingStats;
  logs: StageStepLog[];
  warnings: string[];
  preview: PreviewData;
  stage2Sample: Array<{ unidade: string; matricula: string }>;
}

export class CemigValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CemigValidationError';
  }
}

type Grid = string[][];

function normalizeText(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
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
  } catch {
    throw new CemigValidationError(
      'Não foi possível ler o arquivo enviado. Verifique se é uma planilha Excel válida (.xlsx ou .xls).',
    );
  }

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new CemigValidationError('O arquivo Excel enviado não contém nenhuma aba.');
  }

  const targetName =
    preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
      ? preferredSheetName
      : workbook.SheetNames[0];

  const extractSheet = (name: string): Grid => {
    const ws = workbook.Sheets[name];
    if (!ws || !ws['!ref']) return [];

    const range = XLSX.utils.decode_range(ws['!ref']);
    const numRows = range.e.r + 1;
    const numCols = range.e.c + 1;

    const grid: Grid = Array.from({ length: numRows }, () =>
      Array.from({ length: numCols }, () => ''),
    );

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSX.CellObject | undefined;
        if (!cell) continue;

        let strVal = '';
        if (typeof cell.w === 'string' && cell.w.length > 0) {
          strVal = cell.w;
        } else if (cell.v !== undefined && cell.v !== null) {
          strVal = String(cell.v);
        }
        grid[r][c] = strVal;
      }
    }

    // Unmerge cells: Excel UnMerge keeps top-left cell value and clears remaining cells
    const merges = ws['!merges'] ?? [];
    for (const m of merges) {
      const topLeftVal = grid[m.s.r]?.[m.s.c] ?? '';
      for (let r = m.s.r; r <= m.e.r; r++) {
        for (let c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) {
            if (grid[r]) grid[r][c] = topLeftVal;
          } else if (grid[r] && c < grid[r].length) {
            grid[r][c] = '';
          }
        }
      }
    }

    return trimTrailingEmptyRows(grid);
  };

  const mainGrid = extractSheet(targetName);

  let planilha3Grid: Grid | null = null;
  const p3Name = workbook.SheetNames.find(
    (s) => normalizeText(s) === 'planilha3' && s !== targetName,
  );
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
    (joined.includes('pag.') && joined.includes('/')) ||
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

  const allowedSet = new Set(allowedLeituristas.map((s) => s.trim()));
  const populateP3FromStage2 = options.populatePlanilha3FromStage2 ?? true;

  // Read Matriz workbook
  const { sheetName, grid, planilha3Grid: embeddedP3 } =
    readWorkbookSheetAsGrid(matrizBuffer);

  validateCemigLayout(grid);

  const rawRows = grid.length;
  const rawCols = getGridMaxCols(grid);
  ensureGridRect(grid, Math.max(rawCols, 16));

  logs.push({
    stage: 'Etapa 1',
    step: 'Leitura e Desmesclagem',
    detail: `Aba "${sheetName}" lida com ${rawRows} linhas e ${rawCols} colunas preservando zeros à esquerda e datas.`,
  });

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
  for (const row of grid) {
    const key = row.slice(0, dedupCols).join('\u0001');
    if (!seenRows.has(key)) {
      seenRows.add(key);
      dedupGrid.push(row);
    }
  }
  const afterDeduplicationRows = dedupGrid.length;

  logs.push({
    stage: 'Etapa 1',
    step: 'Remoção de Duplicatas (Passo 9)',
    detail: `Deduplicação concluída: ${rawRows} → ${afterDeduplicationRows} linhas.`,
  });

  const foundLeituristasSet = new Set<string>();
  for (let r = 0; r < dedupGrid.length; r++) {
    const val = (dedupGrid[r][4] ?? '').trim();
    if (val.startsWith('Leiturista:')) {
      foundLeituristasSet.add(val);
    }
  }
  const foundLeituristasInFile = Array.from(foundLeituristasSet).sort();

  // (10) AutoFilter na linha 11 (index 10) e filtrar coluna E (Field:=5)
  const isHeaderCandidate = (row: string[] | undefined): boolean => {
    if (!row) return false;
    const normJoined = row.map((c) => normalizeText(c)).join(' | ');
    return (
      normJoined.includes('medidor') ||
      normJoined.includes('leit. exec') ||
      normJoined.includes('leit exec') ||
      normJoined.includes('consumidora')
    );
  };

  let headerRow1Based = 11;
  const rowsMatchingAfter11 = dedupGrid
    .slice(11)
    .filter((row) => allowedSet.has((row[4] ?? '').trim())).length;

  if (rowsMatchingAfter11 === 0 || !isHeaderCandidate(dedupGrid[10])) {
    const detectedHeaderIdx = dedupGrid.findIndex((row) => isHeaderCandidate(row));
    if (detectedHeaderIdx >= 0) {
      headerRow1Based = detectedHeaderIdx + 1;
    } else if (rowsMatchingAfter11 === 0) {
      const firstMatchIdx = dedupGrid.findIndex((row) =>
        allowedSet.has((row[4] ?? '').trim()),
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
  const leituristaBreakdown: Record<string, number> = {};

  for (let r = headerRowIdx + 1; r < dedupGrid.length; r++) {
    const row = dedupGrid[r];
    const colE = (row[4] ?? '').trim();
    if (allowedSet.has(colE) && !isPageNoiseRow(row)) {
      filteredDataRows.push([...row]);
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

  logs.push({
    stage: 'Etapa 2',
    step: 'ParteDois_UlMatr',
    detail: `${stage2Pairs.length} pares únicos de Unidade de Leitura ↔ Matrícula gerados.`,
  });

  // --------------------------------------------------------------------------
  // Preparação da Planilha3
  // --------------------------------------------------------------------------
  let p3Grid: Grid = [];
  let planilha3Provided = false;

  if (planilha3Buffer) {
    const parsedP3 = readWorkbookSheetAsGrid(planilha3Buffer, 'Planilha3');
    p3Grid = parsedP3.grid;
    planilha3Provided = true;
  } else if (embeddedP3 && embeddedP3.length > 0) {
    p3Grid = embeddedP3;
    planilha3Provided = true;
  } else {
    if (populateP3FromStage2) {
      p3Grid = stage2Pairs.map((pair) => [pair.unidade, pair.matricula]);
    } else {
      p3Grid = [];
    }
    warningsSet.add('Tabela completa do tratamento.');
  }

  ensureGridRect(p3Grid, 11);

  const p3LookupAB = new Map<string, string>();
  const p3LookupDE = new Map<string, string>();
  const p3LookupGH = new Map<string, string>();
  const p3LookupJK = new Map<string, string>();

  for (let r = 0; r < Math.min(p3Grid.length, 50000); r++) {
    const row = p3Grid[r];
    const kA = (row[0] ?? '').trim();
    if (kA && !p3LookupAB.has(kA)) p3LookupAB.set(kA, (row[1] ?? '').trim());

    const kD = (row[3] ?? '').trim();
    if (kD && !p3LookupDE.has(kD)) p3LookupDE.set(kD, (row[4] ?? '').trim());

    const kG = (row[6] ?? '').trim();
    if (kG && !p3LookupGH.has(kG)) p3LookupGH.set(kG, (row[7] ?? '').trim());

    const kJ = (row[9] ?? '').trim();
    if (kJ && !p3LookupJK.has(kJ)) p3LookupJK.set(kJ, (row[10] ?? '').trim());
  }

  // --------------------------------------------------------------------------
  // ETAPA 3 — ParteTrês_RelatórioFinal (Ctrl+Shift+V) + Mapeamento Exato 17 Colunas
  // --------------------------------------------------------------------------
  // Preservar todas as colunas de headerRow + filteredDataRows para localizar OC, Endereço e Qtd Dig. sem perda
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
  const srcLeituristaRawIdx = 6; // G: "Leiturista: XXX" (para fallback de matrícula)
  const srcCodIdx = findSourceCol(['oc', 'cod'], ['sequencia'], 7);
  const srcDataIdx = findSourceCol(['data'], [], 8);
  const srcHoraIdx = findSourceCol(['hora'], [], 9);
  const srcEnderecoIdx = findSourceCol(['endereco', 'endereço', 'apres'], [], 15);
  const srcLeitAntIdx = findSourceCol(['leit. ant', 'leit ant'], ['coluna6'], 10);
  const srcLeitExecIdx = findSourceCol(['leit. exec', 'leit exec'], ['coluna7'], 11);
  const srcImpIdx = findSourceCol(['imp.', 'imp'], ['coluna8'], 13);
  const srcQtdDigIdx = findSourceCol(['qtd dig', 'qtd. dig', 'digitacao'], [], 14);

  // Detectar ordem das datas no lote (M/D/YYYY vs D/M/YYYY)
  const rawDatesSample = fullStage3Grid
    .slice(1)
    .map((row) => row[srcDataIdx] ?? '')
    .filter(Boolean);
  const preferMDY = detectPreferMDY(rawDatesSample);

  // Montar registros normalizados nas 17 colunas exatas
  interface NormalizedRecord {
    instalacao: string;
    medidor: string;
    registro: string;
    rzUlLv: string;
    matricula: string;
    cod: string;
    data: string;
    hora: string;
    endereco: string;
    leitAntRaw: string;
    leitExecRaw: string;
    consumoCalc: number | '';
    impedimento: string;
    digitacao: string;
    consistencia: string;
    simulacao: string;
    fotos: string;
    dateSortKey: number;
    timeSortKey: number;
    originalOrder: number;
  }

  const records: NormalizedRecord[] = [];
  for (let r = 1; r < fullStage3Grid.length; r++) {
    const row = fullStage3Grid[r];
    const instalacao = (row[srcInstalacaoIdx] ?? '').trim();
    const medidor = (row[srcMedidorIdx] ?? '').trim();
    const registro = (row[srcRegistroIdx] ?? '').trim();
    const rzUlLv = (row[srcRzUlLvIdx] ?? '').trim();

    const leituristaFallback = (row[srcLeituristaRawIdx] ?? '')
      .replace(/^Leiturista:\s*/i, '')
      .trim();
    const matricula = p3LookupAB.get(rzUlLv) || leituristaFallback;

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

    // Passo 13: enriquecimento Planilha3 com fallback IFERROR
    const consistencia = p3LookupDE.get(instalacao) || 'Faturada';
    const simulacao = p3LookupGH.get(instalacao) || 'Faturada';
    const fotos = p3LookupJK.get(instalacao) || 'N-OK';

    records.push({
      instalacao,
      medidor,
      registro,
      rzUlLv,
      matricula,
      cod,
      data,
      hora,
      endereco,
      leitAntRaw,
      leitExecRaw,
      consumoCalc,
      impedimento,
      digitacao,
      consistencia,
      simulacao,
      fotos,
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

  logs.push({
    stage: 'Etapa 3',
    step: 'Tabela1 (17 Colunas), Datas dd/mm/aaaa e Ordenação',
    detail: `${records.length} registros ordenados por Leiturista → DATA (dd/mm/aaaa) → HORA crescente.`,
  });

  // --------------------------------------------------------------------------
  // Gerar Workbook Final (17 colunas exatas + Planilha3)
  // --------------------------------------------------------------------------
  const outWb = new ExcelJS.Workbook();
  outWb.creator = 'Controle de Sequenciamento de Leitura';
  outWb.created = new Date();

  const wsFinal = outWb.addWorksheet('ControleLeitura', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  const wsP3 = outWb.addWorksheet('Planilha3');

  if (p3Grid.length > 0) {
    for (let r = 0; r < p3Grid.length; r++) {
      const outRow = wsP3.getRow(r + 1);
      for (let c = 0; c < p3Grid[r].length; c++) {
        const val = p3Grid[r][c] ?? '';
        if (val !== '') {
          const cell = outRow.getCell(c + 1);
          cell.value = val;
          cell.numFmt = '@';
        }
      }
    }
  } else {
    wsP3.getCell('A1').value = '';
  }

  const allPreviewRows: string[][] = [];
  const tableRowsForExcelJS: Array<
    Array<string | number | { formula: string; result?: string | number }>
  > = [];

  for (let i = 0; i < records.length; i++) {
    const rec = records[i];
    const excelRowNum = i + 2;

    const antNumeric = parseNumericMaybe(rec.leitAntRaw);
    const execNumeric = parseNumericMaybe(rec.leitExecRaw);

    const rowCellsForTable: Array<
      string | number | { formula: string; result?: string | number }
    > = [
      rec.instalacao, // 1. A: INSTALAÇÃO
      rec.medidor, // 2. B: MEDIDOR
      rec.registro, // 3. C: REGISTRO
      rec.rzUlLv, // 4. D: RZ_UL_LV
      {
        // 5. E: MATRÍCULA
        formula: `VLOOKUP(D${excelRowNum},Planilha3!$A$1:$B$50000,2,0)`,
        result: rec.matricula || undefined,
      },
      rec.cod, // 6. F: COD
      rec.data, // 7. G: DATA (dd/mm/aaaa)
      rec.hora, // 8. H: HORA
      rec.endereco, // 9. I: ENDEREÇO
      antNumeric !== null ? antNumeric : rec.leitAntRaw, // 10. J: Leit. Anterior
      execNumeric !== null ? execNumeric : rec.leitExecRaw, // 11. K: Leit. Executada
      {
        // 12. L: CONSUMO
        formula: `SUM(K${excelRowNum}-J${excelRowNum})`,
        result: typeof rec.consumoCalc === 'number' ? rec.consumoCalc : 0,
      },
      rec.impedimento, // 13. M: NOSB´IMPEDIMENTO
      rec.digitacao, // 14. N: DIGITAÇÃO
      {
        // 15. O: CONSISTÊNCIA
        formula: `IFERROR(VLOOKUP(A${excelRowNum},Planilha3!$D$1:$E$50000,2,0),"Faturada")`,
        result: rec.consistencia,
      },
      {
        // 16. P: NOSB´SIMULAÇÃO
        formula: `IFERROR(VLOOKUP(A${excelRowNum},Planilha3!$G$1:$H$50000,2,0),"Faturada")`,
        result: rec.simulacao,
      },
      {
        // 17. Q: FOTOS
        formula: `IFERROR(VLOOKUP(A${excelRowNum},Planilha3!$J$1:$K$50000,2,0),"N-OK")`,
        result: rec.fotos,
      },
    ];

    tableRowsForExcelJS.push(rowCellsForTable);

    allPreviewRows.push([
      rec.instalacao,
      rec.medidor,
      rec.registro,
      rec.rzUlLv,
      rec.matricula,
      rec.cod,
      rec.data,
      rec.hora,
      rec.endereco,
      rec.leitAntRaw,
      rec.leitExecRaw,
      typeof rec.consumoCalc === 'number' ? String(rec.consumoCalc) : '',
      rec.impedimento,
      rec.digitacao,
      rec.consistencia,
      rec.simulacao,
      rec.fotos,
    ]);
  }

  // Criar Tabela1 em A1:Q<totalRows>
  wsFinal.addTable({
    name: 'Tabela1',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: FINAL_HEADERS.map((name) => ({
      name,
      filterButton: true,
    })),
    rows: tableRowsForExcelJS,
  });

  // Estilizar cabeçalho (1..17)
  const outHeaderRow = wsFinal.getRow(1);
  for (let c = 1; c <= 17; c++) {
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
      fgColor: { argb: c >= 15 ? 'FF0F766E' : 'FF0F172A' },
    };
    cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: false };
  }

  // Alinhamento, zeros à esquerda como texto '@' e cálculo de largura AutoFit
  const colMaxWidths = FINAL_HEADERS.map((h) => h.length + 4);
  const totalSheetRows = records.length + 1;

  for (let r = 2; r <= totalSheetRows; r++) {
    const excelRow = wsFinal.getRow(r);
    for (let c = 1; c <= 17; c++) {
      const cell = excelRow.getCell(c);
      cell.alignment = {
        horizontal: 'left',
        vertical: 'bottom',
        wrapText: false,
      };
      if (typeof cell.value === 'string') {
        if (/^0\d+$/.test(cell.value) || c === 7) {
          cell.numFmt = '@';
        }
        if (cell.value.length + 2 > colMaxWidths[c - 1]) {
          colMaxWidths[c - 1] = cell.value.length + 2;
        }
      } else if (
        cell.value &&
        typeof cell.value === 'object' &&
        'result' in cell.value &&
        cell.value.result !== undefined
      ) {
        const resLen = String(cell.value.result).length + 2;
        if (resLen > colMaxWidths[c - 1]) colMaxWidths[c - 1] = resLen;
      } else if (typeof cell.value === 'number') {
        const numLen = String(cell.value).length + 2;
        if (numLen > colMaxWidths[c - 1]) colMaxWidths[c - 1] = numLen;
      }
    }
  }

  for (let c = 1; c <= 17; c++) {
    wsFinal.getColumn(c).width = Math.min(Math.max(colMaxWidths[c - 1], 11), 46);
  }
  wsP3.getColumn(1).width = 18;
  wsP3.getColumn(2).width = 14;

  const outArrayBuffer = await outWb.xlsx.writeBuffer();
  const outputBuffer = Buffer.from(outArrayBuffer);

  const durationMs = Date.now() - startTime;

  return {
    outputBuffer,
    fileName: 'Matriz_ControleDeSequenciaLeitura.xlsx',
    stats: {
      sheetName,
      rawRows,
      rawCols,
      afterDeduplicationRows,
      afterFilterRows,
      finalRows: records.length,
      finalCols: 17,
      uniqueUnitsCount: stage2Pairs.length,
      planilha3Provided,
      planilha3Rows: p3Grid.length,
      leituristaBreakdown,
      foundLeituristasInFile,
      durationMs,
    },
    logs,
    warnings: Array.from(warningsSet),
    preview: {
      headers: [...FINAL_HEADERS],
      rows: allPreviewRows,
      totalDataRows: records.length,
    },
    stage2Sample: stage2Pairs.slice(0, 50),
  };
}
