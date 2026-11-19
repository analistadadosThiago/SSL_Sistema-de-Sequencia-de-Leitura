import JSZip from 'jszip';
import * as XLSX from 'xlsx';

export const RGERAL_COLUMN_HEADERS: readonly string[] = [
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

export interface RGeralRowData {
  instalacao: string;
  medidor: string;
  registro: string;
  rzUlLv: string;
  ul: string;
  fotos: string;
  tipo: string;
  matricula: string;
  cod: string;
  data: string;
  hora: string;
  endereco: string;
  leitAnterior: string | number;
  leitAtual: string | number;
  consumo: string | number;
  impedimento: string;
  simulacao: string;
  digitacao: string;
  apres: string;
  consistencia: string;
}

export interface RGeralFillResult {
  buffer: Buffer;
  fileName: string;
}

const TEXT_PRESERVING_COLS = new Set([1, 2, 3, 4, 5, 9, 10, 11, 19]);

function sanitizeXmlText(val: string): string {
  // Remove caracteres de controle XML 1.0 inválidos (\x00-\x08, \x0B, \x0C, \x0E-\x1F)
  return val.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function escapeXml(val: string): string {
  return sanitizeXmlText(val)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXmlEntities(val: string): string {
  return val
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => {
      const code = parseInt(hex, 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : '';
    })
    .replace(/&#(\d+);/g, (_, dec) => {
      const code = parseInt(dec, 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : '';
    })
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function colNumToLetters(colNum: number): string {
  let n = colNum;
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function colLettersToNum(letters: string): number {
  let num = 0;
  for (let i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 64);
  }
  return num;
}

function parseCellRef(ref: string): { colNum: number; rowNum: number } | null {
  const m = /^([A-Za-z]+)(\d+)$/.exec(ref.trim());
  if (!m) return null;
  return {
    colNum: colLettersToNum(m[1].toUpperCase()),
    rowNum: parseInt(m[2], 10),
  };
}

function getAttr(attrString: string, attrName: string): string | null {
  const escaped = attrName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*"([^"]*)"`);
  const m = re.exec(attrString);
  if (m) return m[1];
  const reSingle = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*'([^']*)'`);
  const m2 = reSingle.exec(attrString);
  return m2 ? m2[1] : null;
}

function setAttr(attrString: string, attrName: string, value: string): string {
  const escaped = attrName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(\\s|^)${escaped}\\s*=\\s*(?:"[^"]*"|'[^']*')`);
  if (re.test(attrString)) {
    return attrString.replace(re, `$1${attrName}="${value}"`);
  }
  const trimmed = attrString.trim();
  return trimmed ? `${trimmed} ${attrName}="${value}"` : `${attrName}="${value}"`;
}

function parseSharedStrings(xml: string | null): string[] {
  if (!xml) return [];
  const result: string[] = [];
  const siRegex = /<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g;
  let siMatch: RegExpExecArray | null;
  while ((siMatch = siRegex.exec(xml)) !== null) {
    const inner = siMatch[1];
    let text = '';
    const tRegex = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let tMatch: RegExpExecArray | null;
    while ((tMatch = tRegex.exec(inner)) !== null) {
      text += decodeXmlEntities(tMatch[1]);
    }
    result.push(text);
  }
  return result;
}

interface ParsedCell {
  colNum: number;
  ref: string;
  attrs: string;
  styleAttr: string | null;
  typeAttr: string | null;
  rawXml: string;
  textValue: string;
}

interface ParsedRow {
  rowNum: number;
  rowAttrs: string;
  rawRowXml: string;
  cells: Map<number, ParsedCell>;
}

function extractCellText(
  typeAttr: string | null,
  innerXml: string,
  sharedStrings: string[],
): string {
  if (typeAttr === 'inlineStr') {
    let text = '';
    const tRegex = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let m: RegExpExecArray | null;
    while ((m = tRegex.exec(innerXml)) !== null) {
      text += decodeXmlEntities(m[1]);
    }
    return text;
  }
  const vMatch = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(innerXml);
  if (!vMatch) return '';
  const rawV = decodeXmlEntities(vMatch[1]);
  if (typeAttr === 's') {
    const idx = parseInt(rawV, 10);
    return Number.isFinite(idx) && sharedStrings[idx] !== undefined
      ? sharedStrings[idx]
      : '';
  }
  return rawV;
}

function parseCellsFromRowInner(
  rowInner: string,
  rowNum: number,
  sharedStrings: string[],
): Map<number, ParsedCell> {
  const cells = new Map<number, ParsedCell>();
  const cellRegex = /<c(\s[^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let cm: RegExpExecArray | null;
  let seqCol = 1;
  while ((cm = cellRegex.exec(rowInner)) !== null) {
    const attrs = cm[1] ?? '';
    const inner = cm[2] ?? '';
    const refAttr = getAttr(attrs, 'r');
    const parsedRef = refAttr ? parseCellRef(refAttr) : null;
    const colNum = parsedRef ? parsedRef.colNum : seqCol;
    seqCol = colNum + 1;
    const ref = refAttr ?? `${colNumToLetters(colNum)}${rowNum}`;
    const styleAttr = getAttr(attrs, 's');
    const typeAttr = getAttr(attrs, 't');
    const textValue = extractCellText(typeAttr, inner, sharedStrings);
    cells.set(colNum, {
      colNum,
      ref,
      attrs,
      styleAttr,
      typeAttr,
      rawXml: cm[0],
      textValue,
    });
  }
  return cells;
}

function parseSheetDataRows(
  sheetDataInner: string,
  sharedStrings: string[],
): Map<number, ParsedRow> {
  const rows = new Map<number, ParsedRow>();
  const rowRegex = /<row(\s[^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let rm: RegExpExecArray | null;
  let seqRow = 1;
  while ((rm = rowRegex.exec(sheetDataInner)) !== null) {
    const rowAttrs = rm[1] ?? '';
    const rowInner = rm[2] ?? '';
    const rAttr = getAttr(rowAttrs, 'r');
    const rowNum = rAttr ? parseInt(rAttr, 10) : seqRow;
    seqRow = rowNum + 1;
    const cells = rowInner
      ? parseCellsFromRowInner(rowInner, rowNum, sharedStrings)
      : new Map<number, ParsedCell>();
    rows.set(rowNum, {
      rowNum,
      rowAttrs,
      rawRowXml: rm[0],
      cells,
    });
  }
  return rows;
}

interface StylesContext {
  stylesXml: string | null;
  getTextStyleId: (baseStyleId: string | null) => string | null;
  finalizeStylesXml: () => string | null;
}

function createStylesContext(stylesXml: string | null): StylesContext {
  if (!stylesXml) {
    return {
      stylesXml: null,
      getTextStyleId: (base) => base,
      finalizeStylesXml: () => null,
    };
  }

  const cellXfsMatch = /(<cellXfs\b([^>]*)>)([\s\S]*?)(<\/cellXfs>)/.exec(stylesXml);
  if (!cellXfsMatch) {
    return {
      stylesXml,
      getTextStyleId: (base) => base,
      finalizeStylesXml: () => stylesXml,
    };
  }

  const xfsInner = cellXfsMatch[3];
  const xfList: string[] = [];
  const xfRegex = /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g;
  let xm: RegExpExecArray | null;
  while ((xm = xfRegex.exec(xfsInner)) !== null) {
    xfList.push(xm[0]);
  }

  const cache = new Map<string, string>();
  let plainTextStyleIdx: string | null = null;
  for (let i = 0; i < xfList.length; i++) {
    const numFmtId = getAttr(xfList[i], 'numFmtId');
    if (numFmtId === '49') {
      const fontId = getAttr(xfList[i], 'fontId') ?? '0';
      const fillId = getAttr(xfList[i], 'fillId') ?? '0';
      const borderId = getAttr(xfList[i], 'borderId') ?? '0';
      if (fontId === '0' && fillId === '0' && borderId === '0') {
        plainTextStyleIdx = String(i);
        break;
      }
    }
  }

  let modified = false;

  const getTextStyleId = (baseStyleId: string | null): string | null => {
    const key = baseStyleId ?? '__none__';
    const existing = cache.get(key);
    if (existing !== undefined) return existing;

    if (baseStyleId !== null) {
      const idx = parseInt(baseStyleId, 10);
      if (Number.isFinite(idx) && idx >= 0 && idx < xfList.length) {
        const baseXf = xfList[idx];
        if (getAttr(baseXf, 'numFmtId') === '49') {
          cache.set(key, baseStyleId);
          return baseStyleId;
        }
        // Clonar o estilo base preservando fonte/borda/preenchimento/alinhamento, trocando apenas numFmtId="49"
        const openTagMatch = /^<xf\b([^>]*?)(\/>|>)/.exec(baseXf);
        if (openTagMatch) {
          let newAttrs = setAttr(openTagMatch[1], 'numFmtId', '49');
          newAttrs = setAttr(newAttrs, 'applyNumberFormat', '1');
          const rest = baseXf.slice(openTagMatch[0].length);
          const cloned = `<xf ${newAttrs}${openTagMatch[2]}${rest}`;
          const newIdx = String(xfList.length);
          xfList.push(cloned);
          modified = true;
          cache.set(key, newIdx);
          return newIdx;
        }
      }
    }

    if (plainTextStyleIdx === null) {
      plainTextStyleIdx = String(xfList.length);
      xfList.push(
        '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      );
      modified = true;
    }
    cache.set(key, plainTextStyleIdx);
    return plainTextStyleIdx;
  };

  const finalizeStylesXml = (): string | null => {
    if (!modified) return stylesXml;
    const openTagAttrs = setAttr(cellXfsMatch[2], 'count', String(xfList.length));
    const replacement = `<cellXfs ${openTagAttrs}>${xfList.join('')}</cellXfs>`;
    return (
      stylesXml.slice(0, cellXfsMatch.index) +
      replacement +
      stylesXml.slice(cellXfsMatch.index + cellXfsMatch[0].length)
    );
  };

  return {
    stylesXml,
    getTextStyleId,
    finalizeStylesXml,
  };
}

function buildInlineStrCellXml(
  ref: string,
  val: string,
  styleId: string | null,
): string {
  const sPart = styleId !== null ? ` s="${styleId}"` : '';
  const clean = sanitizeXmlText(val);
  const preserveSpace =
    clean.length > 0 && (clean.startsWith(' ') || clean.endsWith(' '))
      ? ' xml:space="preserve"'
      : '';
  return `<c r="${ref}"${sPart} t="inlineStr"><is><t${preserveSpace}>${escapeXml(clean)}</t></is></c>`;
}

function buildNumericCellXml(
  ref: string,
  val: number,
  styleId: string | null,
): string {
  const sPart = styleId !== null ? ` s="${styleId}"` : '';
  return `<c r="${ref}"${sPart}><v>${val}</v></c>`;
}

function needsTextFormat(colNum: number, val: string): boolean {
  if (TEXT_PRESERVING_COLS.has(colNum)) return true;
  return /^0\d+$/.test(val);
}

function isRow1NonEmpty(row1: ParsedRow | undefined): boolean {
  if (!row1) return false;
  for (let c = 1; c <= RGERAL_COLUMN_HEADERS.length; c++) {
    const cell = row1.cells.get(c);
    if (cell && cell.textValue.trim() !== '') {
      return true;
    }
  }
  return false;
}

function createFreshRGeralWorksheetXml(): string {
  const colsXml = RGERAL_COLUMN_HEADERS.map((h, i) => {
    const w = Math.max(h.length + 4, 14);
    return `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`;
  }).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1:T1"/>' +
    '<sheetViews><sheetView workbookViewId="0">' +
    '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
    '</sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    `<cols>${colsXml}</cols>` +
    '<sheetData/>' +
    '</worksheet>'
  );
}

function normalizeZipPath(target: string): string {
  const clean = target.replace(/^\/+/, '');
  return clean.startsWith('xl/') ? clean : `xl/${clean}`;
}

export async function fillRGeralTemplateOOXML(params: {
  modeloBuffer: Buffer;
  modeloFileName: string;
  records: RGeralRowData[];
}): Promise<RGeralFillResult> {
  const { modeloBuffer, modeloFileName, records } = params;
  const lowerName = (modeloFileName || '').toLowerCase();
  const isXlsm = lowerName.endsWith('.xlsm');
  const extLabel = isXlsm ? '.xlsm' : '.xlsx';
  const outFileName = `Relatorio Geral${extLabel}`;

  console.log(
    `[RGeral] Modelo recebido: ${modeloFileName || 'modelo' + extLabel}, ${modeloBuffer.length} bytes, extensão: ${extLabel}`,
  );

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(modeloBuffer);
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    console.error(`[RGeral] ERRO: Falha ao abrir arquivo modelo como OOXML ZIP (${detail})`);
    throw new Error(
      `Arquivo da Planilha Modelo RGeral inválido ou não suportado (${detail}). Envie um arquivo .xlsx ou .xlsm válido.`,
    );
  }

  const inputHasVba = Boolean(zip.file('xl/vbaProject.bin'));
  const zipFileNames = Object.keys(zip.files);
  const inputPivotCount = zipFileNames.filter((n) =>
    /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(n),
  ).length;
  const inputChartCount = zipFileNames.filter((n) =>
    /^xl\/charts\/chart\d+\.xml$/i.test(n),
  ).length;

  const workbookEntry = zip.file('xl/workbook.xml');
  const relsEntry = zip.file('xl/_rels/workbook.xml.rels');
  const contentTypesEntry = zip.file('[Content_Types].xml');

  if (!workbookEntry || !relsEntry || !contentTypesEntry) {
    console.error('[RGeral] ERRO: Estrutura OOXML incompleta (workbook.xml / rels / Content_Types ausente)');
    throw new Error('Estrutura OOXML inválida no arquivo modelo enviado.');
  }

  let workbookXml = await workbookEntry.async('string');
  let relsXml = await relsEntry.async('string');
  let contentTypesXml = await contentTypesEntry.async('string');

  const ssEntry = zip.file('xl/sharedStrings.xml');
  const sharedStrings = parseSharedStrings(ssEntry ? await ssEntry.async('string') : null);

  const stylesEntry = zip.file('xl/styles.xml');
  const stylesCtx = createStylesContext(stylesEntry ? await stylesEntry.async('string') : null);

  // Mapear relacionamentos rId -> Target
  const relTargetMap = new Map<string, string>();
  const existingRIds = new Set<string>();
  const relRegex = /<Relationship\b([^>]*?)\/?>(?:<\/Relationship>)?/g;
  let relM: RegExpExecArray | null;
  while ((relM = relRegex.exec(relsXml)) !== null) {
    const id = getAttr(relM[1], 'Id');
    const target = getAttr(relM[1], 'Target');
    if (id) {
      existingRIds.add(id);
      if (target) relTargetMap.set(id, target);
    }
  }

  // Listar todas as abas originais e localizar aba "RGeral" em xl/workbook.xml
  const originalSheetNames: string[] = [];
  let targetSheetZipPath: string | null = null;
  let targetSheetId: string | null = null;
  let maxSheetId = 0;
  const sheetTagRegex = /<sheet\b([^>]*?)\/?>(?:<\/sheet>)?/g;
  let sm: RegExpExecArray | null;
  while ((sm = sheetTagRegex.exec(workbookXml)) !== null) {
    const attrs = sm[1];
    const nameRaw = getAttr(attrs, 'name') ?? '';
    const decodedName = decodeXmlEntities(nameRaw);
    if (decodedName) originalSheetNames.push(decodedName);
    const sheetIdRaw = getAttr(attrs, 'sheetId');
    const rId = getAttr(attrs, 'r:id') ?? getAttr(attrs, 'id');
    if (sheetIdRaw) {
      const sid = parseInt(sheetIdRaw, 10);
      if (Number.isFinite(sid) && sid > maxSheetId) maxSheetId = sid;
    }
    if (decodedName === 'RGeral' && rId) {
      const relTarget = relTargetMap.get(rId);
      if (relTarget) {
        targetSheetZipPath = normalizeZipPath(relTarget);
        targetSheetId = sheetIdRaw;
      }
    }
  }

  console.log(`[RGeral] Abas do modelo original: ${JSON.stringify(originalSheetNames)}`);
  const sheetFound = Boolean(targetSheetZipPath && zip.file(targetSheetZipPath));
  console.log(`[RGeral] Aba 'RGeral' encontrada: ${sheetFound ? 'sim' : 'não'}`);

  let sheetXml: string;
  if (sheetFound && targetSheetZipPath) {
    sheetXml = await zip.file(targetSheetZipPath)!.async('string');
  } else {
    // Criar aba "RGeral" automaticamente preservando o resto do workbook e vbaProject.bin
    let nextSheetNum = 1;
    while (zip.file(`xl/worksheets/sheet${nextSheetNum}.xml`)) {
      nextSheetNum++;
    }
    let nextRIdNum = 1;
    while (existingRIds.has(`rId${nextRIdNum}`)) {
      nextRIdNum++;
    }
    const nextRId = `rId${nextRIdNum}`;
    const nextSheetId = String(maxSheetId + 1);
    targetSheetId = nextSheetId;
    targetSheetZipPath = `xl/worksheets/sheet${nextSheetNum}.xml`;

    if (workbookXml.includes('</sheets>')) {
      workbookXml = workbookXml.replace(
        '</sheets>',
        `<sheet name="RGeral" sheetId="${nextSheetId}" r:id="${nextRId}"/></sheets>`,
      );
      zip.file('xl/workbook.xml', workbookXml);
    } else {
      throw new Error('Tag <sheets> não encontrada em xl/workbook.xml.');
    }

    if (relsXml.includes('</Relationships>')) {
      relsXml = relsXml.replace(
        '</Relationships>',
        `<Relationship Id="${nextRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nextSheetNum}.xml"/></Relationships>`,
      );
      zip.file('xl/_rels/workbook.xml.rels', relsXml);
    }

    const partName = `/${targetSheetZipPath}`;
    if (!contentTypesXml.includes(`PartName="${partName}"`)) {
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      );
      zip.file('[Content_Types].xml', contentTypesXml);
    }

    sheetXml = createFreshRGeralWorksheetXml();
  }

  // Extrair <sheetData> da aba RGeral
  const sheetDataSelfClose = /<sheetData\b[^>]*\/>/.exec(sheetXml);
  const sheetDataBlock = /(<sheetData\b[^>]*>)([\s\S]*?)(<\/sheetData>)/.exec(sheetXml);
  const sheetDataMatch = sheetDataBlock ?? sheetDataSelfClose;
  if (!sheetDataMatch) {
    throw new Error('Tag <sheetData> não encontrada na aba RGeral.');
  }

  const existingRows = sheetDataBlock
    ? parseSheetDataRows(sheetDataBlock[2], sharedStrings)
    : new Map<number, ParsedRow>();

  // Linha 1: manter cabeçalhos existentes; se estiver vazia, escrever os 20 cabeçalhos na linha 1
  const existingRow1 = existingRows.get(1);
  const row1Populated = isRow1NonEmpty(existingRow1);
  const effectiveRow1Headers = row1Populated
    ? RGERAL_COLUMN_HEADERS.map((_, idx) => existingRow1!.cells.get(idx + 1)?.textValue ?? '')
    : [...RGERAL_COLUMN_HEADERS];

  console.log(`[RGeral] Cabeçalhos na linha 1: ${JSON.stringify(effectiveRow1Headers)}`);
  console.log(`[RGeral] Linhas de dados a colar: ${records.length}`);

  const rebuiltRows: string[] = [];

  if (row1Populated && existingRow1) {
    rebuiltRows.push(existingRow1.rawRowXml);
  } else {
    const hdrCells: string[] = [];
    for (let c = 1; c <= RGERAL_COLUMN_HEADERS.length; c++) {
      const ref = `${colNumToLetters(c)}1`;
      const existingCellStyle = existingRow1?.cells.get(c)?.styleAttr ?? null;
      hdrCells.push(
        buildInlineStrCellXml(ref, RGERAL_COLUMN_HEADERS[c - 1], existingCellStyle),
      );
    }
    if (existingRow1) {
      for (const [colNum, cell] of Array.from(existingRow1.cells.entries()).sort((a, b) => a[0] - b[0])) {
        if (colNum > RGERAL_COLUMN_HEADERS.length) {
          hdrCells.push(cell.rawXml);
        }
      }
    }
    const row1Attrs = existingRow1
      ? setAttr(existingRow1.rowAttrs, 'r', '1')
      : 'r="1" spans="1:20"';
    rebuiltRows.push(`<row ${row1Attrs.trim()}>${hdrCells.join('')}</row>`);
  }

  // Linha 2: NÃO sobrescrever (preservar integralmente caso exista no modelo)
  const existingRow2 = existingRows.get(2);
  if (existingRow2) {
    rebuiltRows.push(existingRow2.rawRowXml);
  }

  // Linhas 3 em diante: colar os dados mapeados nas colunas A..T (1..20) a partir da célula A3
  let maxRowWritten = existingRow2 ? 2 : 1;
  let maxColUsed = 20;

  for (let i = 0; i < records.length; i++) {
    const rowNum = 3 + i;
    maxRowWritten = rowNum;
    const rec = records[i];
    const rowValues: Array<string | number> = [
      rec.instalacao,
      rec.medidor,
      rec.registro,
      rec.rzUlLv,
      rec.ul,
      rec.fotos,
      rec.tipo,
      rec.matricula,
      rec.cod,
      rec.data,
      rec.hora,
      rec.endereco,
      rec.leitAnterior,
      rec.leitAtual,
      rec.consumo,
      rec.impedimento,
      rec.simulacao,
      rec.digitacao,
      rec.apres,
      rec.consistencia,
    ];

    const prevRow = existingRows.get(rowNum);
    const rowCellsXml: string[] = [];

    for (let c = 1; c <= rowValues.length; c++) {
      const val = rowValues[c - 1];
      const ref = `${colNumToLetters(c)}${rowNum}`;
      const origStyle = prevRow?.cells.get(c)?.styleAttr ?? null;

      if (typeof val === 'number' && Number.isFinite(val)) {
        rowCellsXml.push(buildNumericCellXml(ref, val, origStyle));
      } else {
        const strVal = String(val ?? '');
        if (strVal === '') {
          if (origStyle !== null) {
            rowCellsXml.push(`<c r="${ref}" s="${origStyle}"/>`);
          }
        } else {
          const styleId = needsTextFormat(c, strVal)
            ? stylesCtx.getTextStyleId(origStyle)
            : origStyle;
          rowCellsXml.push(buildInlineStrCellXml(ref, strVal, styleId));
        }
      }
    }

    // Preservar células em colunas > 20 (U em diante) caso existam no modelo
    if (prevRow) {
      for (const [colNum, cell] of Array.from(prevRow.cells.entries()).sort((a, b) => a[0] - b[0])) {
        if (colNum > 20) {
          if (colNum > maxColUsed) maxColUsed = colNum;
          rowCellsXml.push(cell.rawXml);
        }
      }
    }

    const rowAttrs = prevRow
      ? setAttr(prevRow.rowAttrs, 'r', String(rowNum))
      : `r="${rowNum}" spans="1:20"`;
    rebuiltRows.push(`<row ${rowAttrs.trim()}>${rowCellsXml.join('')}</row>`);
  }

  // Limpar linhas excedentes de A..T caso o modelo já tivesse dados a partir de A3, preservando colunas > 20
  const trailingRowNums = Array.from(existingRows.keys())
    .filter((r) => r >= 3 + records.length)
    .sort((a, b) => a - b);

  for (const r of trailingRowNums) {
    const prevRow = existingRows.get(r)!;
    const extraCells = Array.from(prevRow.cells.entries())
      .filter(([colNum]) => colNum > 20)
      .sort((a, b) => a[0] - b[0]);
    if (extraCells.length > 0) {
      maxRowWritten = Math.max(maxRowWritten, r);
      for (const [colNum] of extraCells) {
        if (colNum > maxColUsed) maxColUsed = colNum;
      }
      const rowAttrs = setAttr(prevRow.rowAttrs, 'r', String(r));
      rebuiltRows.push(
        `<row ${rowAttrs.trim()}>${extraCells.map(([, c]) => c.rawXml).join('')}</row>`,
      );
    }
  }

  const newSheetDataXml = `<sheetData>${rebuiltRows.join('')}</sheetData>`;
  sheetXml =
    sheetXml.slice(0, sheetDataMatch.index) +
    newSheetDataXml +
    sheetXml.slice(sheetDataMatch.index + sheetDataMatch[0].length);

  // Atualizar <dimension ref="..."/> garantindo que inclua até o final dos dados colados
  const newDimRef = `A1:${colNumToLetters(maxColUsed)}${Math.max(maxRowWritten, 3)}`;
  if (/<dimension\b[^>]*\/>/.test(sheetXml)) {
    sheetXml = sheetXml.replace(
      /<dimension\b([^>]*?)\/>/,
      (_, attrs) => `<dimension ${setAttr(attrs, 'ref', newDimRef).trim()}/>`,
    );
  }

  zip.file(targetSheetZipPath, sheetXml);

  const updatedStylesXml = stylesCtx.finalizeStylesXml();
  if (updatedStylesXml && stylesEntry) {
    zip.file('xl/styles.xml', updatedStylesXml);
  }

  // Se xl/calcChain.xml referenciava células antigas sobrescritas em RGeral (A3:T...), remover apenas essas referências
  const calcChainEntry = zip.file('xl/calcChain.xml');
  if (calcChainEntry && targetSheetId) {
    let calcChainXml = await calcChainEntry.async('string');
    const cRegex = /<c\b([^>]*?)\/>/g;
    let calcModified = false;
    calcChainXml = calcChainXml.replace(cRegex, (full, attrs: string) => {
      const iAttr = getAttr(attrs, 'i');
      const rAttr = getAttr(attrs, 'r');
      if (iAttr === targetSheetId && rAttr) {
        const parsed = parseCellRef(rAttr);
        if (parsed && parsed.rowNum >= 3 && parsed.colNum <= 20) {
          calcModified = true;
          return '';
        }
      }
      return full;
    });
    if (calcModified) {
      if (!/<c\b/.test(calcChainXml)) {
        zip.remove('xl/calcChain.xml');
        relsXml = relsXml.replace(
          /<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/?>(?:<\/Relationship>)?/g,
          '',
        );
        zip.file('xl/_rels/workbook.xml.rels', relsXml);
        contentTypesXml = contentTypesXml.replace(
          /<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/?>(?:<\/Override>)?/g,
          '',
        );
        zip.file('[Content_Types].xml', contentTypesXml);
      } else {
        zip.file('xl/calcChain.xml', calcChainXml);
      }
    }
  }

  // Garantir recálculo automático nas demais abas (Resumo, Dashboard, etc.) ao abrir no Excel
  if (/<calcPr\b[^>]*\/?>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b([^>]*?)(\/?)>/, (_, attrs: string, slash: string) => {
      const updatedAttrs = setAttr(attrs, 'fullCalcOnLoad', '1');
      return `<calcPr ${updatedAttrs.trim()}${slash}>`;
    });
    zip.file('xl/workbook.xml', workbookXml);
  } else if (workbookXml.includes('</workbook>')) {
    if (workbookXml.includes('<extLst')) {
      workbookXml = workbookXml.replace(
        '<extLst',
        '<calcPr fullCalcOnLoad="1"/><extLst',
      );
    } else {
      workbookXml = workbookXml.replace(
        '</workbook>',
        '<calcPr fullCalcOnLoad="1"/></workbook>',
      );
    }
    zip.file('xl/workbook.xml', workbookXml);
  }

  // Se existirem caches de Tabela Dinâmica (PivotTables), marcar refreshOnLoad="1" mantendo estrutura intacta
  for (const fname of zipFileNames) {
    if (/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(fname)) {
      const entry = zip.file(fname);
      if (entry) {
        const xml = await entry.async('string');
        const updated = xml.replace(
          /<pivotCacheDefinition\b([^>]*?)>/,
          (_, attrs: string) => `<pivotCacheDefinition ${setAttr(attrs, 'refreshOnLoad', '1').trim()}>`,
        );
        if (updated !== xml) zip.file(fname, updated);
      }
    }
  }

  console.log(
    `[RGeral] Dados colados como VALORES (não fórmulas): ${records.length} linhas x 20 colunas`,
  );
  if (records.length > 0) {
    console.log(`[RGeral] Dados colados a partir de A3 até T${2 + records.length}`);
  } else {
    console.log('[RGeral] Dados colados a partir de A3 (0 linhas)');
  }

  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  console.log(
    `[RGeral] Arquivo gerado: ${outFileName}, ${outBuffer.length} bytes (modelo enviado: ${modeloBuffer.length} bytes)`,
  );

  // Validação obrigatória pós-gravação: abas preservadas, vbaProject.bin, PivotTables, Gráficos e valores estáticos (sem fórmulas)
  try {
    const verifyZip = await JSZip.loadAsync(outBuffer);
    const outZipFiles = Object.keys(verifyZip.files);
    const hasVbaAfter = Boolean(verifyZip.file('xl/vbaProject.bin'));
    const outPivotCount = outZipFiles.filter((n) =>
      /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(n),
    ).length;
    const outChartCount = outZipFiles.filter((n) =>
      /^xl\/charts\/chart\d+\.xml$/i.test(n),
    ).length;

    console.log(`[RGeral] PivotTables preservadas: ${outPivotCount}`);
    console.log(`[RGeral] Gráficos preservados: ${outChartCount}`);

    if (outPivotCount < inputPivotCount) {
      throw new Error(`PivotTables perdidas na gravação (${outPivotCount} < ${inputPivotCount}).`);
    }
    if (outChartCount < inputChartCount) {
      throw new Error(`Gráficos perdidos na gravação (${outChartCount} < ${inputChartCount}).`);
    }

    if (isXlsm) {
      console.log(`[RGeral] vbaProject.bin preservado: ${hasVbaAfter ? 'sim' : 'não'}`);
      if (inputHasVba && !hasVbaAfter) {
        throw new Error('xl/vbaProject.bin foi perdido após a gravação do arquivo .xlsm.');
      }
    } else {
      console.log('[RGeral] vbaProject.bin preservado: não-aplicável');
    }

    const verifySheet = verifyZip.file(targetSheetZipPath);
    if (!verifySheet) {
      throw new Error(`Aba RGeral (${targetSheetZipPath}) ausente após gravação.`);
    }
    const verifySheetText = await verifySheet.async('string');
    if (!verifySheetText.includes('<worksheet') || !verifySheetText.includes('</worksheet>')) {
      throw new Error('XML da aba RGeral malformado após gravação.');
    }

    const reopenWb = XLSX.read(outBuffer, {
      type: 'buffer',
      bookVBA: isXlsm,
    });
    const finalSheetNames = [...reopenWb.SheetNames];
    const missingSheets = originalSheetNames.filter(
      (origName) => !finalSheetNames.includes(origName),
    );
    console.log(`[RGeral] Abas preservadas no arquivo final: ${JSON.stringify(finalSheetNames)}`);
    console.log(
      `[RGeral] Match de abas (original = final): ${missingSheets.length === 0 ? 'sim' : 'não'}`,
    );
    if (missingSheets.length > 0) {
      console.log(`[RGeral] Abas perdidas: ${JSON.stringify(missingSheets)}`);
      for (const missing of missingSheets) {
        console.error(`[ERRO] Aba ${missing} perdida na gravação`);
      }
      throw new Error(
        `Abas originais perdidas na gravação: ${missingSheets.join(', ')}`,
      );
    }

    if (!finalSheetNames.includes('RGeral')) {
      console.error('[ERRO] Aba RGeral perdida na gravação');
      throw new Error('Aba "RGeral" não encontrada ao reabrir o arquivo gerado.');
    }

    // Confirmar que nenhuma célula de A3:T... em RGeral contém fórmula (<f>)
    const reopenWs = reopenWb.Sheets['RGeral'];
    if (records.length > 0 && reopenWs) {
      for (let r = 3; r <= 2 + Math.min(records.length, 50); r++) {
        for (let c = 1; c <= 20; c++) {
          const addr = `${colNumToLetters(c)}${r}`;
          const cellObj = reopenWs[addr] as XLSX.CellObject | undefined;
          if (cellObj?.f) {
            throw new Error(
              `Célula ${addr} na aba RGeral contém fórmula ("=${cellObj.f}") em vez de valor estático.`,
            );
          }
        }
      }
    }

    console.log('[RGeral] Reabertura pós-gravação: sucesso');
  } catch (verifyErr) {
    const errDetail = verifyErr instanceof Error ? verifyErr.message : String(verifyErr);
    console.error('[ERRO] Arquivo gerado está corrompido');
    console.error(`[RGeral] ERRO: ${errDetail}`);
    console.log('[RGeral] Reabertura pós-gravação: falha');
    throw new Error(`Falha na validação pós-gravação de ${outFileName}: ${errDetail}`);
  }

  return {
    buffer: outBuffer,
    fileName: outFileName,
  };
}
