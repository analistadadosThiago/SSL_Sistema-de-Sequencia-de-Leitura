import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import {
  CemigValidationError,
  CONTROLE_HORARIO_HEADERS,
  FINAL_HEADERS,
  RELATORIO_GERAL_HEADERS,
  classifyTipoFromRzUlLv,
  extractULFromRzUlLv,
  processCemigWorkbook,
} from '../server/cemigProcessor.ts';
import { generateSampleCemigMatrizBuffer } from '../server/sampleCemigGenerator.ts';

async function runTests() {
  console.log('0. Testando helpers UL e TIPO...');
  assert.equal(extractULFromRzUlLv('09134696'), '1346');
  assert.equal(extractULFromRzUlLv('06102896'), '1028');
  assert.equal(extractULFromRzUlLv('00123456'), '1234');
  assert.equal(extractULFromRzUlLv('31020001'), '0200');
  assert.equal(extractULFromRzUlLv('0912'), '12');
  assert.equal(extractULFromRzUlLv('09'), '');
  assert.equal(classifyTipoFromRzUlLv('31020001'), 'Urbano');
  assert.equal(classifyTipoFromRzUlLv('31020079'), 'Urbano');
  assert.equal(classifyTipoFromRzUlLv('31020080'), 'Povoado');
  assert.equal(classifyTipoFromRzUlLv('31020089'), 'Povoado');
  assert.equal(classifyTipoFromRzUlLv('31020090'), 'Rural');
  assert.equal(classifyTipoFromRzUlLv('31020099'), 'Rural');
  assert.equal(classifyTipoFromRzUlLv(''), '');

  console.log('1. Gerando planilha Matriz de exemplo (layout CEMIG)...');
  const sampleBuf = await generateSampleCemigMatrizBuffer();
  assert.ok(sampleBuf.length > 0, 'Sample buffer deve ser não-vazio');

  console.log('2. Executando pipeline completo (20 colunas + extras brutos, Apres., UL 3..6, ordenação)...');
  const res = await processCemigWorkbook(sampleBuf, null);

  assert.equal(res.fileName, 'Matriz_ControleDeSequenciaLeitura.xlsx');
  assert.ok(res.stats.finalCols > 20, `Deve conter as 20 colunas principais mais colunas extras preservadas (recebido: ${res.stats.finalCols})`);
  assert.deepEqual(
    res.preview.headers.slice(0, 20),
    [...FINAL_HEADERS],
    'As 20 primeiras colunas devem seguir exatamente a ordem definida',
  );
  assert.ok(
    !res.preview.headers.includes('CNA'),
    'Coluna CNA NÃO deve existir',
  );
  assert.ok(
    !res.preview.headers.includes('Leiturista'),
    'Coluna Leiturista NÃO deve existir na tabela final',
  );
  assert.ok(
    !res.preview.headers.includes('c') &&
      !res.preview.headers.includes('A Exec.') &&
      !res.preview.headers.includes('Exec.'),
    'Colunas c, A Exec. e Exec. NÃO devem existir',
  );
  assert.ok(
    res.preview.headers.includes('Seqüência de Leitura Executada'),
    'Deve preservar coluna Seqüência de Leitura Executada',
  );
  assert.ok(
    res.preview.headers.includes('Contrato_Aux'),
    'Deve preservar coluna Contrato_Aux sem descartar informações importadas',
  );

  // 10 registros no sample: 1 duplicado removido, 1 excluído pelo filtro 99Z -> 9 linhas finais
  assert.equal(res.stats.finalRows, 9, 'Deve conter 9 linhas válidas após deduplicação e filtro');
  assert.equal(res.stats.uniqueInstalacoesCount, 9, 'Quantidade Instalações = 9');
  assert.equal(res.stats.urbanoCount, 5, 'Urbano = 5');
  assert.equal(res.stats.povoadoCount, 2, 'Povoado = 2');
  assert.equal(res.stats.ruralCount, 2, 'Rural = 2');

  // Verificar ordenação por Leiturista (4MW, índice 5) e DATA crescente (índice 8): 06/11/2026 -> 07/11/2026 -> 12/11/2026
  const rows4MW = res.preview.rows.filter((r) => r[5] === '4MW');
  assert.equal(rows4MW.length, 3, 'Devem existir 3 registros do leiturista 4MW');
  assert.deepEqual(
    rows4MW.map((r) => r[8]),
    ['06/11/2026', '07/11/2026', '12/11/2026'],
    'Datas dentro de 4MW devem estar em dd/mm/aaaa e ordem crescente',
  );

  // Verificar agrupamento contíguo por Leiturista (MATRÍCULA índice 5)
  const seenGroups = new Set<string>();
  let prevGroup = '';
  for (const row of res.preview.rows) {
    const grp = row[5];
    if (grp !== prevGroup) {
      assert.ok(!seenGroups.has(grp), `Grupo ${grp} não pode aparecer fragmentado`);
      seenGroups.add(grp);
      prevGroup = grp;
    }
  }

  // Verificar amostra principal: MARIA CELESTE M MORAIS / ARL211115451 / 23873 -> 24182 (309)
  const mariaRow = res.preview.rows.find((r) => r[0] === 'MARIA CELESTE M MORAIS');
  assert.ok(mariaRow, 'Registro MARIA CELESTE M MORAIS deve estar presente');
  assert.equal(mariaRow[1], 'ARL211115451', 'MEDIDOR');
  assert.equal(mariaRow[2], '00060', 'REGISTRO com zero à esquerda');
  assert.equal(mariaRow[3], '31020001', 'RZ_UL_LV');
  assert.equal(mariaRow[4], '0200', 'UL (posições 3..6 de 31020001 preservando zero à esquerda)');
  assert.equal(mariaRow[5], '4MW', 'MATRÍCULA');
  assert.equal(mariaRow[6], 'Urbano', 'TIPO');
  assert.equal(mariaRow[7], '00', 'COD');
  assert.equal(mariaRow[8], '06/11/2026', 'DATA dd/mm/aaaa');
  assert.equal(mariaRow[9], '10:03:55 AM', 'HORA preservada');
  assert.equal(mariaRow[10], 'RUA DAS FLORES, 142 - CENTRO', 'ENDEREÇO');
  assert.equal(mariaRow[11], '23873', 'Leit. Anterior');
  assert.equal(mariaRow[12], '24182', 'Leit. Executada');
  assert.equal(mariaRow[13], '309', 'CONSUMO = 309');
  assert.equal(mariaRow[14], 'Faturada', 'NOSB´IMPEDIMENTO');
  assert.equal(mariaRow[15], 'N00', 'DIGITAÇÃO');
  assert.equal(mariaRow[16], 'N00', 'Apres. (valor != "00" convertido para "N00")');
  assert.equal(mariaRow[17], 'Faturada', 'CONSISTÊNCIA');
  assert.equal(mariaRow[18], 'Faturada', 'NOSB´SIMULAÇÃO');
  assert.equal(mariaRow[19], 'N-OK', 'FOTOS');

  const bomPrecoRow = res.preview.rows.find((r) => r[0] === 'SUPERMERCADO BOM PRECO LTDA');
  assert.ok(bomPrecoRow, 'Registro SUPERMERCADO BOM PRECO LTDA deve estar presente');
  assert.equal(bomPrecoRow[3], '09134699', 'RZ_UL_LV');
  assert.equal(bomPrecoRow[4], '1346', 'UL (posições 3..6 de 09134699 = 1346)');

  const seqColIdx = res.preview.headers.indexOf('Seqüência de Leitura Executada');
  assert.equal(mariaRow[seqColIdx], '00136', 'Seqüência preservada com zeros à esquerda');

  console.log('3. Validando estrutura física do .xlsx gerado e paridade completa de linhas/colunas...');
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.load(new Uint8Array(res.outputBuffer).buffer);

  const wsMain = outWb.getWorksheet('ControleLeitura');
  const wsFiltros = outWb.getWorksheet('Filtros');
  const wsOldP3 = outWb.getWorksheet('Planilha3');
  assert.ok(wsMain, 'Aba ControleLeitura deve existir');
  assert.ok(wsFiltros, 'Aba Filtros deve existir');
  assert.ok(!wsOldP3, 'Aba Planilha3 não deve mais existir (renomeada para Filtros)');

  assert.equal(wsMain.actualRowCount - 1, res.preview.rows.length, '.xlsx deve ter exatamente todas as linhas tratadas');
  assert.equal(wsMain.actualColumnCount, res.preview.headers.length, '.xlsx deve ter exatamente todas as colunas tratadas');

  assert.equal(wsFiltros.getCell('D1').value, 'Consistencia');
  assert.equal(wsFiltros.getCell('G1').value, 'Nosb');
  assert.equal(wsFiltros.getCell('J1').value, 'Fotos');

  const d1Fill = wsFiltros.getCell('D1').fill as { fgColor?: { argb?: string } } | undefined;
  const g1Fill = wsFiltros.getCell('G1').fill as { fgColor?: { argb?: string } } | undefined;
  const j1Fill = wsFiltros.getCell('J1').fill as { fgColor?: { argb?: string } } | undefined;
  assert.ok(d1Fill?.fgColor?.argb, 'Coluna D em Filtros deve possuir cor de preenchimento');
  assert.ok(g1Fill?.fgColor?.argb, 'Coluna G em Filtros deve possuir cor de preenchimento');
  assert.ok(j1Fill?.fgColor?.argb, 'Coluna J em Filtros deve possuir cor de preenchimento');
  assert.notEqual(d1Fill?.fgColor?.argb, g1Fill?.fgColor?.argb);
  assert.notEqual(g1Fill?.fgColor?.argb, j1Fill?.fgColor?.argb);

  const table = wsMain.getTable('Tabela1');
  assert.ok(table, 'ListObject "Tabela1" deve existir na aba principal');
  assert.ok(!wsMain.hasMerges, 'Não deve haver células mescladas na saída');

  const f2 = wsMain.getCell('F2').value as { formula?: string; result?: string };
  const n2 = wsMain.getCell('N2').value as { formula?: string; result?: number };
  const r2 = wsMain.getCell('R2').value as { formula?: string; result?: string };
  const s2 = wsMain.getCell('S2').value as { formula?: string; result?: string };
  const t2 = wsMain.getCell('T2').value as { formula?: string; result?: string };

  assert.equal(f2.formula, 'VLOOKUP(D2,Filtros!$A$1:$B$50000,2,0)');
  assert.equal(f2.result, '4MW');
  assert.equal(n2.formula, 'SUM(M2-L2)');
  assert.equal(n2.result, 309);
  assert.equal(r2.formula, 'IFERROR(VLOOKUP(B2,Filtros!$D$1:$E$50000,2,0),"Faturada")');
  assert.equal(s2.formula, 'IFERROR(VLOOKUP(B2,Filtros!$G$1:$H$50000,2,0),"Faturada")');
  assert.equal(t2.formula, 'IFERROR(VLOOKUP(B2,Filtros!$J$1:$K$50000,2,0),"N-OK")');

  const c2 = wsMain.getCell('C2');
  assert.equal(c2.value, '00060', 'Célula C2 mantém texto "00060"');
  assert.equal(wsMain.getCell('E2').value, '0200', 'Célula E2 (UL)');
  assert.equal(wsMain.getCell('G2').value, 'Urbano', 'Célula G2 (TIPO)');
  assert.equal(wsMain.getCell('Q2').value, 'N00', 'Célula Q2 (Apres. = "N00" pois bruto != "00")');
  assert.equal(c2.alignment?.horizontal, 'left');
  assert.equal(c2.alignment?.vertical, 'bottom');
  assert.ok(!c2.alignment?.wrapText, 'wrapText não deve estar ativado');

  console.log('4. Testando exclusão das colunas "c", "A Exec.", "Exec." e "Leiturista"...');
  const exclWb = new ExcelJS.Workbook();
  await exclWb.xlsx.load(new Uint8Array(sampleBuf).buffer);
  const exclWs = exclWb.getWorksheet('Leituras')!;
  exclWs.getRow(11).getCell(14).value = 'c';
  exclWs.getRow(18).getCell(14).value = 'val_c';
  exclWs.getRow(11).getCell(15).value = 'A Exec.';
  exclWs.getRow(18).getCell(15).value = 'val_aexec';
  exclWs.getRow(11).getCell(16).value = 'Exec.';
  exclWs.getRow(18).getCell(16).value = 'val_exec';
  const exclBuf = Buffer.from(await exclWb.xlsx.writeBuffer());
  const exclRes = await processCemigWorkbook(exclBuf, null);
  assert.ok(!exclRes.preview.headers.includes('c'), 'Coluna "c" deve ser removida');
  assert.ok(!exclRes.preview.headers.includes('A Exec.'), 'Coluna "A Exec." deve ser removida');
  assert.ok(!exclRes.preview.headers.includes('Exec.'), 'Coluna "Exec." deve ser removida');
  assert.ok(!exclRes.preview.headers.includes('Leiturista'), 'Coluna "Leiturista" deve ser removida');

  console.log('5. Testando validação contra planilha inválida...');
  const badWb = new ExcelJS.Workbook();
  const badWs = badWb.addWorksheet('Plan1');
  badWs.getCell('A1').value = 'Produto';
  badWs.getCell('B1').value = 'Preço';
  const badBuf = Buffer.from(await badWb.xlsx.writeBuffer());

  await assert.rejects(
    async () => processCemigWorkbook(badBuf, null),
    CemigValidationError,
  );

  console.log('6. Testando arquivo corrompido...');
  await assert.rejects(
    async () =>
      processCemigWorkbook(
        Buffer.from([0x50, 0x4b, 0x03, 0x04, 0xff, 0xff, 0x00, 0x00]),
        null,
      ),
    (err: unknown) =>
      err instanceof CemigValidationError && err.code === 'CORRUPT_FILE',
  );

  console.log('7. Testando escala completa (500 registros + espaçamento irregular Leiturista)...');
  const largeWb = new ExcelJS.Workbook();
  const largeWs = largeWb.addWorksheet('Leituras');
  const setL = (r: number, c: number, v: string) => {
    const cell = largeWs.getRow(r).getCell(c);
    cell.value = v;
    cell.numFmt = '@';
  };
  setL(1, 1, 'CEMIG Distribuição S.A.');
  setL(1, 10, 'Pág. 1/50');
  setL(2, 1, 'Seqüência de Leitura Executada');
  setL(11, 1, 'Consumidora');
  setL(11, 2, 'Medidor');
  setL(11, 3, 'T. Reg');
  setL(11, 4, 'Seqüência de Leitura Executada');
  setL(11, 5, 'Contrato_Aux');
  setL(11, 6, 'Data');
  setL(11, 7, 'Hora');
  setL(13, 7, 'Leit. Ant.');
  setL(11, 8, 'Tipo_Aux1');
  setL(11, 9, 'Tipo_Aux2');
  setL(11, 10, 'Leit. Exec');
  setL(11, 12, 'Qtd Dig.');
  setL(11, 13, 'OC');
  setL(13, 13, 'Apres.');

  let lRow = 18;
  for (let i = 0; i < 500; i++) {
    setL(lRow - 5, 1, `Unidade de Leitura: 310200${String(i % 90).padStart(2, '0')}`);
    setL(lRow - 4, 1, i % 2 === 0 ? 'Leiturista:  4MW' : 'Leiturista: 5GL');
    setL(lRow, 1, `INSTALACAO_${String(i + 1).padStart(5, '0')}`);
    setL(lRow, 2, `MED${String(i + 1).padStart(6, '0')}`);
    setL(lRow, 3, '00060');
    setL(lRow, 4, String(i + 1).padStart(5, '0'));
    setL(lRow, 5, '4600018294');
    setL(lRow, 6, '11/17/2026');
    setL(lRow, 7, '10:00:00 AM');
    setL(lRow + 2, 7, '100');
    setL(lRow, 8, 'CONV');
    setL(lRow, 9, 'NORMAL');
    setL(lRow, 10, String(110 + i));
    setL(lRow, 11, '');
    setL(lRow, 12, '05');
    setL(lRow, 13, '00');
    setL(lRow + 2, 13, `RUA TESTE, ${i + 1}`);
    lRow += 8;
  }
  const largeBuf = Buffer.from(await largeWb.xlsx.writeBuffer());
  const largeRes = await processCemigWorkbook(largeBuf, null);
  assert.equal(largeRes.stats.finalRows, 500, 'Todas as 500 linhas devem ser exportadas sem truncamento');
  assert.ok(largeRes.preview.headers.includes('Contrato_Aux'), 'Deve detectar corretamente cabeçalho mesmo com linhas 3..10 vazias');
  assert.equal(largeRes.preview.rows[0][7], '00', 'COD extraído corretamente na escala');
  assert.equal(largeRes.preview.rows[0][16], 'N00', 'Apres. normalizado para N00 na escala');

  const verifyLargeWb = new ExcelJS.Workbook();
  await verifyLargeWb.xlsx.load(new Uint8Array(largeRes.outputBuffer).buffer);
  const verifyLargeWs = verifyLargeWb.getWorksheet('ControleLeitura')!;
  assert.equal(verifyLargeWs.actualRowCount - 1, 500, '.xlsx final deve conter todas as 500 linhas gravadas');
  assert.equal(verifyLargeWs.actualColumnCount, largeRes.preview.headers.length, '.xlsx final deve conter todas as colunas');

  console.log('8. Testando molde Leitura.xlsx (bloco 3 linhas por medidor + múltiplas páginas + caso de aceite)...');
  const mWb = new ExcelJS.Workbook();
  const mWs = mWb.addWorksheet('Leitura');
  const setM = (r: number, c: number, v: string | number) => {
    const cell = mWs.getRow(r).getCell(c);
    cell.value = v;
    if (typeof v === 'string') cell.numFmt = '@';
  };

  const addLeituraPageHeader = (
    startRow: number,
    pageNum: number,
    totalPages: number,
    ulCode: string,
    leituristaCode: string,
  ): number => {
    setM(startRow, 1, 'CEMIG Distribuição S.A.');
    setM(startRow, 12, `Pág. ${pageNum}/${totalPages}`);
    setM(startRow + 1, 1, 'Seqüência de Leitura Executada');
    setM(startRow + 2, 1, 'Contrato');
    setM(startRow + 2, 2, 'Mês/Ano');
    setM(startRow + 2, 4, 'Unidade de Leitura Inicial');
    setM(startRow + 2, 7, 'Unidade de Leitura Final');
    setM(startRow + 3, 1, '4600018294');
    setM(startRow + 3, 2, '11/2026');
    setM(startRow + 4, 1, `Unidade de Leitura: ${ulCode}`);
    setM(startRow + 4, 6, 'Data Início: 17/11/2026');
    setM(startRow + 5, 1, `Leiturista: ${leituristaCode}`);
    setM(startRow + 5, 6, 'Tipo de Leitura: CONVENCIONAL');
    const hRow = startRow + 6;
    setM(hRow, 1, 'Unidade');
    setM(hRow, 2, 'Medidor');
    setM(hRow, 3, 'T. Reg');
    setM(hRow, 4, 'OC');
    setM(hRow, 5, 'c');
    setM(hRow, 6, 'Data');
    setM(hRow, 7, 'Hora');
    setM(hRow, 8, 'A Exec.');
    setM(hRow, 9, 'Exec.');
    setM(hRow, 10, 'Leit. Ant.');
    setM(hRow, 11, 'Leit. Exec');
    setM(hRow, 12, 'Imp.');
    setM(hRow, 13, 'Apres.');
    setM(hRow + 1, 1, 'Consumidora');
    setM(hRow + 2, 1, 'Cliente');
    setM(hRow + 2, 7, 'Endereço');
    setM(hRow + 2, 13, 'Qtd Dig.');
    return hRow + 3;
  };

  let curR = addLeituraPageHeader(1, 1, 2, '09102801', '5JL');
  // Registro de aceite exato (Serial Excel 46343 = 17/11/2026)
  setM(curR, 1, '2.796.450.018-02');
  setM(curR, 2, 'APG176044248');
  setM(curR, 3, '03');
  setM(curR, 4, '1800');
  setM(curR, 5, 'X');
  setM(curR, 6, 46343);
  setM(curR, 7, '09:11');
  setM(curR, 8, '1');
  setM(curR, 9, '1');
  setM(curR, 10, '15741');
  setM(curR, 11, '15927');
  setM(curR, 12, '');
  setM(curR, 13, '00');
  setM(curR + 1, 6, '9:11:15 AM');
  setM(curR + 1, 7, 'RUA NILO PECANHA, 288 -  - CS');
  setM(curR + 2, 1, 'CARLOS AUGUSTO DE SOUZA PEREIRA');
  setM(curR + 2, 4, '123.456.789-00');
  setM(curR + 2, 13, '1');
  curR += 3;

  // Mais medidores na página 1 testando regra de Apres.: "", "0", "N00", "994", "1" -> "N00"
  const rawApresSamples = ['', '0', 'N00', '994', '1', '00', '00', '00', '00'];
  for (let i = 1; i < 10; i++) {
    setM(curR, 1, `2.796.450.0${18 + i}-02`);
    setM(curR, 2, `APG1760442${48 + i}`);
    setM(curR, 3, '03');
    setM(curR, 4, '1800');
    setM(curR, 5, 'X');
    setM(curR, 6, '17/11/2026');
    setM(curR, 7, `09:${String(11 + i).padStart(2, '0')}`);
    setM(curR, 8, String(i + 1));
    setM(curR, 9, String(i + 1));
    setM(curR, 10, String(1000 + i * 10));
    setM(curR, 11, String(1050 + i * 10));
    setM(curR, 12, i % 3 === 0 ? 'IMP1' : '');
    setM(curR, 13, rawApresSamples[i - 1]);
    setM(curR + 1, 6, `9:${String(11 + i).padStart(2, '0')}:00 AM`);
    setM(curR + 1, 8, `AV BRASIL, ${100 + i} - CENTRO`);
    setM(curR + 2, 1, `CONSUMIDOR TESTE ${i}`);
    setM(curR + 2, 4, `999.000.111-0${i}`);
    setM(curR + 2, 14, '2');
    curR += 3;
  }

  curR = addLeituraPageHeader(curR, 2, 2, '09102895', 'XYZ');
  for (let i = 0; i < 5; i++) {
    setM(curR, 1, `3.100.000.00${i}-01`);
    setM(curR, 2, `MED_RURAL_${i}`);
    setM(curR, 3, '05');
    setM(curR, 4, '00');
    setM(curR, 6, '11/17/2026');
    setM(curR, 7, `10:${String(10 + i).padStart(2, '0')}`);
    setM(curR, 8, String(i + 1));
    setM(curR, 9, String(i + 1));
    setM(curR, 10, '500');
    setM(curR, 11, '620');
    setM(curR, 12, '');
    setM(curR, 13, '00');
    setM(curR + 1, 6, `10:${String(10 + i).padStart(2, '0')}:30 AM`);
    setM(curR + 1, 7, `FAZENDA SANTO ANTONIO, KM ${i + 1}`);
    setM(curR + 2, 1, `PROPRIETARIO RURAL ${i}`);
    setM(curR + 2, 4, `00.111.222/0001-0${i}`);
    setM(curR + 2, 13, '1');
    curR += 3;
  }

  const leituraBuf = Buffer.from(await mWb.xlsx.writeBuffer());
  const leituraRes = await processCemigWorkbook(leituraBuf, null);

  assert.equal(leituraRes.stats.finalRows, 15, 'Deve retornar TODOS os 15 medidores das 2 páginas sem perder nenhum');
  assert.equal(leituraRes.stats.finalCols, 20, 'Tabela final deve conter exatamente as 20 colunas (sem Consumidora e Cliente)');
  assert.deepEqual(leituraRes.preview.headers, [...FINAL_HEADERS], 'Ordem exata das 20 colunas');
  assert.equal(leituraRes.stats.urbanoCount, 10);
  assert.equal(leituraRes.stats.ruralCount, 5);
  assert.ok(!leituraRes.preview.headers.includes('c'), 'Coluna c excluída');
  assert.ok(!leituraRes.preview.headers.includes('A Exec.'), 'Coluna A Exec. excluída');
  assert.ok(!leituraRes.preview.headers.includes('Exec.'), 'Coluna Exec. excluída');
  assert.ok(!leituraRes.preview.headers.includes('Leiturista'), 'Coluna Leiturista excluída');
  assert.ok(!leituraRes.preview.headers.includes('Consumidora'), 'Coluna Consumidora removida da tabela final');
  assert.ok(!leituraRes.preview.headers.includes('Cliente'), 'Coluna Cliente removida da tabela final');

  const carlosRow = leituraRes.preview.rows.find((r) => r[1] === 'APG176044248');
  assert.ok(carlosRow, 'Registro APG176044248 deve estar presente');
  assert.equal(carlosRow[0], '2.796.450.018-02', 'INSTALAÇÃO');
  assert.equal(carlosRow[2], '03', 'REGISTRO');
  assert.equal(carlosRow[3], '09102801', 'RZ_UL_LV');
  assert.equal(carlosRow[4], '1028', 'UL');
  assert.equal(carlosRow[5], '5JL', 'MATRÍCULA');
  assert.equal(carlosRow[6], 'Urbano', 'TIPO');
  assert.equal(carlosRow[7], '1800', 'COD');
  assert.equal(carlosRow[8], '17/11/2026', 'DATA convertida do serial Excel');
  assert.equal(carlosRow[9], '09:11', 'HORA');
  assert.equal(carlosRow[10], 'RUA NILO PECANHA, 288 -  - CS', 'ENDEREÇO');
  assert.equal(carlosRow[11], '15741', 'Leit. Anterior');
  assert.equal(carlosRow[12], '15927', 'Leit. Executada');
  assert.equal(carlosRow[13], '186', 'CONSUMO = 186');
  assert.equal(carlosRow[14], 'Faturada', 'NOSB´IMPEDIMENTO');
  assert.equal(carlosRow[15], '1', 'DIGITAÇÃO (Qtd Dig. da linha 3, não "00")');
  assert.equal(carlosRow[16], '00', 'Apres. = "00" quando bruto = "00"');

  // Conferir regra N00 para valores "", "0", "N00", "994", "1"
  for (let i = 1; i <= 5; i++) {
    const rN00 = leituraRes.preview.rows.find((r) => r[1] === `APG1760442${48 + i}`);
    assert.ok(rN00);
    assert.equal(rN00[16], 'N00', `Apres. diferente de "00" deve virar "N00" (caso ${rawApresSamples[i - 1]})`);
  }

  // Sem planilhas auxiliares, Filtros D2:K2 devem ficar vazios e VLOOKUP retorna fallbacks
  const noAuxWb = new ExcelJS.Workbook();
  await noAuxWb.xlsx.load(new Uint8Array(leituraRes.outputBuffer).buffer);
  const noAuxFiltros = noAuxWb.getWorksheet('Filtros')!;
  assert.equal(noAuxFiltros.getCell('D1').value, 'Consistencia');
  assert.equal(noAuxFiltros.getCell('G1').value, 'Nosb');
  assert.equal(noAuxFiltros.getCell('J1').value, 'Fotos');
  assert.ok(!noAuxFiltros.getCell('D2').value, 'Sem planilha auxiliar, Filtros!D2 vazio');
  assert.ok(!noAuxFiltros.getCell('G2').value, 'Sem planilha auxiliar, Filtros!G2 vazio');
  assert.ok(!noAuxFiltros.getCell('J2').value, 'Sem planilha auxiliar, Filtros!J2 vazio');

  console.log('9. Testando escala Leitura.xlsx com 2.500 medidores (250 páginas)...');
  const scaleWb = new ExcelJS.Workbook();
  const scaleWs = scaleWb.addWorksheet('Leitura');
  const setS = (r: number, c: number, v: string) => {
    const cell = scaleWs.getRow(r).getCell(c);
    cell.value = v;
    cell.numFmt = '@';
  };
  let sr = 1;
  for (let p = 0; p < 250; p++) {
    setS(sr, 1, 'CEMIG Distribuição S.A.');
    setS(sr, 12, `Pág. ${p + 1}/250`);
    setS(sr + 1, 1, 'Seqüência de Leitura Executada');
    setS(sr + 2, 1, `Unidade de Leitura: 091028${String(p % 99).padStart(2, '0')}`);
    setS(sr + 3, 1, `Leiturista: ${p % 2 === 0 ? '5JL' : '4MW'}`);
    setS(sr + 4, 1, 'Unidade');
    setS(sr + 4, 2, 'Medidor');
    setS(sr + 4, 3, 'T. Reg');
    setS(sr + 4, 4, 'OC');
    setS(sr + 4, 6, 'Data');
    setS(sr + 4, 7, 'Hora');
    setS(sr + 4, 10, 'Leit. Ant.');
    setS(sr + 4, 11, 'Leit. Exec');
    setS(sr + 4, 13, 'Apres.');
    setS(sr + 5, 1, 'Consumidora');
    setS(sr + 6, 1, 'Cliente');
    setS(sr + 6, 7, 'Endereço');
    setS(sr + 6, 13, 'Qtd Dig.');
    sr += 7;
    for (let m = 0; m < 10; m++) {
      const idx = p * 10 + m + 1;
      setS(sr, 1, `2.796.${String(idx).padStart(6, '0')}-01`);
      setS(sr, 2, `APG${String(idx).padStart(9, '0')}`);
      setS(sr, 3, '03');
      setS(sr, 4, '1800');
      setS(sr, 6, '17/11/2026');
      setS(sr, 7, '09:15');
      setS(sr, 10, '100');
      setS(sr, 11, '125');
      setS(sr, 13, '00');
      setS(sr + 1, 6, '9:15:10 AM');
      setS(sr + 1, 7, `RUA ESCALA, ${idx}`);
      setS(sr + 2, 1, `CLIENTE NOME ${idx}`);
      setS(sr + 2, 4, `000.000.${String(idx).padStart(3, '0')}-00`);
      setS(sr + 2, 13, '1');
      sr += 3;
    }
  }
  const scaleBuf = Buffer.from(await scaleWb.xlsx.writeBuffer());
  const scaleRes = await processCemigWorkbook(scaleBuf, null);
  assert.equal(scaleRes.stats.finalRows, 2500, '2.500 medidores exatos na tabela final');
  assert.equal(scaleRes.stats.finalCols, 20, 'Exatamente 20 colunas');
  assert.ok(scaleRes.preview.rows.every((r) => r[10] !== ''), 'Nenhum Endereço vazio');
  assert.ok(scaleRes.preview.rows.every((r) => r[15] === '1'), 'Nenhum Qtd Dig. perdido');

  const verifyScaleWb = new ExcelJS.Workbook();
  await verifyScaleWb.xlsx.load(new Uint8Array(scaleRes.outputBuffer).buffer);
  const verifyScaleWs = verifyScaleWb.getWorksheet('ControleLeitura')!;
  assert.equal(verifyScaleWs.actualRowCount - 1, 2500, '.xlsx final deve conter TODOS os 2.500 medidores');
  assert.equal(verifyScaleWs.actualColumnCount, 20, '.xlsx final deve conter exatamente 20 colunas');

  console.log('10. Testando planilhas auxiliares Consistencia, Nosb (sem cabeçalho) e Fotos -> aba Filtros...');
  const makeAuxBuf = async (rows: Array<[string, string]>) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Dados');
    rows.forEach(([a, b], idx) => {
      const rObj = ws.getRow(idx + 1);
      rObj.getCell(1).value = a;
      rObj.getCell(1).numFmt = '@';
      rObj.getCell(2).value = b;
      rObj.getCell(2).numFmt = '@';
    });
    return Buffer.from(await wb.xlsx.writeBuffer());
  };

  const consBuf = await makeAuxBuf([
    ['MEDIDOR', 'Retorno Consistencia'],
    ['APG176044248', 'CONS-PENDENTE'],
    ['APG176044249', 'CONS-OK'],
  ]);
  const nosbBuf = await makeAuxBuf([
    ['APG176044248', 'NOSB-SIMULADA'],
    ['MED_RURAL_0', 'NOSB-RURAL'],
  ]);
  const fotosBuf = await makeAuxBuf([
    ['Medidor', 'Status Foto'],
    ['APG176044248', 'OK'],
    ['APG176044250', 'OK-FOTO'],
  ]);

  const auxRes = await processCemigWorkbook(leituraBuf, null, {
    consistenciaBuffer: consBuf,
    nosbBuffer: nosbBuf,
    fotosBuffer: fotosBuf,
  });

  const carlosAux = auxRes.preview.rows.find((r) => r[1] === 'APG176044248')!;
  assert.equal(carlosAux[17], 'CONS-PENDENTE', 'CONSISTÊNCIA preenchida via Consistencia.xlsx (chave MEDIDOR col B)');
  assert.equal(carlosAux[18], 'NOSB-SIMULADA', 'NOSB´SIMULAÇÃO preenchida via Nosb.xlsx (sem cabeçalho, chave MEDIDOR col B)');
  assert.equal(carlosAux[19], 'OK', 'FOTOS preenchida via Fotos.xlsx (chave MEDIDOR col B)');

  const unmatchedAux = auxRes.preview.rows.find((r) => r[1] === 'MED_RURAL_1')!;
  assert.equal(unmatchedAux[17], 'Faturada', 'Fallback CONSISTÊNCIA');
  assert.equal(unmatchedAux[18], 'Faturada', 'Fallback NOSB´SIMULAÇÃO');
  assert.equal(unmatchedAux[19], 'N-OK', 'Fallback FOTOS');

  const auxOutWb = new ExcelJS.Workbook();
  await auxOutWb.xlsx.load(new Uint8Array(auxRes.outputBuffer).buffer);
  const auxMainWs = auxOutWb.getWorksheet('ControleLeitura')!;
  assert.equal(
    (auxMainWs.getCell('R2').value as { formula?: string }).formula,
    'IFERROR(VLOOKUP(B2,Filtros!$D$1:$E$50000,2,0),"Faturada")',
  );
  const auxFiltros = auxOutWb.getWorksheet('Filtros')!;
  assert.equal(auxFiltros.getCell('D1').value, 'Consistencia');
  assert.equal(auxFiltros.getCell('E1').value, 'Retorno Consistencia');
  assert.equal(auxFiltros.getCell('D2').value, 'APG176044248');
  assert.equal(auxFiltros.getCell('E2').value, 'CONS-PENDENTE');
  assert.equal(auxFiltros.getCell('D3').value, 'APG176044249');
  assert.equal(auxFiltros.getCell('G1').value, 'Nosb');
  assert.equal(auxFiltros.getCell('H1').value, 'Valor');
  assert.equal(auxFiltros.getCell('G2').value, 'APG176044248');
  assert.equal(auxFiltros.getCell('H2').value, 'NOSB-SIMULADA');
  assert.equal(auxFiltros.getCell('G3').value, 'MED_RURAL_0');
  assert.equal(auxFiltros.getCell('J1').value, 'Fotos');
  assert.equal(auxFiltros.getCell('K1').value, 'Status Foto');
  assert.equal(auxFiltros.getCell('J2').value, 'APG176044248');
  assert.equal(auxFiltros.getCell('K2').value, 'OK');

  const d2Fill = (auxFiltros.getCell('D2').fill as { fgColor?: { argb?: string } } | undefined)?.fgColor?.argb;
  const e2Fill = (auxFiltros.getCell('E2').fill as { fgColor?: { argb?: string } } | undefined)?.fgColor?.argb;
  const g2Fill = (auxFiltros.getCell('G2').fill as { fgColor?: { argb?: string } } | undefined)?.fgColor?.argb;
  const j2Fill = (auxFiltros.getCell('J2').fill as { fgColor?: { argb?: string } } | undefined)?.fgColor?.argb;
  assert.equal(d2Fill, 'FFDCE6F1');
  assert.equal(e2Fill, 'FFDCE6F1');
  assert.equal(g2Fill, 'FFE2EFDA');
  assert.equal(j2Fill, 'FFFFF2CC');

  console.log('11. Testando Controle Horário (agrupamento Matricula+Data, cálculo HH:MM, 1 leitura = 00:00, aba "Controle Horario")...');
  const chWb = new ExcelJS.Workbook();
  const chWs = chWb.addWorksheet('Leitura');
  const setC = (r: number, c: number, v: string | number) => {
    const cell = chWs.getRow(r).getCell(c);
    cell.value = v;
    if (typeof v === 'string') cell.numFmt = '@';
  };
  let chR = 1;
  setC(chR, 1, 'CEMIG Distribuição S.A.');
  setC(chR, 12, 'Pág. 1/1');
  setC(chR + 1, 1, 'Seqüência de Leitura Executada');
  setC(chR + 2, 1, 'Unidade de Leitura: 09102801');
  setC(chR + 3, 1, 'Leiturista: 5JL');
  setC(chR + 4, 1, 'Unidade');
  setC(chR + 4, 2, 'Medidor');
  setC(chR + 4, 3, 'T. Reg');
  setC(chR + 4, 4, 'OC');
  setC(chR + 4, 6, 'Data');
  setC(chR + 4, 7, 'Hora');
  setC(chR + 4, 10, 'Leit. Ant.');
  setC(chR + 4, 11, 'Leit. Exec');
  setC(chR + 4, 13, 'Apres.');
  setC(chR + 5, 1, 'Consumidora');
  setC(chR + 6, 1, 'Cliente');
  setC(chR + 6, 7, 'Endereço');
  setC(chR + 6, 13, 'Qtd Dig.');
  chR += 7;

  // Leiturista 5JL no dia 17/11/2026: 09:11 até 16:20 -> Tempo de Serviço = 07:09
  const addMeter = (inst: string, med: string, dt: string, hr: string) => {
    setC(chR, 1, inst);
    setC(chR, 2, med);
    setC(chR, 3, '03');
    setC(chR, 4, '1800');
    setC(chR, 6, dt);
    setC(chR, 7, hr);
    setC(chR, 10, '100');
    setC(chR, 11, '120');
    setC(chR, 13, '00');
    setC(chR + 1, 7, `RUA TESTE, 1`);
    setC(chR + 2, 1, `NOME ${med}`);
    setC(chR + 2, 13, '1');
    chR += 3;
  };
  addMeter('2.100.000.001-01', 'MED_CH_1', '17/11/2026', '09:11');
  addMeter('2.100.000.002-01', 'MED_CH_2', '17/11/2026', '16:20');
  // Leiturista 5JL no dia 18/11/2026: apenas 1 leitura (11:05) -> Tempo de Serviço = 00:00
  addMeter('2.100.000.003-01', 'MED_CH_3', '18/11/2026', '11:05');

  // Outro leiturista 4MW no dia 17/11/2026: 08:00 até 10:30 -> 02:30 (deve vir antes de 5JL na ordenação)
  setC(chR, 1, 'Unidade de Leitura: 09102802');
  setC(chR + 1, 1, 'Leiturista: 4MW');
  chR += 2;
  addMeter('2.100.000.004-01', 'MED_CH_4', '17/11/2026', '10:30');
  addMeter('2.100.000.005-01', 'MED_CH_5', '17/11/2026', '08:00');

  const chBuf = Buffer.from(await chWb.xlsx.writeBuffer());
  const chRes = await processCemigWorkbook(chBuf, null);

  assert.deepEqual(chRes.controleHorario, [
    {
      matricula: '4MW',
      data: '17/11/2026',
      horaInicial: '08:00',
      horaFinal: '10:30',
      tempoServico: '02:30',
    },
    {
      matricula: '5JL',
      data: '17/11/2026',
      horaInicial: '09:11',
      horaFinal: '16:20',
      tempoServico: '07:09',
    },
    {
      matricula: '5JL',
      data: '18/11/2026',
      horaInicial: '11:05',
      horaFinal: '11:05',
      tempoServico: '00:00',
    },
  ]);

  const chOutWb = new ExcelJS.Workbook();
  await chOutWb.xlsx.load(new Uint8Array(chRes.outputBuffer).buffer);
  const wsControleHorario = chOutWb.getWorksheet('Controle Horario');
  assert.ok(wsControleHorario, 'Aba "Controle Horario" deve existir no .xlsx exportado');
  const chSheetHeaders = [1, 2, 3, 4, 5].map((c) =>
    String(wsControleHorario.getRow(1).getCell(c).value ?? ''),
  );
  assert.deepEqual(chSheetHeaders, [...CONTROLE_HORARIO_HEADERS]);
  assert.equal(wsControleHorario.actualRowCount, 4, '1 cabeçalho + 3 linhas agrupadas');
  assert.equal(wsControleHorario.getCell('A3').value, '5JL');
  assert.equal(wsControleHorario.getCell('B3').value, '17/11/2026');
  assert.equal(wsControleHorario.getCell('C3').value, '09:11');
  assert.equal(wsControleHorario.getCell('D3').value, '16:20');
  assert.equal(wsControleHorario.getCell('E3').value, '07:09');
  assert.equal(wsControleHorario.getCell('E4').value, '00:00');

  console.log('12. Testando card "Quantidade Instalações" (contagem total de linhas não vazias, incluindo repetições)...');
  const dupInstWb = new ExcelJS.Workbook();
  const dupWs = dupInstWb.addWorksheet('Leitura');
  const setD = (r: number, c: number, v: string) => {
    const cell = dupWs.getRow(r).getCell(c);
    cell.value = v;
    cell.numFmt = '@';
  };
  setD(1, 1, 'CEMIG Distribuição S.A.');
  setD(1, 12, 'Pág. 1/1');
  setD(2, 1, 'Seqüência de Leitura Executada');
  setD(3, 1, 'Unidade de Leitura: 09102801');
  setD(4, 1, 'Leiturista: 5JL');
  setD(5, 1, 'Unidade'); setD(5, 2, 'Medidor'); setD(5, 3, 'T. Reg'); setD(5, 4, 'OC');
  setD(5, 6, 'Data'); setD(5, 7, 'Hora'); setD(5, 10, 'Leit. Ant.'); setD(5, 11, 'Leit. Exec'); setD(5, 13, 'Apres.');
  setD(6, 1, 'Consumidora'); setD(7, 1, 'Cliente'); setD(7, 7, 'Endereço'); setD(7, 13, 'Qtd Dig.');
  // 2 linhas com mesma INSTALAÇÃO + 1 linha com outra INSTALAÇÃO + 1 linha com INSTALAÇÃO vazia -> total = 3 (de 4 medidores)
  const addDup = (rBase: number, inst: string, med: string) => {
    setD(rBase, 1, inst); setD(rBase, 2, med); setD(rBase, 3, '03'); setD(rBase, 4, '1800');
    setD(rBase, 6, '17/11/2026'); setD(rBase, 7, '09:00'); setD(rBase, 10, '10'); setD(rBase, 11, '20'); setD(rBase, 13, '00');
    setD(rBase + 1, 7, 'RUA X'); setD(rBase + 2, 1, 'CLI'); setD(rBase + 2, 13, '1');
  };
  addDup(8, '2.796.450.018-02', 'M1');
  addDup(11, '2.796.450.018-02', 'M2');
  addDup(14, '2.796.450.019-02', 'M3');
  addDup(17, '', 'M4');
  const dupBuf = Buffer.from(await dupInstWb.xlsx.writeBuffer());
  const dupRes = await processCemigWorkbook(dupBuf, null);
  assert.equal(dupRes.stats.finalRows, 4, '4 medidores na tabela final');
  assert.equal(dupRes.stats.uniqueInstalacoesCount, 3, 'Card Quantidade Instalações deve contar 3 linhas com INSTALAÇÃO preenchida (incluindo repetida e ignorando vazia)');

  console.log('13. Testando upload de Planilha Modelo RGeral (preenchimento a partir de A3 e omissão quando não enviada)...');
  // Sem modeloRGeralBuffer: NÃO deve gerar Relatorio Geral
  assert.equal(auxRes.relatorioGeralFileName, null, 'Sem modeloRGeral, relatorioGeralFileName deve ser null');
  assert.equal(auxRes.relatorioGeralBuffer, null, 'Sem modeloRGeral, relatorioGeralBuffer deve ser null');

  // Conferir que o primeiro arquivo mantém os cabeçalhos originais
  const mainHeadersActual = Array.from({ length: 20 }, (_, idx) =>
    String(auxMainWs.getRow(1).getCell(idx + 1).value ?? ''),
  );
  assert.deepEqual(mainHeadersActual, [...FINAL_HEADERS], 'Matriz_ControleDeSequenciaLeitura.xlsx deve manter os cabeçalhos originais');

  // Criar planilha modelo com aba extra "Resumo", aba "RGeral" com cabeçalhos na linha 1, subtítulo na linha 2 e largura/cores
  const tplWb = new ExcelJS.Workbook();
  const wsResumo = tplWb.addWorksheet('Resumo');
  wsResumo.getCell('A1').value = 'Total de medidores';
  wsResumo.getCell('B1').value = { formula: 'COUNTA(RGeral!B3:B5000)', result: 0 };

  const wsTplRGeral = tplWb.addWorksheet('RGeral');
  RELATORIO_GERAL_HEADERS.forEach((hdr, idx) => {
    const cell = wsTplRGeral.getRow(1).getCell(idx + 1);
    cell.value = hdr;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    wsTplRGeral.getColumn(idx + 1).width = 22;
  });
  wsTplRGeral.getRow(2).getCell(1).value = 'SUBTITULO PRESERVADO';
  wsTplRGeral.getRow(2).getCell(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFEF08A' },
  };
  wsTplRGeral.getRow(3).getCell(21).value = { formula: 'O3*2', result: 0 }; // Coluna U preservada

  const tplBuffer = Buffer.from(await tplWb.xlsx.writeBuffer());
  const auxWithTplRes = await processCemigWorkbook(leituraBuf, null, {
    consistenciaBuffer: consBuf,
    nosbBuffer: nosbBuf,
    fotosBuffer: fotosBuf,
    modeloRGeralBuffer: tplBuffer,
  });

  assert.equal(auxWithTplRes.relatorioGeralFileName, 'Relatorio Geral.xlsx');
  assert.ok(auxWithTplRes.relatorioGeralBuffer && auxWithTplRes.relatorioGeralBuffer.length > 0);

  const filledWb = new ExcelJS.Workbook();
  await filledWb.xlsx.load(new Uint8Array(auxWithTplRes.relatorioGeralBuffer).buffer);
  assert.deepEqual(
    filledWb.worksheets.map((w) => w.name),
    ['Resumo', 'RGeral'],
    'Outras abas do modelo (Resumo) devem ser preservadas',
  );
  assert.equal(
    (filledWb.getWorksheet('Resumo')!.getCell('B1').value as { formula?: string })?.formula,
    'COUNTA(RGeral!B3:B5000)',
    'Fórmulas em outras abas do modelo devem ser preservadas',
  );

  const wsRGeral = filledWb.getWorksheet('RGeral')!;
  const rgHeadersActual = Array.from({ length: 20 }, (_, idx) =>
    String(wsRGeral.getRow(1).getCell(idx + 1).value ?? ''),
  );
  assert.deepEqual(rgHeadersActual, [...RELATORIO_GERAL_HEADERS]);
  assert.equal(wsRGeral.getRow(1).getCell(1).fill?.type, 'pattern', 'Formatação da linha 1 preservada');
  assert.equal(wsRGeral.getColumn(1).width, 22, 'Largura de coluna do modelo preservada');
  assert.equal(wsRGeral.getRow(2).getCell(1).value, 'SUBTITULO PRESERVADO', 'Linha 2 NÃO deve ser sobrescrita');
  assert.equal((wsRGeral.getRow(3).getCell(21).value as { formula?: string })?.formula, 'O3*2', 'Colunas além de T preservadas');
  assert.equal(wsRGeral.actualRowCount - 2, auxWithTplRes.stats.finalRows, 'Dados colados a partir da linha 3 (A3)');

  let carlosRgRowIdx = -1;
  for (let r = 3; r <= wsRGeral.actualRowCount; r++) {
    if (wsRGeral.getRow(r).getCell(2).value === 'APG176044248') {
      carlosRgRowIdx = r;
      break;
    }
  }
  assert.ok(carlosRgRowIdx >= 3, 'Linha APG176044248 deve existir em RGeral a partir de A3');
  const carlosRgRow = wsRGeral.getRow(carlosRgRowIdx);
  assert.equal(carlosRgRow.getCell(1).value, '2.796.450.018-02', 'INSTALAÇÃO');
  assert.equal(carlosRgRow.getCell(3).value, '03', 'REG.');
  assert.equal(carlosRgRow.getCell(6).value, 'OK', 'Foto? resolvida do Filtros');
  assert.equal(carlosRgRow.getCell(7).value, 'Urbano', 'Tipo');
  assert.equal(carlosRgRow.getCell(8).value, '5JL', 'MATR');
  assert.equal(carlosRgRow.getCell(9).value, '1800', 'NL');
  assert.equal(carlosRgRow.getCell(10).value, '17/11/2026', 'DATA');
  assert.equal(carlosRgRow.getCell(15).value, 186, 'CONSUMO = 186');
  assert.equal(carlosRgRow.getCell(17).value, 'NOSB-SIMULADA', 'SIMULAÇÃO resolvida do Filtros');
  assert.equal(carlosRgRow.getCell(18).value, '1', 'DIG');
  assert.equal(carlosRgRow.getCell(19).value, '00', 'CNA = 00');
  assert.equal(carlosRgRow.getCell(20).value, 'CONS-PENDENTE', 'CONSISTENCIA resolvida do Filtros');

  let n00RgRowIdx = -1;
  for (let r = 3; r <= wsRGeral.actualRowCount; r++) {
    if (wsRGeral.getRow(r).getCell(2).value === 'APG176044249') {
      n00RgRowIdx = r;
      break;
    }
  }
  assert.ok(n00RgRowIdx >= 3);
  assert.equal(wsRGeral.getRow(n00RgRowIdx).getCell(19).value, 'N00', 'CNA = N00 quando Apres. != "00"');

  // Testar criação automática da aba RGeral caso a planilha enviada não possua a aba RGeral
  const noSheetWb = new ExcelJS.Workbook();
  noSheetWb.addWorksheet('Capa').getCell('A1').value = 'Planilha Modelo Sem RGeral';
  const noSheetRes = await processCemigWorkbook(leituraBuf, null, {
    modeloRGeralBuffer: Buffer.from(await noSheetWb.xlsx.writeBuffer()),
  });
  assert.ok(noSheetRes.relatorioGeralBuffer);
  const autoSheetWb = new ExcelJS.Workbook();
  await autoSheetWb.xlsx.load(new Uint8Array(noSheetRes.relatorioGeralBuffer).buffer);
  const autoWsRG = autoSheetWb.getWorksheet('RGeral');
  assert.ok(autoWsRG, 'Aba RGeral deve ser criada automaticamente quando ausente');
  assert.deepEqual(
    Array.from({ length: 20 }, (_, i) => String(autoWsRG.getRow(1).getCell(i + 1).value ?? '')),
    [...RELATORIO_GERAL_HEADERS],
  );
  assert.ok(String(autoWsRG.getRow(3).getCell(2).value ?? '').trim() !== '', 'Dados iniciam em A3 mesmo quando aba é criada');

  console.log('14. Testando Planilha Modelo .XLSM e .XLSX multi-abas (Resumo, Dashboard, PivotTable, Gráfico, fullCalcOnLoad, valores estáticos em A3)...');
  // Construir um .xlsm multi-abas a partir de um .xlsx base adicionando xl/vbaProject.bin, PivotTable, PivotCache e Gráfico
  const macroBaseWb = new ExcelJS.Workbook();
  const wsMacroResumo = macroBaseWb.addWorksheet('Resumo');
  wsMacroResumo.getCell('A1').value = { formula: 'COUNTA(RGeral!B3:B5000)', result: 0 };
  wsMacroResumo.getCell('B1').value = { formula: 'SUM(RGeral!O3:O5000)', result: 0 };

  const wsMacroDashboard = macroBaseWb.addWorksheet('Dashboard');
  wsMacroDashboard.getCell('A1').value = 'Indicadores Gerais';
  wsMacroDashboard.getCell('B1').value = { formula: 'COUNTIF(RGeral!G:G,"Urbano")', result: 0 };
  wsMacroDashboard.getCell('B2').value = { formula: 'COUNTIF(RGeral!S:S,"00")', result: 0 };

  const wsMacroRG = macroBaseWb.addWorksheet('RGeral');
  RELATORIO_GERAL_HEADERS.forEach((h, i) => {
    const c = wsMacroRG.getRow(1).getCell(i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
  });
  wsMacroRG.getRow(2).getCell(1).value = 'SUBTITULO MACRO PRESERVADO';
  // Preencher 25 linhas antigas (A3..T27) para garantir sobrescrita e limpeza das linhas excedentes (15 < 25)
  for (let r = 3; r <= 27; r++) {
    wsMacroRG.getRow(r).getCell(1).value = `ANTIGO_${r}`;
    wsMacroRG.getRow(r).getCell(2).value = `MED_ANTIGO_${r}`;
  }
  const macroBaseBuf = Buffer.from(await macroBaseWb.xlsx.writeBuffer());
  const macroZip = await JSZip.loadAsync(macroBaseBuf);

  const fakeVbaPayload = Buffer.from('FAKE_VBA_PROJECT_BIN_MACROS_DATA_1234567890_PRESERVED', 'utf8');
  const fakePivotXml = Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition name="TabelaDinamica1" cacheId="1"/>', 'utf8');
  const fakePivotCacheXml = Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition r:id="rId1"><cacheSource type="worksheet"><worksheetSource ref="A1:T500" sheet="RGeral"/></cacheSource></pivotCacheDefinition>', 'utf8');
  const fakeChartXml = Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>', 'utf8');

  macroZip.file('xl/vbaProject.bin', fakeVbaPayload);
  macroZip.file('xl/pivotTables/pivotTable1.xml', fakePivotXml);
  macroZip.file('xl/pivotCache/pivotCacheDefinition1.xml', fakePivotCacheXml);
  macroZip.file('xl/charts/chart1.xml', fakeChartXml);

  // Adicionar PivotTable/Chart ao .xlsx multi-abas também (sem vbaProject.bin) para garantir preservação em .xlsx
  const xlsxMultiZip = await JSZip.loadAsync(macroBaseBuf);
  xlsxMultiZip.file('xl/pivotTables/pivotTable1.xml', fakePivotXml);
  xlsxMultiZip.file('xl/pivotCache/pivotCacheDefinition1.xml', fakePivotCacheXml);
  xlsxMultiZip.file('xl/charts/chart1.xml', fakeChartXml);
  const xlsxMultiBuffer = await xlsxMultiZip.generateAsync({ type: 'nodebuffer' });

  let wbRels = await macroZip.file('xl/_rels/workbook.xml.rels')!.async('string');
  wbRels = wbRels.replace(
    '</Relationships>',
    '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
  );
  macroZip.file('xl/_rels/workbook.xml.rels', wbRels);

  let ctXml = await macroZip.file('[Content_Types].xml')!.async('string');
  ctXml = ctXml
    .replace(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    )
    .replace(
      '</Types>',
      '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
    );
  macroZip.file('[Content_Types].xml', ctXml);

  const xlsmInputBuffer = await macroZip.generateAsync({ type: 'nodebuffer' });

  const xlsmRes = await processCemigWorkbook(leituraBuf, null, {
    consistenciaBuffer: consBuf,
    nosbBuffer: nosbBuf,
    fotosBuffer: fotosBuf,
    modeloRGeralBuffer: xlsmInputBuffer,
    modeloRGeralFileName: 'Modelo_Relatorio.xlsm',
  });

  assert.equal(xlsmRes.relatorioGeralFileName, 'Relatorio Geral.xlsm', 'Deve preservar extensão .xlsm');
  assert.ok(xlsmRes.relatorioGeralBuffer && xlsmRes.relatorioGeralBuffer.length > 0);

  const outXlsmZip = await JSZip.loadAsync(xlsmRes.relatorioGeralBuffer);
  const vbaOutEntry = outXlsmZip.file('xl/vbaProject.bin');
  assert.ok(vbaOutEntry, 'xl/vbaProject.bin DEVE ser preservado no .xlsm gerado');
  const vbaOutBytes = await vbaOutEntry.async('nodebuffer');
  assert.deepEqual(vbaOutBytes, fakeVbaPayload, 'Bytes de xl/vbaProject.bin devem ser idênticos aos do modelo');

  assert.deepEqual(
    await outXlsmZip.file('xl/pivotTables/pivotTable1.xml')!.async('nodebuffer'),
    fakePivotXml,
    'PivotTable XML deve ser preservada byte a byte no .xlsm',
  );
  assert.deepEqual(
    await outXlsmZip.file('xl/charts/chart1.xml')!.async('nodebuffer'),
    fakeChartXml,
    'Gráfico XML deve ser preservado byte a byte no .xlsm',
  );
  const outPivotCacheXml = await outXlsmZip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string');
  assert.ok(outPivotCacheXml.includes('refreshOnLoad="1"'), 'PivotCache deve ter refreshOnLoad="1"');

  const outXlsmWbXml = await outXlsmZip.file('xl/workbook.xml')!.async('string');
  assert.ok(outXlsmWbXml.includes('fullCalcOnLoad="1"'), 'xl/workbook.xml deve ter fullCalcOnLoad="1"');

  const outCtXml = await outXlsmZip.file('[Content_Types].xml')!.async('string');
  assert.ok(
    outCtXml.includes('application/vnd.ms-excel.sheet.macroEnabled.main+xml'),
    'Content-Type macroEnabled deve ser preservado',
  );

  // Reabrir com SheetJS (bookVBA: true) e conferir abas, fórmulas dependentes e valores estáticos a partir de A3
  const xlsmParsed = XLSX.read(xlsmRes.relatorioGeralBuffer, { type: 'buffer', bookVBA: true });
  assert.deepEqual(xlsmParsed.SheetNames, ['Resumo', 'Dashboard', 'RGeral'], 'Todas as abas originais devem permanecer na ordem exata');
  assert.equal(xlsmParsed.Sheets['Resumo']['A1']?.f, 'COUNTA(RGeral!B3:B5000)');
  assert.equal(xlsmParsed.Sheets['Resumo']['B1']?.f, 'SUM(RGeral!O3:O5000)');
  assert.equal(xlsmParsed.Sheets['Dashboard']['B1']?.f, 'COUNTIF(RGeral!G:G,"Urbano")');
  assert.equal(xlsmParsed.Sheets['Dashboard']['B2']?.f, 'COUNTIF(RGeral!S:S,"00")');

  const xlsmRgWs = xlsmParsed.Sheets['RGeral'];
  const xlsmGrid = XLSX.utils.sheet_to_json<string[]>(xlsmRgWs, { header: 1, defval: '' });
  assert.deepEqual(xlsmGrid[0].slice(0, 20), [...RELATORIO_GERAL_HEADERS]);
  assert.equal(xlsmGrid[1][0], 'SUBTITULO MACRO PRESERVADO', 'Linha 2 preservada no .xlsm');
  // 2 linhas de cabeçalho/subtítulo + 15 registros tratados (linhas 18..27 antigas devem ter sido limpas)
  assert.equal(xlsmGrid.length, 2 + xlsmRes.stats.finalRows, 'Linhas excedentes antigas devem ser removidas');
  const carlosXlsmRow = xlsmGrid.find((row) => row[1] === 'APG176044248')!;
  const carlosXlsmRowNum = xlsmGrid.findIndex((row) => row[1] === 'APG176044248') + 1;
  assert.ok(carlosXlsmRow && carlosXlsmRowNum >= 3, 'APG176044248 encontrado no .xlsm');
  assert.equal(carlosXlsmRow[5], 'OK', 'Foto? no .xlsm');
  assert.equal(carlosXlsmRow[16], 'NOSB-SIMULADA', 'SIMULAÇÃO no .xlsm');
  assert.equal(carlosXlsmRow[18], '00', 'CNA = 00 no .xlsm');
  assert.equal(carlosXlsmRow[19], 'CONS-PENDENTE', 'CONSISTENCIA no .xlsm');

  // Conferir que células em RGeral são valores estáticos (sem .f)
  for (const colLetter of ['F', 'J', 'K', 'O', 'S', 'T']) {
    const cellObj = xlsmRgWs[`${colLetter}${carlosXlsmRowNum}`];
    assert.ok(cellObj && cellObj.v !== undefined, `${colLetter}${carlosXlsmRowNum} deve possuir valor`);
    assert.equal(cellObj.f, undefined, `${colLetter}${carlosXlsmRowNum} NÃO deve conter fórmula`);
  }

  // Testar o mesmo modelo multi-abas com PivotTable e Gráfico em .xlsx
  const xlsxMultiRes = await processCemigWorkbook(leituraBuf, null, {
    consistenciaBuffer: consBuf,
    nosbBuffer: nosbBuf,
    fotosBuffer: fotosBuf,
    modeloRGeralBuffer: xlsxMultiBuffer,
    modeloRGeralFileName: 'Modelo_Dashboard.xlsx',
  });
  assert.equal(xlsxMultiRes.relatorioGeralFileName, 'Relatorio Geral.xlsx');
  assert.ok(xlsxMultiRes.relatorioGeralBuffer);
  const outXlsxZip = await JSZip.loadAsync(xlsxMultiRes.relatorioGeralBuffer);
  assert.deepEqual(
    await outXlsxZip.file('xl/pivotTables/pivotTable1.xml')!.async('nodebuffer'),
    fakePivotXml,
    'PivotTable XML deve ser preservada byte a byte no .xlsx',
  );
  assert.deepEqual(
    await outXlsxZip.file('xl/charts/chart1.xml')!.async('nodebuffer'),
    fakeChartXml,
    'Gráfico XML deve ser preservado byte a byte no .xlsx',
  );
  assert.ok((await outXlsxZip.file('xl/workbook.xml')!.async('string')).includes('fullCalcOnLoad="1"'));
  const xlsxMultiParsed = XLSX.read(xlsxMultiRes.relatorioGeralBuffer, { type: 'buffer' });
  assert.deepEqual(xlsxMultiParsed.SheetNames, ['Resumo', 'Dashboard', 'RGeral']);
  assert.equal(xlsxMultiParsed.Sheets['Dashboard']['B1']?.f, 'COUNTIF(RGeral!G:G,"Urbano")');
  assert.equal(xlsxMultiParsed.Sheets['RGeral'][`O${carlosXlsmRowNum}`]?.f, undefined);
  assert.equal(xlsxMultiParsed.Sheets['RGeral'][`S${carlosXlsmRowNum}`]?.f, undefined);
  assert.equal(xlsxMultiParsed.Sheets['RGeral'][`T${carlosXlsmRowNum}`]?.f, undefined);

  console.log('✅ Todos os testes (Planilha Modelo .xlsm/.xlsx multi-abas + PivotTables + Gráficos + fórmulas dependentes + valores estáticos em A3) passaram!');
}

runTests().catch((err) => {
  console.error('❌ Falha nos testes:', err);
  process.exit(1);
});
