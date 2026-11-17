import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {
  CemigValidationError,
  FINAL_HEADERS,
  processCemigWorkbook,
} from '../server/cemigProcessor.ts';
import { generateSampleCemigMatrizBuffer } from '../server/sampleCemigGenerator.ts';

async function runTests() {
  console.log('1. Gerando planilha Matriz de exemplo (layout CEMIG)...');
  const sampleBuf = await generateSampleCemigMatrizBuffer();
  assert.ok(sampleBuf.length > 0, 'Sample buffer deve ser não-vazio');

  console.log('2. Executando pipeline completo (17 colunas, dd/mm/aaaa, ordenação)...');
  const res = await processCemigWorkbook(sampleBuf, null);

  assert.equal(res.fileName, 'Matriz_ControleDeSequenciaLeitura.xlsx');
  assert.equal(res.stats.finalCols, 17, 'Deve conter exatamente 17 colunas');
  assert.deepEqual(res.preview.headers, [...FINAL_HEADERS], 'Cabeçalhos devem seguir exatamente a ordem das 17 colunas');
  // 10 registros no sample: 1 duplicado removido, 1 excluído pelo filtro 99Z -> 9 linhas finais
  assert.equal(res.stats.finalRows, 9, 'Deve conter 9 linhas válidas após deduplicação e filtro');

  // Verificar ordenação por Leiturista (4MW) e DATA crescente: 06/11/2026 -> 07/11/2026 -> 12/11/2026
  const rows4MW = res.preview.rows.filter((r) => r[4] === '4MW');
  assert.equal(rows4MW.length, 3, 'Devem existir 3 registros do leiturista 4MW');
  assert.deepEqual(
    rows4MW.map((r) => r[6]),
    ['06/11/2026', '07/11/2026', '12/11/2026'],
    'Datas dentro de 4MW devem estar em dd/mm/aaaa e ordem crescente',
  );

  // Verificar agrupamento contíguo por Leiturista
  const seenGroups = new Set<string>();
  let prevGroup = '';
  for (const row of res.preview.rows) {
    const grp = row[4];
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
  assert.equal(mariaRow[4], '4MW', 'MATRÍCULA');
  assert.equal(mariaRow[5], '00', 'COD');
  assert.equal(mariaRow[6], '06/11/2026', 'DATA dd/mm/aaaa');
  assert.equal(mariaRow[7], '10:03:55 AM', 'HORA preservada');
  assert.equal(mariaRow[8], 'RUA DAS FLORES, 142 - CENTRO', 'ENDEREÇO');
  assert.equal(mariaRow[9], '23873', 'Leit. Anterior');
  assert.equal(mariaRow[10], '24182', 'Leit. Executada');
  assert.equal(mariaRow[11], '309', 'CONSUMO = 309');
  assert.equal(mariaRow[12], 'Faturada', 'NOSB´IMPEDIMENTO');
  assert.equal(mariaRow[13], 'N00', 'DIGITAÇÃO');
  assert.equal(mariaRow[14], 'Faturada', 'CONSISTÊNCIA');
  assert.equal(mariaRow[15], 'Faturada', 'NOSB´SIMULAÇÃO');
  assert.equal(mariaRow[16], 'N-OK', 'FOTOS');

  // Verificar que mensagens técnicas proibidas NÃO aparecem
  for (const w of res.warnings) {
    assert.ok(!w.includes('offset'), 'Não deve emitir mensagem técnica de offset');
    assert.ok(!w.includes('colunas D:K vazias'), 'Não deve emitir mensagem técnica de Planilha3');
  }

  console.log('3. Validando estrutura física do .xlsx gerado...');
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.load(new Uint8Array(res.outputBuffer).buffer);

  const wsMain = outWb.getWorksheet('ControleLeitura');
  const wsP3 = outWb.getWorksheet('Planilha3');
  assert.ok(wsMain, 'Aba ControleLeitura deve existir');
  assert.ok(wsP3, 'Aba Planilha3 deve existir');

  const table = wsMain.getTable('Tabela1');
  assert.ok(table, 'ListObject "Tabela1" deve existir na aba principal');
  assert.ok(!wsMain.hasMerges, 'Não deve haver células mescladas na saída');

  const e2 = wsMain.getCell('E2').value as { formula?: string; result?: string };
  const l2 = wsMain.getCell('L2').value as { formula?: string; result?: number };
  const o2 = wsMain.getCell('O2').value as { formula?: string; result?: string };
  const p2 = wsMain.getCell('P2').value as { formula?: string; result?: string };
  const q2 = wsMain.getCell('Q2').value as { formula?: string; result?: string };

  assert.equal(e2.formula, 'VLOOKUP(D2,Planilha3!$A$1:$B$50000,2,0)');
  assert.equal(e2.result, '4MW');
  assert.equal(l2.formula, 'SUM(K2-J2)');
  assert.equal(l2.result, 309);
  assert.equal(o2.formula, 'IFERROR(VLOOKUP(A2,Planilha3!$D$1:$E$50000,2,0),"Faturada")');
  assert.equal(p2.formula, 'IFERROR(VLOOKUP(A2,Planilha3!$G$1:$H$50000,2,0),"Faturada")');
  assert.equal(q2.formula, 'IFERROR(VLOOKUP(A2,Planilha3!$J$1:$K$50000,2,0),"N-OK")');

  const c2 = wsMain.getCell('C2');
  assert.equal(c2.value, '00060', 'Célula C2 mantém texto "00060"');
  assert.equal(c2.alignment?.horizontal, 'left');
  assert.equal(c2.alignment?.vertical, 'bottom');
  assert.ok(!c2.alignment?.wrapText, 'wrapText não deve estar ativado');

  console.log('4. Testando validação contra planilha inválida...');
  const badWb = new ExcelJS.Workbook();
  const badWs = badWb.addWorksheet('Plan1');
  badWs.getCell('A1').value = 'Produto';
  badWs.getCell('B1').value = 'Preço';
  const badBuf = Buffer.from(await badWb.xlsx.writeBuffer());

  await assert.rejects(
    async () => processCemigWorkbook(badBuf, null),
    CemigValidationError,
  );

  console.log('✅ Todos os testes das 17 colunas, datas dd/mm/aaaa e ordenação passaram!');
}

runTests().catch((err) => {
  console.error('❌ Falha nos testes:', err);
  process.exit(1);
});
