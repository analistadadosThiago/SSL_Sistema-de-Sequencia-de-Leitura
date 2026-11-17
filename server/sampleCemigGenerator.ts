import ExcelJS from 'exceljs';

interface SampleMeterRecord {
  consumidora: string;
  medidor: string;
  tReg: string;
  sequencia: string;
  data: string;
  hora: string;
  unidadeLeitura: string;
  leiturista: string;
  leitAnt: string;
  leitExec: string;
  imp: string;
  qtdDig: string;
  apres: string;
}

const SAMPLE_RECORDS: SampleMeterRecord[] = [
  {
    consumidora: 'MARIA CELESTE M MORAIS',
    medidor: 'ARL211115451',
    tReg: '00060',
    sequencia: '00136',
    data: '11/6/2026',
    hora: '10:03:55 AM',
    unidadeLeitura: '31020001',
    leiturista: '4MW',
    leitAnt: '23873',
    leitExec: '24182',
    imp: '',
    qtdDig: '',
    apres: 'RUA DAS FLORES, 142 - CENTRO',
  },
  {
    consumidora: 'JOSE CARLOS FERREIRA',
    medidor: 'CEM994812034',
    tReg: '00060',
    sequencia: '00137',
    data: '11/12/2026',
    hora: '10:06:12 AM',
    unidadeLeitura: '31020001',
    leiturista: '4MW',
    leitAnt: '14210',
    leitExec: '14495',
    imp: 'Lida',
    qtdDig: '05',
    apres: 'RUA DAS FLORES, 158 - CENTRO',
  },
  {
    consumidora: 'ROBERTO CARLOS MENDES',
    medidor: 'ARL211115999',
    tReg: '00060',
    sequencia: '00145',
    data: '11/7/2026',
    hora: '09:15:10 AM',
    unidadeLeitura: '31020001',
    leiturista: '4MW',
    leitAnt: '10200',
    leitExec: '10350',
    imp: '',
    qtdDig: '05',
    apres: 'RUA DAS FLORES, 200 - CENTRO',
  },
  {
    consumidora: 'ANA PAULA RIBEIRO',
    medidor: 'ELG772104882',
    tReg: '00120',
    sequencia: '00138',
    data: '11/6/2026',
    hora: '10:09:44 AM',
    unidadeLeitura: '31020002',
    leiturista: '4NI',
    leitAnt: '08450',
    leitExec: '08712',
    imp: '',
    qtdDig: '05',
    apres: 'AV AFONSO PENA, 890 - FUNCIONARIOS',
  },
  {
    consumidora: 'SUPERMERCADO BOM PRECO LTDA',
    medidor: 'LAN550291837',
    tReg: '00240',
    sequencia: '00139',
    data: '11/6/2026',
    hora: '10:14:21 AM',
    unidadeLeitura: '31020003',
    leiturista: '5GL',
    leitAnt: '51200',
    leitExec: '52840',
    imp: 'Lida',
    qtdDig: '',
    apres: 'AV AMAZONAS, 2104 - BARRO PRETO',
  },
  {
    consumidora: 'MARCOS VINICIUS ALMEIDA',
    medidor: 'ARL211119872',
    tReg: '00060',
    sequencia: '00140',
    data: '11/6/2026',
    hora: '10:18:03 AM',
    unidadeLeitura: '31020004',
    leiturista: '80J',
    leitAnt: '04110',
    leitExec: '04298',
    imp: '',
    qtdDig: '',
    apres: 'RUA BAHIA, 412 - LOURDES',
  },
  {
    consumidora: 'CLARA NOGUEIRA SANTOS',
    medidor: 'WEG334918201',
    tReg: '00060',
    sequencia: '00141',
    data: '11/6/2026',
    hora: '10:22:49 AM',
    unidadeLeitura: '31020005',
    leiturista: '89O',
    leitAnt: '19320',
    leitExec: '19505',
    imp: 'Lida',
    qtdDig: '05',
    apres: 'RUA TUPIS, 77 - CENTRO',
  },
  {
    consumidora: 'REGISTRO EXCLUIDO FILTRO',
    medidor: 'XXX000000999',
    tReg: '00060',
    sequencia: '00142',
    data: '11/6/2026',
    hora: '10:27:10 AM',
    unidadeLeitura: '31029999',
    leiturista: '99Z',
    leitAnt: '10000',
    leitExec: '10100',
    imp: '',
    qtdDig: '',
    apres: 'RUA TESTE FILTRO, 01',
  },
  {
    consumidora: 'PADARIA ESTRELA DALVA',
    medidor: 'ELG772109901',
    tReg: '00120',
    sequencia: '00143',
    data: '11/6/2026',
    hora: '10:31:15 AM',
    unidadeLeitura: '31020006',
    leiturista: '51O',
    leitAnt: '32150',
    leitExec: '32910',
    imp: '',
    qtdDig: '',
    apres: 'RUA RIO DE JANEIRO, 1180 - LOURDES',
  },
  {
    consumidora: 'HELENA CRISTINA SOUZA',
    medidor: 'ARL211120045',
    tReg: '00060',
    sequencia: '00144',
    data: '11/6/2026',
    hora: '10:35:08 AM',
    unidadeLeitura: '31020007',
    leiturista: '5UL',
    leitAnt: '07890',
    leitExec: '08045',
    imp: 'Lida',
    qtdDig: '05',
    apres: 'RUA TIMBIRAS, 530 - FUNCIONARIOS',
  },
];

/**
 * Gera uma planilha Matriz de exemplo no layout bruto "Seqüência de Leitura Executada"
 * da CEMIG, com cabeçalhos repetidos, células mescladas, zeros à esquerda,
 * datas/horas originais e blocos verticais compatíveis com as 3 macros VBA.
 */
export async function generateSampleCemigMatrizBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Leituras');

  const setText = (r: number, c: number, val: string) => {
    const cell = ws.getRow(r).getCell(c);
    cell.value = val;
    cell.numFmt = '@';
  };

  // Cabeçalho institucional CEMIG (linhas 1 a 10) com células mescladas
  setText(1, 1, 'CEMIG Distribuição S.A.');
  setText(1, 10, 'Pág. 1/446');
  ws.mergeCells(1, 1, 1, 4);

  setText(2, 1, 'Seqüência de Leitura Executada');
  ws.mergeCells(2, 1, 2, 6);

  setText(3, 1, 'Contrato: 4600018294');
  setText(3, 6, 'Mês/Ano: 11/2026');
  setText(4, 1, 'Unidade de Leitura Inicial: 31020001');
  setText(4, 6, 'Unidade de Leitura Final: 31020007');
  setText(5, 1, 'Tipo de Leitura: CONVENCIONAL');
  setText(6, 1, 'Data Início: 11/6/2026');
  setText(7, 1, 'Data Término: 11/6/2026');
  setText(8, 1, 'Regional: METROPOLITANA');
  setText(9, 1, 'Lote: 001');
  setText(10, 1, '--------------------------------------------------------------------------------');

  // Linha 11 (cabeçalho da tabela pós-tratamento)
  const h = 11;
  setText(h, 1, 'Consumidora');
  setText(h, 2, 'Medidor');
  setText(h, 3, 'T. Reg');
  setText(h, 4, 'Seqüência de Leitura Executada');
  setText(h, 5, 'Contrato_Aux'); // deletada no passo 1
  setText(h, 6, 'Data');
  setText(h, 7, 'Hora'); // copiada para F no passo 5
  setText(h + 2, 7, 'Leit. Ant.'); // sobe 2 linhas para G no passo 6
  setText(h, 8, 'Tipo_Aux1'); // deletada no passo 2
  setText(h, 9, 'Tipo_Aux2'); // deletada no passo 2
  setText(h, 10, 'Leit. Exec');
  setText(h, 11, ''); // Imp. / Coluna8
  setText(h, 12, 'Qtd Dig.');
  setText(h, 13, 'OC');
  setText(h + 2, 13, 'Apres.'); // sobe 2 linhas para J no passo 4

  const writeRecordBlock = (r: number, rec: SampleMeterRecord) => {
    setText(r - 5, 1, `Unidade de Leitura: ${rec.unidadeLeitura}`);
    setText(r - 4, 1, `Leiturista: ${rec.leiturista}`);

    setText(r, 1, rec.consumidora);
    setText(r, 2, rec.medidor);
    setText(r, 3, rec.tReg);
    setText(r, 4, rec.sequencia);
    setText(r, 5, '4600018294');
    setText(r, 6, rec.data);
    setText(r, 7, rec.hora);
    setText(r + 2, 7, rec.leitAnt);
    setText(r, 8, 'CONV');
    setText(r, 9, 'NORMAL');
    setText(r, 10, rec.leitExec);
    setText(r, 11, rec.imp);
    setText(r, 12, rec.qtdDig);
    setText(r, 13, '00');
    setText(r + 2, 13, rec.apres);
  };

  let cursorRow = 18;
  SAMPLE_RECORDS.forEach((rec, idx) => {
    if (idx === 4) {
      // Quebra de página repetida (Pág. 2/446)
      setText(cursorRow, 1, 'CEMIG Distribuição S.A.');
      setText(cursorRow, 10, 'Pág. 2/446');
      ws.mergeCells(cursorRow, 1, cursorRow, 4);
      setText(cursorRow + 1, 1, 'Seqüência de Leitura Executada');
      setText(cursorRow + 2, 1, 'Unidade de Leitura Final: 31020007');
      cursorRow += 5;
    }

    writeRecordBlock(cursorRow, rec);

    // Duplicata proposital do 1o registro para validar deduplicação (Passo 9)
    if (idx === 0) {
      cursorRow += 8;
      writeRecordBlock(cursorRow, rec);
    }

    cursorRow += 8;
  });

  const arr = await wb.xlsx.writeBuffer();
  return Buffer.from(arr);
}
