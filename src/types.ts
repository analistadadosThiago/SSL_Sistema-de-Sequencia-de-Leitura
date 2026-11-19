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

export interface PreviewData {
  headers: string[];
  rows: string[][];
  totalDataRows: number;
}

export interface ControleHorarioRow {
  matricula: string;
  data: string;
  horaInicial: string;
  horaFinal: string;
  tempoServico: string;
}

export interface ProcessApiResponse {
  fileName: string;
  fileBase64: string;
  relatorioGeralFileName: string | null;
  relatorioGeralBase64: string | null;
  stats: ProcessingStats;
  logs: StageStepLog[];
  warnings: string[];
  preview: PreviewData;
  controleHorario: ControleHorarioRow[];
  stage2Sample: Array<{ unidade: string; matricula: string }>;
  sourceName?: string;
}
