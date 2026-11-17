import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Filter,
  Play,
  RotateCcw,
  Search,
  Upload,
  X,
} from 'lucide-react';
import type { ProcessApiResponse } from './types';

const DEFAULT_LEITURISTAS = [
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

const PAGE_SIZE = 30;
const FETCH_TIMEOUT_MS = 300_000;

type WorkspaceTab = 'preview' | 'logs';

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function downloadBase64Xlsx(base64: string, fileName: string): void {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function safeFetchJson(
  url: string,
  init: RequestInit,
): Promise<Record<string, unknown>> {
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

  try {
    const response = await fetch(url, {
      ...init,
      signal: controller.signal,
    });

    const contentType = response.headers.get('content-type') || '';
    if (!contentType.toLowerCase().includes('application/json')) {
      const text = await response.text();
      throw new Error(
        text || `Resposta inesperada do servidor (HTTP ${response.status}).`,
      );
    }

    const payload = (await response.json()) as Record<string, unknown>;
    if (!response.ok) {
      const msg =
        (typeof payload.erro === 'string' && payload.erro) ||
        (typeof payload.error === 'string' && payload.error) ||
        `Erro HTTP ${response.status}`;
      const detail =
        typeof payload.detalhe === 'string' && payload.detalhe !== msg
          ? ` (${payload.detalhe})`
          : '';
      throw new Error(`${msg}${detail}`);
    }

    return payload;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

function getVisiblePageNumbers(current: number, total: number): Array<number | 'ellipsis'> {
  if (total <= 7) {
    return Array.from({ length: total }, (_, i) => i + 1);
  }
  const pages: Array<number | 'ellipsis'> = [1];
  const start = Math.max(2, current - 1);
  const end = Math.min(total - 1, current + 1);

  if (start > 2) pages.push('ellipsis');
  for (let p = start; p <= end; p++) pages.push(p);
  if (end < total - 1) pages.push('ellipsis');
  pages.push(total);
  return pages;
}

export default function App() {
  const [matrizFile, setMatrizFile] = useState<File | null>(null);
  const [matrizDragActive, setMatrizDragActive] = useState(false);

  const [showFilterConfig, setShowFilterConfig] = useState(false);
  const [leituristaText, setLeituristaText] = useState(
    DEFAULT_LEITURISTAS.join('\n'),
  );

  const [isProcessing, setIsProcessing] = useState(false);
  const [progressPct, setProgressPct] = useState(0);
  const [progressLabel, setProgressLabel] = useState('');
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [result, setResult] = useState<ProcessApiResponse | null>(null);

  const [activeTab, setActiveTab] = useState<WorkspaceTab>('preview');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedLeituristaFilter, setSelectedLeituristaFilter] = useState('ALL');
  const [currentPage, setCurrentPage] = useState(1);

  const matrizInputRef = useRef<HTMLInputElement>(null);
  const pipelineSectionRef = useRef<HTMLElement>(null);
  const workspaceSectionRef = useRef<HTMLElement>(null);

  useEffect(() => {
    const onUnhandledRejection = (e: PromiseRejectionEvent) => {
      console.error(e.reason);
    };
    const onWindowError = (e: ErrorEvent) => {
      console.error(e.error || e.message);
    };
    window.addEventListener('unhandledrejection', onUnhandledRejection);
    window.addEventListener('error', onWindowError);
    return () => {
      window.removeEventListener('unhandledrejection', onUnhandledRejection);
      window.removeEventListener('error', onWindowError);
    };
  }, []);

  const parsedAllowedLeituristas = useMemo(() => {
    return leituristaText
      .split(/[\n,;]+/)
      .map((item) => item.trim())
      .filter(Boolean)
      .map((item) =>
        item.startsWith('Leiturista:') ? item : `Leiturista: ${item}`,
      );
  }, [leituristaText]);

  const handleMatrizSelect = useCallback((file: File | undefined) => {
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith('.xlsx') && !lower.endsWith('.xls')) {
      setErrorMsg(
        'Formato inválido para a planilha Matriz. Envie um arquivo .xlsx ou .xls.',
      );
      return;
    }
    setErrorMsg(null);
    setMatrizFile(file);
  }, []);

  const runProcessing = useCallback(async () => {
    if (!matrizFile) {
      setErrorMsg(
        'Selecione a planilha Matriz (.xlsx ou .xls) antes de iniciar o processamento.',
      );
      return;
    }
    setErrorMsg(null);
    setIsProcessing(true);
    setProgressPct(18);
    setProgressLabel('Lendo e desmesclando planilha...');

    const formData = new FormData();
    formData.append('matriz', matrizFile);
    formData.append(
      'allowedLeituristas',
      JSON.stringify(parsedAllowedLeituristas),
    );
    formData.append('populatePlanilha3FromStage2', 'true');

    const timer1 = window.setTimeout(() => {
      setProgressPct(52);
      setProgressLabel('Executando normalização, ordenação e formatação...');
    }, 260);
    const timer2 = window.setTimeout(() => {
      setProgressPct(82);
      setProgressLabel('Gerando planilha final...');
    }, 620);

    try {
      const payload = (await safeFetchJson('/api/process', {
        method: 'POST',
        body: formData,
      })) as unknown as ProcessApiResponse;

      window.clearTimeout(timer1);
      window.clearTimeout(timer2);

      setProgressPct(100);
      setProgressLabel('Concluído');
      setResult(payload);
      setActiveTab('preview');
      setSelectedLeituristaFilter('ALL');
      setSearchQuery('');
      setCurrentPage(1);
    } catch (err) {
      window.clearTimeout(timer1);
      window.clearTimeout(timer2);
      setProgressPct(0);
      setProgressLabel('');
      const msg =
        err instanceof Error
          ? err.name === 'AbortError'
            ? 'Tempo limite excedido (300s) ao processar a planilha.'
            : err.message
          : 'Erro de comunicação ao processar a planilha.';
      setErrorMsg(msg);
    } finally {
      setIsProcessing(false);
    }
  }, [matrizFile, parsedAllowedLeituristas]);

  const runSampleProcessing = useCallback(async () => {
    setErrorMsg(null);
    setIsProcessing(true);
    setProgressPct(35);
    setProgressLabel('Processando exemplo...');

    try {
      const payload = (await safeFetchJson('/api/process-sample', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ populatePlanilha3FromStage2: true }),
      })) as unknown as ProcessApiResponse;

      setProgressPct(100);
      setProgressLabel('Concluído');
      setResult(payload);
      setActiveTab('preview');
      setSelectedLeituristaFilter('ALL');
      setSearchQuery('');
      setCurrentPage(1);
    } catch (err) {
      setProgressPct(0);
      setProgressLabel('');
      setErrorMsg(
        err instanceof Error
          ? err.message
          : 'Erro ao executar a planilha de exemplo.',
      );
    } finally {
      setIsProcessing(false);
    }
  }, []);

  const handleReset = useCallback(() => {
    setMatrizFile(null);
    setResult(null);
    setErrorMsg(null);
    setProgressPct(0);
    setProgressLabel('');
    setSearchQuery('');
    setSelectedLeituristaFilter('ALL');
    setCurrentPage(1);
    if (matrizInputRef.current) matrizInputRef.current.value = '';
  }, []);

  const handleDownloadResult = useCallback(() => {
    if (!result) return;
    downloadBase64Xlsx(result.fileBase64, result.fileName);
  }, [result]);

  const navigateToTab = useCallback((tab: WorkspaceTab) => {
    setActiveTab(tab);
    workspaceSectionRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, []);

  const filteredPreviewRows = useMemo(() => {
    if (!result) return [];
    const q = searchQuery.trim().toLowerCase();
    return result.preview.rows.filter((row) => {
      if (selectedLeituristaFilter !== 'ALL') {
        const rowMatr = (row[4] ?? '').trim();
        if (rowMatr !== selectedLeituristaFilter) {
          return false;
        }
      }
      if (!q) return true;
      return row.some((cell) => cell.toLowerCase().includes(q));
    });
  }, [result, searchQuery, selectedLeituristaFilter]);

  const totalPages = Math.max(
    1,
    Math.ceil(filteredPreviewRows.length / PAGE_SIZE),
  );
  const safeCurrentPage = Math.min(currentPage, totalPages);

  const paginatedRows = useMemo(() => {
    const startIdx = (safeCurrentPage - 1) * PAGE_SIZE;
    return filteredPreviewRows.slice(startIdx, startIdx + PAGE_SIZE);
  }, [filteredPreviewRows, safeCurrentPage]);

  const leituristaCountsInFiltered = useMemo(() => {
    const counts: Record<string, number> = {};
    for (const row of filteredPreviewRows) {
      const key = (row[4] ?? '').trim() || 'Sem Leiturista';
      counts[key] = (counts[key] ?? 0) + 1;
    }
    return counts;
  }, [filteredPreviewRows]);

  const leituristaOptions = useMemo(() => {
    if (!result) return [];
    return Object.keys(result.stats.leituristaBreakdown).sort();
  }, [result]);

  const pageNumbers = useMemo(
    () => getVisiblePageNumbers(safeCurrentPage, totalPages),
    [safeCurrentPage, totalPages],
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Top Bar Contract: exatamente 3 zonas separadas por gap-8 */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200">
        <div className="max-w-[1380px] mx-auto px-6 h-16 flex items-center justify-between gap-8">
          <a
            href="#pipeline"
            onClick={(e) => {
              e.preventDefault();
              pipelineSectionRef.current?.scrollIntoView({ behavior: 'smooth' });
            }}
            className="font-display text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
          >
            Controle de Sequenciamento de Leitura
          </a>

          <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
            <button
              type="button"
              onClick={() =>
                pipelineSectionRef.current?.scrollIntoView({
                  behavior: 'smooth',
                })
              }
              className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Arquivos de Entrada
            </button>
            <button
              type="button"
              onClick={() => navigateToTab('preview')}
              className={`hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'preview' ? 'text-teal-700 font-semibold' : ''
              }`}
            >
              Pré-visualização
            </button>
            <button
              type="button"
              onClick={() => navigateToTab('logs')}
              className={`hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'logs' ? 'text-teal-700 font-semibold' : ''
              }`}
            >
              Log Resumido
            </button>
          </nav>

          <div className="shrink-0">
            {result ? (
              <button
                type="button"
                onClick={handleDownloadResult}
                className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
              >
                <Download className="w-4 h-4 shrink-0" />
                <span>Baixar Planilha</span>
              </button>
            ) : (
              <button
                type="button"
                onClick={runSampleProcessing}
                disabled={isProcessing}
                className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-teal-700 hover:bg-teal-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
              >
                <Play className="w-4 h-4 shrink-0" />
                <span>Testar Exemplo</span>
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-[1380px] w-full mx-auto px-6 py-8 space-y-8">
        {/* Cabeçalho e Card Único "Arquivos de Entrada" */}
        <section ref={pipelineSectionRef} id="pipeline" className="space-y-6">
          <div className="flex flex-col lg:flex-row lg:items-end justify-between gap-4 border-b border-slate-200 pb-6">
            <div className="space-y-1.5 max-w-3xl">
              <h1 className="font-display text-2xl sm:text-3xl font-bold tracking-tight text-slate-900 text-balance">
                Controle de Sequenciamento de Leitura
              </h1>
              <p className="text-sm sm:text-base text-slate-600 leading-relaxed">
                Importe os dados a serem tratados e faça o download depois do tratamento.
              </p>
            </div>
            <div className="flex items-center gap-3 text-xs text-slate-500 shrink-0">
              <span>Processamento em memória</span>
              <span aria-hidden="true">·</span>
              <a
                href="/api/sample-matriz"
                download="Matriz_Leituras_Exemplo.xlsx"
                className="text-teal-700 font-medium hover:underline underline-offset-4 whitespace-nowrap"
              >
                Baixar Matriz de Exemplo
              </a>
            </div>
          </div>

          {/* Card Único: Arquivos de Entrada */}
          <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-5">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h2 className="font-display text-base font-semibold text-slate-900">
                  Arquivos de Entrada
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Envie a planilha Matriz (.xlsx ou .xls) para tratamento
                </p>
              </div>
              {matrizFile && (
                <button
                  type="button"
                  onClick={() => {
                    setMatrizFile(null);
                    if (matrizInputRef.current) matrizInputRef.current.value = '';
                  }}
                  className="inline-flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-900 px-2 py-1 rounded cursor-pointer"
                  aria-label="Remover arquivo Matriz"
                >
                  <X className="w-3.5 h-3.5" />
                  <span>Remover</span>
                </button>
              )}
            </div>

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setMatrizDragActive(true);
              }}
              onDragLeave={() => setMatrizDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setMatrizDragActive(false);
                handleMatrizSelect(e.dataTransfer.files?.[0]);
              }}
              onClick={() => matrizInputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  matrizInputRef.current?.click();
                }
              }}
              className={`border border-dashed rounded-lg px-5 py-8 text-center cursor-pointer transition-colors focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700 ${
                matrizDragActive
                  ? 'border-teal-600 bg-teal-50/50'
                  : matrizFile
                    ? 'border-teal-300 bg-slate-50'
                    : 'border-slate-300 hover:border-teal-600 hover:bg-slate-50/70'
              }`}
            >
              <input
                ref={matrizInputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => handleMatrizSelect(e.target.files?.[0])}
              />
              {matrizFile ? (
                <div className="flex flex-col items-center gap-1.5">
                  <FileSpreadsheet className="w-7 h-7 text-teal-700" />
                  <p className="text-sm font-semibold text-slate-900 break-all">
                    {matrizFile.name}
                  </p>
                  <p className="text-xs text-slate-500 font-mono tabular-nums">
                    {formatFileSize(matrizFile.size)} · Pronto para processamento
                  </p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-2">
                  <Upload className="w-6 h-6 text-slate-400" />
                  <p className="text-sm font-medium text-slate-700">
                    Arraste a planilha Matriz aqui ou clique para selecionar
                  </p>
                  <p className="text-xs text-slate-500">
                    Suporta arquivos .xlsx e .xls até 50 MB
                  </p>
                </div>
              )}
            </div>

            <div className="flex flex-wrap items-center justify-between gap-4 pt-1">
              <div className="flex flex-wrap items-center gap-3">
                <button
                  type="button"
                  onClick={runProcessing}
                  disabled={isProcessing || !matrizFile}
                  className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-teal-700 hover:bg-teal-800 disabled:bg-slate-300 disabled:text-slate-500 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:cursor-not-allowed focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
                >
                  <Play className="w-4 h-4 shrink-0" />
                  <span>{isProcessing ? 'Processando...' : 'Processar'}</span>
                </button>

                {result && (
                  <button
                    type="button"
                    onClick={handleDownloadResult}
                    className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
                  >
                    <Download className="w-4 h-4 shrink-0" />
                    <span>Baixar Matriz_ControleDeSequenciaLeitura.xlsx</span>
                  </button>
                )}

                <button
                  type="button"
                  onClick={runSampleProcessing}
                  disabled={isProcessing}
                  className="inline-flex items-center gap-2 px-4 py-2.5 text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <span>Usar Exemplo</span>
                </button>

                {(matrizFile || result) && (
                  <button
                    type="button"
                    onClick={handleReset}
                    disabled={isProcessing}
                    className="inline-flex items-center gap-1.5 px-3 py-2.5 text-sm font-medium text-slate-600 hover:text-slate-900 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <RotateCcw className="w-4 h-4 shrink-0" />
                    <span>Limpar</span>
                  </button>
                )}
              </div>

              <button
                type="button"
                onClick={() => setShowFilterConfig((prev) => !prev)}
                className="inline-flex items-center gap-2 px-3 py-2 text-xs font-medium text-slate-700 hover:text-slate-900 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Filter className="w-3.5 h-3.5 shrink-0" />
                <span>
                  Filtro Leituristas ({parsedAllowedLeituristas.length} códigos)
                </span>
              </button>
            </div>

            {showFilterConfig && (
              <div className="pt-4 border-t border-slate-200 space-y-2">
                <div className="flex items-center justify-between">
                  <label
                    htmlFor="leituristas-textarea"
                    className="text-xs font-semibold text-slate-800"
                  >
                    Leituristas mantidos no filtro
                  </label>
                  <button
                    type="button"
                    onClick={() =>
                      setLeituristaText(DEFAULT_LEITURISTAS.join('\n'))
                    }
                    className="text-xs text-teal-700 hover:underline cursor-pointer"
                  >
                    Restaurar 15 padrões
                  </button>
                </div>
                <textarea
                  id="leituristas-textarea"
                  rows={3}
                  value={leituristaText}
                  onChange={(e) => setLeituristaText(e.target.value)}
                  className="w-full font-mono text-xs bg-slate-50 border border-slate-200 rounded-lg p-3 text-slate-800 focus:outline-none focus:border-teal-600"
                />
              </div>
            )}

            {(isProcessing || progressPct > 0) && (
              <div className="pt-2 space-y-1.5">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-medium text-slate-700">
                    {progressLabel}
                  </span>
                  <span className="font-mono tabular-nums text-slate-600">
                    {progressPct}%
                  </span>
                </div>
                <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-teal-700 transition-all duration-200"
                    style={{ width: `${progressPct}%` }}
                  />
                </div>
              </div>
            )}

            {errorMsg && (
              <div
                role="alert"
                className="flex items-start gap-3 p-4 rounded-lg bg-red-50 border border-red-200 text-red-900 text-sm"
              >
                <AlertTriangle className="w-5 h-5 text-red-700 shrink-0 mt-0.5" />
                <div>
                  <p className="font-semibold">Erro no processamento</p>
                  <p className="text-red-800 mt-0.5">{errorMsg}</p>
                </div>
              </div>
            )}
          </div>
        </section>

        {/* Faixa de KPIs */}
        <section className="grid grid-cols-2 lg:grid-cols-4 bg-white border border-slate-200 rounded-xl divide-y lg:divide-y-0 lg:divide-x divide-slate-200">
          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">
              Linhas Brutas Lidas
            </p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result ? result.stats.rawRows.toLocaleString('pt-BR') : '—'}
            </p>
            <p className="text-xs text-slate-500">
              {result
                ? `Aba ${result.stats.sheetName} · ${result.stats.rawCols} colunas`
                : 'Aguardando arquivo Matriz'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">
              Após Deduplicação
            </p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result
                ? result.stats.afterDeduplicationRows.toLocaleString('pt-BR')
                : '—'}
            </p>
            <p className="text-xs text-slate-500">
              {result
                ? `${(result.stats.rawRows - result.stats.afterDeduplicationRows).toLocaleString('pt-BR')} duplicatas removidas`
                : 'Deduplicação por linha'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">
              Após Filtro Leiturista
            </p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result
                ? result.stats.afterFilterRows.toLocaleString('pt-BR')
                : '—'}
            </p>
            <p className="text-xs text-slate-500">
              {result
                ? `${result.stats.uniqueUnitsCount.toLocaleString('pt-BR')} unidades distintas`
                : 'Leituristas selecionados'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">
              Linhas Finais
            </p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-teal-700">
              {result ? result.stats.finalRows.toLocaleString('pt-BR') : '—'}
            </p>
            <p className="text-xs text-slate-500">
              {result
                ? `${result.stats.finalCols} colunas · ${result.stats.durationMs} ms`
                : 'Tabela completa do tratamento'}
            </p>
          </div>
        </section>

        {/* Área de Pré-visualização Paginada e Log Resumido */}
        <section
          ref={workspaceSectionRef}
          className="bg-white border border-slate-200 rounded-xl overflow-hidden"
        >
          <div className="px-6 py-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="inline-flex items-center gap-1 p-1 bg-slate-100 rounded-lg self-start">
              <button
                type="button"
                onClick={() => setActiveTab('preview')}
                className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  activeTab === 'preview'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Pré-visualização
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('logs')}
                className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  activeTab === 'logs'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Log Resumido
              </button>
            </div>

            {result && activeTab === 'preview' && (
              <div className="flex flex-wrap items-center gap-3">
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      setCurrentPage(1);
                    }}
                    placeholder="Buscar instalação, medidor, endereço..."
                    className="pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg w-64 text-slate-800 focus:outline-none focus:border-teal-600"
                  />
                </div>
                <select
                  value={selectedLeituristaFilter}
                  onChange={(e) => {
                    setSelectedLeituristaFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  aria-label="Filtrar por leiturista"
                  className="px-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:outline-none focus:border-teal-600"
                >
                  <option value="ALL">Todos os leituristas</option>
                  {leituristaOptions.map((code) => (
                    <option key={code} value={code}>
                      Leiturista: {code} ({result.stats.leituristaBreakdown[code]})
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>

          {activeTab === 'preview' && (
            <div>
              {isProcessing ? (
                <div className="p-6 space-y-3" aria-busy="true">
                  {Array.from({ length: 6 }).map((_, idx) => (
                    <div
                      key={idx}
                      className="h-9 bg-slate-100 rounded animate-pulse"
                    />
                  ))}
                </div>
              ) : !result ? (
                <div className="px-6 py-16 text-center space-y-4">
                  <FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto" />
                  <div className="space-y-1 max-w-md mx-auto">
                    <p className="text-sm font-semibold text-slate-900">
                      Nenhuma planilha processada ainda
                    </p>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Importe os dados a serem tratados e faça o download depois do tratamento.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={runSampleProcessing}
                    className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer"
                  >
                    <Play className="w-3.5 h-3.5" />
                    <span>Usar Exemplo</span>
                  </button>
                </div>
              ) : (
                <div>
                  <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-4 text-xs text-slate-700">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-teal-700 shrink-0" />
                      <span className="font-medium">
                        Tabela completa do tratamento.
                      </span>
                    </div>
                    <span className="font-mono tabular-nums text-slate-500">
                      Página {safeCurrentPage} de {totalPages}
                    </span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left">
                      <thead>
                        <tr className="bg-slate-900 text-white text-xs">
                          <th className="px-3 py-2.5 font-mono font-medium text-slate-300 border-r border-slate-800 whitespace-nowrap">
                            #
                          </th>
                          {result.preview.headers.map((header, idx) => (
                            <th
                              key={idx}
                              className={`px-3 py-2.5 font-medium whitespace-nowrap border-r border-slate-800 last:border-r-0 ${
                                idx >= 14 ? 'bg-teal-800' : ''
                              }`}
                            >
                              {header}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-xs">
                        {paginatedRows.length === 0 ? (
                          <tr>
                            <td
                              colSpan={result.preview.headers.length + 1}
                              className="px-6 py-10 text-center text-slate-500"
                            >
                              Nenhum registro corresponde ao filtro atual.
                            </td>
                          </tr>
                        ) : (
                          paginatedRows.map((row, idxInPage) => {
                            const globalRowIndex =
                              (safeCurrentPage - 1) * PAGE_SIZE + idxInPage;
                            const leiturista =
                              (row[4] ?? '').trim() || 'Sem Leiturista';
                            const prevLeiturista =
                              idxInPage > 0
                                ? (paginatedRows[idxInPage - 1]?.[4] ?? '').trim() ||
                                  'Sem Leiturista'
                                : null;
                            const showGroupHeader =
                              idxInPage === 0 || leiturista !== prevLeiturista;

                            return (
                              <React.Fragment key={globalRowIndex}>
                                {showGroupHeader && (
                                  <tr className="bg-slate-100 border-y border-slate-200">
                                    <td
                                      colSpan={result.preview.headers.length + 1}
                                      className="px-3 py-2 text-xs font-semibold text-slate-800"
                                    >
                                      <span className="font-mono">
                                        Leiturista: {leiturista}
                                      </span>
                                      <span className="text-slate-500 font-normal ml-2 font-mono tabular-nums">
                                        ·{' '}
                                        {(
                                          leituristaCountsInFiltered[leiturista] ?? 0
                                        ).toLocaleString('pt-BR')}{' '}
                                        registro(s)
                                      </span>
                                    </td>
                                  </tr>
                                )}
                                <tr className="hover:bg-slate-50 transition-colors h-10">
                                  <td className="px-3 py-2 font-mono tabular-nums text-slate-400 border-r border-slate-100 whitespace-nowrap">
                                    {globalRowIndex + 1}
                                  </td>
                                  {row.map((val, cIdx) => {
                                    const isNumericReading =
                                      cIdx === 9 || cIdx === 10 || cIdx === 11;
                                    const isMonoField =
                                      isNumericReading ||
                                      cIdx === 2 ||
                                      cIdx === 3 ||
                                      cIdx === 4 ||
                                      cIdx === 5 ||
                                      cIdx === 6 ||
                                      cIdx === 7;
                                    return (
                                      <td
                                        key={cIdx}
                                        className={`px-3 py-2 whitespace-nowrap border-r border-slate-100 last:border-r-0 ${
                                          isMonoField ? 'font-mono tabular-nums' : ''
                                        } ${
                                          isNumericReading
                                            ? 'text-right font-medium text-slate-900'
                                            : 'text-slate-700'
                                        } ${
                                          cIdx === 11
                                            ? 'bg-teal-50/40 text-teal-900 font-semibold'
                                            : ''
                                        }`}
                                      >
                                        {val || '—'}
                                      </td>
                                    );
                                  })}
                                </tr>
                              </React.Fragment>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* Controles de Paginação (30 registros por página) */}
                  <div className="px-6 py-3.5 border-t border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-4 text-xs text-slate-600">
                    <div>
                      Exibindo{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {filteredPreviewRows.length === 0
                          ? 0
                          : (safeCurrentPage - 1) * PAGE_SIZE + 1}
                      </strong>
                      –
                      <strong className="font-mono tabular-nums text-slate-900">
                        {Math.min(
                          safeCurrentPage * PAGE_SIZE,
                          filteredPreviewRows.length,
                        )}
                      </strong>{' '}
                      de{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {filteredPreviewRows.length.toLocaleString('pt-BR')}
                      </strong>{' '}
                      registros · Página{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {safeCurrentPage} de {totalPages}
                      </strong>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage((p) => Math.max(1, p - 1))
                        }
                        disabled={safeCurrentPage <= 1}
                        className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium bg-white border border-slate-200 rounded-md text-slate-700 hover:bg-slate-100 disabled:opacity-45 disabled:cursor-not-allowed cursor-pointer whitespace-nowrap"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                        <span>Anterior</span>
                      </button>

                      {pageNumbers.map((item, idx) =>
                        item === 'ellipsis' ? (
                          <span
                            key={`e-${idx}`}
                            className="px-2 py-1 text-slate-400 font-mono"
                          >
                            …
                          </span>
                        ) : (
                          <button
                            key={item}
                            type="button"
                            onClick={() => setCurrentPage(item)}
                            className={`px-2.5 py-1.5 text-xs font-mono tabular-nums rounded-md border transition-colors cursor-pointer ${
                              item === safeCurrentPage
                                ? 'bg-teal-700 text-white border-teal-700 font-semibold'
                                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                            }`}
                          >
                            {item}
                          </button>
                        ),
                      )}

                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage((p) => Math.min(totalPages, p + 1))
                        }
                        disabled={safeCurrentPage >= totalPages}
                        className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium bg-white border border-slate-200 rounded-md text-slate-700 hover:bg-slate-100 disabled:opacity-45 disabled:cursor-not-allowed cursor-pointer whitespace-nowrap"
                      >
                        <span>Próxima</span>
                        <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'logs' && (
            <div className="p-6 space-y-6">
              {!result ? (
                <div className="py-10 text-center text-sm text-slate-500">
                  Processe uma planilha para visualizar o log resumido.
                </div>
              ) : (
                <>
                  <div className="flex flex-wrap items-center justify-between gap-4">
                    <div className="flex items-center gap-2 text-sm font-semibold text-emerald-800">
                      <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                      <span>
                        Tabela completa do tratamento.{' '}
                        {result.stats.finalRows.toLocaleString('pt-BR')} registros
                        prontos em {result.stats.durationMs} ms.
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={handleDownloadResult}
                      className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg cursor-pointer"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Baixar Matriz_ControleDeSequenciaLeitura.xlsx</span>
                    </button>
                  </div>

                  <div className="border border-slate-200 rounded-lg divide-y divide-slate-200">
                    {result.logs.map((entry, idx) => (
                      <div
                        key={idx}
                        className="px-4 py-3 grid grid-cols-1 md:grid-cols-12 gap-2 text-xs items-baseline"
                      >
                        <span className="md:col-span-2 font-mono font-semibold text-teal-800">
                          {entry.stage}
                        </span>
                        <span className="md:col-span-4 font-medium text-slate-900">
                          {entry.step}
                        </span>
                        <span className="md:col-span-6 text-slate-600">
                          {entry.detail}
                        </span>
                      </div>
                    ))}
                  </div>

                  <div className="space-y-3">
                    <h3 className="font-display text-sm font-semibold text-slate-900">
                      Registros por Leiturista
                    </h3>
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 text-xs">
                      {Object.entries(result.stats.leituristaBreakdown)
                        .sort(([a], [b]) => a.localeCompare(b))
                        .map(([code, count]) => (
                          <div
                            key={code}
                            className="px-4 py-2.5 border border-slate-200 rounded-lg flex items-center justify-between"
                          >
                            <span className="font-mono font-medium text-slate-800">
                              {code}
                            </span>
                            <span className="font-mono tabular-nums text-slate-600">
                              {count.toLocaleString('pt-BR')}
                            </span>
                          </div>
                        ))}
                    </div>
                  </div>
                </>
              )}
            </div>
          )}
        </section>
      </main>
    </div>
  );
}
