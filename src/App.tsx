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

type WorkspaceTab = 'preview' | 'controleHorario';

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatFooterDateTime(now: Date): string {
  const data = now.toLocaleDateString('pt-BR');
  const hora = now.toLocaleTimeString('pt-BR', {
    hour: '2-digit',
    minute: '2-digit',
  });
  return `${data} ${hora}`;
}

function downloadBase64Xlsx(base64: string, fileName: string): void {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const isXlsm = fileName.toLowerCase().endsWith('.xlsm');
  const blob = new Blob([bytes], {
    type: isXlsm
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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

function isHtmlOrCookieCheck(text: string): boolean {
  const lower = text.toLowerCase();
  return (
    lower.includes('<!doctype html') ||
    lower.includes('<html') ||
    text.includes('Cookie check') ||
    text.includes('Authenticate in new window')
  );
}

const COOKIE_CHECK_ERROR_MSG =
  "Sessão do Google AI Studio expirada ou cookies bloqueados. Abra o preview em uma nova aba, clique em 'Authenticate in new window' → 'Grant permission' → 'Close and continue', e tente novamente.";

const HEALTH_ERROR_MSG =
  'Backend indisponível ou sessão expirada. Recarregue a página e autentique novamente.';

async function checkBackendHealth(): Promise<void> {
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), 15_000);
  try {
    const response = await fetch('/health', {
      method: 'GET',
      credentials: 'same-origin',
      signal: controller.signal,
    });
    const ct = response.headers.get('content-type') || '';
    if (!ct.toLowerCase().includes('application/json')) {
      await response.text().catch(() => '');
      throw new Error(HEALTH_ERROR_MSG);
    }
    const data = (await response.json()) as { ok?: boolean };
    if (!response.ok || !data?.ok) {
      throw new Error(HEALTH_ERROR_MSG);
    }
  } catch (err) {
    if (err instanceof Error && err.message === HEALTH_ERROR_MSG) {
      throw err;
    }
    throw new Error(HEALTH_ERROR_MSG);
  } finally {
    window.clearTimeout(timeoutId);
  }
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
      credentials: 'same-origin',
      signal: controller.signal,
    });

    const ct = response.headers.get('content-type') || '';
    if (!ct.toLowerCase().includes('application/json')) {
      const text = await response.text();
      if (isHtmlOrCookieCheck(text)) {
        throw new Error(COOKIE_CHECK_ERROR_MSG);
      }
      throw new Error(
        `Servidor retornou resposta inesperada (HTTP ${response.status}). Tente novamente.`,
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

  const [consistenciaFile, setConsistenciaFile] = useState<File | null>(null);
  const [consistenciaDragActive, setConsistenciaDragActive] = useState(false);

  const [nosbFile, setNosbFile] = useState<File | null>(null);
  const [nosbDragActive, setNosbDragActive] = useState(false);

  const [fotosFile, setFotosFile] = useState<File | null>(null);
  const [fotosDragActive, setFotosDragActive] = useState(false);

  const [modeloRGeralFile, setModeloRGeralFile] = useState<File | null>(null);
  const [modeloRGeralDragActive, setModeloRGeralDragActive] = useState(false);

  const [showFilterConfig, setShowFilterConfig] = useState(false);
  const [leituristaText, setLeituristaText] = useState(
    DEFAULT_LEITURISTAS.join('\n'),
  );

  const [isProcessing, setIsProcessing] = useState(false);
  const [progressPct, setProgressPct] = useState(0);
  const [progressLabel, setProgressLabel] = useState('');
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [screenErrorLog, setScreenErrorLog] = useState<string[]>([]);
  const [result, setResult] = useState<ProcessApiResponse | null>(null);

  const [activeTab, setActiveTab] = useState<WorkspaceTab>('preview');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedLeituristaFilter, setSelectedLeituristaFilter] = useState('ALL');
  const [currentPage, setCurrentPage] = useState(1);
  const [controleHorarioPage, setControleHorarioPage] = useState(1);

  const matrizInputRef = useRef<HTMLInputElement>(null);
  const consistenciaInputRef = useRef<HTMLInputElement>(null);
  const nosbInputRef = useRef<HTMLInputElement>(null);
  const fotosInputRef = useRef<HTMLInputElement>(null);
  const modeloRGeralInputRef = useRef<HTMLInputElement>(null);
  const pipelineSectionRef = useRef<HTMLElement>(null);
  const workspaceSectionRef = useRef<HTMLElement>(null);

  const appendScreenError = useCallback((message: string) => {
    const timestamp = new Date().toLocaleTimeString('pt-BR');
    setErrorMsg(message);
    setScreenErrorLog((prev) => [`[${timestamp}] ${message}`, ...prev].slice(0, 10));
  }, []);

  const handleReloadSession = useCallback(() => {
    window.open(window.location.href, '_blank');
  }, []);

  const [footerDateTime, setFooterDateTime] = useState(() =>
    formatFooterDateTime(new Date()),
  );

  useEffect(() => {
    setFooterDateTime(formatFooterDateTime(new Date()));
    const intervalId = window.setInterval(() => {
      setFooterDateTime(formatFooterDateTime(new Date()));
    }, 60_000);
    return () => {
      window.clearInterval(intervalId);
    };
  }, []);

  useEffect(() => {
    const onUnhandledRejection = (e: PromiseRejectionEvent) => {
      const reason = e.reason as { message?: unknown } | undefined;
      const msg = String(reason?.message || e.reason || '');
      if (msg.includes('WebSocket') || msg.includes('vite') || msg.includes('HMR')) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      console.error('[Unhandled]', e.reason);
    };
    const onWindowError = (e: ErrorEvent) => {
      const errObj = e.error as { message?: unknown } | undefined;
      const msg = String(errObj?.message || e.message || '');
      if (msg.includes('WebSocket') || msg.includes('vite') || msg.includes('HMR')) {
        e.preventDefault();
        return;
      }
      console.error('[Error]', e.error || e.message);
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
      appendScreenError(
        'Formato inválido para a planilha Matriz. Envie um arquivo .xlsx ou .xls.',
      );
      return;
    }
    setErrorMsg(null);
    setMatrizFile(file);
  }, [appendScreenError]);

  const handleAuxFileSelect = useCallback(
    (
      file: File | undefined,
      label: string,
      setter: React.Dispatch<React.SetStateAction<File | null>>,
      allowXlsm = false,
    ) => {
      if (!file) return;
      const lower = file.name.toLowerCase();
      const valid = allowXlsm
        ? lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xls')
        : lower.endsWith('.xlsx') || lower.endsWith('.xls');
      if (!valid) {
        appendScreenError(
          allowXlsm
            ? `Formato inválido para a planilha "${label}". Envie um arquivo .xlsx ou .xlsm.`
            : `Formato inválido para a planilha "${label}". Envie um arquivo .xlsx ou .xls.`,
        );
        return;
      }
      setErrorMsg(null);
      setter(file);
    },
    [appendScreenError],
  );

  const runProcessing = useCallback(async () => {
    if (!matrizFile) {
      appendScreenError(
        'Selecione a planilha Matriz (.xlsx ou .xls) antes de iniciar o processamento.',
      );
      return;
    }
    setErrorMsg(null);
    setIsProcessing(true);
    setProgressPct(10);
    setProgressLabel('Processando... aguarde');

    let timer1: number | undefined;
    let timer2: number | undefined;

    try {
      await checkBackendHealth();

      setProgressPct(22);
      setProgressLabel('Processando... aguarde (lendo e desmesclando planilha)');

      const formData = new FormData();
      formData.append('matriz', matrizFile);
      if (consistenciaFile) formData.append('consistencia', consistenciaFile);
      if (nosbFile) formData.append('nosb', nosbFile);
      if (fotosFile) formData.append('fotos', fotosFile);
      if (modeloRGeralFile) formData.append('modeloRGeral', modeloRGeralFile);
      formData.append(
        'allowedLeituristas',
        JSON.stringify(parsedAllowedLeituristas),
      );
      formData.append('populatePlanilha3FromStage2', 'true');

      timer1 = window.setTimeout(() => {
        setProgressPct(52);
        setProgressLabel('Processando... aguarde (normalização, ordenação e formatação)');
      }, 260);
      timer2 = window.setTimeout(() => {
        setProgressPct(82);
        setProgressLabel('Processando... aguarde (gerando planilha final)');
      }, 620);

      const payload = (await safeFetchJson('/api/process', {
        method: 'POST',
        body: formData,
      })) as unknown as ProcessApiResponse;

      window.clearTimeout(timer1);
      window.clearTimeout(timer2);

      setProgressPct(100);
      setProgressLabel(`Processamento concluído: ${payload.stats.finalRows} linhas`);
      setResult(payload);
      setActiveTab('preview');
      setSelectedLeituristaFilter('ALL');
      setSearchQuery('');
      setCurrentPage(1);
      setControleHorarioPage(1);
    } catch (err) {
      if (timer1 !== undefined) window.clearTimeout(timer1);
      if (timer2 !== undefined) window.clearTimeout(timer2);
      setProgressPct(0);
      setProgressLabel('');
      const msg =
        err instanceof Error
          ? err.name === 'AbortError'
            ? 'Tempo limite excedido (300s) ao processar a planilha.'
            : err.message
          : 'Erro de comunicação ao processar a planilha.';
      appendScreenError(msg);
    } finally {
      setIsProcessing(false);
    }
  }, [
    matrizFile,
    consistenciaFile,
    nosbFile,
    fotosFile,
    modeloRGeralFile,
    parsedAllowedLeituristas,
    appendScreenError,
  ]);

  const runSampleProcessing = useCallback(async () => {
    setErrorMsg(null);
    setIsProcessing(true);
    setProgressPct(15);
    setProgressLabel('Processando... aguarde');

    try {
      await checkBackendHealth();
      setProgressPct(45);
      setProgressLabel('Processando... aguarde (gerando exemplo)');

      const payload = (await safeFetchJson('/api/process-sample', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ populatePlanilha3FromStage2: true }),
      })) as unknown as ProcessApiResponse;

      setProgressPct(100);
      setProgressLabel(`Processamento concluído: ${payload.stats.finalRows} linhas`);
      setResult(payload);
      setActiveTab('preview');
      setSelectedLeituristaFilter('ALL');
      setSearchQuery('');
      setCurrentPage(1);
      setControleHorarioPage(1);
    } catch (err) {
      setProgressPct(0);
      setProgressLabel('');
      const msg =
        err instanceof Error
          ? err.message
          : 'Erro ao executar a planilha de exemplo.';
      appendScreenError(msg);
    } finally {
      setIsProcessing(false);
    }
  }, [appendScreenError]);

  const handleReset = useCallback(() => {
    setMatrizFile(null);
    setConsistenciaFile(null);
    setNosbFile(null);
    setFotosFile(null);
    setModeloRGeralFile(null);
    setResult(null);
    setErrorMsg(null);
    setScreenErrorLog([]);
    setProgressPct(0);
    setProgressLabel('');
    setSearchQuery('');
    setSelectedLeituristaFilter('ALL');
    setCurrentPage(1);
    setControleHorarioPage(1);
    if (matrizInputRef.current) matrizInputRef.current.value = '';
    if (consistenciaInputRef.current) consistenciaInputRef.current.value = '';
    if (nosbInputRef.current) nosbInputRef.current.value = '';
    if (fotosInputRef.current) fotosInputRef.current.value = '';
    if (modeloRGeralInputRef.current) modeloRGeralInputRef.current.value = '';
  }, []);

  const handleDownloadResult = useCallback(() => {
    if (!result) return;
    downloadBase64Xlsx(result.fileBase64, result.fileName);
  }, [result]);

  const handleDownloadRelatorioGeral = useCallback(() => {
    if (!result || !result.relatorioGeralBase64) return;
    downloadBase64Xlsx(
      result.relatorioGeralBase64,
      result.relatorioGeralFileName || 'Relatorio Geral.xlsx',
    );
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
        const rowMatr = (row[5] ?? '').trim();
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
      const key = (row[5] ?? '').trim() || 'Sem Leiturista';
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

  const filteredControleHorarioRows = useMemo(() => {
    if (!result?.controleHorario) return [];
    const q = searchQuery.trim().toLowerCase();
    return result.controleHorario.filter((row) => {
      if (selectedLeituristaFilter !== 'ALL') {
        if (row.matricula.trim() !== selectedLeituristaFilter) {
          return false;
        }
      }
      if (!q) return true;
      return (
        row.matricula.toLowerCase().includes(q) ||
        row.data.toLowerCase().includes(q) ||
        row.horaInicial.toLowerCase().includes(q) ||
        row.horaFinal.toLowerCase().includes(q) ||
        row.tempoServico.toLowerCase().includes(q)
      );
    });
  }, [result, searchQuery, selectedLeituristaFilter]);

  const controleHorarioTotalPages = Math.max(
    1,
    Math.ceil(filteredControleHorarioRows.length / PAGE_SIZE),
  );
  const safeControleHorarioPage = Math.min(
    controleHorarioPage,
    controleHorarioTotalPages,
  );

  const paginatedControleHorarioRows = useMemo(() => {
    const startIdx = (safeControleHorarioPage - 1) * PAGE_SIZE;
    return filteredControleHorarioRows.slice(startIdx, startIdx + PAGE_SIZE);
  }, [filteredControleHorarioRows, safeControleHorarioPage]);

  const controleHorarioPageNumbers = useMemo(
    () =>
      getVisiblePageNumbers(safeControleHorarioPage, controleHorarioTotalPages),
    [safeControleHorarioPage, controleHorarioTotalPages],
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
              onClick={() => navigateToTab('controleHorario')}
              className={`hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'controleHorario' ? 'text-teal-700 font-semibold' : ''
              }`}
            >
              Controle Horário
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
                  importe aqui, a planilha que foi exportada da Sequencia de Leitura Executada.
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

            {/* 4 Cards Auxiliares Opcionais: Consistencia, Nosb, Fotos e Importar planilha matriz */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {(
                [
                  {
                    title: 'Consistencia',
                    subtitle: 'importe aqui a planilha retirada do relatório "Consistência completa".',
                    file: consistenciaFile,
                    setFile: setConsistenciaFile,
                    dragActive: consistenciaDragActive,
                    setDragActive: setConsistenciaDragActive,
                    inputRef: consistenciaInputRef,
                    allowXlsm: false,
                    acceptAttr: '.xlsx,.xls',
                    emptyPrompt: 'Arraste ou clique (.xlsx / .xls)',
                  },
                  {
                    title: 'Nosb',
                    subtitle: 'Opcional · Preenche Filtros G:H',
                    file: nosbFile,
                    setFile: setNosbFile,
                    dragActive: nosbDragActive,
                    setDragActive: setNosbDragActive,
                    inputRef: nosbInputRef,
                    allowXlsm: false,
                    acceptAttr: '.xlsx,.xls',
                    emptyPrompt: 'Arraste ou clique (.xlsx / .xls)',
                  },
                  {
                    title: 'Fotos',
                    subtitle: 'importe aqui a planilha retirada do relatório "Relatório de Fotos Executadas".',
                    file: fotosFile,
                    setFile: setFotosFile,
                    dragActive: fotosDragActive,
                    setDragActive: setFotosDragActive,
                    inputRef: fotosInputRef,
                    allowXlsm: false,
                    acceptAttr: '.xlsx,.xls',
                    emptyPrompt: 'Arraste ou clique (.xlsx / .xls)',
                  },
                  {
                    title: 'Importar planilha matriz',
                    subtitle: 'importe aqui a planilha que receberá os dados tratados pelas planilhas anteriormente importadas.',
                    file: modeloRGeralFile,
                    setFile: setModeloRGeralFile,
                    dragActive: modeloRGeralDragActive,
                    setDragActive: setModeloRGeralDragActive,
                    inputRef: modeloRGeralInputRef,
                    allowXlsm: true,
                    acceptAttr:
                      '.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12',
                    emptyPrompt: 'Arraste ou clique (.xlsx / .xlsm)',
                  },
                ] as const
              ).map((aux) => (
                <div
                  key={aux.title}
                  className="bg-slate-50/50 border border-slate-200 rounded-xl p-4 space-y-3"
                >
                  <div className="flex items-center justify-between gap-2">
                    <div>
                      <h3 className="font-display text-sm font-semibold text-slate-900">
                        {aux.title}
                      </h3>
                      <p className="text-[11px] text-slate-500">{aux.subtitle}</p>
                    </div>
                    {aux.file && (
                      <button
                        type="button"
                        onClick={() => {
                          aux.setFile(null);
                          if (aux.inputRef.current) aux.inputRef.current.value = '';
                        }}
                        className="inline-flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-900 px-1.5 py-0.5 rounded cursor-pointer"
                        aria-label={`Remover arquivo ${aux.title}`}
                      >
                        <X className="w-3.5 h-3.5" />
                        <span>Remover</span>
                      </button>
                    )}
                  </div>

                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      aux.setDragActive(true);
                    }}
                    onDragLeave={() => aux.setDragActive(false)}
                    onDrop={(e) => {
                      e.preventDefault();
                      aux.setDragActive(false);
                      handleAuxFileSelect(
                        e.dataTransfer.files?.[0],
                        aux.title,
                        aux.setFile,
                        aux.allowXlsm,
                      );
                    }}
                    onClick={() => aux.inputRef.current?.click()}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        aux.inputRef.current?.click();
                      }
                    }}
                    className={`border border-dashed rounded-lg px-3 py-5 text-center cursor-pointer transition-colors focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700 ${
                      aux.dragActive
                        ? 'border-teal-600 bg-teal-50/50'
                        : aux.file
                          ? 'border-teal-300 bg-white'
                          : 'border-slate-300 bg-white hover:border-teal-600 hover:bg-slate-50/70'
                    }`}
                  >
                    <input
                      ref={aux.inputRef}
                      type="file"
                      accept={aux.acceptAttr}
                      className="hidden"
                      onChange={(e) =>
                        handleAuxFileSelect(
                          e.target.files?.[0],
                          aux.title,
                          aux.setFile,
                          aux.allowXlsm,
                        )
                      }
                    />
                    {aux.file ? (
                      <div className="flex flex-col items-center gap-1">
                        <FileSpreadsheet className="w-5 h-5 text-teal-700" />
                        <p className="text-xs font-semibold text-slate-900 break-all line-clamp-1">
                          {aux.file.name}
                        </p>
                        <p className="text-[11px] text-slate-500 font-mono tabular-nums">
                          {formatFileSize(aux.file.size)}
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center gap-1.5">
                        <Upload className="w-5 h-5 text-slate-400" />
                        <p className="text-xs font-medium text-slate-700">
                          {aux.emptyPrompt}
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              ))}
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
                  <>
                    <button
                      type="button"
                      onClick={handleDownloadResult}
                      className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
                    >
                      <Download className="w-4 h-4 shrink-0" />
                      <span>Baixar Matriz_ControleDeSequenciaLeitura.xlsx</span>
                    </button>
                    {result.relatorioGeralBase64 && (
                      <button
                        type="button"
                        onClick={handleDownloadRelatorioGeral}
                        className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-slate-800 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-700"
                      >
                        <Download className="w-4 h-4 shrink-0" />
                        <span>Baixar {result.relatorioGeralFileName || 'Relatorio Geral.xlsx'}</span>
                      </button>
                    )}
                  </>
                )}

                <button
                  type="button"
                  onClick={runSampleProcessing}
                  disabled={isProcessing}
                  className="inline-flex items-center gap-2 px-4 py-2.5 text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <span>Usar Exemplo</span>
                </button>

                {(matrizFile || consistenciaFile || nosbFile || fotosFile || modeloRGeralFile || result) && (
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

              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  onClick={handleReloadSession}
                  className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 hover:text-slate-900 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RotateCcw className="w-3.5 h-3.5 shrink-0" />
                  <span>Recarregar sessão</span>
                </button>
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
                className="flex flex-col sm:flex-row sm:items-start justify-between gap-3 p-4 rounded-lg bg-red-50 border border-red-200 text-red-900 text-sm"
              >
                <div className="flex items-start gap-3">
                  <AlertTriangle className="w-5 h-5 text-red-700 shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <p className="font-semibold">Erro no processamento</p>
                    <p className="text-red-800">{errorMsg}</p>
                    {screenErrorLog.length > 0 && (
                      <div className="pt-1 font-mono text-xs text-red-700 space-y-0.5">
                        {screenErrorLog.slice(0, 3).map((entry, idx) => (
                          <p key={idx}>{entry}</p>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={handleReloadSession}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-red-900 bg-white hover:bg-red-100 border border-red-300 rounded-md whitespace-nowrap shrink-0 cursor-pointer self-start"
                >
                  <RotateCcw className="w-3.5 h-3.5 shrink-0" />
                  <span>Recarregar sessão</span>
                </button>
              </div>
            )}
          </div>
        </section>

        {/* Faixa de KPIs */}
        <section className="grid grid-cols-2 lg:grid-cols-4 bg-white border border-slate-200 rounded-xl divide-y lg:divide-y-0 lg:divide-x divide-slate-200">
          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">
              Quantidade Instalações
            </p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result
                ? result.stats.uniqueInstalacoesCount.toLocaleString('pt-BR')
                : '—'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">Rural</p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result ? result.stats.ruralCount.toLocaleString('pt-BR') : '—'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">Povoado</p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-slate-900">
              {result ? result.stats.povoadoCount.toLocaleString('pt-BR') : '—'}
            </p>
          </div>

          <div className="p-5 space-y-1">
            <p className="text-xs font-medium text-slate-500">Urbano</p>
            <p className="font-mono tabular-nums text-2xl font-semibold text-teal-700">
              {result ? result.stats.urbanoCount.toLocaleString('pt-BR') : '—'}
            </p>
          </div>
        </section>

        {/* Área de Pré-visualização Paginada e Controle Horário */}
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
                onClick={() => setActiveTab('controleHorario')}
                className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  activeTab === 'controleHorario'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Controle Horário
              </button>
            </div>

            {result && (
              <div className="flex flex-wrap items-center gap-3">
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      setCurrentPage(1);
                      setControleHorarioPage(1);
                    }}
                    placeholder={
                      activeTab === 'controleHorario'
                        ? 'Buscar matrícula, data...'
                        : 'Buscar instalação, medidor, endereço...'
                    }
                    className="pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg w-64 text-slate-800 focus:outline-none focus:border-teal-600"
                  />
                </div>
                <select
                  value={selectedLeituristaFilter}
                  onChange={(e) => {
                    setSelectedLeituristaFilter(e.target.value);
                    setCurrentPage(1);
                    setControleHorarioPage(1);
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
                                idx >= 17 && idx <= 19
                                  ? 'bg-teal-800'
                                  : idx >= 20
                                    ? 'bg-slate-800'
                                    : ''
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
                              (row[5] ?? '').trim() || 'Sem Leiturista';
                            const prevLeiturista =
                              idxInPage > 0
                                ? (paginatedRows[idxInPage - 1]?.[5] ?? '').trim() ||
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
                                      cIdx === 11 || cIdx === 12 || cIdx === 13;
                                    const isMonoField =
                                      isNumericReading ||
                                      cIdx === 2 ||
                                      cIdx === 3 ||
                                      cIdx === 4 ||
                                      cIdx === 5 ||
                                      cIdx === 7 ||
                                      cIdx === 8 ||
                                      cIdx === 9 ||
                                      cIdx === 15;
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
                                          cIdx === 13
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

          {activeTab === 'controleHorario' && (
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
                      Processe uma planilha para visualizar o Controle Horário por matrícula e data.
                    </p>
                  </div>
                </div>
              ) : (
                <div>
                  <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-4 text-xs text-slate-700">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-teal-700 shrink-0" />
                      <span className="font-medium">
                        Controle Horário agrupado por Matrícula e Data.
                      </span>
                    </div>
                    <span className="font-mono tabular-nums text-slate-500">
                      Página {safeControleHorarioPage} de {controleHorarioTotalPages}
                    </span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left">
                      <thead>
                        <tr className="bg-slate-900 text-white text-xs">
                          <th className="px-3 py-2.5 font-mono font-medium text-slate-300 border-r border-slate-800 whitespace-nowrap w-14">
                            #
                          </th>
                          <th className="px-4 py-2.5 font-medium whitespace-nowrap border-r border-slate-800">
                            Matricula
                          </th>
                          <th className="px-4 py-2.5 font-medium whitespace-nowrap border-r border-slate-800">
                            Data
                          </th>
                          <th className="px-4 py-2.5 font-medium whitespace-nowrap border-r border-slate-800">
                            Hora inicial
                          </th>
                          <th className="px-4 py-2.5 font-medium whitespace-nowrap border-r border-slate-800">
                            Hora Final
                          </th>
                          <th className="px-4 py-2.5 font-medium whitespace-nowrap bg-teal-800">
                            Tempo de Serviço
                          </th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-xs">
                        {paginatedControleHorarioRows.length === 0 ? (
                          <tr>
                            <td
                              colSpan={6}
                              className="px-6 py-10 text-center text-slate-500"
                            >
                              Nenhum registro de horário corresponde ao filtro atual.
                            </td>
                          </tr>
                        ) : (
                          paginatedControleHorarioRows.map((row, idxInPage) => {
                            const globalRowIndex =
                              (safeControleHorarioPage - 1) * PAGE_SIZE + idxInPage;
                            return (
                              <tr
                                key={`${row.matricula}-${row.data}-${globalRowIndex}`}
                                className="hover:bg-slate-50 transition-colors h-10"
                              >
                                <td className="px-3 py-2 font-mono tabular-nums text-slate-400 border-r border-slate-100 whitespace-nowrap">
                                  {globalRowIndex + 1}
                                </td>
                                <td className="px-4 py-2 font-mono tabular-nums font-medium text-slate-900 whitespace-nowrap border-r border-slate-100">
                                  {row.matricula || '—'}
                                </td>
                                <td className="px-4 py-2 font-mono tabular-nums text-slate-700 whitespace-nowrap border-r border-slate-100">
                                  {row.data || '—'}
                                </td>
                                <td className="px-4 py-2 font-mono tabular-nums text-slate-700 whitespace-nowrap border-r border-slate-100">
                                  {row.horaInicial || '—'}
                                </td>
                                <td className="px-4 py-2 font-mono tabular-nums text-slate-700 whitespace-nowrap border-r border-slate-100">
                                  {row.horaFinal || '—'}
                                </td>
                                <td className="px-4 py-2 font-mono tabular-nums font-semibold text-teal-900 bg-teal-50/40 whitespace-nowrap">
                                  {row.tempoServico || '00:00'}
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* Controles de Paginação Controle Horário (30 registros por página) */}
                  <div className="px-6 py-3.5 border-t border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-4 text-xs text-slate-600">
                    <div>
                      Exibindo{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {filteredControleHorarioRows.length === 0
                          ? 0
                          : (safeControleHorarioPage - 1) * PAGE_SIZE + 1}
                      </strong>
                      –
                      <strong className="font-mono tabular-nums text-slate-900">
                        {Math.min(
                          safeControleHorarioPage * PAGE_SIZE,
                          filteredControleHorarioRows.length,
                        )}
                      </strong>{' '}
                      de{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {filteredControleHorarioRows.length.toLocaleString('pt-BR')}
                      </strong>{' '}
                      registros · Página{' '}
                      <strong className="font-mono tabular-nums text-slate-900">
                        {safeControleHorarioPage} de {controleHorarioTotalPages}
                      </strong>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() =>
                          setControleHorarioPage((p) => Math.max(1, p - 1))
                        }
                        disabled={safeControleHorarioPage <= 1}
                        className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium bg-white border border-slate-200 rounded-md text-slate-700 hover:bg-slate-100 disabled:opacity-45 disabled:cursor-not-allowed cursor-pointer whitespace-nowrap"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                        <span>Anterior</span>
                      </button>

                      {controleHorarioPageNumbers.map((item, idx) =>
                        item === 'ellipsis' ? (
                          <span
                            key={`che-${idx}`}
                            className="px-2 py-1 text-slate-400 font-mono"
                          >
                            …
                          </span>
                        ) : (
                          <button
                            key={item}
                            type="button"
                            onClick={() => setControleHorarioPage(item)}
                            className={`px-2.5 py-1.5 text-xs font-mono tabular-nums rounded-md border transition-colors cursor-pointer ${
                              item === safeControleHorarioPage
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
                          setControleHorarioPage((p) =>
                            Math.min(controleHorarioTotalPages, p + 1),
                          )
                        }
                        disabled={safeControleHorarioPage >= controleHorarioTotalPages}
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
        </section>
      </main>

      <footer className="mt-8 border-t border-slate-200 bg-white/70 px-6 py-5 text-center text-xs text-slate-500">
        <p className="max-w-[1380px] mx-auto leading-relaxed">
          Controle de Sequenciamento de Leitura • {footerDateTime} • Criado por: Thiago Marques Lopes
        </p>
      </footer>
    </div>
  );
}
