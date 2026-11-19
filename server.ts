import express, {
  type NextFunction,
  type Request,
  type Response,
} from 'express';
import http from 'http';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import {
  CemigValidationError,
  DEFAULT_ALLOWED_LEITURISTAS,
  processCemigWorkbook,
} from './server/cemigProcessor.ts';
import { generateSampleCemigMatrizBuffer } from './server/sampleCemigGenerator.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);

// Proteção global contra encerramento inesperado do processo Node
process.on('uncaughtException', (err) => {
  console.error('[UNCAUGHT]', err);
});

process.on('unhandledRejection', (err) => {
  console.error('[UNHANDLED]', err);
});

// Processamento 100% em memória (limite 50 MB)
const ALLOWED_UPLOAD_MIMES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel.sheet.macroenabled.12',
  'application/vnd.ms-excel',
  'application/octet-stream',
]);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
  fileFilter: (_req, file, cb) => {
    const lower = (file.originalname || '').toLowerCase();
    const mimeLower = (file.mimetype || '').toLowerCase();
    const extAllowed =
      file.fieldname === 'modeloRGeral'
        ? lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xls')
        : lower.endsWith('.xlsx') || lower.endsWith('.xls');
    if (extAllowed || ALLOWED_UPLOAD_MIMES.has(mimeLower)) {
      cb(null, true);
      return;
    }
    cb(
      new Error(
        `Formato não suportado para "${file.originalname}". Envie uma planilha .xlsx, .xlsm ou .xls.`,
      ),
    );
  },
});

async function startServer() {
  const app = express();
  const httpServer = http.createServer(app);
  httpServer.requestTimeout = 300000;
  httpServer.headersTimeout = 305000;
  httpServer.keepAliveTimeout = 65000;

  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ extended: true, limit: '50mb' }));

  app.get('/health', (_req, res) => {
    res.json({ ok: true });
  });

  app.get('/api/health', (_req, res) => {
    res.json({ ok: true, status: 'ok', defaultLeituristas: DEFAULT_ALLOWED_LEITURISTAS });
  });

  // Download da planilha de exemplo no layout CEMIG
  app.get('/api/sample-matriz', async (_req, res) => {
    try {
      const buffer = await generateSampleCemigMatrizBuffer();
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
      res.setHeader(
        'Content-Disposition',
        'attachment; filename="Matriz_Leituras_Exemplo.xlsx"',
      );
      res.send(buffer);
    } catch (err) {
      console.error('Erro em /api/sample-matriz:', err);
      if (!res.headersSent) {
        res.status(500).json({
          erro: 'Falha ao processar',
          detalhe: String(err instanceof Error ? err.message : err),
          stack: String(err instanceof Error ? err.stack ?? '' : ''),
        });
      }
    }
  });

  // Endpoint principal: recebe Matriz (obrigatória) + Planilha3 (opcional) e retorna JSON
  app.post(
    '/api/process',
    (req, res, next) => {
      upload.fields([
        { name: 'matriz', maxCount: 1 },
        { name: 'planilha3', maxCount: 1 },
        { name: 'consistencia', maxCount: 1 },
        { name: 'nosb', maxCount: 1 },
        { name: 'fotos', maxCount: 1 },
        { name: 'modeloRGeral', maxCount: 1 },
      ])(req, res, (uploadErr) => {
        if (uploadErr) {
          console.error('Erro de upload multer:', uploadErr);
          if (!res.headersSent) {
            res.status(400).json({
              erro: 'Falha no recebimento do arquivo.',
              detalhe: String(uploadErr),
            });
          }
          return;
        }
        next();
      });
    },
    async (req, res) => {
      console.log('[PROCESSAR] Início', new Date().toISOString());
      try {
        const files = req.files as
          | Record<string, Express.Multer.File[]>
          | undefined;
        const matrizFile = files?.matriz?.[0];
        const planilha3File = files?.planilha3?.[0];
        const consistenciaFile = files?.consistencia?.[0];
        const nosbFile = files?.nosb?.[0];
        const fotosFile = files?.fotos?.[0];
        const modeloRGeralFile = files?.modeloRGeral?.[0];

        if (!matrizFile) {
          res.status(400).json({
            erro: 'Nenhum arquivo enviado',
            detalhe: 'Selecione uma planilha .xlsx ou .xls antes de processar.',
          });
          return;
        }

        const lowerName = (matrizFile.originalname || '').toLowerCase();
        if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
          res.status(400).json({
            erro: 'Formato inválido',
            detalhe: 'Envie uma planilha com extensão .xlsx ou .xls.',
          });
          return;
        }

        for (const [label, f] of [
          ['Consistencia', consistenciaFile],
          ['Nosb', nosbFile],
          ['Fotos', fotosFile],
        ] as const) {
          if (f) {
            const auxLower = (f.originalname || '').toLowerCase();
            if (!auxLower.endsWith('.xlsx') && !auxLower.endsWith('.xls')) {
              res.status(400).json({
                erro: 'Formato inválido',
                detalhe: `A planilha "${label}" deve ter extensão .xlsx ou .xls.`,
              });
              return;
            }
          }
        }

        if (modeloRGeralFile) {
          const modelLower = (modeloRGeralFile.originalname || '').toLowerCase();
          const modelMime = (modeloRGeralFile.mimetype || '').toLowerCase();
          const validExt =
            modelLower.endsWith('.xlsx') ||
            modelLower.endsWith('.xlsm') ||
            modelLower.endsWith('.xls');
          const validMime =
            modelMime ===
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            modelMime === 'application/vnd.ms-excel.sheet.macroenabled.12';
          if (!validExt && !validMime) {
            res.status(400).json({
              erro: 'Formato inválido',
              detalhe:
                'A planilha "Planilha Modelo RGeral" deve ter extensão .xlsx ou .xlsm.',
            });
            return;
          }
        }

        if (!matrizFile.buffer || matrizFile.size === 0 || matrizFile.buffer.length === 0) {
          res.status(400).json({
            erro: 'Arquivo vazio',
            detalhe: 'A planilha enviada não contém dados (0 bytes).',
          });
          return;
        }

        let allowedLeituristas: string[] | undefined;
        if (
          typeof req.body?.allowedLeituristas === 'string' &&
          req.body.allowedLeituristas.trim()
        ) {
          try {
            const parsed = JSON.parse(req.body.allowedLeituristas);
            if (Array.isArray(parsed)) {
              allowedLeituristas = parsed.map((v) => String(v));
            }
          } catch {
            // Mantém lista padrão
          }
        }

        const populatePlanilha3FromStage2 =
          req.body?.populatePlanilha3FromStage2 !== 'false';

        const result = await processCemigWorkbook(
          matrizFile.buffer,
          planilha3File?.buffer ?? null,
          {
            allowedLeituristas,
            populatePlanilha3FromStage2,
            consistenciaBuffer: consistenciaFile?.buffer ?? null,
            nosbBuffer: nosbFile?.buffer ?? null,
            fotosBuffer: fotosFile?.buffer ?? null,
            modeloRGeralBuffer: modeloRGeralFile?.buffer ?? null,
            modeloRGeralFileName: modeloRGeralFile?.originalname ?? null,
          },
        );

        for (const entry of result.logs) {
          console.log(`[ETAPA] ${entry.stage} — ${entry.step}: ${entry.detail}`);
        }

        if (!res.headersSent) {
          res.json({
            fileName: result.fileName,
            fileBase64: result.outputBuffer.toString('base64'),
            relatorioGeralFileName: result.relatorioGeralFileName,
            relatorioGeralBase64: result.relatorioGeralBuffer
              ? result.relatorioGeralBuffer.toString('base64')
              : null,
            stats: result.stats,
            logs: result.logs,
            warnings: result.warnings,
            preview: result.preview,
            controleHorario: result.controleHorario,
            stage2Sample: result.stage2Sample,
          });
        }
      } catch (err) {
        console.error('Erro em /api/process:', err);
        if (res.headersSent) return;
        if (err instanceof CemigValidationError) {
          res.status(400).json({
            erro: err.code === 'CORRUPT_FILE' ? 'Arquivo corrompido' : err.message,
            detalhe: err.message,
          });
          return;
        }
        res.status(500).json({
          erro: 'Falha ao processar',
          detalhe: String(err instanceof Error ? err.message : err),
          stack: String(err instanceof Error ? err.stack ?? '' : ''),
        });
      } finally {
        console.log('[PROCESSAR] Fim', new Date().toISOString());
      }
    },
  );

  // Processa a planilha de exemplo diretamente no backend
  app.post('/api/process-sample', async (req, res) => {
    console.log('[PROCESSAR] Início', new Date().toISOString());
    try {
      const sampleBuf = await generateSampleCemigMatrizBuffer();
      const populatePlanilha3FromStage2 =
        req.body?.populatePlanilha3FromStage2 !== false;

      const result = await processCemigWorkbook(sampleBuf, null, {
        populatePlanilha3FromStage2,
      });

      for (const entry of result.logs) {
        console.log(`[ETAPA] ${entry.stage} — ${entry.step}: ${entry.detail}`);
      }

      if (!res.headersSent) {
        res.json({
          fileName: result.fileName,
          fileBase64: result.outputBuffer.toString('base64'),
          relatorioGeralFileName: result.relatorioGeralFileName,
          relatorioGeralBase64: result.relatorioGeralBuffer
            ? result.relatorioGeralBuffer.toString('base64')
            : null,
          stats: result.stats,
          logs: result.logs,
          warnings: result.warnings,
          preview: result.preview,
          controleHorario: result.controleHorario,
          stage2Sample: result.stage2Sample,
          sourceName: 'Matriz_Leituras_Exemplo.xlsx',
        });
      }
    } catch (err) {
      console.error('Erro em /api/process-sample:', err);
      if (!res.headersSent) {
        res.status(500).json({
          erro: 'Falha ao processar',
          detalhe: String(err instanceof Error ? err.message : err),
          stack: String(err instanceof Error ? err.stack ?? '' : ''),
        });
      }
    } finally {
      console.log('[PROCESSAR] Fim', new Date().toISOString());
    }
  });

  // Garantir que rotas /api/* nunca caiam em HTML de fallback
  app.use('/api', (_req, res) => {
    res.status(404).json({
      erro: 'Endpoint da API não encontrado.',
      detalhe: 'Rota inexistente.',
    });
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: false,
      },
      appType: 'spa',
    });

    // Impede que @vite/client tente abrir WebSocket na porta 24678 quando HMR está desabilitado
    app.get('/@vite/client', (_req, res) => {
      res.type('application/javascript');
      res.send(`
const sheetsMap = new Map();
export function updateStyle(id, content) {
  let style = sheetsMap.get(id);
  if (!style) {
    style = document.createElement('style');
    style.setAttribute('type', 'text/css');
    style.setAttribute('data-vite-dev-id', id);
    style.textContent = content;
    document.head.appendChild(style);
    sheetsMap.set(id, style);
  } else if (style.textContent !== content) {
    style.textContent = content;
  }
}
export function removeStyle(id) {
  const style = sheetsMap.get(id);
  if (style && style.parentNode) {
    style.parentNode.removeChild(style);
  }
  sheetsMap.delete(id);
}
export function createHotContext() {
  return {
    accept() {},
    dispose() {},
    prune() {},
    invalidate() {},
    on() {},
    off() {},
    send() {},
    data: {},
  };
}
export function injectQuery(url, queryToInject) {
  if (url[0] !== '.' && url[0] !== '/') return url;
  const pathname = url.replace(/[?#].*$/, '');
  const parsed = new URL(url, 'http://vite.dev');
  return pathname + '?' + queryToInject + (parsed.search ? '&' + parsed.search.slice(1) : '') + (parsed.hash || '');
}
export class ErrorOverlay {}
`);
    });

    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('/{*splat}', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  // Middleware final para rotas não encontradas — sempre JSON
  app.use((_req, res) => {
    res.status(404).json({ erro: 'Rota não encontrada' });
  });

  // Middleware global de erro Express — sempre retorna JSON válido
  app.use((err: unknown, _req: Request, res: Response, _next: NextFunction) => {
    console.error('[ERRO BACKEND]', err);
    if (!res.headersSent) {
      res.status(500).json({
        erro: 'Erro interno',
        detalhe: String(err instanceof Error ? err.message : err),
      });
    }
  });

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
});
