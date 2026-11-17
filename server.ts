import express, {
  type NextFunction,
  type Request,
  type Response,
} from 'express';
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
  console.error('uncaught:', err);
});

process.on('unhandledRejection', (err) => {
  console.error('unhandled:', err);
});

// Processamento 100% em memória (limite 50 MB)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
});

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));

  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok', defaultLeituristas: DEFAULT_ALLOWED_LEITURISTAS });
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
      res.status(500).json({
        erro: 'Erro ao gerar planilha de exemplo.',
        detalhe: String(err),
      });
    }
  });

  // Endpoint principal: recebe Matriz (obrigatória) + Planilha3 (opcional) e retorna JSON
  app.post(
    '/api/process',
    (req, res, next) => {
      upload.fields([
        { name: 'matriz', maxCount: 1 },
        { name: 'planilha3', maxCount: 1 },
      ])(req, res, (uploadErr) => {
        if (uploadErr) {
          console.error('Erro de upload multer:', uploadErr);
          res.status(400).json({
            erro: 'Falha no recebimento do arquivo.',
            detalhe: String(uploadErr),
          });
          return;
        }
        next();
      });
    },
    async (req, res) => {
      try {
        const files = req.files as
          | Record<string, Express.Multer.File[]>
          | undefined;
        const matrizFile = files?.matriz?.[0];
        const planilha3File = files?.planilha3?.[0];

        if (!matrizFile) {
          res.status(400).json({
            erro: 'Arquivo "Matriz" não enviado.',
            detalhe: 'Selecione uma planilha .xlsx ou .xls antes de processar.',
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
          },
        );

        res.json({
          fileName: result.fileName,
          fileBase64: result.outputBuffer.toString('base64'),
          stats: result.stats,
          logs: result.logs,
          warnings: result.warnings,
          preview: result.preview,
          stage2Sample: result.stage2Sample,
        });
      } catch (err) {
        console.error('Erro em /api/process:', err);
        if (err instanceof CemigValidationError) {
          res.status(400).json({
            erro: err.message,
            detalhe: err.message,
          });
          return;
        }
        res.status(500).json({
          erro: 'Erro ao processar planilha Matriz.',
          detalhe: String(err),
        });
      }
    },
  );

  // Processa a planilha de exemplo diretamente no backend
  app.post('/api/process-sample', async (req, res) => {
    try {
      const sampleBuf = await generateSampleCemigMatrizBuffer();
      const populatePlanilha3FromStage2 =
        req.body?.populatePlanilha3FromStage2 !== false;

      const result = await processCemigWorkbook(sampleBuf, null, {
        populatePlanilha3FromStage2,
      });

      res.json({
        fileName: result.fileName,
        fileBase64: result.outputBuffer.toString('base64'),
        stats: result.stats,
        logs: result.logs,
        warnings: result.warnings,
        preview: result.preview,
        stage2Sample: result.stage2Sample,
        sourceName: 'Matriz_Leituras_Exemplo.xlsx',
      });
    } catch (err) {
      console.error('Erro em /api/process-sample:', err);
      res.status(500).json({
        erro: 'Erro ao processar planilha de exemplo.',
        detalhe: String(err),
      });
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
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('/{*splat}', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  // Middleware global de erro Express — sempre retorna JSON válido em rotas de API
  app.use((err: unknown, _req: Request, res: Response, _next: NextFunction) => {
    console.error('Erro interno Express:', err);
    if (!res.headersSent) {
      res.status(500).json({
        erro: 'Erro interno',
        detalhe: String(err),
      });
    }
  });

  const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
  server.requestTimeout = 300000;
  server.headersTimeout = 305000;
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
});
