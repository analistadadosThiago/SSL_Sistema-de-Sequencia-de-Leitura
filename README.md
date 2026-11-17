# Controle de Sequência de Leitura — CEMIG

Sistema web completo para recebimento da planilha bruta **Matriz** (relatório `"Seqüência de Leitura Executada"` da CEMIG em `.xlsx` ou `.xls`) e geração automatizada do arquivo **`Matriz_ControleDeSequenciaLeitura.xlsx`**, replicando fielmente as 3 macros VBA (`ParteUm_Tratamento`, `ParteDois_UlMatr` e `ParteTrês_RelatórioFinal`) em memória e sem dependência do Microsoft Excel.

---

## Funcionalidades

- **Upload Drag-and-Drop**: Aceita a planilha principal `Matriz` (`.xlsx` / `.xls`) e, opcionalmente, a planilha auxiliar `Planilha3` (caso ausente, gera a aba `Planilha3` automaticamente com os pares `Unidade ↔ Matrícula` da Etapa 2 e fallback `"Faturada"` / `"N-OK"` via `IFERROR`).
- **Processamento 100% em Memória**: Nenhum arquivo do usuário é persistido em disco.
- **Preservação Estrita de Dados**: Mantém zeros à esquerda (`00060`, `00136`), formatos originais de data e hora (`11/6/2026`, `10:03:55 AM`), desmescla células mescladas antes do tratamento e detecta dinamicamente o intervalo máximo (`maxRow`/`maxCol`).
- **Saída Idêntica ao Modelo**: Gera `Matriz_ControleDeSequenciaLeitura.xlsx` com a aba principal contendo a tabela Excel **`Tabela1`**, fórmulas reais (`VLOOKUP`, `IFERROR`, `SUM`), preenchimento de brancos (`"Faturada"`, `"N00"`), AutoFit, alinhamento à esquerda/inferior e a aba auxiliar **`Planilha3`**.
- **Pré-visualização & Log de Auditoria**: Exibe os contadores de linhas em cada etapa (brutas, após deduplicação, após filtro de leiturista, finais), distribuição por leiturista e grade interativa com busca.

---

## Execução da Aplicação Web (Node.js + Express + React)

### Pré-requisitos
- Node.js 20+ (recomendado 22+)

### Instalação e Execução

```bash
npm install
npm run dev
```

A aplicação estará disponível em `http://localhost:3000`.

### Build e Produção

```bash
npm run build
npm start
```

### Testes Automatizados (Pipeline Web)

```bash
npx tsx tests/verify_pipeline.ts
```

---

## Execução Standalone em Python 3.10+ (`pandas` + `openpyxl`)

O diretório `python/` contém a implementação equivalente completa em Python 3.10+ utilizando `pandas`, `openpyxl` e `xlrd`:

```bash
pip install -r python/requirements.txt
pytest python/test_cemig_processor.py
```

Para comparar diretamente com um arquivo de referência:

```bash
python python/test_cemig_processor.py \
  --matriz Leituras.xlsx \
  --ref Matriz_ControleDeSequenciaLeitura.xlsx
```

---

## Estrutura do Projeto

- `server.ts` — Servidor Express (upload em memória via `multer`, endpoints `/api/process`, `/api/process-sample`, `/api/sample-matriz`).
- `server/cemigProcessor.ts` — Motor das 3 etapas VBA em TypeScript (`xlsx` + `exceljs`).
- `server/sampleCemigGenerator.ts` — Gerador de planilha CEMIG de exemplo para testes imediatos.
- `src/App.tsx` — Interface web (upload, progresso, KPIs, grade de pré-visualização, log das 3 etapas, mapeamento VBA).
- `python/cemig_processor.py` — Implementação standalone Python 3.10+ (`pandas` + `openpyxl`).
- `python/test_cemig_processor.py` — Script de testes unitários e comparação com arquivo de referência.
- `tests/verify_pipeline.ts` — Suite de verificação end-to-end do pipeline Node.js.
- `docs/TRANSFORMACOES_VBA_PYTHON.md` — Mapeamento detalhado de cada instrução VBA para o código.
