# Mapeamento das 3 Macros VBA → Implementação Python & TypeScript

Este documento explica cada transformação das 3 macros VBA originais do relatório CEMIG ("Seqüência de Leitura Executada") e como elas foram replicadas de forma determinística e dinâmica (sem dependência do Excel instalado e sem hardcode de linhas).

---

## Pré-processamento & Robustez

1. **Leitura como texto (`dtype=str` / `cellText: true`)**: Zeros à esquerda (ex.: `00060`, `00136`) e strings originais de data/hora (`11/6/2026`, `10:03:55 AM`) são lidos e preservados sem coerção numérica ou conversão de fuso horário.
2. **Desmesclagem prévia (`UnMerge`)**: Todas as células mescladas do relatório bruto são desmescladas antes da manipulação de colunas, mantendo o valor no canto superior esquerdo e limpando as células restantes.
3. **Detecção dinâmica (`max_row` / `max_column`)**: Os intervalos fixos das macros gravadas (`A1:P25929`, `A11:P23647`, `D12:E24561`, `A1:M9132`) são calculados dinamicamente com base no tamanho real da planilha enviada.
4. **Validação de layout**: Verifica dimensões mínimas e marcadores CEMIG (`CEMIG`, `Seqüência de Leitura`, `Unidade de Leitura`, `Leiturista:`, `Medidor`, `Leit. Ant.` / `Leit. Exec`).

---

## Etapa 1 — `ParteUm_Tratamento` (`Ctrl+Shift+F`)

Objetivo: Normalizar o relatório multi-linha CEMIG para formato tabular.

| Passo VBA | Operação | Implementação (`cemigProcessor.ts` / `cemig_processor.py`) |
|---|---|---|
| (1) | Excluir coluna E (`Shift:=xlToLeft`) | `_delete_cols(grid, 5, 1)` |
| (2) | Excluir colunas G:H (`Shift:=xlToLeft`) | `_delete_cols(grid, 7, 2)` |
| (3) | Copiar coluna J → coluna K | `_copy_col(grid, 10, 11)` |
| (4) | Excluir `J3:J4` (`Shift:=xlUp` x2) | `_delete_cells_shift_up(grid, 10, 3, 2)` — alinha campos da sublinha +2 com a linha principal |
| (5) | Inserir coluna em F e copiar G → F | `_insert_cols(grid, 6, 1)` + `_copy_col(grid, 7, 6)` |
| (6) | Excluir `G5:G6` (`Shift:=xlUp` x2) | `_delete_cells_shift_up(grid, 7, 5, 2)` — alinha leitura anterior da sublinha +2 |
| (7) | Inserir 2 colunas em D e copiar A → D e A → E | `_insert_cols(grid, 4, 2)` + `_copy_col(grid, 1, 4)` + `_copy_col(grid, 1, 5)` |
| (8) | `Range("D1").Insert Shift:=xlDown` x5 e `Range("E2").Insert Shift:=xlDown` x4 | `_insert_cells_shift_down(grid, 4, 1, 5)` e `_insert_cells_shift_down(grid, 5, 2, 4)` — alinha `Unidade de Leitura:` (col D) e `Leiturista:` (col E) na mesma linha do medidor |
| (9) | Remover duplicatas em 16 colunas `A:P` (`Header:=xlNo`) | `df.drop_duplicates(keep="first")` / `Set` por tupla das colunas `1..max(16, maxCol)` |
| (10) | AutoFilter na linha 11, Field 5 (`Col E`), mantendo os 15 leituristas | Filtro exato (com `.strip()`) para `4MW, 4NI, 51O, 56A, 58A, 5GL, 5HL, 5IL, 5JL, 5LL, 5UL, 5ZP, 80J, 80L, 89O` e remoção de ruído de quebra de página (`Pág. X/Y`, `CEMIG Distribuição`) |
| (11) | Selecionar `D12:E<fim>` e copiar | Extração dos pares `(col D, col E)` das linhas filtradas |

*Nota*: Eventos `ActiveWindow.ScrollRow` gravados pelo gravador de macros são ignorados por não afetarem dados ou formatação.

---

## Etapa 2 — `ParteDois_UlMatr` (`Ctrl+Shift+T`)

Objetivo: Extrair a tabela auxiliar `Unidade de Leitura ↔ Matrícula do Leiturista`.

| Passo VBA | Operação | Implementação |
|---|---|---|
| (1) | Inserir coluna em B (empurra `Leiturista:` para C) | `_insert_cols(stage2, 2, 1)` |
| (2) | `TextToColumns` largura fixa em A (quebras `0` e `19`), excluir A | Fatia `raw_d[19:].strip()`, removendo o prefixo `"Unidade de Leitura:"` |
| (3) | `TextToColumns` largura fixa em B (quebras `0` e `11`), excluir B | Fatia `raw_e[11:].strip()`, removendo o prefixo `"Leiturista:"` |
| (4) | AutoFit na coluna A | Ajuste automático da largura da coluna A na aba `Planilha3` |

---

## Etapa 3 — `ParteTrês_RelatórioFinal` (`Ctrl+Shift+V`)

Objetivo: Gerar o relatório final `Matriz_ControleDeSequenciaLeitura.xlsx` com tabela `Tabela1`, fórmulas e `Planilha3`.

| Passo VBA | Operação | Implementação |
|---|---|---|
| (1) | Excluir linhas `1:10` (`Shift:=xlUp`) | A linha 11 (cabeçalho) torna-se a linha 1, seguida dos registros filtrados |
| (2) | Criar `ListObject` `"Tabela1"` (`xlSrcRange`, `xlYes`) | `Table(displayName="Tabela1", ref="A1:P<n>")` no `openpyxl` / `ws.addTable(...)` no `ExcelJS` |
| (3) | Alinhamento `xlLeft` / `xlBottom`, sem wrap, sem merge | Aplicado em todas as células de dados (`wrapText=False`) |
| (4) | AutoFit de todas as colunas | Cálculo da largura máxima por coluna (`min(max(len + 3, 10), 46)`) |
| (5)–(6) | Inserir col E; `TextToColumns` em D (quebra `0` e `20`); excluir D | Remove prefixo `"Unidade de Leitura: "` (20 chars) mantendo código limpo na coluna D |
| (7) | AutoFit na coluna D | Recálculo automático de largura |
| (8) | Inserir 2 colunas em E (`E` e `F`) | Desloca colunas originais `E..M` para `G..O` |
| (9) | Fórmula em F2: `=VLOOKUP(RC[-2], Planilha3!R1C1:R50000C2, 2, 0)` | Gravada em sintaxe padrão `=VLOOKUP(D2,Planilha3!$A$1:$B$50000,2,0)` |
| (10) | Inserir col M e fórmula `=SUM([@Coluna7]-[@Coluna6])` | Gravada como `=SUM(L2-K2)` (ou letras detectadas para `Leit. Exec` e `Leit. Ant.`) |
| (11) | Cabeçalhos auxiliares: `Q1="cons"`, `R1="sim"`, `S1="fotos"` | Escritos nas colunas 17, 18 e 19 |
| (12) | Preencher brancos em `[Coluna8]` (`"Faturada"`) e `[Qtd Dig.]` (`"N00"`) | Substituição automática de células vazias antes da gravação |
| (13) | Fórmulas em `Q2`, `R2`, `S2` | `=IFERROR(VLOOKUP(A2,Planilha3!$D$1:$E$50000,2,0),"Faturada")`, `...$G$1:$H$50000...`, `...$J$1:$K$50000... "N-OK"` |
| (14) | Aba final + aba `Planilha3` | Se `Planilha3` não for enviada, cria `Planilha3` com colunas `A:B` da Etapa 2 e `D:K` vazias, garantindo fallback `"Faturada"` / `"N-OK"` |
