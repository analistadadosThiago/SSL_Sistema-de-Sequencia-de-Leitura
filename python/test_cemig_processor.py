"""
Script de teste e validação do processador Python CEMIG (17 colunas, dd/mm/aaaa, ordenação).
"""

from __future__ import annotations

import argparse
import io
import sys
import unittest

import openpyxl

from cemig_processor import CemigValidationError, FINAL_HEADERS, process_matriz_bytes


def build_synthetic_cemig_matriz() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leituras"

    def set_text(r: int, c: int, val: str) -> None:
        cell = ws.cell(r, c, val)
        cell.number_format = "@"

    set_text(1, 1, "CEMIG Distribuição S.A.")
    set_text(1, 10, "Pág. 1/446")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    set_text(2, 1, "Seqüência de Leitura Executada")
    set_text(3, 1, "Contrato: 4600018294")
    set_text(4, 1, "Unidade de Leitura Inicial: 31020001")
    set_text(5, 1, "Tipo de Leitura: CONVENCIONAL")
    set_text(6, 1, "Data Início: 11/6/2026")
    set_text(7, 1, "Data Término: 11/6/2026")
    set_text(8, 1, "Regional: METROPOLITANA")
    set_text(9, 1, "Lote: 001")
    set_text(10, 1, "--------------------------------------------------------------------------------")

    h = 11
    set_text(h, 1, "Consumidora")
    set_text(h, 2, "Medidor")
    set_text(h, 3, "T. Reg")
    set_text(h, 4, "Seqüência de Leitura Executada")
    set_text(h, 5, "Contrato_Aux")
    set_text(h, 6, "Data")
    set_text(h, 7, "Hora")
    set_text(h + 2, 7, "Leit. Ant.")
    set_text(h, 8, "Tipo_Aux1")
    set_text(h, 9, "Tipo_Aux2")
    set_text(h, 10, "Leit. Exec")
    set_text(h, 11, "")
    set_text(h, 12, "Qtd Dig.")
    set_text(h, 13, "OC")
    set_text(h + 2, 13, "Apres.")

    r = 18
    set_text(r - 5, 1, "Unidade de Leitura: 31020001")
    set_text(r - 4, 1, "Leiturista: 4MW")
    set_text(r, 1, "MARIA CELESTE M MORAIS")
    set_text(r, 2, "ARL211115451")
    set_text(r, 3, "00060")
    set_text(r, 4, "00136")
    set_text(r, 5, "4600018294")
    set_text(r, 6, "11/6/2026")
    set_text(r, 7, "10:03:55 AM")
    set_text(r + 2, 7, "23873")
    set_text(r, 8, "CONV")
    set_text(r, 9, "NORMAL")
    set_text(r, 10, "24182")
    set_text(r, 11, "")
    set_text(r, 12, "")
    set_text(r, 13, "00")
    set_text(r + 2, 13, "RUA DAS FLORES, 142 - CENTRO")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestCemigProcessor(unittest.TestCase):
    def test_pipeline_17_columns_and_ddmmyyyy(self) -> None:
        raw_bytes = build_synthetic_cemig_matriz()
        out_bytes, report = process_matriz_bytes(raw_bytes, "Leituras.xlsx")

        self.assertEqual(report.final_rows, 1)
        wb = openpyxl.load_workbook(io.BytesIO(out_bytes), data_only=False)
        ws = wb["ControleLeitura"]
        self.assertIn("Tabela1", ws.tables)
        headers = [ws.cell(1, c).value for c in range(1, 18)]
        self.assertEqual(headers, FINAL_HEADERS)

        self.assertEqual(ws["A2"].value, "MARIA CELESTE M MORAIS")
        self.assertEqual(ws["B2"].value, "ARL211115451")
        self.assertEqual(ws["C2"].value, "00060")
        self.assertEqual(ws["D2"].value, "31020001")
        self.assertEqual(ws["E2"].value, "=VLOOKUP(D2,Planilha3!$A$1:$B$50000,2,0)")
        self.assertEqual(ws["F2"].value, "00")
        self.assertEqual(ws["G2"].value, "06/11/2026")
        self.assertEqual(ws["H2"].value, "10:03:55 AM")
        self.assertEqual(ws["I2"].value, "RUA DAS FLORES, 142 - CENTRO")
        self.assertEqual(ws["J2"].value, 23873)
        self.assertEqual(ws["K2"].value, 24182)
        self.assertEqual(ws["L2"].value, "=SUM(K2-J2)")
        self.assertEqual(ws["M2"].value, "Faturada")
        self.assertEqual(ws["N2"].value, "N00")

    def test_invalid_layout_raises_validation_error(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Planilha Aleatória"
        buf = io.BytesIO()
        wb.save(buf)
        with self.assertRaises(CemigValidationError):
            process_matriz_bytes(buf.getvalue(), "invalida.xlsx")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--matriz")
    parser.add_argument("--ref")
    args, unknown = parser.parse_known_args()
    sys.argv = [sys.argv[0], *unknown]
    unittest.main()
