"""
Script de teste e validação do processador Python CEMIG (17 colunas, dd/mm/aaaa, ordenação).
"""

from __future__ import annotations

import argparse
import io
import sys
import unittest

import openpyxl

from cemig_processor import (
    CONTROLE_HORARIO_HEADERS,
    CemigValidationError,
    FINAL_HEADERS,
    process_matriz_bytes,
)


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
    def test_pipeline_20_columns_and_filtros(self) -> None:
        raw_bytes = build_synthetic_cemig_matriz()
        out_bytes, report = process_matriz_bytes(raw_bytes, "Leituras.xlsx")

        self.assertEqual(report.final_rows, 1)
        self.assertEqual(report.unique_instalacoes_count, 1)
        self.assertEqual(report.urbano_count, 1)
        self.assertEqual(report.povoado_count, 0)
        self.assertEqual(report.rural_count, 0)

        wb = openpyxl.load_workbook(io.BytesIO(out_bytes), data_only=False)
        ws = wb["ControleLeitura"]
        self.assertIn("Filtros", wb.sheetnames)
        ws_f = wb["Filtros"]
        self.assertEqual(ws_f["D1"].value, "Consistencia")
        self.assertEqual(ws_f["G1"].value, "Nosb")
        self.assertEqual(ws_f["J1"].value, "Fotos")

        self.assertIn("Tabela1", ws.tables)
        headers = [ws.cell(1, c).value for c in range(1, 21)]
        self.assertEqual(headers, FINAL_HEADERS)
        all_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        self.assertIn("Seqüência de Leitura Executada", all_headers)
        self.assertIn("Contrato_Aux", all_headers)

        self.assertEqual(ws["A2"].value, "MARIA CELESTE M MORAIS")
        self.assertEqual(ws["B2"].value, "ARL211115451")
        self.assertEqual(ws["C2"].value, "00060")
        self.assertEqual(ws["D2"].value, "31020001")
        self.assertEqual(ws["E2"].value, "0200")
        self.assertEqual(ws["F2"].value, "=VLOOKUP(D2,Filtros!$A$1:$B$50000,2,0)")
        self.assertEqual(ws["G2"].value, "Urbano")
        self.assertEqual(ws["H2"].value, "00")
        self.assertEqual(ws["I2"].value, "06/11/2026")
        self.assertEqual(ws["J2"].value, "10:03:55 AM")
        self.assertEqual(ws["K2"].value, "RUA DAS FLORES, 142 - CENTRO")
        self.assertEqual(ws["L2"].value, 23873)
        self.assertEqual(ws["M2"].value, 24182)
        self.assertEqual(ws["N2"].value, "=SUM(M2-L2)")
        self.assertEqual(ws["O2"].value, "Faturada")
        self.assertEqual(ws["P2"].value, "N00")
        self.assertEqual(ws["Q2"].value, "N00")
        self.assertEqual(ws["R2"].value, '=IFERROR(VLOOKUP(B2,Filtros!$D$1:$E$50000,2,0),"Faturada")')
        self.assertEqual(ws["S2"].value, '=IFERROR(VLOOKUP(B2,Filtros!$G$1:$H$50000,2,0),"Faturada")')
        self.assertEqual(ws["T2"].value, '=IFERROR(VLOOKUP(B2,Filtros!$J$1:$K$50000,2,0),"N-OK")')
        self.assertIn("Controle Horario", wb.sheetnames)
        ws_ch = wb["Controle Horario"]
        self.assertEqual([ws_ch.cell(1, c).value for c in range(1, 6)], CONTROLE_HORARIO_HEADERS)
        self.assertEqual(ws_ch["A2"].value, "4MW")
        self.assertEqual(ws_ch["B2"].value, "06/11/2026")
        self.assertEqual(ws_ch["C2"].value, "10:03")
        self.assertEqual(ws_ch["D2"].value, "10:03")
        self.assertEqual(ws_ch["E2"].value, "00:00")
        self.assertNotIn("CNA", all_headers)
        self.assertNotIn("Leiturista", all_headers)
        self.assertNotIn("c", all_headers)
        self.assertNotIn("A Exec.", all_headers)
        self.assertNotIn("Exec.", all_headers)

    def test_invalid_layout_raises_validation_error(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Planilha Aleatória"
        buf = io.BytesIO()
        wb.save(buf)
        with self.assertRaises(CemigValidationError):
            process_matriz_bytes(buf.getvalue(), "invalida.xlsx")

    def test_leitura_3line_block_layout(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Leitura"

        def s(r: int, c: int, v: object) -> None:
            cell = ws.cell(r, c, v)
            if isinstance(v, str):
                cell.number_format = "@"

        s(1, 1, "CEMIG Distribuição S.A.")
        s(1, 12, "Pág. 1/2")
        s(2, 1, "Seqüência de Leitura Executada")
        s(3, 1, "Unidade de Leitura: 09102801")
        s(4, 1, "Leiturista: 5JL")
        s(5, 1, "Unidade")
        s(5, 2, "Medidor")
        s(5, 3, "T. Reg")
        s(5, 4, "OC")
        s(5, 5, "c")
        s(5, 6, "Data")
        s(5, 7, "Hora")
        s(5, 10, "Leit. Ant.")
        s(5, 11, "Leit. Exec")
        s(5, 12, "Imp.")
        s(5, 13, "Apres.")
        s(6, 1, "Consumidora")
        s(7, 1, "Cliente")
        s(7, 7, "Endereço")
        s(7, 13, "Qtd Dig.")

        s(8, 1, "2.796.450.018-02")
        s(8, 2, "APG176044248")
        s(8, 3, "03")
        s(8, 4, "1800")
        s(8, 6, 46343)
        s(8, 7, "09:11")
        s(8, 10, "15741")
        s(8, 11, "15927")
        s(8, 13, "00")
        s(9, 6, "9:11:15 AM")
        s(9, 7, "RUA NILO PECANHA, 288 -  - CS")
        s(10, 1, "CARLOS AUGUSTO DE SOUZA PEREIRA")
        s(10, 4, "123.456.789-00")
        s(10, 13, "1")

        buf = io.BytesIO()
        wb.save(buf)
        out_bytes, report = process_matriz_bytes(buf.getvalue(), "Leitura.xlsx")
        self.assertEqual(report.final_rows, 1)
        out_wb = openpyxl.load_workbook(io.BytesIO(out_bytes), data_only=False)
        out_ws = out_wb["ControleLeitura"]
        self.assertEqual(out_ws["A2"].value, "2.796.450.018-02")
        self.assertEqual(out_ws["B2"].value, "APG176044248")
        self.assertEqual(out_ws["C2"].value, "03")
        self.assertEqual(out_ws["D2"].value, "09102801")
        self.assertEqual(out_ws["E2"].value, "1028")
        self.assertEqual(out_ws["G2"].value, "Urbano")
        self.assertEqual(out_ws["H2"].value, "1800")
        self.assertEqual(out_ws["I2"].value, "17/11/2026")
        self.assertEqual(out_ws["J2"].value, "09:11")
        self.assertEqual(out_ws["K2"].value, "RUA NILO PECANHA, 288 -  - CS")
        self.assertEqual(out_ws["L2"].value, 15741)
        self.assertEqual(out_ws["M2"].value, 15927)
        self.assertEqual(out_ws["P2"].value, "1")
        self.assertEqual(out_ws["Q2"].value, "00")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--matriz")
    parser.add_argument("--ref")
    args, unknown = parser.parse_known_args()
    sys.argv = [sys.argv[0], *unknown]
    unittest.main()
