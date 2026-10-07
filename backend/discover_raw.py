"""
Roda isso se, num próximo mês, o painel avisar que teve blocos ignorados
(tamanho de bloco fora de 7/8 campos) — mostra os valores brutos da planilha,
linha a linha, pra você conferir visualmente onde o formato mudou.

Uso:
    cd backend
    python discover_raw.py data/chamados.xlsx [linha_inicial] [linha_final]
"""

import sys
from pathlib import Path

import openpyxl

caminho = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/chamados.xlsx")
inicio = int(sys.argv[2]) if len(sys.argv) > 2 else 1
fim = int(sys.argv[3]) if len(sys.argv) > 3 else 80

wb = openpyxl.load_workbook(caminho, data_only=True)
ws = wb.active
print(f"Planilha: {ws.title} — dimensões: {ws.dimensions}")
for i, row in enumerate(ws.iter_rows(min_row=inicio, max_row=fim, values_only=True), start=inicio):
    print(i, repr(row[0]))
