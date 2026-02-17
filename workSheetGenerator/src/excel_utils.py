# =========================
# PLANILHAS PORTO
# =========================

from copy import copy

def copy_style(ws, rows):
    base = ws[10]
    for r in range(10, rows):
        for c, ref in enumerate(base, start=1):
            cell = ws.cell(r, c)
            cell.font = copy(ref.font)
            cell.border = copy(ref.border)
            cell.fill = copy(ref.fill)
            cell.number_format = copy(ref.number_format)
            cell.alignment = copy(ref.alignment)

def copy_formulas(ws, start_row, end_row):
    for r in range(start_row + 1, end_row):
        ws[f'B{r}'] = f'=TEXT(A{r},"dddd")'
        ws[f'E{r}'] = f'=IF(D{r}="SIM",C{r}*2,C{r})'