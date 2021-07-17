from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side
from statistics import mean, stdev, median
from datetime import date

thin_border = Border(
    left = Side(style = 'thin'), 
    right = Side(style = 'thin'), 
    top = Side(style = 'thin'), 
    bottom = Side(style = 'thin')
)

wb = load_workbook('relatorio_painel.xlsx')
ws = wb.active

lastrow = ws.max_row

unit_values = [float(cell[0].value.replace(',', '.')) for cell in ws['H6':f'H{lastrow}']]
meses = [int(cell[0].value.split('/')[1]) for cell in ws['L6':f'L{lastrow}']]
anos = [int(cell[0].value.split('/')[2]) for cell in ws['L6':f'L{lastrow}']]

mes_atual = date.today().month
ano_atual = date.today().year

valores = []
for mes, ano, i in zip(meses, anos, range(0, lastrow)):
    if ano == 2020 and mes in range(mes_atual, 13):
        valores.append(unit_values[i])

    if ano == 2021 and mes in range(mes_atual + 1):
        valores.append(unit_values[i])
        
media = mean(valores)
desvio = stdev(valores)
coeficiente = desvio / media
mediana = median(valores)
preco = mediana if coeficiente > 0.25 else media

ws[f'A{lastrow + 2}'] = 'Média'
ws[f'A{lastrow + 3}'] = 'Desvio'
ws[f'A{lastrow + 4}'] = 'Coeficiente'
ws[f'A{lastrow + 5}'] = 'Mediana'
ws[f'A{lastrow + 6}'] = 'Preço'

ws[f'B{lastrow + 2}'] = media
ws[f'B{lastrow + 3}'] = desvio
ws[f'B{lastrow + 4}'] = coeficiente
ws[f'B{lastrow + 5}'] = mediana
ws[f'B{lastrow + 6}'] = preco

ws[f'A{lastrow + 2}'].border = thin_border
ws[f'A{lastrow + 3}'].border = thin_border
ws[f'A{lastrow + 4}'].border = thin_border
ws[f'A{lastrow + 5}'].border = thin_border
ws[f'A{lastrow + 6}'].border = thin_border
ws[f'B{lastrow + 2}'].border = thin_border
ws[f'B{lastrow + 3}'].border = thin_border
ws[f'B{lastrow + 4}'].border = thin_border
ws[f'B{lastrow + 5}'].border = thin_border
ws[f'B{lastrow + 6}'].border = thin_border

wb.save('resultado.xlsx')