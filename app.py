import openpyxl as xl
from openpyxl.chart import BarChart, Reference

def process_workbook(input_filename, output_filename):
    wb = xl.load_workbook(input_filename)
    sheet = wb['Sheet1']
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price

    values = Reference(sheet,
              min_row = 2,
              max_row = sheet.max_row,
              min_col = 4,
              max_col = 4)

    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'g2')
    wb.save(output_filename)

process_workbook('transactions.xlsx','transactions2.xlsx')