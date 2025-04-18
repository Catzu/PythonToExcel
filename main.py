from openpyxl import workbook, load_workbook

wk = load_workbook('excel.xlsx')
ws = wk.active

ws.title = "Sales Data"
ws.append(['Date', 'Sales Rep', 'Products', 'Units', 'Price'])
ws.append(['05/04/2025', 'Kim Possible', 'Phone', 3, 600])
ws.append(['18/04/2025', 'Kraus Harvestein', 'Laptop', 9, 2900])

wk.save('excel.xlsx')