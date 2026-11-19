from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference

ws_in = load_workbook("sales.xlsx").active

by_manager = defaultdict(float)
by_product = defaultdict(float)
by_day = defaultdict(float)

for row in ws_in.iter_rows(min_row=2, values_only=True):
    date_, manager, product, qty, price, total = row
    by_manager[manager] += total
    by_product[product] += total
    by_day[date_] += total

wb_out = Workbook()
ws1 = wb_out.active
ws1.title = "По менеджерам"
ws2 = wb_out.create_sheet("По товарам")
ws3 = wb_out.create_sheet("По дням")

def fill(ws, title, data, sort_by_value=True):
    ws.append([title, "Сумма продаж"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    if sort_by_value:
        items = sorted(data.items(), key=lambda x: x[1], reverse=True)
    else:
        items = sorted(data.items(), key=lambda x: x[0])
    for name, total in items:
        ws.append([name, total])
    ws.append(["Итого", sum(data.values())])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18

def fill_with_share(ws, title, data):
    total_all = sum(data.values())
    ws.append([title, "Сумма продаж", "Доля, %"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for name, total in sorted(data.items(), key=lambda x: x[1], reverse=True):
        share = total / total_all if total_all else 0   # дробь 0..1
        ws.append([name, total, share])
    ws.append(["Итого", total_all, 1 if total_all else 0])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = "0.0%"
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12

fill_with_share(ws1, "Менеджер", by_manager)
fill(ws2, "Товар", by_product)
fill(ws3, "Дата", by_day, sort_by_value=False)

last_data_row = ws1.max_row - 1
chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "Продажи по менеджерам"
chart.y_axis.title = "Сумма продаж"
chart.x_axis.title = "Менеджер"

data_categories = Reference(ws1, min_col=1, min_row=2, max_row=last_data_row)
data_values = Reference(ws1, min_col=2, min_row=1, max_row=last_data_row)

chart.add_data(data_values, titles_from_data=True)
chart.set_categories(data_categories)
chart.width = 18
chart.height = 10
ws1.add_chart(chart, "E2")

wb_out.save("report.xlsx")
print("Готово: report.xlsx")