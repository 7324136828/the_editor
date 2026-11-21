"""Generate charts_freeze_filter.xlsx (fixture for xlsx.rs tests).

Run: python gen_charts_freeze_filter.py  (needs openpyxl). Writes the .xlsx
next to this script.

Sheet "Data": header row A1:D1 + 5 data rows, frozen at B2, auto filter
A1:D6, and six charts (column, bar, line, pie, scatter, area) that each
reference the header row (series names) and the data.
"""
import os
from openpyxl import Workbook
from openpyxl.chart import (AreaChart, BarChart, LineChart, PieChart,
                            Reference, ScatterChart, Series)

wb = Workbook()
ws = wb.active
ws.title = "Data"
rows = [
    ("Month", "Sales", "Costs", "Profit"),
    ("Jan", 10, 7, 3),
    ("Feb", 12, 8, 4),
    ("Mar", 15, 9, 6),
    ("Apr", 11, 10, 1),
    ("May", 18, 11, 7),
]
for r in rows:
    ws.append(r)
ws.freeze_panes = "B2"
ws.auto_filter.ref = "A1:D6"

cats = Reference(ws, min_col=1, min_row=2, max_row=6)


def add(chart, data_ref, anchor, title):
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    if title:
        chart.title = title
    ws.add_chart(chart, anchor)


col = BarChart()
col.type = "col"
add(col, Reference(ws, min_col=2, min_row=1, max_col=3, max_row=6), "F2", "Sales vs costs")

bar = BarChart()
bar.type = "bar"
add(bar, Reference(ws, min_col=2, min_row=1, max_col=4, max_row=6), "P2", "Horizontal")

line = LineChart()
add(line, Reference(ws, min_col=2, min_row=1, max_col=2, max_row=6), "F20", "Trend")

pie = PieChart()
add(pie, Reference(ws, min_col=2, min_row=1, max_col=2, max_row=6), "P20", None)

sc = ScatterChart()
sc.title = "Scatter"
sc.style = 13
xs = Reference(ws, min_col=2, min_row=2, max_row=6)
ys = Reference(ws, min_col=3, min_row=1, max_row=6)
s = Series(ys, xs, title_from_data=True)
s.marker.symbol = "circle"
s.graphicalProperties.line.noFill = True
sc.series.append(s)
ws.add_chart(sc, "F38")

area = AreaChart()
add(area, Reference(ws, min_col=2, min_row=1, max_col=3, max_row=6), "P38", "Area")

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts_freeze_filter.xlsx")
wb.save(out)
print("wrote", out)
