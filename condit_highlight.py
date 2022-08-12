from operator import index
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule

faves = [
    "the great gatsby",
    "harry potter and the prisoner of azkaban",
    "TFiOS",
    "the bell jar",
    "slaughterhouse five",
    "looking for alaska",
]


def condit_highlight():
    # open file
    excel = "excel_docs\in_progress.xlsx"
    wb = load_workbook(excel)
    ws = wb.active
    yellow_fill = PatternFill(bgColor="FFFF00")
    new_style = DifferentialStyle(fill=yellow_fill)
    rule = Rule(type="expression", dxf=new_style, stopIfTrue=True)
    y = ""
    # for y in faves:
    #     rule.formula = [f'$A1:A17= "{y}"']
    # #     ws.conditional_formatting.add("A1:D17", rule)
    # ALMOST WORKS
    # for rows in ws.iter_rows(min_row=1, max_row=15, min_col=0, max_col=3):
    #     for cell in rows:
    #         if cell.value in faves:
    #             print(cell)
    #             rule.formula = [f'$A1:A17= "{cell}"']
    #             ws.conditional_formatting.add("A1:D17", rule)
    # wb.save("my_test.xlsx")
    # cycle over cells in all rows


if __name__ == "__main__":
    condit_highlight()
