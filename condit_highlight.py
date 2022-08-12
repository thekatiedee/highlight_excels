from operator import index
from turtle import color
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule


def condit_highlight():

    faves = [
        "the great gatsby",
        "harry potter and the prisoner of azkaban",
        "TFiOS",
        "the bell jar",
        "slaughterhouse five",
        "looking for alaska",
    ]

    # open file
    excel = "excel_docs\in_progress.xlsx"
    # openpyxl way of opening a workbook
    wb = load_workbook(excel)
    ws = wb.active
    # highlight cell yellow
    yellow_fill = PatternFill(bgColor="FFFF00")
    new_style = DifferentialStyle(fill=yellow_fill)
    rule = Rule(type="expression", dxf=new_style, stopIfTrue=True)

    # NEARLY WORKS
    # this prints the individual rows that contain the faves items out of the list of all books
    for rows in ws.iter_rows(min_row=1, max_row=15, min_col=0, max_col=3):
        for cell in rows:
            # if statement works now!
            if cell.value in faves:
                # correctly prints six times
                print("yes")
                # rule only applies to line 7 - it's overwriting
                rule.formula = [f'$A1:A17= "{cell.value}"']
                ws.conditional_formatting.add("A1:D17", rule)
                wb.save("my_test.xlsx")
                # only saving the last iteration ('looking for alaska')


if __name__ == "__main__":
    condit_highlight()
