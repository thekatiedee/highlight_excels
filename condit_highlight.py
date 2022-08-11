from operator import index
import pandas as pd


def condit_highlight():
    # open file
    excel = "excel_docs\in_progress.xlsx"
    df = pd.read_excel(excel)
    # print(df)

    title = "the catcher in the rye"

    # for x in rows
    # if x in list
    # highlight yellow

    faves = ["the great gatsby", "the catcher in the rye"]

    # for x in df[df["books"]]:
    #     print("ex:")
    #     print(x)
    #     if x in faves:
    #         print(x)

    # print("df.loc:", df.loc[3])

    (row, col) = df.shape

    # format1 = df.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

    writer = pd.ExcelWriter("output\condit_highlight.xlsx", engine="xlsxwriter")

    # research: data = pd.read_excel(target_file, sheet_name = None) returns a Dict. When you loop over data you get the keys of the Dict.
    df_dict = pd.read_excel(excel, sheet_name=None, dtype=str, engine="openpyxl")
    # to extract values in dictionary: feature3 = [d.get('Feature3') for d in a]

    values = [z.get("the great gatsby") for z in df_dict]

    print(values)

    return
    # '.book' is like calling 'xlsxwriter.Workbook('file name')
    wb = writer.book

    format_yellow = wb.add_format({"bg_color": "yellow"})

    print(row, col)
    for x in faves:
        my_title = df[df["books"] == x]
        print(my_title)
        # conditionally format
        # (first row, first col, last row, last col)

    print(df.iloc[5, 0])

    print(df.loc[0])

    df.style.apply(
        lambda x: "background-color : yellow" if x.value() == df.iloc[5, 0] else ""
    )

    df.to_excel("conditional_highlighting.xlsx")


if __name__ == "__main__":
    condit_highlight()
