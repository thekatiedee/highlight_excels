# for x in df[df["books"]]:
#     print("ex:")
#     print(x)
#     if x in faves:
#         print(x)

# print("df.loc:", df.loc[3])

# format1 = df.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

# from condit formatting:
"""
excel = "excel_docs\in_progress.xlsx"
df = pd.read_excel(excel, dtype=str, engine="openpyxl")
print(df)

# for x in rows
# if x in list
# highlight yellow

faves = ["the great gatsby", "the catcher in the rye"]

(row, col) = df.shape

writer = pd.ExcelWriter("output\condit_highlight.xlsx", engine="xlsxwriter")

# research: data = pd.read_excel(target_file, sheet_name = None) returns a Dict. When you loop over data you get the keys of the Dict.
# df_dict = pd.read_excel(excel, sheet_name=None, dtype=str, engine="openpyxl")

# '.book' is like calling 'xlsxwriter.Workbook('file name')
wb = writer.book
format_yellow = wb.add_format({"bg_color": "yellow"})

print(row, col)

ws = writer.sheets
# (first row, first col, last row, last col)
df.conditional_format(1, 0, 1, 0, {"type": "no_blanks", "format": format_yellow})

writer.save()

# conditionally format one

# conditionally format another

# loop the conditionally format
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

"""
# loop the conditionally format
for x in faves:
    my_title = df[df["books"] == x]
    print(my_title)
    # conditionally format
    # (first row, first col, last row, last col)
