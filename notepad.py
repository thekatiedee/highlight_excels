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


# [f(x) if condition else g(x) for x in sequence]

# ALMOST WORKS
# for rows in ws.iter_rows(min_row=1, max_row=15, min_col=0, max_col=3):
#     for cell in rows:
#         if cell.value in faves:
#             print(cell)
#             rule.formula = [f'$A1:A17= "{cell}"']
#             ws.conditional_formatting.add("A1:D17", rule)
# wb.save("my_test.xlsx")
# cycle over cells in all rows

# rule.formula = [f'$A1:A17= "{x}"']
# ws.conditional_formatting.add("A1:D17", rule)

# -----

y = ""
# for y in faves:
#     rule.formula = [f'$A1:A17= "{y}"']
#     ws.conditional_formatting.add("A1:D17", rule)

# if cell.value in faves:
#     print(cell)
#     rule.formula = [f'$A1:A17= "{cell}"']
#     ws.conditional_formatting.add("A1:D17", rule)

# cycle over cells in all rows
# book = "the great gatsby"
# if book in faves:
#     print("it's here!")
# pull col A into a list
# compare lists
# if col A item in faves list
# then highlight col A row

# new_list = []
# # this turns col A into a list
# for cell in ws["A"]:
#     new_list.append(cell.value)
#     print(cell)

# print(new_list)

# find if subset
# if True: highlight cell

#     cell.font = Font(color="ff007f", italic=True)
# wb.save("font_test.xlsx")

# print number of times an ISBN occurs in the ISBN col (list 'isbn'):
for x in isbn:
    print(f"ISBN: {x}")
    print(f"count: {isbn.count(x)}")

# print all key value pairs (ISBNs with counts)
for key, value in counted.items():
# print(key)
# k_list.append(key)
# print(value)
# v_list.append(key)
for it in range(1, 4):
    keys = sheet5[f"A{it}"]
    keys.value = key
    it += 1
    wb.save("my_test.xlsx")
print(value)

for it in counted:
    for key, value in counted.items():
        keys = sheet5[f"A{it}"]
        keys.value = key
        it += 1
        wb.save("my_test.xlsx")
print(value)


    # letters = ["a", "b", "c"]
    # num = [1, 2, 3]

    # for x in letters:
    #     print(x)
    # for y in num:
    #     print(y)

    # print all key value pairs (ISBNs with counts)
    # for key, value in counted.items():
    # print(key)
    # k_list.append(key)
    # print(value)
    # v_list.append(key)

    # 14
    # print(len(counted))

    # create a new sheet for a test
    # wb.create_sheet("teeeest")
    # sheet5 = wb["teeeest"]

    # letters = ["a", "b", "c"]
    # num = [1, 2, 3]

    # for x in letters:
    #     print(x)
    # for y in num:
    #     print(y)

    # # test = sheet5[""]

    # for it in range(1, 4):
    #     lets = sheet5[f"A{it}"]
    #     lets.value = key
    #     it += 1
    #     wb.save("my_test.xlsx")

    # # create a new sheet for the key
    # wb.create_sheet("what_does_it_mean")
    # sheet3 = wb["what_does_it_mean"]
    # # yellow are the books, pink are the authors
    # # intro
    # intro = sheet3["A1"]
    # intro.font = Font(name="Helvetica", size=14, bold=True)
    # intro.value = "KEY:"