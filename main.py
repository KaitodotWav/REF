import openpyxl
from thefuzz import fuzz
#fuzz.token_sort_ratio("fuzzy wuzzy was a bear", "wuzzy fuzzy was a bear")

dataframe = openpyxl.load_workbook("DILIMAN Google Sheeet.xlsx")
dataframe1 = dataframe.active

def check(item):
    save = True
    for part in item:
        if "," in part:
            save = False
    return save
n = input("name>> ")
collected = []
for row in dataframe1.iter_rows(0, dataframe1.max_row):
    name = row[1].value
    c = fuzz.token_sort_ratio(name, n)
    if c > 55:
        collected.append(f"{c} : {name}")
for i in range(len(collected)):
    print(sorted(collected, reverse=True, key=lambda x: int(x.split(":")[0]))[i])

print(dataframe1.max_row)

