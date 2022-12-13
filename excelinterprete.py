# This is mian file
import xlrd
import math
from pprint import pprint

SRCBOOK = "/home/eugeneai/tmp/08.04.21 АФК.xls"
PAGE = "Лист1"
LABELS = "ABCDEFGH"

print("Hello, World!, This is a cool excal parsing program!")


book = xlrd.open_workbook(SRCBOOK)
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))

sheet = book.sheet_by_name(PAGE)

# tmpsell = page.

# print(page.__class__)

F = {"Дата": "expldate",
    "Время": "expltime",
    "Планшет": "desk",
    "Методика": "technique",
    "Фильтр": "filter"}

def gettext(row):
    return " ".join([str(i.value) for i in row])

startmatrix = False
previndex = None

def convert(val):
    # val = val.replace(",",".")
    # val = float(val)
    return val

def interprete(num, row, text):
    global previndex
    try:
        fname, fval = text.split(":")
        fval = fval.strip()
        if fname in ["Планшет", "Фильтр"]:
            fval = math.floor(float(fval))
        # print("{} = {}".format(fname, fval))
        if fname == "Фильтр":
            startmatrix =True
        return (fname, fval)
    except ValueError:
        pass

    ch = str(row[0].value)
    index = LABELS.find(ch) if ch else -1
    # print(ch, LABELS, index)
    if index>=0:
        previndex = index
        return None
    else:
        if previndex is not None:
            row = row[1:]
            row = [convert(c.value) for c in row]
            digits = range(12)
            digits = [i + previndex*12 + 1 for i in digits]
            return list(zip(digits, row))
        else:
            return None

def main():
    answer = {}
    probes = []
    
    for i, row in enumerate(sheet.get_rows()):
        res = interprete(i, row, gettext(row))
        if res is None:
            continue
        if len(res) == 2:
            answer[res[0]] = res[1]
            continue
        probes = probes + res

    return (answer, probes)
if __name__=="__main__":
    pprint(main())
