import openpyxl
import pandas



def open_roster(path, skip=2):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = ws.values

    for s in range(skip):
        next(data)

    cols = next(data)
    df = pandas.DataFrame(list(data), columns=cols)

    return df
