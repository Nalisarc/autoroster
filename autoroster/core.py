import openpyxl
import pandas


def open_roster(path, skip=2):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = ws.values

    for s in range(skip):
        next(data)

    cols = next(data)[1:]
    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    df = pandas.DataFrame(data,index=idx,columns=cols)

    return df
