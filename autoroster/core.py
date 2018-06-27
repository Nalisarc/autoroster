
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
def get_exam_list(dataframe):
    raw_exam_list = list(dataframe["SP Exam"].unique())
    exam_list = [exam for exam in raw_exam_list if str(type(exam)) == "<class 'str'>"]
    exam_list.sort()
    return exam_list
