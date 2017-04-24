#!/usr/bin/env python
import pandas
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import sys

def read_in_roster(path):
    df = pandas.read_excel(path, skiprows=2)
    return df


def get_exam_list(dataframe):
    raw_exam_list = list(dataframe["SP Exam"].unique())
    exam_list = [exam for exam in raw_exam_list if str(
        type(exam)) == "<class 'str'>"]
    exam_list.sort()
    return exam_list


def get_student_info(exam, dataframe):
    student_info = pandas.DataFrame(
        dataframe, columns=["Student Name", "No Show", "Completed"])
    by_exam = dataframe["SP Exam"] == exam
    processed = student_info[by_exam]
    sorted_info = processed.sort_values("Student Name")
    return sorted_info



def make_sheet(workbook, exam, date, student_info):
    # Text Constants
    title = "Exam Roster Report"
    headers = ["Student Name", "No Show", "Completed", "Check 1", "Check 2"]
    # Create workbook
    ws = workbook.create_sheet(exam)

    # Column widths
    ws.column_dimensions["A"].width = 19
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12

    ws["A1"] = title
    ws["B1"] = exam
    ws["A2"] = "Date: " + date
    ws["B2"] = "Check #1: ___________"
    ws["D2"] = "Check #2: ___________"
    ws.append(headers)

    for r in dataframe_to_rows(student_info, index=False, header=False):
        ws.append(r)

    return None


def save_workbook(workbook,path):
    workbook.save(path)
