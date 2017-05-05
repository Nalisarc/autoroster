import time
from autoroster import core
from argparse import ArgumentParser
import sys


def main():
    date = time.strftime("%x", time.localtime())
    parser = ArgumentParser(description="Generate reports from roster")
    parser.add_argument("-f", "--file", 
                        type=str,
                        default=None,
                        help="File to operate on")
    args = parser.parse_args()
    path = args.file
    if path is None:
        print("Please enter the path to file")
        path = input("==> ")
    else:
        pass
    wb = core.make_workbook()
    exam_exports = core.read_in_roster(path)
    exam_list = core.get_exam_list(exam_exports)
    exams_to_process = prompt_for_exams(exam_list)
    for exam in exams_to_process:
        student_info = core.get_student_info(exam,exam_exports)
        core.make_sheet(wb,exam,date,student_info)

    sheet_to_delete = wb.get_sheet_by_name('Sheet')
    wb.remove_sheet(sheet_to_delete)    
    outpath = get_outpath()
    core.save_workbook(wb, outpath)
    return None


def prompt_for_exams(exam_list):
    for i, item in enumerate(exam_list):
        print(i, item)

    output = []
    run = True
    print("Enter exam number to add it to list")
    print("Enter exit when finished")
    while run:
        user_input = input("==> ")
        if user_input.lower() == "exit":
            run = False
            continue
        try:
            output.append(exam_list[int(user_input)])
            continue
        except:
            print("Error: {0} is an invalid request".format(user_input))

    return output

def get_outpath():
    default = "report" + time.strftime("%m-%d-%y",time.localtime()) + ".xlsx"
    print("Enter name of the new file [Default: {0}]".format(default))
    outpath = input("==> ")
    if outpath == '':
        outpath = default
    return outpath



if __name__ == '__main__':
    main()
    sys.exit()
