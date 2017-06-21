#!/usr/bin/env python
import tkinter
import tkinter.filedialog
import tkinter.messagebox
import time
import autoroster.core

FILEOPENOPTIONS = dict(defaultextension='.xlsx',
                       filetypes=[('Excel File','*.xlsx'),("All Files",'*.*')])
class Application:

    def __init__(self, master):
        self.master = master
        self.frame = tkinter.Frame(self.master)
        self.roster_name = tkinter.Label(master, text="No Roster Selected")
        self.roster_name.grid(row=0, column=0)
        self.master.title("ARR")
        self.master.resizable(False, False)

        self.report_variable = tkinter.StringVar(master)
        self.report_variable.set("daily") # default value
        self.report_text = tkinter.Label(master, text="Report Type:").grid(row=1, column=0, sticky=tkinter.W)

        self.report_type = tkinter.OptionMenu(master, self.report_variable, "daily", "weekly")
        self.report_type.grid(row=1, column=1)

        self.open_roster_button = tkinter.Button(master, text="Open", command=self.open_roster).grid(row=0, column=1)

        self.checkbox = tkinter.Listbox(master, selectmode="extended")
        self.checkbox.grid(row=2,
                           column=0,
                           padx=5,
                           pady=5,
                           sticky= tkinter.W + tkinter.E + tkinter.S,
                           rowspan=2,
                           columnspan=3
        )
        self.generate_button = tkinter.Button(master, text="Generate Report", command=self.generate_report)
        self.generate_button.grid(columnspan=2)

    def open_roster(self):
        filename = tkinter.filedialog.askopenfilename(**FILEOPENOPTIONS)
        roster = filename.split('/')[-1]
        self.roster_name.config(text=roster)

        self.report_dataframe = autoroster.core.read_in_roster(filename)
        exam_list = autoroster.core.get_exam_list(self.report_dataframe)

        for exam in exam_list:
            self.checkbox.insert('end', exam)

    def generate_report(self):
        type_ = self.report_variable.get()
        exams = [self.checkbox.get(idx) for idx in self.checkbox.curselection()]
        wb = autoroster.core.make_workbook()
        if type_ == "daily":
            date = time.strftime("%x", time.localtime())
            for exam in exams:
                student_info = autoroster.core.get_student_info(exam, self.report_dataframe)
                asterisk_free = autoroster.core.ignore_asterisk(student_info)
                autoroster.core.make_daily_report(wb,exam,date,asterisk_free)
        elif type_ == "weekly":
            year = time.strftime("%Y", time.localtime())
            for exam in exams:
                student_info = autoroster.core.get_student_info(exam, self.report_dataframe)
                asterisk_free = autoroster.core.ignore_asterisk(student_info)
                autoroster.core.make_weekly_report(wb,exam,year, asterisk_free)
        else:
             tkinter.messagebox.showerror("Unexpected Error", """An invalid report type was selected,
                                          Please send an email to u0346076@utah.edu with what option you selected""")

        autoroster.core.delete_blank_sheets(wb)
        outpath = tkinter.filedialog.asksaveasfilename(**FILEOPENOPTIONS)
        autoroster.core.save_workbook(wb, outpath)
        tkinter.messagebox.showinfo("Sucess!","File was sucessfully made!")





def do_nothing():
    pass

def main():
    root = tkinter.Tk()
    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    main()
    sys.exit()
