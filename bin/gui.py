#!/usr/bin/env python
import tkinter
import tkinter.filedialog
import tkinter.messagebox
import time
import autoroster.core

class Application:

    def __init__(self, master):
        self.master = master
        self.frame = tkinter.Frame(self.master)
        self.roster_name = tkinter.Label(master, text="No Roster Selected")
        self.roster_name.grid(row=0, column=0)
        self.master.title("ARR")
        self.master.resizable(False, False)

        #variable = tkinter.StringVar(master)
        #variable.set("daily") # default value
        #self.report_text = tkinter.Label(master, text="Report Type:").grid(row=1, column=0, sticky=tkinter.W)

        #self.report_type = tkinter.OptionMenu(master, variable, "daily", "weekly")
        #self.report_type.grid(row=1, column=1)

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
        filename = tkinter.filedialog.askopenfilename()
        roster = filename.split('/')[-1]
        self.roster_name.config(text=roster)

        self.report_dataframe = autoroster.core.read_in_roster(filename)
        exam_list = autoroster.core.get_exam_list(self.report_dataframe)

        for exam in exam_list:
            self.checkbox.insert('end', exam)

    def generate_report(self):

        exams = [self.checkbox.get(idx) for idx in self.checkbox.curselection()]
        wb = autoroster.core.make_workbook()
        date = time.strftime("%x", time.localtime())
        for exam in exams:
            student_info = autoroster.core.get_student_info(exam, self.report_dataframe)
            autoroster.core.make_daily_report(wb,exam,date,student_info)
        autoroster.core.delete_blank_sheets(wb)
        outpath = tkinter.filedialog.asksaveasfilename()
        autoroster.core.save_workbook(wb, outpath)
        tkinter.messagebox.showinfo("Sucess!","File was sucessfully made!")





def do_nothing():
    pass
if __name__ == '__main__':
    root = tkinter.Tk()
    app = Application(root)
    root.mainloop()
