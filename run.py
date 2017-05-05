#!/usr/bin/env python
import sys
import tkinter
import bin.gui


if __name__ == '__main__':
    root = tkinter.Tk()
    app = bin.gui.Application(root)
    root.mainloop()
