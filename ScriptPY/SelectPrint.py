# -*- coding: utf-8 -*-
"""
Created on Tue Jun 16 10:41:35 2020

@author: eduardo.berga
"""

import tkinter as tk
import tkinter.ttk as ttk

import win32print


def get_available_printers():
    return [printer[2] for printer in win32print.EnumPrinters(2)]


class PrinterManager(tk.Frame):

    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        self.configure_interface()
        self.create_widgets()

    def configure_interface(self):
        self.master.title('Printer Manager')
        self.master.geometry('350x100')
        self.master.resizable(False, False)
        self.master.config(background='#626a77')

    def create_widgets(self):
        self.default_printer_label = tk.Label(self.master, bg='#626a77', fg='white')
        self.default_printer_label.place(x=10, y=12)
        self.update_default_printer_label()

        refresh_button = tk.Button(self.master, text='Refresh', command=self.update_default_printer_label)
        refresh_button.place(x=285, y=10)

        selected_printer = tk.StringVar()
        printer_choice_menu = ttk.Combobox(self.master, textvariable=selected_printer, values=get_available_printers(), width=35, state='readonly')
        printer_choice_menu.place(x=12, y=62)

        set_default_printer_button = tk.Button(self.master, text='Set', command=lambda: self.set_default_printer(selected_printer))
        set_default_printer_button.place(x=285, y=60, width=50)

    def update_default_printer_label(self):
        default_printer = win32print.GetDefaultPrinter()
        default_printer_text = 'Default printer: {}'.format(default_printer)
        self.default_printer_label.config(text=default_printer_text)
		print(default_printer_text)

    def set_default_printer(self, printer_name):
        win32print.SetDefaultPrinter(printer_name.get())
        self.update_default_printer_label()


if __name__ == '__main__':
    root = tk.Tk()
    PrinterManager(root)
    root.mainloop()