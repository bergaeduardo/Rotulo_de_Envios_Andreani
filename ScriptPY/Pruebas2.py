# import SelectPrint as sp
# import tkinter as tk
import win32print
# import win32api
# root = tk.Tk()
# sp.PrinterManager(root)
# root.mainloop()

#tempprinter = "\\\\192.168.0.64\\ZDesigner GC420t (EPL)"
tempprinter = "Microsoft Print to PDF"
currentprinter = win32print.GetDefaultPrinter()
print('currentprinter: '+ currentprinter)
win32print.SetDefaultPrinter(tempprinter)
aux=win32print.GetDefaultPrinter()
print('la impresora actual es: '+ aux)
# win32api.ShellExecute(0, "print", filename, None,  ".",  0)
# win32print.SetDefaultPrinter(currentprinter)
