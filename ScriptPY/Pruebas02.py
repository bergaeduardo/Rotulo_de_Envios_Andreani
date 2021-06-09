import win32print

tempprinter = "\\\\192.168.0.64\\ZDesigner GC420t (EPL)"
print('tempprinter: '+ tempprinter)
currentprinter = win32print.GetDefaultPrinter()
print('currentprinter: '+ currentprinter)
win32print.SetDefaultPrinter(tempprinter)
print('Se establecio la siguiente impresora de rotulos: \n'+ tempprinter)
print('Impresora actual: \n' + currentprinter)