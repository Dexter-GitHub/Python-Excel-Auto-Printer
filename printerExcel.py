import os
import win32api
import win32print

def selectPrinter():
    index = 0
    Printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
    print("[Print List]")
    print("Current Printer Name: " + win32print.GetDefaultPrinter())
    print("0. Close")
    for Printer in Printers:   
        index += 1              
        print(str(index) + ". " + Printer['pPrinterName'])

    number = input("Select Printer Number: ")
    if len(Printers) >= int(number) and  int(number) > 0:  
        win32print.SetDefaultPrinter(Printers[int(number)-1]['pPrinterName'])
    print("Setting Printer: " + win32print.GetDefaultPrinter())
    return  win32print.GetDefaultPrinter()

def search(dirname, printerName):
    filenames = os.listdir(dirname)
    select = input("Are you sure that you want to print the whole Excel file ?(y or n): ")
    if select == 'y' or select == 'Y':        
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)                       
            ext = os.path.splitext(full_filename)[-1]        
            if ext == '.xlsm' or ext == '.xlsx':
                print(full_filename + " 인쇄 시작 !")
                win32api.ShellExecute(0, 'printto', full_filename, '"' + printerName + '"', None,  0)
    else:
        os.system('cls')
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)                       
            ext = os.path.splitext(full_filename)[-1]        
            if ext == '.xlsm' or ext == '.xlsx':
                select = input('"' + filename + '"' + " you want to print Excel file ?(y or n): ")
                if select == 'y' or select == 'Y': 
                    print('"' + filename + '"' + " 인쇄 시작 !")
                    win32api.ShellExecute(0, 'printto', full_filename, '"' + printerName + '"', None,  0)
                        
if __name__ == "__main__":    
    printerName = selectPrinter()
    os.system('cls')
   
    search(os.getcwd(), printerName) 
    os.system("Pause")

# VBA Insert Module Code 
# Sub Auto_Close()
#     If ThisWorkbook.Saved = False Then
#         ThisWorkbook.Save
#     End If
# End Sub





