from cgitb import text
import tkinter as tk
import ctypes
from tkinter import ANCHOR, OFF, ON, Menu, ttk, filedialog, Label
from win32com.client import Dispatch
import pathlib
import ctypes
import sys

def boxlabel():

    root = tk.Tk()
    root.geometry("500x300")
    root.configure(background= "Light Sky Blue")
    root.title("White Tape box Label")
    root.resizable(0,0)
    ctypes.windll.shcore.SetProcessDpiAwareness(1)

    #Menu Bar
    menuBar = Menu(root, font=("Times new roman", 13, 'bold'))
    menuOptions = Menu(menuBar,font=("Times new roman", 13, 'bold'), tearoff=0)
    
    menuOptions.add_command(label="Home", command= lambda:switchwindow())
    menuOptions.add_command(label="Exit", command= lambda:sys.exit())
    menuBar.add_cascade(label="More options",font=("Times new roman", 10, 'bold'), menu= menuOptions)
    root.config(menu=menuBar)

    #Labels
    Title: Label= tk.Label(root,text="This is the White Tape Label App",font=('Times new roman', 24, 'bold'),  bg='Light Sky Blue')
    palletlabel: Label = tk.Label(root,text="Pallet number :", font=('Times new roman', 18, 'bold'), bg='Light Sky Blue')
    boxlabel: Label = tk.Label(root,text="Box number    :", font=('Times new roman', 18, 'bold'), bg='Light Sky Blue' )
    #tapelabel: Label = tk.Label(root,text="Tapes", font=('Times new roman', 18, 'bold'), bg='aquamarine' )

    #Entry Boxes
    PalletEntryBox = tk.Entry(root, width=10, font=('Times new roman', 12, 'bold'))
    BoxEntryBox = tk.Entry(root, width=10, font=('Times new roman', 12, 'bold'))
    #tapeEntryBox = tk.Entry(root, width=10, font=('Times new roman', 12, 'bold'))

    #Printing Product Label
    printbtn = tk.Button(text = 'Print', command= lambda:(print_label()))


    #Labels grid
    palletlabel.place(x=100,y=100)
    boxlabel.place(x=102,y=150)
    #tapelabel.place(x=100,y=200)
    PalletEntryBox.place(x=275,y=105)
    BoxEntryBox.place(x=275,y=155)
    #tapeEntryBox.place(x=275,y=205)
    printbtn.place(x=400, y=125)
    Title.place(x=10, y=20)

    #How to print out the Label
    def print_label():
        my_printer = 'DYMO LabelWriter 450 Turbo'
        Pallet_Value = PalletEntryBox.get()
        Box_Value = BoxEntryBox.get()
        qr_code_value = BoxEntryBox.get(),PalletEntryBox.get()
        qr_path = pathlib.Path ('./US White tape label.label')

        printer_com = Dispatch ('Dymo.DymoAddIn')
        printer_com.SelectPrinter(my_printer)
        printer_com.Open(qr_path)
        printer_label = Dispatch ('Dymo.DymoLabels')
        printer_label.SetField("BARCODE_1", qr_code_value)
        printer_label.SetField("Box Number", Box_Value)
        printer_label.SetField("Pallet Number", Pallet_Value)
        printer_com.SetGraphicsAndBarcodePrintMode(ON)


    #Print Label
        printer_com.StartPrintJob()
        printer_com.Print(1,False)
        printer_com.EndPrintJob()

    def switchwindow():
        root.destroy()
        import main
        main.main()

    root.mainloop()
