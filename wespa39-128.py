import math
import serial
from tkinter import *
from tkinter import ttk

#Major parts of the program:
#Inputs:
    #Barcode scanner, connected via USB, keyboard-like input in Line39 or Line128 format.
    # Line39 currently, simply the length
    # Line128 may be a future expansion with the Work Order as the first 4 digits, length as remainder.
    #Laser scanner, COM port, Acuity AR1000 format (see pdf)
    #Config file, located in a sufficiently obscure place (like AppData)
    #Laser offset value, tolerance, and other non-operator settings in Config file
#Outputs:
    #Zebra Printer, connected via USB, uses ZPL II
#GUI:
    #Display info from laser scanner and barcode, refresh laser in realtime:
    #Table Length/Actual Length in Ft + Decimal Inches
    #Order Length/Desired Length in Ft + Decimal Inches
    #"Off by" Length being difference of the two
    #Green/Yellow/Red Tolerance indicator
    #Laser On/Off toggle buttons
    #Clear button to refresh info from inputs (X button on keyboard)
    #Print button toggled by spacebar; disable if measurements not within tolerance
    
    
        
#Libraries:
    #Pyinstaller for compiling to a single exe
    #tkinter for GUI
    #pyserial for COM IO

class MainMenu:

    def __init__(self, root):

        content = ttk.Frame(root, width=900, height=600)
        content.grid(column=0, row=0, columnspan=5, rowspan=6)
        root.resizable(width=True, height=True)
        root.title("WESPA 39/128")
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.
        orderVal = "1234" #AKA Order Number - first 4 digits of a line128 barcode
        orderLength = 284.75 #23ft 8.75in - the entire line39 code, or the remaining digits of a line128
        tolerance = 0.10 #Fill this in from config file
        tableLength = 284.79 #23ft 8.79in #Fill this in from laser scanner
        lbl_lastBarcode = ttk.Label(content, text="Last Barcode Scanned:")
        lbl_lastBarcode.grid(column=0, row=0)
        lbl_workOrder = ttk.Label(content, text="Work Order:")
        lbl_workOrder.grid(column=1, row=0)
        
        lbl_length = ttk.Label(content, text="Length:")
        lbl_length.grid(column=1, row=1)

        btn_clear = ttk.Button(content, text="CLEAR (X)")
        btn_clear.grid(column=3, row=0, rowspan=2)
        btn_resetLaser = ttk.Button(content, text="RESET LASER (L)")
        btn_resetLaser.grid(column=4, row=0, rowspan=2)
        
        lbl_tableLength = ttk.Label(content, text="TABLE LENGTH:")
        lbl_tableLength.grid(column=0, row=2, columnspan=2)
        
        lbl_offBy = ttk.Label(content, text="OFF BY:")
        lbl_offBy.grid(column=2, row=2, columnspan=1)
        #This value is tableLength - orderLength
        offByVal = tableLength - orderLength
        
        lbl_orderLength = ttk.Label(content, text="ORDER LENGTH:")
        lbl_orderLength.grid(column=3, row=2, columnspan=2)

        #Will change between green, yellow, and red based on tolerance, with text changing as well (Within/Near/Outside Tolerance)
        if abs(offByVal) <= tolerance:
            lbl_toleranceIndicator = ttk.Label(content, text="Within Tolerance")
            lbl_toleranceIndicator.configure(foreground="green")
        elif abs(offByVal) <= tolerance*2:
            lbl_toleranceIndicator = ttk.Label(content, text="Near Tolerance")
            lbl_toleranceIndicator.configure(foreground="yellow")
        else:
            lbl_toleranceIndicator = ttk.Label(content, text="Outside Tolerance")
            lbl_toleranceIndicator.configure(foreground="red")
            
        lbl_toleranceIndicator.grid(column=1, row=4, columnspan=3)
        
        btn_print = ttk.Button(content, text="PRINT\n(spacebar)")
        btn_print.grid(column=2, row=5, rowspan=1)



    #Takes a float (dec_inches) and returns string formatted as XX ft YY in
    def decInchesToFtIn(dec_inches):
        feet = dec_inches/12
        inches = dec_inches - feet
    
        return "{0}ft {1}in".format(feet, inches)



        

root = Tk()
MainMenu(root)
root.mainloop()
