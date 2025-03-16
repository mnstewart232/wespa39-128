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

#Takes a float (dec_inches) and returns string formatted as XXft YYin or YYin if no feet
def decInchesToFtIn(dec_inches):
    feet = math.trunc(dec_inches/12)
    inches = format(dec_inches - feet * 12, '.2f')

    if (feet > 0):
        return "{0} FT {1} IN".format(feet, inches)
    else:   
        return "{0} IN".format(inches)


class MainMenu:

    def clear():
        print("Clearing Barcodes...")


    def resetLaser():
        print("Resetting Laser...")


    def printLabel():
        print("Printing Label...")

    scannerInput = ""
    currentBarcode = ""
    orderVal = "" #First 4 digits of a line128 barcode
    orderLength = 0.0 #line39 code, or the remaining digits of a line128
    tableLength = 284.79 #23ft 8.79in #Fill this in from laser scanner
    offByVal = 0.0    
    minTolerance = 0.1 #Fill this in from config file?
    maxTolerance = 6.0 #Fill this in from config file?
    allowPrint = "disabled"
    toleranceIndicatorVal = ""
    toleranceColorVal = ""

    #Test values
    #orderLength = 284.75 #23ft 8.75in
    #orderVal = "1234"

    #Call this once the bardcode has been detected. Update the GUI with the new barcode.
    def updateGUI(self):
        print("Updating GUI. Barcode: {self.currentBarcode}")

        self.orderVal = ""
        self.orderLength = 0.0
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.

        #This assumes a Line128 style code. How to better detect what kind of code it is?
        #Simple, check if the first 4 chars are all digits. If so, it's a Line128 code.
        if (self.currentBarcode[0:4].isdigit()):
            self.orderVal = self.currentBarcode[0:4]
            self.orderLength = float(self.currentBarcode[4:])
        else:
            #Otherwise it's line39!
            self.orderLength = float(self.currentBarcode)


        self.offByVal = self.tableLength - self.orderLength

        #Will change between green, yellow, and red based on tolerance, with text changing as well (Within/Near/Outside Tolerance)
        if abs(self.offByVal) <= self.minTolerance and abs(self.offByVal) > 0:
            self.toleranceIndicatorVal = "Within Tolerance"
            self.toleranceColorVal = "green"
            self.allowPrint = "normal"
        elif abs(self.offByVal) <= self.maxTolerance and abs(self.offByVal) > self.minTolerance:
            self.toleranceIndicatorVal = "Near Tolerance"
            self.toleranceColorVal = "yellow"
        else:
            self.toleranceIndicatorVal = "Outside Tolerance"
            self.toleranceColorVal = "red"

        print(f"Order Length: {self.orderLength}, Order Value: {self.orderVal}")

    def captureInput(self, event):
        if event.keysym == 'Return':
            print(f"Received input: {self.scannerInput}")
            self.currentBarcode = self.scannerInput
            self.scannerInput = ""  # Clear the input after processing
            self.updateGUI()
        elif (event.char >= '0' and event.char <= '9' or event.char == '.'):
            self.scannerInput += event.char  # Append the character to the input string

    def __init__(self, root):
        #Could of course make the window size a config item...

        content = ttk.Frame(root, width=1500, height=900)
        content.grid(column=0, row=0, columnspan=5, rowspan=6)

        for i in range (0, 5):
            root.columnconfigure(i, weight=3)

        for i in range (0, 6):
            root.rowconfigure(i, weight=2)

        root.resizable(width=False, height=False)
        root.title("WESPA 39-128")

        # Bind keyboard shortcuts; also detect barcode input
        root.bind('<x>', lambda event: self.clear())
        root.bind('<l>', lambda event: self.resetLaser())
        root.bind('<space>', lambda event: self.printLabel())
        root.bind('<Key>', self.captureInput)
        
        #How to input?
        #root.bind('<Key>', self.captureInput)
        
        lbl_lastBarcode = ttk.Label(content, text="Last Barcode Scanned:")
        lbl_lastBarcode.grid(column=0, row=0, padx=5, pady=5, sticky="W")

        lbl_workOrder = ttk.Label(content, text="Work Order: ")
        lbl_workOrder.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        lbl_workOrderVal = ttk.Label(content, text=self.orderVal)
        lbl_workOrderVal.grid(column=2, row=0, padx=5, pady=5, sticky="W")
        
        lbl_length = ttk.Label(content, text="Length: ")
        lbl_length.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        lbl_lengthVal = ttk.Label(content, text=decInchesToFtIn(self.orderLength))
        lbl_lengthVal.grid(column=2, row=1, padx=5, pady=5, sticky="W")

        btn_clear = ttk.Button(content, text="CLEAR (X)")
        btn_clear.grid(column=3, row=0, rowspan=2)
        btn_clear.bind('<Button-1>', lambda event: self.clear())

        btn_resetLaser = ttk.Button(content, text="RESET LASER (L)")
        btn_resetLaser.grid(column=4, row=0, rowspan=2)
        btn_resetLaser.bind('<Button-1>', lambda event: self.resetLaser())
        
        lbl_tableLength = ttk.Label(content, text="TABLE LENGTH: " + decInchesToFtIn(self.tableLength))
        lbl_tableLength.grid(column=0, row=2, columnspan=2, padx=5, pady=5, sticky="W")

        lbl_offBy = ttk.Label(content, text="OFF BY: " + decInchesToFtIn(self.offByVal))
        lbl_offBy.grid(column=2, row=2, columnspan=1, padx=5, pady=5, sticky="W")

        lbl_orderLength = ttk.Label(content, text="ORDER LENGTH: " + decInchesToFtIn(self.orderLength))
        lbl_orderLength.grid(column=3, row=2, columnspan=2, padx=5, pady=5, sticky="W")


        lbl_toleranceIndicator = ttk.Label(content, text=self.toleranceIndicatorVal, background=self.toleranceColorVal)
        lbl_toleranceIndicator.grid(column=1, row=4, columnspan=3, padx=5, pady=5, sticky="S")
        
        btn_print = ttk.Button(content, text="PRINT\n(spacebar)")
        btn_print.configure(state=self.allowPrint)
        btn_print.bind("<Button-1>", lambda event: self.printLabel())
        btn_print.grid(column=2, row=5, rowspan=1, padx=5, pady=5, sticky="S")


root = Tk()
MainMenu(root)
root.mainloop()

