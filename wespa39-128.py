import math
import serial
import tkinter as ttk
from tkinter import font

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
    
def resetLaser():
    print("Resetting Laser...")


def printLabel(isEnabled):
    if isEnabled == "normal":
        print("Printing Label...")


class MainMenu:
    scannerInput = "" #Barcode scanner input
    currentBarcode = "" #Last barcode scanned - delimited by newlines with the scanner
    orderVal = "" #First 4 digits of a line128 barcode
    orderLength = 0.0 #line39 code, or the remaining digits of a line128
    tableLength = 123.51 #Fill this in from laser scanner
    offByVal = 0.0
    minTolerance = 0.1 #Fill this in from config file?
    maxTolerance = 6.0 #Fill this in from config file?
    allowPrint = "disabled"
    toleranceIndicatorVal = "Outside Tolerance"
    toleranceColorVal = "red"

    lbl_workOrderVal = None
    lbl_lengthVal = None
    lbl_toleranceIndicator = None
    lbl_tableLengthBox = None
    lbl_offByBox = None
    lbl_orderLengthBox = None
    
    btn_print = None

    def clear(self):
        print("Clearing Barcodes...")
        self.scannerInput = ""
        self.currentBarcode = ""
        self.updateGUI()
        

    #Call this once a bardcode has been detected and as the laser refreshes. Update the GUI with the new information.
    def updateGUI(self):
        print(f"Updating GUI. Barcode: {self.currentBarcode}; Laser Length: {self.tableLength}")

        self.orderVal = ""
        self.orderLength = 0.0
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.
        #This assumes a Line128 style code. How to better detect what kind of code it is?
        #Simple, check if the first 4 chars are all digits (Line39 has 3 max). If so, it's a Line128 code.
        if (self.currentBarcode[0:4].isdigit()):
            self.orderVal = self.currentBarcode[0:4]
            self.orderLength = float(self.currentBarcode[4:]).__round__(2)
        elif(self.currentBarcode != ""):
            #If we get here and the barcode isn't empty, it's probably a valid Line39 code.
            self.orderLength = float(self.currentBarcode).__round__(2)

        self.offByVal = self.tableLength - self.orderLength

        #Will change between green, yellow, and red based on tolerance, with text changing as well (Within/Near/Outside Tolerance)
        if abs(self.offByVal) <= self.minTolerance and abs(self.offByVal) > 0:
            self.toleranceIndicatorVal = "Within Tolerance"
            self.toleranceColorVal = "green"
            self.allowPrint = "normal"
        elif abs(self.offByVal) <= self.maxTolerance and abs(self.offByVal) > self.minTolerance:
            self.toleranceIndicatorVal = "Near Tolerance"
            self.toleranceColorVal = "yellow"
            self.allowPrint = "disabled"
        else:
            self.toleranceIndicatorVal = "Outside Tolerance"
            self.toleranceColorVal = "red"
            self.allowPrint = "disabled"

        self.lbl_workOrderVal.config(text=self.orderVal)
        self.lbl_lengthVal.config(text=decInchesToFtIn(self.orderLength))
        self.lbl_toleranceIndicator.config(text=self.toleranceIndicatorVal, background=self.toleranceColorVal)
        self.btn_print.configure(state=self.allowPrint)
        self.lbl_tableLengthBox.config(text=decInchesToFtIn(self.tableLength))
        self.lbl_offByBox.config(text=decInchesToFtIn(self.offByVal))
        self.lbl_orderLengthBox.config(text=decInchesToFtIn(self.orderLength))

        self.lbl_workOrderVal.update()
        self.lbl_lengthVal.update()
        self.lbl_toleranceIndicator.update()
        self.btn_print.update()
        self.lbl_tableLengthBox.update()
        self.lbl_offByBox.update()
        self.lbl_orderLengthBox.update()

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
        root.bind('<l>', lambda event: resetLaser())
        root.bind('<space>', lambda event: printLabel(self.allowPrint))
        root.bind('<Key>', self.captureInput)
        
        small_bold_font = font.Font(size=18, weight="bold")
        medium_font = font.Font(size=24)
        large_font = font.Font(size=36)
        xl_font = font.Font(size=48)
        
        lbl_lastBarcode = ttk.Label(content, text="Last Barcode Scanned:", font=small_bold_font)
        lbl_lastBarcode.grid(column=0, row=0, padx=5, pady=5, sticky="W")

        lbl_workOrder = ttk.Label(content, text="Work Order: ", font = medium_font)
        lbl_workOrder.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        self.lbl_workOrderVal = ttk.Label(content, text=self.orderVal, font=medium_font)
        self.lbl_workOrderVal.grid(column=2, row=0, padx=5, pady=5, sticky="W")
        
        lbl_length = ttk.Label(content, text="Length: ", font=medium_font)
        lbl_length.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        self.lbl_lengthVal = ttk.Label(content, text=decInchesToFtIn(self.orderLength), font=medium_font)
        self.lbl_lengthVal.grid(column=2, row=1, padx=5, pady=5, sticky="W")

        btn_clear = ttk.Button(content, text="CLEAR (X)", font=small_bold_font)
        btn_clear.grid(column=3, row=0, rowspan=2)
        btn_clear.bind('<Button-1>', lambda event: self.clear())

        btn_resetLaser = ttk.Button(content, text="RESET LASER (L)", font=small_bold_font)
        btn_resetLaser.grid(column=4, row=0, rowspan=2)
        btn_resetLaser.bind('<Button-1>', lambda event: resetLaser())
        
        lbl_tableLength = ttk.Label(content, text="TABLE LENGTH:", justify="left", font=medium_font)
        lbl_tableLength.grid(column=0, row=2, columnspan=2, padx=5, pady=5, sticky="W")

        self.lbl_tableLengthBox = ttk.Label(content, text=decInchesToFtIn(self.tableLength), justify="left", font=large_font, background="white", relief="solid")
        self.lbl_tableLengthBox.grid(column=0, row=3, columnspan=2, padx=5, pady=5, sticky="W")

        #Could make 3 more labels for the relevant stats...

        lbl_offBy = ttk.Label(content, text="OFF BY:", justify="left", font=medium_font)
        lbl_offBy.grid(column=2, row=2, columnspan=1, padx=5, pady=5, sticky="W")

        self.lbl_offByBox = ttk.Label(content, text=decInchesToFtIn(self.offByVal), justify="left", font=xl_font, background="white", relief="solid")
        self.lbl_offByBox.grid(column=2, row=3, columnspan=1, padx=5, pady=5, sticky="W")

        lbl_orderLength = ttk.Label(content, text="ORDER LENGTH:", justify="left", font=medium_font)
        lbl_orderLength.grid(column=3, row=2, columnspan=2, padx=5, pady=5, sticky="E")

        self.lbl_orderLengthBox = ttk.Label(content, text=decInchesToFtIn(self.orderLength), justify="right", font=xl_font, background="white", relief="solid")
        self.lbl_orderLengthBox.grid(column=3, row=3, columnspan=2, padx=5, pady=5, sticky="E")

        self.lbl_toleranceIndicator = ttk.Label(content, text=self.toleranceIndicatorVal, background=self.toleranceColorVal, font=xl_font)
        self.lbl_toleranceIndicator.grid(column=1, row=4, columnspan=3, padx=5, pady=5, sticky="S")
        
        self.btn_print = ttk.Button(content, text="PRINT\n(spacebar)", font=small_bold_font)
        self.btn_print.grid(column=2, row=5, rowspan=1, padx=5, pady=5, sticky="S")
        self.btn_print.bind("<Button-1>", lambda event: printLabel(self.allowPrint))
        self.btn_print.configure(state=self.allowPrint)
        


root = ttk.Tk()
MainMenu(root)
root.mainloop()