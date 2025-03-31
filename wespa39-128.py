import math
import time
import serial
import win32print
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
    
def metersToDecFt(meters):
    return meters * 3.28084

def parseErrorString(err: str):
    print(f"Parsing error string: {err}")
    match err:
        case "E15": return err + ": Sensor slow to respond"
        case "E16": return err + ": Too much target reflectance"
        case "E17": return err + ": Too much ambient light"
        case "E18": return err + ": DX mode: Measured greater than specified range"
        case "E19": return err + ": DX mode: Target speed > 10m/s"
        case "E23": return err + ": Temp below 14F"
        case "E24": return err + ": Temp above 140F"
        case "E31": return err + ": Faulty memory hardware, EEPROM error"
        case "E51": return err + ": High ambient light or hardware error"
        case "E52": return err + ": Faulty laser diode"
        case "E53": return err + ": EEPROM parameter not set (or divide by zero error)"
        case "E54": return err + ": Hardware error (PLL)"
        case "E55": return err + ": Hardware error"
        case "E61": return err + ": Invalid serial command"
        case "E62": return err + ": Hardware error or Parity error in serial settings"
        case "E63": return err + ": SIO Overflow"
        case "E64": return err + ": Framing - error SIO"

    #Either return nothing or a generic error, depending on how I want to handle it in the GUI
    return ""
        

def printLabel(isEnabled, desiredLength, actualLength, tolerance, offset, workOrder):
    if isEnabled == "normal":
        print("Printing Label...")
        #Format is gonna be something like this; very simple once I get the config right.        
        raw_label = "^XA"
        raw_label += "^CFA,25"
        raw_label += "^FO20,20^FDWO#" + workOrder + ":   " + decInchesToFtIn(desiredLength) + "^FS"
        raw_label += "^FO20,55^FDProduced:  " + decInchesToFtIn(actualLength) + "^FS"
        raw_label += "^FO20,90^FDTolerance: " + decInchesToFtIn(tolerance) + "^FS"
        raw_label += "^FO20,125^FDOff by:   " + decInchesToFtIn(offset) + "^FS"
        raw_label += "^XZ"

        ##Turn this into a formatted string and plop in our own data!
        labelBytes=bytes(raw_label, "utf-8")
        
        #os.startfile(raw_label "print")
        #The only flaw here is that the default printer must be selected in Windows.
        #Eventually I'll add a menu option for printer selection, but get it working first.
        printer = win32print.GetDefaultPrinter()
        try:
            printJob = win32print.StartDocPrinter(printer, 1, ("Label", None, "RAW"))
            try:
                win32print.StartPagePrinter(printJob)
                win32print.WritePrinter(printJob, labelBytes)
                win32print.EndPagePrinter(printJob)
            finally:
                win32print.EndDocPrinter(printJob)
        finally:
            win32print.EndDocPrinter(printJob)
        win32print.ClosePrinter(printer)
        print("Label Printed!")
    return

#Main class for the GUI
class MainMenu:
    scannerInput = "" #Barcode scanner input
    currentBarcode = "" #Last barcode scanned - delimited by newlines with the scanner
    orderVal = "" #First 4 digits of a line128 barcode
    orderLength = 0.0 #line39 code, or the remaining digits of a line128
    tableLength = 0.0 #Fill this in from laser scanner
    offByVal = 0.0

    minTolerance = 0.1 #Fill this in from config file
    maxTolerance = 6.0 #Fill this in from config file
    toleranceIndicatorVal = "Outside Tolerance"
    toleranceColorVal = "red"
    
    allowPrint = "disabled"
    printLabelText = "Cut To Length"
    laserStatusString = ""
    
    laserObject = None #Gets initialized in setupLaser()
    laserComPort = "COM5" #Fill this in from config file

    lbl_workOrderVal = None
    lbl_lengthVal = None
    lbl_toleranceIndicator = None
    lbl_tableLengthBox = None
    lbl_offByBox = None
    lbl_orderLengthBox = None
    lbl_errorCode = None
    
    btn_print = None
    btn_resetLaser = None

    def clear(self):
        print("Clearing Barcodes...")
        self.scannerInput = ""
        self.currentBarcode = ""
        self.updateGUI()
        return

    #Call this once a bardcode has been detected or as the laser refreshes. Update the GUI with the new information.
    #Also run if error codes are detected.
    def updateGUI(self):
        print(f"Updating GUI. Barcode: {self.currentBarcode}; Laser Length: {self.tableLength}")

        self.orderVal = "    "
        self.orderLength = 0.0
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.
        #This assumes a Line128 style code. How to better detect what kind of code it is?
        #Simple, check if the first 4 chars are all digits (Line39 has 3 max). If so, it's a Line128 code.
        if (self.currentBarcode[0:4].isdigit()):
            self.orderVal = self.currentBarcode[0:4]
            self.orderLength = float(self.currentBarcode[4:]).__round__(2)
        elif(self.currentBarcode != ""):
            #If we get here and the barcode isn't empty, it's probably a valid Line39 code.
            #TODO: more proof testing on 128 codes once a proper scanner is available.
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
        self.lbl_errorCode.config(text=self.laserStatusString)

        self.lbl_workOrderVal.update()
        self.lbl_lengthVal.update()
        self.lbl_toleranceIndicator.update()
        self.btn_print.update()
        self.lbl_tableLengthBox.update()
        self.lbl_offByBox.update()
        self.lbl_orderLengthBox.update()
        self.lbl_errorCode.update()

        print(f"Order Length: {self.orderLength}, Order Value: {self.orderVal}")
        return

    def resetLaser(self):
        print("Resetting Laser...")
        ##Send a LF followed by LO after a short delay
        # Try using ascii("LF\n") to send the LF and LO commands if a string literal doesn't work.
        try:
            self.laserObject.write(b'LF\n')
            time.sleep(0.5) #Wait for the laser to reset
            self.laserObject.write(b'LO\n')
            print("Checking for laser response...")
            #Verify the connection, but how?
            # Apparently ID command gets sent on laser auto-start (per setttings dump), some 30 lines.
            # So readlines and see if it lists commands (ID) - really any response other than an error code or timeout (6 seconds).
            rl = ""
            rl = self.laserObject.readline().toString()
            if (rl.startswith("E")):
                self.laserStatusString = parseErrorString(rl)

        except serial.SerialTimeoutException:
            print("Laser reset timed out.")
            self.laserStatusString = "Laser reset failed."

        
        self.updateGUI()
        return

    # Run this after the GUI inits. Establish serial communication.
    def setupLaser(self):
        self.laserObject = serial.Serial(self.laserComPort, baudrate=9600, timeout=10)
        self.laserObject.write_timeout = 1
        if self.laserObject.isOpen():
            print()
            self.laserStatusString = "Laser connected on " + self.laserComPort
        else:
            self.laserStatusString = "Laser not connected - check connection and configuration."

        return

    def getLaserLength(self):
        #Manual override key is g - ideally we do this every half second or so.
        try:
            self.laserObject.write(b'DM\n') #Send the command to get the length
            time.sleep(0.5) #Wait for the laser to respond
            re = self.laserObject.readline()
            print(f"Laser response: {re}")
            #Laser is by default configured to return in meters, so convert to decimal feet-inches.
            self.tableLength = metersToDecFt(float(re.decode('utf-8').strip()))
        except serial.SerialTimeoutException:
            print("Laser read timed out.")
            self.laserStatusString = "Laser read failed."

        self.updateGUI()
        return 


    #Deals with keyboard input from the barcode scanner.
    def captureInput(self, event):
        if event.keysym == 'Return':
            print(f"Received input: {self.scannerInput}")
            self.currentBarcode = self.scannerInput
            self.scannerInput = ""  # Clear the input after processing
            self.updateGUI()
        elif (event.char >= '0' and event.char <= '9' or event.char == '.'):
            self.scannerInput += event.char  # Append the character to the input string

    ##Set up GUI elements here.
    def __init__(self, root):
        content = ttk.Frame(root, width=900, height=600)
        content.grid(column=0, row=0, columnspan=5, rowspan=7)

        for i in range (0, 5):
            root.columnconfigure(i, weight=3)

        for i in range (0, 7):
            root.rowconfigure(i, weight=2)

        root.resizable(width=False, height=False)
        root.title("WESPA 39-128")

        # Bind keyboard shortcuts; also detect barcode input
        root.bind('<x>', lambda event: self.clear())
        root.bind('<l>', lambda event: self.resetLaser())
        root.bind('<g>', lambda event: self.getLaserLength())
        root.bind('<space>', lambda event: printLabel(self.allowPrint))
        root.bind('<Key>', self.captureInput)
        
        smaller_font = font.Font(size=14)
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

        self.btn_resetLaser = ttk.Button(content, text="RESET LASER (L)", font=small_bold_font)
        self.btn_resetLaser.grid(column=4, row=0, rowspan=2)
        self.btn_resetLaser.bind('<Button-1>', lambda event: self.resetLaser())
        
        lbl_tableLength = ttk.Label(content, text="TABLE LENGTH:", justify="left", font=medium_font)
        lbl_tableLength.grid(column=0, row=2, columnspan=2, padx=50, pady=5)

        self.lbl_tableLengthBox = ttk.Label(content, text=decInchesToFtIn(self.tableLength), justify="left", width=12, font=large_font, background="white", relief="solid")
        self.lbl_tableLengthBox.grid(column=0, row=3, columnspan=2, padx=5, pady=5)

        lbl_offBy = ttk.Label(content, text="OFF BY:", justify="center", font=medium_font)
        lbl_offBy.grid(column=2, row=2, columnspan=1, padx=5, pady=5)

        self.lbl_offByBox = ttk.Label(content, text=decInchesToFtIn(self.offByVal), justify="left", width=12, font=large_font, background="white", relief="solid")
        self.lbl_offByBox.grid(column=2, row=3, columnspan=1, padx=5, pady=5,sticky="W") 

        lbl_orderLength = ttk.Label(content, text="ORDER LENGTH:", justify="center", font=medium_font)
        lbl_orderLength.grid(column=3, row=2, columnspan=2, padx=5, pady=5)

        self.lbl_orderLengthBox = ttk.Label(content, text=decInchesToFtIn(self.orderLength), justify="right", width=12, font=large_font, background="white", relief="solid")
        self.lbl_orderLengthBox.grid(column=3, row=3, columnspan=2, padx=5, pady=5, sticky="E")

        self.lbl_toleranceIndicator = ttk.Label(content, text=self.toleranceIndicatorVal, background=self.toleranceColorVal, font=xl_font)
        self.lbl_toleranceIndicator.grid(column=1, row=4, columnspan=3, padx=5, pady=5)
        
        self.btn_print = ttk.Button(content, text="PRINT\n(spacebar)", font=small_bold_font)
        self.btn_print.grid(column=2, row=5, rowspan=1, padx=5, pady=5)
        self.btn_print.bind("<Button-1>", lambda event: printLabel(self.allowPrint, self.orderLength, self.tableLength, self.offByVal, self.minTolerance, self.orderVal))
        self.btn_print.configure(state=self.allowPrint)

        self.setupLaser()

        self.lbl_errorCode = ttk.Label(content, text=self.laserStatusString, font=smaller_font)
        self.lbl_errorCode.grid(column=0, row=6, columnspan=1, padx=5, pady=5)
        


root = ttk.Tk()
MainMenu(root)
root.mainloop()