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
    #win32print for printer handling

#Takes a float (dec_inches) and returns string formatted as XXft YYin or YYin if no feet
def inchesToStr(dec_inches: float):
    feet = math.trunc(dec_inches/12)
    inches = format(dec_inches - feet * 12, '.2f')

    if (feet > 0):
        return "{0} FT {1} IN".format(feet, inches)
    else:   
        return "{0} IN".format(inches)
    
def metersToInches(meters: float):
    return meters * 39.3701

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
        case "LO": return err + ": Laser is on"
        case "LF": return err + ": Laser is off"
        case '': return "No response from laser."

    #Either return nothing or a generic error, depending on how I want to handle it in the GUI
    return ""     

def printLabel(isEnabled: str, orderLength: float, tableLength: float, tolerance: float, offset: float, workOrder: str):
    #https://timgolden.me.uk/python/win32_how_do_i/print.htm
    if isEnabled == "normal":
        print("Printing Label...")
        
        raw_label = "^XA"
        raw_label += "^CFA,20" #TODO: Make font size configurable from GUI to find sweet spot.
        raw_label += "^FO0,90^FDWO#" + workOrder + ":   " + inchesToStr(orderLength) + "^FS"
        raw_label += "^FO0,110^FDProduced:  " + inchesToStr(tableLength) + "^FS"
        raw_label += "^FO0,130^FDTolerance: " + inchesToStr(tolerance) + "^FS"
        raw_label += "^FO0,150^FDOff by:    " + inchesToStr(offset) + "^FS"
        raw_label += "^XZ"

        ##Turn this into a formatted string and plop in our own data!
        labelBytes=bytes(raw_label, "utf-8")
        print("Label Bytes: " + str(labelBytes))
        #The only flaw here is that the default printer must be selected in Windows.
        #TODO: add GUI option for printer selection
        defaultPrinter = win32print.GetDefaultPrinter()
        myPrinter = win32print.OpenPrinter(defaultPrinter)
        print("Default Printer: " + defaultPrinter)
        try:
            print("Starting print job...")
            printJob = win32print.StartDocPrinter(myPrinter, 1, ("label", None, "RAW"))
            print("Starting page...")
            win32print.StartPagePrinter(myPrinter)
            print("Writing label bytes...")
            win32print.WritePrinter(myPrinter, labelBytes)
            print("Ending page...")
            win32print.EndPagePrinter(myPrinter)
            print("Ending print job...")
            win32print.EndDocPrinter(myPrinter)
        except Exception as e:
            print(f"Error printing label: {e}")
        finally:
            print("Finally, Closing printer...")
            win32print.ClosePrinter(myPrinter)

#Main class for the GUI
class MainMenu:
    scannerInput: str = "" #Barcode scanner input
    currentBarcode: str = "" #Last barcode scanned - delimited by newlines with the scanner
    orderVal: str = "" #First 4 digits of a line128 barcode
    orderLength: float = 0.0 #line39 code, or the remaining digits of a line128
    tableLength: float = 0.0 #Fill this in from laser scanner
    offByVal: float = 0.0

    minTolerance: float = 0.1 #Fill this in from config file
    maxTolerance: float = 6.0 #Fill this in from config file
    toleranceIndicatorVal: str = "Outside Tolerance"
    toleranceColorVal: str = "red" #red/yellow/green
    
    allowPrint: str = "normal" #normal/disabled
    printLabelText: str = "Cut To Length"
    laserStatusString: str = ""
    
    laserObject: serial.Serial = None #Gets initialized in setupLaser()
    laserIsConnected: bool = False #True if the laser is connected, false if not.
    laserComPort: str = "COM3" #Fill this in from config file

    lbl_workOrderVal: ttk.Label = None
    lbl_lengthVal: ttk.Label = None
    lbl_toleranceIndicator: ttk.Label = None
    lbl_tableLengthBox: ttk.Label = None
    lbl_offByBox: ttk.Label = None
    lbl_orderLengthBox: ttk.Label = None
    lbl_errorCode: ttk.Label = None
    
    btn_print: ttk.Button = None
    btn_resetLaser: ttk.Button = None

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
            #If we get here and the barcode isn't empty, it's probably a Line39 code.
            #TODO: more proof testing on 128 codes once a proper scanner is available.
            self.orderLength = float(self.currentBarcode).__round__(2)
        else:
            #currentBarcode is empty or wrong format.
            self.orderVal = "    "
            self.orderLength = 0.0

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
        self.lbl_lengthVal.config(text=inchesToStr(self.orderLength))
        self.lbl_toleranceIndicator.config(text=self.toleranceIndicatorVal, background=self.toleranceColorVal)
        self.btn_print.configure(state=self.allowPrint)
        self.lbl_tableLengthBox.config(text=inchesToStr(self.tableLength))
        self.lbl_offByBox.config(text=inchesToStr(self.offByVal))
        self.lbl_orderLengthBox.config(text=inchesToStr(self.orderLength))
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

    def tweakLaserOffset(self, newOffset: float):

        return

    #Send an off / on signal to the laser, or try to reconnect if it's not connected.
    def resetLaser(self):
        if (self.laserIsConnected is False):
            print("Laser not connected. Attempting to reconnect...")
            self.laserObject.close() #Close the serial port if it's open
            self.setupLaser()
            return

        print("Resetting Laser...")
        ##Send a LF followed by LO after a short delay
        # Try using ascii("LF\n") to send the LF and LO commands if a string literal doesn't work.
        try:
            print("Writing LF (laser off)")
            self.laserObject.write(b'LF\r\n')
            print("Checking laser response...")
            rl = str(self.laserObject.readline()).strip()
            self.laserStatusString = parseErrorString(rl)
            print(f"Laser response: {rl}")
            time.sleep(1) #Wait for the laser to reset
            self.laserObject.flush()
            print("Writing LO (laser on)")
            self.laserObject.write(b'LO\r\n')
            print("Checking laser response...")
            rl = str(self.laserObject.readline()).strip()
            print(f"Laser response: {rl}")
            self.laserStatusString = parseErrorString(rl)
        except serial.SerialTimeoutException:
            print("Laser reset timed out.")
            self.laserStatusString = "Laser offline."
            self.laserIsConnected = False

        print("Flushing buffer...")
        self.laserObject.flush() #Clear the input buffer to avoid reading old data
        self.updateGUI()
        return

    # Run this after the GUI inits. Establish serial communication.
    def setupLaser(self):
        try:
            self.laserObject = serial.Serial(self.laserComPort, baudrate=9600, timeout=3, write_timeout=3)
            self.laserObject.write(b'ID\r\n') #Send the ID command to check the connection
            time.sleep(0.5) #Wait for the laser to respond
            re = self.laserObject.readlines()
            if (re is None or len(re) == 0):
                raise serial.SerialTimeoutException("No response from laser.")

            self.laserStatusString = "Laser connected on " + self.laserComPort
            print("Laser connected!")
            self.laserIsConnected = True
        except serial.SerialException as e:
            self.laserStatusString = "Laser not found on " + self.laserComPort + " - check connection and configuration."
            print(f"Serial exception: {e}")
        except serial.SerialTimeoutException:
            self.laserStatusString = "Laser connection on " + self.laserComPort + " timed out."
            print(f"Laser read timed out: {e}")

    #TODO: Make this read on a regular timer instead of a button press.
    def getLaserLength(self):
        if (self.laserObject.is_open is False):
            self.laserStatusString = "Laser not connected."
            print("Laser not connected.")
            return
        
        #Manual override key is g - ideally we do this every half second or so, or set continuous read mode on init.
        print("Getting laser length (DM)...")
        #TODO: Consider using the more precise DS command instead of DM (note that this is not instant, but is faster than DT, which can take up to 6 seconds).
        #Whatever I choose will need to take this timing into account - the laser currently has an ST of 0 (no limit).
        try:
            self.laserObject.write(b'DM\n') #Send the command to get the length
            time.sleep(0.5) #Wait for the laser to respond
            re = self.laserObject.readline()
            print(f"Laser response: {re}")
            self.tableLength = metersToInches(float(re.decode('utf-8').strip()))
        except serial.SerialTimeoutException:
            print("Laser read timed out.")
            self.laserStatusString = "Laser offline."
            self.laserIsConnected = False
        except ValueError:
            self.laserStatusString = parseErrorString(str(re).strip())
            print("Non-numeric value received from laser.")
            self.tableLength = 0.0

        print("Flushing buffer...")
        self.laserObject.flush() #Clear the input buffer to avoid reading old data
        self.updateGUI()
        return

    #Deals with keyboard input from the barcode scanner.
    #Side effect of this is that it allows for manual input of the barcode scanner, which is actually a desired feature.
    def captureInput(self, event):
        if event.keysym == 'Return':
            print(f"Received input: {self.scannerInput}")
            self.currentBarcode = self.scannerInput
            self.scannerInput = ""  # Clear the input after processing
            self.updateGUI()
        elif (event.char >= '0' and event.char <= '9' or event.char == '.'):
            self.scannerInput += event.char  # Append the character to the input string

    #TODO: Make GUI scalable
    def __init__(self, root):
        print("Initializing GUI...")
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
        root.bind('<space>', lambda event: printLabel(
            isEnabled=self.allowPrint, orderLength=self.orderLength,
              tableLength=self.tableLength, offset=self.offByVal,
                tolerance=self.minTolerance, workOrder=self.orderVal))
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

        self.lbl_lengthVal = ttk.Label(content, text=inchesToStr(self.orderLength), font=medium_font)
        self.lbl_lengthVal.grid(column=2, row=1, padx=5, pady=5, sticky="W")

        btn_clear = ttk.Button(content, text="CLEAR (X)", font=small_bold_font)
        btn_clear.grid(column=3, row=0, rowspan=2)
        btn_clear.bind('<Button-1>', lambda event: self.clear())

        self.btn_resetLaser = ttk.Button(content, text="RESET LASER (L)", font=small_bold_font)
        self.btn_resetLaser.grid(column=4, row=0, rowspan=2)
        self.btn_resetLaser.bind('<Button-1>', lambda event: self.resetLaser())
        
        lbl_tableLength = ttk.Label(content, text="TABLE LENGTH:", justify="left", font=medium_font)
        lbl_tableLength.grid(column=0, row=2, columnspan=2, padx=50, pady=5)

        self.lbl_tableLengthBox = ttk.Label(content, text=inchesToStr(self.tableLength), justify="left", width=12, font=large_font, background="white", relief="solid")
        self.lbl_tableLengthBox.grid(column=0, row=3, columnspan=2, padx=5, pady=5)

        lbl_offBy = ttk.Label(content, text="OFF BY:", justify="center", font=medium_font)
        lbl_offBy.grid(column=2, row=2, columnspan=1, padx=5, pady=5)

        self.lbl_offByBox = ttk.Label(content, text=inchesToStr(self.offByVal), justify="left", width=12, font=large_font, background="white", relief="solid")
        self.lbl_offByBox.grid(column=2, row=3, columnspan=1, padx=5, pady=5,sticky="W") 

        lbl_orderLength = ttk.Label(content, text="ORDER LENGTH:", justify="center", font=medium_font)
        lbl_orderLength.grid(column=3, row=2, columnspan=2, padx=5, pady=5)

        self.lbl_orderLengthBox = ttk.Label(content, text=inchesToStr(self.orderLength), justify="right", width=12, font=large_font, background="white", relief="solid")
        self.lbl_orderLengthBox.grid(column=3, row=3, columnspan=2, padx=5, pady=5, sticky="E")

        self.lbl_toleranceIndicator = ttk.Label(content, text=self.toleranceIndicatorVal, background=self.toleranceColorVal, font=xl_font)
        self.lbl_toleranceIndicator.grid(column=1, row=4, columnspan=3, padx=5, pady=5)
        
        self.btn_print = ttk.Button(content, text="PRINT\n(spacebar)", font=small_bold_font)
        self.btn_print.grid(column=2, row=5, rowspan=1, padx=5, pady=5)
        self.btn_print.bind("<Button-1>", lambda event: printLabel(
            isEnabled=self.allowPrint, orderLength=self.orderLength,
              tableLength=self.tableLength, offset=self.offByVal,
                tolerance=self.minTolerance, workOrder=self.orderVal))
        self.btn_print.configure(state=self.allowPrint)

        self.setupLaser()

        self.lbl_errorCode = ttk.Label(content, text=self.laserStatusString, font=smaller_font)
        self.lbl_errorCode.grid(column=0, row=6, columnspan=1, padx=5, pady=5)

        print("GUI Initialized!")
        


root = ttk.Tk()
MainMenu(root)
root.mainloop()
