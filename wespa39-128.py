import math
import time
from time import strftime
from configparser import ConfigParser
import tkinter as ttk
from tkinter import font

import serial
import win32print

#Major parts of the program:
#Inputs:
    #Barcode scanner, connected via USB, keyboard-like input in Line39 or Line128 format.
    # Line39 currently, simply the length
    # Line128 may be a future expansion with the Work Order as the first 4 digits,
    #  length as remainder.
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
def get_inches_str(dec_inches: float):
    feet = abs(math.trunc(dec_inches/12))
    inches = format(abs(dec_inches) - feet * 12, '.2f')

    if feet > 0:
        return "{0} FT {1} IN".format(feet, inches)
    else:
        return "{0} IN".format(inches)
    
def meters_to_inches(meters: float):
    return meters * 39.3701

def parse_laser_error(err: str):
    print(strftime('%G-%m-%d %H:%M:%S') + " Parsing error string: " + err)
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

#Send ZPL code to the default system printer along with the data to print.
def send_print_label(is_enabled: str, order_length: float, table_length: float,
                      tolerance: float, offset: float, work_order: str):
    #https://timgolden.me.uk/python/win32_how_do_i/print.htm
    if is_enabled == "normal":
        print(strftime('%G-%m-%d %H:%M:%S') + " Printing Label...")
        
        raw_label = "^XA"
        raw_label += "^CFA,20"
        raw_label += "^FO0,90^FDWO#" + work_order + ":   " + get_inches_str(order_length) + "^FS"
        raw_label += "^FO0,110^FDProduced:  " + get_inches_str(table_length) + "^FS"
        raw_label += "^FO0,130^FDTolerance: " + get_inches_str(tolerance) + "^FS"
        raw_label += "^FO0,150^FDOff by:    " + get_inches_str(offset) + "^FS"
        raw_label += "^XZ"

        ##Turn this into a formatted string and plop in our own data!
        label_bytes=bytes(raw_label, "utf-8")
        print("Label Bytes: " + str(label_bytes))
        #The only flaw here is that the default printer must be selected in Windows.
        #TODO: add GUI option for printer selection
        default_printer = win32print.GetDefaultPrinter()
        my_printer = win32print.OpenPrinter(default_printer)
        print("Default Printer: " + default_printer)
        try:
            print("Starting print job...")
            win32print.StartDocPrinter(my_printer, 1, ("label", None, "RAW"))
            print("Starting page...")
            win32print.StartPagePrinter(my_printer)
            print("Writing label bytes...")
            win32print.WritePrinter(my_printer, label_bytes)
            print("Ending page...")
            win32print.EndPagePrinter(my_printer)
            print("Ending print job...")
            win32print.EndDocPrinter(my_printer)
        except Exception as e:
            print(f"Error printing label: {e}")
        finally:
            print(strftime('%G-%m-%d %H:%M:%S') + " Closing printer...")
            win32print.ClosePrinter(my_printer)

#Main class for the GUI
class MainMenu:
    scanner_input: str = "" #Barcode scanner input
    current_barcode: str = "" #Last barcode scanned - delimited by newlines with the scanner
    order_str: str = "" #First 4 digits of a line128 barcode
    order_length: float = 0.0 #line39 code, or the remaining digits of a line128
    table_length: float = 0.0 #Fill this in from laser scanner
    off_by_val: float = 0.0
    laser_offset: float = 0.0 #Fill this in from config file
    min_tolerance: float = 0.1 #Fill this in from config file
    max_tolerance: float = 6.0 #Fill this in from config file
    tolerance_indicator: str = "Outside Tolerance"
    tolerance_color: str = "red" #red/yellow/green
    
    allow_print: str = "disabled" #normal/disabled
    print_text: str = "Cut To Length"
    laser_status: str = ""
    
    laser_object: serial.Serial = None #Gets initialized in setupLaser()
    laser_is_connected: bool = False #True if the laser is connected, false if not.
    laser_port: str = "COM3" #Fill this in from config file

    lbl_order: ttk.Label = None
    lbl_length: ttk.Label = None
    lbl_tolerance_indicator: ttk.Label = None
    lbl_table_length_box: ttk.Label = None
    lbl_off_by_box: ttk.Label = None
    lbl_order_length_box: ttk.Label = None
    lbl_error_code: ttk.Label = None
    
    btn_print: ttk.Button = None
    btn_laser_reset: ttk.Button = None

    def clear(self):
        print(strftime('%G-%m-%d %H:%M:%S') + " Clearing Barcodes...")
        self.scanner_input = ""
        self.current_barcode = ""
        self.update()

    #Call this once a bardcode has been detected or as the laser refreshes. Update the GUI with the new information.
    #Also run if error codes are detected.
    def update(self):
        print(strftime('%G-%m-%d %H:%M:%S') +
               f" Updating GUI. Barcode: {self.current_barcode}; Laser Length: {self.table_length}")

        self.order_str = "    "
        self.order_length = 0.0
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.
        #This assumes a Line128 style code. How to better detect what kind of code it is?
        #Simple, check if the first 4 chars are all digits (Line39 has 3 max). If so, it's a Line128 code.
        if (len(self.current_barcode) > 4 and self.current_barcode[0:4].isdigit()):
            self.order_str = self.current_barcode[0:4]
            try:
                self.order_length = round(float(self.current_barcode[4:]), 2)
            except ValueError:
                print(strftime('%G-%m-%d %H:%M:%S')
                       + f" ValueError: Could not convert {self.current_barcode[4:]} to float.")
        elif(self.current_barcode is not None and len(self.current_barcode) > 0):
            #If we get here and the barcode isn't empty, it's probably a Line39 code.
            try:
                self.order_length = round(float(self.current_barcode), 2)
            except ValueError:
                print(strftime('%G-%m-%d %H:%M:%S')
                       + f" ValueError: Could not convert {self.current_barcode} to float.")
        else:
            #currentBarcode is empty or wrong format.
            self.order_str = "    "
            self.order_length = 0.0

        self.off_by_val = abs((self.table_length + self.laser_offset) - self.order_length)

        tolerance_position: str = ""
        if (self.order_length < self.table_length + self.laser_offset):
            tolerance_position = ": Too Long"
        elif (self.order_length > self.table_length + self.laser_offset):
            tolerance_position = ": Too Short"

        #Will change between green, yellow, and red based on tolerance, with text changing as well (Within/Near/Outside Tolerance)
        if abs(self.off_by_val) <= self.min_tolerance and self.off_by_val >= 0:
            self.tolerance_indicator = "Within Tolerance"
            self.tolerance_color = "green"
            self.allow_print = "normal"
        elif abs(self.off_by_val) <= self.max_tolerance and self.off_by_val > self.min_tolerance:
            self.tolerance_indicator = "Near Tolerance" + tolerance_position
            self.tolerance_color = "yellow"
            self.allow_print = "disabled"
        else:
            self.tolerance_indicator = "Outside Tolerance" + tolerance_position
            self.tolerance_color = "red"
            self.allow_print = "disabled"

        self.lbl_order.config(text=self.order_str)
        self.lbl_length.config(text=get_inches_str(self.order_length))
        self.lbl_tolerance_indicator.config(text=self.tolerance_indicator, background=self.tolerance_color)
        self.btn_print.configure(state=self.allow_print)
        self.lbl_table_length_box.config(text=get_inches_str(self.table_length + self.laser_offset))
        self.lbl_off_by_box.config(text=get_inches_str(self.off_by_val))
        self.lbl_order_length_box.config(text=get_inches_str(self.order_length + self.laser_offset))
        self.lbl_error_code.config(text=self.laser_status)

        self.lbl_order.update()
        self.lbl_length.update()
        self.lbl_tolerance_indicator.update()
        self.btn_print.update()
        self.lbl_table_length_box.update()
        self.lbl_off_by_box.update()
        self.lbl_order_length_box.update()
        self.lbl_error_code.update()

        print(f"Order Length: {self.order_length}, Order Number: {self.order_str}, Table Length: {self.table_length + self.laser_offset}, Off By: {self.off_by_val}")

    #Send an off / on signal to the laser, or try to reconnect if it's not connected.
    def reset_laser(self):
        if (not self.laser_is_connected):
            print("Laser not connected. Attempting to reconnect...")
            if (self.laser_object is not None):
                self.laser_object.close() #Close the serial port if it's open
            self.setup_laser()
            return

        print(strftime('%G-%m-%d %H:%M:%S') + " Resetting Laser...")
        ##Send a LF followed by LO after a short delay
        # Try using ascii("LF\n") to send the LF and LO commands if a string literal doesn't work.
        try:
            print("Writing LF (laser off)")
            self.laser_object.write(b'LF\r\n')
            print("Checking laser response...")
            rl = str(self.laser_object.readline()).strip()
            self.laser_status = parse_laser_error(rl)
            print(f"Laser response: {rl}")
            time.sleep(1) #Wait for the laser to reset
            self.laser_object.flush()
            print("Writing LO (laser on)")
            self.laser_object.write(b'LO\r\n')
            print("Checking laser response...")
            rl = str(self.laser_object.readline()).strip()
            print(f"Laser response: {rl}")
            self.laser_status = parse_laser_error(rl)
        except serial.SerialTimeoutException:
            print("Laser reset timed out.")
            self.laser_status = "Laser offline."
            self.laser_is_connected = False
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            print(f"Unhandled Exception: {e}")

        print("Flushing buffer...")
        self.laser_object.flush() #Clear the input buffer to avoid reading old data
        self.update()

        #Call ReadLaserLength() here again since connection has been refreshed?

    # Run this after the GUI inits. Establish serial communication.
    def setup_laser(self):
        try:
            self.laser_object = serial.Serial(self.laser_port, baudrate=9600, timeout=3, write_timeout=3)
            self.laser_object.write(b'ID\r\n') #Send the ID command to check the connection
            time.sleep(0.5) #Wait for the laser to respond
            re = self.laser_object.readlines()
            if (re is None or len(re) == 0):
                raise serial.SerialTimeoutException("No response from laser.")

            self.laser_status = "Laser connected on " + self.laser_port
            print(strftime('%G-%m-%d %H:%M:%S') + " Laser connected!")
            self.laser_is_connected = True
        except serial.SerialException as e:
            self.laser_status = "Laser not found on " + self.laser_port + " - check connection and configuration."
            print(strftime('%G-%m-%d %H:%M:%S') + f" Serial exception: {e}")
        except serial.SerialTimeoutException:
            self.laser_status = "Laser connection on " + self.laser_port + " timed out."
            print(strftime('%G-%m-%d %H:%M:%S') + f" Laser read timed out: {e}")
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            print(strftime('%G-%m-%d %H:%M:%S') + f" Unhandled Exception: {e}")

    def getLaserLength(self, root):
        if (not self.laser_is_connected):
            self.laser_status = "Laser not connected."
            print(strftime('%G-%m-%d %H:%M:%S') + " Laser not connected.")
            #Note that this will prevent the getLaserLength function from running again.
            #Use the reset button to reconnect, then manually override with g to get it going again.
            #Hopefully that won't be necessary very often...or at all.
            return
        
        print(strftime('%G-%m-%d %H:%M:%S') +" Getting laser length (DM)...")
        #Consider using the more precise DS command instead of DM
        # Note that this is not instant, but is faster than DT, which can take up to 6 seconds.
        #Whatever I choose will need to take this timing into account
        # The laser currently has an ST of 0 (no limit).
        try:
            self.laser_object.write(b'DM\n') #Send the command to get the length
            print(strftime('%G-%m-%d %H:%M:%S') + " Waiting for laser response...")
            time.sleep(0.5) #Wait for the laser to respond
            re = self.laser_object.readline()
            print(strftime('%G-%m-%d %H:%M:%S') + f" Laser response: {re}")
            self.table_length = meters_to_inches(float(re.decode('utf-8').strip()))
        except serial.SerialTimeoutException:
            print(strftime('%G-%m-%d %H:%M:%S') + " Laser read timed out.")
            self.laser_status = "Laser offline."
            self.laser_is_connected = False
        except ValueError:
            self.laser_status = parse_laser_error(str(re).strip())
            print(strftime('%G-%m-%d %H:%M:%S') +" Non-numeric value received from laser.")
            self.table_length = 0.0
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            print(strftime('%G-%m-%d %H:%M:%S') + f" Unhandled Exception: {e}")
            self.table_length = 0.0

        print("Flushing buffer...")
        self.laser_object.flush() #Clear the input buffer to avoid reading old data
        self.update()

        root.after(500, self.getLaserLength(root)) #Call this function again after 500ms


    #Deals with keyboard input from the barcode scanner.
    #Side effect of this is that it allows for manual input of the barcode scanner, which is actually a desired feature.
    def captureInput(self, event):
        if event.keysym == 'Return':
            print(strftime('%G-%m-%d %H:%M:%S') + f" Received input: {self.scanner_input}")
            self.current_barcode = self.scanner_input
            self.scanner_input = ""  # Clear the input after processing
            self.update()
        elif (event.char >= '0' and event.char <= '9' or event.char == '.'):
            self.scanner_input += event.char  # Append the character to the input string

    def load_debug_vals(self):
        self.order_str = "DBUG"
        self.laser_offset = 0.42
        self.order_length = 128.0
        self.table_length = 256.0
        self.off_by_val = self.table_length + self.laser_offset - self.order_length

    
    def read_config_file(self):
        c = ConfigParser()
        try:
            c.read('wespa39-128.ini')
            self.laser_port = c.get('ports', 'laserComPort')
            self.laser_offset = c.getfloat('offsets', 'laserOffset')
            self.min_tolerance = c.getfloat('offsets', 'minTolerance')
            self.max_tolerance = c.getfloat('offsets', 'maxTolerance')
        except Exception as e:
            print(strftime('%G-%m-%d %H:%M:%S') + f" Error reading config file: {e}")
            print(strftime('%G-%m-%d %H:%M:%S') + " Using default values.")

        print(strftime('%G-%m-%d %H:%M:%S') + " Config file loaded.")


    def __init__(self, root: ttk.Tk):
        print(strftime('%G-%m-%d %H:%M:%S') + " Loading config file...")
        self.read_config_file()

        print(strftime('%G-%m-%d %H:%M:%S') + " Initializing GUI...")
        #root is already initialized as ttk.Tk()
        root.resizable(True, True)
        root.title("WESPA 39-128")
        
        #Number of columns and rows in the grid - all resize at the same rate
        for i in range(3):
            root.columnconfigure(i, weight=1)
        for i in range(7):
            root.rowconfigure(i, weight=1)

        #self.loadDebugVals()

        # Bind keyboard shortcuts; also detect barcode input
        root.bind('<x>', lambda event: self.clear())
        root.bind('<l>', lambda event: self.reset_laser())
        root.bind('<g>', lambda event: self.getLaserLength(root))
        root.bind('<space>', lambda event: send_print_label(
            is_enabled=self.allow_print, order_length=self.order_length,
              table_length=self.table_length, offset=self.off_by_val,
                tolerance=self.min_tolerance, work_order=self.order_str))
        root.bind('<Key>', self.captureInput)
        
        font_baseSize = 12
        smallest_font = font.Font(size=font_baseSize)
        small_bold_font = font.Font(size=font_baseSize+4, weight="bold")
        medium_bold_font = font.Font(size=font_baseSize+8, weight="bold")
        large_bold_font = font.Font(size=font_baseSize+16, weight="bold")
        
        lbl_last_barcode = ttk.Label(root, text="Last Barcode Scanned:", justify="left", font=smallest_font)
        lbl_last_barcode.grid(column=0, row=0, padx=5, pady=5, sticky="nw")

        #WO number label
        lbl_work_order = ttk.Label(root, text="Work Order: ", justify="left", font=small_bold_font)
        lbl_work_order.grid(column=1, row=0, padx=5, pady=5, sticky="ne")

        #WO number textbox
        self.lbl_order = ttk.Label(root, text=self.order_str, justify="left", font=small_bold_font)
        self.lbl_order.grid(column=2, row=0, padx=5, pady=5, sticky="nw")
        
        #Length label (last scanned)
        lbl_length = ttk.Label(root, text="Length: ", justify="left", font=small_bold_font)
        lbl_length.grid(column=1, row=1, padx=5, pady=5, sticky="ne")

        #Length textbox (last scanned)
        self.lbl_length = ttk.Label(root, text=get_inches_str(self.order_length), justify="left", font=small_bold_font)
        self.lbl_length.grid(column=2, row=1, padx=5, pady=5, sticky="nw")

        #Table Length label
        lbl_table_length = ttk.Label(root, text="TABLE LENGTH:", justify="left", font=medium_bold_font)
        lbl_table_length.grid(column=0, row=2, padx=25, pady=5, sticky="nsew")

        #Table Length textbox
        self.lbl_table_length_box = ttk.Label(root, text=get_inches_str(self.table_length), justify="center", background="white", relief="solid", font=medium_bold_font)
        self.lbl_table_length_box.grid(column=0, row=3, padx=5, pady=5, sticky="nsew")

        #OffBy Label
        lbl_off_by = ttk.Label(root, text="OFF BY:", justify="center", font=medium_bold_font)
        lbl_off_by.grid(column=1, row=2, padx=25, pady=5, sticky="nsew")

        #OffBy Textbox
        self.lbl_off_by_box = ttk.Label(root, text=get_inches_str(self.off_by_val), background="white", relief="solid", font=medium_bold_font)
        self.lbl_off_by_box.grid(column=1, row=3, padx=5, pady=5, sticky="nsew")

        #Order Length Label
        lbl_order_length = ttk.Label(root, text="ORDER LENGTH:", justify="left", font=medium_bold_font)
        lbl_order_length.grid(column=2, row=2, padx=25, pady=5, sticky="nsew")

        #Order Length Textbox
        self.lbl_order_length_box = ttk.Label(root, text=get_inches_str(self.order_length), justify="right", background="white", relief="solid", font=medium_bold_font)
        self.lbl_order_length_box.grid(column=2, row=3, padx=5, pady=5, sticky="nsew")

        #Tolerance Indicator
        self.lbl_tolerance_indicator = ttk.Label(root, text=self.tolerance_indicator, background=self.tolerance_color, font=large_bold_font)
        self.lbl_tolerance_indicator.grid(column=0, row=4, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        #Clear button
        btn_clear = ttk.Button(root, text="CLEAR\n(X)", font=medium_bold_font)
        btn_clear.grid(column=0, row=5, padx=5, pady=5)
        btn_clear.bind('<Button-1>', lambda event: self.clear())
        
        #Print button
        self.btn_print = ttk.Button(root, text="PRINT\n(space)", font=medium_bold_font)
        self.btn_print.grid(column=1, row=5, padx=5, pady=5)
        self.btn_print.bind("<Button-1>", lambda event: send_print_label(
            is_enabled=self.allow_print, order_length=self.order_length,
              table_length=self.table_length, offset=self.off_by_val,
                tolerance=self.min_tolerance, work_order=self.order_str))
        self.btn_print.configure(state=self.allow_print)

        #Reset/reconnect button
        self.btn_laser_reset = ttk.Button(root, text="RESET LASER\n(L)", font=medium_bold_font)
        self.btn_laser_reset.grid(column=2, row=5, padx=5, pady=5)
        self.btn_laser_reset.bind('<Button-1>', lambda event: self.reset_laser())

        self.setup_laser()

        #Laser status
        self.lbl_error_code = ttk.Label(root, text=self.laser_status, justify="left", font=smallest_font)
        self.lbl_error_code.grid(column=0, row=6, columnspan=3, padx=5, pady=5, sticky="w")

        print(strftime('%G-%m-%d %H:%M:%S') + " GUI Initialized!")

        root.after(1000, self.getLaserLength(root))
        

root = ttk.Tk()
MainMenu(root)
root.mainloop()
