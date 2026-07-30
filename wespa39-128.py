import math
import logging
import os
import sys
import time
import tkinter as ttk
import serial
import win32print

from configparser import ConfigParser
from tkinter import font

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

#Main class for the GUI
class MainMenu(ttk.Tk):
    scanner_input: str = "" #Barcode scanner input
    current_barcode: str = "" #Last barcode scanned - delimited by newlines with the scanner
    order_str: str = "" #First 4 digits of a line128 barcode
    order_length: float = 0.0 #line39 code, or the remaining digits of a line128
    laser_length: float = 0.0 #Raw measurement from laser scanner
    order_difference: float = 0.0 #Laser Length + Laser Offset - Order Length
    laser_offset: float = 0.0 #Fill this in from config file; adjusts laser length
    adjusted_length: float = 0.0 #Laser Length + Laser Offset
    min_tolerance: float = 0.1 #Fill this in from config file
    max_tolerance: float = 6.0 #Fill this in from config file
    tolerance_indicator: str = "Outside Tolerance"
    tolerance_color: str = "red" #red/yellow/green

    #Debugging mode
    enable_test_mode: bool = False
    
    allow_print: str = "disabled" #normal/disabled
    print_text: str = "Cut To Length"
    laser_status: str = ""
    
    laser_object: serial.Serial = serial.Serial() #Gets initialized in setupLaser()
    laser_is_connected: bool = False #True if the laser is connected, false if not.
    laser_port: str = "COM3" #Fill this in from config file

    lbl_order: ttk.Label
    lbl_length: ttk.Label
    lbl_tolerance_indicator: ttk.Label
    lbl_table_length_box: ttk.Label
    lbl_off_by_box: ttk.Label
    lbl_order_length_box: ttk.Label
    lbl_error_code: ttk.Label
    
    btn_print: ttk.Button
    btn_laser_reset: ttk.Button

    def read_config_file(self):
        c = ConfigParser()

        #Debug / all by default, gets changed by config file when not in test mode.
        logLevel = 10

        # Resolve config file path relative to the exe (or script) location
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_dir, 'wespa39-128.ini')

        #No colons in the logfile name, just a yyyy-mm-dd hhmmss timestamp
        logging.basicConfig(filename=os.path.join(base_dir, time.strftime('%Y-%m-%d %H%M%S') + ' wespa39-128.log'),
            level=logLevel,
            format='%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S')
        
        try:
            # Set up logging first so config read errors are captured
            c.read(config_path)

            if c.has_option('debug', 'enableTestMode') and c.getboolean('debug', 'enableTestMode'):
                #Test mode is hardcoded False by default
                self.enable_test_mode = c.getboolean('debug', 'enableTestMode')
            elif c.has_option('debug', 'enableLogging') and c.getboolean('debug', 'enableLogging'):
                #If not in test mode, set log level per the config file.
                logLevel = c.getint('debug', 'loggingLevel')
                logging.getLogger().setLevel(logLevel)

            self.laser_port = c.get('ports', 'laserComPort')
            self.laser_offset = c.getfloat('offsets', 'laserOffset')
            self.min_tolerance = c.getfloat('offsets', 'minTolerance')
            self.max_tolerance = c.getfloat('offsets', 'maxTolerance')

        except Exception as e:
            logging.error("Error reading config file: %s", e)
            logging.error("Using default config values.")

        logging.info("Config file loaded.")

    #Takes a float (dec_inches) and returns string formatted as XXft YYin or YYin if no feet
    def get_inches_str(self, dec_inches: float):
        feet = abs(round(dec_inches/12))
        inches = abs(round(dec_inches - feet * 12, 2))
        neg_sign = "-" if dec_inches < 0 else ""

        if feet > 0:
            return "{0}{1} FT {2} IN".format(neg_sign, feet, inches)
        else:
            return "{0}{1} IN".format(neg_sign, inches)

        
    #Laser outputs in meters, convert here
    def meters_to_inches(self, meters: float):
        return meters * 39.3701
    

    #Send ZPL code to the default system printer along with the data to print.
    def send_print_label(self):       
        #https://timgolden.me.uk/python/win32_how_do_i/print.htm
        if self.allow_print == "normal":
            logging.info("Printing Label...")
            #Formatted order length
            ol_string = self.get_inches_str(self.order_length)
            
            raw_label = "^XA"
            raw_label += "^CFA,20"
            raw_label += "^FO0,90^FDWO#" + self.order_str + ":   " + ol_string + "^FS"
            raw_label += "^FO0,110^FDProduced:  " + self.get_inches_str(self.adjusted_length) + "^FS"
            raw_label += "^FO0,130^FDTolerance: " + self.get_inches_str(self.min_tolerance) + "^FS"
            raw_label += "^FO0,150^FDOff by:    " + self.get_inches_str(self.order_difference) + "^FS"
            raw_label += "^XZ"

            label_bytes=bytes(raw_label, "utf-8")
            #logging.info("Label Bytes: " + str(label_bytes))
            #The only catch here is that the label printer must be selected as system default.
            default_printer = win32print.GetDefaultPrinterW()
            my_printer = win32print.OpenPrinter(default_printer)
            logging.info("Default Printer: " + default_printer)
            try:
                #print("Starting print job...")
                #Per win32print documentation, arg 2 must be None to print to a printer.
                win32print.StartDocPrinter(my_printer, 1, ("Label" + ol_string, None, "RAW"))
                #print("Starting page...")
                win32print.StartPagePrinter(my_printer)
                #print("Writing label bytes...")
                win32print.WritePrinter(my_printer, label_bytes)
                #print("Ending page...")
                win32print.EndPagePrinter(my_printer)
                #print("Ending print job...")
                win32print.EndDocPrinter(my_printer)
            except Exception as e:
                logging.error("Error printing label: %s", e)
            finally:
                logging.info("Closing printer...")
                win32print.ClosePrinter(my_printer)


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
            logging.info("Laser connected!")
            self.laser_is_connected = True
        except serial.SerialTimeoutException as e:
            self.laser_status = "Laser connection on " + self.laser_port + " timed out."
            logging.error("Laser read timed out: %s", e)
        except serial.SerialException as e:
            self.laser_status = "Laser not found on " + self.laser_port + " - check connection and configuration."
            logging.error("Serial exception: %s", e)
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            logging.error(" Unhandled Exception: %s", e)


    def get_laser_length(self):
        if not self.laser_is_connected:
            self.laser_status = "Laser not connected."
            logging.warning("Laser not connected.")
            #Note that this will prevent the getLaserLength function from running again.
            #Use the reset button to reconnect, then manually override with g to get it going again.
            #Hopefully that won't be necessary very often...or at all.
            return
        
        logging.info("Getting laser length (DM)")
        #Consider using the more precise DS command instead of DM
        # Note that DM is not instant, but is faster than DT, which can take up to 6 seconds.
        #Whatever I choose will need to take this timing into account
        # The laser currently has an ST of 0 (no limit).
        re = ""
        try:
            self.laser_object.write(b'DM\n') #Send the command to get the length
            logging.info("Waiting for laser response...")
            time.sleep(0.25) #Wait for the laser to respond
            re = self.laser_object.readline()
            logging.info("Laser response: %s", re)
            self.laser_length = self.meters_to_inches(float(re.decode('utf-8').strip()))
            self.adjusted_length = self.laser_length + self.laser_offset
        except serial.SerialTimeoutException:
            logging.error("Laser read timed out.")
            self.laser_status = "Laser offline."
            self.laser_is_connected = False
        except ValueError:
            self.laser_status = self.parse_laser_error(str(re).strip())
            logging.error("Non-numeric value received from laser.")
            self.laser_length = 0.0
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            logging.error("Unhandled Exception: %s", e)
            self.laser_length = 0.0

        logging.info("Flushing buffer...")
        self.laser_object.flush() #Clear the input buffer to avoid reading old data
        self.update()

        #Call this function again after 500ms
        self.after(500, self.get_laser_length)


    #Send an off / on signal to the laser, or try to reconnect if it's not connected.
    def reset_laser(self):
        if not self.laser_is_connected:
            logging.warning("Laser not connected. Attempting to reconnect...")
            if self.laser_object is not None:
                self.laser_object.close() #Close the serial port if it's open
            self.setup_laser()
            return

        logging.info("Resetting Laser...")
        ##Send a LF followed by LO after a short delay
        # Try using ascii("LF\n") to send the LF and LO commands if a string literal doesn't work.
        try:
            logging.info("Writing LF (laser off)")
            self.laser_object.write(b'LF\r\n')
            logging.info("Checking laser response...")
            rl = str(self.laser_object.readline()).strip()
            self.laser_status = self.parse_laser_error(rl)
            logging.info("Laser response: %s", rl)
            time.sleep(1) #Wait for the laser to reset
            self.laser_object.flush()
            logging.info("Writing LO (laser on)")
            self.laser_object.write(b'LO\r\n')
            logging.info("Checking laser response...")
            rl = str(self.laser_object.readline()).strip()
            logging.info("Laser response: %s", rl)
            self.laser_status = self.parse_laser_error(rl)
        except serial.SerialTimeoutException as e:
            logging.error("Laser reset timed out: %s", e)
            self.laser_status = "Laser offline."
            self.laser_is_connected = False
        except Exception as e:
            self.laser_status = "Unhandled exception. Restart program."
            logging.error("Unhandled Exception: %s", e)

        logging.info("Flushing buffer...")
        self.laser_object.flush() #Clear the input buffer to avoid reading old data
        self.update()   


    def parse_laser_error(self, err: str):
        response = ""
        match err:
            case "E15": response = err + ": Sensor slow to respond"
            case "E16": response = err + ": Too much target reflectance"
            case "E17": response = err + ": Too much ambient light"
            case "E18": response = err + ": DX mode: Measured greater than specified range"
            case "E19": response = err + ": DX mode: Target speed > 10m/s"
            case "E23": response = err + ": Temp below 14F"
            case "E24": response = err + ": Temp above 140F"
            case "E31": response = err + ": Faulty memory hardware, EEPROM error"
            case "E51": response = err + ": High ambient light or hardware error"
            case "E52": response = err + ": Faulty laser diode"
            case "E53": response = err + ": EEPROM parameter not set (or divide by zero error)"
            case "E54": response = err + ": Hardware error (PLL)"
            case "E55": response = err + ": Hardware error"
            case "E61": response = err + ": Invalid serial command"
            case "E62": response = err + ": Hardware error or Parity error in serial settings"
            case "E63": response = err + ": SIO Overflow"
            case "E64": response = err + ": Framing - error SIO"
            case "LO": response = err + ": Laser is on"
            case "LF": response = err + ": Laser is off"
            case '': response = "No response from laser."

        logging.info("Laser status: %s", response)

        return response


    #Deals with keyboard input from the barcode scanner.
    #Side effect of this is that it allows for manual input of the barcode scanner, which is actually a desired feature.
    def capture_barcode(self, event):
        if event.keysym == 'Return':
            logging.info("Received input: %s", self.scanner_input)
            self.current_barcode = self.scanner_input
            self.scanner_input = ""  # Clear the input after processing
            self.update()
        elif (event.char >= '0' and event.char <= '9' or event.char == '.'):
            self.scanner_input += event.char  # Append the character to the input string


    def clear_barcode(self):
        logging.info("Clearing Barcodes...")
        self.scanner_input = ""
        self.current_barcode = ""
        self.update()


    def parse_barcode(self):
        logging.info("Updating GUI. Barcode: %s; Laser Length: %f", self.current_barcode, self.laser_length)

        if (not self.enable_test_mode):
            self.order_str = "    "
            self.order_length = 0.0
        
        #Until Line128 is used, Work Order won't be in the barcode - be sure to code for it not being there.
        #This assumes a Line128 style code. How to better detect what kind of code it is?
        #Simple, check if the first 4 chars are all digits (Line39 has 3 max). If so, it's a Line128 code.
        if len(self.current_barcode) > 4 and self.current_barcode[0:4].isdigit():
            self.order_str = self.current_barcode[0:4]
            try:
                self.order_length = round(float(self.current_barcode[4:]), 2)
            except ValueError:
                logging.error("ValueError: Could not convert %s to float.", self.current_barcode[4:])
        elif self.current_barcode is not None and len(self.current_barcode) > 0:
            #If we get here and the barcode isn't empty, it's probably a Line39 code.
            try:
                self.order_length = round(float(self.current_barcode), 2)
            except ValueError:
                logging.error(" ValueError: Could not convert %s to float.", self.current_barcode)
        else:
            #currentBarcode is empty or wrong format.
            self.order_str = "    "
            self.order_length = 0.0
            logging.error("Error: Barcode %s is empty or in the wrong format.", self.current_barcode)


    #Check tolerance values and update the GUI accordingly
    def check_tolerance(self):
        tolerance_position: str = ""
        if self.order_length < self.adjusted_length:
            tolerance_position = ": Too Long"
        elif self.order_length > self.adjusted_length:
            tolerance_position = ": Too Short"

        #Will change between green, yellow, and red based on tolerance, with text changing as well (Within/Near/Outside Tolerance)
        if abs(self.order_difference) <= self.min_tolerance and abs(self.order_difference) >= 0:
            self.tolerance_indicator = "Within Tolerance"
            self.tolerance_color = "green"
            self.allow_print = "normal"
        elif abs(self.order_difference) <= self.max_tolerance and abs(self.order_difference) > self.min_tolerance:
            self.tolerance_indicator = "Near Tolerance" + tolerance_position
            self.tolerance_color = "yellow"
            self.allow_print = "disabled"
        else:
            self.tolerance_indicator = "Outside Tolerance" + tolerance_position
            self.tolerance_color = "red"
            self.allow_print = "disabled"

    
    #Call this once a barcode has been detected or as the laser refreshes.
    #Update the GUI with the new information.
    #Also run if error codes are detected.
    def update(self):

        self.parse_barcode()

        #Round values
        self.laser_length = round(self.laser_length, 2)
        self.laser_offset = round(self.laser_offset, 2)
        self.order_length = round(self.order_length, 2)

        #Compute adjustments
        self.adjusted_length = round(self.laser_length + self.laser_offset, 2)
        self.order_difference = round(self.adjusted_length - self.order_length, 2)

        self.check_tolerance()

        #Write values to labels and update
        self.lbl_order.config(text=self.order_str)
        self.lbl_order.update()
        self.lbl_length.config(text=self.get_inches_str(self.order_length))
        self.lbl_length.update()
        self.lbl_tolerance_indicator.config(text=self.tolerance_indicator, background=self.tolerance_color)
        self.lbl_tolerance_indicator.update()
        self.btn_print.configure(state=self.allow_print)
        self.btn_print.update()
        self.lbl_table_length_box.config(text=self.get_inches_str(self.adjusted_length))
        self.lbl_table_length_box.update()
        self.lbl_off_by_box.config(text=self.get_inches_str(self.order_difference))
        self.lbl_off_by_box.update()
        self.lbl_order_length_box.config(text=self.get_inches_str(self.order_length))
        self.lbl_order_length_box.update()
        self.lbl_error_code.config(text=self.laser_status)
        self.lbl_error_code.update()

        logging.info("Order Length: %f, Order Number: %s, Raw Table Length: %f, Laser Offset: %f, Order Off By: %f",
                      self.order_length, self.order_str, self.laser_length, self.laser_offset, self.order_difference)


    #Called when the program closes.
    def on_exit(self):
        logging.warning("Closing serial port and program...")
        #This should prevent get_laser_length() from running again.
        self.laser_is_connected = False
        try:
            self.laser_object.close() #Close the serial port if it's open
        except Exception as e:
            logging.error("Error closing serial port: %s", e)

        self.destroy()


    def __init__(self, *args, **kwargs):
        ttk.Tk.__init__(self, *args, **kwargs)

        self.read_config_file()

        logging.info("Initializing GUI...")
        self.resizable(True, True)
        #Really should set this externally; just need to remember to update manually.
        self.title("WESPA 39-128 v1.4")
        
        #Number of columns and rows in the grid - all resize at the same rate
        for i in range(3):
            self.columnconfigure(i, weight=1)
        for i in range(7):
            self.rowconfigure(i, weight=1)

        # Bind keyboard shortcuts; also detect barcode input
        self.bind('<x>', lambda event: self.clear_barcode())
        self.bind('<l>', lambda event: self.reset_laser())
        self.bind('<g>', lambda event: self.get_laser_length())
        self.bind('<space>', lambda event: self.send_print_label())
        #All other keys need to be captured for the barcode scanner, which is keyboard-like input.
        self.bind('<Key>', self.capture_barcode)

        #Call on_exit() when the window is closed, for a more graceful shutdown.
        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        
        base_size = 12
        smallest_font = font.Font(size=base_size)
        small_bold_font = font.Font(size=base_size+4, weight="bold")
        medium_bold_font = font.Font(size=base_size+8, weight="bold")
        large_bold_font = font.Font(size=base_size+16, weight="bold")
        
        lbl_last_barcode = ttk.Label(self, text="Last Barcode Scanned:", justify="left", font=smallest_font)
        lbl_last_barcode.grid(column=0, row=0, padx=5, pady=5, sticky="nw")

        #WO number label
        lbl_work_order = ttk.Label(self, text="Work Order: ", justify="left", font=small_bold_font)
        lbl_work_order.grid(column=1, row=0, padx=5, pady=5, sticky="ne")

        #WO number textbox
        self.lbl_order = ttk.Label(self, text=self.order_str, justify="left", font=small_bold_font)
        self.lbl_order.grid(column=2, row=0, padx=5, pady=5, sticky="nw")
        
        #Length label (last scanned)
        lbl_length = ttk.Label(self, text="Length: ", justify="left", font=small_bold_font)
        lbl_length.grid(column=1, row=1, padx=5, pady=5, sticky="ne")

        #Length textbox (last scanned)
        self.lbl_length = ttk.Label(self, text=self.get_inches_str(self.order_length), justify="left", font=small_bold_font)
        self.lbl_length.grid(column=2, row=1, padx=5, pady=5, sticky="nw")

        #Table Length label
        lbl_table_length = ttk.Label(self, text="TABLE LENGTH:", justify="left", font=medium_bold_font)
        lbl_table_length.grid(column=0, row=2, padx=25, pady=5, sticky="nsew")

        #Table Length textbox
        self.lbl_table_length_box = ttk.Label(self, text=self.get_inches_str(self.adjusted_length),
                                               justify="center", background="white", relief="solid", font=medium_bold_font)
        self.lbl_table_length_box.grid(column=0, row=3, padx=5, pady=5, sticky="nsew")

        #OffBy Label
        lbl_off_by = ttk.Label(self, text="OFF BY:", justify="center", font=medium_bold_font)
        lbl_off_by.grid(column=1, row=2, padx=25, pady=5, sticky="nsew")

        #OffBy Textbox
        self.lbl_off_by_box = ttk.Label(self, text=self.get_inches_str(self.order_difference),
                                         background="white", relief="solid", font=medium_bold_font)
        self.lbl_off_by_box.grid(column=1, row=3, padx=5, pady=5, sticky="nsew")

        #Order Length Label
        lbl_order_length = ttk.Label(self, text="ORDER LENGTH:", justify="left", font=medium_bold_font)
        lbl_order_length.grid(column=2, row=2, padx=25, pady=5, sticky="nsew")

        #Order Length Textbox
        self.lbl_order_length_box = ttk.Label(self, text=self.get_inches_str(self.order_length),
                                               justify="right", background="white", relief="solid", font=medium_bold_font)
        self.lbl_order_length_box.grid(column=2, row=3, padx=5, pady=5, sticky="nsew")

        #Tolerance Indicator
        self.lbl_tolerance_indicator = ttk.Label(self, text=self.tolerance_indicator,
                                                  background=self.tolerance_color, font=large_bold_font)
        self.lbl_tolerance_indicator.grid(column=0, row=4, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        #Clear button
        btn_clear = ttk.Button(self, text="CLEAR\n(X)", font=medium_bold_font)
        btn_clear.grid(column=0, row=5, padx=5, pady=5)
        btn_clear.bind('<Button-1>', lambda event: self.clear_barcode())
        
        #Print button
        self.btn_print = ttk.Button(self, text="PRINT\n(space)", font=medium_bold_font)
        self.btn_print.grid(column=1, row=5, padx=5, pady=5)
        self.btn_print.bind("<Button-1>", lambda event: self.send_print_label())
        self.btn_print.configure(state="disabled" if self.allow_print == "disabled" else "normal")

        #Reset/reconnect button
        self.btn_laser_reset = ttk.Button(self, text="RESET LASER\n(L)", font=medium_bold_font)
        self.btn_laser_reset.grid(column=2, row=5, padx=5, pady=5)
        self.btn_laser_reset.bind('<Button-1>', lambda event: self.reset_laser())

        self.setup_laser()

        #Laser status
        self.lbl_error_code = ttk.Label(self, text=self.laser_status, justify="left", font=smallest_font)
        self.lbl_error_code.grid(column=0, row=6, columnspan=3, padx=5, pady=5, sticky="w")

        logging.info("GUI Initialized!")

        if(self.enable_test_mode):
            logging.info("Test mode enabled.")
            self.run_tests()
        else:
            #Repeat running the get_laser_length function - lack of parentheses indicate a function pointer.
            self.after(1000, self.get_laser_length)

    def run_tests(self):
        #Test edge cases like 10.00ft turning into 9ft 12in, problems with negatives, general gui check etc.
        #Need to run via existing code and not simply assign variables.
        self.laser_status = "Test Mode Active"

        #No offset for testing. Want to diagnose two bugs
        #1. 19Ft 12In on table length - should be 20Ft 00In
        #2. Tolerance issue for negative differences
        self.single_test("Ctrl", 120.00, 120.00)
        self.single_test("Rounding1", 120.00, 119.99)
        self.single_test("Rounding2", 120.00, 120.01)
        self.single_test("Rounding3", 119.999, 120.001)
        #That's the case for the first bug - a rounding issue 
        self.single_test("Rounding4", 120.001, 119.999)



        return

    def single_test(self, order_str, current_barcode, laser_length):
        self.order_str = order_str
        self.current_barcode = str(current_barcode)
        self.laser_length = laser_length
        self.update()
        time.sleep(3)

        
if __name__== "__main__":
    app = MainMenu()
    app.mainloop()
