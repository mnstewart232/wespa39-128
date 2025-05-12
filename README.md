Requires pyinstaller, pyserial, and win32print libraries to build.

Build with

> pyinstaller --onefile --noconsole --icon "logo_sq_notext_ctr_256.ico" "wespa39-128.py"

Make sure that the label printer is set as the system default in order for the program to work.

Make sure the .ini and .exe files are in the same directory. Log files (if enabled) will save to this directory as well.
