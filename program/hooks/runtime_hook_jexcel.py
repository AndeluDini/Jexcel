import sys
import os

# In onefile mode, PyInstaller extracts files to sys._MEIPASS.
if getattr(sys, 'frozen', False):
    # Ensure that sys._MEIPASS is in sys.path
    sys.path.insert(0, sys._MEIPASS)