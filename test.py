import os
import sys
import tkinter

tcl_path = os.path.dirname(os.path.abspath(tkinter.Tcl().eval('info library')))
tk_path = os.path.dirname(os.path.abspath(tkinter.Tk().eval('info library')))

print("Tcl library path:", tcl_path)
print("Tk library path:", tk_path)
