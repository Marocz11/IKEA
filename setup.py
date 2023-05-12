from cx_Freeze import setup, Executable
import os
import sys
sys.setrecursionlimit(5000)

TCL_LIBRARY = "/Users/MarekHalska/opt/anaconda3/envs/myenv/lib"
TK_LIBRARY = "/Users/MarekHalska/opt/anaconda3/envs/myenv/lib"

setup(
    name="YourAppName",
    version="0.1",
    description="Your app description",
    options={
        "build_exe": {
            "include_files": [
                (os.path.join(TCL_LIBRARY, "tcl8.6"), "tcl"),
                (os.path.join(TK_LIBRARY, "tk8.6"), "tk"),
                # Include other files if needed
            ],
            # Add any other required options
        }
    },
    executables=[Executable("main.py", base="Console")],
)
