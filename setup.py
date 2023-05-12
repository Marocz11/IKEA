import sys
from cx_Freeze import setup, Executable
import ssl

build_exe_options = {
    "packages": ["os", "selenium", "openpyxl", "tkinter", "bs4", "yahoofinancials", "forex_python", "re", "json", "requests", "datetime"],
    "include_files": [(ssl.get_default_verify_paths().openssl_cafile, "cacert.pem")]
}

setup(
    name="IKEA Product Scraper",
    version="0.1",
    description="A script to scrape IKEA product information",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=None)],
)
