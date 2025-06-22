"""
Test program for extraction of Word doc via source code
"""

import os
import shutil
import zipfile
import sys

# get the file

original_filename = input()
try:
    base_name, old_extension = os.path.splitext(original_filename)

    NEW_EXTENSION = ".zip"

    file_Source = "./" + original_filename
    dst_source = "./" + base_name + NEW_EXTENSION

    shutil.copy(file_Source, dst_source)
except FileNotFoundError:
    print("You Goofed")
    sys.exit(1)


with zipfile.ZipFile(dst_source, "r") as zip_ref:
    zip_ref.extractall("./Image Dump")
