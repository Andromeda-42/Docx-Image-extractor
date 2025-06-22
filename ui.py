"""Test UI program"""

import os
import shutil
import zipfile
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


def open_file_dialog():
    """Opens a file dialog to select a file"""
    file_path = filedialog.askopenfile(title="Select a File", initialdir="./")
    if file_path:
        image_extraction(file_path)
        file_path.close()
    else:
        print("No file selected")


def image_extraction(original_file_source):
    try:
        source_no_extension, old_extension = os.path.splitext(original_file_source.name)

        new_extension = ".zip"

        dst_source = source_no_extension + new_extension

        shutil.copy(original_file_source.name, dst_source)
    except FileNotFoundError:
        print("You Goofed")
        sys.exit(1)

    with zipfile.ZipFile(dst_source, "r") as zip_ref:
        zip_ref.extractall("./Image Dump")
    os.remove(dst_source)
    shutil.move("./Image Dump/word/media", "./")
    os.remove("./Image Dump")


root = tk.Tk()
root.title("Image Extractor")
root.geometry("500x300")

header1 = ttk.Label(text="Dan's Image Extractor", font=("Arial", 20, "bold"))
header1.pack(side="top", anchor="nw", padx=20, pady=10)

label1 = ttk.Label(
    text='Please only open word documents, specifically ".DOCX"', font=("Arial", 10)
)
label1.pack(side="top", anchor="w", padx=20)

button1 = ttk.Button(text="Open File", command=open_file_dialog)
button1.config(width=30)
button1.pack(side="top", pady=20, ipady=5)

root.mainloop()
