"""Test UI program"""

import os
import shutil
import zipfile
import sys
import tkinter as tk
from pathlib import Path
from pypdf import PdfReader
from tkinter import ttk
from tkinter import filedialog


def open_file_dialog():
    """Opens a file dialog to select a file"""
    file_path = filedialog.askopenfilename(title="Select a File", initialdir="./")
    if file_path.endswith(".docx"):
        docx_image_extraction(file_path)
    elif file_path.endswith(".pdf"):
        pdf_image_extraction(file_path)
    else:
        print("No file selected")


def docx_image_extraction(original_file_source):
    """Strips the images out of a docx file by converting to .zip extracting and moving the image source folder to the forefront"""
    try:
        source_no_extension, old_extension = original_file_source.split(".")

        new_extension = ".zip"

        dst_source = source_no_extension + new_extension

        shutil.copy(original_file_source, dst_source)
    except FileNotFoundError:
        print("You Goofed")
        sys.exit(1)

    with zipfile.ZipFile(dst_source, "r") as zip_ref:
        zip_ref.extractall("./Program Data")

    # File clean up at the end
    os.remove(dst_source)
    try:
        shutil.copytree("./Program Data/word/media", "./Images")
    except FileExistsError:
        shutil.rmtree("./Images")
        shutil.copytree("./Program Data/word/media", "./Images")


def pdf_image_extraction(pdf_path):
    """
    Extract all images from *pdf_path* and save them in ./images/.
    If the images folder already exists, it is deleted and recreated.
    """
    pdf_path = Path(pdf_path).expanduser().resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"No such file: {pdf_path}")

    # --- (Re)create ./images ---
    images_dir = Path("Images")
    if images_dir.exists():
        shutil.rmtree(images_dir)  # blow away old contents
    images_dir.mkdir()

    saved_paths = []
    with pdf_path.open("rb") as fh:
        reader = PdfReader(fh)

        for page_num, page in enumerate(reader.pages, start=1):
            xobjects = (page.get("/Resources") or {}).get("/XObject") or {}

            for name, obj_ref in xobjects.items():
                xobj = obj_ref.get_object()
                if xobj.get("/Subtype") != "/Image":
                    continue

                flt = xobj.get("/Filter")
                if isinstance(flt, list):
                    flt = flt[0]
                ext = {
                    "/DCTDecode": ".jpg",
                    "/JPXDecode": ".jp2",
                    "/FlateDecode": ".png",
                    "/LZWDecode": ".tiff",
                    "/CCITTFaxDecode": ".tiff",
                }.get(flt, ".bin")

                img_data = xobj.get_data()
                out_name = f"page{page_num:03d}_{name[1:]}{ext}"
                out_file = images_dir / out_name
                out_file.write_bytes(img_data)
                saved_paths.append(out_file)

    print(f"Extracted {len(saved_paths)} image(s) to {images_dir.resolve()}")
    return saved_paths


root = tk.Tk()
root.title("Image Extractor")
root.geometry("500x300")

header1 = ttk.Label(text="Dan's Image Extractor", font=("Arial", 20, "bold"))
header1.pack(side="top", anchor="nw", padx=20, pady=10)

label1 = ttk.Label(
    text='Please only open word documents, specifically ".DOCX"', font=("Arial", 10)
)
label1.pack(side="top", anchor="w", padx=20)

button1 = ttk.Button(text="Open Word Document File", command=open_file_dialog)
button1.config(width=30)
button1.pack(side="top", pady=20, ipady=5)

button2 = ttk.Button(text="Open PDF ", command=open_file_dialog)
button2.config(width=30)
button2.pack(side="top", pady=20, ipady=5)

root.mainloop()
