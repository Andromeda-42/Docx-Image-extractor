"""Test UI program"""

import os
import shutil
import zipfile
import sys
import tkinter as tk
from pathlib import Path
import zipfile
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

    # Ensure the output folder exists
    os.makedirs("./pics", exist_ok=True)

    # Open the .docx file as a zip archive
    with zipfile.ZipFile(original_file_source, "r") as docx_zip:
        # List all files in the archive
        all_files = docx_zip.namelist()

        # Filter for image files (usually in 'word/media/')
        image_files = [f for f in all_files if f.startswith("word/media/")]

        # Extract and save each image
        for image_file in image_files:
            image_data = docx_zip.read(image_file)
            image_name = os.path.basename(image_file)
            output_path = os.path.join("./pics", image_name)

            with open(output_path, "wb") as f:
                f.write(image_data)

        print(f"Extracted {len(image_files)} images to pics.")


# Example usage
# extract_images_from_docx("example.docx", "output_images")


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

button1 = ttk.Button(text="Open Word Document or PDF", command=open_file_dialog)
button1.config(width=30)
button1.pack(side="top", pady=20, ipady=5)

root.mainloop()
