import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
from pdf2docx import Converter

# ---------------------------
# Conversion Functions
# ---------------------------

def word_to_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("Word Files", "*.docx")]
    )

    if not file_path:
        return

    try:
        output_path = os.path.splitext(file_path)[0] + ".pdf"
        convert(file_path, output_path)
        messagebox.showinfo("Success", f"Converted to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def pdf_to_word():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_path:
        return

    try:
        output_path = os.path.splitext(file_path)[0] + ".docx"
        cv = Converter(file_path)
        cv.convert(output_path)
        cv.close()
        messagebox.showinfo("Success", f"Converted to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# ---------------------------
# GUI Setup
# ---------------------------

root = tk.Tk()
root.title("Word ↔ PDF Converter")
root.geometry("400x250")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="Word ↔ PDF Converter",
    font=("Arial", 16, "bold")
)
title_label.pack(pady=20)

btn_word_to_pdf = tk.Button(
    root,
    text="Convert Word to PDF",
    width=25,
    height=2,
    command=word_to_pdf
)
btn_word_to_pdf.pack(pady=10)

btn_pdf_to_word = tk.Button(
    root,
    text="Convert PDF to Word",
    width=25,
    height=2,
    command=pdf_to_word
)
btn_pdf_to_word.pack(pady=10)

exit_button = tk.Button(
    root,
    text="Exit",
    width=10,
    command=root.quit
)
exit_button.pack(pady=15)

root.mainloop()
