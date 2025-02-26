import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
import subprocess
from pdf2docx import Converter  # PDF to DOCX conversion
import pandas as pd  # Excel to CSV conversion

try:
    from docx2pdf import convert  # type: ignore
except ImportError:
    messagebox.showerror("Error", "docx2pdf module not installed. Run: pip install docx2pdf")
    exit()

supported_formats = {
    ".pdf": ".docx",
    ".docx": ".pdf",
    ".xlsx": ".csv",
    ".csv": ".xlsx",
}

def get_file_icon(extension):
    icon_map = {
        ".pdf": "icons/pdf_icon.png",
        ".docx": "icons/docx_icon.png",
        ".xlsx": "icons/xlsx_icon.png",
        ".csv": "icons/csv_icon.png"
    }
    default_icon = "icons/default_icon.png"
    return icon_map.get(extension, default_icon) if os.path.exists(icon_map.get(extension, default_icon)) else default_icon

def open_file(file_path):
    if file_path and os.path.exists(file_path):
        try:
            subprocess.run(["open" if os.name == "posix" else "start", file_path], shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {e}")

def browse_file():
    global file_path, img_tk
    file_path = filedialog.askopenfilename(filetypes=[
        ("PDF Files", "*.pdf"),
        ("Word Files", "*.docx"),
        ("Excel Files", "*.xlsx"),
        ("CSV Files", "*.csv")
    ])
    
    if file_path:
        ext = os.path.splitext(file_path)[1]
        icon_path = get_file_icon(ext)
        label_file.config(text=os.path.basename(file_path), fg="blue")
        
        try:
            img = Image.open(icon_path)
            img = img.resize((50, 50), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)
            img_label.config(image=img_tk)
            img_label.image = img_tk
            img_label.bind("<Button-1>", lambda e: open_file(file_path))
        except Exception as e:
            print(f"Error loading image: {e}")

def convert_file():
    global file_path, converted_file_path, converted_img_tk
    if not file_path:
        messagebox.showerror("Error", "No file selected")
        return

    ext = os.path.splitext(file_path)[1]
    
    if ext not in supported_formats:
        messagebox.showerror("Error", "Unsupported file format")
        return

    converted_file_path = os.path.splitext(file_path)[0] + supported_formats[ext]

    try:
        if ext == ".pdf":
            convert_pdf_to_docx(file_path, converted_file_path)
        elif ext == ".docx":
            convert_docx_to_pdf(file_path)
        elif ext == ".xlsx":
            convert_excel_to_csv(file_path, converted_file_path)
        elif ext == ".csv":
            convert_csv_to_excel(file_path, converted_file_path)

        label_output.config(text=f"Converted: {os.path.basename(converted_file_path)}", fg="green")
        
        converted_icon_path = get_file_icon(supported_formats[ext])
        img = Image.open(converted_icon_path)
        img = img.resize((50, 50), Image.LANCZOS)
        converted_img_tk = ImageTk.PhotoImage(img)
        converted_img_label.config(image=converted_img_tk)
        converted_img_label.image = converted_img_tk
        converted_img_label.bind("<Button-1>", lambda e: open_file(converted_file_path))

        btn_convert.config(text="Converted ✅", bg="green", fg="white")
        root.update()
        root.after(2000, lambda: btn_convert.config(text="Convert", bg="green", fg="white"))  
    
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed: {e}")

def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

def convert_docx_to_pdf(docx_path):
    try:
        convert(docx_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert DOCX to PDF: {e}")

def convert_excel_to_csv(excel_path, csv_path):
    df = pd.read_excel(excel_path)
    df.to_csv(csv_path, index=False)

def convert_csv_to_excel(csv_path, excel_path):
    df = pd.read_csv(csv_path)
    df.to_excel(excel_path, index=False)

def main():
    global root, img_label, converted_img_label, label_file, label_output, btn_convert
    root = tk.Tk()
    root.title("Growth Mindset Challenge: Web App with Streamlit")
    root.geometry("500x500")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="File Converter Web App", font=("Arial", 16, "bold"), bg="#f0f0f0", fg="black")
    header_label.pack(pady=10)

    img_label = tk.Label(root, bg="#f0f0f0")
    img_label.pack(pady=5)

    label_file = tk.Label(root, text="No file selected", fg="blue", bg="#f0f0f0", font=("Arial", 12))
    label_file.pack(pady=10)
    
    btn_browse = tk.Button(root, text="Browse", command=browse_file, bg="green", fg="white", font=("Arial", 12))
    btn_browse.pack(pady=5)
    
    btn_convert = tk.Button(root, text="Convert", command=convert_file, bg="green", fg="white", font=("Arial", 12))
    btn_convert.pack(pady=5)

    label_output = tk.Label(root, text="", fg="blue", bg="#f0f0f0", font=("Arial", 12))
    label_output.pack(pady=10)
    
    converted_img_label = tk.Label(root, bg="#f0f0f0")
    converted_img_label.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
