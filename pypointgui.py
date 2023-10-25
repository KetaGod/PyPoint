import tkinter as tk
from tkinter import filedialog
from pptx import Presentation
from docx import Document
import subprocess
import sys
import random

def extract_text_from_pptx(pptx_file, word_output_file, summary_output_file):
    presentation = Presentation(pptx_file)
    
    all_text = ""
    for slide in presentation.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                all_text += shape.text + "\n"
    
    # Save to a Word document
    doc = Document()
    doc.add_paragraph(all_text)
    doc.save(word_output_file)


def browse_pptx():
    pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    pptx_entry.delete(0, "end")
    pptx_entry.insert(0, pptx_file)

def browse_word_output():
    word_output_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    word_output_entry.delete(0, "end")
    word_output_entry.insert(0, word_output_file)

def extract_text():
    pptx_file = pptx_entry.get()
    word_output_file = word_output_entry.get()
    extract_text_from_pptx(pptx_file, word_output_file)
    status_label.config(text=f"Text extracted and saved to '{word_output_file}'")

def run_script_silently():
    script = "pypointgui.py"
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    subprocess.Popen([sys.executable, script], startupinfo=startupinfo)

# Create a Tkinter window
window = tk.Tk()
window.title("PowerPoint Notes ExtractorV0.0.3")

# Allow resizing of the GUI
window.geometry("450x250") # Initial Size

# Set a random background color
random_color = "#{:02x}{:02x}{:02x}".format(random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))
window.configure(bg=random_color)

# Apply high contrast theme for better clarity
window.tk_setPalette(background=random_color, foreground='white')

# Create and configure GUI elements
author_label = tk.Label(window, text="Author: KetaGod | keta666")
author_label.grid(row=0, column=0, columnspan=3)

pptx_label = tk.Label(window, text="Select PowerPoint file:")
pptx_entry = tk.Entry(window, width=40)
browse_pptx_button = tk.Button(window, text="Browse", command=browse_pptx)

word_output_label = tk.Label(window, text="Select output Word file:")
word_output_entry = tk.Entry(window, width=40)
browse_word_output_button = tk.Button(window, text="Browse", command=browse_word_output)

extract_button = tk.Button(window, text="Extract Text", command=extract_text)
status_label = tk.Label(window, text="")

word_output_label.grid(row=2, column=0)
word_output_entry.grid(row=2, column=1)
browse_word_output_button.grid(row=2, column=2)

extract_button.grid(row=3, column=0, columnspan=3)
status_label.grid(row=4, column=0, columnspan=3)

# Add a multi-line information label
info_text = (
    "Thank you for using PyPoint.\n"
    "If you haven't already, please read the README.txt file.\n"
    "This program was created for those who do not want to take their own notes on PowerPoints.\n"
    "This program is able to write everything down within a second or less.\n"
    "Enjoy. Created by KetaGod"
)
info_label = tk.Label(window, text=info_text, bg=random_color, wraplength=380, justify="left")
info_label.grid(row=6, column=0, columnspan=3)

# Arrange GUI elements
pptx_label.grid(row=1, column=0)
pptx_entry.grid(row=1, column=1)
browse_pptx_button.grid(row=1, column=2)

word_output_label.grid(row=2, column=0)
word_output_entry.grid(row=2, column=1)
browse_word_output_button.grid(row=2, column=2)

extract_button.grid(row=3, column=0, columnspan=3)
status_label.grid(row=4, column=0, columnspan=3)

# Start the GUI main loop
window.mainloop()
