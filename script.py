import tkinter as tk
from docx import Document
from docx2pdf import convert
import os


placeholders = {
        "RecipientName": "",
        "RecipientTitle": "",
        "Position": "",
        "CompanyName": "",
        "CompanyAddress": "",
        "CityProvincePostal": ""
}

def replace_placeholder(doc, placeholder, replacement):
    for p in doc.paragraphs:
        if placeholder in p.text:
            inline = p.runs
            for item in range(len(inline)):
                if placeholder in inline[item].text:
                    inline[item].text = inline[item].text.replace(placeholder, replacement)


def generate_cover_letter():
    filename = "coverletter.docx"
    doc = Document(filename)
    global placeholders
    for placeholder in placeholders:
        replace_placeholder(doc, placeholder, str(placeholders[placeholder].get()))
    output_filename = f"{str(placeholders['CompanyName'].get())}_{str(placeholders['Position'].get())}.docx"
    doc.save(output_filename)
    convert(output_filename)
    os.remove(output_filename)


def quit_application():
    root.destroy()


# Create the GUI
root = tk.Tk()
root.title("Cover Letter Generator")
root.geometry("400x400")

# Add labels and entry fields for each placeholder
for placeholder in placeholders:
    label = tk.Label(root, text=placeholder)
    label.pack(pady=3)
    entry = tk.Entry(root)
    entry.pack()
    placeholders[placeholder] = entry

# Add generate and quit buttons
generate_button = tk.Button(root, text="Generate Cover Letter", command=generate_cover_letter)
generate_button.pack(pady=5)

quit_button = tk.Button(root, text="Quit", command=quit_application)
quit_button.pack(pady=5)

# Start the main loop
root.mainloop()
