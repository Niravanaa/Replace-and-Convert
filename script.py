import docx
from docx2pdf import convert
import os


def replace_placeholder(doc, placeholder, replacement):
    for p in doc.paragraphs:
        if placeholder in p.text:
            inline = p.runs
            for item in range(len(inline)):
                print(placeholder, inline[item].text)
                if placeholder in inline[item].text:
                    print("found!")
                    inline[item].text = inline[item].text.replace(placeholder,
                                                                  replacement)


def main():
    while True:
        filename = "coverletter.docx"
        doc = docx.Document(filename)
        placeholders = {
            "RecipientName": "",
            "RecipientTitle": "",
            "Position": "",
            "CompanyName": "",
            "CompanyAddress": "",
            "CityProvincePostal": ""
        }
        for placeholder in placeholders:
            replacement = input(f"Enter {placeholder}: ")
            placeholders[placeholder] = replacement
            replace_placeholder(doc, placeholder, replacement)
        doc.save(
            placeholders["CompanyName"] + "_" +
            placeholders["Position"] + ".docx"
        )
        convert(placeholders["CompanyName"] + "_" +
                placeholders["Position"] + ".docx")
        os.remove(placeholders["CompanyName"] + "_" +
                  placeholders["Position"] + ".docx")
        choice = input("Press q to quit or any other key to continue: ")
        if choice.lower() == "q":
            break


if __name__ == "__main__":
    main()
