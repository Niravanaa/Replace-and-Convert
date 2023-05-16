# Cover Letter Generator
This Python script generates personalized cover letters based on a template document. It utilizes the docx library to work with Microsoft Word documents and the docx2pdf library to convert the generated cover letters to PDF format.

# Prerequisites
* Python 3.x
* `docx` library (python-docx package)
* `docx2pdf` library

# Installation
1. Clone the repository:
```shell
$ git clone https://github.com/Niravanaa/cover-letter-generator.git
# cd cover-letter-generator
```
2. Install the required dependencies
```shell
$ pip install python-docx docx2pdf
```

# Usage
3. Place your cover letter document (named `coverletter.docx`) in the same directory as the script.
4. Run the script:
```shell
python script.py
```
5. Follow the prompts to enter the requested information for each placeholder in the cover letter template.
6. The generated cover letter will be saved as a Word document (`[CompanyName]_[Position].docx`) and automatically converted into PDF format.
7. The original Word document will be deleted.
8. To quit the script, press `q` when prompted to continue generating cover letters.

# Acknowledgements
* Idea inspired by [Keshan Kathiripilay](https://www.linkedin.com/in/keshankathiripilay/).
