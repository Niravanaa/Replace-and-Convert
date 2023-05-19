# Replace-and-Convert
This Python script generates PDF files based on a template .docx document. It utilizes the docx library to work with Microsoft Word documents and the docx2pdf library to convert the generated cover letters to PDF format.

**If you wish not to install the Python packages, an executable (.exe) form of the script can be downloaded <a target="_blank" href="https://drive.google.com/uc?export=download&id=1Y2sj97DiBZbzH5Gjs8M-lvH-QDu0UDuO">using this link</a>.**

**Support for MacOS will be provided shortly.**

## Prerequisites
* Python 3.x
* `docx` library (python-docx package)
* `docx2pdf` library

## Installation
1. Clone the repository:
```shell
$ git clone https://github.com/Niravanaa/replace-and-convert.git
# cd replace-and-convert
```
2. Install the required dependencies
```shell
$ pip install python-docx docx2pdf
```

## Usage
3. Place your .docx document (named `coverletter.docx`) in the same directory as the script.
4. Run the script, either from an IDE or using the following terminal command (at the directory containing the script):
```shell
python script.py
```
5. Follow the prompts to enter the requested information for each placeholder in the cover letter template.
NOTE: This script contains the following pre-coded placeholders:
```text
RecipientName: The name of the recruiter
RecipientTitle: The position of the recruiter
Position: The position you are applying for
CompanyName: The name of the company 
CompanyAddress: The address of the company (Number, Street Name)
CityProvincePostal: The city, province, and postal code of the company (e.g. Montreal, QC A0A 0A0)
```
7. The generated file will be saved as a Word document (`[CompanyName]_[Position].docx`) and automatically converted into PDF format.
8. The original Word document will be deleted.
9. To quit the script, press `q` when prompted to repeat the process.

## Acknowledgements
* Idea inspired by [Keshan Kathiripilay](https://www.linkedin.com/in/keshankathiripilay/).
