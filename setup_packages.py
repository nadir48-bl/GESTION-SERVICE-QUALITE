import subprocess
import os
from setuptools import setup, find_packages
"""dependencies = [
    'PyQt6',
    'PyQt6-Qt6',
    'docx',
    'docxtpl',
    'docx2pdf',
    'colorama',
    'comtypes',
    'python-docx',
    'python-bidi',
    'arabic_reshaper',
    'mysql-connector-python',
    'PyQt6',
    'colorama',
    'comtypes',
]
for dependency in dependencies:
    subprocess.check_call(['pip', 'install', dependency])
from setuptools import setup, find_packages
"""
setup(
    name='GSQR',
    version='1.0.3',
    scripts=['login.pyw'],  # List your main Python script here
    py_modules=['agréage.py', 'confirmité.py','gsqr.py','lgmsec.py','moulin.py','phytosanitaire.py','refus.py'],
    author='Belbacha Mohamed Nadir',
    author_email='naadirlazi48@gmail.com',
    description='An example Python package',
    packages=find_packages(),
    install_requires=[
        'PyQt6',
        'docx',
        'docxtpl',
        'docx2pdf',
        'colorama',
        'comtypes',
        'python-docx',
        'python-bidi',
        'arabic_reshaper',
        'mysql-connector-python',
        'xhtml2pdf',  # Add this line if needed
    ],
)

file=subprocess.run(["filles_exe/okular-master-1685-windows-cl-msvc2019-x86_64.exe"])
file1=subprocess.run(["filles_exe/DesktopEditors_x64.exe"])
file3 = os.startfile("reademe.txt","open")




"""def convert_to_pdf(docx_path, pdf_path):
    # Create a Microsoft Word application instance
    word = CreateObject("Word.Application")
    # Open the input DOCX file
    doc = word.Documents.Open(os.path.abspath("Docxfiles/_bulletin entré/template_BULLETIN_ENTRE.docx"))
    # Save the document as PDF
    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    # Close the document and Word application
    doc.Close()
    word.Quit()
# Example usage:
convert_to_pdf("input.docx", "output123.pdf")
"""

