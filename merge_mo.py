import os
import shutil
import time
import pyautogui
import getpass
import PyPDF2

#TODO - merge pdfs after pdf_mo has made the pdfs. use the .json file to know the order to merge them.
#TODO  DD goes before any dash number. 

job = input("What job would you like to merge?: ")

print("\nMerging...\n")

os.chdir(r"C:\Users\jlee.NTPV\Documents\MO Reports\{}".format(job))

pdf_merge = []

for file in os.listdir('.'):
    if "MERGE" in str(file):
        continue
    if file.endswith('.pdf'):
        pdf_merge.append(file)

pdf_writer = PyPDF2.PdfFileWriter()

for file in pdf_merge:
    pdfFileObj = open(file,'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdfFileObj)

    for pageNum in range(pdf_reader.numPages):
        pageObj = pdf_reader.getPage(pageNum)
        pdf_writer.addPage(pageObj)

        pdf_output = open(job+' MERGE.pdf', 'wb')

        pdf_writer.write(pdf_output)

        pdf_output.close()

    pdfFileObj.close()

os.chdir(r"\\NTPV-SERVER2008\ntpv data\Justyn's MISys\Pick List PDFs\Jobs")
