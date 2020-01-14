import os
import PyPDF2
import json

job_list = []
with open("QS-100 MASTER SCHEDULE.json") as f:
    data = json.load(f)
    for k in data["MOs"]:
        job = k["Job"]
        if job not in job_list:
            job_list.append(job)

pdf_merge = []
for p in range(0,len(job_list)):
    for i in data["MOs"]:
        job = i["Job"]
        mo = i["Manufacturing Order"]
        if job == job_list[p]:
            job_split = job.split("-")
            job = job_split[1]
            os.chdir(r"C:\Users\jlee.NTPV\Documents\MO Reports\{}".format(job))
            for pdf in os.listdir('.'):
                if "DD" in str(pdf):
                    if mo in pdf:
                        pdf_merge.append(pdf)
                        break
            
            for pdf in os.listdir('.'):
                if "MERGE" in str(pdf):
                    continue
                if "PL" in str(pdf):
                    if mo in pdf:
                        pdf_merge.append(pdf)
                        break
            

    pdf_writer = PyPDF2.PdfFileWriter()
    for file in pdf_merge:
        pdfFileObj = open(file,'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdfFileObj)

        for pageNum in range(pdf_reader.numPages):
            pageObj = pdf_reader.getPage(pageNum)
            pdf_writer.addPage(pageObj)

            pdf_output = open('{} MERGE.pdf'.format(job_list[p]), 'wb')

            pdf_writer.write(pdf_output)

            pdf_output.close()

        pdfFileObj.close()
    pdf_merge = []

    os.chdir(r"C:\Users\jlee.NTPV\Documents\MO Reports")
