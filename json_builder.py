import json
import openpyxl
import datetime

def format_date(d):
    year = str(d.year)
    month = str(d.month).zfill(2)
    day = str(d.day).zfill(2)
    final_date = "{}{}{}".format(month,day,year)
    return(final_date)

wb = openpyxl.load_workbook('QS-100 MASTER SCHEDULE.xlsx', data_only = True)

jobs = []
job_list = []

for sheet in wb:

    #MAX = 35
    JOB = str(sheet.cell(row=1, column=1).value)
    CUSTOMER = str(sheet.cell(row=3, column=1).value)
    LOCATION = str(sheet.cell(row=3, column=2).value)
    MAX = sheet.max_row

    #* MO number
    #* Drawing number & Description

    
    job_builder = {}

    for row in range(5,MAX+1):
        builder = {}
        builder['Job'] = JOB
        builder['Customer'] = CUSTOMER
        builder['Location'] = LOCATION
        for i in range(1,5):
            mo = str(sheet.cell(row=row, column=i).value)
            description_drawing = sheet.cell(row=row, column=i+1).value
            if "None" not in mo:
                break
        
        builder['Manufacturing Order'] = mo
        #if "PKG" in description_drawing:
            #description_drawing_split = description_drawing.split(" ")
            #builder['Drawing No.'] = description_drawing_split[0]
            #builder['Description'] = description_drawing_split[0]
        #else:
        description_drawing_split = description_drawing.split(" - ")
        builder['Drawing No.'] = description_drawing_split[0]

        try:
            builder['Description'] = description_drawing_split[1]
        except:
            continue

        #* Start date
        start_date = sheet.cell(row=row, column=9).value
        builder['Start Date'] = format_date(start_date)

        #* Finish date
        finish_date = sheet.cell(row=row, column=10).value
        builder['Finish Date'] = format_date(finish_date)

        #* QTY
        qty = sheet.cell(row=row, column=11).value
        builder['Quantity'] = qty
        job_list.append(builder)

    #* Save all info into a json file
    job = {JOB:job_list}
    jobs.append(job)

complete_job = {"MOs": job_list}
    
with open("QS-100 MASTER SCHEDULE.json", 'w', encoding='utf-8') as f:
        json.dump(complete_job, f, ensure_ascii=False, indent=2)

#*Sort json file in alphabetical order. Or use a range to go in numerical order.

"""
job_list = []
skip_me = ['02','03']
with open("QS-100 MASTER SCHEDULE.json") as f:
    data = json.load(f)
    for j in data["MOs"]:
        if j["Job"] not in job_list:
            job_list.append(j["Job"])

    dash_count = int(len(data["MOs"]) / len(job_list))

    for i in range(0, len(job_list)):
        for k in range(1,dash_count+1):
            skip = str(k).zfill(2)
            if skip in skip_me:
                continue
            for j in data["MOs"]:
                job = j["Job"]
                customer = j["Customer"]
                mo = j["Manufacturing Order"]
                drawing_no = j["Drawing No."]
                description = j["Description"]
                start = j["Start Date"]
                finish = j["Finish Date"]
                quantity = j["Quantity"]
                mo_split = mo.split("-")
                dash = int(mo_split[2])
                if job in job_list[i]:
                    if k == dash:
                        break
            print(mo)
            print(customer)
            print(drawing_no)
            print(description)
            print(start)
            print(finish)
            print(quantity)
"""

#TODO - Use that info to release MO's
#* Should be able to just copy the pyautogui and pygewindow portion of the original script to input the given data
#* talk to howard or phillip about the customer box and ask about the vessels that are already released. maybe put a skip number option