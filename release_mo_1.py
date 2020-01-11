import json
import openpyxl
import datetime

JOB = "0519-LR6"
#JOB = input("What job would you like to input?: ")

wb = openpyxl.load_workbook('QS-100 MASTER SCHEDULE.xlsx', data_only = True)
#sheet = wb.sheetnames[0]
sheet = wb[JOB]
MAX = 35
#MAX = sheet.max_row

def format_date(d):
    year = str(d.year)
    month = str(d.month).zfill(2)
    day = str(d.day).zfill(2)
    final_date = "{}{}{}".format(month,day,year)
    return(final_date)

#* MO number
#* Drawing number & Description
# mo_list = []
# drawing_list = []
# description_list = []
# start_date_list = []
# finish_date_list = []
# qty_list = []

job_list = []
builder = {}

for row in range(4,MAX+1):
    for i in range(1,5):
        mo = str(sheet.cell(row=row, column=i).value)
        description_drawing = sheet.cell(row=row, column=i+1).value
        if "None" not in mo:
            break
    # mo_list.append(mo)
    builder['mo'] = mo
    description_drawing_split = description_drawing.split(" - ")
    # drawing_list.append(description_drawing_split[0])
    builder['drawing_list'] = description_drawing_split[0]

    try:
        # description_list.append(description_drawing_split[1])
        builder['description_list'] = description_drawing_split[1]
    except:
        # description_list.append("")
        continue

    #print(mo_list)
    #print(drawing_list)
    #print(description_list)

    #* Start date
    start_date = sheet.cell(row=row, column=9).value
    builder['start_date'] = format_date(start_date)
    #print(start_date_list)

    #* Finish date
    finish_date = sheet.cell(row=row, column=10).value
    builder['finish_date'] = format_date(finish_date)
    # finish_date_list.append(new_finish_date)
    #print(finish_date_list)

    #* QTY
    qty = str(sheet.cell(row=row, column=11).value)
    # qty_list.append(qty)
    builder['qty'] = qty
    #print(qty_list)
    job_list.append(builder)

complete_job = {JOB:job_list}
# print(complete_job)
a = json.dumps(complete_job,indent=2)
print(a)
#TODO - Save all info into a json file

# for i in range(0,len(mo_list)):

#     my_dict = { "0519-LR6": {
#             mo_list[i]: {
#                 "Drawing Number": drawing_list[i],
#                 "Description": description_list[i],
#                 "QTY": qty_list[i],
#                 "Start Date": start_date_list[i],
#                 "Finish Date": finish_date_list[i],
#             }
#         }
#     }
#     data = json.dumps(my_dict, indent=2, sort_keys=True) # This would be the JSON file

#     print(data)

#TODO - Sort json file in alphabetical order. Or use a range to go in numerical order.
#TODO - Use that info to release MO's