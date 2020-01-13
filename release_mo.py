# MO release using Json

import pyautogui
import pygetwindow as gw
import time
import openpyxl
import json

pyautogui.PAUSE = 0.5

skip_me = ['01','02','03','04','05','06','07']

print('Press Ctrl-C to quit.')

try:
    start_total_time = time.time()

    # VARIABLES
    #----------------
    loop_counter = 0
    total_loop_time = 0

    # LOOP
    # ---------------
    job_list = []
    with open("QS-100 MASTER SCHEDULE.json") as f:
        data = json.load(f)
        for j in data["MOs"]:
            if j["Job"] not in job_list:
                job_list.append(j["Job"])

        dash_count = int(len(data["MOs"]) / len(job_list))

        for i in range(0, len(job_list)):
            for k in range(9,dash_count+1):
                skip = str(k).zfill(2)
                if skip in skip_me:
                    continue
                for j in data["MOs"]:
                    job = j["Job"]
                    customer = j["Customer"]
                    location = j["Location"]
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
                print(location)
                print(drawing_no)
                print(description)
                print(start)
                print(finish)
                print(quantity)

                # NEW
                #----------------
                print('-------------------------')
                start_loop_time = time.time()
                #input('Press ENTER to continue.')
                #print('-------------------------')
                #pyautogui.click(380, 80)
                #time.sleep(1)

                # WAIT FOR MO WINDOW
                # ---------------
                
                #while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 0:
                    #time.sleep(2)
                #time.sleep(8)

                # MO NO.
                # ---------------
                pyautogui.doubleClick(210, 93)
                pyautogui.typewrite(str(mo))
                pyautogui.typewrite(['tab'])
                while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) < 2:
                    time.sleep(2)
                time.sleep(1)
                pyautogui.doubleClick(1032, 595)
                time.sleep(1)
                
                # DESCRIPTION
                # ---------------

                pyautogui.doubleClick(345, 118)
                pyautogui.typewrite(str(description))
                time.sleep(1)

                # BUILD ITEM
                # ---------------

                pyautogui.doubleClick(246, 174)
                pyautogui.typewrite(str(drawing_no))
                time.sleep(1)

                # JOB NUMBER
                # ---------------

                pyautogui.doubleClick(195, 306)
                pyautogui.typewrite(str(mo))
                # tab
                # check pop up window
                pyautogui.click(200, 331)
                time.sleep(1)
                pyautogui.click(928, 537)
                pyautogui.typewrite(str(description))
                time.sleep(1)
                pyautogui.click(888, 588)
                time.sleep(1)

                # LOCATION
                # ---------------

                pyautogui.doubleClick(200, 331)
                pyautogui.typewrite(location)
                time.sleep(1)

                # CUSTOMER (LIFTROCK)
                # ---------------

                pyautogui.doubleClick(245, 354)
                pyautogui.typewrite(customer)
                time.sleep(1)

                # ORDERED
                # ---------------

                pyautogui.doubleClick(214, 417)
                pyautogui.typewrite(str(quantity))
                time.sleep(1)
                # tab
                pyautogui.click(237, 338)
                time.sleep(2)
                # CLICK YES
                pyautogui.click(1024, 595)
                time.sleep(1)
                # wait for selcet screen
                while len(gw.getWindowsWithTitle('Select Revision to Use - Bills of Material')) < 1:
                    time.sleep(2)
                time.sleep(1)
                pyautogui.click(1166, 762)
                # wait for select screen to exit
                while len(gw.getWindowsWithTitle('Select Revision to Use - Bills of Material')) == 1:
                    time.sleep(2)
                time.sleep(2)

                # COMPLETION DATE
                # ---------------

                pyautogui.doubleClick(571, 488)
                pyautogui.typewrite(str(finish))
                time.sleep(1)

                # START DATE
                # ---------------

                pyautogui.doubleClick(571, 465)
                pyautogui.typewrite(str(start))
                time.sleep(1)

            
                # CLICK DESCRIPTION BOX
                # ---------------

                pyautogui.click(345, 118)
                time.sleep(1)

                # SAVE
                #----------------
                #input()
                pyautogui.click(77, 58)
                time.sleep(1)

                # RELEASE
                # ---------------

                pyautogui.click(427, 58)
                time.sleep(1)

                # CLICK OK
                # ---------------

                while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 1:
                    time.sleep(1)
                pyautogui.click(1082, 595)
                time.sleep(5)

                loop_time = time.time() - start_loop_time
                loop_counter = loop_counter + 1
                total_loop_time = total_loop_time + loop_time
                print('Loop time: ' + str(round(loop_time, 3)) + ' Seconds')
                print('-------------------------')
                
except KeyboardInterrupt:
    print('\nDone')

# CLOSE
# ---------------

pyautogui.click(347, 57)
time.sleep(2)
# wait for window to close.
while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 1:
    time.sleep(1)
time.sleep(5)

end_total_time = time.time()
elapsed_time = round(end_total_time - start_total_time, 3)
minutes = 0
loop_minutes = 0
while elapsed_time >= 60:
    elapsed_time = elapsed_time - 60
    minutes = minutes + 1
average_loop_time = total_loop_time / loop_counter
while average_loop_time >= 60:
    average_loop_time = average_loop_time - 60
    loop_minutes = loop_minutes + 1
print('\nAverage time per loop: ' +str(loop_minutes) + ' Minutes, ' + str(round(average_loop_time, 3)) + ' Seconds')
print('\nElapsed time: ' + str(minutes) + ' Minutes, ' + str(round(elapsed_time, 3)) + ' Seconds')

print('\nCompleted.')
