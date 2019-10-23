# MO release using Excel

import pyautogui
import pygetwindow as gw
import time
import openpyxl

pyautogui.PAUSE = 1

print('Press Ctrl-C to quit.')

try:
    start_total_time = time.time()
    start = int(input('What row do you want to start on?: '))
    end = int(input('What row do you want to end on?: '))
    #job_number = input('What is the job number?: ')
    path = 'C:\\Users\\jlee.NTPV\\Desktop\\MO-Release\\HT-202NC MASTER PRODUCTION SCHEDULE.xlsx'
    wb = openpyxl.load_workbook(path)
    #wb.get_sheet_names()
    #sheet = wb.get_sheet_by_name(job_number)
    sheet = wb.active

    # COORDINATES
    #------------------
    # New           360, 87
    # MO No.        210, 100    C
    # Description   234, 125    E
    # Build Item    261, 180    D
    # Job No.       228, 313
    # Location No.  227, 338
    # Customer      237, 362
    # Ordered       242, 423    
    # Order Date    597, 439
    # Start Date    600, 473    I
    # Comp. Date    599, 497    J

    # VARIABLES
    #----------------

    #end = sheet.max_row
    job_range = range(start, end+1, 1)
    loop_counter = 0
    total_loop_time = 0

    # LOOP
    # ---------------

    for i in job_range:
        start_loop_time = time.time()
        print('Current value = ' + str(i))
        mo_number = sheet.cell(row = i, column = 3)
        description = sheet.cell(row = i, column = 5)
        build_item = sheet.cell(row = i, column = 4)
        job_number = mo_number
        location_number = str('NTPV')
        customer = str('Kinder Equipment Co')
        ordered = str('1')
        start_date = sheet.cell(row = i, column = 9)
        completion_date = sheet.cell(row = i, column = 10)

        print('MO = ' + str(mo_number.value))
        print('Description = ' + str(description.value))
        print('Build Item = ' + str(build_item.value))
        print('Start Date = ' + str(start_date.value))
        print('Completion Date = ' + str(completion_date.value))

    # NEW
    #----------------
        print('-------------------------')
        #input('Press ENTER to continue.')
        #print('-------------------------')
        pyautogui.click(360, 87)

    # WAIT FOR MO WINDOW
    # ---------------
        
        while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 1:
            time.sleep(2)
        time.sleep(8)

    # MO NO.
    # ---------------

        pyautogui.doubleClick(210, 100)
        pyautogui.typewrite(str(mo_number.value))
        
    # DESCRIPTION
    # ---------------

        pyautogui.doubleClick(234, 125)
        pyautogui.typewrite(str(description.value))

    # BUILD ITEM
    # ---------------

        pyautogui.doubleClick(261, 180)
        pyautogui.typewrite(str(build_item.value))

    # JOB NUMBER
    # ---------------

        pyautogui.doubleClick(228, 313)
        pyautogui.typewrite(str(job_number.value))
        # tab
        pyautogui.click(237, 338)
        time.sleep(2)
        pyautogui.click(928, 537)
        pyautogui.typewrite(str(description.value))
        pyautogui.click(888, 588)
        time.sleep(2)

    # LOCATION
    # ---------------

        pyautogui.doubleClick(227, 338)
        pyautogui.typewrite(location_number)

    # CUSTOMER (LIFTROCK)
    # ---------------

        pyautogui.doubleClick(237, 362)
        pyautogui.typewrite(customer)

    # ORDERED
    # ---------------

        pyautogui.doubleClick(242, 423)
        pyautogui.typewrite(ordered)
        # tab
        pyautogui.click(237, 338)
        # CLICK YES
        pyautogui.click(1025, 595)
        # wait for selcet screen
        while len(gw.getWindowsWithTitle('Select Revision to Use - Bills of Material')) < 1:
            time.sleep(2)
        time.sleep(1)
        pyautogui.click(1166, 762)
        # wait for select screen to exit
        while len(gw.getWindowsWithTitle('Select Revision to Use - Bills of Material')) == 1:
            time.sleep(2)
        time.sleep(1)

    # COMPLETION DATE
    # ---------------

        pyautogui.doubleClick(599, 497)
        pyautogui.typewrite(str(completion_date.value))

    # START DATE
    # ---------------

        pyautogui.doubleClick(600, 473)
        pyautogui.typewrite(str(start_date.value))

    
    # CLICK DESCRIPTION BOX
    # ---------------

        pyautogui.click(234, 125)

    # SAVE
    #----------------
        #input()
        pyautogui.click(76, 57)

    # RELEASE
    # ---------------

        pyautogui.click(427, 58)

    # CLICK OK
    # ---------------

        while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 1:
            time.sleep(1)
        pyautogui.click(1082, 595)
        time.sleep(5)

    # CLOSE
    # ---------------
        
        pyautogui.click(347, 57)
        time.sleep(2)
        # wait for window to close.
        while len(gw.getWindowsWithTitle('Manufacturing Orders - North Texas Pressure Vessels Inc.')) == 1:
            time.sleep(1)
        time.sleep(2)
        pyautogui.click(1889, 1007)
        #else:
         #   input('Press ENTER to continue.')
          #  print('-------------------------')
        loop_time = time.time() - start_loop_time
        loop_counter = loop_counter + 1
        total_loop_time = total_loop_time + loop_time
        print('Loop time: ' + str(round(loop_time, 3)) + ' Seconds')
        print('-------------------------')
        
except KeyboardInterrupt:
    print('\nDone')

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
