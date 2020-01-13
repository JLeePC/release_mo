# MISys semi-auto MO

import pyautogui
import pygetwindow as gw
import time

pyautogui.PAUSE = 0.005

print('Press Ctrl-C to quit.')

job = input('Whats the base job?: ')

range_start = int(input('Enter the range start: '))

range_ceiling = int(input('Enter the range ceiling: '))+1

z_fill = input('Do you need zfill? (1/0): ')

placeholder = []
num_to_skip = []
job_range = []
stop_loop = False
skip_me = str(input("Do you have numbers to skip? (Y/N): "))
if 'Y' in skip_me or 'y' in skip_me:
    # stop_loop is a secondary measure to prevent infinite loops, not required, but precautionary
    while not stop_loop:
        user_input = input("Please enter the number you would like to skip (enter STOP to quit): ")
        try:
            if 'STOP' in str(user_input) or 'stop' in str(user_input):
                stop_loop = True
                break
        except ValueError:
            continue
            
        try:
            placeholder.append(int(user_input))
        except ValueError:
            print("Please enter a valid number or STOP to quit")
            continue
    # We need to remove possible duplicates
    for num in placeholder:
        if num not in num_to_skip:
            num_to_skip.append(num)
    
    # Now we want to build a disjointed list to make the future for loop 1000 times easier

    temp_range = range(range_start,range_ceiling,1)
    disjointer_a = [number for number in num_to_skip if number not in temp_range]
    disjointer_b = [number for number in temp_range if number not in num_to_skip]
    
    # Combining the two lists to make the completed iteration
    job_range = disjointer_a + disjointer_b

# Just checking to see if it's empty, that way we won't error out in future
if not job_range:
    job_range = range(range_start,range_ceiling,1)


try:
    start_time = time.time()
    print('Sart Time')
    loop_counter = 0
    loop_time = 0
    print('------------------------------')
    for position_in_range in job_range:

        start_loop = time.time()

        if '1' in z_fill:            
            print("Current position: {}".format(str(position_in_range).zfill(2)))
            mo = "{0}-{1}".format(str(job),(str(position_in_range).zfill(2)))
        else:
            print("Current position: {}".format(position_in_range))
            mo = "{0}-{1}".format(str(job),str(position_in_range))
            
        time.sleep(0.5)
        print(mo)
        pyautogui.click(1114,573)
        pyautogui.click(490,1008)
        statusWindow = gw.getWindowsWithTitle('Filter Conditions')
        while len(gw.getWindowsWithTitle('Filter Conditions')) == 0:
            time.sleep(1)
        time.sleep(1)
        pyautogui.doubleClick(824,446)
        time.sleep(1)
        pyautogui.typewrite(str(mo))
        time.sleep(1)
        pyautogui.click(1158,728)
        statusWindow = gw.getWindowsWithTitle('MANUFACTURING ORDER DOCUMENT DISTRIBUTION')
        while len(gw.getWindowsWithTitle('MANUFACTURING ORDER DOCUMENT DISTRIBUTION')) == 0:
            time.sleep(1)
        time.sleep(1)
        pyautogui.click(385,36)
        statusWindow = gw.getWindowsWithTitle('Print Options')
        while len(gw.getWindowsWithTitle('Print Options')) == 0:
            time.sleep(1)
        time.sleep(1)
        pyautogui.click(1138,464)
        time.sleep(1)
        pyautogui.click(1096,543)
        time.sleep(1)
        pyautogui.click(1127,620)
        statusWindow = gw.getWindowsWithTitle('Save Print Output As')
        while len(gw.getWindowsWithTitle('Save Print Output As')) == 0:
            time.sleep(1)
        time.sleep(1)
        pyautogui.click(631,591)
        pyautogui.typewrite(str(mo)+' DD')
        pyautogui.click(1084,660)
        time.sleep(1)
        pyautogui.click(1895,9)
        stop_loop = round(time.time() - start_loop, 3)
        loop_counter = loop_counter + 1
        loop_time = loop_time + stop_loop
        print('Loop time: ' + str(stop_loop) + ' Seconds')
        print('------------------------------')
        time.sleep(4)

except KeyboardInterrupt:
    print('\nCanceled')

end_time = time.time()
elapsed_time = round(end_time - start_time, 3)
minutes = 0
while elapsed_time >= 60:
    elapsed_time = elapsed_time - 60
    minutes = minutes + 1
average_loop = loop_time / loop_counter
print('\nAverage time per loop: ' + str(round(average_loop, 3)) + ' Seconds')
print('\nElapsed time: ' + str(minutes) + ' Minutes ' + str(round(elapsed_time, 3)) + ' Seconds')

print('\nCompleted.')
