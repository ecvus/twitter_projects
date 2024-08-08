import pyautogui
import cv2
import numpy as np
import time
import threading

# Set up the window title and image paths
window_title = "anty.exe"
image_path1 = "img/Screenshot_1.png"
image_path2 = "img/Screenshot_2.png"

# Set up the delay between actions
delay = 2

# Set up the scroll variable
scroll = False

# Set up the next page variable
next_page_go = False

def turnde_off_on():
    global switch_button
    switch_button = not switch_button
    if switch_button:
        tab_work()

def tab_work():
    global scroll
    global next_page_go
    # Get the window position and size
    window_pos = pyautogui.position()
    window_size = pyautogui.size()

    # Scroll to the top of the window
    pyautogui.scroll(-window_size[1])

    # Search for the first image
    img1 = cv2.imread(image_path1)
    img_gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    template1 = cv2.imread(image_path1, cv2.IMREAD_GRAYSCALE)
    result1 = cv2.matchTemplate(img_gray1, template1, cv2.TM_CCOEFF_NORMED)
    min_val1, max_val1, min_loc1, max_loc1 = cv2.minMaxLoc(result1)
    if max_val1 > 0.5:
        x1, y1 = max_loc1
        print(f"Image 1 found at ({x1}, {y1})")
    else:
        x1, y1 = None, None

    # Search for the second image
    img2 = cv2.imread(image_path2)
    img_gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    template2 = cv2.imread(image_path2, cv2.IMREAD_GRAYSCALE)
    result2 = cv2.matchTemplate(img_gray2, template2, cv2.TM_CCOEFF_NORMED)
    min_val2, max_val2, min_loc2, max_loc2 = cv2.minMaxLoc(result2)
    if max_val2 > 0.5:
        x2, y2 = max_loc2
        print(f"Image 2 found at ({x2}, {y2})")
    else:
        x2, y2 = None, None

    # Perform actions based on the images found
    if x1 is not None and x2 is not None:
        if y1 < y2:
            scroll = True
            first_case_repost(x1, y1)
            next_page_go = True
        else:
            scroll = True
            second_case_repost(x2, y2)
            next_page_go = True
    elif x1 is not None:
        scroll = True
        first_case_repost(x1, y1)
        next_page_go = True
    elif x2 is not None:
        scroll = True
        second_case_repost(x2, y2)
        next_page_go = True

    # Wait for a certain amount of time
    time.sleep(3.5)

def first_case_repost(x, y):
    # Perform actions for the first case
    if x == 0 and y == 0:
        print("Invalid coordinates (0, 0). Skipping click.")
        return
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.18)
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.18)
    pyautogui.hotkey('ctrl', 'w')

def second_case_repost(x, y):
    # Perform actions for the second case
    if x == 0 and y == 0:
        print("Invalid coordinates (0, 0). Skipping click.")
        return
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.18)
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.1)
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.18)
    pyautogui.moveTo(100, 100)  # Move the mouse cursor to a safe location
    pyautogui.click(x, y)
    time.sleep(0.18)
    pyautogui.hotkey('ctrl', 'w')

def next_page():
    global next_page_go
    if next_page_go:
        tab_work()

# Set up the timer
def timer_thread():
    while True:
        next_page()
        time.sleep(0.001)

threading.Thread(target=timer_thread).start()

# Start the script
switch_button = False
turnde_off_on()