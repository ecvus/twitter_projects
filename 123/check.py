import cv2
import pyautogui
import numpy as np
import time
from pynput.mouse import Button, Controller

# Загрузить шаблоны
template1 = cv2.imread('template.png', 0)
w1, h1 = template1.shape[::-1]

template2 = cv2.imread('template2.png', 0)
w2, h2 = template2.shape[::-1]

# Сделать снимок экрана
screenshot = pyautogui.screenshot()
screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

# Преобразовать скриншот в оттенки серого
screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

# Использовать шаблонное сопоставление для поиска шаблонов на скриншоте
res1 = cv2.matchTemplate(screenshot_gray, template1, cv2.TM_CCOEFF_NORMED)
min_val1, max_val1, min_loc1, max_loc1 = cv2.minMaxLoc(res1)

res2 = cv2.matchTemplate(screenshot_gray, template2, cv2.TM_CCOEFF_NORMED)
min_val2, max_val2, min_loc2, max_loc2 = cv2.minMaxLoc(res2)

# Найти центры шаблонов
center1 = (max_loc1[0] + w1 // 2, max_loc1[1] + h1 // 2)
center2 = (max_loc2[0] + w2 // 2, max_loc2[1] + h2 // 2)

# Кликнуть по центрам шаблонов
mouse = Controller()

mouse.position = (center1)

