import time
import pyautogui
import keyboard

# Set the delay between hotkey presses (in seconds)
delay = 0.5

# Set the confidence level for image matching
confidence = 0.9

# Set the list of image paths to search for
image_paths = ['image1.png', 'image2.png', 'image3.png']

def click_on_image():
    for image_path in image_paths:
        try:
            # Search for the image on the screen
            image_location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            
            if image_location:
                # Move the mouse to the center of the image
                x, y = pyautogui.center(image_location)
                pyautogui.moveTo(x, y, duration=0)
                
                # Click on the image
                pyautogui.click()
                print(f"Clicked on {image_path}")
                return
        except pyautogui.ImageNotFoundException:
            print(f"Could not find {image_path} on the screen")
    print("No image found")

def hotkey_function():
    time.sleep(delay)
    click_on_image()

# Bind hotkey to function
keyboard.add_hotkey('alt', hotkey_function)

print("Press 'alt' to click on one of the images")
keyboard.wait()