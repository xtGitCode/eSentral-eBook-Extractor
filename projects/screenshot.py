import pyautogui
import pynput
import os
import os.path 
from pynput.mouse import Listener

filepath = os.getenv('APPDATA') + "\Screenshot"
if not os.path.exists(filepath):
    os.makedirs(filepath)
    
def on_click(x, y, button, pressed):
    i=0
    while os.path.exists(filepath + r'\image%s.jpg'%i):
        i += 1
    pyautogui.screenshot(filepath + r'\image%s.jpg'%i)


with Listener(on_click=on_click) as listener:
    listener.join()