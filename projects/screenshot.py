import pyautogui
import pynput
import os.path 
from pynput.mouse import Listener

def on_click(x, y, button, pressed):
    i=0
    while os.path.exists(r'C:\Users\xiaot\Downloads\Screenshots\image%s.jpg'%i):
        i += 1
    pyautogui.screenshot(r'C:\Users\xiaot\Downloads\Screenshots\image%s.jpg'%i)


with Listener(on_click=on_click) as listener:
    listener.join()
