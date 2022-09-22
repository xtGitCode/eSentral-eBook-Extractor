import pyautogui
import os
import os.path
import tkinter as tk
from TkinterDnD2 import DND_FILES, TkinterDnD
from tkinter import ttk
import easygui
import time
from PIL import ImageGrab
from utils.utils import next_button, next_page
from utils.utils import doubledig
import pytesseract
import glob
import re
from utils.processing import preprocessing
from utils.processing import splitTwo
from utils.utils import replace_text
from utils.utils import split_companies
from utils.utils import text_process
from utils.utils import toExcel
from utils.utils import next_button
import os.path
from tkinter import *

pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'

def page1():
    page2text.pack_forget()
    page3text.pack_forget()
    page_label.pack_forget()
    page_entry.pack_forget()
    inst.pack_forget()
    inst2.pack_forget()
    textlabel.pack_forget()
    folder_entry.pack_forget()
    file_button.pack_forget()
    clear_button.pack_forget()
    convert_button.place_forget()
    startconvert_button.place_forget()
    page1text.pack(anchor=tk.W)
    show_folder()
    page_label.pack(anchor=tk.W,padx=10, pady=5)
    page_entry.pack(anchor=tk.W,padx=10)
    inst.pack(anchor=tk.W,padx=10)
    inst2.pack(anchor=tk.W,padx=10)
    sc_button.place(relx=0.33, rely=0.92)


def page2():
    page1text.pack_forget()
    page3text.pack_forget()
    page_label.pack_forget()
    page_entry.pack_forget()
    inst.pack_forget()
    inst2.pack_forget()
    textlabel.pack_forget()
    folder_entry.pack_forget()
    file_button.pack_forget()
    clear_button.pack_forget()
    sc_button.place_forget()
    startconvert_button.place_forget()
    page2text.pack(anchor=tk.W)
    show_folder()
    convert_button.place(relx=0.33, rely=0.92)

def page3():
    page1text.pack_forget()
    page2text.pack_forget()
    page_label.pack_forget()
    page_entry.pack_forget()
    inst.pack_forget()
    inst2.pack_forget()
    textlabel.pack_forget()
    folder_entry.pack_forget()
    file_button.pack_forget()
    clear_button.pack_forget()
    sc_button.place_forget()
    convert_button.place_forget()
    page3text.pack(anchor=tk.W)
    show_folder()
    page_label.pack(anchor=tk.W,padx=10, pady=5)
    page_entry.pack(anchor=tk.W,padx=10)
    inst.pack(anchor=tk.W,padx=10)
    inst2.pack(anchor=tk.W,padx=10)
    startconvert_button.place(relx=0.20, rely=0.92)

window = TkinterDnD.Tk()
window.geometry("500x620")

page1btn = tk.Button(window, text="Screenshot", width = 10,command=page1)
page2btn = tk.Button(window, text="Convert", width = 10,command=page2)
page3btn = tk.Button(window, text="Both", width = 10,command=page3)

page1text = tk.Label(window, text="SCREENSHOT")
page2text = tk.Label(window, text="CONVERT")
page3text = tk.Label(window, text="BOTH")

page1btn.pack(fill='both')
page2btn.pack(fill='both')
page3btn.pack(fill='both')


def display(string):
    label = tk.Label(
        window,
        text=string)
    label.pack()

def fleet_display(top,string):
    label = tk.Label(
        top,
        text=string)
    label.pack(anchor=tk.CENTER,pady=15)
    top.after(1000,label.destroy)  

def convert_txt():
    filepath = folder_entry.get() + '\*.jpg'
    extract_txt(filepath)

def extract_txt(path):
    folderpath = folder_entry.get()
    pages = glob.glob(path)
    totalimg = len(pages)
    incr = 100 / totalimg
    progress = incr
    pages.sort()

    text2=''
    total = []
    top = Toplevel(window)
    top.geometry("400x150")
    top.title("Conversion Progress")
    for page in pages:
        text = ''
        filetext = ''
        companies = []
        fleet_display(top,'converting: {}'.format(page))
        fleet_display(top,'{:.2f}%'.format(progress))
        window.update()
        page_num = re.findall(r"page(\d*).jpg",page)[0]
        page1, page2 = splitTwo(page)
        page1 = preprocessing(page1)
        page2 = preprocessing(page2)
        image_string1 = pytesseract.image_to_string(page1, config='--oem 3 --psm 6')
        image_string2 = pytesseract.image_to_string(page2, config='--oem 3 --psm 6')
        text += image_string1 + image_string2
        text2 += image_string1 + image_string2
        filetext = replace_text(text)
        text2 = replace_text(text2)
        companies = split_companies(filetext)
        total = text_process(total,companies,page_num)
        progress = progress + incr
        window.update_idletasks()
    top.destroy()
    txtfile = "result.txt"
    completeName = os.path.join(folderpath, txtfile) 
    with open(completeName, mode = 'w') as f:
        f.write(text2)

    # convert to xlsx
    toExcel(total,folderpath)

    display('Extraction Completed')
    tk.messagebox.showinfo("Message","result.txt and result.xlsx are generated")
    return

def startconvert_sc():
    filepath = folder_entry.get()

    #get coordinates 
    w,h = pyautogui.size() #get screen size
    x1 = int(w * 0.268)
    x2 = int(w * 0.728)

    x4 = int(w * 0.975)
    y4 = int(h * 0.386)
    x5 = int(w * 0.485)
    y5 = int(h * 0.269)
    pyautogui.click(x4,y4) #to make gui dissappear
    pyautogui.click(clicks=2, x=x5,y=y5) #fullscreen
    time.sleep(5) #wait for press esc warning to fade
    x3,y3 = next_button(filepath)
    pyautogui.moveTo(x3,y3)

    page = page_entry.get()
    if page == 'all':
        page = '12-40,54-63,66-78,82-88,92-108,112-152,156-188,192,196-197,200-305,308-357,360-369,372-382'
    page = list(page)
    arr = doubledig(page)
    length = len(arr)
    x = 0 

    starting_page = int(arr[x])
    next_page(starting_page,x3,y3) #move to starting page
    image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc first page
    image.save(filepath + r'\page%s.jpg'%(str(arr[x]).zfill(3))) # save using current page value

    x += 1 # go to next element
    while x < length:
        if arr[x] == ',': # if element is ',' jump to next page given
            x += 1 
            next = int(arr[x]) - int(arr[x-2])
            next_page(next,x3,y3) # jump to next page
            image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc 
            image.save(filepath + r'\page%s.jpg'%(str(arr[x]).zfill(3)))
        
        if arr[x] =='-': # sc all following pages
            x += 1

            cur_page = int(arr[x-2])
            save_page = cur_page + 1
            while cur_page < int(arr[x]):
                next_page(1,x3,y3)
                image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc 
                image.save(filepath + r'\page%s.jpg'%(str(save_page).zfill(3)))
                save_page += 1
                cur_page += 1
        x += 1
    tk.messagebox.showinfo("Message","Sreenshot Completed\nCurrently converting to xlsx file")
    convert_txt()

def drop_inside_box(event):
    folder_entry.insert("end",event.data)

def open_file():
    folder=easygui.diropenbox()
    folder_entry.insert("end",folder)

def start_sc():
    filepath = folder_entry.get()

    #get coordinates 
    w,h = pyautogui.size() #get screen size
    x1 = int(w * 0.268)
    x2 = int(w * 0.728)

    x4 = int(w * 0.975)
    y4 = int(h * 0.386)
    x5 = int(w * 0.485)
    y5 = int(h * 0.269)
    pyautogui.click(x4,y4) #to make gui dissappear
    pyautogui.click(clicks=2, x=x5,y=y5) #fullscreen
    time.sleep(5) #wait for press esc warning to fade
    x3,y3 = next_button(filepath)
    pyautogui.moveTo(x3,y3)

    page = page_entry.get()
    if page == 'all':
        page = '12-40,54-63,66-78,82-88,92-108,112-152,156-188,192,196-197,200-305,308-357,360-369,372-382'
    page = list(page)
    arr = doubledig(page)
    length = len(arr)
    x = 0 

    starting_page = int(arr[x])
    next_page(starting_page,x3,y3) #move to starting page
    image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc first page
    image.save(filepath + r'\page%s.jpg'%(str(arr[x]).zfill(3))) # save using current page value

    x += 1 # go to next element
    while x < length:
        if arr[x] == ',': # if element is ',' jump to next page given
            x += 1 
            next = int(arr[x]) - int(arr[x-2])
            next_page(next,x3,y3) # jump to next page
            image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc 
            image.save(filepath + r'\page%s.jpg'%(str(arr[x]).zfill(3)))
        
        if arr[x] =='-': # sc all following pages
            x += 1

            cur_page = int(arr[x-2])
            save_page = cur_page + 1
            while cur_page < int(arr[x]):
                next_page(1,x3,y3)
                image = ImageGrab.grab(bbox=(x1,0,x2,h)) #sc 
                image.save(filepath + r'\page%s.jpg'%(str(save_page).zfill(3)))
                save_page += 1
                cur_page += 1
        x += 1
    tk.messagebox.showinfo("Message","Sreenshot Completed")
    

#convert button
startconvert_button = tk.Button(
    window,
    text='Screenshot + Convert',
    width=18,
    bg='green',
    fg='white',
    command = startconvert_sc
)
convert_button = tk.Button(
    window,
    text='Convert',
    width=7,
    bg='green',
    fg='white',
    command = convert_txt)

sc_button = tk.Button(
    window,
    text = "Start",
    width=7,
    bg='green',
    fg='white',
    command = start_sc,
    )

cancel_button = tk.Button(
    window,
    text = "Quit",
    width=7,
    bg='red',
    fg='white',
    command = window.destroy
    )

#drop folder
testvariable = tk.StringVar()
textlabel = tk.Label(window,
                  text='Folder: ')
folder_entry = tk.Entry(window,
                 textvar=testvariable,
                 width=90)
folder_entry.drop_target_register(DND_FILES)
folder_entry.dnd_bind('<<Drop>>', drop_inside_box)

def clear_text():
   folder_entry.delete(0, END)

#open explorer button
file_button = tk.Button(
    window,
    text='Choose Folder',
    command = open_file)
clear_button = tk.Button(
    window,
    text = 'Clear',
    command=clear_text
)

page_label = tk.Label(text="Pages")
page_entry = tk.Entry()


#instructions
inst = tk.Label(
    window,
    text='eg. 1-5,8,10-13'
)

inst2 = tk.Label(
    window,
    text = 'Pages for companies in:\n Johor (12-40), Kedah (54-63), Kelantan (66-78), Melaka (82-88),\nNegeri Sembilan (92-108), Pahang (112-152), Perak (156-188), \nPerlis (192), Pulau Pinang (196-197), Sabah (200-305), \nSarawak (308,357), Selangor (360-369), Terengganu (372-382)\nType all for all companies pages'
)

#show
cancel_button.place(relx=0.53, rely=0.92)

#def
def show_folder():
    textlabel.pack (anchor=tk.W,padx=10, pady=5)
    folder_entry.pack(fill=tk.X, padx=10, ipady=10)
    file_button.pack(anchor=tk.W,padx=10,pady=10)
    clear_button.pack(anchor=tk.W,padx=10)


window.mainloop()