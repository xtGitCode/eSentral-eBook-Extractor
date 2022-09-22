import re
import pyautogui
import time
import pandas as pd
from utils.spellcheck import spell
from utils.spellcheck import spell_state
import os.path
import xlsxwriter
import cv2
import numpy as np
import os

def replace_text(filetext):
    filetext = filetext.replace('\n\n',"\n")
    filetext = filetext.replace('MPO8',"MPOB")
    filetext = filetext.replace('Mpob',"MPOB")
    filetext = filetext.replace('IMPOB',"MPOB")
    filetext = filetext.replace('MPQB',"MPOB")
    filetext = filetext.replace('MPGB',"MPOB")
    filetext = filetext.replace('Liconse',"License")
    filetext = filetext.replace('50397 1402000',"503971402000")
    filetext = filetext.replace('Licsnse',"License")
    filetext = filetext.replace('Licenso',"License")
    filetext = filetext.replace('Licensa',"License")
    filetext = filetext.replace('licensa',"License")
    filetext = filetext.replace('hcense ',"License ")
    filetext = filetext.replace('Licanse',"License")
    filetext = filetext.replace('Liconso',"License")
    filetext = filetext.replace('Licarse',"License")
    filetext = filetext.replace('license',"License")
    filetext = filetext.replace('Licerse',"License")
    filetext = filetext.replace('Licansa',"License")
    filetext = filetext.replace('Ticense',"License")
    filetext = filetext.replace('Te!ephone',"Telephone")
    filetext = filetext.replace('Teephone',"Telephone")
    filetext = filetext.replace('Telepone',"Telephone")
    filetext = filetext.replace('Teleptvone',"Telephone")
    filetext = filetext.replace('Emait',"Email")
    filetext = filetext.replace('Email ',"Email:")
    filetext = filetext.replace('Ema :',"Email:")
    filetext = filetext.replace('Te\'ephona',"Telephone")
    filetext = filetext.replace('Te\'ophone',"Telephone")
    filetext = filetext.replace('Telepnione',"Telephone")
    filetext = filetext.replace('Teleptione',"Telephone")
    filetext = filetext.replace('Telephon ',"Telephone ")
    filetext = filetext.replace('Telaphon ',"Telephone ")
    filetext = filetext.replace('Tolophono',"Telephone")
    filetext = filetext.replace('Tolophione',"Telephone")
    filetext = filetext.replace('Tolophione',"Telephone")
    filetext = filetext.replace('Talepthione',"Telephone")
    filetext = filetext.replace('Teleptrone',"Telephone")
    filetext = filetext.replace('Teloptrone',"Telephone")
    filetext = filetext.replace('Te\'ephone',"Telephone")
    filetext = filetext.replace('Telephones',"Telephone")
    filetext = filetext.replace('Teltaphone',"Telephone")
    filetext = filetext.replace('Adress',"Address")
    filetext = filetext.replace('Addess',"Address")
    filetext = filetext.replace('address',"Address")
    filetext = filetext.replace('Poti ',"Peti ")
    filetext = filetext.replace('Pet} ',"Peti ")
    filetext = filetext.replace('Jin.',"Jln.")
    filetext = filetext.replace('Jin ',"Jln ")
    filetext = filetext.replace('Jatan ',"Jalan ")
    filetext = filetext.replace('Jaian',"Jalan")
    filetext = filetext.replace('SON.',"SDN.")
    filetext = filetext.replace('SDN,',"SDN.")
    filetext = filetext.replace('Sdn,',"Sdn.")
    filetext = filetext.replace('BHD,',"BHD.")
    filetext = filetext.replace('Kiuang,',"Kluang")
    filetext = filetext.replace('Joho:',"Johor")
    filetext = filetext.replace('Ist ',"1st ")
    filetext = filetext.replace('Iskandar Puten',"Iskandar Puteri")
    filetext = filetext.replace('Buioh',"Buloh")
    filetext = filetext.replace('Sn ',"Sri ")
    filetext = filetext.replace('Puiai ',"Pulai ")
    filetext = filetext.replace('Kulal',"Kulai")
    filetext = filetext.replace('Dangar',"Dengar")
    filetext = filetext.replace('Kluana',"Kluang")
    filetext = filetext.replace('Klueng',"Kluang")
    filetext = filetext.replace('Kota Tinggl',"Kota Tinggi")
    filetext = filetext.replace('Kots Tinggi',"Kota Tinggi")
    filetext = filetext.replace('Pet Surat',"Peti Surat")
    filetext = filetext.replace('Pojabat',"Pejabat")
    filetext = filetext.replace('Pejabal',"Pejabat")
    filetext = filetext.replace('Pojabal',"Pejabat")
    filetext = filetext.replace('Pejabst',"Pejabat")
    filetext = filetext.replace('Pefabat',"Pejabat")
    filetext = filetext.replace('Borhad',"Berhad")
    filetext = filetext.replace('Bertrad',"Berhad")
    filetext = filetext.replace('Pesabat',"Pejabat")
    filetext = filetext.replace('Peyabat',"Pejabat")
    filetext = filetext.replace('Bokit',"Bukit")
    filetext = filetext.replace('Posta!',"Postal")
    filetext = filetext.replace('Posla!',"Postal")
    filetext = filetext.replace('Postall',"Postal")
    filetext = filetext.replace('Postal!',"Postal")
    filetext = filetext.replace('Poste!',"Postal")
    filetext = filetext.replace('Posia!',"Postal")
    filetext = filetext.replace('Postai',"Postal")
    filetext = filetext.replace('KAW,',"KAW.")
    filetext = filetext.replace('FELDAAIR',"FELDA AIR")
    filetext = filetext.replace('Perndusiran',"Perindustrian")
    filetext = filetext.replace('Pasta!',"Postal")
    filetext = filetext.replace('5008 12602000',"500812602000")
    filetext = filetext.replace('Sagamat',"Segamat")
    filetext = filetext.replace('SUNGA!',"SUNGAI")
    filetext = filetext.replace('KOPERAS! ',"KOPERASI ")
    filetext = filetext.replace('KOPERAS!I ',"KOPERASI ")
    filetext = filetext.replace('AGROMACC8',"AGROMACC")
    filetext = filetext.replace('10! Corporation Berhad',"IOI Corporation Berhad")
    filetext = filetext.replace('Telephone no',"Telephone No")
    filetext = filetext.replace('OF-',"07-")
    filetext = filetext.replace('O7-',"07-")
    return filetext

 

def split_companies(filetext):
    companies = re.split(r"(.*)(?:\n)(?:.*)(?:MPOB)",filetext)
    length = len(companies)
    x = 0
    while x < length-1:
        count = len(companies[x])
        if count < 100:
            if count == 0:
                companies.pop(x) # if element is empty delete it
                length = length-1
                x -= 1
            else:
                companies[x]=''.join(companies[x:x+2]) # else join with next element
                companies.pop(x+1) # delete next element
                length = length-1
        x += 1
    return companies

def text_process(total,companies,page_num):
    for company in companies:
        companies_dict={}

        #Page
        companies_dict['Page Number']=page_num

        #Name
        if re.findall(r".*.(?=License)", company):
            c_name = re.findall(r".*.(?=License)", company)[0]
        else:
            c_name = "NA"
        try:
            c_name = re.findall(r"[a-zA-Z].*",c_name)[0]
        except Exception:
            pass
        c_name =c_name.strip()
        companies_dict['Name']=c_name

        #MPOB
        if re.findall(r"(\d{11,12})",company):
            c_mpob = re.findall(r"(\d{11,12})",company)[0]
        else:
            c_mpob = "NA"
        c_mpob =c_mpob.strip()
        companies_dict['MPOB']=c_mpob

        #Parent Company
        if re.findall(r"P\w{5} C\w{6}..(.*)",company):
            c_parent = re.findall(r"P\w{5} C\w{6}..(.*)",company)[0]
        else:
            c_parent = "NA"
        try: 
            c_p = re.findall(r"[a-zA-Z0-9\s.(),&]+",c_parent)[0]
        except Exception:
            pass
        if c_p == ' ':
            c_p = re.findall(r"[a-zA-Z0-9\s.(),&]+",c_parent)[1]
        c_p =c_p.strip()
        companies_dict['Parent Company']=c_p

        # change to no newline for address
        address = company.replace('\n',"")

        #Address 
        if re.findall(r"(?:P\w{5}(?:\s|)A\w{6}.)(.*)(?=T\w{8})",address):
            c_address = re.findall(r"(?:P\w{5}(?:\s|)A\w{6}.)(.*)(?=T\w{8})",address)[0]
            c_address = c_address.strip()
            c_address = re.findall(r"[a-zA-Z0-9\s.,\/&-].*",c_address)[0]
            c_a =c_address.strip()
            # spell check last word in address
            if re.findall(r"\w*\s\w*$",c_a):
                state = re.findall(r"\w*\s\w*$",c_a)[0]
                state = state.strip()
                try:
                    c_a = c_a.rsplit(re.findall(r"\w*\s\w*$",c_a)[0], 1)[0]
                except Exception:
                    pass
                state = spell_state(state)
                state = state.strip()
                c_address = c_a + ' ' + state
            else: c_address = c_a
        else:
            c_address = "NA"

        companies_dict['Address']=c_address

        #Telephone
        if re.findall(r"T\w{8}\sN\w.(.*)",company):
            c_telephone = re.findall(r"T\w{8}\sN\w.(.*)",company)[0]
            c_telephone = c_telephone.replace(" ","")
            if re.findall(r"(?:0|O).*",c_telephone):
                c_telephone=re.findall(r"(?:0|O).*",c_telephone)[0]
        else:
            c_telephone = "NA"
        
        c_t =c_telephone.strip()
        companies_dict['Telephone']=c_t

        #Fax
        if re.findall(r"F\w{2}(?:\s|)N\w.(.*)",company):
            c_fax = re.findall(r"F\w{2}(?:\s|)N\w.(.*)",company)[0]
            c_fax = c_fax.replace(" ","")
            try:
                c_fax=re.findall(r"\d.*",c_fax)[0]
            except Exception:
                pass            
            if re.findall(r"(?:0|O).*",c_fax):
                c_fax=re.findall(r"(?:0|O).*",c_fax)[0]
        else:
            c_fax = "NA"
        c_f =c_fax.strip()
        companies_dict['Fax']=c_f

        #Email
        if re.findall(r".*@.*.(?:c|n|m).*",company):
            c_email = re.findall(r".*@.*.(?:c|n|m).*",company)[0]
            c_email = c_email.replace(" ","")
            try:
                c_email = re.findall(r"^\w*.(.*)",c_email)[0]
            except Exception:
                pass
            try:
                c_email = re.findall(r"[a-zA-Z].*",c_email)[0]
            except Exception:
                pass
        else:
            c_email = "NA"
        c_email =c_email.strip()
        companies_dict['Email']=c_email

        #District
        if re.findall(r"Dis\w* Pr\w{5}.(.*)",company):
            c_district = re.findall(r"Dis\w* Pr\w{5}.(.*)",company)[0]
        else:
            c_district = "NA"
        try:
            c_district = re.findall(r"[a-zA-Z\s]+",c_district)[0]
        except Exception:
            pass
        c_district =c_district.strip()
        c_d,state = spell(c_district,page_num)
        companies_dict['District Premise']=c_d
        companies_dict['State']=state

        total.append(companies_dict)
    return total

def toExcel(total,folderpath):
    df = pd.DataFrame(data=total)
    xlxsfile = "result.xlsx"
    completeName = os.path.join(folderpath, xlxsfile)
    writer = pd.ExcelWriter(completeName, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Companies')
    workbook = writer.book
    worksheet = writer.sheets['Companies']
    error_format = workbook.add_format({'bg_color':'#FFC7CE'})
    worksheet.conditional_format('D2:D1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('C2:C1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('J2:J1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('K2:K1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('G2:G1048576',{'type':'cell','criteria':'==','value':'"NA"','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'Poti','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'£','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'{','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'}','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'$','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'°','format':error_format})
    worksheet.conditional_format('F2:F1048576',{'type':'text','criteria':'containing','value':'%','format':error_format})
    worksheet.conditional_format('G2:G1048576',{'type':'text','criteria':'containing','value':'%','format':error_format})
    worksheet.conditional_format('G2:G1048576',{'type':'text','criteria':'containing','value':'$','format':error_format})
    worksheet.conditional_format('D2:D1048576',{'type':'formula','criteria':'=LEN(D2)<12','format':error_format})
    workbook.close()
    
def next_button(filepath):
    img = pyautogui.screenshot()
    filepath = filepath + r'\sc.png'
    img.save(filepath)
    img = cv2.imread(filepath,0)
    h,w = img.shape
    blurred = cv2.GaussianBlur(img,(11,11),0)
    c_width = w * 0.78
    c_width = int(c_width)
    crop_img = blurred[1:-1, c_width:-1] 
    # Finds circles in a grayscale image using the Hough transform
    circles = cv2.HoughCircles(crop_img, cv2.HOUGH_GRADIENT, 1, 100,
                             param1=100,param2=90,minRadius=0,maxRadius=200)
    os.remove(filepath)
    if circles is not None:
        circles = np.round(circles[0, :]).astype("int")
    for (x, y, r) in circles:
        x=x+c_width
        return x,y

def next_page(int,x,y):
    i=0
    while i < int:
        pyautogui.click(x,y)
        time.sleep(1)
        i += 1
    return

def doubledig(arr):
    length = len(arr) 
    x = 0
    while x < length-1: 
        if arr[x].isdigit(): 
            x+=1 
            while (x < length) and (arr[x].isdigit()): 
                arr[x-1]=''.join(arr[x-1:x+1])
                arr.pop(x)
                length = length-1
        x+=1
    return arr


