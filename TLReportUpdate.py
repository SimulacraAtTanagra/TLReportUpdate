# -*- coding: utf-8 -*-
"""
Created on Fri Jan 24 09:16:07 2020

@author: sayers
"""
import pandas as pd
import pyautogui
import time
from datetime import datetime
import os
from os import path
import re
import win32com.client as win32

def str(self, item):
    print(item)

    prev, current = None, self.__iter.next()
    while isinstance(current, int):
        print (current)
        prev, current = current, self.__iter.next()

def listextract(ls,pos):
    try:
        return(ls[pos])
    except:
        return("blank")

def dfgen(df,xdf,list1,r1):
    for i in range(r1):
        df[list1[i]] = xdf.apply(listextract,args=(i,))
    return(df)

def successmail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "adavis901@york.cuny.edu"
    mail.Subject = "Successful Time and Leave Backup"
    
       
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    text = """
    Good Day,
    
    The T&L backup file has been refreshed. You can find the most recent copy on the shared HR drive with today's date in the filename    


    Best Regards,
    Shane Ayers
    Acting Human Resources Information Systems Manager
    Office of Human Resources
    York College
    The City University of New York"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <body><p>Good Day,</p>
    <p>The T&L backup file has been refreshed. You can find the most recent copy on the shared HR drive with today's date in the filename. </p>
    <p></p>
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
    # above line took every col inside csv as list
    mail.Body = text
    mail.HTMLBody = html
    mail.Send()

def failmail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "adavis901@york.cuny.edu"
    mail.Subject = "Successful Time and Leave Backup"
    
       
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    text = """
    Good Day,
    
    Today's backup has failed. Please notify IT as this may indicate that the Time and Leave system is down.


    Best Regards,
    Shane Ayers
    Acting Human Resources Information Systems Manager
    Office of Human Resources
    York College
    The City University of New York"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <body><p>Good Day,</p>
    <p>Today's backup has failed. Please notify IT as this may indicate that the Time and Leave system is down. </p>
    <p></p>
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
    # above line took every col inside csv as list
    mail.Body = text
    mail.HTMLBody = html
    mail.Send()

def listofliststrip(Z):
    
    for x,b in enumerate(Z):
        for i, a in enumerate(b):
            b[i]=  ''.join(re.sub(r',\s,+|,,',',',re.sub(' +', ' ',re.sub('Days|Holidays|\n|Comp|Hrs|Mns|Annual|Sick|Leave|Balance|Ending', ',',a))))
        Z[x] = ''.join(b)
    return(Z)
def liststrip(Z):
    for x,b in enumerate(Z):
        Z[x]=re.sub(r',\s,+|,,',',',re.sub(' +', ' ',re.sub('Days|Holidays|\n|Comp|Hrs|Mns|Annual|Sick|Leave|Balance|Ending| ', ',',b)))
    Z = [i.split(',') for i in Z if i] 
    for n,x in enumerate(Z):
        Z[n] =[i for i in x if i]
    return(Z)
def main(un,pass1,pass2):
    time.sleep(5)
    os.startfile('U:\Fulltime2016.accdb')
    
    """
    #Size(width=1920, height=1080)
    useridfield = [598,206]
    pwordfield = [592,243]
    loginbutton = [618, 280]
    reportsbutton = [929,646]
    
    
    reports_FYear = [691,413]
    reports_printnamesummary = [909,523]
    reports_textexport = [752,91]
    """
    #reports_tlreportclose = [550,219]
    time.sleep(1)
    #pyautogui.click(useridfield) 
    pyautogui.typewrite(un)
    #pyautogui.click(pwordfield) 
    pyautogui.press('tab')
    pyautogui.typewrite(pass1)
    pyautogui.press('enter')
    #pyautogui.click(loginbutton)
    time.sleep(1)
    pyautogui.click(1096,501)           # clicking a random place to put the cursor on buttons
    pyautogui.press('tab',presses=5)    # navigating to the reports button
    pyautogui.press('enter')            # opening reports
    #pyautogui.click(reportsbutton)
    time.sleep(1)
    pyautogui.press('tab')              # navigating to the year field
    #pyautogui.click(reports_FYear)
    pyautogui.typewrite('2020')
    time.sleep(1)
    pyautogui.press('tab',presses=6)    # navigating to the Print Name Summary button
    pyautogui.press('enter')            # executing this report type
    time.sleep(2) 
    #pyautogui.click(reports_printnamesummary)
    pyautogui.typewrite('2020')
    pyautogui.press('enter')
    time.sleep(10) 
    pyautogui.press(['alt','down'])
    pyautogui.press('left',presses=5)
    pyautogui.press('enter')
    time.sleep(1) 
    #pyautogui.click(reports_textexport)
    #pyautogui.click(loginbutton)
    curtitle = f"s:\documents\TLSummary_{datetime.now().month}_{datetime.now().day}_{datetime.now().year}.txt"
    #reports_export_filename = [718,390]
    #pyautogui.click(reports_export_filename)
    pyautogui.hotkey('ctrl','a')
    pyautogui.typewrite(curtitle)
    #reports_export_save = [1158,768]
    time.sleep(4)
    pyautogui.press(['tab','tab','tab','enter','y','enter'])
    time.sleep(120)
    pyautogui.press('k')
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl','f4')
    #pyautogui.click(reports_tlreportclose)
    time.sleep(2)
    #reports_exit = [11367,281]
    #pyautogui.click(reports_exit)
    pyautogui.hotkey('ctrl','f4')
    #tlquit = [995,686]
    time.sleep(2)
    pyautogui.press(['tab','enter'])
    #pyautogui.click(tlquit)
    
    if path.isfile(curtitle) == True:
        rows = []
        with open (curtitle) as f:
            for row in f:
                if len(row) < 5:
                    continue
                else:
                 rows.append(row)   
        names = [i for i in rows if "Name" in i] 
        balances = [i for i in rows if "Days" in i] 
        
        #balances = [[j,next(rows2),next(rows2)] for j in rows2 if "Balance Ending" in j]
        #balances = [[j,next(rows2),next(rows2),next(rows2)] for j in rows2 if "Ending" in j]
        rows3 = rows.copy()
        rows3.reverse()
        rows4 = iter(rows3)
        these_dates = [next(rows4) for i in rows4 if "Ending" in i]
        these_dates.reverse()
        dates = these_dates.copy()
        #dates = [i for i in rows if "/20" in i or "/19" in i] 
        #dates = [i for i in dates if "Hire" not in i]
        #dates = [i for i in dates if "Series" not in i]
        #dates = [i for i in dates if "Balance" in i]
        #dates = [i for i in dates if "Beginning" not in i]
        dates =liststrip(dates)   
        #dates = [i[0] for i in dates if i]
        #balances =[i for i in balances if i]    
        for i,a in enumerate(names):
            names[i]= re.sub(r'Dept:\s\d+\D+', '',re.sub(r' Name: ','',(re.sub(' +', ' ',a))))
        balances = liststrip(balances)
        balancescur1 = balances[4::6]
        balancescur2 = balances[5::6]
        final = []
        for i,x in enumerate(names):
            #print(names[i])
            #print(dates[i])
            #print(balancescur1[i])
            #print(balancescur2[i])
            final.append((names[i],dates[i],balancescur1[i],balancescur2[i]))
        #created a function for this 
        #dates = [i.split(',') for i in dates if i] 
        #for n,x in enumerate(dates):
        #    dates[n] =[i for i in x if i]
        xdf = pd.DataFrame(final)
        #df = pd.DataFrame({'name':names,'Ending':balances})  
        new = xdf[0].str.split(" ", n = 1, expand = True) 
        df=pd.DataFrame()
        df["Last_Name"] = new[0]
        df["First_Name"] = new[1]
        #new = df[1].str.split(",| ", n = 40, expand = True) 
        df["Ending"] = xdf[1].apply(listextract,args=(0,))
        #new2 = new.replace('', np.nan)
        #new2.dropna(axis=1,how='all',inplace=True)
        #df["Ending"] = pd.DataFrame(dates)
        #new2=df[2].apply(listextract,0)
        list1=["Annual_Days","Annual_hr","Annual_min","Sick_Days","Sick_hr","Sick_min"]
        df = dfgen(df,xdf[2],list1,6)
        list2 = ["Hol_days","Hol_hr","Hol_min","Comp_days","Comp_hr","Comp_min"]
        df = dfgen(df,xdf[3],list2,6)
        """
        This portion has been deprecated in favor of the above 4 line function call
        df["Annual_Days"] = xdf[2].apply(listextract,args=(0,))
        df["Annual_hr"] = xdf[2].apply(listextract,args=(1,))
        df["Annual_min"] = xdf[2].apply(listextract,args=(2,))
        df["sick_Days"] = xdf[2].apply(listextract,args=(3,))
        df["Sick_hr"] = xdf[2].apply(listextract,args=(4,))
        df["Sick_min"] = xdf[2].apply(listextract,args=(5,))
        df["Hol_days"] = xdf[3].apply(listextract,args=(0,))
        df["Hol_hr"] = xdf[3].apply(listextract,args=(1,))
        df["Hol_min"] = xdf[3].apply(listextract,args=(2,))
        df["Comp_days"] = xdf[3].apply(listextract,args=(3,))
        df["Comp_hr"] = xdf[3].apply(listextract,args=(4,))
        df["Comp_min"] = xdf[3].apply(listextract,args=(5,))
        """
        
        
        path = 'Y:\\'    
        for root, dirs, files in os.walk(path): 
            for file in files:  
          
                # change the extension from '.mp3' to  
                # the one of your choice. 
                if file.startswith('TLSummary'): 
                    filepath = f'{path}{file}'
                    os.remove(filepath) 
                    
        
        curtitle = f"Y:\TLSummary_{datetime.now().month}_{datetime.now().day}_{datetime.now().year}.xlsx"
        df.to_excel(curtitle)
        
        
        os.startfile(curtitle)
        time.sleep(1)
        pyautogui.press(['alt','f'])
        time.sleep(1)
        pyautogui.press('i')
        time.sleep(1)
        pyautogui.press('e')
        time.sleep(1)
        pyautogui.press(['alt','f','i','p','e'])
        time.sleep(1)
        pyautogui.typewrite(pass2)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.typewrite(pass2)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.hotkey('ctrl','s')
        time.sleep(1)
        pyautogui.hotkey('alt','F4')
        successmail()
        print("S'all good, man")
    else:
        print("You done motherfucking goofed")
        failmail()
        
        
if __name__ == "__main__":
    main(un,pass1,pass2)
