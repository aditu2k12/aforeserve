import pyautogui
import pywinauto
import time
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import logging

logging.basicConfig()
flag=0
def mailConfig():
    
    global flag
    
    #open control panel
   
    control_panel=pywinauto.Application(backend='uia').start(r'C:\Windows\System32\control.exe',timeout=50)
    
    time.sleep(2)
    
    pyautogui.hotkey('win','up')

    time.sleep(5)

#click mail icon on control panel
    control_panel_window=pywinauto.findwindows.find_windows(best_match=u'All Control Panel Items')
    if control_panel_window:
         mail=control_panel.window_(handle=control_panel_window[0])
         try:
             mail.child_window(title="Mail (Microsoft Outlook 2016) (32-bit)", auto_id="name", control_type="Hyperlink").wait('visible', timeout=30, retry_interval=0.5).click_input()
         except Exception as e:
             mail.child_window(title="Mail", auto_id="name", control_type="Hyperlink").wait('visible', timeout=30, retry_interval=0.5).click_input()
         
         time.sleep(2)
         #click on sEmail Account
         mail_setup_outlook=None
         try:
             mail_setup_outlook=pywinauto.findwindows.find_windows(title='Mail Setup - Outlook')
         except:
             for i in range(3):
                 if mail_setup_outlook is None:
                     mail_setup_outlook=pywinauto.findwindows.find_windows(title='Mail Setup - Outlook')
                     time.sleep(2)
                 else:
                     break
         
         if mail_setup_outlook:
                mail_setup_outlook=control_panel.window_(handle=mail_setup_outlook[0])
                mail_setup_outlook.child_window(title="Email Accounts...").wait('visible', timeout=30, retry_interval=0.5).click()
                time.sleep(5)
                # Clicking new for new account
                #time.sleep(5)
         #click on Email Accounts
         #pyautogui.press("enter")
         time.sleep(5)
         pyautogui.press("tab")
         pyautogui.press("enter")
        
        # Clicking for manual setup
         time.sleep(2)
         pyautogui.press("down")
         pyautogui.press("enter")
        
        # clicking for POP
         time.sleep(2)
         pyautogui.press("down")
         pyautogui.press("enter")
         user_details=pd.read_csv('http://127.0.0.1:8000/input.csv').columns
        
        # Adding account details
         time.sleep(2)
         pyautogui.typewrite(user_details[1])
         pyautogui.press("tab")
         #pyautogui.typewrite('aditya.singh@algo8.ai')
         pyautogui.typewrite(user_details[2])
         pyautogui.press("tab",presses=2)
        #pyautogui.press("down")
        #pyautogui.press("tab")
         pyautogui.typewrite('outlook.office365.com')
         pyautogui.press("tab")
         pyautogui.typewrite('smtp.office365.com')
         pyautogui.press("tab")
         pyautogui.typewrite(user_details[1])
         pyautogui.press("tab")
         pyautogui.typewrite(user_details[3])
        
        
        # MY Code
         pyautogui.press("tab",presses=6)
         pyautogui.press("enter")
         pyautogui.hotkey('ctrl', 'shift','tab',presses=2)
         pyautogui.typewrite('995')
         pyautogui.press("tab",presses=2)
         pyautogui.press("space")
         pyautogui.press("tab")
         pyautogui.typewrite('587')
         pyautogui.press("tab")
         pyautogui.press("down",presses=3)
         pyautogui.press("tab",presses=4)
         pyautogui.typewrite('7')
         pyautogui.press("tab")
         pyautogui.press("space")
         pyautogui.press("tab")
         pyautogui.press("enter")
         pyautogui.press("enter")
         pyautogui.press("enter")
         time.sleep(4)
         pyautogui.press("enter")
    flag=1
    return flag

#mailConfig()
