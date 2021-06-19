from flask import Flask,request,render_template
import ITSM_Fetching_Tickets
from selenium.webdriver.common.keys import Keys
from csv import writer
import Printer_Latest
import time
import smtplib
import test_fetch_mail
import diskCleanup
import pyautogui
from selenium import webdriver
import updateTicketStatus
from datetime import datetime, timedelta

app = Flask(__name__)

def sendMail(ticket_id,issue,email):
    sender_email = "bramhesh.srivastav@algo8.ai"
    rec_email = 'aditya.singh@algo8.ai'
    password = "Dev@1234@"
    now = datetime.now()
    if email:
        message = "\nlink -> http://127.0.0.1:9004/bot/"+issue+"/"+ticket_id+"/"+str(now)+"/12345698638463986398563986398753298137190583=True"+"\nWe have a password reset request from "+email\
                   +"\n\n Please click on this link only if you had raised the password reset request"
        
    else:
        message = "\nlink -> http://127.0.0.1:9004/bot/"+issue+"/"+ticket_id+"/12345698638463986398563986398753298137190583=True"
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    server.starttls()
    server.login(sender_email, password)
    print("Login success")
    server.sendmail(sender_email, rec_email, message)
    print("Email has been sent to ", rec_email)

sendMail('1234','shusu','a')

def append_list_as_row(file_name, list_of_elem):
    with open(file_name, 'w+', newline='') as write_obj:
        # Create a writer object from csv module
      csv_writer = writer(write_obj)
        # Add contents of list as last row in the csv file
      csv_writer.writerow(list_of_elem)
      
@app.route('/updateTickets')      
def updateTicketsStatus():
     
    driver = webdriver.Chrome(r'C:\Users\Aditya\Aforeserve\Demo\chromedriver.exe')
    # Gets the URL 
    driver.get('https://staging-ci.symphonysummit.com/Afsservicedesk/Summit_WebLogin.aspx')
    
    driver.maximize_window()
    
    time.sleep(3)
    
    # Finds the user name and password by id from HTML
    element = driver.find_element_by_id("txtLogin")
    element1 = driver.find_element_by_id("txtPassword")
    
    # Enters the user name and password
    element.send_keys("afsauto@aforeserve.co.in")
    element1.send_keys("Afs@123#")
    
    # Clicks the login button
    driver.find_element_by_id("butSubmit").click()
    
    
    # This is used to check if a duplicate login window pop ups, if it does press continue otherwise pass 
    try:
        if driver.find_element_by_class_name("TitlePanel").text == 'DUPLICATE LOGIN':
            driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
            driver.find_element_by_id("ContentPanel_btnContinue").click()
        else:
            pass
    except:
        print(" ")
    
    time.sleep(1)
    driver.find_element_by_id("IM").click()
    driver.find_element_by_id("IM_WORKGROUP_TICKETS").click()
    
    num_of_tickets = int(driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_lblCurrentRange"]').text.split()[-1].strip())
    
    for i in range(0,len(df)):
        in_id = int(df['Incident ID'][i])
        solu = df['Solution'][i]
        stat = df['Status'][i]
        
        for j in range(1,num_of_tickets+1):
            incident=int(driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_gvMyTickets"]/tbody/tr['+str(j+1)+']/td[2]/div[2]/a[1]').text)
            if incident == in_id:
                driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_gvMyTickets"]/tbody/tr['+str(j+1)+']/td[2]/div[2]/a[1]').click()
                time.sleep(2)
                
                # This clicks on the assigned
                driver.find_element_by_xpath('//*[@id="ticketdetail"]/div[2]/div/div[2]/div/div[1]/div/div/ul/li[2]/a').click()
                
                # This clicks on the assigned to option
                driver.find_element_by_xpath('//*[@id="s2id_BodyContentPlaceHolder_ddlAssignedExecutive"]/a/span[2]/b').click()
                #time.sleep(3)
                afs= driver.find_element_by_xpath('//*[@id="s2id_autogen34_search"]')
                afs.send_keys("AFS Automation",Keys.ENTER)
                #pyautogui.press("down")
                #pyautogui.press("enter")
                
                # This clicks on  the communication panel
                driver.find_element_by_xpath('//*[@id="aCommunication"]').click()
                
                # This fills the communication panel
                text = driver.find_element_by_xpath('//*[@id="Communication"]/div/div[2]/div[3]/div[2]/div')
                text.send_keys("Ticket is in progress")
                
                # Clicks on the resolved part 
                
                driver.find_element_by_xpath('//*[@id="ticketdetail"]/div[2]/div/div[2]/div/div[1]/div/div/ul/li[5]/a').click()
                driver.find_element_by_xpath('//*[@id="ticketdetail"]/div[2]/div/div[2]/div/div[1]/div/div/ul/li[5]/a').text
                
                time.sleep(10)
                # Clicks on the general panel
                driver.find_element_by_xpath('//*[@id="general"]/a').click()
               
           
                time.sleep(3)
                # This scrolls the window
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
               
                
                
                # Clicks on the solution panel and fills it
                driver.find_element_by_xpath('//*[@id="divSolutionRow"]/div[2]/div/div[2]/div').click()
                #time.sleep(2)
                solution=driver.find_element_by_xpath('//*[@id="divSolutionRow"]/div[2]/div/div[2]/div')
                print('Solution is ',solu)
                if solu:
                    solution.send_keys(solu)
                else:
                    solution.send_keys('Solution still Pending')
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                
                # This clicks on the resolution code and selects  
                #all_sol =['SELECT','Resolved','Completed With Errors','User not responding','Other Support Required','Out-of-Scope']
                
                if stat=='Resolved':
                      driver.find_element_by_xpath('//*[@id="s2id_BodyContentPlaceHolder_ddlResolutionCode"]/a/span[2]/b').click()
                      time.sleep(3)
                      
                      pyautogui.press("down")
                      pyautogui.press("enter")
                else:
                      driver.find_element_by_xpath('//*[@id="s2id_BodyContentPlaceHolder_ddlResolutionCode"]/a/span[2]/b').click()
                      time.sleep(3)
                      pyautogui.press("down",presses=4)
                      pyautogui.press("enter")
                
                
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                
                
                # This clicks on the violation panel , if its is yes then it fills the text
                
                if driver.find_element_by_xpath('//*[@id="General"]/div[2]/div/div[2]/div[2]/div/div[2]/label').text[-3:] == 'Yes':
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    driver.find_element_by_xpath('//*[@id="iRespReasonOpener"]').click()
                    text1 = driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_txtRespViolationReason"]')
                    text1.send_keys('This happened because of system latency')
                    driver.find_element_by_xpath('//*[@id="divRespViolationReason"]/div[3]/input').click()
                else:
                    pass
            
                #This goes one window back
                driver.execute_script("window.history.go(-1)")
                
                
            else:
                continue
                 
                
    driver.find_element_by_xpath('//*[@id="imgProfile"]').click()
    driver.find_element_by_xpath('//*[@id="hrefLogout"]').click()
    driver.quit()

@app.route('/bot/<issue>/<email>/<num>',methods=['GET','POST'])
def gdata(issue,email,num):
    global fdata
    fdata = list() 
    if issue=='Printer':
        return render_template('pr.html')
    elif issue=='Email':
        return render_template('index_E.html')
    elif issue=='Password':
        return render_template('p.html')
    elif issue=='DiskCLeanup':
        return render_template('disk.html')
    elif issue=='SoftwareInstall':
        return render_template('softwareInstall.html')
    else:
        return render_template('softwareInstall.html')

@app.route('/get')
def getdata():
    text = request.args.get('msg')
    fdata.append(text)
    append_list_as_row('input.csv', fdata)
    
@app.route('/Printer/<ticket_id>/<time>')
def printerConf(ticket_id,time):
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
        time.sleep(5)
        flag=Printer_Latest.printerConfig()
        if flag==1:
            df.loc[df['Incident ID']==int(ticket_id),'Status']='Resolved'
            df.loc[df['Incident ID']==int(ticket_id),'Solution']='Printer Configuration done'
        else:
            df.loc[df['Incident ID']==int(ticket_id),'Status']='Out-of-Scope'
            df.loc[df['Incident ID']==int(ticket_id),'Solution']='Printer Configuration not done'
        print(df)
    
    

@app.route('/Email/<ticket_id>/<time>')
def emailConf(ticket_id,time):
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
        time.sleep(5)
        flag=test_fetch_mail.mailConfig()
        if flag==1:
            df.loc[df['Incident ID']==int(ticket_id),'Status']='Resolved'
            df.loc[df['Incident ID']==int(ticket_id),'Solution']='Email Configuration done'
        else:
            df.loc[df['Incident ID']==int(ticket_id),'Status']='Out-of-Scope'
            df.loc[df['Incident ID']==int(ticket_id),'Solution']='Email Configuration not done'
        
        print(df)

@app.route('/Password/<ticket_id>/<time>')
def passwordConf(ticket_id,time):
    
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
        df.loc[df['Incident ID']==int(ticket_id),'Status']='Resolved'
        df.loc[df['Incident ID']==int(ticket_id),'Solution']='Password reset done'
        
        print(df)
    
@app.route('/DiskCleanup/<ticket_id>/<time>')
def diskCleaning(ticket_id,time):
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
        time.sleep(5)
        errors=diskCleanup.startCleanup()
        print('errors returned for diskcleanup is',errors)
        if errors==None:
            df.loc[df['Incident ID']==int(ticket_id),'Status']='Resolved'
            df.loc[df['Incident ID']==int(ticket_id),'Solution']='PC normal working restored'
            
        print(df)
    
@app.route('/SoftwareInstall/<ticket_id>/<time>')
def softwareInstall(ticket_id,time):
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
    
        df.loc[df['Incident ID']==int(ticket_id),'Status']='Other Support Required'
        df.loc[df['Incident ID']==int(ticket_id),'Solution']='Software installation not done'
        
        print(df)
    
@app.route('/OtherIssue/<ticket_id>/<time>')
def otherIssues(ticket_id,time):
    received_time = str(time)
    datetime_object = datetime.strptime(received_time, '%Y-%m-%d %H:%M:%S.%f')
    now_plus_30 = datetime_object + timedelta(minutes = 30)
    now = datetime.now()
    if now > now_plus_30:
        return render_template('404.html')
    else:
        df.loc[df['Incident ID']==int(ticket_id),'Status']='Other Support Required'
        df.loc[df['Incident ID']==int(ticket_id),'Solution']='Other Issue'
    
@app.route('/')
def predict():
    global df
    
    df,num_of_tickets=ITSM_Fetching_Tickets.loginAndFetchTickets()
    
    class_mapping={0:'DiskCLeanup_issue',1:'Email_issue',2:'Others_issue',3:'Password_issue',4:'Printer_issue',5:'SoftwareInstall_issue'}
    
    df['Issue_Class']=df['predicted_class_num'].map(class_mapping)
    
    df.to_excel('All_Incidents.xlsx',sheet_name='Incidents',index=False)
    
    df3=df.copy()
    df3.drop(['predicted_class_num'],axis=1,inplace=True)
    df['Status']=None
    df['Solution']=None
    
    print(df.dtypes)
    
    for inc_id,user_mail,pred_class_num in zip(df['Incident ID'].values,df['User_Mail'].values,df['predicted_class_num'].values):
         
        if pred_class_num==0:
            issue='DiskCLeanup'
            sendMail(inc_id,issue,customer_name=None,department=None,email=None)
             
        if pred_class_num==1:
            issue='Email'
            sendMail(inc_id,issue,customer_name,department,email=None)
            
        if pred_class_num==2:
            issue='otherIssue'
            sendMail(inc_id,issue,customer_name=None,department=None,email=None)
        
        if pred_class_num==3:
            issue='Password'
            sendMail(inc_id,issue,customer_name=None,department=None,user_mail)
                        
        if pred_class_num==4:
            issue='Printer'
            sendMail(inc_id,issue,customer_name=None,department=None,email=None)
            
        if pred_class_num==5:
            issue='SoftwareInstall'
            sendMail(inc_id,issue,customer_name=None,department=None,email=None)
            
        
    df['Incident ID']=df['Incident ID'].astype(int)
    
    updateTicketStatus.ticketUpdate(df)
    
        
    return render_template('index.html',tables=[df3.to_html(classes='data')])

if __name__ == '__main__':
    app.run(port=9004)
    
        