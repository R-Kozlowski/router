import os, bs4, time, xlrd, datetime
import threading #threading processes
import pandas as pd
import pyautogui as gui #screen handling

from selenium import webdriver
from selenium.webdriver import ActionChains

#open the browser with hiding module
from selenium.webdriver.firefox.options import Options

#waiting modules
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException

#option for keyboard keys
from selenium.webdriver.common.keys import Keys

#info window handling
from tkinter import *
from tkinter import ttk
import tkinter as tk



#----------------------------------#
#ROBOT-USER communication

#generating simple user window
def okno_informacyjne(info):
    try:
        root = Tk()
        frm = ttk.Frame(root, padding=10)
        frm.grid()
        ttk.Label(frm, text='SPOT changed in:\n\n'+str(info)).grid(column=0, row=0)
        ttk.Button(frm, text="OK", command=root.destroy).grid(column=0, row=1)
        root.mainloop()
    except:
        root = Tk()
        frm = ttk.Frame(root, padding=10)
        frm.grid()
        ttk.Label(frm, text='Unable to display information window').grid(column=0, row=0)
        ttk.Button(frm, text="OK", command=root.destroy).grid(column=0, row=1)
        root.mainloop()



#----------------------------------#
#DOWNLOAD DATA

#I change the path to the one where the .py program is located
os.chdir(os.getcwd())

#downloading URL address from the excel file on the same folder direction
url = pd.read_excel("START.xlsm", sheet_name="url")
url = url.loc[0][0]
bu = url[:2]
url = url[5:]


#if url is empty then error window and exits browser
if url == '':
    okno_informacyjne('No URL address was given')
    browser.quit()

#take input criteria
net_config = pd.read_excel("START.xlsm", sheet_name="net_config")
net_config = net_config.fillna(value='') #delete NaN
net_config['To hub'] = net_config['To hub'].fillna(' ')
net_config['By hub'] = net_config['By hub'].fillna(' ')


#split pandas table to 2
split = net_config.index[len(net_config)//2]

while True:
    if net_config.iloc[split]['ID'] == net_config.iloc[split+1]['ID']:
        split = split - 1
    else:
        break

net_config1 = net_config.loc[:split].copy()
net_config2 = net_config.loc[split+1:].copy()
net_config2 = net_config2.reset_index(drop=True)

#downloading paths to the fileds in TMS
xpath = pd.read_excel("START.xlsm", sheet_name="xpath")

#creating empty table for existing rows
existing_rows = pd.DataFrame(columns=['Network code',
                                      'From hub',
                                      'To hub',
                                      'Priority',
                                      'By hub',
                                      'Description',
                                      'Restriction',
                                      'Number',
                                      'Table',
                                      'Column',
                                      'Oper',
                                      'Table 2',
                                      'Column 2',
                                      'Condition value'])



#----------------------------------#
#LOGGING

def pobieranie_logowania():
    global login, password
    login = e1.get()
    password = e2.get()
    master.destroy()

#building a window
master = Tk()
master.title("Logging")
master.geometry('400x100')
ttk.Label(master, text="Login").grid(row=0, column=0)
ttk.Label(master, text="Password").grid(row=1, column=0)
e1 = Entry(master)
e1.grid(row=0, column=1,sticky=W)
e2 = Entry(master)
e2.grid(row=1, column=1,sticky=W)
button_przekaz = Button(master,text='Log in',command=pobieranie_logowania)
button_przekaz.grid(row=2, column=1,sticky=W)
master.mainloop()


#open 2 browsers
browser1 = webdriver.Firefox()
browser2 = webdriver.Firefox()

browser1.maximize_window() #maximize window 1
browser2.maximize_window() #maximize window 2

#I pass on the size of the first window
width = browser1.execute_script("return screen.width")
height = browser1.execute_script("return screen.height")


#adjust first window
browser1.set_window_size(width, height//2)
browser1.set_window_position(0, height//2)

#adjust second window
browser2.set_window_position(0, 0)
browser2.set_window_size(width, height//2)

#open TMS system on both browsers
browser1.get(url)
browser2.get(url)



#----------------------------------------------#
#MAIN PROCESS

#function for each process
def dzielenie_procesow(proces,browser,login,password,net_config,bu,xpath,existing_rows,id):

    #waiting for logging site
    WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.ID, "username")))


    proc = 'logging'
    #logging
    browser.find_element(By.XPATH, 'example_xpath').send_keys(str(login))
    browser.find_element(By.XPATH, 'example_xpath').send_keys(str(password))
    browser.find_element(By.XPATH, 'example_xpath').submit()
    time.sleep(3)
    try:
        browser.find_element(By.XPATH, 'example_xpath').send_keys(str(login))
        browser.find_element(By.XPATH, 'example_xpath').send_keys(str(password))
        browser.find_element(By.XPATH, 'example_xpath').click()
    except:
        next


    time.sleep(2)

    proc = 'raport searching'
    #waiting for role appearance
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath')))
    WebDriverWait(browser, 50).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath')))
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('my_role', Keys.ENTER)
    time.sleep(2)

    #go to the report
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath')))
    
    #waiting little bit longer on the specific TMS
    if bu == 'PL':
        time.sleep(2)
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

    time.sleep(3)

    try:
        #loop for each row in pandas table
        while True:

            proc = 'save txt file'
            try:
                print('INPUT ROW(process '+str(proces)+'):\n------------------\n'+str(net_config.iloc[id])+'\n')
                with open('mr_robot_logs_proc_'+str(proces)+'.txt', 'a') as f:
                    f.write('INPUT ROW:\n------------------\n'+str(net_config.iloc[id])+'\n')
            except:
                break

            proc = 'check filters'
            
            #checking filters if given row exists there
            try:
                WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['Network code']), Keys.TAB)
            except:
                WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['Network code']), Keys.TAB)
        
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['From hub']), Keys.TAB)
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['To hub']), Keys.TAB)
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['Priority']), Keys.TAB)
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(str(net_config.loc[id]['By hub']), Keys.TAB)
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

            #checing if field is empty
            while True:
                try:
                    try:
                        czy_jest_hub_to = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).text
                        break
                    except:
                        time.sleep(1)
                        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                        czy_jest_hub_to = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).text
                        break
                except:
                    continue

            #if row doesn't exist robot will insert data
            if czy_jest_hub_to == ' ':

                #click on the first row
                try:
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                except:
                    time.sleep(2)
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    
                #click insert
                try:
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                except:
                    time.sleep(2)
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

                try:
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Network code'])))).send_keys(str(net_config.loc[id]['Network code']), Keys.TAB)
                except:
                    time.sleep(2)
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Network code'])))).send_keys(str(net_config.loc[id]['Network code']), Keys.TAB)

                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['From hub'])))).send_keys(str(net_config.loc[id]['From hub']), Keys.TAB)
                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['To hub'])))).send_keys(str(net_config.loc[id]['To hub']), Keys.TAB)
                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Priority'])))).send_keys(str(net_config.loc[id]['Priority']), Keys.TAB)
                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['By hub'])))).send_keys(str(net_config.loc[id]['By hub']), Keys.TAB)
                #if nobody gave description of new row I will insert 'MrRobot' for better overview who did this
                if net_config.loc[id]['Description'] != '':
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Description'])))).send_keys(str(net_config.loc[id]['Description']), Keys.TAB)
                else:
                    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Description'])))).send_keys('MrRobot', Keys.TAB)
                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Restriction'])))).send_keys(str(net_config.loc[id]['Restriction']), Keys.ENTER)

                time.sleep(1)

                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

                if net_config.loc[id]['Table'] != '':
                    x = 3
                    proc = 'menu'
                    try:
                        #klikam menu
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    except:
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    proc = 'conditions'
                    try:
                        #exception for different TMS
                        if bu == 'PL':
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                        else:
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    except:
                        try:
                            #if nothing happend I click again
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            time.sleep(2)
                            #exception for different TMS
                            if bu == 'PL':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            else:
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()          
                        except:
                            #exception for different TMS
                            if bu == 'PL':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            else:
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

                    #insert rows in conditions window
                    while True:
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Number'].replace('tr[3]', f'tr[{x}]'))))).send_keys(Keys.TAB)
                        #close first window
                        if bu == 'PL':
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                        else:
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Table'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Table']),Keys.TAB)
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Column'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Column']),Keys.TAB)

                        #exception for specific TMS
                        if bu == 'BS':
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                            time.sleep(1)
                            #depends from the operator I will choose specific line
                            if net_config.loc[id]['Oper'] == '=':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('equal to selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '<':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('less than the selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '<=':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('less than or equal to the selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '>':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('greater than the selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '>=':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('greater than or equal to the sel. value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '<>':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('not equal to selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '~':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('like the selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '!~':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('not like the selected value',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '^':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('starts with',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '!^':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys("doesn't start with",Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '$':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys('ends with',Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                            elif net_config.loc[id]['Oper'] == '!$':
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys("doesn't end with",Keys.TAB)
                                WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).send_keys(Keys.ENTER)
                        else:
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Oper'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Oper']),Keys.TAB)

                        #close second window
                        if bu == 'PL':
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                        else:
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()

                        #specific xpath taken from the excel file with replaced values
                        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Table 2'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Table 2']),Keys.TAB)
                        
                        if net_config.loc[id]['Table 2'] == '':
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Condition value'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Condition value']),Keys.TAB)
                        else:
                            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, str(xpath.loc[0]['Column 2'].replace('tr[3]', f'tr[{x}]'))))).send_keys(str(net_config.loc[id]['Column 2']),Keys.TAB)                    

                        x = x + 1

                        #if row exists I will go to the next row creation
                        try:
                            if net_config.loc[id+1]['ID'] == net_config.loc[id]['ID']:
                                id = id + 1
                            else:
                                break
                        except:
                            break

                    proc = 'close conditions window'
                    #close main window
                    if bu == 'PL':
                        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    else:
                        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
                    time.sleep(2)
                    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, 'example_xpath'))).click()
            else:
                print('Given ID exists - ' + str(net_config.loc[id]['ID']) + '\n')
                #and I move this row to separate dataframe
                existing_rows = pd.concat([existing_rows, (net_config.head(id+1)).tail(1)])
            id = id + 1

    except:
        print('robot stop working during ID - '+str(net_config.iloc[id]['ID']))
        print('in the process '+str(proces))
        #save everything which robot wasn't able to insert to crash table
        net_config.iloc[id:].to_excel(f"Crash table (proc "+str(proces)+").xlsx")
        print('process '+str(proces)+' was stopped during:\n===================\n'+str(proc)+'\n=================\n')

        now = datetime.datetime.now()
        date_time = now.strftime("%d-%m-%Y %H-%M-%S")
        existing_rows.to_excel(f"Existing rows (proc {proces}) - {date_time}.xlsx")
        return

    now = datetime.datetime.now()
    date_time = now.strftime("%d-%m-%Y %H-%M-%S")
    existing_rows.to_excel(f"Existing rows (proc {proces}) - {date_time}.xlsx")

    browser.quit()


#threading of processes
id = 0
thread1 = threading.Thread(target=dzielenie_procesow, args=(1,browser1,login,password,net_config1,bu,xpath,existing_rows,id))
idn = 0
thread2 = threading.Thread(target=dzielenie_procesow, args=(2,browser2,login,password,net_config2,bu,xpath,existing_rows,idn))

#start threads
thread1.start()
thread2.start()

#waiting for the end of each precess
thread1.join()
thread2.join()
