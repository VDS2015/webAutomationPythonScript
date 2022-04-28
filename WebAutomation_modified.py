import os
import time
import base64
import wget
from bs4 import BeautifulSoup
import psycopg2
import sys
import csv
import smtplib
import datetime	
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from datetime import datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from contextlib import contextmanager

try:
#get the path of ChromeDriverServer
	dir = os.path.dirname(__file__)
	folder_path="C:\\temp"
	chrome_options = Options()
	#chrome_options.add_argument("--headless")
	chrome_options.add_argument("--disable-gpu")
	chrome_options.add_argument("--window-size=1280,1696")
	chrome_driver_path = dir + "\chromedriver.exe"
	chrome_options.add_experimental_option("prefs",{
												"download.default_directory": r"C:\temp",
												"download.prompt_for_download": False,
												"download.directory_upgrade": True,
												"safebrowsing.enabled": True
												
											})
	#create a new Chrome session
	driver = webdriver.Chrome(executable_path=os.path.abspath("chromedriver"),chrome_options=chrome_options)
	driver.delete_all_cookies();
	driver.implicitly_wait(300)
	driver.maximize_window()
	#navigate to the application home page
	driver.get("http://10.100.202.37:8080/console")
	#get the search textbox
	search_field = driver.find_element_by_id("txtUserName")
	#enter search keyword and submit
	search_field.send_keys("administrator@google.com")
	search_field= driver.find_element_by_id("btnLogin")
	search_field.click()
	#wait for next page element id 
	timeout = 10000
	Password= "12345678"
	
	try:
		WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'password')))
	except Exception as e:
		print("Timed out waiting for page to load1")
	search_field = driver.find_element_by_id("password")
	search_field.send_keys(Password)
	search_field=driver.find_element_by_id("btnLogin")
	search_field.click()
	my_list=["Event","View", "Outstanding","xyz View", "xyz Summary","Service Level","ptr Summary"]
	if str(sys.argv[5]) in my_list:
		search_field = driver.find_element_by_id("xys")
		search_field.click()
		if(str(sys.argv[5]) == 'Event Summary'):
					search_field=driver.find_element_by_id("xya")
					search_field.click()
					print("Event Summary")
					search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlTimeSpan"))
					search_field.select_by_value(str(sys.argv[2]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdf")
					search_field.click()
		elif(str(sys.argv[5]) == 'Executive View'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("Executive View")
					search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlTimeSpan"))
					search_field.select_by_value(str(sys.argv[2]))
					search_field=Select(driver.find_element_by_id("ddlCompany"))
					search_field.select_by_value(str(sys.argv[1]))
					search_field=Select(driver.find_element_by_id("ddlAssignedGrp"))
					search_field.select_by_value(str(sys.argv[3]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdfev")
					search_field.click()
		elif(str(sys.argv[5]) == 'Outstanding'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("Outstanding")
					#search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlCompany"))
					search_field.select_by_value(str(sys.argv[1]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdfos")
					search_field.click()
		elif(str(sys.argv[5]) == 'IMC View'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("IMC View")
					#search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlCompany"))
					search_field.select_by_value(str(sys.argv[1]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdfid")
					search_field.click()
		elif(str(sys.argv[5]) == 'Activity Summary'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("Activity Summary")
					#search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlCompany"))
					search_field.select_by_value(str(sys.argv[1]))
					search_field=Select(driver.find_element_by_id("ddlTimeSpan"))
					search_field.select_by_value(str(sys.argv[2]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdfas")
					search_field.click()
		elif(str(sys.argv[5]) == 'Service Level'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("Service Level")
					#search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlTimeSpan"))
					search_field.select_by_value(str(sys.argv[2]))
					search_field=Select(driver.find_element_by_id("ddlCompany"))
					search_field.select_by_value(str(sys.argv[1]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdfsl")
					search_field.click()
		elif(str(sys.argv[5]) == 'IAP Summary'):
					search_field=driver.find_element_by_id("xyz")
					search_field.click()
					print("IAP Summary")
					#search_field=driver.find_element_by_id("ddlTimeSpan")
					search_field=Select(driver.find_element_by_id("ddlTimeSpan"))
					search_field.select_by_value(str(sys.argv[2]))
					time.sleep(10)
					search_field= driver.find_element_by_id("btnExportToPdf")
					search_field.click()
		else:
			print("ok")
	else:
		print("Please write argument like xyz Summary Executive, View xyz,xyz View,  xyz  Summary, Service Level, xyz Summary")
	print("downloaded file to desired location")
	time.sleep(10)
	driver.close()
	#close the browser window
except Exception as e:
	print(str(e))
	
def infomail(subject,message,attachment=''):
	try:
		fromaddr = "vidhayasagar@google.ca"
		toaddr = "vidhayasagar.h@gmail.com"
		#smtpserver = "xyz.groupinfra.com"
		smtpserver = "mail.google.com"
		smtpserverport = "25"
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Cc'] = ""
		msg['Subject'] = subject
		body = str(message) 
		msg.attach(MIMEText(body, 'html'))
		
		if (attachment!='') :
			try :
					reportfilename = [str(sys.argv[5]) +".pdf"]
					for f in reportfilename:
						reportfilepath = folder_path + "\\" + f
						part = MIMEApplication(open(reportfilepath,"rb").read())
						part.add_header('Content-Disposition', 'attachment', filename=f)
						msg.attach(part)
			except:
				pass
			
		server = smtplib.SMTP(smtpserver,smtpserverport)
		server.ehlo()
		server.starttls()
		server.ehlo()
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)	
		server.close()
		print("Send_Mail")
	except Exception as e:
		print(str(e))	
# --------------------------------------------------------------------------------------------------------------		




#------------------------- Sending Mail ------
date = str(datetime.datetime.today());
subject =  str(sys.argv[4])
body =  """\
<p>Dear User,</p>

<p>Please find the attached report</p>

<p>Note: This is an auto generated e-mail, kindly do not reply.</p>

<p><br />
Best Regards,<br />
xyz Communication<br />
<br />
"""
attachment = "yes"
infomail(subject,str(body),attachment);