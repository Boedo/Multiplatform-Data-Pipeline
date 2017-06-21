import boto3
import imaplib
import email
import os, sys
from os import listdir
from os.path import isfile, join
import datetime

import selenium
import time
import shutil
from os import rename
import re
import glob
from selenium import webdriver
from selenium.webdriver.common.keys import Keys #
from selenium.webdriver import ActionChains


################################################################################
############################### GENERAL VARIABLES ###############################
################################################################################

# Where are all the files sent to - All platforms
svdir = 'C:\\Users\\nicol\\Desktop\\Activity Monitor\\dailyData'

# Defining date for the email fetching - Avocet/AdWords
date = (datetime.date.today()).strftime("%d-%b-%Y")

# Selenium reauirement - Facebook/Bing
driver = webdriver.Chrome('C:\\Users\\nicol\\Desktop\\Scripts\\chromedriver.exe')

# Centralised folder
dest_dir = "C:\\Users\\nicol\\Desktop\\Activity Monitor\\dailyData"

##########################################
# Centralising files - TO DO: ADD TO EVERY PLATFORM DOWNLOADS TO MATCH ON FILENAME
##########################################

def moving_file():
    for filee in glob.glob(r"C:\\Users\\nicol\\Downloads\\*.xls"):
        shutil.move(filee, dest_dir)
        stored_filee = dest_dir + "/" + filee
        print stored_filee

    for filee in glob.glob(r"C:\\Users\\nicol\\Downloads\\*.xlsx"):
        shutil.move(filee, dest_dir)
        stored_filee = dest_dir + "/" + filee
        print stored_filee

    time.sleep(2)

###############################################################################
################################ BING REPORTS #################################
###############################################################################

#########################################
# Get Bing Reports using Selenium
#########################################

def bingData():
    time.sleep(5)

    driver.get('https://secure.azure.bingads.microsoft.com/Auth')
    driver.maximize_window()

    time.sleep(3)

    login = driver.find_element_by_id("LoginModel_Username")
    login.send_keys("Bob@gmail.com")

    time.sleep(1)

    login_next = driver.find_element_by_id("LoginSectionNextButton")
    login_next.click()

    time.sleep(1)

    password = driver.find_element_by_xpath('//*[@id="i0118"]')
    password.send_keys("password")

    time.sleep(1)

    validation = driver.find_element_by_id("idSIButton9")
    validation.click()

    ##########################################
    # Delivery report
    ##########################################

    time.sleep(1)

    driver.get("https://ui.bingads.microsoft.com/Reporting/?cid=160309793&aid=141009974#createreport/getselectedscopefilters?_=1490040454724&cid=160309793")

    time.sleep(2)

    my_reports = driver.find_element_by_xpath('//*[@id="MyReports"]/a')
    my_reports.click()

    time.sleep(1)

    delivery_report = driver.find_element_by_id('Grid_1046956_cd')
    delivery_report.click()

    time.sleep(1)

    download_report = driver.find_element_by_id('Grid_runDownload')
    download_report.click()

    time.sleep(20)

    ##########################################
    # Conversion report
    ##########################################

    time.sleep(1)

    driver.get("https://ui.bingads.microsoft.com/Reporting/?cid=160309793&aid=141009974#createreport/getselectedscopefilters?_=1490040454724&cid=160309793")

    time.sleep(2)

    my_reports = driver.find_element_by_xpath('//*[@id="MyReports"]/a')
    my_reports.click()

    time.sleep(1)

    delivery_report = driver.find_element_by_id('Grid_1046957_cd')
    delivery_report.click()

    time.sleep(1)

    download_report = driver.find_element_by_id('Grid_runDownload')
    download_report.click()

    time.sleep(20)

bingData()
moving_file()
################################################################################
################################# FACEBOOK REPORT ##############################
################################################################################

##########################################
# Get Facebook Reports using Selenium
##########################################

def facebookData():
    time.sleep(5)

    driver.get('https://business.facebook.com/')
    driver.maximize_window()

    time.sleep(1)

    login = driver.find_element_by_id("email")
    login.send_keys("Bob@gmail.com")

    time.sleep(1)

    password = driver.find_element_by_id("pass")
    password.send_keys("password")

    time.sleep(1)

    validation = driver.find_element_by_id("u_0_2")

    if validation.is_displayed():
        validation.click()
        print "click 2"
    else:
        driver.find_element_by_id("u_0_1").click()
        print "click 1"

    time.sleep(1)

    driver.get("https://business.facebook.com/ads/manage/powereditor/reporting?act=1232797566750998")

    time.sleep(1)

    ##########################################
    # Delivery/Conversion report
    ##########################################

    delivery_report = driver.find_element_by_link_text("Facebook Delivery Report")
    delivery_report.click()

    time.sleep(1)

    export_report = driver.find_element_by_class_name('_3s9l')
    export_report.click()

    time.sleep(1)

    # Export from pop-up overlay

    export_finalise = driver.find_element_by_xpath('//*[@id="facebook"]/body/div[7]/div[2]/div/div/div/div/div[3]/div/div/div[2]/div/button')
    export_finalise.click()

    time.sleep(6)

facebookData()
moving_file()
################################################################################
################################# AVOCET, ADWORDS REPORTS ###############################
################################################################################

##########################################
# Connecting to the server, to the account, get the directory
##########################################
connection = imaplib.IMAP4_SSL('outlook.office365.com',993)
connection.login('Bob@gmail;com','password')
connection.list(directory='""', pattern='*')
connection.select("Inbox")

##########################################
# Find the target emails
##########################################

liste = ['(SUBJECT "AdWords Delivery Data" SENTSINCE {date})','(SUBJECT "Avocet Delivery Report" SENTSINCE {date})',
'(SUBJECT "Your scheduled report is ready to view" SENTSINCE {date})']

#########################################
## Actual code parsing emails
#########################################
i = 0
for query in range(len(liste)):

    liste_item = liste[i]
    print liste_item
    typ, data = connection.uid('search', None, liste_item.format(date=date))
    print typ, data
    target_email = data[0].split()

    for emailid in target_email:
        typ, data = connection.uid('fetch', emailid, '(RFC822)') # fetch the email body (RFC822) for the given ID
        raw_email = data[0][1]
        email_msg = email.message_from_string(raw_email)

        if email_msg.get_content_maintype() != 'multipart':
            continue

        for part in email_msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            print filename
            sv_path = os.path.join(svdir, filename)

            if not os.path.isfile(sv_path):
                fp = open(sv_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

            i += 1

# moving_file()
################################################################################
################################# UPLOADING TO Amazon S3 ##############################
################################################################################

##########################################
# Sending to S3 Bucket
##########################################

svdr = 'C:\\Users\\nicol\\Desktop\\Scripts\\S3_Control\\Files\\'

onlyfiles = [f for f in listdir('C:\\Users\\nicol\\Desktop\\Scripts\\S3_Control\\Files') if
isfile(join('C:\\Users\\nicol\\Desktop\\Scripts\\S3_Control\\Files',f))]

s3 = boto3.resource('s3')

target_bucket = 'kfc-bucket'

n = 0
for dossier in onlyfiles:
    file_to_bucket = 'C:\\Users\\nicol\\Desktop\\Scripts\\S3_Control\\Files\\' + onlyfiles[n]

    for file in file_to_bucket:
        try:
            response = s3.Object(target_bucket, onlyfiles[n]).put(Body=open(file_to_bucket,'r'))
        except Exception as error:
            print error
    n += 1
