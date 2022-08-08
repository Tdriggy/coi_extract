import pdfquery
import clipboard
import time
from selenium import webdriver
# from selenium.webdriver.firefox.options import Options
# from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys


coordinate_dict = {'gl_policy_number': (2.95, 5.58, 4.58, 6.73),
                   'al_policy_number': (2.95, 4.74, 4.58, 5.57),
                   'gl_eff_date':      (4.61, 5.59, 5.22, 6.70),
                   'gl_exp_date':      (5.25, 5.58, 5.80, 6.70),
                   'gl_each_occur':    (7.30, 6.60, 8.23, 6.72)
                   }

file_name = 'COI - JWB Construction.pdf'


# This function creates the webdriver object for use with all subsequent functions

def chromedriver_setup():
    my_options = webdriver.ChromeOptions()
    # my_options = webdriver.FirefoxOptions()
    my_options.add_argument('user-data-dir=C:/Users/Tyler/ProgramingProjects/coi_extract/User_Data')
    # driver = webdriver.Firefox(profile)
    driver_setup = webdriver.Chrome(options=my_options)
    return driver_setup


# This function opens Riskonnect, makes the browser full-screen, and logs in using the password stored in passkey.txt

def logon_vendor_add():
    driver.get('https://riskonnectvmc.lightning.force.com/lightning/o/Vendor__c/new?count=1&nooverride=1'
               '&useRecordTypeCheck=1&navigationLocation=LIST_VIEW&uid=165956621188761667&backgroundContext'
               '=%2Flightning%2Fo%2FVendor__c%2Flist%3FfilterName%3DRecent&recordTypeId=012f1000000n7uAAAQ')
    driver.maximize_window()
    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'Login')))
    password = driver.find_element(By.XPATH, '//*[@id="password"]')
    f = open('passkey.txt', 'r')
    password.send_keys(f.read())
    time.sleep(1)
    submit = driver.find_element(By.ID, "Login")
    submit.click()
    f.close()
    #WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, '//*[@id="input-181"]')))
    #WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, '//*[@id="input-199"]')))


# This function uses the PDFquery module to extract a string from a PDF within the coordinates provided as an argument

def pdf_extract(x1, y1, x2, y2, pdf_file=file_name):
    coordinates = (x1 * 72, y1 * 72, x2 * 72, y2 * 72)
    pdf = pdfquery.PDFQuery(pdf_file)
    pdf.load(0)
    extracted_text = pdf.pq('LTTextLineHorizontal:overlaps_bbox("%d, %d, %d, %d")' % coordinates).text()
    # print(extracted_text)
    return extracted_text


def get_gl_policy_number(coordinates):
    gl_number = pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
    if gl_number[0:4] in ['X X ', 'x x ']:
        gl_number = gl_number[4:]
    print('GL Number: ' + gl_number)


def get_gl_eff_date(coordinates):
    gl_eff_date = pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
    print('GL Eff Date: ' + gl_eff_date)


def get_gl_exp_date(coordinates):
    gl_exp_date = pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
    print('GL Exp Date: ' + gl_exp_date)


def get_gl_each_occurrence(coordinates):
    gl_each_occurrence = pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
    print('GL EO Amount: ' + gl_each_occurrence)


def get_al_policy_number(coordinates):
    al_number = pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
    if al_number[0:4] in ['X X ', 'x x ']:
        al_number = al_number[4:]
    print('AL Number: ' + al_number)


if __name__ == '__main__':

    # driver = chromedriver_setup()
    # logon_vendor_add()
    get_gl_policy_number(coordinate_dict['gl_policy_number'])
    get_gl_eff_date(coordinate_dict['gl_eff_date'])
    get_gl_exp_date(coordinate_dict['gl_exp_date'])
    get_gl_each_occurrence(coordinate_dict['gl_each_occur'])
    get_al_policy_number(coordinate_dict['al_policy_number'])


