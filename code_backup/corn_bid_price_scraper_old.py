import logging
import os
import sys
import time
from datetime import datetime
#import pytesseract
import glob
import requests
import xlwings as xw
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import bu_alerts

sender_email = 'biourjapowerdata@biourja.com'
sender_password = r'bY3mLSQ-\Q!9QmXJ'
receiver_email = 'manish.gupta@biourja.com,devina.ligga@biourja.com,Adarsh.Bhandari@biourja.com'
#path = 'Cornbids.xlsx'
path=r"S:\IT Dev\Production_Environment\cron-bid-price-automation"

logfile = os.getcwd() + '\\' + str(datetime.today().date()) + '_CornBidPrice_Logfile.txt'
logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')

#download_path = os.getcwd() + '\\download\\'
download_path = path + '\\download\\'

logger = logging.getLogger()
logger.setLevel(logging.INFO)

# browser user_agent
headers = {'User-Agent': 'Mozilla/5.0'}
DRIVER_PATH = r'S:\IT Dev\Production_Environment\cron-bid-price-automation\chromedriver.exe'
options = Options()
options.headless = True
driver = webdriver.Chrome(options=options, executable_path=DRIVER_PATH)

# browser user_agent
'''headers = {'User-Agent': 'Mozilla/5.0'}
DRIVER_PATH = r'S:\IT Dev\Production_Environment\cron-bid-price-automation\geckodriver.exe'
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.dir", download_path)
driver=webdriver.Firefox(executable_path=DRIVER_PATH, firefox_profile=profile)'''

# month list used to check name of months on websites
month_list = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
month_number_dic = {'jan': '01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06','jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'}

# this section is used to define future months
# that are to be set as new columns in the new sheet
base_date = datetime.today().date().replace(day=1)
future_months = list()
for i in range(0, 6):
    future_months.append(base_date + relativedelta(months=i))


# creates a new sheet in the excel file and copies the default columns
# over to the new sheet. also creates new columns in the sheet.
# if the new sheet is already present, new sheet is not created then
def initialize_new_sheet(bid_prices):
    try:
        latest_sheet = bid_prices.sheets.active
        # opening the latest tab to copy initial columns from it --
        new_sheet_name = str(datetime.now().month) + '.' + str(datetime.now().day)
        sheet_names = [i.name for i in bid_prices.sheets]
        if new_sheet_name not in sheet_names:
            # copying the common columns to the newly created tab --
            bid_prices.sheets.add(new_sheet_name, after=latest_sheet)
            latest_sheet = latest_sheet.range('A:G')
            latest_sheet.copy(bid_prices.sheets[new_sheet_name].range('A:G'))

            # adding the date columns --
            new_date_columns = ['H1', 'I1', 'J1', 'K1', 'L1', 'M1']
            i = 0
            for column in new_date_columns:
                xw.Range(column).value = future_months[i]
                i += 1
            bid_prices.save()
            bid_prices.close()
        else:
            return False
        return True
    except Exception:
        print(sys.exc_info()[0])
        logger.info(sys.exc_info()[0])
        print("error occoured in initalizing new sheet")
        logger.info("error occoured in initalizing new sheet")
        return False


# to scrape from a singular website.
# these return a dictionary of month date as keys and basis as values.
# returns an empty dict if it fails or excpetion occours
def scrape_absenergy():
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        time.sleep(10)
        page = requests.get("http://www.absenergy.org/grainbids.html")
        time.sleep(15)
        soup = BeautifulSoup(page.text, features='lxml')
        table_rows = soup.find_all('table')[6].find_all('tr')
        month_to_basis = dict()
        for i in table_rows[3:9]:
            month = i.find_all('td')[0].text.strip().lower()
            basis = i.find_all('td')[2].text.strip()
            if month:
                basis = float(basis[:-5])
                if 'first' in month or 'second' in month:
                    month = month[-3:]
                month = '01' + month[0:3] + str(year)
                month = datetime.strptime(month, '%d%b%Y').date()
                if month in month_to_basis:
                    current = month_to_basis[month]
                    if basis <= 0 and current <= 0:
                        month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                    else:
                        month_to_basis[month] = round((basis + current) / 2, 3)
                else:
                    month_to_basis[month] = basis
                if month.month == 12:
                    year += 1

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website: http://www.absenergy.org/grainbids.html")
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: http://www.absenergy.org/grainbids.html")
        return month_to_basis

# to scrape from a singular website
def scrape_midwestagenergy():
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        #driver.get("https://www.midwestagenergy.com/fccp-blue-flint-bids-19639")
        #driver.get("https://www.midwestagenergy.com/cashbidssingle-1703")
        time.sleep(20)
        #WebDriverWait(driver, 15).until(EC.frame_to_be_available_and_switch_to_it(driver.find_element_by_xpath( "/html/body/form/div[2]/div[2]/div[2]/div/div[2]/div[1]/div/p[15]/iframe")))
        #a23=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,"/html/body/form/div[2]/div[2]/div/div/div/div/div/div/div[1]/div")))
        res = requests.get('https://www.midwestagenergy.com/cashbidssingle-1703')
        time.sleep(15)
        soup = BeautifulSoup(res.content, features='lxml')
        table = soup.find_all('div', attrs={'class': 'cashBidLocation'})[0].find_all('ul')
        for row in table[1:]:
            month = row.find_all('li')[0].text.strip().lower()
            basis = float(row.find_all('li')[2].text.strip())
            if month:
                month = '01' + month + str(year)
                month = datetime.strptime(month, '%d%B%Y').date()
                month_to_basis[month] = basis
                if month.month == 12:
                    year += 1
        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website: https://www.midwestagenergy.com/fccp-blue-flint-bids-19639")
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: https://www.midwestagenergy.com/fccp-blue-flint-bids-19639")
        return month_to_basis


# to scrape from a singular website
def scrape_frvethanol():
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        driver.get("https://www.frvethanol.com/cashbids/")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'cashbids-data-table')))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('table', attrs={'id': 'cashbids-data-table'})[0].find_all('tr')
        for row in table[1:7]:
            month = row.find_all('td')[0].find('span').text.strip().lower()
            basis = row.find_all('td')[3].find('span').text.strip()
            if month:
                month = '01' + month + str(year)
                month = datetime.strptime(month, '%d%B%Y').date()
                month_to_basis[month] = float(basis)
                if month.month == 12:
                    year += 1
        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website: https://www.frvethanol.com/cashbids/")
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: https://www.frvethanol.com/cashbids/")
        return month_to_basis


# to scrape from a singular website, but different locations
def scrape_fhr(url):
    month_to_basis = dict()
    try:
        driver.get(url)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, 'pricingTable')))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        tables = soup.find_all('table', attrs={'class': 'priceYear'})
        for table in tables:
            year = table.find('tr').text.strip()
            rows = table.find_all('tr')
            for row in rows[1:]:
                month = row.find_all('td')[0].text.strip().lower()[0:3]
                basis = row.find_all('td')[2].text.strip()
                if month in month_list:
                    month = '01' + month + year
                    month = datetime.strptime(month, '%d%b%Y').date()
                    month_to_basis[month] = float(basis)

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:", url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website:", url)
        return month_to_basis



# calls the fhr_scrape function and then insert_to_sheet function
def fetch_and_insert_fhr():
    try:
        fhr_urls = {"https://www.fhr.com/corn-prices/arthur": 54, "https://www.fhr.com/corn-prices/fairbank": 56,
                    "https://www.fhr.com/corn-prices/Fairmont": 57, "https://www.fhr.com/corn-prices/iowa-falls": 58,
                    "https://www.fhr.com/corn-prices/Menlo": 59, "https://www.fhr.com/corn-prices/shell-rock": 60}

        for url in fhr_urls:
            bids = scrape_fhr(url)
            if insert_into_sheet(fhr_urls[url], bids):
                print("success for row " + str(fhr_urls[url]))
                logger.info("success for row " + str(fhr_urls[url]))
        return True
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured in fhr_urls")
        logger.info(sys.exc_info()[0])
        logger.info("error occoured in fhr_urls")


# for singular website but different locations, scrapes and inserts one by one
def scrape_and_insert_gpreinc():
    try:
        driver.get("http://www.gpreinc.com/corn-bids")
        city_select_values = {'Atkinson': 73, 'Central City': 74, 'Mount Vernon': 79, 'Obion': 80, 'Ord': 81,
                              'Shenandoah': 82, 'Superior': 83, 'York': 84}
        time.sleep(15)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "cash-bids-table-widget-1xfyw7x")))
        select = Select(
            driver.find_element_by_xpath("//*[@id=\"dtn_cashBids_container\"]/div/div[1]/form/div/div[1]"
        
                                         "/div/dtn-select[2]/label/select"))
                                         
        select.select_by_value("Corn")
        for city in city_select_values:
            month_to_basis = dict()
            select = Select(driver.find_element_by_xpath("//*[@id=\"dtn_cashBids_container\"]/div/div[1]/form/div/div[1]/div/dtn-select[1]/label/select"))
            select.select_by_value(city)
            soup = BeautifulSoup(driver.page_source, features='lxml')
            table = soup.find('table').find_all('tr')
            for row in table[1:7]:
                month = row.find_all('td')[1].text.strip().lower()
                basis = row.find_all('td')[3].text.strip()
                if month[:3] in month_list:
                    month = '01' + month[0:3] + '20' + month[-2:]
                    month = datetime.strptime(month, '%d%b%Y').date()
                    month_to_basis[month] = basis

            if insert_into_sheet(city_select_values[city], month_to_basis):
                print("success for row " + str(city_select_values[city]))
                logger.info("success for row " + str(city_select_values[city]))
            del month_to_basis
        return True
    except Exception:
        print(sys.exc_info()[0])
        logger.info(sys.exc_info()[0])
        print("error occoured in gpreinc_urls")
        logger.info("error occoured in gpreinc_urls")
        return False


# to scrape from a singular website
def poet_biorefining2(url, basis_index):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        months, basis_values = list(), list()
        driver.get(url)
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CLASS_NAME, "DataGrid")))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        basis_rows = soup.find_all('table', attrs={'class': 'DataGrid'})[0].find_all('tr')[basis_index].find_all('td')
        month_rows = soup.find_all('table', attrs={'class': 'DataGrid'})[0].find_all('tr')[0].find_all('th')
        for row in month_rows[1:]:
            month = row.find('span').text.strip().lower()
            if month[:3] in month_list:
                if month[-2:].isdigit():
                    month = '01' + month[0:3] + '20' + month[-2:]
                    month = datetime.strptime(month, '%d%b%Y').date()
                else:
                    month = '01' + month[0:3] + str(year)
                    month = datetime.strptime(month, '%d%b%Y').date()
                    if month.month == 12:
                        year += 1
                months.append(month)

        for row in basis_rows:
            basis = row.text.strip()
            basis_values.append(float(basis))

        index = 0
        for i in months:
            month_to_basis[i] = basis_values[index]
            index += 1

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:"+url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website:"+url)
        return month_to_basis


# to scrape from a singular website
def scrape_admfarm(url):
    month_to_basis = dict()
    try:
        driver.get(url)

        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "future-basis-cash")))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('table', attrs={'class': 'future-basis-cash'})[0].find_all('tr')
        for row in table[1:11]:
            month = row.find_all('td')[0].find('span').text.strip().lower()
            basis = float(row.find_all('td')[3].find('span').text.strip())
            index1 = month.find('/')
            index2 = month.find('/', index1 + 1, len(month))
            month = month[:index1+1] + '01' + month[index2:]
            month = datetime.strptime(month, '%m/%d/%y').date()
            if month in month_to_basis:
                current = month_to_basis[month]
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:"+url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: {}".format(url))
        return month_to_basis


# function to scrape from a type of webiste where month occours in the header
def scrape_regular_website_1(url, basis_index, iframe_xpath=""):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        driver.get(url)
        if iframe_xpath:
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it(driver.find_element_by_xpath(iframe_xpath)))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "DataGrid")))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('table', attrs={'class': 'DataGrid'})[0].find_all('tr')
        for row in table[2:]:
            month = row.find_all('th')[0].text.strip().lower()
            # for dates like - lh dec, fh dec
            if month[0:2] in ('lh', 'fh'):
                month = month[2:].strip()
            basis = row.find_all('td')[basis_index].text.strip()
            if basis and month[:3] in month_list:
                if month[-2:].isdigit():
                    month = '01' + month[0:3] + '20' + month[-2:]
                    month = datetime.strptime(month, '%d%b%Y').date()
                    month_to_basis[month] = float(basis)
                else:
                    month = '01' + month[0:3] + str(year)
                    month = datetime.strptime(month, '%d%b%Y').date()
                    month_to_basis[month] = float(basis)
                    if month.month == 12:
                        year += 1
                if month in month_to_basis:
                    current = month_to_basis[month]
                    basis = float(basis)
                    if float(basis) <= 0 and current <= 0:
                        month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                    else:
                        month_to_basis[month] = round((basis + current) / 2, 3)
                else:
                    month_to_basis[month] = float(basis)

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:"+url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website:"+url)
        return month_to_basis


def scrape_cvec(url):
    month_to_basis = {}
    try:
        time.sleep(10)
        driver.get(url)
        time.sleep(10)
        html_content = requests.get(url)
        time.sleep(10)
        a=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "DataGrid DataGridPlus")))
        soup = BeautifulSoup(html_content.content, features='lxml')
        table = soup.find_all('table', attrs={'class': 'DataGrid DataGridPlus'})[0].find_all('tr')
        for row in table[1:8]:
            month = row.find_all('td')[0].text.strip().lower()
            basis = float(row.find_all('td')[3].find('span').text.strip())
            index1 = month.find('/')
            index2 = month.find('/', index1 + 1, len(month))
            month = month[:index1+1] + '01' + month[index2:]
            month = datetime.strptime(month, '%m/%d/%y').date()
            if month in month_to_basis:
                current = month_to_basis[month]
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:", url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: {}".format(url))
        return month_to_basis


def delete_all_files(folder_path:str):
    files = glob.glob(folder_path+'*')
    if len(files)>0:
        for f in files:
            os.remove(f)

def scrape_homeland(url):
    month_to_basis = {}
    try:
        #download the image and save in local folder
        year = datetime.now().date().year
        html_content = requests.get(url)
        soup = BeautifulSoup(html_content.content, features='lxml')
        img = soup.find_all('img', attrs={'class': 'vc_single_image-img attachment-full'})
        img_url = img[0]['src']
        delete_all_files(download_path)
        img_data = requests.get(img_url).content
        with open(download_path+'data_file.png', 'wb') as handler:
            handler.write(img_data)
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract\tesseract.exe'
        str_img_content = pytesseract.image_to_string(download_path+'data_file.png')
        lst_bids = str_img_content.split('\n')[34:40]
        for row in lst_bids:
            month = row.split()[0].strip().lower()
            basis = float(row.split()[3].strip())
            month = '01' + month[0:3] + str(year)
            month = datetime.strptime(month, '%d%b%Y').date()
            month_to_basis[month] = basis            
        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:", url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: {}".format(url))
        return month_to_basis

# the function which is used for most of the types of webistes, handles multiple cases
def scrape_regular_website_2(url, find_by_option, basis_index, month_index=0, table_name='cashbids-data-table',
                             wait_by_option=0, time_flag=0,
                             xpath_for_table="", class_name='DataGrid DataGridPlus', row_start_index=1, table_index=0,
                             table_id='', iframe_xpath="", row_end_index=8):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        driver.get(url)
        if time_flag:
            time.sleep(10)
        if iframe_xpath:
            WebDriverWait(driver, 15).until(EC.frame_to_be_available_and_switch_to_it(
                driver.find_element_by_xpath(iframe_xpath)))

        if wait_by_option == 1:
            a=WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, xpath_for_table)))
        elif wait_by_option == 2:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.NAME, table_name)))
        elif wait_by_option == 3:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CLASS_NAME, class_name)))
        elif wait_by_option == 4:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, table_id)))
        
        soup = BeautifulSoup(driver.page_source, features='lxml')
        if find_by_option == 1:
            table = soup.find_all('table', attrs={'class': class_name})[table_index].find_all('tr')
        elif find_by_option == 2:
            table = soup.find_all('table', attrs={'xpath': xpath_for_table})[table_index].find_all('tr')
        elif find_by_option == 3:
            table = soup.find_all('table')[table_index].find_all('tr')
        elif find_by_option == 4:
            table = soup.find_all('table', attrs={'id': table_id})[table_index].find_all('tr')
    
        for row in table[row_start_index:row_end_index]:
            month = row.find_all('td')[month_index].text.strip().lower()
            basis = float(row.find_all('td')[basis_index].text.strip())
            raw_month = month
            month = month.replace('mch', 'mar') if 'mch' in month else month
            
            # for dates like - lh dec, fh dec
            if month[0:2] in ('lh', 'fh'):
                month = month[2:].strip()

            # for dates like - '12/01/2020'
            if month[:2].isdigit():
                month = month.replace(month[3:5], '01')
                month = datetime.strptime(month, '%m/%d/%Y').date()
                if month in month_to_basis:
                    current = month_to_basis[month]
                    if basis <= 0 and current <= 0:
                        month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                    else:
                        month_to_basis[month] = round((basis + current) / 2, 3)
                else:
                    month_to_basis[month] = basis
                continue

            # for dates like - december 2020
            if month[:3] in month_list and '/' not in month:
                if month[-2:].isdigit() and month[-2:] == str(year)[-2:]:
                    month = '01' + month[0:3] + '20' + month[-2:]
                    month = datetime.strptime(month, '%d%b%Y').date()
                    if month in month_to_basis:
                        current = month_to_basis[month]
                        if basis <= 0 and current <= 0:
                            month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                        else:
                            month_to_basis[month] = round((basis + current) / 2, 3)
                    else:
                        month_to_basis[month] = basis
                else:
                    # for dates like - december
                    month = '01' + month[0:3] + str(year)
                    month = datetime.strptime(month, '%d%b%Y').date()
                    if month in month_to_basis:
                        current = month_to_basis[month]
                        if basis <= 0 and current <= 0:
                            month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                        else:
                            month_to_basis[month] = round((basis + current) / 2, 3)
                    else:
                        month_to_basis[month] = basis
                    if month.month == 12:
                        year += 1
             #for dates jfm 21 --> Jan Feb March
            if 'JFM'.lower() in raw_month:
                for mth in list(month[0:3]):
                    if mth == 'j':
                        month = str(year) + '-01-01'
                    elif mth == 'f':
                        month =  str(year) + '-02-01'
                    elif mth == 'm':
                        month = str(year) + '-03-01'
                    month = datetime.strptime(month, '%Y-%m-%d').date()
                    month_to_basis[month] = basis
            #for dates April/May 21 like that
            if str(raw_month[0:3]).isalnum() and raw_month[:3] in month_list and '/' in raw_month:
                arr_months = raw_month.split()[0].split('/')
                if len(arr_months) > 0:
                    for mth in arr_months:
                        month = str(year) + '-' + str(month_number_dic[mth.lower()[:3]]) + '-01'
                        month = datetime.strptime(month, '%Y-%m-%d').date()
                        month_to_basis[month] = basis

        return month_to_basis
    except Exception:
        print(sys.exc_info()[0])
        print("error occoured with website:"+url)
        logger.info(sys.exc_info()[0])
        logger.info("error occoured with website: {}".format(url))
        return month_to_basis


# calls the scrape functions for regular type 1 and 2 and then inserts to sheet
def fetch_and_insert_regular_websitedata():
    #websites using regular scrape 1 method --
    bids = scrape_regular_website_1(url="http://www.glaciallakesenergy.com/corn_mina.htm", basis_index=2,
                                    iframe_xpath="/html/body/div[2]/table[5]/tbody/tr/td[2]/iframe")
    if insert_into_sheet(65, bids):
        print("success for row 65")
        logger.info("success for row 65")

    bids = scrape_regular_website_1(url="http://www.glaciallakesenergy.com/corn_mina.htm", basis_index=7,
                                    iframe_xpath="/html/body/div[2]/table[5]/tbody/tr/td[2]/iframe")
    if insert_into_sheet(67, bids):
        print("success for row 67")
        logger.info("success for row 67")

    bids = scrape_regular_website_1(url="http://dtn.pagrain.com/index.cfm", basis_index=-2)
    if insert_into_sheet(128, bids):
        print("success for row 128")
        logger.info("success for row 128")

    bids = scrape_regular_website_1(url="http://corn.eenergyadams.com/index.cfm?show=11&mid=6", basis_index=0)
    if insert_into_sheet(49, bids):
        print("success for row 49")
        logger.info("success for row 49")

    bids = scrape_regular_website_1(url="http://www.heronlakebioenergy.com/index.cfm?show=11&mid=8", basis_index=2)
    if insert_into_sheet(90, bids):
        print("success for row 90")
        logger.info("success for row 90")

    bids = scrape_regular_website_1(url="http://www.highwaterethanol.com/index.cfm?show=11&mid=36", basis_index=1)
    if insert_into_sheet(91, bids):
        print("success for row 91")
        logger.info("success for row 91")

    bids = scrape_regular_website_1(url="http://dtn.nebraskacornprocessing.com/index.cfm", basis_index=2)
    if insert_into_sheet(113, bids):
        print("success for row 113")
        logger.info("success for row 113")

    # websites using regular scrape 2 method --
    bids = scrape_regular_website_2(url="http://tallcornethanol.aghost.net/index.cfm?show=11&mid=3", wait_by_option=2,
                                    basis_index=-1,
                                    month_index=0, find_by_option=1)
    if insert_into_sheet(139, bids):
        print("success for row 139")
        logger.info("success for row 139")

    bids = scrape_regular_website_2(url="https://auroracoop.com/markets/", wait_by_option=3, find_by_option=3,
                                    class_name='section', month_index=0, basis_index=2, table_index=5, row_end_index=7)
    if insert_into_sheet(122, bids):
        insert_into_sheet(125, bids)
        print("success for row 122 and 125")
        logger.info("success for row 122 and 125")
        time.sleep(10)

    bids = scrape_regular_website_2(url="https://www.hankinsonre.com/janesville", wait_by_option=1, basis_index=3,
                                    month_index=1, find_by_option=1, class_name="cashbid_table cashbid_fulltable",
                                    row_start_index=2,
                                    xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[3]/div[1]/div[2]/div[2]/div["
                                                    "1]/table/tbody/tr/td/div/table")
    if insert_into_sheet(86, bids):
        print("success for row 86")
        logger.info("success for row 86")
        time.sleep(10)

    bids = scrape_regular_website_2(url="https://www.hankinsonre.com/hankinson", wait_by_option=1, basis_index=3,
                                    month_index=1, find_by_option=1, class_name="cashbid_table cashbid_fulltable",
                                    row_start_index=2,
                                    xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[2]/div["
                                                    "1]/table/tbody/tr/td/div/table")
    if insert_into_sheet(87, bids):
        print("success for row 87")
        logger.info("success for row 87")
        time.sleep(10)

    bids = scrape_regular_website_2(url="https://www.hankinsonre.com/lima", wait_by_option=1, basis_index=3,
                                    month_index=1,
                                    find_by_option=1, class_name="cashbid_table cashbid_fulltable", row_start_index=2,
                                    xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[4]/div[1]/div[2]/div[2]/div[1]"
                                                    "/table/tbody/tr/td/div/table")
    if insert_into_sheet(88, bids):
        print("success for row 88")
        logger.info("success for row 88")
        time.sleep(10)

    bids = scrape_regular_website_2(url="http://www.huskerag.com", wait_by_option=1, basis_index=4, month_index=1,
                                    find_by_option=3, table_index=9,
                                    xpath_for_table="/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[1]/div/table[4]/tbody/tr[3]/td[2]/table")
    if insert_into_sheet(93, bids):
        print("success for row 93")
        logger.info("success for row 93")

    bids = scrape_regular_website_2(url="http://www.ibecethanol.com/index.cfm?show=11", wait_by_option=3, basis_index=4,
                                    month_index=3, find_by_option=1, class_name="DataGrid", row_start_index=2)
    if insert_into_sheet(96, bids):
        print("success for row 96")
        logger.info("success for row 96")
        time.sleep(10)

    bids = scrape_regular_website_2(url="https://kaapaethanolcommodities.com/Commodities/Cash-Bids", basis_index=5,
                                    month_index=1, find_by_option=1, row_start_index=2, table_index=2,
                                    class_name="cashbid_table cashbid_fulltable", row_end_index=8)
    if insert_into_sheet(97, bids):
        print("success for row 97")
        logger.info("success for row 97")
        

    bids = scrape_regular_website_2(url="https://kaapaethanolcommodities.com/Commodities/Cash-Bids", basis_index=5,
                                    month_index=1, find_by_option=1, row_start_index=2, table_index=3, row_end_index=8
                                    ,class_name="cashbid_table cashbid_fulltable")
    if insert_into_sheet(98, bids):
        print("success for row 98")
        logger.info("success for row 98")
        

    bids = scrape_regular_website_2(url="http://www.granitefallsenergy.com/index.cfm?show=11&mid=41", wait_by_option=2,
                                    find_by_option=1, basis_index=2, class_name="DataGrid DataGridPlus DataNormal")
    if insert_into_sheet(72, bids):
        print("success for row 72")
        logger.info("success for row 72")

    bids = scrape_regular_website_2(url="https://www.ggecorn.com/bids", wait_by_option=4, table_id="dpTable1",
                                    find_by_option=3, basis_index=3)
    if insert_into_sheet(68, bids):
        print("success for row 68")
        logger.info("success for row 68")

    bids = scrape_regular_website_2(url="http://www.oneearthenergy.com", wait_by_option=3, month_index=1,
                                    find_by_option=1, basis_index=3, class_name="cb_table")
    if insert_into_sheet(116, bids):
        print("success for row 116")
        logger.info("success for row 116")

    bids = scrape_regular_website_2(url="http://www.ldnorfolk.com/index.cfm?show=11&mid=4", wait_by_option=2,
                                    find_by_option=1, basis_index=2)
    if insert_into_sheet(104, bids):
        print("success for row 104")
        logger.info("success for row 104")

    bids = scrape_regular_website_2(url="https://www.ldc.com/us/en/our-facilities/grand-junction-ia/cash-bids/",
                                    wait_by_option=1, find_by_option=3, basis_index=2, table_index=0,
                                    xpath_for_table="//*[@id=\"ldc-root\"]/article/div[1]/div/div[2]/div[2]"
                                                    "/section/div/div/div/div/div/table", row_end_index=11)
    if insert_into_sheet(103, bids):
        print("success for row 103")
        logger.info("success for row 103")

    bids = scrape_regular_website_2(url="https://www.littlesiouxcornprocessors.com", wait_by_option=4,
                                    find_by_option=4, basis_index=3, month_index=1, table_id="dpTable1", row_start_index=2, row_end_index=8)
    if insert_into_sheet(102, bids):
        print("success for row 102")
        logger.info("success for row 102")

    bids = scrape_regular_website_2(url="https://www.lincolnwayenergy.com/corn1.php", wait_by_option=2,
                                    find_by_option=1, basis_index=-1,
                                    iframe_xpath="/html/body/table[2]/tbody/tr/td[2]/table[2]/tbody/tr/td[1]/iframe", row_end_index=8)
    if insert_into_sheet(101, bids):
        print("success for row 101")
        logger.info("success for row 101")

    bids = scrape_regular_website_2(url="https://www.lincolnlandagrienergy.com/pages/custom.php?id=5427", wait_by_option=3,
                                    find_by_option=1, basis_index=3, month_index=1, class_name="homepage_quoteboard",
                                    row_start_index=2)
    if insert_into_sheet(100, bids):
        print("success for row 100")
        logger.info("success for row 100")

    bids = scrape_regular_website_2(url="http://www.sireethanol.com/index.cfm?show=11&mid=8", wait_by_option=2,
                                    find_by_option=1, basis_index=-1)
    if insert_into_sheet(175, bids):
        print("success for row 175")
        logger.info("success for row 175")

    bids = scrape_regular_website_2(url="http://www.southbendethanol.com/index.cfm?show=11&mid=3", wait_by_option=2,
                                    find_by_option=1, basis_index=-2)
    if insert_into_sheet(174, bids):
        print("success for row 174")
        logger.info("success for row 174")

    bids = scrape_regular_website_2(url="https://siouxlandethanol.com/cash-bids/", wait_by_option=1, find_by_option=3,
                                    basis_index=-2, xpath_for_table="/html/body/div[3]/div[1]/div/table",
                                    row_start_index=2)
    if insert_into_sheet(173, bids):
        print("success for row 173")
        logger.info("success for row 173")

    bids = scrape_regular_website_2(url="https://www.quad-county.com/markets/cash.php", wait_by_option=1, table_id="dpTable1",
                                    find_by_option=4, basis_index=-3, month_index=1,
                                    xpath_for_table="//*[@id=\"dpTable1\"]")
    if insert_into_sheet(163, bids):
        print("success for row 163")
        logger.info("success for row 163")

    bids = scrape_regular_website_2(url="http://www.redriverenergy.com/index.php", wait_by_option=3,
                                    find_by_option=1, basis_index=3, month_index=1, class_name="tbl")
    if insert_into_sheet(165, bids):
        print("success for row 165")
        logger.info("success for row 165")

    bids = scrape_regular_website_2(url="https://www.midmissourienergy.com/markets/cash.php", wait_by_option=4,
                                    find_by_option=4, basis_index=-3, month_index=1, table_id='dpTable1', row_end_index=13)
    if insert_into_sheet(111, bids):
        print("success for row 111")
        logger.info("success for row 111")

    bids = scrape_regular_website_2(url="https://www.andersonsgrain.com/locations/in/clymers/", wait_by_option=3,
                                    find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
    if insert_into_sheet(183, bids):
        print("success for row 183")
        logger.info("success for row 183")

    bids = scrape_regular_website_2(url="https://www.andersonsgrain.com/locations/oh/greenville/", wait_by_option=3,
                                    find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
    if insert_into_sheet(185, bids):
        print("success for row 185")
        logger.info("success for row 185")

    bids = scrape_regular_website_2(url="https://www.andersonsgrain.com/locations/ia/denison/", wait_by_option=3,
                                    find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
    if insert_into_sheet(184, bids):
        print("success for row 184")
        logger.info("success for row 184")

    bids = scrape_regular_website_2(url="https://www.andersonsgrain.com/locations/mi/albion/", wait_by_option=3,
                                    find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
    if insert_into_sheet(182, bids):
        print("success for row 182")
        logger.info("success for row 182")
        time.sleep(10)

    bids = scrape_regular_website_2(url="https://goldentriangleenergy.com/corn/", row_start_index=2,row_end_index=8,
                                   find_by_option=1, month_index=1, basis_index=4, class_name='homepage_quoteboard',
                                   iframe_xpath="/html/body/div[1]/div[2]/main/div/section/div/div/div[2]/div/div/div/iframe")
    if insert_into_sheet(69, bids):
        print("success for row 69")
        logger.info("success for row 69")

    bids = scrape_regular_website_2(url="https://www.pacificethanol.com/pekin-il-corn", row_start_index=3,
                                    find_by_option=3, basis_index=2, row_end_index=9)
    if insert_into_sheet(121, bids):
        insert_into_sheet(123, bids)
        insert_into_sheet(124, bids)
        print("success for row 121, 123, 124")
        logger.info("success for row 121, 123, 124")

    bids = scrape_regular_website_2(url="https://www.nugenmarion.com", row_start_index=2, table_id='dpTable1',
                                    wait_by_option=4, find_by_option=4, basis_index=4, month_index=1, row_end_index=8)
    if insert_into_sheet(115, bids):
        print("success for row 115")
        logger.info("success for row 115")

    bids = scrape_regular_website_2(url="http://www.cmgtharaldsonethanol.com/index.cfm?show=11&mid=5", row_start_index=1,
                                    wait_by_option=2, find_by_option=1, basis_index=4)
    if insert_into_sheet(181, bids):
        print("success for row 181")
        logger.info("success for row 181")

    bids = scrape_regular_website_2(url="https://www.unitedethanol.com/markets/cash.php?location_filter=18298",
                                    wait_by_option=3, find_by_option=1, basis_index=6, row_start_index=2, month_index=1,
                                    class_name="homepage_quoteboard")
    if insert_into_sheet(189, bids):
        print("success for row 189")
        logger.info("success for row 189")

    bids = scrape_regular_website_2(url="https://www.uwgp.com/grain/cash-bids-futures/", wait_by_option=4,
                                    find_by_option=4, basis_index=3, row_start_index=2, table_id='cashbids-data-table')
    if insert_into_sheet(190, bids):
        print("success for row 190")
        logger.info("success for row 190")


    valero_urls = {"https://valero-aurora.aghostportal.com/index.cfm?show=11&mid=3": 193,
                   "https://valero-albertcity.aghostportal.com/index.cfm?show=11&mid=3": 191,
                   "https://valero-bluffton.aghostportal.com/index.cfm?show=11&mid=3": 195,
                   "https://valero-hartley.aghostportal.com/index.cfm?show=11&mid=3": 198,
                   "https://valero-lakota.aghostportal.com/index.cfm?show=11&mid=3": 200,
                   "http://valero.aghostportal.com/index.cfm?show=11&mid=3": 203,
                   "https://valero-mtvernon.aghostportal.com/index.cfm?show=11&mid=3": 204}

    for url in valero_urls:
        bids = scrape_regular_website_2(url=url, wait_by_option=2, find_by_option=1, basis_index=2)
        if insert_into_sheet(valero_urls[url], bids):
            print("success for row " + str(valero_urls[url]))
            logger.info("success for row " + str(valero_urls[url]))
            time.sleep(10)
            if valero_urls[url] == 204:
                insert_into_sheet(201, bids)
                print("success for row 201")
                logger.info("success for row 201")

    bids=scrape_regular_website_2(url="https://valero-fortdodge.aghostportal.com/index.cfm?show=11&mid=3",wait_by_option=2,find_by_option=1,basis_index=1,
                                  row_start_index=1, class_name="DataGrid DataGridPlus")
    if insert_into_sheet(197,bids):
        print("success for row 197")
        logger.info("success for row 197")

    bids=scrape_regular_website_2(url="https://valero-charlescity.aghostportal.com/index.cfm?show=11&mid=3",wait_by_option=2,find_by_option=1,basis_index=2,
                                  row_start_index=1, class_name="DataGrid DataGridPlus")
    if insert_into_sheet(196,bids):
        print("success for row 196")
        logger.info("success for row 196")

    
    bids = poet_biorefining2("http://poetbiorefining-cloverdale.aghost.net/index.cfm?show=11&mid=27", 4)
    if insert_into_sheet(138, bids):
        print("success for row 138")
        logger.info("success for row 138")

    bids = poet_biorefining2("http://poetbiorefining-portland.aghost.net/index.cfm?show=11&mid=3", 3)
    if insert_into_sheet(156, bids):
        print("success for row 156")
        logger.info("success for row 156")

    bids = scrape_admfarm("https://www.admfarmview.com/cash-bids/bids/marshall")
    if insert_into_sheet(11, bids):
        print("success for row 11")
        logger.info("success for row 11")
        time.sleep(15)
        
    
    bids = scrape_admfarm("https://www.admfarmview.com/cash-bids/bids/cedarrapids")
    if insert_into_sheet(12, bids):
        print("success for row 12")
        logger.info("success for row 12")

    time.sleep(10)
    bids = scrape_admfarm("https://www.admfarmview.com/cash-bids/bids/columbuscorn")
    if insert_into_sheet(13, bids):
        print("success for row 13")
        logger.info("success for row 13")

    time.sleep(5)
    bids = scrape_admfarm("https://www.admfarmview.com/cash-bids/bids/columbuscorn")
    if insert_into_sheet(15, bids):
        print("success for row 15")
        logger.info("success for row 15")

    poetbiorefining_urls = {"https://poetbiorefining-alexandria.aghost.net/index.cfm?show=11&mid=3": [-1, 132],
                            "http://poetbiorefining-binghamlake.aghost.net/index.cfm?show=11&mid=3": [-1, 135],
                            "http://poetbiorefining-caro.aghost.net/index.cfm?show=11&mid=3": [-1, 136],
                          "https://poetbiorefining-chancellor.aghost.net/index.cfm?show=11&mid=3&ts=550964": [-1, 137],
                            "http://poetbiorefining-corning.aghost.net/index.cfm?show=11&mid=3": [-1, 140],
                            "http://poetbiorefining-fostoria.aghost.net/index.cfm?show=11&mid=3": [-1, 142],
                            "http://poetbiorefining-groton.aghost.net/index.cfm?show=11&mid=3": [-1, 145],
                            "http://poetbiorefining-laddonia.aghost.net/index.cfm?show=11&mid=3": [-1, 149],
                            "http://poetbiorefining-lakecrystal.aghost.net/index.cfm?show=11&mid=17": [-1, 150],
                            "http://poetbiorefining-leipsic.aghost.net/index.cfm?show=11&mid=5": [-1, 151],
                            "http://poetbiorefining-marion.aghost.net/index.cfm?show=11&mid=3": [-1, 152],
                           "http://poetbiorefining-marion.aghost.net/index.cfm?show=11&mid=3": [-1, 153],
                            "http://poetbiorefining-preston.aghost.net/index.cfm?show=11&mid=3": [-1, 157],
                            "http://shb.poetgrain.com/index.cfm?show=11&mid=3": [-1, 158],
                            "http://poetbiorefining-researchcenter.aghost.net/index.cfm?show=11&mid=3": [-1, 159],
                            "http://poetbiorefining-emmetsburg.aghost.net/index.cfm?show=11&mid=3": [2, 141],
                            "http://poetbiorefining-gowrie.aghost.net/index.cfm?show=11&mid=3": [1, 144],
                            "http://poetbiorefining-hanlontown.aghost.net/index.cfm?show=11&mid=3": [2, 146],
                            "http://poetbiorefining-hudson.aghost.net/index.cfm?show=11&mid=3": [3, 147],
                            "http://poetbiorefining-jewell.aghost.net/index.cfm?show=11&mid=3": [2, 148],
                            "https://poetbiorefining-macon.aghost.net/index.cfm?show=11&mid=3": [2, 152],
                          "http://poetbiorefining-mitchell.aghost.net/index.cfm?show=11&mid=3": [-2, 154]}

    for url in poetbiorefining_urls:
        bids = scrape_regular_website_2(url=url, wait_by_option=2, find_by_option=1, basis_index=poetbiorefining_urls[url][0])
        if insert_into_sheet(poetbiorefining_urls[url][1], bids):
            print("success for row " + str(poetbiorefining_urls[url][1]))
            logger.info("success for row " + str(poetbiorefining_urls[url][1]))

    bids = scrape_regular_website_2(url="http://poetbiorefining-northmanchester.aghost.net/index.cfm?show=11&mid=3",
                                    month_index=3, basis_index=5, class_name='DataGrid', row_start_index=2, wait_by_option=3,
                                    find_by_option=1)
    if insert_into_sheet(155, bids):
        print("success for row 155")
        logger.info("success for row 155")

    bids = scrape_regular_website_2(url="http://poetbiorefining-ashton.aghost.net/index.cfm?show=11&mid=3",
                                    month_index=0, basis_index=3, class_name='DataGrid DataGridPlus', row_start_index=1,wait_by_option=2,
                                    find_by_option=1)
    if insert_into_sheet(133, bids):
        print("success for row 133")
        logger.info("success for row 133")

    bids = scrape_regular_website_2(url="http://poetbiorefining-bigstone.aghost.net/index.cfm?show=11&mid=5&ts=527357",
                                    month_index=0, basis_index=4, class_name='DataGrid DataGridPlus', row_start_index=1,wait_by_option=2,
                                    find_by_option=1)
    if insert_into_sheet(134, bids):
        print("success for row 134")
        logger.info("success for row 134")

    bids = scrape_regular_website_2(url="https://www.wnyenergy.com/corn-bids/", class_name='cornbids',
                                    basis_index=3, wait_by_option=3, find_by_option=1)
    if insert_into_sheet(206, bids):
        print("success for row 206")
        logger.info("success for row 206")

    bids = scrape_regular_website_2(url="https://ekaellc.com/grain2/", month_index=2, basis_index=4,
                                    class_name='homepage_quoteboard', find_by_option=1,
                                    iframe_xpath="/html/body/div[1]/div[2]/div[2]/div/main/article/p/iframe")
    if insert_into_sheet(50, bids):
        print("success for row 50")
        logger.info("success for row 50")

    bids = scrape_regular_website_2(url="http://www.dencollc.com", iframe_xpath="/html/body/div[1]/div[2]/div[2]/iframe[2]",
                                    basis_index=4, find_by_option=1)
    if insert_into_sheet(46, bids):
        print("success for row 46")
        logger.info("success for row 46")

    bids = scrape_regular_website_2(url="https://www.dakotaethanol.com/index.cfm?show=11&mid=3",
                                    basis_index=2, find_by_option=1, wait_by_option=2, row_end_index=10)
    if insert_into_sheet(44, bids):
        print("success for row 44")
        logger.info("success for row 44")

    bids = scrape_regular_website_2(url="http://www.cie.us/corn_bids.php", class_name='homepage_quoteboard',
                                    month_index=1, basis_index=5, find_by_option=1, row_start_index=2,
                                    iframe_xpath="/html/body/div/div/iframe",row_end_index=8)
    if insert_into_sheet(35, bids):
        print("success for row 35")
        logger.info("success for row 35")

    bids = scrape_regular_website_2(url="http://www.cardinalethanol.com/markets/cash.php?location_filter=30179&showcwt=0",
                                    basis_index=6, month_index=1, find_by_option=4, wait_by_option=4, table_id='dpTable1')
    if insert_into_sheet(30, bids):
        print("success for row 30")
        logger.info("success for row 30")

    bids = scrape_regular_website_2(url="https://www.cgbioenergy.com/cash-bids/", table_id="cashbids-data-table",
                                    basis_index=2, find_by_option=4, wait_by_option=4)
    if insert_into_sheet(29, bids):
        print("success for row 29")
        logger.info("success for row 29")

    bids = scrape_regular_website_2(url="https://bushmillsethanol.com/corn-procurement-and-bids/", table_id="tablepress-4",
                                    basis_index=3, find_by_option=4, wait_by_option=4)
    if insert_into_sheet(26, bids):
        print("success for row 26")
        logger.info("success for row 26")

    bids = scrape_regular_website_2(url="http://dtn.al-corn.com/index.cfm?show=11&mid=17",
                                    basis_index=-1, find_by_option=1, wait_by_option=2)
    if insert_into_sheet(6, bids):
        print("success for row 6")
        logger.info("success for row 6")

    bids = scrape_regular_website_2(url="https://www.aceethanol.com/cash-bids/", table_id="cashbids-data-table",
                                    basis_index=-1, find_by_option=4, wait_by_option=2)
    if insert_into_sheet(3, bids):
        print("success for row 3")
        logger.info("success for row 3")

    bids = scrape_regular_website_2(url="http://www.bigriverbids.com/index.cfm?show=11&mid=17&theLocation=8&layout=19",
                                    basis_index=3, find_by_option=1, wait_by_option=2)
    if insert_into_sheet(19, bids):
        print("success for row 19")
        logger.info("success for row 19")

    bids = scrape_regular_website_2(url="http://phaellc.com/receiving/cash-bids/", wait_by_option=3, find_by_option=1,
                                    class_name='homepage_quoteboard', month_index=1, basis_index=3, row_start_index=2,
                                    iframe_xpath="//*[@id=\"post-917\"]/div/p/iframe")
    if insert_into_sheet(160, bids):
        print("success for row 160")
        logger.info("success for row 160")

    bids = scrape_regular_website_2(url="https://www.siouxlandenergy.com/markets/cash.php", wait_by_option=3, find_by_option=1,
                                    class_name='homepage_quoteboard', month_index=1, basis_index=-3, row_start_index=2)
    if insert_into_sheet(172, bids):
        print("success for row 172")
        logger.info("success for row 172")

    bids = scrape_regular_website_2(url="https://www.agtegra.com/grain/cash-bids?format=table&groupby=ccommodity&setLocation=3121&commodity=",
                                    find_by_option=1, class_name="cashbid_table cashbid_fulltable", wait_by_option=1,
                                    month_index=1, basis_index=-2, row_start_index=2,
                                    xpath_for_table="//*[@id=\"form\"]/main/div[2]/div/div/div[1]/div/table[2]", row_end_index=8)
    if insert_into_sheet(167, bids):
        print("success for row 167")
        logger.info("success for row 167")

    bids = scrape_regular_website_2(url="https://www.ringneckenergy.com/cashbids", find_by_option=1, wait_by_option=3,
                                    class_name="homepage_quoteboard", month_index=1, basis_index=4, row_start_index=2, row_end_index=8)
    if insert_into_sheet(169, bids):
        print("success for row 169")
        logger.info("success for row 169")

    bids = scrape_regular_website_2(url="https://www.eliteoctane.net", find_by_option=1, wait_by_option=1,
                                    class_name="cashbid_table cashbid_fulltable", month_index=1, basis_index=5,
                                    row_start_index=2, xpath_for_table="/html/body/div[3]/div[1]/div[2]/table")
    if insert_into_sheet(52, bids):
        print("success for row 52")
        logger.info("success for row 52")

    bids = scrape_regular_website_2(url="http://pce-coops.com/resources/cashbids/", find_by_option=1, month_index=1,
                                    basis_index=-3, class_name="homepage_quoteboard", row_start_index=2,
                                    iframe_xpath="//*[@id=\"post-70741\"]/div[3]/div/div/div/div/div/div/div/iframe", row_end_index=9)
    if insert_into_sheet(4, bids):
        print("success for row 4")
        logger.info("success for row 4")

    bids = scrape_regular_website_2(url="http://www.bigriverbids.com/index.cfm?show=11&mid=17&theLocation=2&layout=19",
                                    basis_index=3, wait_by_option=2, find_by_option=1)
    if insert_into_sheet(21, bids):
        print("success for row 21")
        logger.info("success for row 21")

    bids = scrape_regular_website_2(url="http://www.bigriverbids.com/index.cfm?show=11&mid=17&theLocation=1&layout=19",
                                    basis_index=3, wait_by_option=2, find_by_option=1)
    if insert_into_sheet(22, bids):
        print("success for row 22")
        logger.info("success for row 22")

    bids = scrape_regular_website_2(url="http://www.bigriverbids.com/index.cfm?show=11&mid=17&theLocation=5&layout=19",
                                    basis_index=5, find_by_option=1, wait_by_option=2, row_end_index=9)
    if insert_into_sheet(20, bids):
        print("success for row 20")
        logger.info("success for row 20")

    # bids = scrape_cvec("http://dtn.cvec.com/index.cfm?show=11&mid=3&cmid=1&layout=1034")
    # if insert_into_sheet(38, bids):
    #     print("success for row 38")
    #     logger.info("success for row 38")

    # bids = scrape_homeland("https://www.homelandenergysolutions.com/grain-bids/")
    # if insert_into_sheet(92, bids):
    #     print("success for row 92")
    #     logger.info("success for row 92")

# inserts the input dictionary to a sheet row, expects a dictionary with month-date
# as key and its basis as value and the row number
def insert_into_sheet(row_number, bids):
    if bids:
        try:
            columns = ['H', 'I', 'J', 'K', 'L', 'M']
            columns = [col + str(row_number) for col in columns]
            date_to_column, index = dict(), 0
            for col in columns:
                date_to_column[future_months[index]] = col
                index += 1
            for date in date_to_column:
                if date in bids:
                    xw.Range(date_to_column[date]).value = bids[date]
                else:
                    xw.Range(date_to_column[date]).value = '-'
            return True
        except Exception:
            print(sys.exc_info()[0])
            logger.info(sys.exc_info()[0])
            return False
    else:
        print("empty bids dictionary for row number : " + str(row_number))
        logger.info("empty bids dictionary for row number : " + str(row_number))
        return False

def main():
    global bid_prices
    try:
        starttime=datetime.now()
        logging.warning('Start work at {} ...'.format(starttime.strftime('%Y-%m-%d %H:%M:%S')))
        logger.info("initializing new sheet...")
        excel_app = xw.App(visible=False)
        bid_prices = excel_app.books.open(r"\\biourja.local\biourja\India Sync\India\Automated Reports\Corn Bid\Cornbids.xlsx")
        status = initialize_new_sheet(bid_prices)
        if status:
            print("new sheet created, starting the scraping process...")
            logger.info("new sheet created, starting the scraping process...")
        else:
            print("sheet already present, starting the scraping process...")
            logger.info("sheet already, starting the scraping process...")

        no_bids_row_numbers = ['H25', 'H48', 'H55', 'H110', 'H131', 'H186', 'H199']
        for row_num in no_bids_row_numbers:
            xw.Range(row_num).value = 'No Bids'
        
        bids = scrape_absenergy()
        if insert_into_sheet(2, bids):
            print("success for row 2")
            logger.info("success for row 2")
        
        bids = scrape_midwestagenergy()
        if insert_into_sheet(23, bids):
            print("success for row 23")
            logger.info("success for row 23")
        
        bids = scrape_frvethanol()
        if insert_into_sheet(61, bids):
            print("success for row 61")
            logger.info("success for row 61")
        
        scrape_and_insert_gpreinc()
        fetch_and_insert_fhr()
        fetch_and_insert_regular_websitedata()
        time.sleep(10)
        bid_prices.save()
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='SUCCESS - CORN BID PRICE AUTOMATION',mail_body = 'CORN BID PRICE AUTOMATION completed successfully, Attached logs', attachment_location = logfile)
    
    except Exception as ex:
        print("error occoured in main",ex)
        print(sys.exc_info()[0])
        logger.info("error occoured in main",ex)
        logger.info(sys.exc_info()[0])
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='FAILURE - CORN BID PRICE AUTOMATION',mail_body = 'CORN BID PRICE AUTOMATION failed, Attached logs', attachment_location = logfile)


    finally:
        excel_app.quit()
        driver.close()
        driver.quit()
        endtime=datetime.now()
        logging.warning('Complete work at {} ...'.format(endtime.strftime('%Y-%m-%d %H:%M:%S')))
        logging.warning('Total time taken: {} seconds'.format((endtime-starttime).total_seconds()))


if __name__ == "__main__":
    main()
