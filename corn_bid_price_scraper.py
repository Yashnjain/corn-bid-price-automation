import os
import sys
import time
import glob
import psutil
import logging
import requests
import bu_alerts
import bu_config
import numpy as np
import xlwings as xw
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from dateutil.relativedelta import relativedelta
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support import expected_conditions as EC


# creates a new sheet in the excel file and copies the default columns
# over to the new sheet. also creates new columns in the sheet.
# if the new sheet is already present, new sheet is not created then
def initialize_new_sheet(bid_prices):
    try:
        latest_sheet = bid_prices.sheets.active
        # opening the latest tab to copy initial columns from it --
        # change 0 to 1 for manual run
        new_sheet_name = str(datetime.now().month) + '.' + str(datetime.now().day - 0)
        logging.info(f"new sheet name is {new_sheet_name}")
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
            # bid_prices.close()
        else:
            return False
        return True
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured in initializing new sheet method : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured in initializing new sheet method : {e}")
        return False


# to scrape from a singular website.
# these return a dictionary of month date as keys and basis as values.
# returns an empty dict if it fails or exception occours
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
                    month = month.split()[2][:-1]
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
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: http://www.absenergy.org/grainbids.html (scrape_absenergy method): {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: http://www.absenergy.org/grainbids.html (scrape_absenergy method): {e}")
        return month_to_basis


# to scrape from a singular website
def scrape_midwestagenergy():
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        time.sleep(20)
        res = requests.get('https://www.midwestagenergy.com/cashbidssingle-1703')
        time.sleep(15)
        soup = BeautifulSoup(res.content, features='lxml')
        table = soup.find_all('div', attrs={'class': 'cashBidLocation'})[0].find_all('ul')
        for row in table[2:]:
            month = row.find_all('li')[0].text.strip().lower()
            basis = float(row.find_all('li')[2].text.strip())
            if month:
                if 'fh' in month or 'lh' in month:
                    month = month.split()[1]
                month = '01' + month + str(year)
                month = datetime.strptime(month, '%d%B%Y').date()
                month_to_basis[month] = basis
                if month.month == 12:
                    year += 1
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: https://www.midwestagenergy.com/fccp-blue-flint-bids-19639 (scrape_midwestagenergy) : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: https://www.midwestagenergy.com/fccp-blue-flint-bids-19639 (scrape_midwestagenergy): {e}")
        return month_to_basis


# to scrape from a singular website
def scrape_frvethanol(driver):
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        driver.get("https://www.frvethanol.com/cashbids/")
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'cashbids-data-table')))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('table', attrs={'id': 'cashbids-data-table'})[0].find_all('tr')
        for row in table[1:7]:
            month = row.find_all('td')[0].find('span').text.strip().lower()
            basis = row.find_all('td')[3].find('span').text.strip()
            if month:
                if 'fh' in month or 'lh' in month:
                    month = month.split()[1][:-1]
                month = '01' + month + str(year)
                try:
                    month = datetime.strptime(month, '%d%B%Y').date()
                except:
                    if 'sept' in month:
                        month = month.replace('sept','sep')
                    month = datetime.strptime(month, '%d%b%Y').date()
                month_to_basis[month] = float(basis)
                if month.month == 12:
                    year += 1
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: https://www.frvethanol.com/cashbids/  (scrape_frvethanol method): {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: https://www.frvethanol.com/cashbids/  (scrape_frvethanol method): {e}")
        return month_to_basis


# to scrape from a singular website, but different locations
def scrape_fhr(driver, url):
    month_to_basis = dict()
    try:
        driver.get(url)
        time.sleep(10)
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[1]')))
        ml = []
        bl = []
        for i in range(1,8):
            monthCheck=False
            month = driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[1]/p").text
            if "overrun" in month.lower():
                month = driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[1]/span").text
                month = month.split('—')[0]
                month = datetime.strptime(month, '%m/%d/%y').date()
                n_month = datetime.strftime(month.replace(day=1), "%B%d%Y")
                ml.append(n_month)
            # Handling 'October/November 2022'
            elif '/' in month:
                month = month[month.find('/')+1:]
                nMonth = month[:month.find('/')]+'01'+month[month.find('/'):].split()[-1]
                monthCheck = True
                
                for month in [month, nMonth]:
                    ml.append(month)
                    # Handling September Sept Sep Cases for date conversion
                    try:
                        month = datetime.strptime(month, '%B%d%Y').date()
                    except:
                        try:
                            month = datetime.strptime(month, '%b%d%Y').date()
                        except:
                            try:
                                month = driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[1]/p").text
                                month = month.split()
                                # Removing Last Letter Sept to Sep for date conversion
                                month = month[0][:-1]+'01'+month[1]
                                month = datetime.strptime(month, '%b%d%Y').date()
                            except Exception as e:
                                raise e
            else:
                if month[0:2].lower() in ('lh', 'fh'):
                    month = month[2:].strip()
                month = month.split()
                month = month[0]+'01'+month[1]
                ml.append(month)
                # Handling September Sept Sep Cases for date conversion
                try:
                    month = datetime.strptime(month, '%B%d%Y').date()
                except:
                    try:
                        month = datetime.strptime(month, '%b%d%Y').date()
                    except:
                        try:
                            month = driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[1]/p").text
                            month = month.split()
                            # Removing Last Letter Sept to Sep for date conversion
                            month = month[0][:-1]+'01'+month[1]
                            month = datetime.strptime(month, '%b%d%Y').date()
                        except Exception as e:
                            raise e
            if not monthCheck:
                basis = float(driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[4]").text.replace("$",""))
                bl.append(basis)
                if month in month_to_basis:
                    current = month_to_basis[month]
                    if basis <= 0 and current <= 0:
                        month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                    else:
                        month_to_basis[month] = round((basis + current) / 2, 3)
                else:
                    month_to_basis[month] = basis
            else:
                for month in [month, nMonth]:
                    basis = float(driver.find_element_by_xpath(f"/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[{i}]/div/div[4]").text.replace("$",""))
                    bl.append(basis)
                    
                    if month in month_to_basis:
                        current = month_to_basis[month]
                        if basis <= 0 and current <= 0:
                            month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                        else:
                            month_to_basis[month] = round((basis + current) / 2, 3)
                    else:
                        month_to_basis[month] = basis
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website:{url} (scrape_fhr method) : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website:{url} (scrape_fhr method) : {e}")
        return month_to_basis
# to scrape from a singular website
def scrape_ul_table(url, basis_index=2):
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        # time.sleep(20)
        res = requests.get(url)
        # time.sleep(15)
        soup = BeautifulSoup(res.content, features='lxml')
        table = soup.find_all('div', attrs={'class': 'cashBidLocation'})[0].find_all('ul')
        for row in table[1:]:
            month = row.find_all('li')[0].text.strip().lower()
            basis = float(row.find_all('li')[basis_index].text.strip())
            if month and "cont overfill" not in month.lower():
                if 'fh' in month or 'lh' in month:
                    month = month.split()[1]
                if '/' in month and (month[0:3].lower() not in ('f/h', 'l/h')):
                    arr_months = month.split()[0].split('/')
                    if len(arr_months) > 0:
                        for mth in arr_months:
                            month = str(year) + '-' + str(month_number_dic[mth.lower()[:3]]) + '-01'
                            month = datetime.strptime(month, '%Y-%m-%d').date()
                            month_to_basis[month] = basis
                else:
                    if month[0:3].lower() in ('f/h', 'l/h'):
                        month = month[3:].strip()
                    if "-" in month:
                        month = month.split('-')[0].split()[0]
                    if str(year)[2:] in month.split()[-1]:
                        month = month.replace(month.split()[-1], str(year))
                    if str(year) not in month:
                        month = '01' + month + str(year)
                    else:
                        month = '01' + month.replace(" ","")
                    try:
                        month = datetime.strptime(month, '%d%B%Y').date()
                    except:
                        try:
                            if "sept" in month:
                                month = month.replace("sept","sep")
                            month = datetime.strptime(month, '%d%b%Y').date()  
                        except Exception as e:
                            raise e
                    month_to_basis[month] = basis
                    if month.month == 12:
                        year += 1
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url}: {e}")
        return month_to_basis

# calls the fhr_scrape function and then insert_to_sheet function
def fetch_and_insert_fhr(driver):
    try:
        fhr_urls = {"https://www.fhr.com/corn-prices/arthur": 54, "https://www.fhr.com/corn-prices/fairbank": 56,
                    "https://www.fhr.com/corn-prices/Fairmont": 57, "https://www.fhr.com/corn-prices/iowa-falls": 58,
                    "https://www.fhr.com/corn-prices/Menlo": 59, "https://www.fhr.com/corn-prices/shell-rock": 60}

        for url in fhr_urls:
            bids = scrape_fhr(driver, url)
            if insert_into_sheet(fhr_urls[url], bids):
                print("success for row " + str(fhr_urls[url]))
                logging.info("success for row " + str(fhr_urls[url]))
                logging.info(f"inserted bids are: {bids}")
        return True
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured in fhr_urls(fetch_and_insert_fhr method) : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured in fhr_urls(fetch_and_insert_fhr method) : {e}")
        return False


# for singular website but different locations, scrapes and inserts one by one
def scrape_and_insert_gpreinc(driver):
    retry=0
    while retry<3:
        try:
            retry+=1
            #'Ord' is removed
            print("Starting gprenic")
            driver.get("http://www.gpreinc.com/corn-bids")
            city_select_values = {'Atkinson': 73, 'Central City': 74, 'Madison': 78, 'Mount Vernon': 79, 'Obion': 80,
                                'Shenandoah': 82, 'Superior': 83, 'York': 84}
            # WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/main/article/div/section[1]/div/div/div[2]/div/select/option[2]"))).click()
            # WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/main/article/div/section[1]/div/div/div[2]/div/select/option[1]"))).click()
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//select[@aria-label='Cash Bids Location Select']")))
            select = Select(driver.find_element_by_xpath("//select[@aria-label='Cash Bids Location Select']"))
            # select.select_by_value("Corn")
            for city in city_select_values:
                month_to_basis = dict()
                # driver.get("http://www.gpreinc.com/corn-bids")
                # select = Select(driver.find_element_by_xpath("//select[@aria-label='Cash Bids Location Select']"))
                select_element_scroll = WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//select[@aria-label='Cash Bids Location Select']")))
                try:
                    if city == 'Atkinson':
                        select_element = WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//select[@aria-label='Cash Bids Location Select']")))
                        select = Select(select_element)
                        
                        print(f"Selecting city: {city.lower()}")
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        select.select_by_value(city.lower())
                        
                    else:
                        select_element = WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//dtn-select[@label='Locations']//label//select")))
                        
                        
                        select = Select(select_element)
                        
                        print(f"Selecting city: {city.lower()}")
                        driver.execute_script("arguments[0].scrollIntoView(true);", select_element_scroll)
                        select.select_by_visible_text(city)
                except Exception as e:
                    try:
                        logging.exception(f"Inside exception {e}")
                        new_element = WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH,"//select[@aria-label='Cash Bids Location Select']")))
                        driver.execute_script("arguments[0].scrollIntoView(true);", new_element)
                        logging.info(f"Selecting again except for city: {city}")
                        select.select_by_visible_text(city)
                    except Exception as e:
                        raise e
                # select.select_by_visible_text(city)
                WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH,"//dtn-select[@label='Locations']//label//select")))
                soup = BeautifulSoup(driver.page_source, features='lxml')
                table = soup.find('table').find_all('tr')
                for row in table[1:7]:
                    month = row.find_all('td')[1].text.strip().lower()
                    basis = row.find_all('td')[3].text.strip()
                    if month[:3] in month_list:
                        month = '01' + month[0:3] + '20' + month[-2:]
                        month = datetime.strptime(month, '%d%b%Y').date()
                        month_to_basis[month] = basis
                print(f"Inserting {city_select_values[city]}")
                ########################################Uncomment after test done#########################
                if insert_into_sheet(city_select_values[city], month_to_basis):
                    print("success for row " + str(city_select_values[city]))
                    logging.info("success for row " + str(city_select_values[city]))
                    print(f"inserted bids are: {month_to_basis}")
                    logging.info(f"inserted bids are: {month_to_basis}")
                del month_to_basis
            return True
        except Exception as e:

            print(sys.exc_info()[0])
            print(f"error occured in gpreinc_urls (scrape_and_insert_gpreinc method): {e}")
            logging.info(sys.exc_info()[0])
            logging.info(f"error occured in gpreinc_urls (scrape_and_insert_gpreinc method): {e}")
            logging.exception(e)
            if retry==2:
                return False
            else:
                continue

    return False
# to scrape from a singular website
def poet_biorefining2(driver, url, basis_index):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        months, basis_values = list(), list()
        driver.get(url)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "DataGrid")))
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
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (poet_biorefining2 method) : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url} (poet_biorefining2 method): {e}")
        return month_to_basis


# to scrape from a singular website
def scrape_admfarm(driver, url):
    month_to_basis = dict()
    try:
        driver.get(url)
        try:
            WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div/div[3]/div/div/div[1]/div/div[2]/button"))).click()
        except:
            pass
        # Check is data is loaded or not via basis column
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div/div/div[3]/main/div/div[2]/div/div/div[2]/div[3]/div/div/div/div[2]/div/div[1]/div[2]/div/div/div/section/header/div[3]/div[1]")))
        
        for i in range(1,12):
            month = driver.find_element_by_xpath(f"/html/body/div/div/div/div/div[3]/main/div/div[2]/div/div/div[2]/div[3]/div/div/div/div[2]/div/div[1]/div[2]/div/div/div/section/div[{i}]/div[1]").text
            month = month.split('-')[0].split()
            month = month[0]+'01'+month[2]
            basis = driver.find_element_by_xpath(f"/html/body/div/div/div/div/div[3]/main/div/div[2]/div/div/div[2]/div[3]/div/div/div/div[2]/div/div[1]/div[2]/div/div/div/section/div[{i}]/div[3]/div").text
            basis = float(basis)
            month = datetime.strptime(month, '%b%d%Y').date()
            if month in month_to_basis:
                current = month_to_basis[month]
                print(current)
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_admfarm method), {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url} (scrape_admfarm method), {e}")
        return month_to_basis


# function to scrape from a type of webiste where month occours in the header
def scrape_regular_website_1(driver, url, basis_index, iframe_xpath=""):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        driver.get(url)
        if iframe_xpath:
            WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it(driver.find_element_by_xpath(iframe_xpath)))
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "DataGrid")))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('table', attrs={'class': 'DataGrid'})[0].find_all('tr')
        for row in table[2:10]:
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
                        month_to_basis[month] = round(((current/2)+(basis/2)), 3)
                    else:
                        month_to_basis[month] = round((basis + current) / 2, 3)
                else:
                    month_to_basis[month] = float(basis)
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_regular_website_1 method), {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url} (scrape_regular_website_1 method), {e}")
        return month_to_basis
    
    
def scrape_eliteoctane(driver, url):
    month_to_basis = dict()
    try:
        driver.get(url)
        logging.info("scrapping eliteoctane")
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[1]/div[2]/table/tbody/tr[1]/th[4]")))
        for i in range(2,8):
            month = driver.find_element_by_xpath(f"/html/body/div[3]/div[1]/div[2]/table/tbody/tr[{i}]/td[1]").text
            logging.info(month)
            if month == "":
                continue
            month = month.split()
            month = month[0]+'01'+month[1]
            try: 
                driver.find_element_by_xpath(f"/html/body/div[3]/div[1]/div[2]/table/tbody/tr[{i}]/td[7]")
                basis = driver.find_element_by_xpath(f"/html/body/div[3]/div[1]/div[2]/table/tbody/tr[{i}]/td[6]").text
            except:
                basis = driver.find_element_by_xpath(f"/html/body/div[3]/div[1]/div[2]/table/tbody/tr[{i}]/td[5]").text
            basis = float(basis)
            month = datetime.strptime(month, '%b%d%Y').date()
            if month in month_to_basis:
                current = month_to_basis[month]
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_eliteoctane method), {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url} (scrape_eliteoctane method), {e}")
        return month_to_basis

def scrape_ggcorn(driver, url,iframe_xpath):
    month_to_basis = {}
    try:
        
        driver.get(url)
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it(
                driver.find_element_by_xpath(iframe_xpath)))
        WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH, "//body//div//table")))
        # Range(1,6)
        for i in range(2,7):
            month = driver.find_element_by_xpath(f"//body//div//table/tbody/tr[{i}]/td[1]").text
            basis = driver.find_element_by_xpath(f"//body//div//table/tbody/tr[{i}]/td[5]").text
            basis = float(basis)
            #'Jul 2022'  %m/%d/%Yold format
            month = datetime.strptime(month, '%b %Y').date()
            if month in month_to_basis:
                current = month_to_basis[month]
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_ggcorn method), {e}")
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_ggcorn method), {e}")
        return month_to_basis
    
    
def scrape_cvec(driver, url):
    month_to_basis = {}
    try:
        driver.get(url)
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[2]/td/table[3]/thead/tr/th[2]/a")))
        for i in range(1,7):
            month = driver.find_element_by_xpath(f"/html/body/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[2]/td/table[3]/tbody/tr[{i}]/th").text
            month = month.replace("FH","")
            month = month.replace("LH","")
            month = month.replace("Sept","Sep")
            month = month.split()
            month = month[0]+'01'+month[1]
            logging.info(month)
            basis = driver.find_element_by_xpath(f"/html/body/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[2]/td/table[3]/tbody/tr[{i}]/td/a[1]").text
            basis = float(basis)
            month = datetime.strptime(month, '%b%d%y').date()
            
            if month in month_to_basis:
                current = month_to_basis[month]
                if basis <= 0 and current <= 0:
                    month_to_basis[month] = round((((-1 * basis) + (-1 * current)) / 2) * -1, 3)
                else:
                    month_to_basis[month] = round((basis + current) / 2, 3)
            else:
                month_to_basis[month] = basis
        logging.info("cvec Completed")
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} (scrape_cvec method), {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url} (scrape_cvec method), {e}")
        return month_to_basis


def delete_all_files(folder_path:str):
    try:
        files = glob.glob(folder_path+'*')
        if len(files)>0:
            for f in files:
                os.remove(f)
    except Exception as e:
        logging.info(sys.exc_info()[0])
        logging.info("error occured in delete_all_files method : {}".format(e))
        print("error occured in delete_all_files method : {}".format(e))
        raise e

def scrape_ul_table_with_driver(driver,url,iframe_xpath, basis_index=2):
    month_to_basis = dict()
    try:
        year = datetime.today().date().year
        driver.get(url)
        WebDriverWait(driver,60).until(EC.frame_to_be_available_and_switch_to_it(
                driver.find_element_by_xpath(iframe_xpath)))
        
        soup = BeautifulSoup(driver.page_source, features='lxml')
        table = soup.find_all('div', attrs={'class': 'cbCommodity'})[0].find_all('ul')
        for row in table[1:]:
            month = row.find_all('li')[0].text.strip().lower()
            basis = float(row.find_all('li')[basis_index].text.strip())
            if month and "cont overfill" not in month.lower():
                if 'fh' in month or 'lh' in month:
                    month = month.split()[1]
                if '/' in month and (month[0:3].lower() not in ('f/h', 'l/h')):
                    arr_months = month.split()[0].split('/')
                    if len(arr_months) > 0:
                        for mth in arr_months:
                            month = str(year) + '-' + str(month_number_dic[mth.lower()[:3]]) + '-01'
                            month = datetime.strptime(month, '%Y-%m-%d').date()
                            month_to_basis[month] = basis
                else:
                    if month[0:3].lower() in ('f/h', 'l/h'):
                        month = month[3:].strip()
                    if "-" in month:
                        month = month.split('-')[0].split()[0]
                    if str(year)[2:] in month.split()[-1]:
                        month = month.replace(month.split()[-1], str(year))
                    if str(year) not in month:
                        month = '01' + month + str(year)
                    else:
                        month = '01' + month.replace(" ","")
                    try:
                        month = datetime.strptime(month, '%d%B%Y').date()
                    except:
                        try:
                            if "sept" in month:
                                month = month.replace("sept","sep")
                            month = datetime.strptime(month, '%d%b%Y').date()  
                        except Exception as e:
                            raise e
                    month_to_basis[month] = basis
                    if month.month == 12:
                        year += 1
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url} : {e}")
        logging.info(sys.exc_info()[0])
        logging.info(f"error occured with website: {url}: {e}")
        return month_to_basis
# the function which is used for most of the types of webistes, handles multiple cases
def scrape_regular_website_2(driver, url, find_by_option, basis_index, month_index=0, table_name='cashbids-data-table',
    wait_by_option=0, time_flag=0,xpath_for_table="", class_name='DataGrid DataGridPlus', row_start_index=1, table_index=0,table_id='', iframe_xpath="", row_end_index=8):
    month_to_basis = dict()
    try:
        year = datetime.now().date().year
        driver.get(url)
        if time_flag:
            time.sleep(10)
        if iframe_xpath:
            try:
                WebDriverWait(driver,200,poll_frequency=5).until(EC.frame_to_be_available_and_switch_to_it(
                    driver.find_element_by_xpath(iframe_xpath)))
            except Exception as e:
                try:
                   time.sleep(2)
                   WebDriverWait(driver,10,poll_frequency=5).until(EC.frame_to_be_available_and_switch_to_it(
                    driver.find_element_by_xpath(iframe_xpath)))
                except Exception as e:
                    raise e
        if wait_by_option == 1:
            a=WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.XPATH, xpath_for_table)))
        elif wait_by_option == 2:
            WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.NAME, table_name)))
        elif wait_by_option == 3:
            WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.CLASS_NAME, class_name)))
        elif wait_by_option == 4:
            WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.ID, table_id)))
        soup = BeautifulSoup(driver.page_source, features='lxml')
        if find_by_option == 1:
            logging.info(f"table index is {table_index}")
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
            # for dates jfm 21 --> Jan Feb March
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
            # for dates April/May 21 like that
            if str(raw_month[0:3]).isalnum() and raw_month[:3] in month_list and '/' in raw_month:
                arr_months = raw_month.split()[0].split('/')
                if len(arr_months) > 0:
                    for mth in arr_months:
                        month = str(year) + '-' + str(month_number_dic[mth.lower()[:3]]) + '-01'
                        month = datetime.strptime(month, '%Y-%m-%d').date()
                        month_to_basis[month] = basis
        return month_to_basis
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured with website: {url}  (scrape_regular_website_2 method), {e}")
        logging.info(sys.exc_info()[0])
        logging.exception(f"error occured with website: {url} (scrape_regular_website_2 method), {e}")
        return month_to_basis


# calls the scrape functions for regular type 1 and 2 and then inserts to sheet
def fetch_and_insert_regular_websitedata(driver):
    try:
        # websites using regular scrape 1 method --
        bids = scrape_regular_website_1(driver, url="http://www.glaciallakesenergy.com/corn_mina.htm", basis_index=2,
                                        iframe_xpath="/html/body/div[2]/table[5]/tbody/tr/td[2]/iframe")
        if insert_into_sheet(65, bids):
            print("success for row 65")
            logging.info("success for row 65")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://www.glaciallakesenergy.com/corn_mina.htm", basis_index=7,
                                        iframe_xpath="/html/body/div[2]/table[5]/tbody/tr/td[2]/iframe")
        if insert_into_sheet(67, bids):
            print("success for row 67")
            logging.info("success for row 67")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://dtn.pagrain.com/index.cfm", basis_index=-2)
        if insert_into_sheet(128, bids):
            print("success for row 128")
            logging.info("success for row 128")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://corn.eenergyadams.com/index.cfm?show=11&mid=6", basis_index=0)
        if insert_into_sheet(49, bids):
            print("success for row 49")
            logging.info("success for row 49")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://www.heronlakebioenergy.com/index.cfm?show=11&mid=8", basis_index=2)
        if insert_into_sheet(90, bids):
            print("success for row 90")
            logging.info("success for row 90")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://www.highwaterethanol.com/index.cfm?show=11&mid=36", basis_index=1)
        if insert_into_sheet(91, bids):
            print("success for row 91")
            logging.info("success for row 91")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_1(driver, url="http://dtn.nebraskacornprocessing.com/index.cfm", basis_index=2)
        if insert_into_sheet(113, bids):
            print("success for row 113")
            logging.info("success for row 113")
            logging.info(f"inserted bids are: {bids}")

        # websites using regular scrape 2 method --
        bids = scrape_regular_website_2(driver, url="http://tallcornethanol.aghost.net/index.cfm?show=11&mid=3", wait_by_option=2,
                                        basis_index=-1,
                                        month_index=0, find_by_option=1)
        if insert_into_sheet(139, bids):
            print("success for row 139")
            logging.info("success for row 139")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://auroracoop.com/markets/", wait_by_option=3, find_by_option=3,
                                        class_name='section', month_index=0, basis_index=2, table_index=5, row_end_index=7)
        if insert_into_sheet(122, bids):
            insert_into_sheet(125, bids)
            print("success for row 122 and 125")
            logging.info("success for row 122 and 125")
            logging.info(f"inserted bids are: {bids}")
            time.sleep(10)
        bids = scrape_regular_website_2(driver, url="https://www.cargillag.com/check-prices?location=46790", 
                                        find_by_option=1, basis_index=-2,class_name = 'table__main bids-table__main')
        if insert_into_sheet(85, bids):
            print("success for row 85")
            logging.info("success for row 85")
            logging.info(f"inserted bids are: {bids}")
        bids = scrape_regular_website_2(driver, url="https://www.hankinsonre.com/janesville", wait_by_option=1, basis_index=4,
                                        month_index=0, find_by_option=1, class_name="cashbid_table",
                                        row_start_index=1,
                                        xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[3]/div[1]/div[3]/div[2]/div/table")
        if insert_into_sheet(86, bids):
            print("success for row 86")
            logging.info("success for row 86")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G93").value, wait_by_option=1, basis_index=4, month_index=1,
        #                                 find_by_option=3, table_index=9,
        #                                 xpath_for_table="/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[1]/div/table[4]/tbody/tr[3]/td[2]/table")
        bids = scrape_regular_website_2(driver, url=xw.Range("G93").value, wait_by_option=3, basis_index=5, month_index=0,
                                        find_by_option=1, table_index=0,
                                        class_name="cashbid_table")
        if insert_into_sheet(93, bids):
            print("success for row 93")
            logging.info("success for row 93")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url="http://www.ibecethanol.com/index.cfm?show=11", wait_by_option=3, basis_index=4,
        #                                 month_index=3, find_by_option=1, class_name="cbCommodity", row_start_index=2)
        bids = scrape_ul_table_with_driver(driver,url="http://www.ibecethanol.com/index.cfm?show=11", 
                                           iframe_xpath="/html/body/form/div[2]/div[2]/div/div[2]/div[2]/div[1]/div/p/iframe",basis_index=2)
        if insert_into_sheet(96, bids):
            print("success for row 96")
            logging.info("success for row 96")
            logging.info(f"inserted bids are: {bids}")
            time.sleep(10)

        bids = scrape_regular_website_2(driver, url="https://kaapaethanolcommodities.com/Commodities/Cash-Bids", basis_index=5,
                                        month_index=1, find_by_option=1, row_start_index=2, table_index=2,
                                        class_name="cashbid_table cashbid_fulltable", row_end_index=8)
        if insert_into_sheet(97, bids):
            print("success for row 97")
            logging.info("success for row 97")
            logging.info(f"inserted bids are: {bids}")
    

        bids = scrape_regular_website_2(driver, url="http://www.granitefallsenergy.com/corn-cash-bids/", wait_by_option=3,
                                        find_by_option=1, basis_index=4, class_name="cashbid_table")
        if insert_into_sheet(72, bids):
            print("success for row 72")
            logging.info("success for row 72")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.ggecorn.com/cash-bids%2Fcustomers", wait_by_option=4, table_id="dpTable1",
                                        find_by_option=3, basis_index=3)
        bids = scrape_regular_website_2(driver, iframe_xpath='//*[@id="iframe-03"]',url="https://www.ggecorn.com/cash-bids%2Fcustomers", wait_by_option=3, class_name="cashbid_table",
                                        find_by_option=1, basis_index=4)
        bids = scrape_ggcorn(driver, "https://www.ggecorn.com/cash-bids%2Fcustomers",iframe_xpath = "//span[@data-ux='Element']//iframe")
        if insert_into_sheet(68, bids):
            print("success for row 68")
            logging.info("success for row 68")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.oneearthenergy.com", wait_by_option=3, month_index=1,
                                        find_by_option=1, basis_index=3, class_name="cb_table")
        if insert_into_sheet(116, bids):
            print("success for row 116")
            logging.info("success for row 116")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.ldnorfolk.com/index.cfm?show=11&mid=4", 
                                        iframe_xpath="//iframe[@title='Embedded Content']",
                                        find_by_option=1, basis_index=-2,class_name = 'cashbid_table')
        if insert_into_sheet(104, bids):
            print("success for row 104")
            logging.info("success for row 104")
            logging.info(f"inserted bids are: {bids}")
        
        logging.info("Scaraping kapa 3rd table index")
        bids = scrape_regular_website_2(driver, url="https://kaapaethanolcommodities.com/Commodities/Cash-Bids", basis_index=5,
                                        month_index=1, find_by_option=1, row_start_index=2, table_index=3,
                                        class_name="cashbid_table cashbid_fulltable", row_end_index=8) #xpath_for_table="/html/body/form/div[4]/div[2]/div/div[4]/div/div[1]/div[2]/div/div[2]/table/thead/tr[2]/td[1]")
        if insert_into_sheet(98, bids):
            print("success for row 98")
            logging.info("success for row 98")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.ldc.com/us/en/our-facilities/grand-junction-ia/cash-bids/",row_start_index=2,wait_by_option=1, find_by_option=3, basis_index=2, table_index=0,xpath_for_table="/html/body/div[1]/div[3]/article/div[1]/div/div[2]/div[2]/section/div/div/div/div/div/table/tbody",row_end_index=11)
        if insert_into_sheet(103, bids):
            print("success for row 103")
            logging.info("success for row 103")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.littlesiouxcornprocessors.com", wait_by_option=4,
                                        find_by_option=4, basis_index=3, month_index=1, table_id="dpTable1", row_start_index=2, row_end_index=8)
        if insert_into_sheet(102, bids):
            print("success for row 102")
            logging.info("success for row 102")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G101").value, wait_by_option=2,
        #                                 find_by_option=1, basis_index=-1,
        #                                 iframe_xpath="/html/body/table[2]/tbody/tr/td[2]/table[2]/tbody/tr/td[1]/iframe", row_end_index=8)
        bids = scrape_regular_website_2(driver, url="https://www.lincolnwayenergy.com/corn-bids/", wait_by_option=3,
                                        find_by_option=1, basis_index=-2,class_name="cashbid_table",
                                        row_end_index=8)
        if insert_into_sheet(101, bids):
            print("success for row 101")
            logging.info("success for row 101")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.lincolnlandagrienergy.com/pages/custom.php?id=5427", wait_by_option=3,
                                        find_by_option=1, basis_index=3, month_index=1, class_name="homepage_quoteboard",
                                        row_start_index=2)
        if insert_into_sheet(100, bids):
            print("success for row 100")
            logging.info("success for row 100")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.hankinsonre.com/hankinson", wait_by_option=1, basis_index=4,
                                            month_index=0, find_by_option=1, class_name="cashbid_table",
                                            row_start_index=1, row_end_index=8,
                                            xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[1]/table")
        if insert_into_sheet(87, bids):
            print("success for row 87")
            logging.info("success for row 87")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.sireethanol.com/index.cfm?show=11&mid=8", wait_by_option=2,
                                        find_by_option=1, basis_index=-1)
        if insert_into_sheet(175, bids):
            print("success for row 175")
            logging.info("success for row 175")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.southbendethanol.com/index.cfm?show=11&mid=3", wait_by_option=2,
                                        find_by_option=1, basis_index=-2)
        if insert_into_sheet(174, bids):
            print("success for row 174")
            logging.info("success for row 174")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://siouxlandethanol.com/cash-bids/", wait_by_option=1, find_by_option=1,
                                        basis_index=-2, xpath_for_table="/html/body/div[3]/div[1]/div/div/div/div/table",class_name="cashbid_table cashbid_fulltable",
                                        row_start_index=2)
        if insert_into_sheet(173, bids):
            print("success for row 173")
            logging.info("success for row 173")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url="https://www.quad-county.com/cash-bids", wait_by_option=1, table_id="dpTable1",
        #                                 find_by_option=4, basis_index=-3, month_index=1,
        #                                 xpath_for_table="//*[@id=\"dpTable1\"]")
        bids = scrape_ul_table(url="https://www.quad-county.com/cash-bids", basis_index=2)
        if insert_into_sheet(163, bids):
            print("success for row 163")
            logging.info("success for row 163")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.redriverenergy.com/index.php", wait_by_option=3,
                                        find_by_option=1, basis_index=3, month_index=1, class_name="cashbid_table")
        if insert_into_sheet(165, bids):
            print("success for row 165")
            logging.info("success for row 165")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.midmissourienergy.com/markets/cash.php", wait_by_option=4,
                                        find_by_option=4, basis_index=-3, month_index=1, table_id='dpTable1', row_end_index=13)
        if insert_into_sheet(111, bids):
            print("success for row 111")
            logging.info("success for row 111")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.andersonsgrain.com/locations/in/clymers/", wait_by_option=3,
                                        find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
        if insert_into_sheet(183, bids):
            print("success for row 183")
            logging.info("success for row 183")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.andersonsgrain.com/locations/oh/greenville/", wait_by_option=3,
                                        find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
        if insert_into_sheet(185, bids):
            print("success for row 185")
            logging.info("success for row 185")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.andersonsgrain.com/locations/ia/denison/", wait_by_option=3,
                                        find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
        if insert_into_sheet(184, bids):
            print("success for row 184")
            logging.info("success for row 184")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.andersonsgrain.com/locations/mi/albion/", wait_by_option=3,
                                        find_by_option=1, basis_index=2, table_name='styled-table', class_name='styled-table')
        if insert_into_sheet(182, bids):
            print("success for row 182")
            logging.info("success for row 182")
            logging.info(f"inserted bids are: {bids}")
            time.sleep(10)

        bids = scrape_regular_website_2(driver, url="https://goldentriangleenergy.com/corn/", row_start_index=2,row_end_index=8,
                                    find_by_option=1, month_index=1, basis_index=4, class_name='homepage_quoteboard',
                                    wait_by_option=3,iframe_xpath="/html/body/div[1]/div[2]/main/div/section/div/div/div[2]/div/div/div/iframe")
        if insert_into_sheet(69, bids):
            print("success for row 69")
            logging.info("success for row 69")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.pacificethanol.com/pekin-il-corn", row_start_index=3,
                                        find_by_option=3, basis_index=2, row_end_index=9)
        if insert_into_sheet(121, bids):
            insert_into_sheet(123, bids)
            insert_into_sheet(124, bids)
            print("success for row 121, 123, 124")
            logging.info("success for row 121, 123, 124")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.nugenmarion.com", row_start_index=2, table_id='dpTable1',
                                        wait_by_option=4, find_by_option=4, basis_index=4, month_index=1, row_end_index=8)
        if insert_into_sheet(115, bids):
            print("success for row 115")
            logging.info("success for row 115")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.cmgtharaldsonethanol.com/index.cfm?show=11&mid=5", row_start_index=1,wait_by_option=2, find_by_option=1, basis_index=4)
        if insert_into_sheet(181, bids):
            print("success for row 181")
            logging.info("success for row 181")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.unitedethanol.com/markets/cash.php?location_filter=18298",
                                        wait_by_option=3, find_by_option=1, basis_index=6, row_start_index=2, month_index=1,
                                        class_name="homepage_quoteboard")
        if insert_into_sheet(189, bids):
            print("success for row 189")
            logging.info("success for row 189")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G190").value, wait_by_option=4,
        #                                 find_by_option=4, basis_index=3, row_start_index=2, table_id='cashbids-data-table')
        bids = scrape_ul_table(url="https://www.uwgp.com/cash-bids", basis_index=2)
        if insert_into_sheet(190, bids):
            print("success for row 190")
            logging.info("success for row 190")
            logging.info(f"inserted bids are: {bids}")


        valero_urls = {"https://valero-aurora.aghostportal.com/index.cfm?show=11&mid=3": 193,
                    "https://valero-albertcity.aghostportal.com/index.cfm?show=11&mid=3": 191,
                    "https://valero-bluffton.aghostportal.com/index.cfm?show=11&mid=3": 195,
                    "https://valero-hartley.aghostportal.com/index.cfm?show=11&mid=3": 198,
                    "https://valero-lakota.aghostportal.com/index.cfm?show=11&mid=3": 200,
                    "http://valero.aghostportal.com/index.cfm?show=11&mid=3": 203,
                    "https://valero-mtvernon.aghostportal.com/index.cfm?show=11&mid=3": 204}
        for url in valero_urls:
            bids = scrape_regular_website_2(driver, url=url, wait_by_option=2, find_by_option=1, basis_index=2)
            if insert_into_sheet(valero_urls[url], bids):
                print("success for row " + str(valero_urls[url]))
                logging.info("success for row " + str(valero_urls[url]))
                time.sleep(10)
                if valero_urls[url] == 204:
                    insert_into_sheet(201, bids)
                    print("success for row 201")
                    logging.info("success for row 201")
                    logging.info(f"inserted bids are: {bids}")

        bids=scrape_regular_website_2(driver, url="https://valero-fortdodge.aghostportal.com/index.cfm?show=11&mid=3",wait_by_option=2,find_by_option=1,basis_index=1,
                                    row_start_index=1, class_name="DataGrid DataGridPlus")
        if insert_into_sheet(197,bids):
            print("success for row 197")
            logging.info("success for row 197")
            logging.info(f"inserted bids are: {bids}")

        bids=scrape_regular_website_2(driver, url="https://valero-charlescity.aghostportal.com/index.cfm?show=11&mid=3",wait_by_option=2,find_by_option=1,basis_index=2,
                                    row_start_index=1, class_name="DataGrid DataGridPlus")
        if insert_into_sheet(196,bids):
            print("success for row 196")
            logging.info("success for row 196")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.hankinsonre.com/lima", wait_by_option=1, basis_index=3,
                                            month_index=1, row_end_index=8,
                                            find_by_option=1, class_name="cashbid_table", row_start_index=2,
                                            xpath_for_table="/html/body/div[1]/div[2]/div[2]/div/div[4]/div[1]/div[1]/div[2]/div[1]/table")
        if insert_into_sheet(88, bids):
                print("success for row 88")
                logging.info("success for row 88")
                logging.info(f"inserted bids are: {bids}")
                
        bids = poet_biorefining2(driver, "http://poetbiorefining-cloverdale.aghost.net/index.cfm?show=11&mid=27", 4)
        if insert_into_sheet(138, bids):
            print("success for row 138")
            logging.info("success for row 138")
            logging.info(f"inserted bids are: {bids}")

        time.sleep(5)
        bids = poet_biorefining2(driver, "http://shb.poetgrain.com/index.cfm?show=11&mid=3&theLocation=1&layout=1047", 3)
        if insert_into_sheet(158, bids):
            print("success for row 158")
            logging.info("success for row 158")
            logging.info(f"inserted bids are: {bids}")

        bids = poet_biorefining2(driver, "http://poetbiorefining-portland.aghost.net/index.cfm?show=11&mid=3", 3)
        if insert_into_sheet(156, bids):
            print("success for row 156")
            logging.info("success for row 156")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_admfarm(driver, "https://www.admfarmview.com/cash-bids/bids/marshall")
        if insert_into_sheet(11, bids):
            print("success for row 11")
            logging.info("success for row 11")
            time.sleep(15)
        
        bids = scrape_admfarm(driver, "https://www.admfarmview.com/cash-bids/bids/cedarrapids")
        if insert_into_sheet(12, bids):
            print("success for row 12")
            logging.info("success for row 12")
            logging.info(f"inserted bids are: {bids}")

        time.sleep(10)
        bids = scrape_admfarm(driver, "https://www.admfarmview.com/cash-bids/bids/columbuscorn")
        if insert_into_sheet(13, bids):
            print("success for row 13")
            logging.info("success for row 13")
            logging.info(f"inserted bids are: {bids}")

        time.sleep(5)
        bids = scrape_admfarm(driver, "https://www.admfarmview.com/cash-bids/bids/columbuscorn")
        if insert_into_sheet(15, bids):
            print("success for row 15")
            logging.info("success for row 15")
            logging.info(f"inserted bids are: {bids}")

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
                                "http://poetbiorefining-mitchell.aghost.net/index.cfm?show=11&mid=3": [-2, 154],
                                "http://poetbiorefining-researchcenter.aghost.net/index.cfm?show=11&mid=3": [-1, 159],
                                "http://poetbiorefining-emmetsburg.aghost.net/index.cfm?show=11&mid=3": [2, 141],
                                "http://poetbiorefining-gowrie.aghost.net/index.cfm?show=11&mid=3": [1, 144],
                                "http://poetbiorefining-hanlontown.aghost.net/index.cfm?show=11&mid=3": [2, 146],
                                "http://poetbiorefining-hudson.aghost.net/index.cfm?show=11&mid=3": [3, 147],
                                "http://poetbiorefining-jewell.aghost.net/index.cfm?show=11&mid=3": [2, 148],
                                "https://poetbiorefining-macon.aghost.net/index.cfm?show=11&mid=3": [2, 152] }
                            
        for url in poetbiorefining_urls:
            bids = scrape_regular_website_2(driver, url=url, wait_by_option=2, find_by_option=1, basis_index=poetbiorefining_urls[url][0])
            if insert_into_sheet(poetbiorefining_urls[url][1], bids):
                print("success for row " + str(poetbiorefining_urls[url][1]))
                logging.info("success for row " + str(poetbiorefining_urls[url][1]))
                logging.info(f"inserted bids are: {bids}")
        time.sleep(5)       
        bids = scrape_regular_website_2(driver, url="http://poetbiorefining-northmanchester.aghost.net/index.cfm?show=11&mid=3",
                                        month_index=3, basis_index=5, class_name='DataGrid', row_start_index=2, wait_by_option=1,row_end_index=7,
                                        find_by_option=1, xpath_for_table="/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[2]/tbody/tr[2]/td/table[2]/tbody/tr/td/table")
        if insert_into_sheet(155, bids):
            print("success for row 155")
            logging.info("success for row 155")
            logging.info(f"inserted bids are: {bids}")
            
        bids = scrape_regular_website_2(driver, url="http://poetbiorefining-ashton.aghost.net/index.cfm?show=11&mid=3",
                                        month_index=0, basis_index=3, class_name='DataGrid DataGridPlus', row_start_index=1,wait_by_option=2,
                                        find_by_option=1)
        if insert_into_sheet(133, bids):
            print("success for row 133")
            logging.info("success for row 133")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://poetbiorefining-bigstone.aghost.net/index.cfm?show=11&mid=5&ts=527357",
                                        month_index=0, basis_index=5, class_name='DataGrid DataGridPlus', row_start_index=1,wait_by_option=2,
                                        find_by_option=1)
        if insert_into_sheet(134, bids):
            print("success for row 134")
            logging.info("success for row 134")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.wnyenergy.com/corn-bids/", class_name='cornbids',
                                        basis_index=3, wait_by_option=3, find_by_option=1)
        if insert_into_sheet(206, bids):
            print("success for row 206")
            logging.info("success for row 206")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://ekaellc.com/grain2/", month_index=2, basis_index=4,
                                        class_name='homepage_quoteboard', find_by_option=1,
                                        iframe_xpath="/html/body/div[1]/div[2]/div[2]/div/main/article/p/iframe")
        if insert_into_sheet(50, bids):
            print("success for row 50")
            logging.info("success for row 50")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.dencollc.com",basis_index=4, find_by_option=1, xpath_for_table="/html/body/div[1]/div[2]/div[2]/b[5]/div[1]/table",
                                        wait_by_option = 1,class_name="cashbid_table") # iframe_xpath="/html/body/div[1]/div[2]/div[2]/iframe[2]",xpath"/html/body/div[1]/div[2]/div[2]/div[1]/table/tbody/tr[1]/th[4]"
                                        
        if insert_into_sheet(46, bids):
            print("success for row 46")
            logging.info("success for row 46")
            logging.info(f"inserted bids are: {bids}")
    

        bids = scrape_regular_website_2(driver, url="https://www.dakotaethanol.com/index.cfm?show=11&mid=3",
                                        basis_index=2, find_by_option=1, wait_by_option=2, row_end_index=10)
        if insert_into_sheet(44, bids):
            print("success for row 44")
            logging.info("success for row 44")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.cie.us/corn_bids.php", class_name='homepage_quoteboard',
                                        month_index=1, basis_index=5, find_by_option=1, row_start_index=2,
                                        iframe_xpath="/html/body/div/div/iframe",row_end_index=8)
        if insert_into_sheet(35, bids):
            print("success for row 35")
            logging.info("success for row 35")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://www.cardinalethanol.com/markets/cash.php?location_filter=30179&showcwt=0",
                                        basis_index=6, month_index=1, find_by_option=4, wait_by_option=4, table_id='dpTable1')
        if insert_into_sheet(30, bids):
            print("success for row 30")
            logging.info("success for row 30")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.cgbioenergy.com/cash-bids/", table_id="cashbids-data-table",
                                        basis_index=2, find_by_option=4, wait_by_option=4)
        if insert_into_sheet(29, bids):
            print("success for row 29")
            logging.info("success for row 29")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://bushmillsethanol.com/corn-procurement-and-bids/",class_name="cashbid_table",basis_index=5, find_by_option=1, wait_by_option=3)
        if insert_into_sheet(26, bids):
            print("success for row 26")
            logging.info("success for row 26")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://dtn.al-corn.com/index.cfm?show=11&mid=17",
                                        basis_index=-1, find_by_option=1, wait_by_option=2,table_index=0)
        if insert_into_sheet(6, bids):
            print("success for row 6")
            logging.info("success for row 6")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.aceethanol.com/cash-bids/", table_id="cashbids-data-table",
                                        basis_index=-1, find_by_option=4, wait_by_option=2)
        if insert_into_sheet(3, bids):
            print("success for row 3")
            logging.info("success for row 3")
            logging.info(f"inserted bids are: {bids}")
        
        # bids = scrape_regular_website_2(driver, url=xw.Range("G19").value,
        #                                 basis_index=3, find_by_option=1, wait_by_option=2)
        bids = scrape_ul_table(url="http://www.bigriverbids.com/cashbidssingle-2121", basis_index=2)
        if insert_into_sheet(19, bids):
            print("success for row 19")
            logging.info("success for row 19")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://phaellc.com/receiving/cash-bids/", wait_by_option=3, find_by_option=1,
                                        class_name='homepage_quoteboard', month_index=1, basis_index=3, row_start_index=2,
                                        iframe_xpath="//*[@id=\"post-917\"]/div/p/iframe")
        if insert_into_sheet(160, bids):
            print("success for row 160")
            logging.info("success for row 160")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.siouxlandenergy.com/markets/cash.php", wait_by_option=3, find_by_option=1,
                                        class_name='table-responsive', month_index=1, basis_index=-3, row_start_index=1)#'homepage_quoteboard'row_start_index=2
        if insert_into_sheet(172, bids):
            print("success for row 172")
            logging.info("success for row 172")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G167").value,
        #                                 find_by_option=1, class_name="cashbid_table cashbid_fulltable", wait_by_option=1,
        #                                 month_index=1, basis_index=-2, row_start_index=2,
        #                                 xpath_for_table="/html/body/form/main/div[2]/div/div/div[1]/div/div/div[2]/table",
        #                                 row_end_index=8)
        bids = scrape_regular_website_2(driver, url="https://www.agtegra.com/cash-bids?format=table&groupby=ccommodity&setLocation=3121&commodity=",
                                        find_by_option=1, class_name="cashbid_table cashbid_fulltable", wait_by_option=1,
                                        month_index=1, basis_index=-2, row_start_index=2,xpath_for_table="/html/body/main/div/div/div/div[2]/div[2]/div/table",
                                        row_end_index=8)
        if insert_into_sheet(167, bids):
            print("success for row 167")
            logging.info("success for row 167")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.ringneckenergy.com/cashbids", find_by_option=1, wait_by_option=3,
                                        class_name="cashbid_table", month_index=1, basis_index=5, row_start_index=2, row_end_index=9)
        if insert_into_sheet(169, bids):
            print("success for row 169")
            logging.info("success for row 169")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_eliteoctane(driver, url="https://www.eliteoctane.net")
        if insert_into_sheet(52, bids):
            print("success for row 52")
            logging.info("success for row 52")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="http://pce-coops.com/resources/cashbids/", find_by_option=1, month_index=1,
                                        basis_index=-3, time_flag=1, class_name="homepage_quoteboard", row_start_index=2,wait_by_option=3,
                                        iframe_xpath="//iframe[@src='https://pce-coops.agricharts.com/markets/cash.php']", row_end_index=9)
                                        #"//*[@id=\"post-70741\"]/div[3]/div/div/div/div/div/div/div/iframe"
      
        if insert_into_sheet(4, bids):
            print("success for row 4")
            logging.info("success for row 4")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G21").value,
        #                                 basis_index=3, wait_by_option=2, find_by_option=1)
        bids = scrape_ul_table(url="https://www.bigriverbids.com/cashbidssingle-2164", basis_index=2)
        if insert_into_sheet(21, bids):
            print("success for row 21")
            logging.info("success for row 21")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G22").value,
        #                                 basis_index=3, wait_by_option=2, find_by_option=1)
        bids = scrape_ul_table(url="https://www.bigriverbids.com/cashbidssingle-2162", basis_index=2)
        if insert_into_sheet(22, bids):
            print("success for row 22")
            logging.info("success for row 22")
            logging.info(f"inserted bids are: {bids}")

        # bids = scrape_regular_website_2(driver, url=xw.Range("G20").value,
        #                                 basis_index=5, find_by_option=1, wait_by_option=2, row_end_index=9)
        bids = scrape_ul_table(url="https://www.bigriverbids.com/cashbidssingle-2163", basis_index=2)
        if insert_into_sheet(20, bids):
            print("success for row 20")
            logging.info("success for row 20")
            logging.info(f"inserted bids are: {bids}")
        logging.info("Downloading basis price from cvec")
        bids = scrape_cvec(driver, "http://dtn.cvec.com/index.cfm?show=11&mid=3&cmid=1&layout=1034")
        if insert_into_sheet(38, bids):
            print("success for row 38")
            logging.info("success for row 38")
            logging.info(f"inserted bids are: {bids}")

        bids = scrape_regular_website_2(driver, url="https://www.homelandenergysolutions.com/grain-bids/",class_name="cashbid_table",basis_index=-2, find_by_option=1, wait_by_option=3, row_end_index=9)
        if insert_into_sheet(92, bids):
            print("success for row 92")
            logging.info("success for row 92")
            logging.info(f"inserted bids are: {bids}")
    except Exception as e:
        print(sys.exc_info()[0])
        print(f"error occured in fetch_and_insert_regular_websitedata (fetch_and_insert_regular_websitedata method) : {e}")
        logging.info(sys.exc_info()[0])
        logging.info("error occured in fetch_and_insert_regular_websitedata (fetch_and_insert_regular_websitedata method) : {}".format(e))
        raise e


# as key and its basis as value and the row number
# inserts the input dictionary to a sheet row, expects a dictionary with month-date
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
        except Exception as e:
            logging.info(f"Caught Exception in insert rows for {row_number}")
            print(sys.exc_info()[0])
            logging.exception(e)
            logging.info(sys.exc_info()[0])
            return False
    else:
        print("empty bids dictionary for row number : " + str(row_number))
        logging.info("empty bids dictionary for row number : " + str(row_number))
        return False


def kill_excel():
    try:
        print('1. Kill the existing excel process')
        logging.info('1. Kill the existing excel process')
        for proc in psutil.process_iter():
            if proc.name() == "excel.exe":
                print('process name {}'.format(proc.name()))
                logging.info('process name {}'.format(proc.name()))
                proc.kill()
            elif proc.name() == "EXCEL.EXE":
                print('process name {}'.format(proc.name()))
                logging.info('process name {}'.format(proc.name()))
                proc.kill()
    except Exception as e:
        print(f"Unable to kill excel due to {e}")
        logging.info(f"Unable to kill excel due to {e}")
        raise e


def main(bid_price_sheet):    
    global bid_prices
    #initializing sheet for single index.
    try:
        # kill_excel()
        mime_types=['application/pdf'
                            ,'text/plain',
                            'application/vnd.ms-excel',
                            'test/csv',
                            'application/zip',
                            'application/csv',
                            'text/comma-separated-values','application/download','application/octet-stream'
                            ,'binary/octet-stream'
                            ,'application/binary'
                            ,'application/x-unknown']
        profile = webdriver.FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.manager.showWhenStarting', False)
        profile.set_preference('pdfjs.disabled', True)
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ','.join(mime_types))
        profile.set_preference('browser.helperApps.neverAsk.openFile',','.join(mime_types))
        driver = webdriver.Firefox(executable_path=GeckoDriverManager().install(), firefox_profile=profile)
        driver.maximize_window()
        # logging.info("initializing new sheet...")
        # # excel_app = xw.App(visible=True)
        # # bid_prices = excel_app.books.open(bid_price_sheet)
        # bid_prices = xw.Book(bid_price_sheet,update_links=False)
        # status = initialize_new_sheet(bid_prices)
        # if status:
        #     print("new sheet created, starting the scraping process...")
        #     logging.info("new sheet created, starting the scraping process...")
        # else:
        #     print("sheet already present, starting the scraping process...")
        #     logging.info("sheet already, starting the scraping process...")

        # no_bids_row_numbers = ['H25', 'H48', 'H55', 'H110', 'H131', 'H186', 'H199']
        # for row_num in no_bids_row_numbers:
        #     xw.Range(row_num).value = 'No Bids'
        
        # bids = scrape_absenergy()
        # if insert_into_sheet(2, bids):
        #     print("success for row 2")
        #     logging.info("success for row 2")
        #     logging.info(f"inserted bids are: {bids}")
        # time.sleep(10)
        # bid_prices.save()
        # bid_prices.close()
        
    except Exception as ex:
        print("error occured in main",ex)
        print(sys.exc_info()[0])
        logging.info("error occured in main",ex)
        logging.info(sys.exc_info()[0])
        raise ex
    finally:
        try:
            bid_prices.app.quit()
        except:
            pass
    #starting the complete process 
    time.sleep(10)
    try:
        # kill_excel()
        logging.info("initializing new sheet...")
        # excel_app = xw.App(visible=True)
        # bid_prices = excel_app.books.open(bid_price_sheet)
        bid_prices = xw.Book(bid_price_sheet,update_links=False)
        status = initialize_new_sheet(bid_prices)
        if status:
            print("new sheet created, starting the scraping process...")
            logging.info("new sheet created, starting the scraping process...")
        else:
            print("sheet already present, starting the scraping process...")
            logging.info("sheet already, starting the scraping process...")
        no_bids_row_numbers = ['H25', 'H48', 'H55', 'H110', 'H131', 'H186', 'H199']
        for row_num in no_bids_row_numbers:
            xw.Range(row_num).value = 'No Bids'
            
        bids = scrape_absenergy()
        if insert_into_sheet(2, bids):
            print("success for row 2")
            logging.info("success for row 2")
            logging.info(f"inserted bids are: {bids}")
            
        bids = scrape_midwestagenergy()
        if insert_into_sheet(23, bids):
            print("success for row 23")
            logging.info("success for row 23")
            logging.info(f"inserted bids are: {bids}")
        
        bids = scrape_frvethanol(driver)
        if insert_into_sheet(61, bids):
            print("success for row 61")
            logging.info("success for row 61")
            logging.info(f"inserted bids are: {bids}")
        
        scrape_and_insert_gpreinc(driver)
        fetch_and_insert_fhr(driver)
        fetch_and_insert_regular_websitedata(driver)
        time.sleep(10)
        bid_prices.save()
        logging.info("saved file finally")
    except Exception as ex:
        print("error occured in main",ex)
        print(sys.exc_info()[0])
        logging.info("error occured in main",ex)
        logging.info(sys.exc_info()[0])
        raise ex
    finally:
        try:
            bid_prices.app.quit()
        except:
            pass
        try:
            driver.quit()
        except:
            pass

def corn_bid_runner():
    global month_list, future_months , month_number_dic 
    logging.info('Execution Started')
    time_start = time.time()
    try:
        job_id=np.random.randint(1000000,9999999)
        logfile = os.getcwd() + '\\logs\\CORN_BID_PRICE_SCRAPER.txt'
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)
        
        credential_dict = bu_config.get_config('CORN_BID_PRICE_SCRAPER', 'N',other_vert= True)
        database=credential_dict['DATABASE'].split(";")[0]
        warehouse=credential_dict['DATABASE'].split(";")[1]
        table_name = credential_dict['TABLE_NAME']
        bid_price_sheet= credential_dict["API_KEY"]
        
        job_name = credential_dict['PROJECT_NAME']
        owner = credential_dict['IT_OWNER']
        receiver_email = credential_dict['EMAIL_LIST']
        
        ######################## Uncomment For Testing###########################
        # database="BUITDB_DEV"
        # warehouse="BUIT_WH"
        # receiver_email="amanullah.khan@biourja.com,deep.durugkar@biourja.com,imam.khan@biourja.com,yashn.jain@biourja.com"
        # # DRIVER_PATH = r'S:\IT Dev\Production_Environment\corn-bid-price-automation\geckodriver.exe'
        # # DRIVER_PATH = r'S:\IT Dev\Production_Environment\chromedriver\chromedriver.exe'
        # bid_price_sheet = r"E:\testingEnvironment\J_local_drive\India\Automated Reports\Corn Bid\Cornbids.xlsx"
        job_name = "BIO-PAD01_"+job_name
        #########################################################################

        # download_path = os.getcwd() + '\\download\\'
        headers = {'User-Agent': 'Mozilla/5.0'}
        options = Options()
        
        # month list used to check name of months on websites
        month_list = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        month_number_dic = {'jan': '01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06','jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'}

        # this section is used to define future months
        # that are to be set as new columns in the new sheet
        base_date = datetime.today().date().replace(day=1)
        future_months = list()
        for i in range(0, 6):
            future_months.append(base_date + relativedelta(months=i))
            
        #In prod bualerts works on  POWERDB  ITPYTHON_WH
        # BU_LOG entry(started) in PROCESS_LOG table
        log_json = '[{"JOB_ID": "'+str(job_id)+'","JOB_NAME": "'+str(job_name)+'","CURRENT_DATETIME": "'+str(datetime.now())+'","STATUS": "STARTED"}]'
        bu_alerts.bulog(process_name=job_name,table_name=table_name,status='STARTED',process_owner=owner ,row_count=0,log=log_json,database=database,warehouse=warehouse)
        
        main(bid_price_sheet)
        
        # BU_LOG entry(completed) in PROCESS_LOG table
        log_json = '[{"JOB_ID": "'+str(job_id)+'","JOB_NAME": "'+str(job_name)+'","CURRENT_DATETIME": "'+str(datetime.now())+'","STATUS": "COMPLETED"}]'
        bu_alerts.bulog(process_name=job_name,table_name=table_name,status='COMPLETED',process_owner=owner ,row_count=1,log=log_json,database=database,warehouse=warehouse)
        
        logging.info('Execution Done')
        bu_alerts.send_mail(
            receiver_email = receiver_email,
            mail_subject =f'JOB SUCCESS - {job_name}',
            mail_body = f'{job_name} completed successfully, Attached logs',
            attachment_location = logfile)
            
    except Exception as e:
        print("Exception caught during execution: ",e)
        logging.exception(f'Exception caught during execution: {e}')
        
        # BU_LOG entry(Failed) in PROCESS_LOG table
        log_json = '[{"JOB_ID": "'+str(job_id)+'","JOB_NAME": "'+str(job_name)+'","CURRENT_DATETIME": "'+str(datetime.now())+'","STATUS": "FAILED"}]'
        bu_alerts.bulog(process_name=job_name,table_name=table_name,status='FAILED',process_owner=owner ,row_count=0,log=log_json,database=database,warehouse=warehouse)
        
        bu_alerts.send_mail(
            receiver_email = receiver_email,
            mail_subject =f'JOB FAILED - {job_name}',
            mail_body = f'{job_name} failed during execution, Attached logs',
            attachment_location = logfile)
        
        sys.exit(-1)
    time_end = time.time()
    logging.warning('It took {} seconds to run.'.format(time_end - time_start))
    
if __name__ == "__main__":
    corn_bid_runner()
